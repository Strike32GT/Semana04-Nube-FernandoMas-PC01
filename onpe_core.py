import os
import re
import shutil
import subprocess
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

ONPE_URL = "https://consultaelectoral.onpe.gob.pe/inicio"
DEFAULT_EXCEL_NAME = "ExcelPrueba.xlsx"
BROWSER_ENV_VAR = "ONPE_BROWSER_PATH"
HEADLESS_ENV_VAR = "ONPE_HEADLESS"
SUPPORTED_BROWSER_FAMILIES = {"edge", "chrome", "brave"}
BROWSER_COMMAND_MAP = {
    "edge": ["msedge", "microsoft-edge", "microsoft-edge-stable"],
    "chrome": ["chrome", "google-chrome", "google-chrome-stable", "chromium", "chromium-browser"],
    "brave": ["brave", "brave-browser"],
    "firefox": ["firefox"],
    "safari": [],
}
BROWSER_PROGID_MAP = {
    "MSEdgeHTM": "edge",
    "ChromeHTML": "chrome",
    "BraveHTML": "brave",
    "FirefoxURL": "firefox",
    "FirefoxHTML": "firefox",
}
BROWSER_DESKTOP_MAP = {
    "microsoft-edge.desktop": "edge",
    "google-chrome.desktop": "chrome",
    "chromium.desktop": "chrome",
    "brave-browser.desktop": "brave",
    "firefox.desktop": "firefox",
}
BROWSER_BUNDLE_MAP = {
    "com.microsoft.edgemac": "edge",
    "com.google.chrome": "chrome",
    "com.brave.browser": "brave",
    "org.mozilla.firefox": "firefox",
    "com.apple.safari": "safari",
}


@dataclass
class ConsultaResultado:
    values: dict = field(default_factory=dict)
    error: str = ""


class OnpeBrowserClient:
    def __init__(self):
        self.playwright = None
        self.browser = None
        self.page = None
        self.browser_path, self.browser_family = self._resolve_browser()
        self.headless = os.environ.get(HEADLESS_ENV_VAR, "").strip().lower() in {"1", "true", "yes", "on"}

    def __enter__(self):
        self.playwright = sync_playwright().start()
        launch_args = ["--no-sandbox", "--disable-dev-shm-usage"]
        if not self.headless:
            launch_args.append("--start-minimized")

        self.browser = self.playwright.chromium.launch(
            headless=self.headless,
            executable_path=self.browser_path,
            args=launch_args,
        )
        self.page = self.browser.new_page(viewport={"width": 1600, "height": 900})
        self.page.set_default_timeout(45000)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.page is not None:
            self.page.close()
        if self.browser is not None:
            self.browser.close()
        if self.playwright is not None:
            self.playwright.stop()

    def consultar_dni(self, dni):
        self.page.goto(ONPE_URL, wait_until="domcontentloaded")
        self.page.locator("input").first.fill(str(dni))
        self.page.get_by_role("button", name=re.compile("consultar", re.IGNORECASE)).click()

        try:
            self.page.wait_for_url("**/local-de-votacion", timeout=30000)
        except PlaywrightTimeoutError as exc:
            body_text = self.page.locator("body").inner_text()
            raise RuntimeError(self._infer_error_message(body_text) or "ONPE no devolvio resultados para ese DNI.") from exc

        self.page.wait_for_timeout(1500)
        body_text = self.page.locator("body").inner_text()
        return self._parse_result(body_text, str(dni))

    def _parse_result(self, text, expected_dni):
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        normalized_lines = [self._normalize(line) for line in lines]

        miembro = self._extract_member_status(text, normalized_lines)

        nombres = self._value_after_label(lines, normalized_lines, "nombres y apellidos")
        ubicacion = self._value_after_label(lines, normalized_lines, "region provincia distrito")
        local_lines = self._local_block(lines, normalized_lines)
        direccion = " | ".join(local_lines)

        numero_mesa = self._value_after_label(lines, normalized_lines, "n de mesa")
        numero_orden = self._value_after_label(lines, normalized_lines, "n de orden")
        pabellon = self._value_after_label(lines, normalized_lines, "pabellon")
        piso = self._value_after_label(lines, normalized_lines, "piso")
        aula = self._value_after_label(lines, normalized_lines, "aula")

        region, provincia, distrito = self._split_ubicacion(ubicacion)

        values = {
            "dni": expected_dni,
            "miembro_de_mesa": miembro,
            "ubicacion": ubicacion,
            "direccion": direccion,
            "nombres_y_apellidos": nombres,
            "nombre_completo": nombres,
            "region": region,
            "provincia": provincia,
            "distrito": distrito,
            "numero_mesa": numero_mesa,
            "numero_de_mesa": numero_mesa,
            "nro_mesa": numero_mesa,
            "numero_orden": numero_orden,
            "numero_de_orden": numero_orden,
            "nro_orden": numero_orden,
            "pabellon": pabellon,
            "piso": piso,
            "aula": aula,
        }

        if not miembro and not ubicacion and not direccion:
            raise RuntimeError("No se pudo interpretar la respuesta de ONPE para ese DNI.")

        return ConsultaResultado(values=values)

    def _extract_member_status(self, text, normalized_lines):
        normalized_text = self._normalize(text)

        no_patterns = [
            r"\bno\s+eres\s+miembro\s+de\s+mesa\b",
            r"\bno\s+miembro\s+de\s+mesa\b",
        ]
        si_patterns = [
            r"\bsi\s+eres\s+miembro\s+de\s+mesa\b",
            r"\beres\s+miembro\s+de\s+mesa\b",
            r"\bsi\s+miembro\s+de\s+mesa\b",
            r"\beres\s+presidente\b",
            r"\beres\s+secretario\b",
            r"\beres\s+tercer\s+miembro\b",
            r"\beres\s+suplente\b",
        ]

        for pattern in no_patterns:
            if re.search(pattern, normalized_text):
                return "no"

        for pattern in si_patterns:
            if re.search(pattern, normalized_text):
                return "si"

        for index, line in enumerate(normalized_lines):
            if (
                "miembro de mesa" not in line
                and "secretario" not in line
                and "presidente" not in line
                and "tercer miembro" not in line
                and "suplente" not in line
            ):
                continue

            window = " ".join(normalized_lines[max(0, index - 2): min(len(normalized_lines), index + 3)])
            if re.search(r"\bno\s+eres\s+miembro\s+de\s+mesa\b", window):
                return "no"
            if re.search(r"\bsi\s+eres\s+miembro\s+de\s+mesa\b", window):
                return "si"
            if re.search(r"\beres\s+miembro\s+de\s+mesa\b", window):
                return "si"
            if re.search(r"\beres\s+presidente\b", window):
                return "si"
            if re.search(r"\beres\s+secretario\b", window):
                return "si"
            if re.search(r"\beres\s+tercer\s+miembro\b", window):
                return "si"
            if re.search(r"\beres\s+suplente\b", window):
                return "si"
            if line in {"miembro de mesa", "secretario", "presidente", "tercer miembro", "suplente"}:
                return "si"

        return ""

    def _local_block(self, lines, normalized_lines):
        start = self._find_line_index(normalized_lines, "tu local de votacion")
        if start == -1:
            return []

        collected = []
        skip = {"ver", "mapa", "descargar", "croquis", "capacitate"}
        stop_tokens = {"n de mesa", "n de orden", "pabellon", "piso", "aula", "oficina central"}

        for index in range(start + 1, len(lines)):
            current_norm = normalized_lines[index]
            if current_norm in skip:
                continue
            if current_norm in stop_tokens:
                break
            if current_norm.startswith("n de mesa"):
                break
            collected.append(lines[index])

        return collected

    def _value_after_label(self, lines, normalized_lines, label):
        idx = self._find_line_index(normalized_lines, label)
        if idx == -1:
            return ""
        for pos in range(idx + 1, len(lines)):
            if normalized_lines[pos] != label:
                return lines[pos]
        return ""

    def _find_line_index(self, normalized_lines, expected):
        for idx, line in enumerate(normalized_lines):
            if line == expected:
                return idx
        return -1

    def _split_ubicacion(self, ubicacion):
        if not ubicacion:
            return "", "", ""
        parts = [part.strip() for part in ubicacion.split("/")]
        while len(parts) < 3:
            parts.append("")
        return parts[0], parts[1], parts[2]

    def _infer_error_message(self, body_text):
        normalized = self._normalize(body_text)
        if "no se encontraron datos" in normalized:
            return "ONPE no encontro datos para ese DNI."
        if "ingresa un dni valido" in normalized:
            return "El DNI no es valido."
        if "ocurrio un error" in normalized:
            return "ONPE devolvio un error durante la consulta."
        return ""

    def _resolve_browser(self):
        env_path = os.environ.get(BROWSER_ENV_VAR, "").strip()
        if env_path:
            if os.path.exists(env_path):
                return env_path, self._infer_family_from_path(env_path)
            raise RuntimeError(f"La ruta definida en {BROWSER_ENV_VAR} no existe: {env_path}")

        default_family, default_path = self._detect_default_browser()
        if default_family:
            if default_family not in SUPPORTED_BROWSER_FAMILIES:
                raise RuntimeError(
                    f"El navegador predeterminado detectado es {default_family}, pero esta automatizacion "
                    "solo funciona con navegadores Chromium como Edge, Chrome o Brave."
                )
            if default_path:
                return default_path, default_family

        fallback = self._find_supported_browser_in_path()
        if fallback:
            return fallback

        raise RuntimeError(
            "No se pudo detectar un navegador Chromium compatible. "
            f"Define la variable {BROWSER_ENV_VAR} o configura Edge, Chrome o Brave como navegador disponible."
        )

    def _detect_default_browser(self):
        if sys.platform.startswith("win"):
            return self._detect_default_browser_windows()
        if sys.platform == "darwin":
            return self._detect_default_browser_macos()
        return self._detect_default_browser_linux()

    def _detect_default_browser_windows(self):
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice",
            ) as key:
                prog_id = winreg.QueryValueEx(key, "ProgId")[0]
        except Exception:
            return None, None

        family = BROWSER_PROGID_MAP.get(prog_id, "")
        if not family:
            return None, None

        for exe_name in BROWSER_COMMAND_MAP.get(family, []):
            name = exe_name + ".exe" if not exe_name.endswith(".exe") else exe_name
            path = self._resolve_windows_app_path(name)
            if path:
                return family, path
        return family, None

    def _resolve_windows_app_path(self, exe_name):
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}",
            ) as key:
                return winreg.QueryValue(key, None)
        except Exception:
            pass

        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}",
            ) as key:
                return winreg.QueryValue(key, None)
        except Exception:
            return shutil.which(exe_name)

    def _detect_default_browser_linux(self):
        try:
            result = subprocess.run(
                ["xdg-settings", "get", "default-web-browser"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            desktop_file = result.stdout.strip()
        except Exception:
            desktop_file = ""

        family = BROWSER_DESKTOP_MAP.get(desktop_file, "")
        if not family:
            return None, None

        for command in BROWSER_COMMAND_MAP.get(family, []):
            path = shutil.which(command)
            if path:
                return family, path
        return family, None

    def _detect_default_browser_macos(self):
        try:
            result = subprocess.run(
                [
                    "osascript",
                    "-e",
                    'id of application (path to default application for URL "https://consultaelectoral.onpe.gob.pe")',
                ],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            bundle_id = result.stdout.strip().lower()
        except Exception:
            bundle_id = ""

        family = BROWSER_BUNDLE_MAP.get(bundle_id, "")
        if not family:
            return None, None

        app_path = self._find_macos_app(bundle_id)
        return family, app_path

    def _find_macos_app(self, bundle_id):
        try:
            result = subprocess.run(
                ["mdfind", f'kMDItemCFBundleIdentifier == "{bundle_id}"'],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            app_bundle = next((line.strip() for line in result.stdout.splitlines() if line.strip().endswith(".app")), "")
        except Exception:
            app_bundle = ""

        if not app_bundle:
            return None

        macos_dir = Path(app_bundle) / "Contents" / "MacOS"
        if not macos_dir.exists():
            return None

        for child in macos_dir.iterdir():
            if child.is_file():
                return str(child)
        return None

    def _find_supported_browser_in_path(self):
        for family in ("edge", "chrome", "brave"):
            for command in BROWSER_COMMAND_MAP.get(family, []):
                path = shutil.which(command)
                if path:
                    return path, family
        return None

    def _infer_family_from_path(self, path):
        normalized = str(path).lower()
        if "brave" in normalized:
            return "brave"
        if "edge" in normalized or "msedge" in normalized:
            return "edge"
        if "chrome" in normalized or "chromium" in normalized:
            return "chrome"
        if "firefox" in normalized:
            return "firefox"
        if "safari" in normalized:
            return "safari"
        return "desconocido"

    def _normalize(self, text):
        text = str(text or "")
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = text.lower()
        text = text.replace("°", " ")
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return text.strip()


def normalize_header(text):
    text = str(text or "")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip().replace("_", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = text.strip()
    aliases = {
        "miembro de mesa": "miembro_de_mesa",
        "miembro mesa": "miembro_de_mesa",
        "miembro_de_mesa": "miembro_de_mesa",
        "numero mesa": "numero_mesa",
        "n de mesa": "numero_mesa",
        "nro de mesa": "numero_mesa",
        "numero de mesa": "numero_mesa",
        "numero orden": "numero_orden",
        "n de orden": "numero_orden",
        "nro de orden": "numero_orden",
        "numero de orden": "numero_orden",
        "nombres y apellidos": "nombres_y_apellidos",
        "nombre completo": "nombre_completo",
    }
    return aliases.get(text, text.replace(" ", "_"))


def cell_to_str(value):
    return "" if value is None else str(value).strip()
