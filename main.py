import os
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl import load_workbook


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

REQUIRED_COLUMNS = [
    "id",
    "dni",
    "nombres",
    "apellidos",
    "miembro_mesa",
    "region",
    "provincia",
    "distrito",
    "direccion_local",
    "numero_mesa",
    "numero_orden",
    "pabellon",
    "piso",
    "aula",
]


class ConsultaElectoralApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Consulta electoral - Excel")
        self.geometry("1320x820")
        self.minsize(1180, 720)

        self.archivo_excel = os.path.join(os.getcwd(), "electores.xlsx")
        self.registros = []

        self._build_ui()
        self._cargar_excel_inicial()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(self, corner_radius=16)
        main.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            main,
            text="Consulta tu local de votación",
            font=ctk.CTkFont(size=30, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="w")

        ctk.CTkLabel(
            main,
            text=(
                "Carga un archivo Excel y la app mostrará la información de todos los electores "
                "registrados: si son miembros de mesa, su ubicación y su local de votación."
            ),
            wraplength=1180,
            justify="left",
            font=ctk.CTkFont(size=15),
        ).grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        top = ctk.CTkFrame(main)
        top.grid(row=2, column=0, padx=20, pady=(0, 14), sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            top,
            text="Cargar Excel",
            width=160,
            height=40,
            command=self.seleccionar_excel,
        ).grid(row=0, column=0, padx=12, pady=12)

        self.archivo_entry = ctk.CTkEntry(top, height=40)
        self.archivo_entry.grid(row=0, column=1, padx=(0, 12), pady=12, sticky="ew")
        self.archivo_entry.insert(0, self.archivo_excel)

        self.archivo_status = ctk.CTkLabel(
            main,
            text="Archivo: sin cargar",
            anchor="w",
            justify="left",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.archivo_status.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")

        self.lista_frame = ctk.CTkScrollableFrame(main)
        self.lista_frame.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.lista_frame.grid_columnconfigure(0, weight=1)

        self.empty_label = ctk.CTkLabel(
            self.lista_frame,
            text="Todavía no hay electores cargados.",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        self.empty_label.grid(row=0, column=0, padx=20, pady=30, sticky="w")

    def _cargar_excel_inicial(self):
        if os.path.exists(self.archivo_excel):
            try:
                self.cargar_excel(self.archivo_excel)
            except Exception as exc:
                self.archivo_status.configure(
                    text=f"Archivo: no se pudo cargar el Excel inicial ({exc})",
                    text_color="#B91C1C",
                )
        else:
            self.archivo_status.configure(
                text="Archivo: selecciona un Excel para empezar",
                text_color="#92400E",
            )

    def seleccionar_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx")],
        )
        if not path:
            return

        try:
            self.cargar_excel(path)
            messagebox.showinfo(
                "Excel cargado",
                "El archivo se cargó correctamente y ya se muestran todos los registros.",
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel.\n\n{exc}")

    def cargar_excel(self, path):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb["electores"] if "electores" in wb.sheetnames else wb[wb.sheetnames[0]]

        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            raise ValueError("El archivo Excel está vacío.")

        headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
        normalized_headers = [h.lower() for h in headers]

        missing = [col for col in REQUIRED_COLUMNS if col not in normalized_headers]
        if missing:
            raise ValueError("Faltan columnas obligatorias en el Excel: " + ", ".join(missing))

        index_map = {name: normalized_headers.index(name) for name in REQUIRED_COLUMNS}
        registros = []
        for row in rows[1:]:
            if row is None:
                continue
            registro = {}
            empty = True
            for field, idx in index_map.items():
                value = row[idx] if idx < len(row) else ""
                if value is None:
                    value = ""
                value = str(value).strip()
                if value:
                    empty = False
                registro[field] = value
            if not empty:
                registros.append(registro)

        self.registros = registros
        self.archivo_excel = path
        self.archivo_entry.delete(0, "end")
        self.archivo_entry.insert(0, path)
        self.archivo_status.configure(
            text=f"Archivo cargado: {os.path.basename(path)} | Registros: {len(registros)}",
            text_color="#166534",
        )
        self._render_registros()

    def _clear_lista_frame(self):
        for widget in self.lista_frame.winfo_children():
            widget.destroy()

    def _render_registros(self):
        self._clear_lista_frame()

        if not self.registros:
            ctk.CTkLabel(
                self.lista_frame,
                text="No hay registros para mostrar.",
                font=ctk.CTkFont(size=18, weight="bold"),
            ).grid(row=0, column=0, padx=20, pady=30, sticky="w")
            return

        for index, elector in enumerate(self.registros):
            self._crear_card_elector(index, elector)

    def _crear_card_elector(self, row_index, elector):
        card = ctk.CTkFrame(self.lista_frame, corner_radius=14)
        card.grid(row=row_index, column=0, padx=10, pady=10, sticky="ew")
        card.grid_columnconfigure(0, weight=2)
        card.grid_columnconfigure(1, weight=3)

        es_miembro = elector["miembro_mesa"].upper() == "SI"
        estado_texto = "SI ERES MIEMBRO DE MESA" if es_miembro else "NO ERES MIEMBRO DE MESA"
        estado_color = "#15803D" if es_miembro else "#E11D48"

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, padx=16, pady=(16, 10), sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        estado_badge = ctk.CTkFrame(header, fg_color=estado_color, corner_radius=12)
        estado_badge.grid(row=0, column=0, padx=(0, 12), sticky="w")
        ctk.CTkLabel(
            estado_badge,
            text=estado_texto,
            text_color="white",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(padx=16, pady=12)

        nombre_completo = f"{elector['nombres']} {elector['apellidos']}".strip()
        ctk.CTkLabel(
            header,
            text=f"DNI: {elector['dni']}\nNombre completo: {nombre_completo}",
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=1, sticky="ew")

        izquierda = ctk.CTkFrame(card)
        izquierda.grid(row=1, column=0, padx=(16, 8), pady=(0, 16), sticky="nsew")
        izquierda.grid_columnconfigure(0, weight=1)

        derecha = ctk.CTkFrame(card)
        derecha.grid(row=1, column=1, padx=(8, 16), pady=(0, 16), sticky="nsew")
        derecha.grid_columnconfigure((0, 1), weight=1)

        self._add_detail(izquierda, "Ubicación", f"{elector['region']} / {elector['provincia']} / {elector['distrito']}")
        self._add_detail(izquierda, "Dirección del local de votación", elector["direccion_local"])

        self._add_grid_detail(derecha, 0, 0, "Número de mesa", elector["numero_mesa"])
        self._add_grid_detail(derecha, 0, 1, "Número de orden", elector["numero_orden"])
        self._add_grid_detail(derecha, 1, 0, "Pabellón", elector["pabellon"])
        self._add_grid_detail(derecha, 1, 1, "Piso", elector["piso"])
        self._add_grid_detail(derecha, 2, 0, "Aula", elector["aula"])

    def _add_detail(self, parent, title, value):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", padx=10, pady=8)

        ctk.CTkLabel(
            frame,
            text=title,
            anchor="w",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(fill="x", padx=14, pady=(12, 4))

        ctk.CTkLabel(
            frame,
            text=value or "-",
            anchor="w",
            justify="left",
            wraplength=440,
            font=ctk.CTkFont(size=17),
        ).pack(fill="x", padx=14, pady=(0, 12))

    def _add_grid_detail(self, parent, row, col, title, value):
        frame = ctk.CTkFrame(parent)
        frame.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

        ctk.CTkLabel(
            frame,
            text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(padx=14, pady=(12, 4))

        ctk.CTkLabel(
            frame,
            text=value or "-",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(padx=14, pady=(0, 12))


if __name__ == "__main__":
    app = ConsultaElectoralApp()
    app.mainloop()
