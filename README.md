# Consulta ONPE - Miembros de Mesa

Este proyecto permite cargar un archivo Excel con DNIs, consultar la pagina de ONPE y completar automaticamente los datos encontrados, como por ejemplo:

- `miembro_de_mesa`
- `ubicacion`
- `direccion`
- otras columnas reconocidas si existen en el Excel

El proyecto se puede usar de 2 formas:

1. `Entorno WEB`
2. `Entorno Escritorio (CustomTkinter)`

## Requisitos generales

Antes de usar el proyecto, asegurate de tener instalado lo siguiente:

- `Python 3.12` o una version compatible
- `pip`
- `Docker Desktop` si vas a usar la version web dockerizada
- Un navegador compatible con Chromium en tu maquina local para el procesamiento ONPE:
  - `Microsoft Edge`
  - `Google Chrome`
  - `Brave`

## Formato del Excel

El archivo Excel debe ser `.xlsx` y debe tener como minimo la columna:

```text
dni
```

Ejemplo basico:

```text
dni | miembro_de_mesa | ubicacion | direccion
70730440
10372619
60978226
```

Notas:

- La columna `dni` es obligatoria.
- Las demas columnas pueden estar vacias.
- Si agregas columnas reconocidas por el sistema, tambien se intentaran completar.

## Entorno WEB

Esta modalidad usa:

- una web dockerizada para subir y descargar archivos
- un worker local en Python para hacer la consulta real a ONPE

Minitutorial:

- https://youtu.be/RJskMOvevi4

Esto permite que la imagen Docker sea mucho mas ligera que una imagen que intenta llevar Chromium dentro del contenedor.

### 1. Instalar dependencias locales para el worker

En una terminal abierta en la carpeta del proyecto, ejecuta:

```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

Si ya tienes el entorno virtual creado, solo usa:

```powershell
.\venv\Scripts\activate
```

### 2. Construir la imagen web

En la misma carpeta del proyecto, ejecuta:

```powershell
docker build --no-cache -f Dockerfile.optimizado -t pc01-fernandomas-web:optimizado .
```

### 3. Ejecutar la web dockerizada

Abre una terminal y ejecuta:

```powershell
docker run --rm -it -p 8000:8000 pc01-fernandomas-web:optimizado
```

Deja esa terminal abierta.

Cuando todo vaya bien, veras algo parecido a esto:

```text
* Running on http://127.0.0.1:8000
```

### 4. Abrir la pagina web

En tu navegador, entra a:

```text
http://localhost:8000
```

### 5. Subir el Excel

Dentro de la web:

1. haz clic en `Crear trabajo`
2. selecciona tu archivo `.xlsx`
3. espera a que la pagina muestre el trabajo creado

### 6. Ejecutar el worker local

Abre una segunda terminal en la carpeta del proyecto.

Activa el entorno virtual:

```powershell
.\venv\Scripts\activate
```

Luego ejecuta:

```powershell
python worker_host.py --server http://localhost:8000
```

Ese worker hara lo siguiente:

- tomara el trabajo pendiente desde la web
- descargara el Excel subido
- consultara ONPE usando tu navegador local
- devolvera el Excel actualizado a la web

### 7. Descargar el Excel actualizado

Cuando el worker termine:

1. vuelve a la pestaña de la web
2. espera a que cambie el estado del trabajo
3. haz clic en `Descargar Excel actualizado`

## Resumen rapido del Entorno WEB

Usa 2 terminales:

### Terminal 1

```powershell
docker run --rm -it -p 8000:8000 pc01-fernandomas-web:optimizado
```

### Terminal 2

```powershell
.\venv\Scripts\activate
python worker_host.py --server http://localhost:8000
```

### Navegador

```text
http://localhost:8000
```

## Entorno Escritorio (CustomTkinter)

Esta modalidad ejecuta todo de forma local en una aplicacion de escritorio.

Minitutorial:

- https://youtu.be/ucsk1xSiwLM

### 1. Instalar dependencias

En una terminal ubicada en la carpeta del proyecto, ejecuta:

```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

Si ya tienes el entorno virtual creado, solo usa:

```powershell
.\venv\Scripts\activate
```

### 2. Ejecutar la aplicacion de escritorio

```powershell
python main.py
```

### 3. Usar la aplicacion

Dentro de la ventana:

1. haz clic en `Cargar Excel`
2. selecciona tu archivo `.xlsx`
3. espera a que inicie la consulta automaticamente
4. observa la vista previa en pantalla
5. al terminar, el mismo Excel original se sobrescribe con los resultados

### Importante en escritorio

- El archivo Excel debe estar cerrado antes de procesarlo.
- Si el archivo esta abierto en Microsoft Excel, el programa no podra sobrescribirlo.
- La consulta depende de la web de ONPE y puede tardar varios segundos o minutos segun la respuesta del sitio.

## Solucion de problemas

### La web abre pero el worker falla con `ConnectionRefusedError`

Eso significa que la web dockerizada no esta corriendo.

Primero ejecuta:

```powershell
docker run --rm -it -p 8000:8000 pc01-fernandomas-web:optimizado
```

Y despues, en otra terminal:

```powershell
python worker_host.py --server http://localhost:8000
```

### El Excel no se procesa

Verifica lo siguiente:

- que el archivo sea `.xlsx`
- que tenga la columna `dni`
- que el DNI tenga datos validos
- que tu navegador local compatible con Chromium este instalado

### El escritorio no puede guardar el Excel

Eso normalmente significa que el archivo esta abierto.

Solucion:

- cierra el Excel
- vuelve a ejecutar la consulta

## Archivos principales del proyecto

- [app.py](/c:/Users/Fernando/Desktop/TECSUP2026/Ciclo5/Desarrollo_De_Soluciones_En_La_Nube/Semana04-Ejercicios/Caso1/app.py): web local/dockerizada
- [worker_host.py](/c:/Users/Fernando/Desktop/TECSUP2026/Ciclo5/Desarrollo_De_Soluciones_En_La_Nube/Semana04-Ejercicios/Caso1/worker_host.py): worker local para ONPE en el entorno web
- [main.py](/c:/Users/Fernando/Desktop/TECSUP2026/Ciclo5/Desarrollo_De_Soluciones_En_La_Nube/Semana04-Ejercicios/Caso1/main.py): version escritorio con CustomTkinter
- [onpe_core.py](/c:/Users/Fernando/Desktop/TECSUP2026/Ciclo5/Desarrollo_De_Soluciones_En_La_Nube/Semana04-Ejercicios/Caso1/onpe_core.py): logica compartida de consulta ONPE
- [Dockerfile.optimizado](/c:/Users/Fernando/Desktop/TECSUP2026/Ciclo5/Desarrollo_De_Soluciones_En_La_Nube/Semana04-Ejercicios/Caso1/Dockerfile.optimizado): imagen ligera para la web

## Recomendacion final

Si quieres la forma mas simple para un usuario promedio:

- usa `Entorno Escritorio (CustomTkinter)` si todo se ejecutara en una sola maquina
- usa `Entorno WEB` si quieres una interfaz web dockerizada y estas dispuesto a ejecutar tambien el worker local en otra terminal
