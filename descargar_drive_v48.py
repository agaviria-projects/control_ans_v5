import os
import io
import gspread
import pandas as pd
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from gspread.utils import rowcol_to_a1
from pathlib import Path
import zipfile

# ============================================================
# CONFIGURACIÓN BASE
# ============================================================
CRED_PATH = r"C:\Users\hector.gaviria\Desktop\Control_ANS\control-ans-elite-f4ea102db569.json"
SHEET_ID = "1bPLGVVz50k6PlNp382isJrqtW_3IsrrhGW0UUlMf-bM"

# ============================================================
# DETECCIÓN DE ENTORNO
# ============================================================
RUTA_EMPRESA = Path(r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS\Evidencias_PDF")

if RUTA_EMPRESA.exists():
    RUTA_DESTINO = RUTA_EMPRESA
    print("🏢 Entorno detectado: Empresa (OneDrive conectado)")
else:
    RUTA_DESTINO = Path(r"C:\Users\Acer\Desktop\Evidencias_PDF")
    print("💻 Entorno detectado: Personal (modo pruebas en Desktop)")

print(f"📂 Carpeta destino base: {RUTA_DESTINO}")

# ============================================================
# AUTENTICACIÓN GOOGLE DRIVE
# ============================================================
def crear_servicio():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)
# ============================================================
# CONECTAR GOOGLE SHEET
# ============================================================
def conectar_gspread():
    creds = service_account.Credentials.from_service_account_file(
        CRED_PATH, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_ID)

    for ws in spreadsheet.worksheets():
        nombre = ws.title.lower().replace(" ", "")
        if "form" in nombre or "respuesta" in nombre:
            print(f"📄 Hoja activa detectada: {ws.title}")
            return ws

    print("⚠️ No se detectó hoja de respuestas; usando la primera hoja.")
    return spreadsheet.sheet1

# ============================================================
# LEER GOOGLE SHEET COMO CSV
# ============================================================
def leer_google_sheet(service):
    try:
        request = service.files().export_media(fileId=SHEET_ID, mimeType="text/csv")
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        df = pd.read_csv(fh)
        print("✅ Hoja leída correctamente.\n")
        print(df.head())
        return df

    except Exception as e:
        print(f"❌ Error al leer Google Sheet: {e}")
        return None

# ============================================================
# DESCARGAR PDFS + RENOMBRAR + COMPRESIÓN
# ============================================================
def descargar_pdfs(service, df):

    # Normalizar columnas
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("á", "a")
        .str.replace("é", "e")
        .str.replace("í", "i")
        .str.replace("ó", "o")
        .str.replace("ú", "u")
        .str.replace("ñ", "n")
    )

    # ============ FECHA REAL =============
    col_fecha = next((c for c in df.columns if "marca" in c), None)

    if col_fecha:
        # Tomar fecha real desde la Marca Temporal
        df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        df = df.dropna(subset=[col_fecha])

        # Extraer fecha real del formulario
        df["fecha_real"] = df[col_fecha].dt.strftime("%Y-%m-%d")
        print("📌 Registros encontrados:", len(df), "\n")
        
        # 🔥 Solo descargas del día
        fecha_mas_reciente = df["fecha_real"].max()
        df = df[df["fecha_real"] == fecha_mas_reciente]
        print(f"📌 Filtrando solo registros del día: {fecha_mas_reciente} → {len(df)} registros")
    else:
        fecha_form = datetime.now().strftime("%Y-%m-%d")
        print("⚠️ No se detectó columna fecha. Se procesarán todos los registros.")

    # Columnas clave
    col_pedido = next((c for c in df.columns if "pedido" in c), None)
    col_tecnico = next((c for c in df.columns if "tecnic" in c), None)
    col_actividad = next((c for c in df.columns if "actividad" in c), None)
    col_url = next((c for c in df.columns if "evidenc" in c), None)

    if not all([col_pedido, col_tecnico, col_actividad, col_url]):
        print("❌ Columnas clave no encontradas.")
        return None

    RESPONSABLES = {
        "ARTER-(REPLANTEO PREPAGO)": "Lina",
        "ARTER-(REPLANTEO HV)": "Lina",
        "ACREV-(PUNTOS DE CONEXION)": "Lina",
        "ALEGA-(LEGALIZACION RESIDENCIAL)": "Frank",
        "ALEGN-(LEGALIZACION NO RESIDENCIAL)": "Frank",
        "ALECA-(REFORMA RESIDENCIAL)": "Frank",
        "ACAMN-(REFORMA NO RESIDENCIAL)": "Frank",
        "AEJDO-(HV SENCILLO)": "Lina",
        "INPRE-(EJECUCION PREPAGO)": "Lina",
        "AMRTR-(MOVIMIENTOS DE REDES)": "Robinson",
        "AEJDO-(HV MAS INTERNA)": "Lina",
        "REEQU-(TRABAJOS PREPAGO)": "Lina",
        "DIPRE-(RETIRO PREPAGO)": "Lina"
    }
    def obtener_ruta_destino(actividad, fecha_real):
        responsable = RESPONSABLES.get(actividad, "Sin_Asignar")
        carpeta_responsable = RUTA_DESTINO / responsable
        carpeta_responsable.mkdir(parents=True, exist_ok=True)

        ruta_final = carpeta_responsable / fecha_real / actividad
        ruta_final.mkdir(parents=True, exist_ok=True)
        return ruta_final


    errores = 0
    total_encontrados = len(df)
    total_descargados = 0
    total_comprimidos = 0
    total_sin_comprimir = 0


    # ====================================================
    # LOOP PRINCIPAL
    # ====================================================
    for i, fila in df.iterrows():

        pedido = str(fila.get(col_pedido, "")).strip()
        tecnico = str(fila.get(col_tecnico, "")).strip()
        actividad = str(fila.get(col_actividad, "")).strip()
        url = str(fila.get(col_url, "")).strip()

        # 🔎 DEBUG: mostrar actividad exactamente como llega desde Google Sheets
        print(f"Actividad recibida: »{actividad}«")

        if not (pedido and tecnico and url):
            continue
        if "id=" not in url:
            continue

        file_id = url.split("id=")[-1]

        # Validar existencia
        try:
            service.files().get(fileId=file_id).execute()
        except Exception:
            continue

        ruta_destino = obtener_ruta_destino(actividad, fila["fecha_real"])
        base_name = f"EPM-FNX-{pedido}-257"

        # Verificar consecutivos del día
        existentes = list(ruta_destino.glob(f"{base_name}-(*).pdf"))

        consecutivo = (
            max([
                int(e.stem.split("(")[-1].replace(")", ""))
                for e in existentes
            ]) + 1
            if existentes else 1
        )

        nombre_archivo = f"{base_name}-({consecutivo}).pdf"
        ruta_local = ruta_destino / nombre_archivo

        print(f"⬇️ Descargando {nombre_archivo} ...")

        # Si el archivo ya existe, saltarlo (ahorra mucho tiempo)
        if ruta_local.exists():
            print(f"⏩ Ya existe → {ruta_local.name} (se omite)")
            total_sin_comprimir += 1
            continue

        # Intentos con chunk optimizado
        exitoso = False
        for intento in range(1, 4):
            try:
                request = service.files().get_media(fileId=file_id)

                # chunksize reducido → más estable
                with io.FileIO(ruta_local, "wb") as fh:
                    downloader = MediaIoBaseDownload(fh, request, chunksize=512 * 1024)
                    done = False

                    while not done:
                        status, done = downloader.next_chunk()

                exitoso = True
                break

            except Exception as e:
                print(f"⚠️ Error descargando (intentó {intento}/3): {e}")
                if ruta_local.exists():
                    ruta_local.unlink()
                if intento < 3:
                  time.sleep(1)  # pequeño respiro antes del reintento 

        if not exitoso:
            errores += 1
            print("❌ Falló descarga definitiva, se omite.\n")
            continue

        total_descargados += 1

        # COMPRESIÓN ≥ 20 KB
        peso_kb = ruta_local.stat().st_size / 1024
        if peso_kb >= 20:
            zip_path = ruta_local.with_suffix(".zip")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(ruta_local, ruta_local.name)
            ruta_local.unlink()
            total_comprimidos += 1  
        else:
            total_sin_comprimir += 1    

    return df["fecha_real"].iloc[-1]

# ============================================================
# ACTUALIZAR RUTAS EN GOOGLE SHEET
# ============================================================
def actualizar_rutas_locales(df, fecha_form):

    print("\n🔄 Actualizando enlaces en Google Sheet...\n")

    try:
        sheet = conectar_gspread()
    except:
        return

    data = sheet.get_all_records()
    encabezados = sheet.row_values(1)

    col_evid = None
    for idx, name in enumerate(encabezados, start=1):
        if "evidenc" in name.lower().replace(" ", ""):
            col_evid = idx
            break

    if not col_evid:
        return

    for i, fila in enumerate(data, start=2):
        pedido = str(fila.get("Número del pedido", "")).strip()
        if not pedido:
            continue

        patron = f"EPM-FNX-{pedido}-257-(1).*"
        ruta_local = next((p for p in RUTA_DESTINO.glob(f"**/{patron}")), None)

        if ruta_local and ruta_local.exists():

            celda = rowcol_to_a1(i, col_evid)

            ruta_web = str(ruta_local).replace(
                r"C:\Users\hector.gaviria\OneDrive - Elite Ingenieros SAS",
                "https://eliteingenierosas-my.sharepoint.com/personal/h_gaviria_eliteingenieros_com_co/Documents"
            ).replace("\\", "/")

            sheet.update_acell(celda, f'=HIPERVINCULO("{ruta_web}"; "Abrir")')

            print(f"✔️ Enlace actualizado → {ruta_local.name}")

# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================
if __name__ == "__main__":
    service = crear_servicio()
    df = leer_google_sheet(service)

    if df is not None:
        fecha_form = descargar_pdfs(service, df)
        if fecha_form:
            actualizar_rutas_locales(df, fecha_form)
