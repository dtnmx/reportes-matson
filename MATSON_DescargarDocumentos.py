#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MATSON_DescargarDocumentos.py
Descarga automática de documentos (CI, BOL, POD) desde el sistema GM
para el proceso de carga al portal MATSON.

Basado en Sit_Cobranza_DescargarReporte.py — reutiliza login, alert,
pop-up diario, logout y sistema de logs.

Clasificación de documentos a descargar (derivada automáticamente del Excel):
  - "LG ELECTRONICS" en Ruta/Concepto  →  LG   →  solo CI
  - Tipo de Viaje = "EXPORTACION"       →  USA  →  CI, BOL y POD
  - Cualquier otro caso                 →  TODOS →  CI y BOL
"""

import sys
import time
import shutil
import pandas as pd
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from log_SitCobranza_config import escribir_log

# ============================================================
#  CONSTANTES — Ajusta estos valores según sea necesario
# ============================================================

# --- Credenciales GM ---
GM_URL_LOGIN  = "https://www.softwareparatransporte.com/GMTERPV8_WEB/ES/PAGE_CatUsuariosLoginAWP.awp"
GM_EMPRESA    = "SLC080722SM6"
GM_USUARIO    = "DTN"
GM_CONTRASENA = "RenGax8*"

# --- Archivo Excel de entrada ---
# Cambiar esta ruta al Excel del batch correspondiente.
ARCHIVO_EXCEL = r"I:\Mi unidad\DTN\Proyectos\Sega Carriers\2026-04-07_DTN-SEG-009\Documentos\CARGA MATSON 28 ABRIL.xlsx"

# Nombres de columna exactos del Excel de Facturación (no modificar salvo que cambie el formato)
COLUMNA_NUMERO_VIAJE  = "Núm. de Viaje"    # Número de viaje en GM
COLUMNA_TIPO_VIAJE    = "Tipo de Viaje"    # "INTERMODAL" o "EXPORTACION"
COLUMNA_RUTA_CONCEPTO = "Ruta/Concepto"    # Contiene el nombre del cliente (ej. "LG ELECTRONICS")

# Palabra clave en Ruta/Concepto para identificar viajes LG (insensible a mayúsculas)
KEYWORD_LG = "LG ELECTRONICS"

# --- Carpeta de salida ---
# Se crea automáticamente con la fecha del día. Ej: MATSON_Docs_20260428
FECHA_HOY    = datetime.now().strftime("%Y%m%d")
CARPETA_BASE = f"MATSON_Docs_{FECHA_HOY}"

# --- Documentos a descargar por tipo ---
# Cada entrada es un dict con:
#   "id"     → nombre usado para el archivo descargado (ej. CI_4792578.pdf)
#   "buscar" → texto EXACTO que aparece en la fila del popup de GM
#              (verificado en DevTools: CI, PODC, POD USA, BOL)
DOCS_POR_TIPO = {
    "LG":    [
        {"id": "CI",  "buscar": "CI"},
    ],
    "USA":   [
        {"id": "CI",      "buscar": "CI"},
        {"id": "BOL",     "buscar": "BOL"},
        {"id": "POD",     "buscar": "POD USA"},   # Nombre en GM: "POD USA", archivo: POD_XXXXX.pdf
    ],
    "TODOS": [
        {"id": "CI",  "buscar": "CI"},
        {"id": "BOL", "buscar": "BOL"},
    ],
}

# --- Tiempos de espera (segundos) ---
TIMEOUT_ESPERA = 10   # WebDriverWait timeout general
SLEEP_LOGIN    = 5    # Tras clic de login y tras cerrar alert
SLEEP_MODULO   = 3    # Tras navegar entre módulos/secciones del menú
SLEEP_FILTRO   = 2    # Tras escribir en el buscador de la tabla de viajes
SLEEP_POPUP    = 3    # Tras abrir o cerrar la ventana de documentos adjuntos
SLEEP_DESCARGA = 12   # Segundos máximos para que aparezca el archivo descargado
SLEEP_RETRY    = 3    # Espera antes de reintentar una descarga fallida

# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================

def determinar_tipo_descarga(tipo_viaje: str, ruta_concepto: str) -> str:
    """
    Determina el tipo de descarga (LG / USA / TODOS) a partir de las
    columnas del Excel de Facturación, sin necesitar columna extra.
    """
    ruta = str(ruta_concepto).upper()
    tipo = str(tipo_viaje).upper().strip()

    if KEYWORD_LG.upper() in ruta:
        return "LG"
    elif tipo == "EXPORTACION":
        return "USA"
    else:
        return "TODOS"


def esperar_archivo_nuevo(carpeta: Path, archivos_antes: set, timeout: int = SLEEP_DESCARGA):
    """Espera hasta que aparezca un archivo nuevo y completo (no .crdownload) en carpeta."""
    fin = time.time() + timeout
    while time.time() < fin:
        ahora = {f for f in carpeta.iterdir() if f.is_file()}
        nuevos = ahora - archivos_antes
        completados = [f for f in nuevos if f.suffix.lower() != ".crdownload"]
        if completados:
            return sorted(completados, key=lambda f: f.stat().st_mtime)[-1]
        time.sleep(1)
    return None


def fila_coincide(texto_fila: str, termino_busqueda: str) -> bool:
    """
    Verifica si el texto de una fila del popup coincide con el término de búsqueda.
    Usa comparación exacta (strip) para evitar que "CI" matchee "PODC" o
    que "POD" matchee "PODC" en lugar de "POD USA".
    """
    return texto_fila.strip().upper() == termino_busqueda.strip().upper()


# ============================================================
#  INICIALIZACIÓN DE SELENIUM
# ============================================================

hora_actual = datetime.now().hour

# Crear carpeta de salida y subcarpeta temporal de descargas
carpeta_salida = Path(CARPETA_BASE)
carpeta_salida.mkdir(exist_ok=True)
carpeta_temp = carpeta_salida / "_temp_downloads"
carpeta_temp.mkdir(exist_ok=True)

options = webdriver.ChromeOptions()

if 6 <= hora_actual <= 10:
    options.add_argument("--start-maximized")
    escribir_log("🔁 Ejecutando en modo VISIBLE (sin headless) por ser la primera corrida del día.")
else:
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    escribir_log("🚀 Ejecutando en modo headless.")

options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# Configurar descarga automática a carpeta temporal (sin cuadro de diálogo)
prefs = {
    "download.default_directory":         str(carpeta_temp.resolve()),
    "download.prompt_for_download":       False,
    "download.directory_upgrade":         True,
    "plugins.always_open_pdf_externally": True,
    "safebrowsing.enabled":               True,
}
options.add_experimental_option("prefs", prefs)

service = Service(ChromeDriverManager().install())
driver  = webdriver.Chrome(service=service, options=options)
wait    = WebDriverWait(driver, TIMEOUT_ESPERA)

# ============================================================
#  LOGIN
# ============================================================

escribir_log("📌 Iniciando sesión en GM Transport...")
driver.get(GM_URL_LOGIN)
time.sleep(2)
time.sleep(5)  # Esperar carga completa de la página

driver.find_element(By.NAME, "EDT_EMPRESA").send_keys(GM_EMPRESA)
driver.find_element(By.NAME, "EDT_USUARIO").send_keys(GM_USUARIO)
driver.find_element(By.NAME, "EDT_CONTRASENA").send_keys(GM_CONTRASENA)
driver.find_element(By.ID, "BTN_ENTRAR").click()
time.sleep(SLEEP_LOGIN)

# --- Alert de "sesión abierta" ---
time.sleep(3)
try:
    alert = driver.switch_to.alert
    alert_text = alert.text
    escribir_log(f"✅ Alert detectado: {alert_text}")
    if "El sistema ha detectado una sesión abierta con este usuario" in alert_text:
        alert.accept()
        escribir_log("✅ Alert de sesión abierta aceptado.")
    else:
        escribir_log("ℹ️ Alert desconocido, aceptando.")
        alert.accept()
except:
    escribir_log("ℹ️ No se detectó alert de sesión abierta.")

# --- Pop-up diario ---
try:
    time.sleep(8)
    checkbox = driver.find_element(By.ID, "CBOX_CHECKBOX1_1")
    if not checkbox.is_selected():
        checkbox.click()
        escribir_log("✅ Opción 'No volver a mostrar' marcada.")
    boton_ok = driver.find_element(By.XPATH, "//span[@class='btnvalignmiddle']")
    boton_ok.click()
    escribir_log("✅ Pop-up diario cerrado correctamente.")
    time.sleep(2)
except:
    escribir_log("ℹ️ No se detectó pop-up diario. Continuando.")

# ============================================================
#  CARGAR Y VALIDAR EXCEL DE VIAJES
# ============================================================

escribir_log(f"📌 Cargando archivo Excel: {ARCHIVO_EXCEL}")

try:
    df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
    # Limpiar espacios en nombres de columnas por si el Excel trae espacios extra
    df.columns = df.columns.str.strip()
except FileNotFoundError:
    escribir_log(f"❌ No se encontró el archivo: {ARCHIVO_EXCEL}", nivel="error")
    driver.quit()
    sys.exit(1)
except Exception as e:
    escribir_log(f"❌ Error al leer el Excel: {e}", nivel="error")
    driver.quit()
    sys.exit(1)

for col in [COLUMNA_NUMERO_VIAJE, COLUMNA_TIPO_VIAJE, COLUMNA_RUTA_CONCEPTO]:
    if col not in df.columns:
        escribir_log(f"❌ El Excel no contiene la columna requerida: '{col}'", nivel="error")
        escribir_log(f"   Columnas encontradas: {list(df.columns)}", nivel="error")
        driver.quit()
        sys.exit(1)

# Filtrar filas con número de viaje válido y calcular tipo de descarga
df = df.dropna(subset=[COLUMNA_NUMERO_VIAJE]).copy()
df[COLUMNA_NUMERO_VIAJE] = df[COLUMNA_NUMERO_VIAJE].str.strip().str.replace(r"\.0$", "", regex=True)
df["_tipo_descarga"] = df.apply(
    lambda r: determinar_tipo_descarga(r[COLUMNA_TIPO_VIAJE], r[COLUMNA_RUTA_CONCEPTO]),
    axis=1
)

escribir_log(f"✅ {len(df)} viajes cargados del Excel.")
escribir_log("📌 Distribución de tipos de descarga:")
for tipo, conteo in df["_tipo_descarga"].value_counts().items():
    ids = [d["id"] for d in DOCS_POR_TIPO[tipo]]
    escribir_log(f"   {tipo}: {conteo} viajes → {ids}")

# ============================================================
#  NAVEGAR A TRÁFICO → VIAJES
# ============================================================

try:
    escribir_log("📌 Navegando a la sección 'Tráfico'...")
    time.sleep(SLEEP_MODULO)
    trafico_icono = driver.find_element(
        By.XPATH,
        "//img[contains("
        "translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        "'trafico1.jpg')]"
    )
    trafico_icono.click()
    time.sleep(SLEEP_MODULO)
    escribir_log("✅ Sección 'Tráfico' abierta.")
except Exception as e:
    escribir_log(f"❌ Error al acceder a 'Tráfico': {e}", nivel="error")
    driver.quit()
    sys.exit(1)

try:
    escribir_log("📌 Navegando a 'Viajes'...")
    viajes_icono = driver.find_element(
        By.XPATH,
        "//img[contains("
        "translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        "'trafico/viajes1.jpg')]"
    )
    viajes_icono.click()
    time.sleep(SLEEP_MODULO)
    escribir_log("✅ Sección 'Viajes' abierta.")
except Exception as e:
    escribir_log(f"❌ Error al acceder a 'Viajes': {e}", nivel="error")
    driver.quit()
    sys.exit(1)

# ============================================================
#  PROCESAMIENTO DE VIAJES
# ============================================================

viajes_procesados      = 0
documentos_descargados = 0
viajes_con_error       = []

for _, fila_excel in df.iterrows():
    numero_viaje    = str(fila_excel[COLUMNA_NUMERO_VIAJE]).strip()
    tipo_descarga   = fila_excel["_tipo_descarga"]
    ruta_concepto   = str(fila_excel[COLUMNA_RUTA_CONCEPTO]).strip()
    docs_requeridos = DOCS_POR_TIPO[tipo_descarga]

    escribir_log(f"\n{'='*60}")
    escribir_log(
        f"📌 Viaje: {numero_viaje} | Tipo: {tipo_descarga} | "
        f"Ruta: {ruta_concepto[:50]} | Docs: {[d['id'] for d in docs_requeridos]}"
    )

    # Carpeta destino del viaje
    carpeta_viaje = carpeta_salida / numero_viaje
    carpeta_viaje.mkdir(exist_ok=True)

    # ── a) Buscar el viaje en la tabla ──────────────────────────────────────
    try:
        buscador = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[aria-controls='TABLE_ProViajes']")
            )
        )
        buscador.clear()
        buscador.send_keys(numero_viaje)
        time.sleep(SLEEP_FILTRO)
    except Exception as e:
        escribir_log(f"⚠️ No se pudo acceder al buscador para viaje {numero_viaje}: {e}")
        viajes_con_error.append(numero_viaje)
        continue

    # ── b) Clic en ícono de documentos adjuntos (clip azul) ─────────────────
    try:
        link_doc = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@onclick, 'LINK_DOCADJUNTOS')]")
            )
        )
        link_doc.click()
        time.sleep(SLEEP_POPUP)
        escribir_log(f"📌 Ventana de documentos adjuntos abierta para viaje {numero_viaje}.")
    except Exception as e:
        escribir_log(
            f"⚠️ Viaje {numero_viaje} no aparece en la tabla o sin ícono de docs: {e}"
        )
        viajes_con_error.append(numero_viaje)
        try:
            buscador = driver.find_element(
                By.CSS_SELECTOR, "input[aria-controls='TABLE_ProViajes']"
            )
            buscador.clear()
            time.sleep(1)
        except:
            pass
        continue

    # Manejar el caso en que el popup abra una nueva ventana del navegador
    ventana_principal = driver.current_window_handle
    ventanas_abiertas = driver.window_handles
    if len(ventanas_abiertas) > 1:
        for ventana in ventanas_abiertas:
            if ventana != ventana_principal:
                driver.switch_to.window(ventana)
                break

    # ── c) Descargar documentos según tipo de destino ────────────────────────
    # docs_requeridos es una lista de dicts: {"id": "CI", "buscar": "CI"}, etc.
    ids_descargados = []

    for doc in docs_requeridos:
        doc_id     = doc["id"]      # nombre en el archivo resultante (ej. "POD")
        doc_buscar = doc["buscar"]  # texto exacto de la fila en GM (ej. "POD USA")
        descargado = False

        for intento in range(1, 3):  # Máximo 2 intentos por documento
            try:
                # Filas del popup "Documentos Adjuntos de Viaje".
                # Cada fila tiene texto (CI / PODC / POD USA / BOL) y un span.btnvalignmiddle.
                filas_doc = driver.find_elements(
                    By.XPATH,
                    "//tr[.//span[@class='btnvalignmiddle']]"
                )

                if not filas_doc:
                    escribir_log(
                        f"⚠️ Sin filas de documentos en la ventana del viaje {numero_viaje}."
                    )
                    break

                boton_descarga = None
                for fila in filas_doc:
                    if fila_coincide(fila.text, doc_buscar):
                        try:
                            boton_descarga = fila.find_element(
                                By.XPATH, ".//span[@class='btnvalignmiddle']"
                            )
                        except:
                            pass
                        break

                if boton_descarga is None:
                    escribir_log(
                        f"⚠️ '{doc_buscar}' no encontrado en ventana "
                        f"del viaje {numero_viaje}."
                    )
                    break

                # Snapshot previo para detectar archivo nuevo
                archivos_antes = {f for f in carpeta_temp.iterdir() if f.is_file()}

                boton_descarga.click()
                escribir_log(
                    f"📌 Descargando '{doc_buscar}' — viaje {numero_viaje} "
                    f"(intento {intento})..."
                )

                archivo_nuevo = esperar_archivo_nuevo(carpeta_temp, archivos_antes)

                if archivo_nuevo:
                    extension      = archivo_nuevo.suffix or ".pdf"
                    nombre_destino = carpeta_viaje / f"{doc_id}_{numero_viaje}{extension}"
                    shutil.move(str(archivo_nuevo), str(nombre_destino))
                    escribir_log(f"✅ '{doc_id}' guardado: {nombre_destino}")
                    ids_descargados.append(doc_id)
                    documentos_descargados += 1
                    descargado = True
                    break
                else:
                    escribir_log(
                        f"⚠️ No apareció archivo de '{doc_buscar}' (intento {intento}). "
                        f"Reintentando en {SLEEP_RETRY}s..."
                    )
                    time.sleep(SLEEP_RETRY)

            except Exception as e:
                escribir_log(
                    f"⚠️ Error al descargar '{doc_buscar}' viaje {numero_viaje} "
                    f"(intento {intento}): {e}"
                )
                time.sleep(SLEEP_RETRY)

        if not descargado:
            escribir_log(
                f"❌ No se pudo descargar '{doc_buscar}' para viaje {numero_viaje} "
                f"tras 2 intentos."
            )

    # Registrar error si faltaron documentos
    ids_requeridos = [d["id"] for d in docs_requeridos]
    faltantes = [d for d in ids_requeridos if d not in ids_descargados]
    if faltantes:
        escribir_log(f"⚠️ Viaje {numero_viaje} — documentos faltantes: {faltantes}")
        if numero_viaje not in viajes_con_error:
            viajes_con_error.append(numero_viaje)

    viajes_procesados += 1

    # ── d) Volver a ventana principal si se abrió nueva ─────────────────────
    if driver.current_window_handle != ventana_principal:
        driver.close()
        driver.switch_to.window(ventana_principal)
        time.sleep(SLEEP_POPUP)

    # ── e) Cerrar ventana de documentos (botón "Regresar") ──────────────────
    cerrado = False
    # Intento 1: por NAME del input (BTN_CANCELAR es el ID interno en AWP)
    try:
        boton_regresar = wait.until(
            EC.element_to_be_clickable((By.NAME, "BTN_CANCELAR"))
        )
        boton_regresar.click()
        cerrado = True
    except:
        pass
    # Intento 2: por texto visible del botón si el ID no funciona
    if not cerrado:
        try:
            boton_regresar = driver.find_element(
                By.XPATH,
                "//input[@value='Regresar'] | //button[contains(text(),'Regresar')] "
                "| //span[contains(text(),'Regresar')]"
            )
            boton_regresar.click()
            cerrado = True
        except:
            pass
    if cerrado:
        escribir_log(f"📌 Ventana de documentos cerrada para viaje {numero_viaje}.")
    else:
        escribir_log(f"⚠️ No se encontró botón Regresar/BTN_CANCELAR para viaje {numero_viaje}.")
    time.sleep(SLEEP_POPUP)

    # ── f) Limpiar buscador y continuar ─────────────────────────────────────
    try:
        buscador = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[aria-controls='TABLE_ProViajes']")
            )
        )
        buscador.clear()
        time.sleep(1)
    except:
        pass

# ============================================================
#  RESUMEN FINAL
# ============================================================

escribir_log(f"\n{'='*60}")
escribir_log("✅ RESUMEN FINAL")
escribir_log(f"   - Viajes procesados:        {viajes_procesados}")
escribir_log(f"   - Documentos descargados:   {documentos_descargados}")
escribir_log(f"   - Viajes con error:         {len(viajes_con_error)}")
if viajes_con_error:
    escribir_log(f"   - Viajes con error (lista): {', '.join(viajes_con_error)}")

# ============================================================
#  LOGOUT Y CIERRE DEL NAVEGADOR
# ============================================================

try:
    escribir_log("📌 Cerrando sesión...")
    driver.find_element(By.ID, "z_BTN_CERRARSESION_IMG").click()
    time.sleep(3)
    escribir_log("✅ Sesión cerrada correctamente.")
except:
    escribir_log("⚠️ No se encontró el botón 'Cerrar Sesión'. Procediendo con cierre de navegador.")

driver.quit()
escribir_log("✅ Navegador cerrado correctamente.")

# Limpiar carpeta temporal si quedó vacía
try:
    if carpeta_temp.exists() and not any(carpeta_temp.iterdir()):
        carpeta_temp.rmdir()
except:
    pass
