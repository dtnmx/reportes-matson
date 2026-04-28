#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MATSON_DescargarDocumentos.py
Descarga automática de documentos (CI, BOL, POD) desde el sistema GM
para el proceso de carga al portal MATSON.

Basado en Sit_Cobranza_DescargarReporte.py — reutiliza login, alert,
pop-up diario, logout y sistema de logs.
"""

import os
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
ARCHIVO_EXCEL        = "viajes_matson.xlsx"   # Ruta al Excel de viajes
COLUMNA_NUMERO_VIAJE = "Numero_Viaje"          # Columna con el número de viaje
COLUMNA_TIPO_DESTINO = "Tipo_Destino"          # Columna con el tipo de destino

# --- Carpeta de salida ---
FECHA_HOY    = datetime.now().strftime("%Y%m%d")
CARPETA_BASE = f"MATSON_Docs_{FECHA_HOY}"

# --- Documentos a descargar por tipo de destino ---
#     Las claves deben estar en MAYÚSCULAS; el valor es la lista de tipos de doc.
DOCS_POR_TIPO = {
    "LG":    ["CI"],
    "USA":   ["CI", "BOL", "POD"],
    "TODOS": ["CI", "BOL"],        # Valor por defecto cuando Tipo_Destino no es LG ni USA
}

# --- Tiempos de espera (segundos) ---
TIMEOUT_ESPERA  = 10   # WebDriverWait timeout general
SLEEP_LOGIN     = 5    # Tras clic de login y tras cerrar alert
SLEEP_MODULO    = 3    # Tras navegar entre módulos/secciones del menú
SLEEP_FILTRO    = 2    # Tras escribir en el buscador de la tabla de viajes
SLEEP_POPUP     = 3    # Tras abrir o cerrar la ventana de documentos adjuntos
SLEEP_DESCARGA  = 10   # Segundos máximos para que aparezca el archivo descargado
SLEEP_RETRY     = 3    # Espera antes de reintentar una descarga fallida

# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================

def esperar_archivo_nuevo(carpeta: Path, archivos_antes: set, timeout: int = SLEEP_DESCARGA) -> "Path | None":
    """Bloquea hasta que aparezca un archivo nuevo y completo (no .crdownload) en carpeta."""
    fin = time.time() + timeout
    while time.time() < fin:
        ahora = {f for f in carpeta.iterdir() if f.is_file()}
        nuevos = ahora - archivos_antes
        completados = [f for f in nuevos if f.suffix.lower() != ".crdownload"]
        if completados:
            # Devolver el más reciente por fecha de modificación
            return sorted(completados, key=lambda f: f.stat().st_mtime)[-1]
        time.sleep(1)
    return None


def docs_para_viaje(tipo_destino: str) -> list:
    """Retorna la lista de tipos de documento a descargar según el tipo de destino."""
    clave = tipo_destino.upper().strip() if isinstance(tipo_destino, str) else "TODOS"
    return DOCS_POR_TIPO.get(clave, DOCS_POR_TIPO["TODOS"])


def fila_contiene_tipo_doc(texto_fila: str, tipo_doc: str) -> bool:
    """
    Verifica si el texto de una fila de documentos corresponde al tipo buscado.
    Compara en mayúsculas; por ejemplo 'CI', 'BOL', 'POD' deben aparecer en el texto.
    Ajusta esta función si el sistema usa nombres completos como 'CARTA PORTE', etc.
    """
    return tipo_doc.upper() in texto_fila.upper()


# ============================================================
#  INICIALIZACIÓN DE SELENIUM
# ============================================================

hora_actual = datetime.now().hour

# Crear carpeta de salida y carpeta temporal de descargas
carpeta_salida = Path(CARPETA_BASE)
carpeta_salida.mkdir(exist_ok=True)
carpeta_temp   = carpeta_salida / "_temp_downloads"
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

# Configurar carpeta de descarga automática (evita cuadros de diálogo de guardar)
prefs = {
    "download.default_directory":    str(carpeta_temp.resolve()),
    "download.prompt_for_download":  False,
    "download.directory_upgrade":    True,
    "plugins.always_open_pdf_externally": True,
    "safebrowsing.enabled":          True,
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
time.sleep(5)  # Esperar que la página cargue completamente

driver.find_element(By.NAME, "EDT_EMPRESA").send_keys(GM_EMPRESA)
driver.find_element(By.NAME, "EDT_USUARIO").send_keys(GM_USUARIO)
driver.find_element(By.NAME, "EDT_CONTRASENA").send_keys(GM_CONTRASENA)
driver.find_element(By.ID, "BTN_ENTRAR").click()
time.sleep(SLEEP_LOGIN)

# --- Manejo del alert de "sesión abierta" ---
time.sleep(3)
try:
    alert = driver.switch_to.alert
    alert_text = alert.text
    escribir_log(f"✅ Alert detectado: {alert_text}")
    if "El sistema ha detectado una sesión abierta con este usuario" in alert_text:
        alert.accept()
        escribir_log("✅ Alert de sesión abierta aceptado.")
    else:
        escribir_log("ℹ️ Alert desconocido, aceptando de todas formas.")
        alert.accept()
except:
    escribir_log("ℹ️ No se detectó alert de sesión abierta.")

# --- Manejo del pop-up diario ---
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
#  CARGAR EXCEL DE VIAJES
# ============================================================

escribir_log(f"📌 Cargando archivo Excel: {ARCHIVO_EXCEL}")

try:
    df = pd.read_excel(ARCHIVO_EXCEL, dtype={COLUMNA_NUMERO_VIAJE: str})
except FileNotFoundError:
    escribir_log(f"❌ No se encontró el archivo: {ARCHIVO_EXCEL}", nivel="error")
    driver.quit()
    sys.exit(1)
except Exception as e:
    escribir_log(f"❌ Error al leer el Excel: {e}", nivel="error")
    driver.quit()
    sys.exit(1)

if COLUMNA_NUMERO_VIAJE not in df.columns:
    escribir_log(f"❌ El Excel no contiene la columna '{COLUMNA_NUMERO_VIAJE}'.", nivel="error")
    driver.quit()
    sys.exit(1)

if COLUMNA_TIPO_DESTINO not in df.columns:
    escribir_log(
        f"⚠️ El Excel no contiene la columna '{COLUMNA_TIPO_DESTINO}'. "
        "Se asumirá tipo 'TODOS' (CI + BOL) para todos los viajes."
    )
    df[COLUMNA_TIPO_DESTINO] = "TODOS"

viajes_df = df[[COLUMNA_NUMERO_VIAJE, COLUMNA_TIPO_DESTINO]].dropna(subset=[COLUMNA_NUMERO_VIAJE]).copy()
escribir_log(f"✅ {len(viajes_df)} viajes cargados del Excel.")

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

viajes_procesados    = 0
documentos_descargados = 0
viajes_con_error     = []

for _, fila_excel in viajes_df.iterrows():
    numero_viaje    = str(fila_excel[COLUMNA_NUMERO_VIAJE]).strip()
    tipo_destino    = str(fila_excel[COLUMNA_TIPO_DESTINO]).strip()
    docs_requeridos = docs_para_viaje(tipo_destino)

    escribir_log(f"\n{'='*60}")
    escribir_log(
        f"📌 Procesando viaje: {numero_viaje} | "
        f"Tipo destino: {tipo_destino} | Docs a descargar: {docs_requeridos}"
    )

    # Crear carpeta destino del viaje
    carpeta_viaje = carpeta_salida / numero_viaje
    carpeta_viaje.mkdir(exist_ok=True)

    # ---- a) Buscar el viaje en la tabla ----
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
        escribir_log(f"⚠️ No se pudo acceder al buscador de viajes: {e}")
        viajes_con_error.append(numero_viaje)
        continue

    # ---- b) Clic en el ícono azul de documentos adjuntos (clip) ----
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
            f"⚠️ Viaje {numero_viaje} no aparece en la tabla o no tiene ícono de documentos: {e}"
        )
        viajes_con_error.append(numero_viaje)
        # Limpiar buscador antes de continuar
        try:
            buscador = driver.find_element(
                By.CSS_SELECTOR, "input[aria-controls='TABLE_ProViajes']"
            )
            buscador.clear()
            time.sleep(1)
        except:
            pass
        continue

    # Si el popup abre una nueva ventana del navegador, cambiar el foco a ella.
    # Si la ventana de docs es un panel inline en la misma página, este bloque
    # no afectará nada (driver.window_handles tendrá solo 1 handle).
    ventana_principal = driver.current_window_handle
    ventanas_abiertas = driver.window_handles
    if len(ventanas_abiertas) > 1:
        for ventana in ventanas_abiertas:
            if ventana != ventana_principal:
                driver.switch_to.window(ventana)
                break

    # ---- c) Descargar documentos según tipo de destino ----
    docs_descargados_viaje = []

    for tipo_doc in docs_requeridos:
        descargado = False

        for intento in range(1, 3):  # Máximo 2 intentos por documento
            try:
                # Buscar todas las filas que contienen un botón de descarga (span.btnvalignmiddle).
                # NOTA: Si la ventana de docs usa un contenedor/iframe específico, ajusta el
                # XPATH para buscar dentro de ese contenedor, ej:
                #   "//div[@id='ID_CONTENEDOR_DOCS']//tr[.//span[@class='btnvalignmiddle']]"
                filas_doc = driver.find_elements(
                    By.XPATH,
                    "//tr[.//span[@class='btnvalignmiddle']]"
                )

                if not filas_doc:
                    escribir_log(
                        f"⚠️ No se encontraron filas de documentos en la ventana "
                        f"para viaje {numero_viaje}."
                    )
                    break

                boton_descarga = None
                for fila in filas_doc:
                    texto_fila = fila.text
                    if fila_contiene_tipo_doc(texto_fila, tipo_doc):
                        try:
                            boton_descarga = fila.find_element(
                                By.XPATH, ".//span[@class='btnvalignmiddle']"
                            )
                        except:
                            pass
                        break

                if boton_descarga is None:
                    escribir_log(
                        f"⚠️ Documento '{tipo_doc}' no encontrado en la ventana "
                        f"del viaje {numero_viaje}."
                    )
                    break

                # Tomar snapshot de archivos en carpeta temporal antes de descargar
                archivos_antes = {f for f in carpeta_temp.iterdir() if f.is_file()}

                boton_descarga.click()
                escribir_log(
                    f"📌 Descargando '{tipo_doc}' del viaje {numero_viaje} "
                    f"(intento {intento})..."
                )

                # Esperar a que aparezca el archivo descargado
                archivo_nuevo = esperar_archivo_nuevo(carpeta_temp, archivos_antes)

                if archivo_nuevo:
                    extension      = archivo_nuevo.suffix if archivo_nuevo.suffix else ".pdf"
                    nombre_destino = carpeta_viaje / f"{tipo_doc}_{numero_viaje}{extension}"
                    shutil.move(str(archivo_nuevo), str(nombre_destino))
                    escribir_log(f"✅ '{tipo_doc}' guardado como: {nombre_destino}")
                    docs_descargados_viaje.append(tipo_doc)
                    documentos_descargados += 1
                    descargado = True
                    break
                else:
                    escribir_log(
                        f"⚠️ No apareció archivo de '{tipo_doc}' en el intento {intento}. "
                        f"Esperando {SLEEP_RETRY}s antes de reintentar..."
                    )
                    time.sleep(SLEEP_RETRY)

            except Exception as e:
                escribir_log(
                    f"⚠️ Error al descargar '{tipo_doc}' para viaje {numero_viaje} "
                    f"(intento {intento}): {e}"
                )
                time.sleep(SLEEP_RETRY)

        if not descargado:
            escribir_log(
                f"❌ No se pudo descargar '{tipo_doc}' para viaje {numero_viaje} "
                f"tras 2 intentos."
            )

    # Registrar viaje con error si faltaron documentos
    faltantes = [d for d in docs_requeridos if d not in docs_descargados_viaje]
    if faltantes:
        escribir_log(
            f"⚠️ Viaje {numero_viaje}: documentos NO descargados: {faltantes}"
        )
        if numero_viaje not in viajes_con_error:
            viajes_con_error.append(numero_viaje)

    viajes_procesados += 1

    # ---- d) Regresar a ventana principal si se abrió nueva ventana ----
    if driver.current_window_handle != ventana_principal:
        driver.close()
        driver.switch_to.window(ventana_principal)
        time.sleep(SLEEP_POPUP)

    # ---- e) Cerrar ventana de documentos (BTN_CANCELAR) ----
    try:
        boton_cancelar = wait.until(
            EC.element_to_be_clickable((By.NAME, "BTN_CANCELAR"))
        )
        boton_cancelar.click()
        time.sleep(SLEEP_POPUP)
        escribir_log(f"📌 Ventana de documentos cerrada para viaje {numero_viaje}.")
    except Exception as e:
        escribir_log(f"⚠️ No se encontró BTN_CANCELAR para viaje {numero_viaje}: {e}")

    # ---- f) Limpiar buscador y esperar tabla ----
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

# Eliminar carpeta temporal si quedó vacía
try:
    if carpeta_temp.exists() and not any(carpeta_temp.iterdir()):
        carpeta_temp.rmdir()
except:
    pass
