import os
import time
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from log_SitCobranza_config import escribir_log  

# 📌 Determinar si es la primera corrida del día (9 AM)
hora_actual = datetime.now().hour

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


# 📌 Ruta del ChromeDriver (ajusta según la ubicación real en tu PC)
service = Service(ChromeDriverManager().install())

# 📌 Inicializar el navegador con Selenium
driver = webdriver.Chrome(service=service, options=options)

# 📌 URL de inicio de sesión de GM Transport
url = "https://www.softwareparatransporte.com/GMTERPV8_WEB/ES/PAGE_CatUsuariosLoginAWP.awp"
driver.get(url)
time.sleep(2)

# Esperar que la página cargue
time.sleep(5)

# Llenar los campos de inicio de sesión
empresa = driver.find_element(By.NAME, "EDT_EMPRESA")
usuario = driver.find_element(By.NAME, "EDT_USUARIO")
contraseña = driver.find_element(By.NAME, "EDT_CONTRASENA")

# Ingresar los datos (Asegúrate de reemplazar por los datos reales)
empresa.send_keys("SLC080722SM6")
usuario.send_keys("DTN")
contraseña.send_keys("RenGax8*")

# Hacer clic en el botón "INICIAR SESIÓN"
boton_login = driver.find_element(By.ID, "BTN_ENTRAR")
boton_login.click()

# Esperar unos segundos para que cargue la nueva página
time.sleep(5)

# 📌 Verificar si aparece popup de Sesion Abierta 
time.sleep(3)

try:
    # Cambiar el foco al alerta emergente
    alert = driver.switch_to.alert

    # Obtener el texto del alerta (para confirmar que es el correcto)
    alert_text = alert.text
    print(f"✅ Se detectó un alerta: {alert_text}")

    # Si el mensaje detectado es el de sesión abierta, hacer clic en "Aceptar" (Sí)
    if "El sistema ha detectado una sesión abierta con este usuario" in alert_text:
        alert.accept()  # Hace clic en "Sí"
        print("✅ Se aceptó el alerta de sesión abierta.")

    else:
        print("ℹ️ No es el alerta esperado.")

except:
    print("ℹ️ No se detectó ningún alerta.")

# 📌 Verificar si aparece el pop-up diario
try:
    time.sleep(8)  # Esperar para asegurarse de que el popup se abra
    # Intentar encontrar el checkbox "No volver a mostrar"
    checkbox_no_mostrar = driver.find_element(By.ID, "CBOX_CHECKBOX1_1")
    
    if not checkbox_no_mostrar.is_selected():  # Si no está marcado, marcarlo
        checkbox_no_mostrar.click()
        escribir_log("✅ Opción 'No volver a mostrar' marcada.")

    # Intentar encontrar y hacer clic en el botón "OK"
    boton_ok = driver.find_element(By.XPATH, "//span[@class='btnvalignmiddle']")
    boton_ok.click()
    
    escribir_log("✅ Pop-up detectado y cerrado correctamente.")
    time.sleep(2)  # Esperar para asegurarse de que el popup se cierre

except:
    escribir_log("ℹ️ No se detectó ningún pop-up. Continuando con el proceso.")

# 📌 1️⃣ Navegar a "Cobranza"
try:
    escribir_log("📌 Navegando a la sección 'Cobranza'...")
    time.sleep(3)
    cobranza = driver.find_element(By.XPATH, "//img[contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'cobranza1.jpg')]")
    cobranza.click()
    time.sleep(3)

    escribir_log("✅ Sección 'Cobranza' abierta correctamente.")
except Exception as e:
    escribir_log(f"❌ Error al acceder a 'Cobranza': {e}", nivel="error")
    driver.quit()
    exit()

# 📌 2️⃣ Hacer clic en "Reportes"
try:
    escribir_log("📌 Navegando a 'Reportes'...")
    reportes = driver.find_element(By.XPATH, "//img[contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'cobranza/reportes1.jpg')]")
    reportes.click()
    time.sleep(3)

    escribir_log("✅ Sección 'Reportes' abierta correctamente.")
except Exception as e:
    escribir_log(f"❌ Error al acceder a 'Reportes': {e}", nivel="error")
    driver.quit()
    exit()

# 📌 3️⃣ Seleccionar "Situación de Cobranza"
try:
    escribir_log("📌 Seleccionando 'Situación de Cobranza'...")
    
    situacion_cobranza = driver.find_element(By.ID, "zrl_10_LINK_REPORTE")
    situacion_cobranza.click()
    time.sleep(3)

    escribir_log("✅ Reporte 'Situación de Cobranza' seleccionado.")
except Exception as e:
    escribir_log(f"❌ Error al seleccionar 'Situación de Cobranza': {e}", nivel="error")
    driver.quit()
    exit()

# 📌 4️⃣ Seleccionar "Convertir a la Moneda"
try:
    escribir_log("📌 Seleccionando 'Convertir a la Moneda' en el combo...")
    time.sleep(1)
    select_moneda = Select(driver.find_element(By.ID, "COMBO_CONSIDERAR"))
    time.sleep(1)
    select_moneda.select_by_value("2")  # Opción "Convertir a la Moneda"

    escribir_log("✅ Opción 'Convertir a la Moneda' seleccionada correctamente.")
except Exception as e:
    escribir_log(f"❌ Error al seleccionar 'Convertir a la Moneda': {e}", nivel="error")

# 📌 5️⃣ Seleccionar "Fecha de Envío"
try:
    escribir_log("📌 Seleccionando 'Fecha de Envío' en el combo...")
    time.sleep(1)
    select_fecha = Select(driver.find_element(By.ID, "COMBO_FECHADE"))
    time.sleep(1)
    select_fecha.select_by_value("2")  # Opción "Fecha de Envío"

    escribir_log("✅ Opción 'Fecha de Envío' seleccionada correctamente.")
except Exception as e:
    escribir_log(f"❌ Error al seleccionar 'Fecha de Envío': {e}", nivel="error")

# 📌 6️⃣ SELECCIONAR "PESOS" EN MONEDA
try:
    escribir_log("📌 Seleccionando 'PESOS' en el combo de Moneda...")
    time.sleep(1)
    select_moneda = Select(driver.find_element(By.ID, "COMBO_CATMONEDAS"))
    time.sleep(1)
    select_moneda.select_by_value("1")  # Opción "PESOS"

    escribir_log("✅ 'Moneda' cambiada a 'PESOS'.")
except Exception as e:
    escribir_log(f"❌ No se pudo cambiar 'Moneda' a 'PESOS': {e}", nivel="error")

time.sleep(2)

# 📌 6️⃣ Generar Reporte
try:
    escribir_log("📌 Generando reporte...")
    time.sleep(1)
    driver.find_element(By.XPATH, "//span[contains(text(), 'Generar')]").click()
    time.sleep(300)  

    escribir_log("✅ Reporte generado correctamente.")
except Exception as e:
    time.sleep(300) 
    escribir_log(f"❌ No se pudo generar el reporte: {e}", nivel="error")

# 📌 7️⃣ Exportar a XLS
try:
    escribir_log("📌 Exportando reporte a XLS...")
    time.sleep(1)
    driver.find_element(By.XPATH, "//span[contains(text(), 'Exportar XLS')]").click()
    time.sleep(10)  

    escribir_log("✅ Reporte exportado correctamente a XLS.")
except Exception as e:
    escribir_log(f"❌ No se pudo exportar a XLS: {e}", nivel="error")

# 📌 8️⃣ Cerrar sesión y cerrar navegador
try:
    escribir_log("📌 Cerrando sesión...")
    driver.find_element(By.ID, "z_BTN_CERRARSESION_IMG").click()
    time.sleep(3)
    escribir_log("✅ Sesión cerrada correctamente.")
except:
    escribir_log("⚠️ No se encontró el botón de 'Cerrar Sesión'. Procediendo con cierre de navegador.")

driver.quit()
escribir_log("✅ Navegador cerrado correctamente.")
