from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import gspread
import time

from Dados import *

# Início da contagem de tempo
inicio = time.time()

# Autenticação Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("telematica-459812-2dcf92fe61ca.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZFLZR5aTHrtnQR44ofJS8PXHbKbyzr30F5_II_zggaE/edit#gid=0")
sheet = spreadsheet.get_worksheet(1)
valores = sheet.col_values(1)[1:]  # Coluna A (exceto cabeçalho)

# Configuração do Selenium - Headless Chrome
options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Login no sistema
driver.get(url)
driver.find_element(*campo_usuario).send_keys(usuario)
driver.find_element(*campo_senha).send_keys(senha)
driver.find_element(*btn_entrar).click()
driver.execute_script("window.scrollBy(0, 200);")

# Início da consulta
updates = []  # lista de atualizações para batch_update
for i, protocolo in enumerate(valores, start=2):
    print(f"Consultando protocolo: {protocolo}")
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//a[contains(., "Consultar")]'))
        ).click()

        campo = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'))
        )
        campo.clear()
        campo.send_keys(str(protocolo))

        linha = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, f'//tr[td[contains(text(), "{protocolo}")]]'))
        )
        status = linha.find_element(By.XPATH, './td[12]').text.strip()
        responsavel = linha.find_element(By.XPATH, './td[13]').text.strip() or "(sem responsável)"

        print(f"Status: '{status}' | Responsável: '{responsavel}'")
    except Exception as e:
        print(f"Erro na consulta do protocolo {protocolo}: {e}")
        status = "Erro ao buscar status"
        responsavel = "Erro ao buscar responsável"

    # Prepara as atualizações (coluna L: status, coluna E: responsável)
    updates.append({'range': f"L{i}", 'values': [[status]]})
    updates.append({'range': f"E{i}", 'values': [[responsavel]]})

# Atualização em lote
if updates:
    sheet.batch_update(updates)
    print(f"{len(updates)//2} protocolos atualizados com sucesso.")
else:
    print("Nenhuma atualização realizada.")

# Encerra navegador e exibe tempo
driver.quit()
fim = time.time()
minutos = int((fim - inicio) // 60)
segundos = int((fim - inicio) % 60)
print(f"Tempo total de execução: {minutos} minuto(s) e {segundos} segundo(s).")
