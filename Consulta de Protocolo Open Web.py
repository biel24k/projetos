import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials
import gspread

from Dados import *

inicio = time.time()

# === 1. Autenticação Google Sheets ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("telematica-459812-2dcf92fe61ca.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZFLZR5aTHrtnQR44ofJS8PXHbKbyzr30F5_II_zggaE/edit#gid=0")
sheet = spreadsheet.get_worksheet(0)
valores = sheet.col_values(1)[1:]  # Coluna A (sem cabeçalho)

# === 2. Inicialização do navegador (Chrome) ===
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# === 3. Login no sistema ===
driver.get(url)
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[1]/input').send_keys(usuario)
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[2]/input').send_keys(senha)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
driver.execute_script("window.scrollBy(0, 200);")

# === 4. Acesso à tela de consulta ===
WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.XPATH, '//a[contains(., "Consultar")]'))
).click()

# === 5. Consulta de protocolos ===
updates = []  # lista de atualizações para batch_update
for i, protocolo in enumerate(valores[:3], start=404):
    print(f"Consultando protocolo: {protocolo}")
    try:
        # Preencher campo de busca
        campo = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'))
        )
        campo.clear()
        campo.send_keys(str(protocolo))

        # Localiza linha do protocolo
        linha = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, f'//tr[td[contains(text(), "{protocolo}")]]'))
        )

        status = linha.find_element(By.XPATH, './td[12]').text.strip()
        responsavel = linha.find_element(By.XPATH, './td[13]').text.strip()

        if not responsavel:
            responsavel = "(sem responsável)"
            print(f"Protocolo {protocolo}: status='{status}' | sem responsável.")
        else:
            print(f"Protocolo {protocolo}: status='{status}' | responsável='{responsavel}'.")

    except Exception as e:
        print(f"Erro ao consultar protocolo {protocolo}: {e}")
        status = "Status não encontrado"
        responsavel = "Erro ao buscar responsável"

    # Adiciona atualizações em lote
    updates.append({'range': f"L{i}", 'values': [[status]]})
    updates.append({'range': f"E{i}", 'values': [[responsavel]]})

# === 6. Atualização em lote da planilha ===
if updates:
    sheet.batch_update(updates)
    print(f"{len(updates)//2} protocolos atualizados.")
else:
    print("Nenhuma atualização realizada.")

# === 7. Encerramento e tempo de execução ===
driver.quit()
fim = time.time()
minutos = int((fim - inicio) // 60)
segundos = int((fim - inicio) % 60)
print(f"Tempo total de execução: {minutos} minuto(s) e {segundos} segundo(s).")
