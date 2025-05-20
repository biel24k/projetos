import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# CONFIGURAÇÕES
GOOGLE_SHEETS_URL = os.getenv("GOOGLE_SHEETS_URL")
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE", "credentials.json")
USUARIO = os.getenv("PORTAL_USUARIO")
SENHA = os.getenv("PORTAL_SENHA")
URL = os.getenv("PORTAL_URL")

def autorizar_planilha():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(GOOGLE_SHEETS_URL).get_worksheet(1)

def configurar_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def login(driver):
    driver.get(URL)
    driver.find_element(By.XPATH, '//input[@type="text"]').send_keys(USUARIO)
    driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(SENHA)
    driver.find_element(By.XPATH, '//button[@type="submit"]').click()
    driver.execute_script("window.scrollBy(0, 200);")

def consultar_protocolo(driver, protocolo):
    WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//a[contains(., "Consultar")]'))
    ).click()

    campo = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'))
    )
    campo.clear()
    campo.send_keys(str(protocolo))

    try:
        linha = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, f'//tr[td[contains(text(), "{protocolo}")]]'))
        )
        status = linha.find_element(By.XPATH, './td[12]').text.strip()
        responsavel = linha.find_element(By.XPATH, './td[13]').text.strip() or "(sem responsável)"
    except Exception:
        status = "Status não encontrado"
        responsavel = "Erro ao buscar responsável"
    
    return status, responsavel

def main():
    inicio = time.time()

    sheet = autorizar_planilha()
    protocolos = sheet.col_values(1)[1:]

    driver = configurar_driver()
    login(driver)

    for i, protocolo in enumerate(protocolos, start=2):
        print(f"Consultando protocolo: {protocolo}")
        status, responsavel = consultar_protocolo(driver, protocolo)

        print(f"Status: {status}, Responsável: {responsavel}")
        sheet.update_acell(f"L{i}", status)
        sheet.update_acell(f"E{i}", responsavel)

    driver.quit()

    fim = time.time()
    duracao = fim - inicio
    print(f"Tempo total de execução: {int(duracao // 60)}m {int(duracao % 60)}s")

if __name__ == "__main__":
    main()
