from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import os
from datetime import datetime
from webdriver_manager.chrome import ChromeDriverManager
from Dados import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
inicio = time.time()


# Autorização do Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("telematica-459812-2dcf92fe61ca.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZFLZR5aTHrtnQR44ofJS8PXHbKbyzr30F5_II_zggaE/edit?gid=0#gid=0")
sheet = spreadsheet.get_worksheet(0)
valores = sheet.col_values(1)[1:]  # Pega os valores da coluna A a partir da linha 2

# Configurando o Selenium
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

     # Acessar o site
driver.get(url)

     # Login
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[1]/input').send_keys(usuario)
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[2]/input').send_keys(senha)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
driver.execute_script("window.scrollBy(0, 200);")

     # Localiza e clica no Consultar
consultar_link = WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.XPATH, '//a[contains(., "Consultar")]'))
)
consultar_link.click()

#                 LIMITANDO A CONSULTA PARA APENAS OS 21 PRIMEIROS
for i, Protocolo in enumerate(valores[:3], start=404):  # começa na linha 2
    print(f"Consultando valor: {Protocolo}")

    # Localiza e preenche o campo Procurar
    campo = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'))
    )
    campo.click()
    campo.clear()
    campo.send_keys(str(Protocolo))

    # Tenta capturar o status
    try:
        # Aguarda a linha com o protocolo aparecer
        linha = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, f'//tr[td[contains(text(), "{Protocolo}")]]'))
        )
        status_element = linha.find_element(By.XPATH, './td[12]')
        status = status_element.text.strip()
        print(f"Status extraído: '{status}' para o protocolo {Protocolo}") 
    except Exception as e:
        status = "Status não encontrado"
        print(f"Erro ao buscar status do protocolo {Protocolo}: {e}")

    try:
        responsavel_element = linha.find_element(By.XPATH, './td[13]')
        responsavel = responsavel_element.text.strip()

        if not responsavel:
                responsavel = "(sem responsável)"
                print(f"Nenhum responsável atribuído para o protocolo {Protocolo}, marcando como encerrado.")
        else:
                print(f"Responsável extraído: '{responsavel}' para o protocolo {Protocolo}")

    except Exception as e:
        responsavel = "Erro ao buscar responsável"
        print(f"Erro ao buscar responsável do protocolo {Protocolo}: {e}")

    # Atualiza a coluna L com o status e a colune E com o Técnico/Responsavel
    cell_range = f"L{i}"
    sheet.update_acell(cell_range, status)

    cell_responsavel = f"E{i}"
    sheet.update_acell(cell_responsavel, responsavel)


fim = time.time()
duracao = fim - inicio
minutos = int(duracao // 60)
segundos = int(duracao % 60)
print(f"Tempo total de execução: {minutos} minuto(s) e {segundos} segundo(s).")
