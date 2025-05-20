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
from selenium.webdriver.chrome.options import Options
inicio = time.time()



# Autorização do Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("telematica-459812-2dcf92fe61ca.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZFLZR5aTHrtnQR44ofJS8PXHbKbyzr30F5_II_zggaE/edit?gid=0#gid=0")

# SELECIONA A ABA DA PLANILHA
sheet = spreadsheet.get_worksheet(1)
valores = sheet.col_values(1)[1:]  # PEGA OS VALORES DA COLUNA "A" A PARTIR DA LINHA 2

# CONFIGURAÇÃO PARA RODAR EM SEGUNDO PLANO (CLOSE WEB - headless)
options = Options()
options.add_argument("--headless")  # modo invisível
options.add_argument("--disable-gpu")  # recomendação no modo headless
options.add_argument("--window-size=1920,1080")  # previne bugs de layout
options.add_argument("--no-sandbox")  # útil em ambientes Linux
options.add_argument("--disable-dev-shm-usage")  # melhora estabilidade no headless
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico, options=options)

 # ACESSAR O SITE
driver.get(url)

# LOGIN
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[1]/input').send_keys(usuario)
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[2]/input').send_keys(senha)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
driver.execute_script("window.scrollBy(0, 200);")

# LOCALIZA E CLICA NO CONSULTAR
# LIMITANDO A CONSULTA PARA APENAS OS 21 PRIMEIROS
# for i, Protocolo in enumerate(valores[:100], start=2):
for i, Protocolo in enumerate(valores, start=2):
    print(f"Consultando valor: {Protocolo}")
    consultar_link = WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.XPATH, '//a[contains(., "Consultar")]'))
    )
    consultar_link.click()

    # LOCALIZA E PREENCHE O CAMPO "PROCURAR"
    campo = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="search"]'))
    )
    campo.click()
    campo.clear()
    campo.send_keys(str(Protocolo))

    # TENTA CAPTURAR O STATUS
    try:
        # AGUARDA A LINHA COM O PROTOCOLO APARECER
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

    # ATUALIZA A COLUNA "L" COM O status E A COLUNA "E" COM Técnico/Responsavel
    cell_range = f"L{i}"
    sheet.update_acell(cell_range, status)
    cell_responsavel = f"E{i}"
    sheet.update_acell(cell_responsavel, responsavel)


# CALCULA O TEMPO GASTO NO SCRIPT
fim = time.time()
duracao = fim - inicio
minutos = int(duracao // 60)
segundos = int(duracao % 60)
print(f"Tempo total de execução: {minutos} minuto(s) e {segundos} segundo(s).")
