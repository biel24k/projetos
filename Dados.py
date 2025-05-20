from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import gspread
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
from dotenv import load_dotenv

# Carrega variáveis de ambiente do .env
load_dotenv()

# === CONFIGURAÇÃO DA PASTA DE DOWNLOAD ===
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# === URL DO SISTEMA ===
url = "https://portal.epontodespacho.com.br/login?returnUrl=%2Fhome"

# === LOGIN (via variáveis de ambiente) ===
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

campo_usuario = (By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[1]/input')
campo_senha = (By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[2]/input')
btn_entrar = (By.XPATH, '//button[@type="submit"]')

# === NAVEGAÇÃO ===
btn_opcao_relatorios = (By.XPATH, "//a[@href='/relatorio/consultar']")

# === CAMPOS DE DATA ===
campo_data_inicio_iden = (By.CSS_SELECTOR, "input[formcontrolname='DataInicial']")
campo_data_fim_iden = (By.CSS_SELECTOR, "input[formcontrolname='DataFinal']")

# === RELATÓRIO E EXPORTAÇÃO ===
opcao_relatorio = (By.CSS_SELECTOR, ".select2-selection")
opcao_relatorio_atendimento = (By.XPATH, "//li[text()='RELATÓRIO DE ATENDIMENTO']")
btn_exportar = (By.XPATH, "//button[text()='Exportar relatório em CSV']")
btn_confirmacao = (By.XPATH, "//button[text()=' Sim ']")
