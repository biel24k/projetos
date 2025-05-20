from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import gspread
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.chrome.service import Service
from Dados import *
import pandas as pd
import time
import os
inicio = time.time()

# === CONFIGURAR PASTA DE DOWNLOAD ===
download_path = os.path.join(os.path.expanduser("~"), "Downloads")
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})

# === INICIAR NAVEGADOR ===
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get("https://portal.epontodespacho.com.br/login?returnUrl=%2Fhome")

# === LOGIN ===
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[1]/input').send_keys(usuario)
driver.find_element(By.XPATH, '/html/body/app-root/app-layout-login/div/app-login/div/div/div[2]/form/div[2]/input').send_keys(senha)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
time.sleep(2)

# === ACESSAR RELATÓRIOS ===
driver.find_element(By.XPATH, "//a[@href='/relatorio/consultar']").click()
time.sleep(2)

# === DEFINIR DATAS ===
data_inicio = datetime(datetime.now().year, datetime.now().month, 1).strftime('%d/%m/%Y')
campo_data_inicio = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='DataInicial']")
campo_data_inicio.click()
campo_data_inicio.send_keys(data_inicio)
data_fim = datetime.now().strftime('%d/%m/%Y')


# # MÊS ANTERIOR VVVVVVVVVVVVVVVV
# from dateutil.relativedelta import relativedelta
# # === DEFINIR DATAS ===
# primeiro_dia_mes_anterior = datetime.now().replace(day=1) - relativedelta(months=1)
# ultimo_dia_mes_anterior = datetime.now().replace(day=1) - relativedelta(days=1)

# data_inicio = primeiro_dia_mes_anterior.strftime('%d/%m/%Y')
# data_fim = ultimo_dia_mes_anterior.strftime('%d/%m/%Y')


# campo_data_inicio = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='DataInicial']")
# campo_data_inicio.click()
# campo_data_inicio.send_keys(data_inicio)
campo_data_fim = driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='DataFinal']")
campo_data_fim.click()
campo_data_fim.send_keys(data_fim)

# === SELECIONAR RELATÓRIO E EXPORTAR ===
driver.find_element(By.CSS_SELECTOR, ".select2-selection").click()
WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, "//li[text()='RELATÓRIO DE ATENDIMENTO']"))
).click()

driver.find_element(By.XPATH, "//button[text()='Exportar relatório em CSV']").click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[text()=' Sim ']"))
).click()

# === AGUARDAR DOWNLOAD E FECHAR NAVEGADOR ===
time.sleep(3)
driver.quit()

# === LOCALIZAR CSV MAIS RECENTE ===
arquivos_csv = [f for f in os.listdir(download_path) if f.endswith('.csv')]
caminho_completo = os.path.join(
    download_path,
    max(arquivos_csv, key=lambda f: os.path.getmtime(os.path.join(download_path, f)))
)

# === CARREGAR CSV ===
df_csv = pd.read_csv(caminho_completo, delimiter=';', on_bad_lines='skip', encoding='utf-8')

# === AUTENTICAÇÃO GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("telematica-459812-2dcf92fe61ca.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ZFLZR5aTHrtnQR44ofJS8PXHbKbyzr30F5_II_zggaE/edit?gid=0#gid=0")
# TRECHO RESPONSÁVEL PELA ABA DA PLANILHA
sheet = spreadsheet.get_worksheet(1)
print("Planilha aberta com sucesso:", spreadsheet.title)

# === CONFIGURAÇÃO DE CABEÇALHOS ===
colunas_chave = [
    "Protocolo", "Cliente", "Ocorrencia", "FilaAtendimento", "Responsavel/Tecnico",
    "DataAbertura", "Status", "Tipo_de_Chamado", "Nome_Fantasia",
    "Razao_Social", "Local_de_Atendimento", "Status_de_Consulta"
]

titulos = sheet.row_values(1)
if titulos != colunas_chave or '' in titulos or len(set(titulos)) != len(titulos):
    print("Atualizando cabeçalhos...")
    sheet.update('A1', [colunas_chave])
else:
    print("Cabeçalhos válidos encontrados.")

# === OBTER DADOS EXISTENTES ===
registros_existentes = sheet.get_all_records(expected_headers=colunas_chave)
num_linhas = len(sheet.get_all_values())
if num_linhas > 1:
    sheet.batch_clear([f'A2:Z{num_linhas}'])
    print(f"{num_linhas - 1} linhas apagadas da planilha.")
else:
    print("Planilha já limpa.")

# === CRIAR CONJUNTO DE IDS EXISTENTES ===
ids_existentes = {
    ''.join(str(registro.get(col, "")).strip() for col in colunas_chave)
    for registro in registros_existentes
}

# === FILTRAR NOVAS LINHAS ===
novas_linhas = []
for _, row in df_csv.iterrows():
    chave = ''.join(str(row.get(col, "")).strip() for col in colunas_chave)
    if chave not in ids_existentes:
        novas_linhas.append(row)
        ids_existentes.add(chave)

if not novas_linhas:
    print("Nenhuma nova entrada para adicionar.")
else:
    df_novos = pd.DataFrame(novas_linhas).fillna("")

    # Garantir coluna 'Status_de_Consulta'
    if 'Status_de_Consulta' not in df_novos.columns:
        df_novos['Status_de_Consulta'] = ''
    # Garantir coluna 'Responsavel/Tecnico'
    if 'Responsavel/Tecnico' not in df_novos.columns:
        df_novos['Responsavel/Tecnico'] = ''

    # Organizar e adicionar
    dados = df_novos[colunas_chave].values.tolist()
    sheet.append_rows(dados, value_input_option="USER_ENTERED")
    print(f"Novas entradas adicionadas: {len(dados)}")

# === EXCLUIR CSV USADO ===
if os.path.exists(caminho_completo):
    os.remove(caminho_completo)
    print(f"Arquivo CSV removido: {caminho_completo}")
else:
    print("CSV já havia sido removido ou não encontrado.")



fim = time.time()
duracao = fim - inicio
minutos = int(duracao // 60)
segundos = int(duracao % 60)
print(f"Tempo total de execução: {minutos} minuto(s) e {segundos} segundo(s).")