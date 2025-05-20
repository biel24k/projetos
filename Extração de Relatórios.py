from Dados import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import os
import time

# Início do contador de tempo
inicio = time.time()

# Configurações do navegador
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get(url)

# Login no sistema
driver.find_element(*campo_usuario).send_keys(usuario)
driver.find_element(*campo_senha).send_keys(senha)
driver.find_element(*btn_entrar).click()
time.sleep(2)

# Acessar área de relatórios
driver.find_element(*btn_opcao_relatorios).click()
time.sleep(2)

# Preencher intervalo de datas
campo_data_inicio = driver.find_element(*campo_data_inicio_iden)
campo_data_fim = driver.find_element(*campo_data_fim_iden)
data_inicio = datetime(datetime.now().year, datetime.now().month, 1).strftime('%d/%m/%Y')
data_fim = datetime.now().strftime('%d/%m/%Y')
campo_data_inicio.send_keys(data_inicio)
campo_data_fim.send_keys(data_fim)

# Selecionar o tipo de relatório
driver.find_element(*opcao_relatorio).click()
WebDriverWait(driver, 10).until(EC.visibility_of_element_located(opcao_relatorio_atendimento)).click()

# Exportar relatório
driver.find_element(*btn_exportar).click()
WebDriverWait(driver, 10).until(EC.element_to_be_clickable(btn_confirmacao)).click()

# Aguarda o download
time.sleep(3)
driver.quit()

# Localizar o CSV mais recente
arquivos_csv = [f for f in os.listdir(download_path) if f.endswith('.csv')]
caminho_completo = os.path.join(download_path, max(arquivos_csv, key=lambda f: os.path.getmtime(os.path.join(download_path, f))))

# Carregar o CSV
df_csv = pd.read_csv(caminho_completo, delimiter=';', on_bad_lines='skip', encoding='utf-8')

# Conectar ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(os.getenv("CRED_JSON_PATH"), scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url(os.getenv("PLANILHA_URL"))
sheet = spreadsheet.get_worksheet(1)

print("Planilha aberta com sucesso:", spreadsheet.title)

# Cabeçalhos esperados
colunas_chave = [
    "Protocolo", "Cliente", "Ocorrencia", "FilaAtendimento", "Responsavel/Tecnico",
    "DataAbertura", "Status", "Tipo_de_Chamado", "Nome_Fantasia",
    "Razao_Social", "Local_de_Atendimento", "Status_de_Consulta"
]

# Verificar ou criar cabeçalhos
titulos = sheet.row_values(1)
if titulos != colunas_chave or '' in titulos or len(set(titulos)) != len(titulos):
    print("Atualizando cabeçalhos...")
    sheet.update('A1', [colunas_chave])
else:
    print("Cabeçalhos válidos encontrados.")

# Limpar dados anteriores
registros_existentes = sheet.get_all_records(expected_headers=colunas_chave)
num_linhas = len(sheet.get_all_values())
if num_linhas > 1:
    sheet.batch_clear([f'A2:Z{num_linhas}'])
    print(f"{num_linhas - 1} linhas apagadas.")
else:
    print("Planilha já estava limpa.")

# Evitar duplicidade
ids_existentes = {
    ''.join(str(registro.get(col, "")).strip() for col in colunas_chave)
    for registro in registros_existentes
}

novas_linhas = []
for _, row in df_csv.iterrows():
    chave = ''.join(str(row.get(col, "")).strip() for col in colunas_chave)
    if chave not in ids_existentes:
        novas_linhas.append(row)
        ids_existentes.add(chave)

if not novas_linhas:
    print("Nenhuma nova entrada.")
else:
    df_novos = pd.DataFrame(novas_linhas).fillna("")
    for col in ["Responsavel/Tecnico", "Status_de_Consulta"]:
        if col not in df_novos.columns:
            df_novos[col] = ''
    dados = df_novos[colunas_chave].values.tolist()
    sheet.append_rows(dados, value_input_option="USER_ENTERED")
    print(f"{len(dados)} novas entradas adicionadas.")

# Limpar CSV
if os.path.exists(caminho_completo):
    os.remove(caminho_completo)
    print(f"CSV removido: {caminho_completo}")

# Tempo de execução
fim = time.time()
minutos = int((fim - inicio) // 60)
segundos = int((fim - inicio) % 60)
print(f"Execução finalizada em {minutos}m {segundos}s.")
