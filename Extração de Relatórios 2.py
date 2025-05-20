# Configurações iniciais e variáveis de ambiente
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
GOOGLE_SHEETS_URL = os.getenv("GOOGLE_SHEETS_URL")
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE", "credentials.json")
USUARIO = os.getenv("PORTAL_USUARIO")
SENHA = os.getenv("PORTAL_SENHA")
PORTAL_URL = os.getenv("PORTAL_URL")

# Prepara o navegador com diretório de download automático
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
    "directory_upgrade": True
})
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get(PORTAL_URL)

# Faz login no portal com as credenciais
driver.find_element(By.XPATH, '//input[@formcontrolname="Login"]').send_keys(USUARIO)
driver.find_element(By.XPATH, '//input[@formcontrolname="Senha"]').send_keys(SENHA)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
time.sleep(2)

# Navega até a área de relatórios
driver.find_element(By.XPATH, "//a[@href='/relatorio/consultar']").click()
time.sleep(2)

# Preenche os campos de data com o início do mês até hoje
hoje = datetime.now()
data_inicio = datetime(hoje.year, hoje.month, 1).strftime('%d/%m/%Y')
data_fim = hoje.strftime('%d/%m/%Y')

driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='DataInicial']").send_keys(data_inicio)
driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='DataFinal']").send_keys(data_fim)

# Escolhe o tipo de relatório e solicita exportação em CSV
driver.find_element(By.CSS_SELECTOR, ".select2-selection").click()
WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, "//li[text()='RELATÓRIO DE ATENDIMENTO']"))
).click()
driver.find_element(By.XPATH, "//button[text()='Exportar relatório em CSV']").click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[text()=' Sim ']"))
).click()

# Aguarda um tempo para download e encerra o navegador
time.sleep(3)
driver.quit()

# Busca o arquivo CSV mais recente na pasta de downloads
csv_arquivos = [f for f in os.listdir(DOWNLOAD_PATH) if f.endswith('.csv')]
csv_path = os.path.join(DOWNLOAD_PATH, max(csv_arquivos, key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_PATH, f))))

# Lê o conteúdo do CSV
df_csv = pd.read_csv(csv_path, delimiter=';', on_bad_lines='skip', encoding='utf-8')

# Conecta à planilha do Google usando as credenciais de serviço
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url(GOOGLE_SHEETS_URL).get_worksheet(1)

# Define os nomes das colunas que queremos manter na planilha
colunas_chave = [
    "Protocolo", "Cliente", "Ocorrencia", "FilaAtendimento", "Responsavel/Tecnico",
    "DataAbertura", "Status", "Tipo_de_Chamado", "Nome_Fantasia",
    "Razao_Social", "Local_de_Atendimento", "Status_de_Consulta"
]

# Garante que os cabeçalhos na planilha estejam corretos
titulos = sheet.row_values(1)
if titulos != colunas_chave or '' in titulos or len(set(titulos)) != len(titulos):
    sheet.update('A1', [colunas_chave])
    print("Cabeçalhos atualizados.")
else:
    print("Cabeçalhos corretos já existentes.")

# Limpa o conteúdo antigo da planilha, se houver
num_linhas = len(sheet.get_all_values())
if num_linhas > 1:
    sheet.batch_clear([f'A2:Z{num_linhas}'])
    print(f"{num_linhas - 1} linhas removidas da planilha.")
else:
    print("Nenhuma linha para limpar.")

# Remove duplicatas e prepara os dados que ainda não foram adicionados
ids_existentes = set()
novas_linhas = []

for _, row in df_csv.iterrows():
    chave = ''.join(str(row.get(col, "")).strip() for col in colunas_chave)
    if chave not in ids_existentes:
        ids_existentes.add(chave)
        novas_linhas.append([row.get(col, "").strip() for col in colunas_chave])

# Adiciona os dados novos na planilha
if novas_linhas:
    sheet.append_rows(novas_linhas, value_input_option="USER_ENTERED")
    print(f"{len(novas_linhas)} novas entradas adicionadas.")
else:
    print("Nenhuma nova entrada para adicionar.")

# Exclui o arquivo CSV usado para liberar espaço
try:
    os.remove(csv_path)
    print(f"CSV removido: {csv_path}")
except FileNotFoundError:
    print("CSV já removido ou não encontrado.")

# Exibe o tempo total de execução do script
fim = time.time()
print(f"Tempo total: {int((fim - inicio) // 60)} minuto(s) e {int((fim - inicio) % 60)} segundo(s).")
