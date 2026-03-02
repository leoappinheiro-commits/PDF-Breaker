import os
import shutil
import datetime
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd

# Deixe como None para baixar automaticamente uma versão compatível.
# Só preencha se você realmente precisar fixar manualmente uma versão específica.
driver_version = None


def obter_caminho_chromedriver(versao=None):
    """Garante que o caminho retornado pelo webdriver-manager seja o executável do chromedriver."""
    caminho = ChromeDriverManager(driver_version=versao).install() if versao else ChromeDriverManager().install()
    caminho_path = Path(caminho)

    if caminho_path.is_file() and caminho_path.suffix.lower() == ".exe":
        return str(caminho_path)

    # Alguns ambientes retornam o diretório de cache ou um arquivo que não é executável.
    pasta = caminho_path if caminho_path.is_dir() else caminho_path.parent
    candidatos = list(pasta.rglob("chromedriver.exe"))
    if candidatos:
        return str(candidatos[0])

    raise RuntimeError(
        f"Não foi possível localizar o executável 'chromedriver.exe' em: {pasta}. "
        "Defina o caminho manualmente para evitar o erro WinError 193."
    )


i = 0
Chaves = pd.read_excel(r'C:\Users\leonardopinheiro\Documents\Lista de Chaves.xlsx')

# Configurar as opções do Chrome
chrome_options = Options()
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--safebrowsing-disable-download-protection')
chrome_options.add_argument('--no-sandbox')

# Inicializar o driver do Chrome
# 1) Tenta Selenium Manager (resolve automaticamente o driver correto para o Chrome instalado)
# 2) Em caso de falha, usa webdriver-manager com validação de caminho
try:
    navegador = webdriver.Chrome(options=chrome_options)
except Exception:
    driver_path = obter_caminho_chromedriver(driver_version)
    service = Service(executable_path=driver_path)
    navegador = webdriver.Chrome(service=service, options=chrome_options)

# Acessar a página desejada
navegador.get("https://consultadanfe.com/#")

# Diretório de downloads
downloads_dir = r'C:\Users\leonardopinheiro\Downloads'


# Realizar qualquer interação adicional no navegador, se necessário
for i, chave in Chaves.iterrows():
    campo_nome = navegador.find_element(By.ID, 'chave').send_keys(chave['Chaves_Coluna'])
    time.sleep(5)
    campo_nome = navegador.find_element(By.CLASS_NAME, 'g-recaptcha').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[2]/div/div[3]/p[2]/a').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[3]/a').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.ID, 'chave').clear()
    time.sleep(5)

    # Mover o arquivo baixado para a pasta XML
    arquivo_baixado = max([os.path.join(downloads_dir, f) for f in os.listdir(downloads_dir)], key=os.path.getctime)
    data_atual = datetime.datetime.now().strftime("%d-%m-%Y")
    xml_dir = os.path.join(downloads_dir, "XML " + data_atual)
    os.makedirs(xml_dir, exist_ok=True)
    shutil.move(arquivo_baixado, xml_dir)

# Fechar o navegador
navegador.quit()
