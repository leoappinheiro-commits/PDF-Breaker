import os
import shutil
import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager

driver_version = "119.0.6045.106"  # Substitua pelo número da versão desejada


i = 0
Chaves = pd.read_excel(r'C:\Users\leonardopinheiro\Documents\Lista de Chaves.xlsx')

# Configurar as opções do Chrome
chrome_options = Options()
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--safebrowsing-disable-download-protection')
chrome_options.add_argument('--no-sandbox')

# Configurar o serviço do WebDriver
driver_path = ChromeDriverManager().install()
# Inicializar o driver do Chrome
service = Service(ChromeDriverManager().install())
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
