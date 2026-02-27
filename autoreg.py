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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

driver_version = "119.0.6045.106"  # Substitua pelo número da versão desejada

i = 0
Chaves = pd.read_excel(r'C:\Users\leonardopinheiro\Downloads\teste (1).xlsx')
Chaves2 = pd.read_excel(r'C:\Users\leonardopinheiro\Desktop\teste2.xlsx')

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
navegador.get("https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/formularios/impostos/parcelamento/autorregularizacao.html")

# Diretório de downloads
## downloads_dir = r'C:\Users\leonardopinheiro\Downloads'


# Realizar qualquer interação adicional no navegador, se necessário
for i, chave in Chaves.iterrows():
    botao_debito = navegador.find_element(By.XPATH, '//*[@id="btnDebito"]/button')

    # Mover o cursor do mouse para as coordenadas do botão e clicar nele
    ActionChains(navegador).move_to_element(botao_debito).click().perform()

    # Aguardar um segundo após o clique
    time.sleep(1)

    # Criando o XPath dinâmico com base no índice do loop
    xpath_tipo = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[1]/select'
    campo_tipo = navegador.find_element(By.XPATH, xpath_tipo)
    if not pd.isna(chave['Tipo declaração']):
        campo_tipo.send_keys(chave['Tipo declaração'])

    time.sleep(1)

    if not pd.isna(chave['Data entrega']):
        data_entrega_string = chave['Data entrega'].strftime('%d/%m/%Y')
        xpath_data = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[2]/input'
        campo_data = navegador.find_element(By.XPATH, xpath_data)
        campo_data.send_keys(data_entrega_string)

    time.sleep(1)

    if not pd.isna(chave['CPF/CNPJ do débito']):
        xpath_cpf_cnpj = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[3]/input'
        campo_cpf_cnpj = navegador.find_element(By.XPATH, xpath_cpf_cnpj)
        campo_cpf_cnpj.send_keys(chave['CPF/CNPJ do débito'])

    time.sleep(1)

    if not pd.isna(chave['Nº do processo/DEBCAD']):
        xpath_processo = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[4]/input'
        campo_processo = navegador.find_element(By.XPATH, xpath_processo)
        campo_processo.send_keys(chave['Nº do processo/DEBCAD'])

    time.sleep(1)

    if not pd.isna(chave['Código receita']):
        xpath_codigo_receita = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[5]/select'
        campo_codigo_receita = navegador.find_element(By.XPATH, xpath_codigo_receita)
        campo_codigo_receita.send_keys(chave['Código receita'])

    time.sleep(1)

    if not pd.isna(chave['Período de apuração']):
        data_apuracao_string = chave['Período de apuração'].strftime('%d/%m/%Y')
        xpath_apuracao = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[6]/input'
        campo_apuracao = navegador.find_element(By.XPATH, xpath_apuracao)
        campo_apuracao.send_keys(data_apuracao_string)

    time.sleep(1)

    if not pd.isna(chave['Vencimento do tributo']):
        data_vencimento_string = chave['Vencimento do tributo'].strftime('%d/%m/%Y')
        xpath_vencimento = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[7]/input'
        campo_vencimento = navegador.find_element(By.XPATH, xpath_vencimento)
        campo_vencimento.send_keys(data_vencimento_string)

    time.sleep(1)

    if not pd.isna(chave['Valor (R$)']):
        # Enviar o valor como uma string formatada
        valor_string = '{:.2f}'.format(chave['Valor (R$)'])  # Formatando com 2 casas decimais
        xpath_valor = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[8]/input'
        campo_valor = navegador.find_element(By.XPATH, xpath_valor)
        campo_valor.send_keys(valor_string)

    time.sleep(1)

    if not pd.isna(chave['CIB/CNO/CNPJ prestador']):
        xpath_cib_cno_cnpj = f'//*[@id="tabelaDebitosBody"]/tr[{i + 1}]/td[9]/input'
        campo_cib_cno_cnpj = navegador.find_element(By.XPATH, xpath_cib_cno_cnpj)
        campo_cib_cno_cnpj.send_keys(chave['CIB/CNO/CNPJ prestador'])

    time.sleep(1)


for i, chave2 in Chaves2.iterrows():

    if not pd.isna(chave2['Montante']):
        valor_string_montante = '{:.2f}'.format(chave2['Montante'])  # Formatando com 2 casas decimais
        xpath_montante = f'//*[@id="tabelaPropriosBody"]/tr[1]/td[2]/input'
        campo_montante = navegador.find_element(By.XPATH, xpath_montante)
        campo_montante.send_keys(valor_string_montante)

        time.sleep(1)

    if not pd.isna(chave2['Valor']):
        valor_string_valor = '{:.2f}'.format(chave2['Valor'])  # Formatando com 2 casas decimais
        xpath_valor = f'//*[@id="tabelaPropriosBody"]/tr[1]/td[4]/input'
        campo_valor = navegador.find_element(By.XPATH, xpath_valor)
        campo_valor.send_keys(valor_string_valor)

        time.sleep(1)


    if not pd.isna(chave2['Data entrega ECF']):
        data_vencimento_string = chave2['Data entrega ECF'].strftime('%d/%m/%Y')
        xpath_vencimento = f'//*[@id="tabelaPropriosBody"]/tr[1]/td[5]/input'
        campo_vencimento = navegador.find_element(By.XPATH, xpath_vencimento)
        campo_vencimento.send_keys(data_vencimento_string)

        # Esperar até que o botão "Gerar PDF" seja clicável
        # Encontrar o botão "Gerar PDF"
botao_gerar_pdf = navegador.find_element(By.XPATH, "//button[contains(text(), 'Gerar PDF')]")

        # Mover o cursor do mouse para as coordenadas do botão e clicar nele
ActionChains(navegador).move_to_element(botao_gerar_pdf).click().perform()

        # Aguardar um momento para a impressão começar
time.sleep(20)
        

# Fechar o navegador
navegador.quit()
