import pyautogui
import time
import os
import openpyxl as xl
import PyPDF2 as pyf
import pandas as pd


# variáveis

caminho_arquivos_txt = r'C:\Users\leonardopinheiro\Desktop\Teste'
caminho_nova_pasta = r'C:\Users\leonardopinheiro\Desktop' + '/PDFs'
i = 0
lista_arquivos = os.listdir(caminho_arquivos_txt)
print(lista_arquivos)

# cria uma nova pasta para armazenar os PDFs
if not os.path.exists(caminho_nova_pasta):
    os.makedirs(caminho_nova_pasta)

# Cria um arquivo Excel para armazenar os resultados
arquivo_excel = os.path.join(caminho_arquivos_txt, 'resultado.xlsx')
planilha = xl.Workbook()
planilha_ativa = planilha.active
planilha_ativa.append(['Nome do arquivo', 'Linha'])

# abre o validador EFD ICMS IPI

pyautogui.hotkey('win')
pyautogui.PAUSE = 1
pyautogui.write('EFD ICMS')
pyautogui.PAUSE = 1
pyautogui.press('enter')
time.sleep(25)

# loop para importar e consultar escrituração
for arquivo in lista_arquivos:
    i = i + 1
    time.sleep(5)
    with pyautogui.hold('ctrl'):
            pyautogui.press('i')
    caminho_completo = os.path.join(caminho_arquivos_txt, arquivo.replace('"',''))
    pyautogui.write(caminho_completo)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(5)
    with pyautogui.hold('ctrl'):
            pyautogui.press('c')
    time.sleep(5)
    with pyautogui.hold('ctrl'):
            pyautogui.press('a')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(2)
    caminho_completo = os.path.join(caminho_nova_pasta, arquivo.replace('"',''))
    print(caminho_completo)
    pyautogui.write(caminho_completo)
    time.sleep(2)
    caminho_completo_novo = os.path.join(caminho_completo + ".pdf")
    pyautogui.press('enter')
    time.sleep(20)
    with pyautogui.hold('alt'):
        pyautogui.press('F4')
    time.sleep(2)
    with pyautogui.hold('ctrl'):
        pyautogui.press('e')
    time.sleep(2)
    with pyautogui.hold('ctrl'):
        pyautogui.press('a')
    time.sleep(2)
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('esc')

## setor do código para verificação no PDF da situação do arquivo - e transposição para uma planilha em excel

nome_arquivo = r'C:\Users\leonardopinheiro\Desktop\PDFs\90576356000160-0240124146-20140201-20140228-1-1E5197267E2B7AF4D13122C36EBF2399EAB3AC1A-SPED-EFD.txt.pdf'
arquivo = pyf.PdfReader(nome_arquivo)

texto_procurado = "Resultado da Verificação"

i = 1
# percorrendo todas as páginas
for pagina in arquivo.pages:
    # pegar o que está escrito na página
    texto_pagina = pagina.extract_text()
    # verificar se o dentro do texto da página tem o texto_procurado
    if texto_procurado in texto_pagina:
        # se tiver, me diz qual é o número da página
        print(f'Está na Página {i}')
        num_pagina = i
        texto_final = texto_pagina
    i += 1

posicao = texto_final.find("Resultado da Verificação")
posicao_final = texto_final.find(".", posicao + 1)
texto_despesa = texto_final[posicao:posicao_final]
print(texto_despesa)

nome_arquivo_base = os.path.basename(nome_arquivo)

novo_df = pd.DataFrame({'Nome do Arquivo': [nome_arquivo_base], 'Texto da Despesa': [texto_despesa]})
novo_df.to_excel("Resultados.xlsx", index=False)

