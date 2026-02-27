import os
import PyPDF2 as pyf
import pandas as pd

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
