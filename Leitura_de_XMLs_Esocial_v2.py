import os
import pandas as pd
import xmltodict
import xml.parsers.expat
from zipfile import ZipFile
from tqdm import tqdm
import time


caminho_pasta_zip_temp = input("Insira o caminho onde encontram-se os arquivos em formato .zip")
caminho_pasta_zip = caminho_pasta_zip_temp
print(caminho_pasta_zip)


caminho_pasta_resultado_zip_temp = input("Insira o caminho onde encontram-se os arquivos extraídos")
caminho_pasta_resultado_zip = caminho_pasta_resultado_zip_temp


## Funcoes para quebrar TAGs/XML
def quebrar_tags(documento, prefixo=""):
    if isinstance(documento, dict):
        novo_dict = {}
        for chave, valor in documento.items():
            if isinstance(valor, list):
                for i, item in enumerate(valor):
                    sub_dict = quebrar_tags(item, f"{chave}.{i}")
                    for sub_chave, sub_valor in sub_dict.items():
                        novo_dict[sub_chave] = sub_valor
            elif isinstance(valor, dict):
                sub_dict = quebrar_tags(valor, chave)
                for sub_chave, sub_valor in sub_dict.items():
                    novo_dict[sub_chave] = sub_valor
            else:
                novo_dict[f"{prefixo}.{chave}"] = valor
        return novo_dict
    else:
        return {prefixo: documento}


##def ler_xml_danfe(nota):
  ##  with open(nota, 'rb') as arquivo:
    ##    documento = xmltodict.parse(arquivo.read())
    ##documento_quebrado = quebrar_tags(documento)
    ##return documento_quebrado


def ler_xml_servico(nota):
    try:
        with open(nota, 'rb') as arquivo:
            documento = xmltodict.parse(arquivo.read())
        documento_quebrado = quebrar_tags(documento)
        return documento_quebrado
    except xml.parsers.expat.ExpatError as e:
        print(f"Erro ao processar o arquivo {nota}: {e}")
        return None

## Indica o caminho dos ZIPs
lista_arquivos1 = os.listdir(caminho_pasta_zip)

## Cria um DataFrame Vazio para alimentarmos informações

df_final_S1010 = pd.DataFrame()
df_final_S5011 = pd.DataFrame()
df_final_S1200 = pd.DataFrame()

 #Loop para quebrar todos os ZIPs de Esocial contidos no caminho
for arquivo in tqdm(lista_arquivos1):
    if ".zip" in arquivo:
        with ZipFile (caminho_pasta_zip+"/"+arquivo) as arquivo_zipado:
            arquivo_zipado.extractall(caminho_pasta_resultado_zip)


## Indica o caminho dos resultados dos Zips

lista_arquivos = os.listdir(caminho_pasta_resultado_zip)

## Loop para quebrar as tags dos XMLs extraídos, utilizando as funcoes definidas anteriormente
for arquivo in tqdm(lista_arquivos):
    if 'xml' in arquivo:
        if 'S-1010' in arquivo:
            documento = ler_xml_servico(os.path.join(caminho_pasta_resultado_zip, arquivo))
            if documento is not None:
                # Converter o documento quebrado em DataFrame
                df = pd.DataFrame.from_dict(documento, orient='index').T
                df_final_S1010 = df_final_S1010.append(df)
        if 'S-5011' in arquivo:
            documento = ler_xml_servico(os.path.join(caminho_pasta_resultado_zip, arquivo))
            if documento is not None:
                # Converter o documento quebrado em DataFrame
                df = pd.DataFrame.from_dict(documento, orient='index').T
                df_final_S5011 = df_final_S5011.append(df)
        if 'S-1200' in arquivo:
            documento = ler_xml_servico(os.path.join(caminho_pasta_resultado_zip, arquivo))
            if documento is not None:
                # Converter o documento quebrado em DataFrame
                df = pd.DataFrame.from_dict(documento, orient='index').T
                df_final_S1200 = df_final_S1200.append(df)

# Renomear as colunas removendo o prefixo
df_final_S1010.columns = [coluna.split('.')[-1] for coluna in df_final_S1010.columns]
df_final_S5011.columns = [coluna.split('.')[-1] for coluna in df_final_S5011.columns]
df_final_S1200.columns = [coluna.split('.')[-1] for coluna in df_final_S1200.columns]

caminho_arquivo = os.path.join(caminho_pasta_resultado_zip, 'S1010.xlsx')
df_final_S1010.to_excel(caminho_arquivo, index=False)

print(f"O arquivo 'S1010.xlsx' foi salvo em: {caminho_arquivo}")

caminho_arquivo = os.path.join(caminho_pasta_resultado_zip, 'S5011.xlsx')
df_final_S5011.to_excel(caminho_arquivo, index=False)

caminho_arquivo = os.path.join(caminho_pasta_resultado_zip, 'S1200.xlsx')
df_final_S1200.to_excel(caminho_arquivo, index=False)