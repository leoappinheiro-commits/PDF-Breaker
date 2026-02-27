import re
from openpyxl import Workbook

arquivo_entrada = r"C:\Users\leonardopinheiro\Desktop\Teste_INSS\Comprovantes INSS 2014_OCR_COMPLETO.txt"

with open(arquivo_entrada, "r", encoding="utf-8", errors="ignore") as f:
    texto = f.read()

# ========= LIMPEZA OCR =========
texto = re.sub(r"[|']", " ", texto)
texto = re.sub(r"\s+", " ", texto)

# ========= FUNÇÃO PARA AJUSTAR VALORES OCR =========
def corrige_valor(valor):
    if not valor:
        return ""

    # remove espaços internos tipo "50,432 51"
    valor = re.sub(r"\s+", "", valor)

    # corrige formato tipo 50,43251 -> 50.432,51
    if re.match(r"\d{1,3},\d{3}\d{2}$", valor):
        valor = valor.replace(",", ".", 1)
        valor = valor[:-2] + "," + valor[-2:]

    return valor


# ========= SEPARAR GPS =========
blocos = re.split(r"MINIST[ÉE]RIO\s+DA\s+PREVID", texto, flags=re.I)

dados_extraidos = []

for bloco in blocos:
    if len(bloco) < 200:
        continue

    competencia = re.search(
        r"COMPET[ÊE]NCIA[^0-9]{0,20}(\d{2}/\d{4})",
        bloco, re.I)

    identificador = re.search(
        r"IDENTIFICADOR[^0-9]{0,30}([\d\./-]+)",
        bloco, re.I)

    # >>> REGEX MELHORADO <<<
    inss = re.search(
        r"VALORES?\s+DO\s+INSS[^0-9]{0,30}((?:\d[\d\.,\s]{3,}))",
        bloco, re.I)

    outras = re.search(
        r"ENTIDADES[^0-9]{0,30}((?:\d[\d\.,\s]{3,}))",
        bloco, re.I)

    total = re.search(
        r"TOTAL[^0-9]{0,30}((?:\d[\d\.,\s]{3,}))",
        bloco, re.I)

    dados_extraidos.append({
        "Competencia": competencia.group(1) if competencia else "",
        "Identificador": identificador.group(1) if identificador else "",
        "Valor INSS": corrige_valor(inss.group(1)) if inss else "",
        "Outras Entidades": corrige_valor(outras.group(1)) if outras else "",
        "Total": corrige_valor(total.group(1)) if total else ""
    })

# ========= GERAR EXCEL =========
wb = Workbook()
ws = wb.active
ws.title = "GPS"

ws.append([
    "Competencia",
    "Identificador",
    "Valor INSS",
    "Outras Entidades",
    "Total"
])

for linha in dados_extraidos:
    ws.append(list(linha.values()))

arquivo_saida = r"C:\Users\leonardopinheiro\Desktop\Teste_INSS\gps_extraido.xlsx"
wb.save(arquivo_saida)

print("Arquivo gerado com sucesso:")
print(arquivo_saida)
