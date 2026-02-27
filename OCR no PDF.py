import os
import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
import pandas as pd

# ===== CAMINHO DO TESSERACT =====
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\leonardopinheiro\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)

# ===== PDF DE ENTRADA =====
PDF_PATH = r"C:\Users\leonardopinheiro\Downloads\OneDrive_2_2-19-2026\Comprovantes INSS 2014.pdf"

# ===== POPPLER (necessário para pdf2image no Windows) =====
POPPLER_PATH = r"C:\Users\leonardopinheiro\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"

# ===== PASTA FIXA DE SAÍDA =====
PASTA_SAIDA = r"C:\Users\leonardopinheiro\Desktop\Teste_INSS"
os.makedirs(PASTA_SAIDA, exist_ok=True)

nome_pdf = os.path.splitext(os.path.basename(PDF_PATH))[0]
txt_saida = os.path.join(PASTA_SAIDA, f"{nome_pdf}_OCR_COMPLETO.txt")
csv_saida = os.path.join(PASTA_SAIDA, f"{nome_pdf}_OCR_TABELA.csv")


# ===== PRÉ-PROCESSAMENTO OCR =====
def preprocessar(imagem_pil):
    img = np.array(imagem_pil)

    # escala cinza
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # reduzir ruído
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # threshold adaptativo
    thresh = cv2.adaptiveThreshold(
        blur, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 2
    )

    # reforçar contraste
    kernel = np.ones((1, 1), np.uint8)
    img = cv2.dilate(thresh, kernel, 1)
    img = cv2.erode(img, kernel, 1)

    return img


# ===== OCR PDF =====
print("Convertendo PDF em imagens...")
pages = convert_from_path(
    PDF_PATH,
    dpi=400,
    poppler_path=POPPLER_PATH
)

texto_total = []
linhas_csv = []

for i, page in enumerate(pages):
    print(f"OCR página {i+1}/{len(pages)}")

    img = preprocessar(page)

    config = r'--oem 3 --psm 6 -l por'

    texto = pytesseract.image_to_string(img, config=config)

    texto_total.append(texto)

    # tentativa simples de estruturar tabela
    for linha in texto.split("\n"):
        if linha.strip():
            linhas_csv.append(linha.split())


# ===== SALVAR TXT =====
try:
    with open(txt_saida, "w", encoding="utf-8") as f:
        f.write("\n\n".join(texto_total))
    print("TXT salvo em:")
    print(txt_saida)
except Exception as e:
    print("Erro ao salvar TXT:", e)


# ===== SALVAR CSV =====
try:
    df = pd.DataFrame(linhas_csv)
    df.to_csv(csv_saida, index=False, sep=";")
    print("CSV salvo em:")
    print(csv_saida)
except Exception as e:
    print("Erro ao salvar CSV:", e)


print("\n✔ OCR FINALIZADO")
