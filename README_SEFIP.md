# Parser SEFIP RE (OCR + Anchors + Regex)

## Como rodar

1. Crie e ative um ambiente virtual (opcional, recomendado):

```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows
```

2. Instale dependências:

```bash
pip install -r requirements_sefip.txt
```

3. Execute o script principal (mantendo o nome original):

```bash
python "OCR no PDF.py" ./PDF --output SEFIP_CONSOLIDADO.xlsx
```

- `./PDF` é a pasta com os PDFs SEFIP.
- Se você não informar pasta, o padrão já é `./PDF`.

## Argumentos úteis

```bash
python "OCR no PDF.py" --help
```

Argumentos:
- `input_dir`: pasta com PDFs (`default=PDF`)
- `--output`: nome/caminho do Excel final
- `--tesseract-cmd`: caminho do executável do Tesseract
- `--log-level`: nível de log (`INFO`, `DEBUG` etc.)

## Exemplo (Windows)

```bash
python "OCR no PDF.py" "C:\\SEFIP\\PDFs" --output "C:\\SEFIP\\SEFIP_CONSOLIDADO.xlsx" --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
```
