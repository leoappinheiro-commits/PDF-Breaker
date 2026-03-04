"""Entry-point simples para o parser SEFIP.

Uso rápido:
    python "OCR no PDF.py" /caminho/da/pasta_com_pdfs --output SEFIP_CONSOLIDADO.xlsx

Se nenhum diretório for informado, usa a pasta `./PDF`.
"""

from __future__ import annotations

import sys


def _erro_dependencia(exc: ModuleNotFoundError) -> None:
    print(
        "Erro de dependência: biblioteca ausente para executar o parser.\n"
        f"Módulo faltante: {exc.name}\n\n"
        "Instale os pacotes necessários com:\n"
        "    pip install -r requirements_sefip.txt\n"
    )


def main() -> None:
    try:
        from sefip_parser.main import main as parser_main
    except ModuleNotFoundError as exc:
        _erro_dependencia(exc)
        sys.exit(1)

    parser_main()


if __name__ == "__main__":
    main()
