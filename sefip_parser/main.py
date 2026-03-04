from __future__ import annotations

import argparse
import logging
from pathlib import Path

from .excel_exporter import ExcelExporter
from .sefip_parser import SefipParser


def configure_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )


def processar_pasta(input_dir: Path, output_file: Path, tesseract_cmd: str | None = None) -> None:
    parser = SefipParser(tesseract_cmd=tesseract_cmd)
    result = parser.process_folder(input_dir)
    ExcelExporter().export(result.empresa_df, result.trabalhadores_df, result.resumo_df, output_file)


def main() -> None:
    argp = argparse.ArgumentParser(description="Extrator inteligente de PDFs SEFIP RE")
    argp.add_argument("input_dir", type=Path, help="Pasta com PDFs SEFIP")
    argp.add_argument(
        "--output",
        type=Path,
        default=Path("SEFIP_CONSOLIDADO.xlsx"),
        help="Arquivo Excel consolidado",
    )
    argp.add_argument("--tesseract-cmd", type=str, default=None, help="Caminho do executável tesseract")
    argp.add_argument("--log-level", type=str, default="INFO")
    args = argp.parse_args()

    configure_logging(args.log_level)
    processar_pasta(args.input_dir, args.output, args.tesseract_cmd)


if __name__ == "__main__":
    main()
