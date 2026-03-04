from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import pandas as pd
from tqdm import tqdm

from .ocr_engine import OcrEngine
from .page_detector import PageDetector
from .pdf_reader import PdfReader
from .regex_extractor import RegexExtractor

LOGGER = logging.getLogger(__name__)


@dataclass
class ParseResult:
    empresa_df: pd.DataFrame
    trabalhadores_df: pd.DataFrame
    resumo_df: pd.DataFrame


class SefipParser:
    """Pipeline principal para extrair dados SEFIP de múltiplos PDFs."""

    def __init__(self, tesseract_cmd: str | None = None) -> None:
        self.reader = PdfReader()
        self.ocr = OcrEngine(tesseract_cmd=tesseract_cmd)
        self.detector = PageDetector()
        self.regex = RegexExtractor()

    def process_folder(self, input_dir: Path) -> ParseResult:
        pdf_files = sorted(input_dir.glob("*.pdf"))
        empresa_rows: List[Dict] = []
        resumo_rows: List[Dict] = []
        trabalhadores_rows: List[Dict] = []

        for pdf_file in tqdm(pdf_files, desc="Processando PDFs"):
            LOGGER.info("Arquivo sendo processado: %s", pdf_file.name)
            page_texts = self._extract_text_by_strategy(pdf_file)
            classified_pages = self.detector.classify_pages(page_texts)
            LOGGER.info(
                "Páginas detectadas: %s",
                {p.page_number: p.section for p in classified_pages},
            )

            full_text = "\n".join(page_texts)
            extraction = self.regex.extract(full_text)

            empresa_row = {"Arquivo": pdf_file.name, **extraction.empresa}
            resumo_row = {"Arquivo": pdf_file.name, **extraction.resumo}
            empresa_rows.append(empresa_row)
            resumo_rows.append(resumo_row)

            trab_pages_text = "\n".join(
                p.text for p in classified_pages if p.section == "trabalhadores"
            )
            trabalhadores_rows.extend(self._parse_trabalhadores(trab_pages_text, pdf_file.name))

        return ParseResult(
            empresa_df=pd.DataFrame(empresa_rows),
            trabalhadores_df=pd.DataFrame(trabalhadores_rows),
            resumo_df=pd.DataFrame(resumo_rows),
        )

    def _extract_text_by_strategy(self, pdf_file: Path) -> List[str]:
        load_result = self.reader.load(pdf_file)
        if not load_result.needs_ocr:
            LOGGER.info("PDF digital detectado (sem OCR): %s", pdf_file.name)
            return [p.text for p in load_result.page_texts]

        LOGGER.info("PDF digitalizado detectado (OCR): %s", pdf_file.name)
        pixmaps = self.reader.render_pages_as_images(pdf_file)
        return self.ocr.extract_texts(pixmaps)

    def _parse_trabalhadores(self, text: str, file_name: str) -> List[Dict]:
        """Converte bloco textual de trabalhadores para linhas estruturadas."""
        rows: List[Dict] = []
        if not text.strip():
            return rows

        for line in text.splitlines():
            clean = " ".join(line.split())
            if not clean or "PIS" in clean.upper() and "NOME" in clean.upper():
                continue

            # Split por 2+ espaços para preservar nomes com espaço
            parts = re.split(r"\s{2,}", line.strip())
            if len(parts) < 6:
                continue

            pis = self._extract_pis(parts[0])
            if not pis:
                continue

            nome = parts[1].strip()
            rem, base_fgts, fgts, inss = [self._to_float_or_none(v) for v in parts[2:6]]
            rows.append(
                {
                    "Arquivo": file_name,
                    "PIS": pis,
                    "Nome": nome,
                    "Remuneracao": rem,
                    "Base_FGTS": base_fgts,
                    "FGTS": fgts,
                    "INSS": inss,
                }
            )
        return rows

    @staticmethod
    def _extract_pis(text: str) -> str | None:
        match = re.search(r"\b\d{11}\b", re.sub(r"\D", "", text))
        return match.group(0) if match else None

    def _to_float_or_none(self, text: str) -> float | None:
        try:
            return self.regex.normalize_brl_number(text)
        except ValueError:
            LOGGER.warning("Erro de parsing de valor: %s", text)
            return None
