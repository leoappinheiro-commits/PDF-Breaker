from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from rapidfuzz import fuzz
from tqdm import tqdm

from .ocr_engine import OcrEngine
from .page_detector import PageDetector
from .pdf_reader import PdfReader
from .regex_extractor import RegexExtractor

LOGGER = logging.getLogger(__name__)

PIS_PATTERN = re.compile(r"\b(\d{3}\s*[.\-]?\s*\d{5}\s*[.\-]?\s*\d{2}\s*-\s*\d)\b")
MONEY_PATTERN = re.compile(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b")
HEADER_VARIATIONS = [
    "NOME TRABALHADOR",
    "NOME TRABALHADOR PIS",
    "NOME TRABALHADOR PIS/PASEP",
    "NOME TRABALHADOR PIS/PASEP/CI",
]


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

            trab_pages = [p.text for p in classified_pages if p.section == "trabalhadores"]
            trab_pages_text = "\n".join(trab_pages).strip() or full_text
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
        """Parser vertical da SEFIP: linha de nome+PIS seguida por linha de remuneração."""
        rows: List[Dict] = []
        if not text.strip():
            return rows

        linhas = [ln.strip() for ln in text.splitlines() if ln.strip()]

        for idx, linha in enumerate(linhas):
            if self._is_worker_header_line(linha):
                continue

            match_pis = PIS_PATTERN.search(linha)
            if not match_pis:
                continue

            nome = linha[: match_pis.start()].strip(" -:\t")
            pis = self._normalize_pis(match_pis.group(1))
            if not nome or not pis:
                continue

            remuneracao = self._extract_remuneracao_from_next_lines(linhas, idx)
            if remuneracao is None:
                LOGGER.debug("Remuneração não encontrada para PIS %s no arquivo %s", pis, file_name)

            rows.append(
                {
                    "Arquivo": file_name,
                    "Nome": nome,
                    "PIS": pis,
                    "Remuneracao": remuneracao,
                }
            )

        LOGGER.info("Trabalhadores extraídos de %s: %d", file_name, len(rows))
        return rows

    def _extract_remuneracao_from_next_lines(self, linhas: List[str], idx: int) -> Optional[float]:
        # Busca nas 3 linhas seguintes para tolerar quebras de OCR/layout.
        for jump in (1, 2, 3):
            target_idx = idx + jump
            if target_idx >= len(linhas):
                break

            target = linhas[target_idx]
            values = MONEY_PATTERN.findall(target)
            if not values:
                continue
            return self._to_float_or_none(values[0])

        return None

    @staticmethod
    def _normalize_pis(pis_text: str) -> Optional[str]:
        only_digits = re.sub(r"\D", "", pis_text)
        return only_digits if len(only_digits) == 11 else None

    def _is_worker_header_line(self, line: str) -> bool:
        norm = self._normalize_text(line)
        for expected in HEADER_VARIATIONS:
            score = fuzz.partial_ratio(self._normalize_text(expected), norm)
            if score >= 78:
                return True
        return False

    @staticmethod
    def _normalize_text(text: str) -> str:
        return " ".join(text.upper().replace("Ç", "C").replace("Ã", "A").replace("Á", "A").split())

    def _to_float_or_none(self, text: str) -> float | None:
        try:
            return self.regex.normalize_brl_number(text)
        except (TypeError, ValueError):
            LOGGER.warning("Erro de parsing de valor: %s", text)
            return None
