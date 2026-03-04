from __future__ import annotations

import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from rapidfuzz import fuzz
from tqdm import tqdm

if __package__ in (None, ""):
    project_root = Path(__file__).resolve().parent.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    from sefip_parser.ocr_engine import OcrEngine
    from sefip_parser.page_detector import PageDetector
    from sefip_parser.pdf_reader import PdfReader
    from sefip_parser.regex_extractor import RegexExtractor
else:
    from .ocr_engine import OcrEngine
    from .page_detector import PageDetector
    from .pdf_reader import PdfReader
    from .regex_extractor import RegexExtractor

LOGGER = logging.getLogger(__name__)

PIS_PATTERN = re.compile(r"\b(\d{3}\s*[.\-]?\s*\d{5}\s*[.\-]?\s*\d{2}\s*-\s*\d)\b")
MONEY_PATTERN = re.compile(r"\b(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}\b")
OCR_MONEY_PATTERN = re.compile(r"(?:[:.]\s*)?[\dBO][\dBO\s.:,]*,\s*\d{2}")
DATE_PATTERN = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
CATEGORY_PATTERN = re.compile(r"\b\d{2}\b")
CBO_PATTERN = re.compile(r"^\s*(\d{5})\s*$")
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
            rows = self._parse_trabalhadores(trab_pages_text, pdf_file.name)
            trabalhadores_rows.extend(rows)

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
        """Parser estrutural SEFIP: PIS -> remuneração -> admissão -> contribuição."""
        if not text.strip():
            return []

        linhas = [ln.strip() for ln in text.splitlines() if ln.strip()]
        rows: List[Dict] = []

        for idx, linha in enumerate(linhas):
            if self._is_worker_header_line(linha):
                continue

            match_pis = PIS_PATTERN.search(linha)
            if not match_pis:
                continue

            nome = linha[: match_pis.start()].strip(" -:	")
            pis = self._normalize_pis(match_pis.group(1))
            if not nome or not pis:
                continue

            rem_sem_13, rem_13, base_13 = self._extract_remuneracoes(linhas, idx)
            admissao, contrib = self._extract_admissao_contrib_seq(linhas, idx)

            rows.append(
                {
                    "Arquivo": file_name,
                    "Nome_Trabalhador": nome,
                    "PIS": pis,
                    "Rem_Sem_13": rem_sem_13,
                    "Rem_13": rem_13,
                    "Base_13_Prev": base_13,
                    "Admissao": admissao,
                    "Contrib_Segurado": contrib,
                }
            )

        self._log_first_rows_preview(rows)
        LOGGER.info("Trabalhadores extraídos de %s: %d", file_name, len(rows))
        return rows

    def _extract_admissao_contrib_seq(self, linhas: List[str], idx: int) -> tuple[Optional[str], Optional[float]]:
        """Após linha do PIS, busca admissão em +2/+3 e contribuição na linha seguinte."""
        for offset in (2, 3):
            adm_idx = idx + offset
            if adm_idx >= len(linhas):
                break
            adm_line = self._normalize_numbers_text(linhas[adm_idx])
            date_match = DATE_PATTERN.search(adm_line)
            if not date_match:
                continue

            contrib_idx = adm_idx + 1
            if contrib_idx >= len(linhas):
                return date_match.group(0), None

            contrib_values = self._extract_money_values(linhas[contrib_idx])
            contrib = self._to_float_or_none(contrib_values[0]) if contrib_values else None
            contrib = self._sanitize_expected_range(contrib, "Contrib_Segurado", 50, 1000)
            return date_match.group(0), contrib

        return None, None

    def _log_first_rows_preview(self, rows: List[Dict]) -> None:
        """Log rápido das primeiras linhas para validação estrutural sem prints permanentes."""
        for row in rows[:20]:
            LOGGER.debug(
                "DEBUG_PREVIEW | Nome=%s | Rem_Sem_13=%s | Contrib_Segurado=%s",
                row.get("Nome_Trabalhador"),
                row.get("Rem_Sem_13"),
                row.get("Contrib_Segurado"),
            )

    def _extract_identification_block(self, linhas: List[str], file_name: str) -> List[Dict]:
        trabalhadores: List[Dict] = []

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

            rem_sem_13, rem_13, base_13 = self._extract_remuneracoes(linhas, idx)
            trabalhadores.append(
                {
                    "Arquivo": file_name,
                    "Nome_Trabalhador": nome,
                    "PIS": pis,
                    "Rem_Sem_13": rem_sem_13,
                    "Rem_13": rem_13,
                    "Base_13_Prev": base_13,
                    "Admissao": None,
                    "Contrib_Segurado": None,
                    "Categoria": None,
                    "Ocor": None,
                    "Data_Movimentacao": None,
                    "Deposito_FGTS": None,
                    "CBO": None,
                }
            )

        return trabalhadores

    def _extract_remuneracoes(self, linhas: List[str], idx: int) -> tuple[Optional[float], Optional[float], Optional[float]]:
        """Mantém lógica original: remuneração na linha imediatamente após o PIS."""
        target_idx = idx + 1
        if target_idx >= len(linhas):
            return None, None, None

        values = self._extract_money_values(linhas[target_idx])
        if not values:
            return None, None, None

        rem_sem_13 = self._sanitize_expected_range(self._to_float_or_none(values[0]), "Rem_Sem_13", 100, 10000)
        rem_13 = self._to_float_or_none(values[1]) if len(values) > 1 else None
        base_13 = self._to_float_or_none(values[2]) if len(values) > 2 else None
        return rem_sem_13, rem_13, base_13

    def _extract_previdenciario_block(self, linhas: List[str]) -> List[Dict]:
        entries: List[Dict] = []
        for idx, linha in enumerate(linhas):
            admissao_match = DATE_PATTERN.search(linha)
            if not admissao_match:
                continue

            admissao = admissao_match.group(0)
            categoria = self._extract_categoria(linha, admissao, linhas, idx)
            data_mov = self._extract_second_date(linha, admissao)

            entries.append(
                {
                    "Admissao": admissao,
                    "Categoria": categoria,
                    "Ocor": None,
                    "Data_Movimentacao": data_mov,
                }
            )

        return entries


    def _extract_contribuicoes(self, linhas: List[str]) -> List[Optional[float]]:
        """Extrai lista de contribuicoes por bloco previdenciário para mapeamento posicional."""
        contribuicoes: List[Optional[float]] = []
        for idx, linha in enumerate(linhas):
            if not DATE_PATTERN.search(linha):
                continue
            contribuicoes.append(self._extract_contrib_from_window(linhas, idx))
        return contribuicoes

    def _extract_fgts_block(self, linhas: List[str]) -> List[Dict]:
        entries: List[Dict] = []
        for idx, linha in enumerate(linhas):
            cbo_match = CBO_PATTERN.match(linha)
            if not cbo_match:
                continue

            deposito = self._extract_deposito_from_window(linhas, idx)
            entries.append({"CBO": cbo_match.group(1), "Deposito_FGTS": deposito})

        return entries


    def _normalize_numbers_text(self, text: str) -> str:
        """Normaliza fragmentação OCR de números monetários antes da extração."""
        normalized = text.upper()
        normalized = normalized.replace(" . ", ".")
        normalized = normalized.replace(" ,", ",")
        normalized = normalized.replace(", ", ",")
        normalized = normalized.replace(" .", ".")
        normalized = normalized.replace(". ", ".")
        normalized = normalized.replace("B", "8")
        normalized = normalized.replace("O", "0")
        normalized = normalized.replace(":.", "1")
        normalized = normalized.replace("::", "1")
        normalized = re.sub(r"^\s*[:.]\s*(?=\d)", "1", normalized)
        normalized = normalized.replace(";", ",")
        normalized = re.sub(r"\s+,\s*", ",", normalized)
        normalized = re.sub(r"\s+\.\s*", ".", normalized)
        normalized = re.sub(r"(\d)\s*,\s*(\d{2})", r"\1,\2", normalized)
        normalized = re.sub(r"(\d)\s*\.\s*(\d{3})", r"\1.\2", normalized)
        return normalized

    def _normalize_ocr_value_token(self, token: str) -> str:
        """Reconstrói token monetário para formato BRL válido."""
        normalized = self._normalize_numbers_text(token).replace(" ", "")
        normalized = re.sub(r"^[.:](?=\d)", "1", normalized)
        normalized = re.sub(r"[^\d,\.]", "", normalized)

        if "," in normalized:
            inteiro, dec = normalized.rsplit(",", 1)
            inteiro = inteiro.replace(".", "")
            normalized = f"{inteiro},{dec[:2]}"

        return normalized

    def _extract_money_values(self, line: str) -> List[str]:
        """Extrai valores monetários após normalização OCR da linha inteira."""
        normalized_line = self._normalize_numbers_text(line)
        candidates = OCR_MONEY_PATTERN.findall(normalized_line)
        values: List[str] = []

        for candidate in candidates:
            token = self._normalize_ocr_value_token(candidate)
            if MONEY_PATTERN.fullmatch(token):
                values.append(token)

        return values

    def _extract_categoria(self, line: str, admissao: str, linhas: List[str], idx: int) -> Optional[str]:
        tail = line.split(admissao, maxsplit=1)[-1]
        match = CATEGORY_PATTERN.search(tail)
        if match:
            return match.group(0)

        next_idx = idx + 2
        if next_idx < len(linhas):
            next_match = CATEGORY_PATTERN.search(linhas[next_idx])
            if next_match:
                return next_match.group(0)
        return None

    def _extract_second_date(self, line: str, admissao: str) -> Optional[str]:
        dates = DATE_PATTERN.findall(line)
        if len(dates) > 1:
            for dt in dates:
                if dt != admissao:
                    return dt
        return None

    def _extract_contrib_from_window(self, linhas: List[str], idx: int) -> Optional[float]:
        """Contribuição deve vir na linha imediatamente após a admissão."""
        target_idx = idx + 1
        if target_idx >= len(linhas):
            return None

        values = self._extract_money_values(linhas[target_idx])
        if not values:
            return None

        contrib = self._to_float_or_none(values[0])
        return self._sanitize_expected_range(contrib, "Contrib_Segurado", 50, 1000)

    def _extract_deposito_from_window(self, linhas: List[str], idx: int) -> Optional[float]:
        """Depósito FGTS deve vir na linha imediatamente após o CBO."""
        target_idx = idx + 1
        if target_idx >= len(linhas):
            return None

        values = self._extract_money_values(linhas[target_idx])
        if not values:
            return None

        return self._sanitize_expected_range(self._to_float_or_none(values[0]), "Deposito_FGTS", 1, 10000)


    def _sanitize_expected_range(
        self,
        value: Optional[float],
        field_name: str,
        soft_min: float,
        hard_max: float,
    ) -> Optional[float]:
        """Validação de consistência para reduzir capturas erradas de blocos."""
        if value is None:
            return None
        if value < 1 or value > hard_max:
            LOGGER.warning("Valor fora da faixa plausível (%s): %s", field_name, value)
            return None
        if value < soft_min:
            LOGGER.warning("Valor possivelmente inconsistente (%s): %s", field_name, value)
        return value

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
