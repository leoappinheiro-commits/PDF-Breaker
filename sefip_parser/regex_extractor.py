from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class RegexExtraction:
    empresa: Dict[str, Optional[str]]
    resumo: Dict[str, Optional[float]]


class RegexExtractor:
    """Extrai campos do texto da SEFIP com regex resiliente."""

    PATTERNS = {
        "cnpj": re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b"),
        "competencia": re.compile(r"\b\d{2}/\d{4}\b"),
        "codigo_recolhimento": re.compile(r"COD(?:IGO)?\s*DE\s*RECOLHIMENTO\s*[:\-]?\s*(\d{3,4})", re.IGNORECASE),
        "fpas": re.compile(r"\bFPAS\s*[:\-]?\s*(\d{3})\b", re.IGNORECASE),
        "valor": re.compile(r"\b(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}\b"),
        "total_trabalhadores": re.compile(r"TOTAL\s+TRABALHADORES\s*[:\-]?\s*(\d+)", re.IGNORECASE),
        "empresa": re.compile(r"EMPRESA\s*[:\-]?\s*(.+)", re.IGNORECASE),
    }

    def extract(self, text: str) -> RegexExtraction:
        empresa = {
            "Empresa": self._search_group("empresa", text),
            "CNPJ": self._search("cnpj", text),
            "Competencia": self._search("competencia", text),
            "Codigo_Recolhimento": self._search_group("codigo_recolhimento", text),
            "FPAS": self._search_group("fpas", text),
        }

        resumo = {
            "Total_Remuneracao": self._search_value_after_label(text, "TOTAL REMUNERACAO"),
            "Total_FGTS": self._search_value_after_label(text, "TOTAL FGTS"),
            "Total_INSS": self._search_value_after_label(text, "TOTAL INSS"),
            "Total_Trabalhadores": self._parse_int(self._search_group("total_trabalhadores", text)),
        }

        return RegexExtraction(empresa=empresa, resumo=resumo)

    def _search(self, pattern_name: str, text: str) -> Optional[str]:
        match = self.PATTERNS[pattern_name].search(text)
        return match.group(0) if match else None

    def _search_group(self, pattern_name: str, text: str) -> Optional[str]:
        match = self.PATTERNS[pattern_name].search(text)
        return match.group(1).strip() if match else None

    def _search_value_after_label(self, text: str, label: str) -> Optional[float]:
        pattern = re.compile(rf"{label}\s*[:\-]?\s*((?:\d{{1,3}}(?:\.\d{{3}})*|\d+),\d{{2}})", re.IGNORECASE)
        match = pattern.search(text)
        if not match:
            return None
        return self.normalize_brl_number(match.group(1))

    @staticmethod
    def normalize_brl_number(value: str | None) -> Optional[float]:
        if not value:
            return None
        normalized = value.replace(".", "").replace(",", ".")
        return float(normalized)

    @staticmethod
    def _parse_int(value: str | None) -> Optional[int]:
        return int(value) if value and value.isdigit() else None
