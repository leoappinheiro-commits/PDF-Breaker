from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable

from rapidfuzz import fuzz

ANCHORS: Dict[str, list[str]] = {
    "empresa": ["EMPRESA", "CNPJ", "COMPETENCIA", "CÓDIGO DE RECOLHIMENTO", "FPAS"],
    "trabalhadores": ["RELAÇÃO DE TRABALHADORES", "PIS", "NOME", "REMUNERACAO", "BASE FGTS"],
    "resumo": ["RESUMO", "TOTAL", "FGTS", "INSS", "TOTAL TRABALHADORES"],
}


@dataclass
class AnchorScore:
    section: str
    score: float


class AnchorDetector:
    """Classifica textos por seção usando anchors e fuzzy matching."""

    def __init__(self, threshold: float = 65.0) -> None:
        self.threshold = threshold

    def classify(self, text: str) -> AnchorScore:
        normalized = self._normalize(text)
        best = AnchorScore(section="desconhecida", score=0.0)

        for section, anchors in ANCHORS.items():
            score = self._score_section(normalized, anchors)
            if score > best.score:
                best = AnchorScore(section=section, score=score)

        if best.score < self.threshold:
            return AnchorScore(section="desconhecida", score=best.score)
        return best

    def _score_section(self, text: str, anchors: Iterable[str]) -> float:
        scores = []
        for anchor in anchors:
            anchor_norm = self._normalize(anchor)
            full = fuzz.partial_ratio(anchor_norm, text)
            token = fuzz.token_set_ratio(anchor_norm, text)
            scores.append(max(full, token))
        return sum(scores) / max(1, len(scores))

    @staticmethod
    def _normalize(text: str) -> str:
        return " ".join(text.upper().replace("Ç", "C").replace("Ã", "A").replace("Á", "A").split())
