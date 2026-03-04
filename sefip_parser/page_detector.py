from __future__ import annotations

from dataclasses import dataclass
from typing import List

from .anchor_detector import AnchorDetector


@dataclass
class ClassifiedPage:
    page_number: int
    section: str
    text: str
    score: float


class PageDetector:
    """Detecta seção predominante de cada página da SEFIP."""

    def __init__(self, anchor_detector: AnchorDetector | None = None) -> None:
        self.anchor_detector = anchor_detector or AnchorDetector()

    def classify_pages(self, page_texts: List[str]) -> List[ClassifiedPage]:
        output: List[ClassifiedPage] = []
        for idx, text in enumerate(page_texts, start=1):
            anchor_score = self.anchor_detector.classify(text)
            output.append(
                ClassifiedPage(
                    page_number=idx,
                    section=anchor_score.section,
                    text=text,
                    score=anchor_score.score,
                )
            )
        return output
