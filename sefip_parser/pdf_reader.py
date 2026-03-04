from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

import fitz  # PyMuPDF
import pdfplumber


@dataclass
class PdfPageContent:
    """Representa o conteúdo textual de uma página."""

    page_number: int
    text: str


@dataclass
class PdfLoadResult:
    """Resultado da leitura inicial de um PDF."""

    file_path: Path
    page_texts: List[PdfPageContent]
    needs_ocr: bool


class PdfReader:
    """Lê PDFs e detecta se há texto nativo suficiente para evitar OCR."""

    def __init__(self, min_chars_per_page: int = 30, text_page_ratio: float = 0.6) -> None:
        self.min_chars_per_page = min_chars_per_page
        self.text_page_ratio = text_page_ratio

    def load(self, file_path: Path) -> PdfLoadResult:
        page_texts: List[PdfPageContent] = []

        with pdfplumber.open(str(file_path)) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = (page.extract_text() or "").strip()
                page_texts.append(PdfPageContent(page_number=i, text=text))

        pages_with_text = sum(1 for p in page_texts if len(p.text) >= self.min_chars_per_page)
        ratio = pages_with_text / max(1, len(page_texts))
        needs_ocr = ratio < self.text_page_ratio

        return PdfLoadResult(file_path=file_path, page_texts=page_texts, needs_ocr=needs_ocr)

    @staticmethod
    def render_pages_as_images(file_path: Path, dpi: int = 300) -> List["fitz.Pixmap"]:
        """Renderiza todas as páginas em alta resolução para OCR."""
        doc = fitz.open(str(file_path))
        scale = dpi / 72
        matrix = fitz.Matrix(scale, scale)
        images = [page.get_pixmap(matrix=matrix, alpha=False) for page in doc]
        doc.close()
        return images
