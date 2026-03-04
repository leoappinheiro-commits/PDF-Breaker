from __future__ import annotations

from typing import List

import cv2
import numpy as np
import pytesseract


class OcrEngine:
    """Executa OCR com pré-processamento robusto para documentos fiscais."""

    def __init__(self, lang: str = "por", tesseract_cmd: str | None = None) -> None:
        self.lang = lang
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    def extract_texts(self, pixmaps: List[object]) -> List[str]:
        texts: List[str] = []
        for pixmap in pixmaps:
            img = self._pixmap_to_ndarray(pixmap)
            preprocessed = self._preprocess(img)
            config = r"--oem 3 --psm 6"
            text = pytesseract.image_to_string(preprocessed, lang=self.lang, config=config)
            texts.append(text)
        return texts

    @staticmethod
    def _pixmap_to_ndarray(pixmap: object) -> np.ndarray:
        array = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, pixmap.n)
        if pixmap.n == 4:
            return cv2.cvtColor(array, cv2.COLOR_BGRA2BGR)
        return array

    def _preprocess(self, img: np.ndarray) -> np.ndarray:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        denoised = cv2.fastNlMeansDenoising(gray, h=12)
        threshold = cv2.adaptiveThreshold(
            denoised,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11,
        )
        deskewed = self._deskew(threshold)
        return deskewed

    @staticmethod
    def _deskew(img: np.ndarray) -> np.ndarray:
        coords = np.column_stack(np.where(img < 255))
        if coords.size == 0:
            return img

        angle = cv2.minAreaRect(coords)[-1]
        angle = -(90 + angle) if angle < -45 else -angle

        h, w = img.shape[:2]
        center = (w // 2, h // 2)
        matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(
            img,
            matrix,
            (w, h),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_REPLICATE,
        )
        return rotated
