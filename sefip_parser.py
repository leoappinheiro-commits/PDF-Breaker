#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import logging
import re
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence

import pandas as pd
from tqdm import tqdm

try:
    from charset_normalizer import from_path as detect_from_path
except Exception:  # pragma: no cover
    detect_from_path = None


LOGGER = logging.getLogger("sefip_parser")


@dataclass(slots=True)
class ParseError:
    line_number: int
    record_type: str
    reason: str
    content: str


class BaseRecord:
    record_type: str = ""

    @staticmethod
    def _slice(line: str, start: int, end: int) -> str:
        if start >= len(line):
            return ""
        return line[start:end].rstrip("\n\r")

    @staticmethod
    def _clean(value: str) -> str:
        return value.strip()

    @staticmethod
    def _digits(value: str) -> str:
        return re.sub(r"\D+", "", value or "")

    @staticmethod
    def _to_int(value: str) -> Optional[int]:
        digits = re.sub(r"[^0-9-]", "", value or "")
        if not digits:
            return None
        try:
            return int(digits)
        except ValueError:
            return None

    @staticmethod
    def _to_decimal(value: str, scale: int = 2) -> Optional[float]:
        digits = re.sub(r"\D+", "", value or "")
        if not digits:
            return None
        try:
            return int(digits) / (10 ** scale)
        except ValueError:
            return None

    @staticmethod
    def _to_date(value: str) -> Optional[date]:
        digits = re.sub(r"\D+", "", value or "")
        if len(digits) != 8:
            return None
        for fmt in ("%d%m%Y", "%Y%m%d"):
            try:
                return datetime.strptime(digits, fmt).date()
            except ValueError:
                continue
        return None

    def parse(self, line: str) -> Dict[str, Any]:
        raise NotImplementedError


class Record00(BaseRecord):
    record_type = "00"

    # The SEFIP layout has multiple variants by version. We parse fixed sections + regex fallbacks.
    def parse(self, line: str) -> Dict[str, Any]:
        cnpj = self._digits(self._slice(line, 51, 65)) or None
        competence_raw = self._clean(self._slice(line, 368, 374))
        email_raw = self._clean(self._slice(line, 330, 390))
        company_name = self._clean(self._slice(line, 65, 105))
        responsible = self._clean(self._slice(line, 105, 140))
        address = self._clean(self._slice(line, 140, 220))
        city = self._clean(self._slice(line, 240, 270))
        state = self._clean(self._slice(line, 270, 272))
        zip_code = self._digits(self._slice(line, 228, 236)) or None

        if not competence_raw:
            m_comp = re.search(r"\b(0[1-9]|1[0-2])\d{4}\b", line)
            competence_raw = m_comp.group(0) if m_comp else ""

        if not cnpj:
            m_cnpj = re.search(r"\b\d{14}\b", line)
            cnpj = m_cnpj.group(0) if m_cnpj else None

        if not email_raw:
            m_mail = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", line)
            email_raw = m_mail.group(0) if m_mail else ""

        return {
            "record_type": self.record_type,
            "cnpj": cnpj,
            "company_name": company_name or None,
            "address": address or None,
            "city": city or None,
            "state": state or None,
            "zip_code": zip_code,
            "responsible_person": responsible or None,
            "email": email_raw or None,
            "sefip_competence": competence_raw or None,
            "raw_line": line.rstrip("\n\r"),
        }


class Record10(BaseRecord):
    record_type = "10"

    def parse(self, line: str) -> Dict[str, Any]:
        return {
            "record_type": self.record_type,
            "cnpj": self._digits(self._slice(line, 2, 16)) or None,
            "cpf_responsible": self._digits(self._slice(line, 16, 27)) or None,
            "social_reason": self._clean(self._slice(line, 27, 67)) or None,
            "address": self._clean(self._slice(line, 67, 147)) or None,
            "city": self._clean(self._slice(line, 147, 177)) or None,
            "state": self._clean(self._slice(line, 177, 179)) or None,
            "zip_code": self._digits(self._slice(line, 179, 187)) or None,
            "raw_line": line.rstrip("\n\r"),
        }


class Record20(BaseRecord):
    record_type = "20"

    def parse(self, line: str) -> Dict[str, Any]:
        return {
            "record_type": self.record_type,
            "cnpj": self._digits(self._slice(line, 2, 16)) or None,
            "establishment_type": self._clean(self._slice(line, 16, 18)) or None,
            "fpas": self._clean(self._slice(line, 18, 21)) or None,
            "sat_rate": self._to_decimal(self._slice(line, 21, 25), scale=2),
            "third_party_code": self._clean(self._slice(line, 25, 29)) or None,
            "raw_line": line.rstrip("\n\r"),
        }


class Record30(BaseRecord):
    record_type = "30"

    def parse(self, line: str) -> Dict[str, Any]:
        raw = line.rstrip("\n\r")
        pis = self._digits(self._slice(line, 18, 29)) or None
        worker_name = self._clean(self._slice(line, 37, 107))
        admission_date = self._to_date(self._slice(line, 124, 132))
        birth_date = self._to_date(self._slice(line, 132, 140))
        category = self._clean(self._slice(line, 140, 142)) or None
        worker_type = self._clean(self._slice(line, 142, 144)) or None
        remuneration = self._to_decimal(self._slice(line, 144, 159))
        fgts_base = self._to_decimal(self._slice(line, 159, 174))
        salary_13_base = self._to_decimal(self._slice(line, 174, 189))
        movement_code = self._clean(self._slice(line, 189, 191)) or None
        employment_link_type = self._clean(self._slice(line, 191, 193)) or None
        fgts_value = self._to_decimal(self._slice(line, 193, 208))
        contribution_base = self._to_decimal(self._slice(line, 208, 223))
        employment_status = self._clean(self._slice(line, 223, 225)) or None

        if not pis:
            m_pis = re.search(r"\b\d{11}\b", raw)
            pis = m_pis.group(0) if m_pis else None

        if not worker_name:
            # fallback: name chunk between dates / numeric zones
            m_name = re.search(r"\d{11}\s*\d{8}(.{10,80}?)\d{8}", raw)
            worker_name = m_name.group(1).strip() if m_name else ""

        return {
            "record_type": self.record_type,
            "pis": pis,
            "worker_name": worker_name or None,
            "admission_date": admission_date,
            "birth_date": birth_date,
            "category": category,
            "worker_type": worker_type,
            "remuneration": remuneration,
            "fgts_base": fgts_base,
            "salary_13_base": salary_13_base,
            "movement_code": movement_code,
            "employment_link_type": employment_link_type,
            "fgts_value": fgts_value,
            "contribution_base": contribution_base,
            "employment_status": employment_status,
            "raw_line": raw,
        }


class Record40(BaseRecord):
    record_type = "40"

    def parse(self, line: str) -> Dict[str, Any]:
        return {
            "record_type": self.record_type,
            "cnpj": self._digits(self._slice(line, 2, 16)) or None,
            "inss_base": self._to_decimal(self._slice(line, 16, 31)),
            "inss_due": self._to_decimal(self._slice(line, 31, 46)),
            "rat_due": self._to_decimal(self._slice(line, 46, 61)),
            "third_party_due": self._to_decimal(self._slice(line, 61, 76)),
            "raw_line": line.rstrip("\n\r"),
        }


class Record50(BaseRecord):
    record_type = "50"

    def parse(self, line: str) -> Dict[str, Any]:
        return {
            "record_type": self.record_type,
            "cnpj": self._digits(self._slice(line, 2, 16)) or None,
            "pis": self._digits(self._slice(line, 16, 27)) or None,
            "movement_code": self._clean(self._slice(line, 27, 29)) or None,
            "movement_date": self._to_date(self._slice(line, 29, 37)),
            "fgts_movement_base": self._to_decimal(self._slice(line, 37, 52)),
            "fgts_movement_value": self._to_decimal(self._slice(line, 52, 67)),
            "raw_line": line.rstrip("\n\r"),
        }


class Record90(BaseRecord):
    record_type = "90"

    def parse(self, line: str) -> Dict[str, Any]:
        return {
            "record_type": self.record_type,
            "total_records": self._to_int(self._slice(line, 2, 11)),
            "generation_date": self._to_date(self._slice(line, 11, 19)),
            "raw_line": line.rstrip("\n\r"),
        }


class SEFIPParser:
    def __init__(self, input_path: Path, output_dir: Path) -> None:
        self.input_path = input_path
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.errors: List[ParseError] = []
        self.total_lines = 0

        self.record_parsers: Dict[str, BaseRecord] = {
            "00": Record00(),
            "10": Record10(),
            "20": Record20(),
            "30": Record30(),
            "40": Record40(),
            "50": Record50(),
            "90": Record90(),
        }

    def detect_encoding(self) -> str:
        if detect_from_path is not None:
            results = detect_from_path(str(self.input_path))
            best = results.best()
            if best and best.encoding:
                return best.encoding
        # robust fallback for legacy government files
        for encoding in ("latin-1", "cp1252", "utf-8"):
            try:
                with self.input_path.open("r", encoding=encoding) as handle:
                    handle.read(1024)
                return encoding
            except UnicodeDecodeError:
                continue
        return "latin-1"

    def iter_lines(self, encoding: str) -> Iterator[tuple[int, str]]:
        with self.input_path.open("r", encoding=encoding, errors="replace") as handle:
            for line_number, line in enumerate(handle, start=1):
                yield line_number, line.rstrip("\n")

    def parse(self) -> Dict[str, pd.DataFrame]:
        encoding = self.detect_encoding()
        LOGGER.info("Detected encoding: %s", encoding)

        header_rows: List[Dict[str, Any]] = []
        company_rows: List[Dict[str, Any]] = []
        establishment_rows: List[Dict[str, Any]] = []
        employee_rows: List[Dict[str, Any]] = []
        financial_rows: List[Dict[str, Any]] = []
        movement_rows: List[Dict[str, Any]] = []
        trailer_rows: List[Dict[str, Any]] = []

        for line_number, line in tqdm(self.iter_lines(encoding), desc="Parsing SEFIP", unit="lines"):
            self.total_lines += 1
            if not line:
                continue
            record_type = line[:2]
            parser = self.record_parsers.get(record_type)

            if parser is None:
                self.errors.append(ParseError(line_number, record_type, "Unknown record type", line))
                continue

            try:
                parsed = parser.parse(line)
                parsed["line_number"] = line_number
            except Exception as exc:  # defensive parsing
                self.errors.append(ParseError(line_number, record_type, f"Parse failure: {exc}", line))
                continue

            if record_type == "00":
                header_rows.append(parsed)
                company_rows.append(parsed)
            elif record_type == "10":
                company_rows.append(parsed)
            elif record_type == "20":
                establishment_rows.append(parsed)
            elif record_type == "30":
                employee_rows.append(parsed)
            elif record_type == "40":
                financial_rows.append(parsed)
            elif record_type == "50":
                movement_rows.append(parsed)
            elif record_type == "90":
                trailer_rows.append(parsed)

        df_header = pd.DataFrame(header_rows)
        df_company = pd.DataFrame(company_rows)
        df_establishment = pd.DataFrame(establishment_rows)
        df_employees = pd.DataFrame(employee_rows)
        df_financial = pd.DataFrame(financial_rows)
        df_movements = pd.DataFrame(movement_rows)
        df_trailer = pd.DataFrame(trailer_rows)

        self._post_process_frames(
            df_header, df_company, df_establishment, df_employees, df_financial, df_movements, df_trailer
        )

        return {
            "df_header": df_header,
            "df_company": df_company,
            "df_establishment": df_establishment,
            "df_employees": df_employees,
            "df_financial": df_financial,
            "df_movements": df_movements,
            "df_trailer": df_trailer,
        }

    @staticmethod
    def _post_process_frames(*frames: pd.DataFrame) -> None:
        for frame in frames:
            if frame.empty:
                continue
            for col in frame.columns:
                if frame[col].dtype == object:
                    frame[col] = frame[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    def export(self, datasets: Dict[str, pd.DataFrame]) -> None:
        df_employees = datasets["df_employees"]
        df_company = datasets["df_company"]
        df_movements = datasets["df_movements"]

        df_employees.to_excel(self.output_dir / "employees_sefip.xlsx", index=False)
        df_employees.to_csv(self.output_dir / "employees_sefip.csv", index=False, encoding="utf-8-sig")
        df_company.to_excel(self.output_dir / "company_info.xlsx", index=False)
        df_movements.to_excel(self.output_dir / "fgts_movements.xlsx", index=False)

        full_dataset = pd.concat(
            [
                datasets["df_header"].assign(dataset="header"),
                datasets["df_company"].assign(dataset="company"),
                datasets["df_establishment"].assign(dataset="establishment"),
                datasets["df_employees"].assign(dataset="employees"),
                datasets["df_financial"].assign(dataset="financial"),
                datasets["df_movements"].assign(dataset="movements"),
                datasets["df_trailer"].assign(dataset="trailer"),
            ],
            ignore_index=True,
            sort=False,
        )
        full_dataset.to_parquet(self.output_dir / "sefip_full_database.parquet", index=False)

    def write_error_report(self) -> Path:
        error_path = self.output_dir / "sefip_error_report.jsonl"
        with error_path.open("w", encoding="utf-8") as handle:
            for error in self.errors:
                handle.write(
                    json.dumps(
                        {
                            "line_number": error.line_number,
                            "record_type": error.record_type,
                            "reason": error.reason,
                            "content": error.content,
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        return error_path


def configure_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = output_dir / "sefip_parser.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler()],
    )


def build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="High-performance parser for SEFIP .RE files")
    parser.add_argument("input_file", type=Path, help="Path to input .RE file")
    parser.add_argument("output_folder", type=Path, help="Output folder path")
    return parser


def main() -> None:
    args = build_cli().parse_args()
    configure_logging(args.output_folder)

    parser = SEFIPParser(input_path=args.input_file, output_dir=args.output_folder)
    datasets = parser.parse()
    parser.export(datasets)
    error_file = parser.write_error_report()

    total_employees = len(datasets["df_employees"])
    total_employers = len(datasets["df_company"])
    total_movements = len(datasets["df_movements"])

    LOGGER.info("Error report generated: %s", error_file)

    print(f"Total lines processed: {parser.total_lines}")
    print(f"Total employees parsed: {total_employees}")
    print(f"Total employers parsed: {total_employers}")
    print(f"Total movements parsed: {total_movements}")


if __name__ == "__main__":
    main()
