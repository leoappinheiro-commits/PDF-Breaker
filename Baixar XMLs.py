import argparse
import datetime
import logging
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


@dataclass(frozen=True)
class DownloadConfig:
    input_file: Path
    output_dir: Path
    download_dir: Path
    url: str = "https://consultadanfe.com/#"
    key_column: str = "Chaves_Coluna"
    wait_seconds: int = 5


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )


def parse_args() -> DownloadConfig:
    parser = argparse.ArgumentParser(
        description="Baixa XMLs do Consulta DANFE com base em uma planilha de chaves."
    )
    parser.add_argument(
        "--input",
        default=r"C:\Users\leonardopinheiro\Documents\Lista de Chaves.xlsx",
        help="Caminho da planilha de entrada.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Diretório de saída para os XMLs processados. Padrão: pasta de download.",
    )
    parser.add_argument(
        "--download-dir",
        default=r"C:\Users\leonardopinheiro\Downloads",
        help="Diretório padrão de downloads do navegador.",
    )

    args = parser.parse_args()
    download_dir = Path(args.download_dir)
    output_dir = Path(args.output) if args.output else download_dir

    return DownloadConfig(
        input_file=Path(args.input),
        output_dir=output_dir,
        download_dir=download_dir,
    )


def _resolve_chromedriver_binary(installed_path: str) -> Path:
    """
    Garante que o caminho retornado pelo webdriver_manager aponte para o executável.

    Em algumas versões, o manager pode retornar um caminho não executável
    (ex.: THIRD_PARTY_NOTICES), causando WinError 193 no Windows.
    """
    path = Path(installed_path)
    if path.is_file() and path.suffix.lower() == ".exe":
        return path

    candidates = []
    if path.is_file():
        candidates = list(path.parent.rglob("chromedriver*.exe"))
    elif path.is_dir():
        candidates = list(path.rglob("chromedriver*.exe"))

    if not candidates:
        raise FileNotFoundError(
            f"Não foi possível localizar chromedriver.exe a partir de: {installed_path}"
        )

    return max(candidates, key=lambda candidate: candidate.stat().st_mtime)


def initialize_selenium_driver() -> webdriver.Chrome:
    logger = logging.getLogger("selenium_init")
    try:
        chrome_options = Options()
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--safebrowsing-disable-download-protection")
        chrome_options.add_argument("--no-sandbox")

        raw_driver_path = ChromeDriverManager().install()
        driver_binary = _resolve_chromedriver_binary(raw_driver_path)
        logger.info("driver_binary_resolved=true path=%s", driver_binary)

        service = Service(str(driver_binary))
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logger.info("driver_initialized=true")
        return driver
    except (WebDriverException, OSError, FileNotFoundError) as exc:
        logger.exception("driver_initialized=false error=%s", exc)
        raise


def latest_downloaded_file(download_dir: Path) -> Path:
    files = [p for p in download_dir.iterdir() if p.is_file()]
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em: {download_dir}")
    return max(files, key=lambda path: path.stat().st_ctime)


def postprocess_downloaded_file(download_dir: Path, output_dir: Path) -> Path:
    logger = logging.getLogger("postprocess")
    downloaded_file = latest_downloaded_file(download_dir)
    date_suffix = datetime.datetime.now().strftime("%d-%m-%Y")
    destination_dir = output_dir / f"XML {date_suffix}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination_file = destination_dir / downloaded_file.name

    try:
        moved_file = Path(shutil.move(str(downloaded_file), str(destination_file)))
        logger.info(
            "file_moved=true source=%s destination=%s",
            downloaded_file,
            moved_file,
        )
        return moved_file
    except OSError as exc:
        logger.exception(
            "file_moved=false source=%s destination=%s error=%s",
            downloaded_file,
            destination_file,
            exc,
        )
        raise


def process_downloads(config: DownloadConfig) -> None:
    logger = logging.getLogger("download_logic")

    if not config.input_file.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {config.input_file}")
    if not config.download_dir.exists():
        raise FileNotFoundError(f"Diretório de download não encontrado: {config.download_dir}")

    keys_df = pd.read_excel(config.input_file)
    if config.key_column not in keys_df.columns:
        raise KeyError(f"Coluna '{config.key_column}' não encontrada na planilha.")

    browser = initialize_selenium_driver()
    try:
        browser.get(config.url)
        logger.info("page_opened=true url=%s", config.url)

        for index, row in keys_df.iterrows():
            key_value = row[config.key_column]
            logger.info("processing_key=true index=%s key=%s", index, key_value)

            try:
                browser.find_element(By.ID, "chave").send_keys(key_value)
                time.sleep(config.wait_seconds)
                browser.find_element(By.CLASS_NAME, "g-recaptcha").click()
                time.sleep(config.wait_seconds)
                browser.find_element(
                    By.XPATH, '//*[@id="modalNFe"]/div/div/div[2]/div/div[3]/p[2]/a'
                ).click()
                time.sleep(config.wait_seconds)
                browser.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[3]/a').click()
                time.sleep(config.wait_seconds)
                browser.find_element(By.ID, "chave").clear()
                time.sleep(config.wait_seconds)

                postprocess_downloaded_file(config.download_dir, config.output_dir)
            except (NoSuchElementException, WebDriverException, OSError, FileNotFoundError) as exc:
                logger.exception(
                    "processing_key=false index=%s key=%s error=%s",
                    index,
                    key_value,
                    exc,
                )
                continue
    finally:
        browser.quit()
        logger.info("browser_closed=true")


def main() -> int:
    setup_logging()
    logger = logging.getLogger("main")

    try:
        config = parse_args()
        logger.info(
            "config_loaded=true input=%s output=%s download_dir=%s",
            config.input_file,
            config.output_dir,
            config.download_dir,
        )
        process_downloads(config)
        return 0
    except Exception as exc:  # captura final para retorno com código de erro
        logger.exception("execution_failed=true error=%s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
