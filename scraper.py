"""
Web scraper para ONGs da Paraíba no site ONGs Brasil.
Fluxo: Estado (PB) -> Cidades -> Lista de ONGs -> Página de detalhes -> Planilha Excel.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.ongsbrasil.com.br/"
STATE_URL = (
    "https://www.ongsbrasil.com.br/default.asp"
    "?Pag=1&Destino=Instituicoes&Estado=PB"
)

LABEL_TO_KEY = {
    "endereco": "endereco",
    "bairro": "bairro",
    "cep": "cep",
    "cidade": "cidade",
    "estado": "estado",
    "pais": "pais",
    "telefone": "telefone",
    "nome fantasia": "nome_fantasia",
    "email": "email",
    "site": "site",
    "cnpj": "cnpj",
}

OUTPUT_COLUMNS = [
    "codigo_instituicao",
    "nome",
    "cidade_listagem",
    "bairro_listagem",
    "endereco",
    "bairro",
    "cep",
    "cidade",
    "estado",
    "pais",
    "telefone",
    "nome_fantasia",
    "email",
    "site",
    "cnpj",
    "url_detalhe",
]

DEFAULT_DELAY = 1.0
DEFAULT_OUTPUT = "ongs_paraiba.xlsx"
DEFAULT_CHECKPOINT = "checkpoint.json"


def normalize_label(text: str) -> str:
    text = text.strip().lower()
    text = text.replace(":", "")
    replacements = {
        "endereço": "endereco",
        "endereco": "endereco",
        "país": "pais",
        "pais": "pais",
        "e-mail": "email",
        "email": "email",
    }
    return replacements.get(text, text)


def decode_response(response: requests.Response) -> str:
    if response.encoding and response.encoding.lower() not in ("iso-8859-1", "latin-1"):
        return response.text
    return response.content.decode("latin-1", errors="replace")


class OngsBrasilScraper:
    def __init__(
        self,
        delay: float = DEFAULT_DELAY,
        session: requests.Session | None = None,
    ):
        self.delay = delay
        self.session = session or requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Accept-Language": "pt-BR,pt;q=0.9",
            }
        )
        self.logger = logging.getLogger(__name__)

    def fetch(self, url: str) -> BeautifulSoup:
        time.sleep(self.delay)
        response = self.session.get(url, timeout=60)
        response.raise_for_status()
        html = decode_response(response)
        return BeautifulSoup(html, "lxml")

    def get_cities(self) -> list[dict[str, str]]:
        soup = self.fetch(STATE_URL)
        cities: list[dict[str, str]] = []
        seen: set[str] = set()

        for link in soup.find_all("a", href=True):
            href = link["href"]
            if "Estado=PB" not in href or "Cidade=" not in href:
                continue
            full_url = urljoin(BASE_URL, href)
            parsed = urlparse(full_url)
            params = parse_qs(parsed.query)
            cidade = params.get("Cidade", params.get("cidade", [None]))[0]
            if not cidade or cidade in seen:
                continue
            seen.add(cidade)
            cities.append({"nome": unquote(cidade), "url": full_url})

        cities.sort(key=lambda c: c["nome"].lower())
        self.logger.info("Encontradas %d cidades na Paraíba", len(cities))
        return cities

    def get_city_page_urls(self, city_url: str) -> list[str]:
        soup = self.fetch(city_url)
        pages = {city_url}

        for link in soup.find_all("a", href=True):
            href = link["href"]
            if "PageNo=" not in href or "Destino=Instituicoes" not in href:
                continue
            pages.add(urljoin(BASE_URL, href))

        return sorted(pages)

    def get_ongs_from_city_page(self, soup: BeautifulSoup) -> list[dict[str, str]]:
        ongs: list[dict[str, str]] = []
        seen_codes: set[str] = set()

        for link in soup.find_all("a", href=True):
            href = link["href"]
            if "CodigoInstituicao=" not in href:
                continue

            full_url = urljoin(BASE_URL, href)
            parsed = urlparse(full_url)
            params = parse_qs(parsed.query)
            codigo = params.get("CodigoInstituicao", [None])[0]
            if not codigo or codigo in seen_codes:
                continue

            if "mais info" in link.get_text(strip=True).lower():
                nome = params.get("Instituicao", [""])[0]
                nome = unquote(nome.replace("-", " "))
            else:
                h2 = link.find("h2")
                nome = h2.get_text(strip=True) if h2 else link.get_text(strip=True)

            seen_codes.add(codigo)
            ongs.append(
                {
                    "codigo_instituicao": codigo,
                    "nome_listagem": nome,
                    "url_detalhe": full_url,
                }
            )

        bairro_pattern = re.compile(r"Bairro:\s*([^\-]+)", re.I)
        bairro_by_index: list[str] = []
        for p in soup.find_all("p", class_="text-capitalize"):
            if "Bairro:" not in p.get_text():
                continue
            match = bairro_pattern.search(p.get_text(" ", strip=True))
            bairro_by_index.append(match.group(1).strip() if match else "")

        for i, ong in enumerate(ongs):
            if i < len(bairro_by_index) and bairro_by_index[i]:
                ong["bairro_listagem"] = bairro_by_index[i].strip()
            else:
                ong["bairro_listagem"] = ""

        return ongs

    def parse_detail_page(self, soup: BeautifulSoup, url: str) -> dict[str, str]:
        data: dict[str, str] = {"url_detalhe": url}

        h1 = soup.find("h1", itemprop="name")
        data["nome"] = h1.get_text(strip=True) if h1 else ""

        table = soup.find("table", class_="table")
        if table:
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) < 2:
                    continue
                label = normalize_label(cells[0].get_text(strip=True))
                key = LABEL_TO_KEY.get(label)
                if not key:
                    continue

                value_cell = cells[1]
                input_tag = value_cell.find("input")
                if input_tag and input_tag.get("value"):
                    value = input_tag["value"].strip()
                else:
                    value = value_cell.get_text(strip=True)
                data[key] = value

        return data

    def scrape_ong_detail(self, url: str) -> dict[str, str]:
        soup = self.fetch(url)
        return self.parse_detail_page(soup, url)

    def scrape_city(self, city: dict[str, str]) -> list[dict[str, str]]:
        self.logger.info("Cidade: %s", city["nome"])
        results: list[dict[str, str]] = []

        for page_url in self.get_city_page_urls(city["url"]):
            soup = self.fetch(page_url)
            ongs = self.get_ongs_from_city_page(soup)
            self.logger.info("  Página: %d ONG(s) encontrada(s)", len(ongs))

            for ong in ongs:
                try:
                    detail = self.scrape_ong_detail(ong["url_detalhe"])
                except requests.RequestException as exc:
                    self.logger.error(
                        "  Erro ao buscar ONG %s: %s",
                        ong["codigo_instituicao"],
                        exc,
                    )
                    continue

                row = {
                    "codigo_instituicao": ong["codigo_instituicao"],
                    "nome": detail.get("nome") or ong.get("nome_listagem", ""),
                    "cidade_listagem": city["nome"],
                    "bairro_listagem": ong.get("bairro_listagem", ""),
                    "endereco": detail.get("endereco", ""),
                    "bairro": detail.get("bairro", ""),
                    "cep": detail.get("cep", ""),
                    "cidade": detail.get("cidade", ""),
                    "estado": detail.get("estado", ""),
                    "pais": detail.get("pais", ""),
                    "telefone": detail.get("telefone", ""),
                    "nome_fantasia": detail.get("nome_fantasia", ""),
                    "email": detail.get("email", ""),
                    "site": detail.get("site", ""),
                    "cnpj": detail.get("cnpj", ""),
                    "url_detalhe": ong["url_detalhe"],
                }
                results.append(row)
                self.logger.info("    OK: %s", row["nome"][:60])

        return results

    def scrape_all(
        self,
        output_path: Path,
        checkpoint_path: Path,
        resume: bool = True,
    ) -> pd.DataFrame:
        checkpoint = self._load_checkpoint(checkpoint_path) if resume else {}
        scraped_codes: set[str] = set(checkpoint.get("scraped_codes", []))
        all_rows: list[dict] = checkpoint.get("rows", [])

        cities = self.get_cities()
        completed_cities: set[str] = set(checkpoint.get("completed_cities", []))

        for city in cities:
            if city["nome"] in completed_cities:
                self.logger.info("Cidade já processada (pulando): %s", city["nome"])
                continue

            city_rows = self.scrape_city(city)
            new_rows = [
                r
                for r in city_rows
                if r["codigo_instituicao"] not in scraped_codes
            ]

            for row in new_rows:
                scraped_codes.add(row["codigo_instituicao"])

            all_rows.extend(new_rows)
            completed_cities.add(city["nome"])

            self._save_checkpoint(
                checkpoint_path,
                {
                    "scraped_codes": list(scraped_codes),
                    "completed_cities": list(completed_cities),
                    "rows": all_rows,
                },
            )
            self._save_excel(output_path, all_rows)

        df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
        self.logger.info("Concluído: %d ONGs salvas em %s", len(df), output_path)
        return df

    @staticmethod
    def _load_checkpoint(path: Path) -> dict:
        if not path.exists():
            return {}
        with path.open(encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def _save_checkpoint(path: Path, data: dict) -> None:
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @staticmethod
    def _save_excel(path: Path, rows: list[dict]) -> None:
        df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
        rename = {
            "codigo_instituicao": "Código",
            "nome": "Nome",
            "cidade_listagem": "Cidade (listagem)",
            "bairro_listagem": "Bairro (listagem)",
            "endereco": "Endereço",
            "bairro": "Bairro",
            "cep": "CEP",
            "cidade": "Cidade",
            "estado": "Estado",
            "pais": "País",
            "telefone": "Telefone",
            "nome_fantasia": "Nome Fantasia",
            "email": "E-mail",
            "site": "Site",
            "cnpj": "CNPJ",
            "url_detalhe": "URL",
        }
        df = df.rename(columns=rename)
        df.to_excel(path, index=False, sheet_name="ONGs Paraíba")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extrai todas as ONGs da Paraíba do site ONGs Brasil."
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"Arquivo Excel de saída (padrão: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"Segundos entre requisições (padrão: {DEFAULT_DELAY})",
    )
    parser.add_argument(
        "--checkpoint",
        default=DEFAULT_CHECKPOINT,
        help=f"Arquivo de progresso (padrão: {DEFAULT_CHECKPOINT})",
    )
    parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Reinicia do zero, ignorando checkpoint anterior",
    )
    parser.add_argument(
        "--cidade",
        help="Processa apenas uma cidade (nome exato, ex: 'Joao Pessoa')",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Log detalhado",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    scraper = OngsBrasilScraper(delay=args.delay)
    output_path = Path(args.output)
    checkpoint_path = Path(args.checkpoint)

    if args.cidade:
        cities = scraper.get_cities()
        match = next((c for c in cities if c["nome"].lower() == args.cidade.lower()), None)
        if not match:
            raise SystemExit(f"Cidade não encontrada: {args.cidade}")
        rows = scraper.scrape_city(match)
        scraper._save_excel(output_path, rows)
        logging.info("Salvo: %s (%d ONGs)", output_path, len(rows))
        return

    if args.no_resume and checkpoint_path.exists():
        checkpoint_path.unlink()

    scraper.scrape_all(
        output_path=output_path,
        checkpoint_path=checkpoint_path,
        resume=not args.no_resume,
    )


if __name__ == "__main__":
    main()
