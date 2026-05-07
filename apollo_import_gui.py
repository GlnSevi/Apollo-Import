from __future__ import annotations

import csv
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from html import unescape
import io
import json
import os
from pathlib import Path
import random
import re
import string
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Callable
from urllib import error as urllib_error
from urllib import request as urllib_request
from urllib.parse import parse_qs, quote, unquote, urlencode, urlparse
import xml.etree.ElementTree as ET

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

try:
    from PIL import Image, ImageTk
except ImportError:  # pragma: no cover - optional preview dependency
    Image = None
    ImageTk = None

try:
    import pymupdf
except ImportError:  # pragma: no cover - optional preview dependency
    try:
        import fitz as pymupdf  # type: ignore[no-redef]
    except ImportError:  # pragma: no cover - optional preview dependency
        pymupdf = None


APP_TITLE = "Apollo Import GUI Prototype"
DEFAULT_IMPORT_DIR = Path(r"C:\Users\heimbuchner\Desktop\Apollo Import App\Aktuelle Import Datein")
DEFAULT_OUTPUT_DIR = Path.cwd() / "output"
DEEPL_DEFAULT_BASE_URL = "https://api.deepl.com"
ID_ALPHABET = string.ascii_uppercase + string.digits
ID_LENGTH = 6
ATTACHMENT_FORMAT_TYPE_HEADER = "TecDoc Anhangsformattyp ID"
LAST_WRITTEN_HEADER = "Zuletzt geschrieben am"

SHORT_TEXT_HEADERS = [
    "Artikelnummer",
    "Text Modul ID",
    "Modultyp ID",
    "Infotyp ID",
    "Language ID1",
    "TEXT UNI",
    "Language ID2",
    "TEXT DE",
    "Language ID3",
    "TEXT EN",
    "Language ID4",
    "TEXT CZ",
    "Language ID5",
    "TEXT FR",
    "Language ID6",
    "TEXT IT",
    "Language ID7",
    "TEXT NL",
    LAST_WRITTEN_HEADER,
]

SHORT_TEXT_FILE = ("Kurzbezeichnung-NEU.xlsx", "Import")
SHORT_MAPPING_FILE = ("Kurzbezeichnung_zu_ID.xlsx", "Zuordnung")
LONG_TEXT_FILE = ("Text-NEU.xlsx", "Import")
IMAGE_FILE = ("Bilder.xlsx", "Sheet1")
DOCUMENT_FILE = ("Dokumente.xlsx", "Sheet1")
VIDEO_FILE = ("Videos.xlsx", "Tabelle1")
WEB_LINK_FILE = ("Web Link.xlsx", "Tabelle1")

IMAGE_HEADERS = ["Artikelnummer", "BILDPFAD", "Art", "Sprache", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
DOCUMENT_HEADERS = ["Artikelnummer", "Pfad zum Dokument", "Sprache", "Art", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
VIDEO_HEADERS = ["Produktnummer", "Link", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
WEB_HEADERS = ["Artikelnummer ", "Link", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
SHORT_MAPPING_HEADERS = ["Artikelnummer", "Text Modul ID", LAST_WRITTEN_HEADER]
CSV_ARTICLE_HEADER_ALIASES = {
    "artikelnummer",
    "artikelnr",
    "artikelnr",
    "artnr",
    "produktnummer",
    "sku",
    "article",
}
CSV_URL_HEADER_ALIASES = {
    "kunzerprodukturl",
    "kunzerurl",
    "produkturl",
    "produktlink",
    "produktseite",
    "produktseitenurl",
    "url",
    "link",
}

EXPORT_LANGUAGE_LAYOUT = [
    ("255", "uni"),
    ("1", "de"),
    ("4", "en"),
    ("18", "cz"),
    ("6", "fr"),
    ("7", "it"),
    ("9", "nl"),
]

UI_LANGUAGE_ORDER = [
    ("de", "Deutsch"),
    ("en", "Englisch"),
    ("cz", "Tschechisch"),
    ("fr", "Franzoesisch"),
    ("it", "Italienisch"),
    ("nl", "Niederlaendisch"),
    ("uni", "UNI"),
]

DEEPL_TARGET_LANGUAGES = {
    "en": "EN-GB",
    "cz": "CS",
    "fr": "FR",
    "it": "IT",
    "nl": "NL",
}

BROWSER_EXECUTABLE_CANDIDATES = [
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
]

KUNZER_GENERIC_YOUTUBE_LINKS = {
    "https://www.youtube.com/channel/UCUpBSumIkYx8-C6QYZe9Z0Q",
}

ATTACHMENT_FORMAT_TYPE_BY_EXTENSION = {
    ".jpg": "3",
    ".jpeg": "3",
    ".png": "6",
    ".pdf": "2",
    ".gif": "7",
}


@dataclass
class TranslationSet:
    de: str = ""
    en: str = ""
    cz: str = ""
    fr: str = ""
    it: str = ""
    nl: str = ""
    uni: str = ""

    def normalized(self) -> "TranslationSet":
        return TranslationSet(
            de=self.de.strip(),
            en=self.en.strip(),
            cz=self.cz.strip(),
            fr=self.fr.strip(),
            it=self.it.strip(),
            nl=self.nl.strip(),
            uni=self.uni.strip(),
        )

    def export_values(self, auto_uni: bool) -> list[str]:
        values = self.normalized()
        if auto_uni:
            values.uni = values.de
        export = []
        for language_id, key in EXPORT_LANGUAGE_LAYOUT:
            export.extend([language_id, getattr(values, key)])
        return export

    def populated_count(self, auto_uni: bool) -> int:
        values = self.normalized()
        if auto_uni:
            values.uni = values.de
        return sum(1 for value in vars(values).values() if value)


@dataclass
class MediaRow:
    path_or_link: str
    art: str = ""
    sprache: str = ""


@dataclass
class ExportBundle:
    article_number: str
    short_module_id: str
    long_module_id: str
    short_texts: TranslationSet
    short_auto_uni: bool
    long_texts: TranslationSet
    long_auto_uni: bool
    image_rows: list[MediaRow] = field(default_factory=list)
    document_rows: list[MediaRow] = field(default_factory=list)
    video_rows: list[MediaRow] = field(default_factory=list)
    web_rows: list[MediaRow] = field(default_factory=list)
    include_short_text: bool = True
    include_long_text: bool = True
    include_images: bool = True
    include_documents: bool = True
    include_videos: bool = True
    include_web_links: bool = True


@dataclass
class StoredArticleSnapshot:
    article_number: str
    source_label: str
    source_folder: Path
    short_module_id: str = ""
    long_module_id: str = ""
    short_texts: TranslationSet = field(default_factory=TranslationSet)
    short_auto_uni: bool = True
    long_texts: TranslationSet = field(default_factory=TranslationSet)
    long_auto_uni: bool = True
    image_rows: list[MediaRow] = field(default_factory=list)
    document_rows: list[MediaRow] = field(default_factory=list)
    video_rows: list[MediaRow] = field(default_factory=list)
    web_rows: list[MediaRow] = field(default_factory=list)


def normalize_article_number(value: str) -> str:
    return " ".join(value.strip().split())


def infer_document_art(document_path: str) -> str:
    parsed = urlparse(document_path.strip())
    raw_name = unquote(parsed.path if parsed.scheme else document_path)
    lowered = Path(raw_name).name.lower()
    normalized = (
        lowered.replace("_", " ")
        .replace("-", " ")
        .replace(".", " ")
        .replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )
    if any(token in normalized for token in ["bedienungsanleitung", "betriebsanleitung", "einbauanleitung", "montageanleitung"]):
        return "14"
    if any(token in normalized for token in ["produktinfo", "zubehoer", "zubehor", "datenblatt", "broschuere", "flyer"]):
        return "17"
    return "17"


def safe_folder_name(article_number: str) -> str:
    cleaned = article_number.strip().replace(" ", "_")
    invalid = '<>:"/\\|?*'
    for char in invalid:
        cleaned = cleaned.replace(char, "_")
    return cleaned or "ohne_artikelnummer"


def normalize_csv_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())


@dataclass
class CsvImportItem:
    article_number: str = ""
    product_url: str = ""


def _choose_csv_delimiter(sample: str) -> str:
    delimiters = [";", ",", "\t", "|"]
    counts = {delimiter: sample.count(delimiter) for delimiter in delimiters}
    best = max(counts, key=counts.get)
    return best if counts[best] else ";"


def _read_import_items_from_rows(header_values: list[object], data_rows: list[list[object]], source_label: str) -> list[CsvImportItem]:
    header_map = {
        normalize_csv_header(str(field_name)): index
        for index, field_name in enumerate(header_values)
        if field_name is not None and str(field_name).strip()
    }
    article_index = next((header_map[key] for key in CSV_ARTICLE_HEADER_ALIASES if key in header_map), None)
    url_index = next((header_map[key] for key in CSV_URL_HEADER_ALIASES if key in header_map), None)
    if article_index is None and url_index is None:
        raise ValueError(f"Die {source_label} braucht mindestens eine Spalte fuer Artikelnummer oder Produkt-URL.")

    items: list[CsvImportItem] = []
    for row in data_rows:
        article_value = ""
        if article_index is not None and article_index < len(row) and row[article_index] is not None:
            article_value = normalize_article_number(str(row[article_index]))

        url_value = ""
        if url_index is not None and url_index < len(row) and row[url_index] is not None:
            url_value = str(row[url_index]).strip()

        if not article_value and not url_value:
            continue
        items.append(CsvImportItem(article_number=article_value, product_url=url_value))

    if not items:
        raise ValueError(f"In der {source_label} wurden keine importierbaren Produkte gefunden.")
    return items


def read_product_import_items(path: Path) -> list[CsvImportItem]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            worksheet = workbook[workbook.sheetnames[0]]
            rows = list(worksheet.iter_rows(values_only=True))
        finally:
            workbook.close()

        if not rows:
            raise ValueError("Die XLSX-Datei ist leer.")
        return _read_import_items_from_rows(list(rows[0]), [list(row) for row in rows[1:]], "XLSX-Datei")

    if suffix != ".csv":
        raise ValueError("Unterstuetzt werden derzeit CSV- und XLSX-Dateien.")

    last_error: UnicodeDecodeError | None = None
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                delimiter = _choose_csv_delimiter(sample)
                reader = csv.DictReader(handle, delimiter=delimiter)
                if not reader.fieldnames:
                    raise ValueError("Die CSV-Datei hat keine Kopfzeile.")
                rows = [[row.get(field_name, "") for field_name in reader.fieldnames] for row in reader]
                return _read_import_items_from_rows(list(reader.fieldnames), rows, "CSV-Datei")
        except UnicodeDecodeError as exc:
            last_error = exc

    if last_error is not None:
        raise ValueError(f"CSV-Datei konnte nicht gelesen werden: {last_error}") from last_error
    raise ValueError("CSV-Datei konnte nicht gelesen werden.")


def normalize_youtube_url_to_public(url: str) -> str:
    value = url.strip()
    parsed = urlparse(value)
    host = parsed.netloc.lower()
    path = parsed.path.strip("/")
    query = parse_qs(parsed.query)
    video_id = ""

    if host in {"youtu.be", "www.youtu.be"}:
        video_id = path.split("/", 1)[0]
    elif host.endswith("youtube.com"):
        parts = [part for part in path.split("/") if part]
        if len(parts) >= 2 and parts[0] in {"embed", "shorts", "live"}:
            video_id = parts[1]
        elif parts[:1] == ["watch"]:
            video_id = (query.get("v") or [""])[0]

    if not video_id:
        return value

    passthrough: dict[str, str] = {}
    for key in ["si", "start", "t"]:
        if query.get(key):
            passthrough[key] = query[key][0]

    if "start" in passthrough and "t" not in passthrough:
        passthrough["t"] = passthrough.pop("start")

    query_string = f"?{urlencode(passthrough)}" if passthrough else ""
    return f"https://youtu.be/{video_id}{query_string}"


def extract_youtube_video_id(url: str) -> str:
    value = url.strip()
    parsed = urlparse(value)
    host = parsed.netloc.lower()
    path = parsed.path.strip("/")
    query = parse_qs(parsed.query)

    if host in {"youtu.be", "www.youtu.be"}:
        return path.split("/", 1)[0]
    if host.endswith("youtube.com"):
        parts = [part for part in path.split("/") if part]
        if len(parts) >= 2 and parts[0] in {"embed", "shorts", "live"}:
            return parts[1]
        if parts[:1] == ["watch"]:
            return (query.get("v") or [""])[0]
    return ""


def detect_browser_path() -> Path | None:
    for candidate in BROWSER_EXECUTABLE_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def infer_attachment_format_type_id(path_or_link: str) -> str:
    value = path_or_link.strip()
    parsed = urlparse(value)
    if parsed.scheme in {"http", "https"}:
        return "4"
    path_part = parsed.path if parsed.scheme else value
    suffix = Path(unquote(path_part)).suffix.lower()
    if suffix in ATTACHMENT_FORMAT_TYPE_BY_EXTENSION:
        return ATTACHMENT_FORMAT_TYPE_BY_EXTENSION[suffix]
    return ""


class IdRegistry:
    def __init__(self) -> None:
        self.used_ids: set[str] = set()
        self.short_ids_by_article: dict[str, str] = {}
        self.long_ids_by_article: dict[str, str] = {}
        self.random = random.SystemRandom()

    def load_from_folder(self, folder: Path, clear_existing: bool = True) -> tuple[int, list[str]]:
        if clear_existing:
            self.used_ids.clear()
            self.short_ids_by_article.clear()
            self.long_ids_by_article.clear()
        warnings: list[str] = []
        candidates = [
            "Kurzbezeichnung-NEU.xlsx",
            "Text-NEU.xlsx",
            "Text_zu_ID.xlsx",
            "Kurzbezeichnung_zu_ID.xlsx",
        ]

        for file_name in candidates:
            workbook_path = folder / file_name
            if not workbook_path.exists():
                continue
            try:
                workbook = load_workbook(workbook_path, read_only=True, data_only=True)
            except Exception as exc:  # pragma: no cover - defensive user feedback
                warnings.append(f"{file_name}: {exc}")
                continue

            worksheet = workbook[workbook.sheetnames[0]]
            rows = worksheet.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                workbook.close()
                continue
            header_map = {str(value).strip(): index for index, value in enumerate(header) if value is not None}
            if "Text Modul ID" not in header_map:
                workbook.close()
                continue
            id_index = header_map["Text Modul ID"]
            article_index = None
            for candidate_name in ["Artikelnummer", "Artikelnummer ", "Produktnummer"]:
                if candidate_name in header_map:
                    article_index = header_map[candidate_name]
                    break
            for row in rows:
                if row is None or id_index >= len(row):
                    continue
                value = row[id_index]
                if value is None:
                    continue
                candidate = str(value).strip()
                if candidate:
                    self.used_ids.add(candidate)
                    if article_index is not None and article_index < len(row) and row[article_index] is not None:
                        article_number = normalize_article_number(str(row[article_index]))
                        if article_number:
                            if file_name in {"Kurzbezeichnung-NEU.xlsx", "Kurzbezeichnung_zu_ID.xlsx"}:
                                self.short_ids_by_article[article_number] = candidate
                            elif file_name in {"Text-NEU.xlsx", "Text_zu_ID.xlsx"}:
                                self.long_ids_by_article[article_number] = candidate
            workbook.close()

        return len(self.used_ids), warnings

    def reserve(self, value: str) -> None:
        if value:
            self.used_ids.add(value)

    def remember_article_ids(self, article_number: str, short_id: str, long_id: str) -> None:
        article_key = normalize_article_number(article_number)
        if not article_key:
            return
        if short_id:
            self.short_ids_by_article[article_key] = short_id
            self.used_ids.add(short_id)
        if long_id:
            self.long_ids_by_article[article_key] = long_id
            self.used_ids.add(long_id)

    def get_ids_for_article(self, article_number: str) -> tuple[str, str]:
        article_key = normalize_article_number(article_number)
        return (
            self.short_ids_by_article.get(article_key, ""),
            self.long_ids_by_article.get(article_key, ""),
        )

    def generate_unique(self) -> str:
        while True:
            candidate = "".join(self.random.choice(ID_ALPHABET) for _ in range(ID_LENGTH))
            if candidate not in self.used_ids:
                self.used_ids.add(candidate)
                return candidate


class DeepLTranslationError(RuntimeError):
    pass


class DeepLClient:
    def __init__(self, auth_key: str, base_url: str, timeout_seconds: int = 45) -> None:
        self.auth_key = auth_key.strip()
        self.base_url = base_url.rstrip("/")
        self.timeout_seconds = timeout_seconds

    def translate_texts(self, texts: list[str], target_lang: str, source_lang: str = "DE") -> list[str]:
        payload = {
            "text": texts,
            "source_lang": source_lang,
            "target_lang": target_lang,
            "preserve_formatting": True,
        }
        request = urllib_request.Request(
            url=f"{self.base_url}/v2/translate",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"DeepL-Auth-Key {self.auth_key}",
                "Content-Type": "application/json",
                "User-Agent": "ApolloImportGui/1.0",
            },
            method="POST",
        )
        try:
            with urllib_request.urlopen(request, timeout=self.timeout_seconds) as response:
                body = response.read().decode("utf-8")
        except urllib_error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            raise DeepLTranslationError(f"HTTP {exc.code}: {details or exc.reason}") from exc
        except urllib_error.URLError as exc:
            raise DeepLTranslationError(f"Verbindungsfehler: {exc.reason}") from exc

        try:
            parsed = json.loads(body)
            return [item["text"] for item in parsed["translations"]]
        except (KeyError, TypeError, json.JSONDecodeError) as exc:
            raise DeepLTranslationError(f"Unerwartete Antwort von DeepL: {body[:400]}") from exc

    def translate_from_german(self, german_text: str) -> dict[str, str]:
        text = german_text.strip()
        if not text:
            raise DeepLTranslationError("Kein deutscher Quelltext zum Uebersetzen vorhanden.")

        translations: dict[str, str] = {}
        for ui_code, deepl_code in DEEPL_TARGET_LANGUAGES.items():
            translations[ui_code] = self.translate_texts([text], target_lang=deepl_code, source_lang="DE")[0]
        return translations


class KunzerScrapeError(RuntimeError):
    pass


@dataclass
class KunzerScrapeResult:
    article_number: str
    product_url: str
    title: str
    short_text_de: str
    long_text_de: str
    image_links: list[str]
    document_links: list[str]
    video_links: list[str]


class KunzerScraper:
    def __init__(self) -> None:
        self._sitemap_index: dict[str, str] | None = None

    def resolve_product_url(self, article_number_or_url: str) -> str:
        value = article_number_or_url.strip()
        if not value:
            raise KunzerScrapeError("Bitte eine Artikelnummer oder Kunzer Produkt-URL eingeben.")
        if value.lower().startswith(("http://", "https://")):
            return value

        sitemap = self._load_sitemap_index()
        normalized = normalize_article_number(value)
        if normalized in sitemap:
            return sitemap[normalized]

        encoded = quote(normalized, safe=".-")
        return f"https://www.kunzer.de/shop/p/{encoded}"

    def scrape_product(self, article_number_or_url: str) -> KunzerScrapeResult:
        target_url = self.resolve_product_url(article_number_or_url)
        browser_path = self._detect_browser_path()

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(browser_path) if browser_path else None,
                headless=True,
            )
            page = browser.new_page(viewport={"width": 1440, "height": 2400})
            try:
                page.goto(target_url, wait_until="domcontentloaded", timeout=120000)
                page.wait_for_timeout(1500)
                self._accept_cookies_if_present(page)
                page.wait_for_load_state("networkidle", timeout=120000)
                page.wait_for_timeout(800)
                self._expand_all_documents(page)

                title = self._clean_single_line(page.title().replace("| Kunzer", ""))
                heading = self._safe_first_text(page.get_by_role("heading").all_inner_texts())
                short_text_de = self._clean_single_line(heading or title)

                description_text = self._extract_tab_content(page, "Artikelbeschreibung")
                info_text = self._extract_tab_content(page, "Artikelinformationen")
                long_text_de = self._select_long_text(short_text_de, description_text, info_text)

                links = self._extract_links(page)
                image_links = [href for href in links if self._is_image_link(href)]
                document_links = [href for href in links if href.lower().endswith(".pdf")]
                video_links = [href for href in links if self._is_video_link(href)]

                article_number = normalize_article_number(
                    self._extract_article_number(page.url, page.locator("body").inner_text(), short_text_de)
                )

                return KunzerScrapeResult(
                    article_number=article_number,
                    product_url=page.url,
                    title=short_text_de,
                    short_text_de=short_text_de,
                    long_text_de=long_text_de,
                    image_links=image_links,
                    document_links=document_links,
                    video_links=video_links,
                )
            except PlaywrightTimeoutError as exc:
                raise KunzerScrapeError(f"Timeout beim Laden der Kunzer-Seite: {target_url}") from exc
            except Exception as exc:  # pragma: no cover - runtime GUI feedback
                raise KunzerScrapeError(f"Kunzer-Seite konnte nicht gelesen werden: {exc}") from exc
            finally:
                browser.close()

    def _load_sitemap_index(self) -> dict[str, str]:
        if self._sitemap_index is not None:
            return self._sitemap_index

        request = urllib_request.Request(
            url="https://www.kunzer.de/products-sitemap.xml",
            headers={"User-Agent": "ApolloImportGui/1.0"},
            method="GET",
        )
        try:
            with urllib_request.urlopen(request, timeout=60) as response:
                xml_text = response.read().decode("utf-8")
        except urllib_error.URLError as exc:
            raise KunzerScrapeError(f"Kunzer Sitemap konnte nicht geladen werden: {exc.reason}") from exc

        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError as exc:
            raise KunzerScrapeError("Kunzer Sitemap ist nicht lesbar.") from exc

        namespace = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        entries: dict[str, str] = {}
        for loc in root.findall(".//sm:loc", namespace):
            if loc.text:
                url = loc.text.strip()
                segment = unquote(url.rstrip("/").rsplit("/", 1)[-1])
                entries[normalize_article_number(segment)] = url

        self._sitemap_index = entries
        return entries

    def _detect_browser_path(self) -> Path | None:
        return detect_browser_path()

    def _accept_cookies_if_present(self, page: object) -> None:
        try:
            page.get_by_role("button", name="Alle zulassen").click(timeout=4000)
            page.wait_for_timeout(800)
        except Exception:
            return

    def _expand_all_documents(self, page: object) -> None:
        try:
            page.get_by_role("button", name="Alle Dokumente anzeigen").click(timeout=4000)
            page.wait_for_timeout(800)
        except Exception:
            return

    def _extract_links(self, page: object) -> list[str]:
        hrefs = []
        for href in page.locator("a").evaluate_all("(nodes) => nodes.map((node) => node.href || '')"):
            cleaned = href.strip()
            if cleaned and cleaned not in hrefs:
                hrefs.append(cleaned)
        return hrefs

    def _extract_tab_content(self, page: object, button_label: str) -> str:
        try:
            button = page.get_by_role("button", name=button_label)
            button.click(timeout=4000)
            page.wait_for_timeout(400)
        except Exception:
            return ""

        try:
            controls_id = button.get_attribute("aria-controls")
            if controls_id:
                panel = page.locator(f"#{controls_id}")
                if panel.count() > 0:
                    content = panel.inner_text().strip()
                    if content:
                        return self._normalize_multiline_text(content)
        except Exception:
            pass

        body_text = page.locator("body").inner_text()
        sections = re.split(r"\n(?=Artikelbeschreibung|Artikelinformationen)\n", body_text)
        for section in sections:
            if section.startswith(button_label):
                lines = [line.strip() for line in section.splitlines()[1:]]
                collected: list[str] = []
                for line in lines:
                    if not line or line in {"Artikelbeschreibung", "Artikelinformationen", "Anfrage senden", "Verfügbar"}:
                        break
                    collected.append(line)
                if collected:
                    return self._normalize_multiline_text("\n".join(collected))
        return ""

    def _select_long_text(self, title: str, description_text: str, info_text: str) -> str:
        description = self._normalize_multiline_text(description_text)
        if description and description != title:
            return description

        info = self._normalize_multiline_text(info_text)
        if info and info != title:
            return info

        return title

    def _extract_article_number(self, page_url: str, body_text: str, title: str) -> str:
        url_segment = unquote(urlparse(page_url).path.rstrip("/").rsplit("/", 1)[-1])
        if url_segment and url_segment.lower() not in {"p", "shop"}:
            return url_segment

        candidates = body_text.splitlines()
        for line in candidates:
            cleaned = self._clean_single_line(line)
            if cleaned and cleaned != title and len(cleaned) <= 40 and re.fullmatch(r"[A-Z0-9 .\\-]+", cleaned):
                return cleaned
        return title

    def _is_image_link(self, href: str) -> bool:
        lower = href.lower()
        return lower.endswith((".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"))

    def _is_video_link(self, href: str) -> bool:
        lower = href.lower()
        if href in KUNZER_GENERIC_YOUTUBE_LINKS:
            return False
        return "youtube.com" in lower or "youtu.be" in lower

    def _safe_first_text(self, values: list[str]) -> str:
        for value in values:
            cleaned = self._clean_single_line(value)
            if cleaned:
                return cleaned
        return ""

    def _clean_single_line(self, value: str) -> str:
        return " ".join(value.replace("\r", " ").replace("\n", " ").split()).strip()

    def _normalize_multiline_text(self, value: str) -> str:
        lines = [line.strip() for line in value.replace("\r", "\n").split("\n")]
        compacted: list[str] = []
        blank_pending = False
        for line in lines:
            if line:
                if blank_pending and compacted:
                    compacted.append("")
                compacted.append(line)
                blank_pending = False
            else:
                blank_pending = True
        return "\n".join(compacted).strip()


def write_workbook(path: Path, sheet_name: str, headers: list[str], rows: list[list[str]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def read_workbook_rows(path: Path, sheet_name: str, expected_width: int) -> list[list[str]]:
    if not path.exists():
        return []

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook.active
        existing_rows: list[list[str]] = []
        row_iter = worksheet.iter_rows(values_only=True)
        next(row_iter, None)
        for values in row_iter:
            row = ["" if value is None else str(value) for value in values[:expected_width]]
            if len(row) < expected_width:
                row.extend([""] * (expected_width - len(row)))
            if any(cell.strip() for cell in row):
                existing_rows.append(row)
        return existing_rows
    finally:
        workbook.close()


def prepare_existing_rows_for_write(path: Path, sheet_name: str, headers: list[str]) -> list[list[str]]:
    existing_rows = read_workbook_rows(path, sheet_name, len(headers))
    if headers and headers[-1] == LAST_WRITTEN_HEADER and path.exists():
        fallback_written_at = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        for row in existing_rows:
            if row and not row[-1].strip() and any(cell.strip() for cell in row[:-1]):
                row[-1] = fallback_written_at
    return existing_rows


def merge_rows_by_article(
    existing_rows: list[list[str]],
    new_rows: list[list[str]],
    key_index: int = 0,
    replace_article_keys: set[str] | None = None,
) -> list[list[str]]:
    article_keys = set(replace_article_keys or set())
    article_keys.update(
        normalize_article_number(str(row[key_index]))
        for row in new_rows
        if len(row) > key_index and normalize_article_number(str(row[key_index]))
    )
    if not article_keys:
        return existing_rows

    merged_rows = [
        row
        for row in existing_rows
        if len(row) <= key_index or normalize_article_number(str(row[key_index])) not in article_keys
    ]
    merged_rows.extend(new_rows)
    return merged_rows


def write_workbook_with_upsert(
    path: Path,
    sheet_name: str,
    headers: list[str],
    rows: list[list[str]],
    key_index: int = 0,
    replace_article_keys: set[str] | None = None,
) -> None:
    existing_rows = prepare_existing_rows_for_write(path, sheet_name, headers)
    merged_rows = merge_rows_by_article(existing_rows, rows, key_index=key_index, replace_article_keys=replace_article_keys)
    write_workbook(path, sheet_name, headers, merged_rows)


def remove_article_rows_from_workbook(
    path: Path,
    sheet_name: str,
    headers: list[str],
    article_numbers: set[str],
    key_index: int = 0,
) -> None:
    if not path.exists():
        return

    normalized_article_numbers = {normalize_article_number(article_number) for article_number in article_numbers if normalize_article_number(article_number)}
    if not normalized_article_numbers:
        return

    existing_rows = prepare_existing_rows_for_write(path, sheet_name, headers)
    remaining_rows = [
        row
        for row in existing_rows
        if len(row) <= key_index or normalize_article_number(str(row[key_index])) not in normalized_article_numbers
    ]
    write_workbook(path, sheet_name, headers, remaining_rows)


def replace_article_rows_preserving_timestamps(
    path: Path,
    sheet_name: str,
    headers: list[str],
    article_number: str,
    rows: list[list[str]],
    key_index: int = 0,
) -> None:
    existing_rows = prepare_existing_rows_for_write(path, sheet_name, headers)
    article_key = normalize_article_number(article_number)
    has_last_written = bool(headers) and headers[-1] == LAST_WRITTEN_HEADER
    written_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rebuilt_rows = rows
    if has_last_written:
        existing_timestamps_by_signature: dict[tuple[str, ...], list[str]] = {}
        signature_width = len(headers) - 1
        for row in existing_rows:
            if len(row) <= key_index or normalize_article_number(str(row[key_index])) != article_key:
                continue
            signature = tuple(str(cell) for cell in row[:signature_width])
            existing_timestamps_by_signature.setdefault(signature, []).append(row[-1].strip() or written_at)

        rebuilt_rows = []
        for row in rows:
            signature_values = [str(cell) for cell in row[:signature_width]]
            if len(signature_values) < signature_width:
                signature_values.extend([""] * (signature_width - len(signature_values)))
            signature = tuple(signature_values)
            known_timestamps = existing_timestamps_by_signature.get(signature, [])
            timestamp = known_timestamps.pop(0) if known_timestamps else written_at
            rebuilt_rows.append(signature_values + [timestamp])

    merged_rows: list[list[str]] = []
    inserted = False
    for existing_row in existing_rows:
        if len(existing_row) > key_index and normalize_article_number(str(existing_row[key_index])) == article_key:
            if not inserted:
                merged_rows.extend(rebuilt_rows)
                inserted = True
            continue
        merged_rows.append(existing_row)

    if not inserted:
        merged_rows.extend(rebuilt_rows)

    write_workbook(path, sheet_name, headers, merged_rows)


def build_short_text_export_row(bundle: ExportBundle) -> list[str]:
    return [
        bundle.article_number,
        bundle.short_module_id,
        "1",
        "1",
        *bundle.short_texts.export_values(bundle.short_auto_uni),
    ]


def build_long_text_export_row(bundle: ExportBundle) -> list[str]:
    return [
        bundle.article_number,
        bundle.long_module_id,
        "2",
        "1",
        *bundle.long_texts.export_values(bundle.long_auto_uni),
    ]


def build_short_mapping_export_row(bundle: ExportBundle) -> list[str]:
    return [bundle.article_number, bundle.short_module_id]


def build_image_export_rows(bundle: ExportBundle) -> list[list[str]]:
    export_rows: list[list[str]] = []
    for row in bundle.image_rows:
        format_type_id = infer_attachment_format_type_id(row.path_or_link)
        if not format_type_id:
            raise ValueError(f"Kein TecDoc Anhangsformattyp fuer Bild ableitbar: {row.path_or_link}")
        export_rows.append([bundle.article_number, row.path_or_link, row.art or "5", row.sprache or "255", format_type_id])
    return export_rows


def build_document_export_rows(bundle: ExportBundle) -> list[list[str]]:
    export_rows: list[list[str]] = []
    for row in bundle.document_rows:
        format_type_id = infer_attachment_format_type_id(row.path_or_link)
        if not format_type_id:
            raise ValueError(f"Kein TecDoc Anhangsformattyp fuer Dokument ableitbar: {row.path_or_link}")
        export_rows.append([bundle.article_number, row.path_or_link, row.sprache or "255", row.art or "17", format_type_id])
    return export_rows


def build_video_export_rows(bundle: ExportBundle) -> list[list[str]]:
    export_rows: list[list[str]] = []
    for row in bundle.video_rows:
        normalized_video_link = normalize_youtube_url_to_public(row.path_or_link)
        format_type_id = infer_attachment_format_type_id(normalized_video_link)
        if not format_type_id:
            raise ValueError(f"Kein TecDoc Anhangsformattyp fuer Video-Link ableitbar: {row.path_or_link}")
        export_rows.append([bundle.article_number, normalized_video_link, format_type_id])
    return export_rows


def build_web_export_rows(bundle: ExportBundle) -> list[list[str]]:
    export_rows: list[list[str]] = []
    for row in bundle.web_rows:
        format_type_id = infer_attachment_format_type_id(row.path_or_link)
        if not format_type_id:
            raise ValueError(f"Kein TecDoc Anhangsformattyp fuer Web-Link ableitbar: {row.path_or_link}")
        export_rows.append([bundle.article_number, row.path_or_link, format_type_id])
    return export_rows


def append_written_at(row: list[str], written_at: str) -> list[str]:
    return [*row, written_at]


def export_bundle(bundle: ExportBundle, output_root: Path, use_timestamp_subdir: bool) -> Path:
    written_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if use_timestamp_subdir:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_dir = output_root / f"{safe_folder_name(bundle.article_number)}_{timestamp}"
    else:
        export_dir = output_root
    export_dir.mkdir(parents=True, exist_ok=True)

    short_row = append_written_at(build_short_text_export_row(bundle), written_at)
    long_row = append_written_at(build_long_text_export_row(bundle), written_at)
    short_mapping_row = append_written_at(build_short_mapping_export_row(bundle), written_at)
    image_rows = [append_written_at(row, written_at) for row in build_image_export_rows(bundle)] if bundle.include_images else []
    document_rows = [append_written_at(row, written_at) for row in build_document_export_rows(bundle)] if bundle.include_documents else []
    video_rows = [append_written_at(row, written_at) for row in build_video_export_rows(bundle)] if bundle.include_videos else []
    web_rows = [append_written_at(row, written_at) for row in build_web_export_rows(bundle)] if bundle.include_web_links else []

    if use_timestamp_subdir:
        if bundle.include_short_text:
            write_workbook(export_dir / SHORT_TEXT_FILE[0], SHORT_TEXT_FILE[1], SHORT_TEXT_HEADERS, [short_row])
            write_workbook(export_dir / SHORT_MAPPING_FILE[0], SHORT_MAPPING_FILE[1], SHORT_MAPPING_HEADERS, [short_mapping_row])
        if bundle.include_long_text:
            write_workbook(export_dir / LONG_TEXT_FILE[0], LONG_TEXT_FILE[1], SHORT_TEXT_HEADERS, [long_row])
        if bundle.include_images:
            write_workbook(export_dir / IMAGE_FILE[0], IMAGE_FILE[1], IMAGE_HEADERS, image_rows)
        if bundle.include_documents:
            write_workbook(export_dir / DOCUMENT_FILE[0], DOCUMENT_FILE[1], DOCUMENT_HEADERS, document_rows)
        if bundle.include_videos:
            write_workbook(export_dir / VIDEO_FILE[0], VIDEO_FILE[1], VIDEO_HEADERS, video_rows)
        if bundle.include_web_links:
            write_workbook(export_dir / WEB_LINK_FILE[0], WEB_LINK_FILE[1], WEB_HEADERS, web_rows)
    else:
        article_keys = {bundle.article_number}
        if bundle.include_short_text:
            write_workbook_with_upsert(
                export_dir / SHORT_TEXT_FILE[0],
                SHORT_TEXT_FILE[1],
                SHORT_TEXT_HEADERS,
                [short_row],
                replace_article_keys=article_keys,
            )
            write_workbook_with_upsert(
                export_dir / SHORT_MAPPING_FILE[0],
                SHORT_MAPPING_FILE[1],
                SHORT_MAPPING_HEADERS,
                [short_mapping_row],
                replace_article_keys=article_keys,
            )
        if bundle.include_long_text:
            write_workbook_with_upsert(
                export_dir / LONG_TEXT_FILE[0],
                LONG_TEXT_FILE[1],
                SHORT_TEXT_HEADERS,
                [long_row],
                replace_article_keys=article_keys,
            )
        if bundle.include_images:
            write_workbook_with_upsert(
                export_dir / IMAGE_FILE[0],
                IMAGE_FILE[1],
                IMAGE_HEADERS,
                image_rows,
                replace_article_keys=article_keys,
            )
        if bundle.include_documents:
            write_workbook_with_upsert(
                export_dir / DOCUMENT_FILE[0],
                DOCUMENT_FILE[1],
                DOCUMENT_HEADERS,
                document_rows,
                replace_article_keys=article_keys,
            )
        if bundle.include_videos:
            write_workbook_with_upsert(
                export_dir / VIDEO_FILE[0],
                VIDEO_FILE[1],
                VIDEO_HEADERS,
                video_rows,
                replace_article_keys=article_keys,
            )
        if bundle.include_web_links:
            write_workbook_with_upsert(
                export_dir / WEB_LINK_FILE[0],
                WEB_LINK_FILE[1],
                WEB_HEADERS,
                web_rows,
                replace_article_keys=article_keys,
            )

    return export_dir


def build_preview(bundle: ExportBundle) -> str:
    lines = [
        f"Artikelnummer: {bundle.article_number or '(leer)'}",
        "",
        "IDs",
        f"- Kurzbezeichnung: {bundle.short_module_id or '(noch nicht generiert)'}",
        f"- Text: {bundle.long_module_id or '(noch nicht generiert)'}",
        "",
        "Texte",
        f"- Kurzbezeichnung: {bundle.short_texts.populated_count(bundle.short_auto_uni)} / 7 Sprachfelder befuellt",
        f"- Text: {bundle.long_texts.populated_count(bundle.long_auto_uni)} / 7 Sprachfelder befuellt",
        "",
        "Medien",
        f"- Bilder: {len(bundle.image_rows)}",
        f"- Dokumente: {len(bundle.document_rows)}",
        f"- Videos: {len(bundle.video_rows)}",
        f"- Web Links: {len(bundle.web_rows)}",
        "",
        "Exportdateien",
        f"- {SHORT_TEXT_FILE[0]}",
        f"- {SHORT_MAPPING_FILE[0]}",
        f"- {LONG_TEXT_FILE[0]}",
        f"- {IMAGE_FILE[0]}",
        f"- {DOCUMENT_FILE[0]}",
        f"- {VIDEO_FILE[0]}",
        f"- {WEB_LINK_FILE[0]}",
    ]
    return "\n".join(lines)


def format_translation_set(label: str, translations: TranslationSet) -> str:
    values = translations.normalized()
    lines = [label]
    for code, language_label in UI_LANGUAGE_ORDER:
        text = getattr(values, code) or "-"
        lines.append(f"{language_label}:")
        lines.append(text)
        lines.append("")
    return "\n".join(lines).rstrip()


def format_media_rows(label: str, rows: list[MediaRow], include_meta: bool = True) -> str:
    lines = [label]
    if not rows:
        lines.append("-")
        return "\n".join(lines)

    for index, row in enumerate(rows, start=1):
        lines.append(f"{index}. {row.path_or_link}")
        if include_meta:
            lines.append(f"   Art: {row.art or '-'} | Sprache: {row.sprache or '-'}")
    return "\n".join(lines)


def format_article_snapshot(snapshot: StoredArticleSnapshot) -> str:
    sections = [
        f"Artikelnummer: {snapshot.article_number}",
        f"Quelle: {snapshot.source_label}",
        f"Ordner: {snapshot.source_folder}",
        "",
        f"Kurz-ID: {snapshot.short_module_id or '-'}",
        f"Text-ID: {snapshot.long_module_id or '-'}",
        "",
        format_translation_set("Kurzbezeichnung", snapshot.short_texts),
        "",
        format_translation_set("Text", snapshot.long_texts),
        "",
        format_media_rows("Bilder", snapshot.image_rows),
        "",
        format_media_rows("Dokumente", snapshot.document_rows),
        "",
        format_media_rows("Videos", snapshot.video_rows, include_meta=False),
        "",
        format_media_rows("Web Links", snapshot.web_rows, include_meta=False),
    ]
    return "\n".join(sections).strip()


def translation_set_from_export_row(row: list[str]) -> TranslationSet:
    padded = list(row) + [""] * max(0, len(SHORT_TEXT_HEADERS) - len(row))
    return TranslationSet(
        uni=padded[5].strip(),
        de=padded[7].strip(),
        en=padded[9].strip(),
        cz=padded[11].strip(),
        fr=padded[13].strip(),
        it=padded[15].strip(),
        nl=padded[17].strip(),
    )


def load_article_snapshots_from_folder(folder: Path, source_label: str) -> dict[str, StoredArticleSnapshot]:
    if not folder.exists():
        return {}

    snapshots: dict[str, StoredArticleSnapshot] = {}

    def ensure_snapshot(article_number: str) -> StoredArticleSnapshot:
        article_key = normalize_article_number(article_number)
        snapshot = snapshots.get(article_key)
        if snapshot is None:
            snapshot = StoredArticleSnapshot(article_number=article_key, source_label=source_label, source_folder=folder)
            snapshots[article_key] = snapshot
        return snapshot

    short_rows = read_workbook_rows(folder / SHORT_TEXT_FILE[0], SHORT_TEXT_FILE[1], len(SHORT_TEXT_HEADERS))
    for row in short_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        snapshot = ensure_snapshot(article_number)
        snapshot.short_module_id = row[1].strip()
        snapshot.short_texts = translation_set_from_export_row(row)
        snapshot.short_auto_uni = not snapshot.short_texts.uni or snapshot.short_texts.uni == snapshot.short_texts.de

    long_rows = read_workbook_rows(folder / LONG_TEXT_FILE[0], LONG_TEXT_FILE[1], len(SHORT_TEXT_HEADERS))
    for row in long_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        snapshot = ensure_snapshot(article_number)
        snapshot.long_module_id = row[1].strip()
        snapshot.long_texts = translation_set_from_export_row(row)
        snapshot.long_auto_uni = not snapshot.long_texts.uni or snapshot.long_texts.uni == snapshot.long_texts.de

    image_rows = read_workbook_rows(folder / IMAGE_FILE[0], IMAGE_FILE[1], len(IMAGE_HEADERS))
    for row in image_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        ensure_snapshot(article_number).image_rows.append(MediaRow(path_or_link=row[1].strip(), art=row[2].strip(), sprache=row[3].strip()))

    document_rows = read_workbook_rows(folder / DOCUMENT_FILE[0], DOCUMENT_FILE[1], len(DOCUMENT_HEADERS))
    for row in document_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        ensure_snapshot(article_number).document_rows.append(MediaRow(path_or_link=row[1].strip(), art=row[3].strip(), sprache=row[2].strip()))

    video_rows = read_workbook_rows(folder / VIDEO_FILE[0], VIDEO_FILE[1], len(VIDEO_HEADERS))
    for row in video_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        ensure_snapshot(article_number).video_rows.append(MediaRow(path_or_link=row[1].strip()))

    web_rows = read_workbook_rows(folder / WEB_LINK_FILE[0], WEB_LINK_FILE[1], len(WEB_HEADERS))
    for row in web_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        ensure_snapshot(article_number).web_rows.append(MediaRow(path_or_link=row[1].strip()))

    return snapshots


class SingleLineTranslationFrame(ttk.LabelFrame):
    def __init__(self, master: tk.Misc, title: str, on_change: Callable[[], None] | None = None) -> None:
        super().__init__(master, text=title, padding=14)
        self.on_change = on_change
        self.auto_uni_var = tk.BooleanVar(value=True)
        self.fields: dict[str, tk.StringVar] = {code: tk.StringVar() for code, _ in UI_LANGUAGE_ORDER}
        self.uni_entry: ttk.Entry | None = None
        self.columnconfigure(1, weight=1)

        auto_uni_check = ttk.Checkbutton(
            self,
            text="UNI automatisch aus Deutsch uebernehmen",
            variable=self.auto_uni_var,
            command=self._handle_auto_uni_toggle,
        )
        auto_uni_check.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        for row_index, (code, label) in enumerate(UI_LANGUAGE_ORDER, start=1):
            ttk.Label(self, text=label).grid(row=row_index, column=0, sticky="w", padx=(0, 10), pady=4)
            entry_state = "readonly" if code == "uni" else "normal"
            entry = ttk.Entry(self, textvariable=self.fields[code], width=90, state=entry_state)
            entry.grid(row=row_index, column=1, sticky="ew", pady=4)
            entry.bind("<FocusOut>", self._handle_focus_out)
            if code == "uni":
                self.uni_entry = entry
            if code == "de":
                self.fields["de"].trace_add("write", self._sync_uni)

        self._toggle_uni_state()

    def _sync_uni(self, *_args: object) -> None:
        if self.auto_uni_var.get():
            self.fields["uni"].set(self.fields["de"].get())

    def _toggle_uni_state(self) -> None:
        if self.auto_uni_var.get():
            self.fields["uni"].set(self.fields["de"].get())
        state = "readonly" if self.auto_uni_var.get() else "normal"
        if self.uni_entry is not None:
            self.uni_entry.configure(state=state)

    def _handle_auto_uni_toggle(self) -> None:
        self._toggle_uni_state()
        self._emit_change()

    def _handle_focus_out(self, _event: tk.Event[tk.Misc]) -> None:
        self._emit_change()

    def _emit_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    def get_value(self) -> TranslationSet:
        return TranslationSet(**{code: variable.get() for code, variable in self.fields.items()})

    def get_german_text(self) -> str:
        return self.fields["de"].get().strip()

    def apply_translations(self, translations: dict[str, str]) -> None:
        for code, value in translations.items():
            if code in self.fields and code != "de":
                self.fields[code].set(value)
        self._toggle_uni_state()

    def set_value(self, value: TranslationSet, auto_uni: bool | None = None) -> None:
        if auto_uni is not None:
            self.auto_uni_var.set(auto_uni)
        for code, variable in self.fields.items():
            variable.set(getattr(value, code))
        self._toggle_uni_state()


class MultiLineTranslationFrame(ttk.LabelFrame):
    def __init__(self, master: tk.Misc, title: str, on_change: Callable[[], None] | None = None) -> None:
        super().__init__(master, text=title, padding=14)
        self.on_change = on_change
        self.auto_uni_var = tk.BooleanVar(value=True)
        self.text_widgets: dict[str, ScrolledText] = {}
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        controls = ttk.Frame(self)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        controls.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            controls,
            text="UNI automatisch aus Deutsch uebernehmen",
            variable=self.auto_uni_var,
            command=self._handle_auto_uni_toggle,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            controls,
            text="Jede Sprache hat ein eigenes grosses Textfeld.",
            foreground="#5E6472",
        ).grid(row=0, column=1, sticky="e")

        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew")

        for code, label in UI_LANGUAGE_ORDER:
            tab = ttk.Frame(notebook, padding=12)
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)
            widget = ScrolledText(tab, wrap="word", height=18, font=("Segoe UI", 10))
            widget.grid(row=0, column=0, sticky="nsew")
            if code == "uni":
                widget.configure(state="disabled")
            if code == "de":
                widget.bind("<<Modified>>", self._on_de_modified)
            widget.bind("<FocusOut>", self._handle_focus_out)
            self.text_widgets[code] = widget
            notebook.add(tab, text=label)

    def _set_text(self, code: str, value: str) -> None:
        widget = self.text_widgets[code]
        previous_state = str(widget.cget("state"))
        if previous_state == "disabled":
            widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)
        if previous_state == "disabled":
            widget.configure(state="disabled")

    def _on_de_modified(self, event: tk.Event[tk.Misc]) -> None:
        widget = event.widget
        if widget.edit_modified():
            if self.auto_uni_var.get():
                self._set_text("uni", self.get_text("de"))
            widget.edit_modified(False)

    def _toggle_uni_state(self) -> None:
        uni_widget = self.text_widgets["uni"]
        if self.auto_uni_var.get():
            self._set_text("uni", self.get_text("de"))
            uni_widget.configure(state="disabled")
        else:
            uni_widget.configure(state="normal")

    def _handle_auto_uni_toggle(self) -> None:
        self._toggle_uni_state()
        self._emit_change()

    def _handle_focus_out(self, _event: tk.Event[tk.Misc]) -> None:
        self._emit_change()

    def _emit_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    def get_text(self, code: str) -> str:
        return self.text_widgets[code].get("1.0", "end").strip()

    def get_german_text(self) -> str:
        return self.get_text("de")

    def apply_translations(self, translations: dict[str, str]) -> None:
        for code, value in translations.items():
            if code in self.text_widgets and code != "de":
                self._set_text(code, value)
        self._toggle_uni_state()

    def get_value(self) -> TranslationSet:
        return TranslationSet(**{code: self.get_text(code) for code, _ in UI_LANGUAGE_ORDER})

    def set_value(self, value: TranslationSet, auto_uni: bool | None = None) -> None:
        if auto_uni is not None:
            self.auto_uni_var.set(auto_uni)
        for code, _label in UI_LANGUAGE_ORDER:
            self._set_text(code, getattr(value, code))
        self._toggle_uni_state()


class MediaTableFrame(ttk.LabelFrame):
    def __init__(
        self,
        master: tk.Misc,
        title: str,
        path_label: str,
        path_dialog_title: str,
        default_art: str,
        default_sprache: str,
        browse_filetypes: list[tuple[str, str]],
        infer_art: bool = False,
        preview_kind: str = "document",
        on_change: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master, text=title, padding=14)
        self.on_change = on_change
        self.path_dialog_title = path_dialog_title
        self.default_art = default_art
        self.default_sprache = default_sprache
        self.browse_filetypes = browse_filetypes
        self.infer_art = infer_art
        self.preview_kind = preview_kind
        self.preview_photo: object | None = None
        self.inline_editor: ttk.Entry | None = None
        self.inline_editor_item = ""
        self.inline_editor_column = ""
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Zeilen kopieren", command=self.copy_selected_rows)
        self.context_menu.add_command(label="Zeilen loeschen", command=self.remove_selected)

        self.columnconfigure(0, weight=1)
        self.columnconfigure(2, weight=0)
        self.rowconfigure(1, weight=1)

        form = ttk.Frame(self)
        form.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        form.columnconfigure(1, weight=1)

        self.path_var = tk.StringVar()
        self.art_var = tk.StringVar(value=default_art)
        self.sprache_var = tk.StringVar(value=default_sprache)

        ttk.Label(form, text=path_label).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.path_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(form, text="Dateien waehlen", command=self.browse_files).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(form, text="Art").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.art_var, width=12).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(form, text="Sprache").grid(row=1, column=1, sticky="e", padx=(0, 80), pady=4)
        ttk.Entry(form, textvariable=self.sprache_var, width=12).grid(row=1, column=2, sticky="w", pady=4)

        actions = ttk.Frame(form)
        actions.grid(row=2, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(actions, text="Zeile hinzufuegen", command=self.add_manual_row).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Auswahl aktualisieren", command=self.update_selected_row).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="Auswahl entfernen", command=self.remove_selected).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Formular leeren", command=self.clear_form).grid(row=0, column=3)

        self.tree = ttk.Treeview(self, columns=("path", "art", "sprache"), show="headings", height=10, selectmode="extended")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self.tree.heading("path", text=path_label)
        self.tree.heading("art", text="Art")
        self.tree.heading("sprache", text="Sprache")
        self.tree.column("path", width=840, anchor="w")
        self.tree.column("art", width=100, anchor="center")
        self.tree.column("sprache", width=100, anchor="center")

        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self._handle_selection)
        self.tree.bind("<Double-1>", self._begin_inline_edit)
        self.tree.bind("<Button-3>", self._open_context_menu)

        preview_frame = ttk.LabelFrame(self, text="Vorschau", padding=10)
        preview_frame.grid(row=1, column=2, sticky="nsew", padx=(12, 0))
        preview_frame.columnconfigure(0, weight=1)

        self.preview_title_var = tk.StringVar(value="Keine Auswahl")
        self.preview_meta_var = tk.StringVar(value="Waehle eine Zeile aus, um mehr Details zu sehen.")
        self.preview_badge_var = tk.StringVar(value="DATEI")

        ttk.Label(preview_frame, textvariable=self.preview_title_var, font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w")

        self.preview_visual = ttk.Label(preview_frame, text="", anchor="center", justify="center")
        self.preview_visual.grid(row=1, column=0, sticky="ew", pady=(10, 10))

        ttk.Label(preview_frame, textvariable=self.preview_meta_var, foreground="#5E6472", wraplength=250, justify="left").grid(
            row=2, column=0, sticky="w"
        )
        self._reset_preview()

    def browse_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title=self.path_dialog_title,
            filetypes=self.browse_filetypes,
        )
        if not selected:
            return
        for path in selected:
            art = infer_document_art(path) if self.infer_art else self.art_var.get().strip() or self.default_art
            sprache = self.sprache_var.get().strip() or self.default_sprache
            self.tree.insert("", "end", values=(path, art, sprache))
        self.path_var.set(selected[-1])
        self._select_first_row()
        self._emit_change()

    def add_manual_row(self) -> None:
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst einen Pfad eingeben oder Dateien waehlen.")
            return
        art = self.art_var.get().strip() or self.default_art
        sprache = self.sprache_var.get().strip() or self.default_sprache
        if self.infer_art and not self.art_var.get().strip():
            art = infer_document_art(path)
            self.art_var.set(art)
        self.tree.insert("", "end", values=(path, art, sprache))
        self._select_first_row(select_last=True)
        self._emit_change()

    def update_selected_row(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst eine Zeile zum Bearbeiten auswaehlen.")
            return

        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst einen Pfad oder eine URL eingeben.")
            return

        art = self.art_var.get().strip() or self.default_art
        sprache = self.sprache_var.get().strip() or self.default_sprache
        self.tree.item(selected[0], values=(path, art, sprache))
        self.tree.focus(selected[0])
        self._handle_selection()
        self._emit_change()

    def remove_selected(self) -> None:
        for item_id in self.tree.selection():
            self.tree.delete(item_id)
        self.clear_form()
        self._select_first_row()
        self._emit_change()

    def copy_selected_rows(self) -> None:
        selected_items = list(self.tree.selection())
        if not selected_items:
            return
        rows = []
        for item_id in selected_items:
            values = [str(value).strip() for value in self.tree.item(item_id, "values")]
            rows.append("\t".join(values))
        self.clipboard_clear()
        self.clipboard_append("\n".join(rows))

    def clear_form(self) -> None:
        self.path_var.set("")
        self.art_var.set(self.default_art)
        self.sprache_var.set(self.default_sprache)

    def _emit_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    def get_rows(self) -> list[MediaRow]:
        rows = []
        for item_id in self.tree.get_children():
            path, art, sprache = self.tree.item(item_id, "values")
            rows.append(MediaRow(path_or_link=str(path).strip(), art=str(art).strip(), sprache=str(sprache).strip()))
        return rows

    def set_rows(self, rows: list[MediaRow]) -> None:
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
        for row in rows:
            self.tree.insert("", "end", values=(row.path_or_link, row.art, row.sprache))
        self._select_first_row()

    def _select_first_row(self, select_last: bool = False) -> None:
        children = list(self.tree.get_children())
        if not children:
            self._reset_preview()
            return
        target = children[-1] if select_last else children[0]
        self.tree.selection_set(target)
        self.tree.focus(target)
        self._handle_selection()

    def _reset_preview(self) -> None:
        self.preview_photo = None
        self.preview_title_var.set("Keine Auswahl")
        self.preview_meta_var.set("Waehle eine Zeile aus, um mehr Details zu sehen.")
        self.preview_visual.configure(text="Keine Vorschau", image="")

    def _handle_selection(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        selected = self.tree.selection()
        if not selected:
            self._reset_preview()
            return
        path, art, sprache = self.tree.item(selected[0], "values")
        self.path_var.set(str(path).strip())
        self.art_var.set(str(art).strip())
        self.sprache_var.set(str(sprache).strip())
        self._update_preview(str(path).strip(), str(art).strip(), str(sprache).strip())

    def _open_context_menu(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.tree.identify_row(event.y)
        if item_id:
            if item_id not in self.tree.selection():
                self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self._handle_selection()
        has_selection = bool(self.tree.selection())
        self.context_menu.entryconfigure("Zeilen kopieren", state="normal" if has_selection else "disabled")
        self.context_menu.entryconfigure("Zeilen loeschen", state="normal" if has_selection else "disabled")
        self.context_menu.tk_popup(event.x_root, event.y_root)
        self.context_menu.grab_release()

    def _begin_inline_edit(self, event: tk.Event[tk.Misc]) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or column_id not in {"#1", "#2", "#3"}:
            return
        bbox = self.tree.bbox(item_id, column_id)
        if not bbox:
            return

        self._destroy_inline_editor()
        values = list(self.tree.item(item_id, "values"))
        column_index = int(column_id[1:]) - 1
        current_value = str(values[column_index]) if column_index < len(values) else ""
        x_pos, y_pos, width, height = bbox

        editor = ttk.Entry(self.tree)
        editor.place(x=x_pos, y=y_pos, width=width, height=height)
        editor.insert(0, current_value)
        editor.select_range(0, "end")
        editor.focus_set()
        editor.bind("<Return>", lambda _event: self._commit_inline_edit())
        editor.bind("<Escape>", lambda _event: self._destroy_inline_editor())
        editor.bind("<FocusOut>", lambda _event: self._commit_inline_edit())
        self.inline_editor = editor
        self.inline_editor_item = item_id
        self.inline_editor_column = column_id

    def _commit_inline_edit(self) -> None:
        if self.inline_editor is None or not self.inline_editor.winfo_exists():
            self.inline_editor = None
            return

        item_id = self.inline_editor_item
        column_id = self.inline_editor_column
        new_value = self.inline_editor.get().strip()
        self._destroy_inline_editor()

        if not item_id or column_id not in {"#1", "#2", "#3"}:
            return

        values = list(self.tree.item(item_id, "values"))
        column_index = int(column_id[1:]) - 1
        while len(values) < 3:
            values.append("")

        if column_index == 0 and not new_value:
            messagebox.showwarning(APP_TITLE, "Pfad oder URL darf nicht leer sein.")
            return
        if column_index == 1 and not new_value:
            new_value = self.default_art
        if column_index == 2 and not new_value:
            new_value = self.default_sprache

        values[column_index] = new_value
        if self.infer_art and column_index == 0 and not str(values[1]).strip():
            values[1] = infer_document_art(new_value)

        self.tree.item(item_id, values=values)
        self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        self._handle_selection()
        self._emit_change()

    def _destroy_inline_editor(self) -> None:
        if self.inline_editor is not None and self.inline_editor.winfo_exists():
            self.inline_editor.destroy()
        self.inline_editor = None
        self.inline_editor_item = ""
        self.inline_editor_column = ""

    def _update_preview(self, path_or_link: str, art: str, sprache: str) -> None:
        parsed = urlparse(path_or_link)
        raw_name = unquote(parsed.path if parsed.scheme else path_or_link)
        name = Path(raw_name).name or path_or_link
        extension = Path(raw_name).suffix.lower().replace(".", "").upper() or "LINK"
        self.preview_title_var.set(name)

        if self.preview_kind == "image":
            self._update_image_preview(path_or_link, art, sprache, extension)
            return

        if self.preview_kind == "document":
            self._update_document_preview(path_or_link, art, sprache, extension)
            return

        self.preview_photo = None
        self.preview_visual.configure(
            text=extension,
            image="",
            font=("Segoe UI Semibold", 18),
            width=18,
            anchor="center",
            justify="center",
        )
        self.preview_meta_var.set(
            f"Art: {art or '-'}\nSprache: {sprache or '-'}\nTyp: {extension}\nPfad/URL: {path_or_link}"
        )

    def _update_image_preview(self, path_or_link: str, art: str, sprache: str, extension: str) -> None:
        self.preview_meta_var.set(
            f"Art: {art or '-'}\nSprache: {sprache or '-'}\nTyp: {extension}\nPfad/URL: {path_or_link}"
        )
        if Image is None or ImageTk is None:
            self.preview_photo = None
            self.preview_visual.configure(text="Pillow fuer Bildvorschau installieren", image="", wraplength=220, justify="center")
            return

        try:
            image = self._load_preview_image(path_or_link)
        except Exception as exc:  # pragma: no cover - user-specific preview issues
            self.preview_photo = None
            self.preview_visual.configure(text=f"Keine Bildvorschau verfuegbar\n{exc}", image="", wraplength=220, justify="center")
            return

        image.thumbnail((260, 180))
        photo = ImageTk.PhotoImage(image)
        self.preview_photo = photo
        self.preview_visual.configure(image=photo, text="")

    def _update_document_preview(self, path_or_link: str, art: str, sprache: str, extension: str) -> None:
        self.preview_meta_var.set(
            f"Art: {art or '-'}\nSprache: {sprache or '-'}\nTyp: {extension}\nPfad/URL: {path_or_link}"
        )

        if extension == "PDF":
            self._update_pdf_preview(path_or_link)
            return

        self.preview_photo = None
        self.preview_visual.configure(
            text=extension,
            image="",
            font=("Segoe UI Semibold", 18),
            width=18,
            anchor="center",
            justify="center",
        )

    def _update_pdf_preview(self, path_or_link: str) -> None:
        if Image is None or ImageTk is None:
            self.preview_photo = None
            self.preview_visual.configure(text="Pillow fuer PDF-Vorschau installieren", image="", wraplength=220, justify="center")
            return
        if pymupdf is None:
            self.preview_photo = None
            self.preview_visual.configure(text="PyMuPDF fuer PDF-Vorschau installieren", image="", wraplength=220, justify="center")
            return

        try:
            image = self._render_pdf_preview(path_or_link)
        except Exception as exc:  # pragma: no cover - environment-specific PDF issues
            self.preview_photo = None
            self.preview_visual.configure(text=f"Keine PDF-Vorschau verfuegbar\n{exc}", image="", wraplength=220, justify="center")
            return

        image.thumbnail((260, 180))
        photo = ImageTk.PhotoImage(image)
        self.preview_photo = photo
        self.preview_visual.configure(image=photo, text="")

    def _load_preview_image(self, path_or_link: str) -> object:
        parsed = urlparse(path_or_link)
        if parsed.scheme in {"http", "https"}:
            request = urllib_request.Request(path_or_link, headers={"User-Agent": "ApolloImportGui/1.0"})
            with urllib_request.urlopen(request, timeout=12) as response:
                data = response.read()
            return Image.open(io.BytesIO(data))

        image_path = Path(path_or_link)
        if not image_path.exists():
            raise FileNotFoundError("Datei nicht gefunden")
        return Image.open(image_path)

    def _load_preview_bytes(self, path_or_link: str) -> bytes:
        parsed = urlparse(path_or_link)
        if parsed.scheme in {"http", "https"}:
            request = urllib_request.Request(path_or_link, headers={"User-Agent": "ApolloImportGui/1.0"})
            with urllib_request.urlopen(request, timeout=12) as response:
                return response.read()

        file_path = Path(path_or_link)
        if not file_path.exists():
            raise FileNotFoundError("Datei nicht gefunden")
        return file_path.read_bytes()

    def _render_pdf_preview(self, path_or_link: str) -> object:
        pdf_bytes = self._load_preview_bytes(path_or_link)
        document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        try:
            page = document[0]
            matrix = pymupdf.Matrix(1.5, 1.5)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            return pixmap.pil_image()
        finally:
            document.close()


class LinkTableFrame(ttk.LabelFrame):
    def __init__(
        self,
        master: tk.Misc,
        title: str,
        label_text: str,
        preview_kind: str = "web",
        normalize_link_fn: Callable[[str], str] | None = None,
        on_change: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master, text=title, padding=14)
        self.preview_kind = preview_kind
        self.normalize_link_fn = normalize_link_fn
        self.on_change = on_change
        self.inline_editor: ttk.Entry | None = None
        self.inline_editor_item = ""
        self.preview_photo: object | None = None
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Zeilen kopieren", command=self.copy_selected_rows)
        self.context_menu.add_command(label="Zeilen loeschen", command=self.remove_selected)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(2, weight=0)
        self.rowconfigure(1, weight=1)

        form = ttk.Frame(self)
        form.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        form.columnconfigure(1, weight=1)

        self.link_var = tk.StringVar()

        ttk.Label(form, text=label_text).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(form, textvariable=self.link_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(form, text="Hinzufuegen", command=self.add_row).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(form, text="Auswahl aktualisieren", command=self.update_selected_row).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(form, text="Auswahl entfernen", command=self.remove_selected).grid(row=0, column=4, padx=(8, 0))
        ttk.Button(form, text="Leeren", command=self.clear_form).grid(row=0, column=5, padx=(8, 0))

        self.tree = ttk.Treeview(self, columns=("link",), show="headings", height=7, selectmode="extended")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self.tree.heading("link", text=label_text)
        self.tree.column("link", width=780, anchor="w")

        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self._handle_selection)
        self.tree.bind("<Double-1>", self._begin_inline_edit)
        self.tree.bind("<Button-3>", self._open_context_menu)

        preview_frame = ttk.LabelFrame(self, text="Vorschau", padding=10)
        preview_frame.grid(row=1, column=2, sticky="nsew", padx=(12, 0))
        preview_frame.columnconfigure(0, weight=1)

        self.preview_title_var = tk.StringVar(value="Keine Auswahl")
        self.preview_meta_var = tk.StringVar(value="Waehle eine Zeile aus, um mehr Details zu sehen.")

        ttk.Label(preview_frame, textvariable=self.preview_title_var, font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w")

        self.preview_visual = ttk.Label(preview_frame, text="", anchor="center", justify="center")
        self.preview_visual.grid(row=1, column=0, sticky="ew", pady=(10, 10))

        ttk.Label(preview_frame, textvariable=self.preview_meta_var, foreground="#5E6472", wraplength=250, justify="left").grid(
            row=2, column=0, sticky="w"
        )
        self._reset_preview()

    def add_row(self) -> None:
        link = self.link_var.get().strip()
        if not link:
            messagebox.showwarning(APP_TITLE, "Bitte einen Link eingeben.")
            return
        if self.normalize_link_fn is not None:
            link = self.normalize_link_fn(link)
        self.tree.insert("", "end", values=(link,))
        children = list(self.tree.get_children())
        if children:
            self.tree.selection_set(children[-1])
            self.tree.focus(children[-1])
            self._handle_selection()
        self._emit_change()

    def update_selected_row(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst eine Zeile zum Bearbeiten auswaehlen.")
            return

        link = self.link_var.get().strip()
        if not link:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst einen Link eingeben.")
            return
        if self.normalize_link_fn is not None:
            link = self.normalize_link_fn(link)
        self.tree.item(selected[0], values=(link,))
        self.tree.focus(selected[0])
        self._handle_selection()
        self._emit_change()

    def remove_selected(self) -> None:
        for item_id in self.tree.selection():
            self.tree.delete(item_id)
        children = list(self.tree.get_children())
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
            self._handle_selection()
        else:
            self.clear_form()
            self._reset_preview()
        self._emit_change()

    def copy_selected_rows(self) -> None:
        selected_items = list(self.tree.selection())
        if not selected_items:
            return
        rows = []
        for item_id in selected_items:
            values = [str(value).strip() for value in self.tree.item(item_id, "values")]
            rows.append("\t".join(values))
        self.clipboard_clear()
        self.clipboard_append("\n".join(rows))

    def clear_form(self) -> None:
        self.link_var.set("")

    def _emit_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    def _handle_selection(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        selected = self.tree.selection()
        if not selected:
            self._reset_preview()
            return
        (link,) = self.tree.item(selected[0], "values")
        normalized_link = str(link).strip()
        self.link_var.set(normalized_link)
        self._update_preview(normalized_link)

    def _open_context_menu(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.tree.identify_row(event.y)
        if item_id:
            if item_id not in self.tree.selection():
                self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self._handle_selection()
        has_selection = bool(self.tree.selection())
        self.context_menu.entryconfigure("Zeilen kopieren", state="normal" if has_selection else "disabled")
        self.context_menu.entryconfigure("Zeilen loeschen", state="normal" if has_selection else "disabled")
        self.context_menu.tk_popup(event.x_root, event.y_root)
        self.context_menu.grab_release()

    def _begin_inline_edit(self, event: tk.Event[tk.Misc]) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or column_id != "#1":
            return
        bbox = self.tree.bbox(item_id, column_id)
        if not bbox:
            return

        self._destroy_inline_editor()
        (current_value,) = self.tree.item(item_id, "values")
        x_pos, y_pos, width, height = bbox
        editor = ttk.Entry(self.tree)
        editor.place(x=x_pos, y=y_pos, width=width, height=height)
        editor.insert(0, str(current_value))
        editor.select_range(0, "end")
        editor.focus_set()
        editor.bind("<Return>", lambda _event: self._commit_inline_edit())
        editor.bind("<Escape>", lambda _event: self._destroy_inline_editor())
        editor.bind("<FocusOut>", lambda _event: self._commit_inline_edit())
        self.inline_editor = editor
        self.inline_editor_item = item_id

    def _commit_inline_edit(self) -> None:
        if self.inline_editor is None or not self.inline_editor.winfo_exists():
            self.inline_editor = None
            return

        item_id = self.inline_editor_item
        link = self.inline_editor.get().strip()
        self._destroy_inline_editor()

        if not item_id:
            return
        if not link:
            messagebox.showwarning(APP_TITLE, "Link darf nicht leer sein.")
            return
        if self.normalize_link_fn is not None:
            link = self.normalize_link_fn(link)

        self.tree.item(item_id, values=(link,))
        self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        self._handle_selection()
        self._emit_change()

    def _destroy_inline_editor(self) -> None:
        if self.inline_editor is not None and self.inline_editor.winfo_exists():
            self.inline_editor.destroy()
        self.inline_editor = None
        self.inline_editor_item = ""

    def _select_first_row(self, select_last: bool = False) -> None:
        children = list(self.tree.get_children())
        if not children:
            self._reset_preview()
            return
        target = children[-1] if select_last else children[0]
        self.tree.selection_set(target)
        self.tree.focus(target)
        self._handle_selection()

    def _reset_preview(self) -> None:
        self.preview_photo = None
        self.preview_title_var.set("Keine Auswahl")
        self.preview_meta_var.set("Waehle eine Zeile aus, um mehr Details zu sehen.")
        self.preview_visual.configure(text="Keine Vorschau", image="", wraplength=220, justify="center")

    def _update_preview(self, link: str) -> None:
        parsed = urlparse(link)
        host = parsed.netloc.lower().removeprefix("www.") if parsed.netloc else "-"
        path = unquote(parsed.path or "/")

        if self.preview_kind == "video":
            self._update_video_preview(link, host, path)
            return

        self._update_web_preview(link, host, path)

    def _update_video_preview(self, link: str, host: str, path: str) -> None:
        video_id = extract_youtube_video_id(link)
        self.preview_title_var.set(f"Video: {video_id}" if video_id else (host or "Video-Link"))
        self.preview_meta_var.set(f"Host: {host or '-'}\nPfad: {path or '/'}\nLink: {link}")

        if not video_id or Image is None or ImageTk is None:
            self.preview_photo = None
            self.preview_visual.configure(text="VIDEO", image="", font=("Segoe UI Semibold", 18), wraplength=220, justify="center")
            return

        thumbnail_url = f"https://img.youtube.com/vi/{video_id}/hqdefault.jpg"
        try:
            image = self._load_remote_image(thumbnail_url)
        except Exception:  # pragma: no cover - preview only
            self.preview_photo = None
            self.preview_visual.configure(text="VIDEO", image="", font=("Segoe UI Semibold", 18), wraplength=220, justify="center")
            return

        image.thumbnail((260, 180))
        photo = ImageTk.PhotoImage(image)
        self.preview_photo = photo
        self.preview_visual.configure(image=photo, text="")

    def _update_web_preview(self, link: str, host: str, path: str) -> None:
        title = self._fetch_page_title(link)
        self.preview_title_var.set(title or host or "Web-Link")
        self.preview_meta_var.set(f"Host: {host or '-'}\nPfad: {path or '/'}\nLink: {link}")
        if Image is not None and ImageTk is not None:
            try:
                image, page_title = self._render_web_preview(link)
                if page_title:
                    self.preview_title_var.set(page_title)
                image.thumbnail((260, 180))
                photo = ImageTk.PhotoImage(image)
                self.preview_photo = photo
                self.preview_visual.configure(image=photo, text="")
                return
            except Exception:  # pragma: no cover - preview only
                pass

        badge = (host or "LINK").upper()
        self.preview_photo = None
        self.preview_visual.configure(text=badge, image="", font=("Segoe UI Semibold", 16), wraplength=220, justify="center")

    def _load_remote_image(self, url: str) -> object:
        request = urllib_request.Request(url, headers={"User-Agent": "ApolloImportGui/1.0"})
        with urllib_request.urlopen(request, timeout=10) as response:
            data = response.read()
        return Image.open(io.BytesIO(data))

    def _render_web_preview(self, link: str) -> tuple[object, str]:
        parsed = urlparse(link)
        if parsed.scheme not in {"http", "https"}:
            raise ValueError("Nur http/https Links koennen als Webseite vorgeladen werden.")

        browser_path = detect_browser_path()
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(browser_path) if browser_path else None,
                headless=True,
            )
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                ignore_https_errors=True,
            )
            page = context.new_page()
            try:
                page.goto(link, wait_until="domcontentloaded", timeout=20000)
                try:
                    page.wait_for_load_state("networkidle", timeout=5000)
                except Exception:
                    pass
                page.wait_for_timeout(800)
                screenshot = page.screenshot(full_page=False, type="png")
                title = " ".join(page.title().split()).strip()
            finally:
                context.close()
                browser.close()

        return Image.open(io.BytesIO(screenshot)), title

    def _fetch_page_title(self, link: str) -> str:
        parsed = urlparse(link)
        if parsed.scheme not in {"http", "https"}:
            return ""

        try:
            request = urllib_request.Request(link, headers={"User-Agent": "ApolloImportGui/1.0"})
            with urllib_request.urlopen(request, timeout=8) as response:
                raw = response.read(32768)
                charset = response.headers.get_content_charset() or "utf-8"
        except Exception:  # pragma: no cover - preview only
            return ""

        text = raw.decode(charset, errors="ignore")
        match = re.search(r"<title[^>]*>(.*?)</title>", text, flags=re.IGNORECASE | re.DOTALL)
        if not match:
            return ""
        return " ".join(unescape(match.group(1)).split()).strip()

    def get_rows(self) -> list[MediaRow]:
        rows = []
        for item_id in self.tree.get_children():
            (link,) = self.tree.item(item_id, "values")
            normalized = str(link).strip()
            if self.normalize_link_fn is not None:
                normalized = self.normalize_link_fn(normalized)
            rows.append(MediaRow(path_or_link=normalized))
        return rows

    def set_rows(self, rows: list[MediaRow]) -> None:
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
        for row in rows:
            link = row.path_or_link
            if self.normalize_link_fn is not None:
                link = self.normalize_link_fn(link)
            self.tree.insert("", "end", values=(link,))
        self._select_first_row()


class ApolloImportApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1360x920")
        self.root.minsize(1180, 780)

        self.id_registry = IdRegistry()
        self.kunzer_scraper = KunzerScraper()

        self.import_dir_var = tk.StringVar(value=str(DEFAULT_IMPORT_DIR))
        self.output_dir_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.product_list_path_var = tk.StringVar()
        self.deepl_api_key_var = tk.StringVar(value=os.getenv("DEEPL_API_KEY", ""))
        self.deepl_base_url_var = tk.StringVar(value=os.getenv("DEEPL_API_BASE_URL", DEEPL_DEFAULT_BASE_URL))
        self.kunzer_product_url_var = tk.StringVar()
        self.auto_translate_after_scrape_var = tk.BooleanVar(value=True)
        self.fixed_export_path_var = tk.BooleanVar(value=True)
        self.batch_short_text_var = tk.BooleanVar(value=True)
        self.batch_long_text_var = tk.BooleanVar(value=True)
        self.batch_image_var = tk.BooleanVar(value=True)
        self.batch_document_var = tk.BooleanVar(value=True)
        self.batch_video_var = tk.BooleanVar(value=True)
        self.batch_web_var = tk.BooleanVar(value=True)
        self.article_number_var = tk.StringVar()
        self.short_module_id_var = tk.StringVar()
        self.long_module_id_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Bereit.")
        self.known_id_count_var = tk.StringVar(value="0 IDs geladen")
        self.current_id_article_number = ""
        self.article_browser_records: dict[str, StoredArticleSnapshot] = {}
        self.live_write_suspended = False
        self.article_browser_context_menu: tk.Menu | None = None

        self._configure_style()
        self._build_layout()
        self._load_known_ids(initial=True)
        self.refresh_preview()

    def _configure_style(self) -> None:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("TFrame", background="#F6F2EA")
        style.configure("TLabelframe", background="#F6F2EA")
        style.configure("TLabelframe.Label", background="#F6F2EA", foreground="#2E4057", font=("Segoe UI Semibold", 11))
        style.configure("TLabel", background="#F6F2EA", foreground="#243447", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10))
        style.configure("TCheckbutton", background="#F6F2EA", foreground="#243447")
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        self.root.configure(background="#F6F2EA")

    def _build_layout(self) -> None:
        shell = ttk.Frame(self.root, padding=18)
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Apollo Import GUI", font=("Segoe UI Semibold", 20)).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Prototyp fuer die Erfassung eines Artikels und den Export der zugehoerigen Importdateien.",
            foreground="#5E6472",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        notebook = ttk.Notebook(shell)
        notebook.grid(row=1, column=0, sticky="nsew")

        self.project_tab = ttk.Frame(notebook, padding=18)
        self.short_tab = ttk.Frame(notebook, padding=18)
        self.long_tab = ttk.Frame(notebook, padding=18)
        self.image_tab = ttk.Frame(notebook, padding=18)
        self.document_tab = ttk.Frame(notebook, padding=18)
        self.links_tab = ttk.Frame(notebook, padding=18)

        notebook.add(self.project_tab, text="Projekt")
        notebook.add(self.short_tab, text="Kurzbezeichnung")
        notebook.add(self.long_tab, text="Text")
        notebook.add(self.image_tab, text="Bilder")
        notebook.add(self.document_tab, text="Dokumente")
        notebook.add(self.links_tab, text="Links")

        self._build_project_tab()
        self._build_short_tab()
        self._build_long_tab()
        self._build_media_tabs()

        status_bar = ttk.Label(shell, textvariable=self.status_var, anchor="w", foreground="#5E6472")
        status_bar.grid(row=2, column=0, sticky="ew", pady=(12, 0))

    def _build_project_tab(self) -> None:
        self.project_tab.columnconfigure(0, weight=1)
        self.project_tab.columnconfigure(1, weight=1)
        self.project_tab.rowconfigure(3, weight=1)

        file_frame = ttk.LabelFrame(self.project_tab, text="Ordner", padding=14)
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew")
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Importordner").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(file_frame, textvariable=self.import_dir_var).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(file_frame, text="Waehlen", command=self.choose_import_dir).grid(row=0, column=2, padx=(8, 0), pady=6)
        ttk.Button(file_frame, text="IDs laden", command=self._load_known_ids).grid(row=0, column=3, padx=(8, 0), pady=6)

        ttk.Label(file_frame, text="Ausgabeordner / Importpfad").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(file_frame, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(file_frame, text="Waehlen", command=self.choose_output_dir).grid(row=1, column=2, padx=(8, 0), pady=6)
        ttk.Label(file_frame, textvariable=self.known_id_count_var, foreground="#5E6472").grid(row=1, column=3, sticky="w", padx=(8, 0), pady=6)

        ttk.Label(file_frame, text="Produktliste CSV/XLSX").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(file_frame, textvariable=self.product_list_path_var).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(file_frame, text="Datei waehlen", command=self.choose_product_list_file).grid(row=2, column=2, padx=(8, 0), pady=6)
        ttk.Button(file_frame, text="Liste importieren", command=self.import_products_from_file).grid(row=2, column=3, padx=(8, 0), pady=6)

        article_frame = ttk.LabelFrame(self.project_tab, text="Artikel", padding=14)
        article_frame.grid(row=1, column=0, sticky="new", pady=(14, 0), padx=(0, 9))
        article_frame.columnconfigure(1, weight=1)

        ttk.Label(article_frame, text="Artikelnummer").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        article_entry = ttk.Entry(article_frame, textvariable=self.article_number_var)
        article_entry.grid(row=0, column=1, sticky="ew", pady=6)
        article_entry.bind("<FocusOut>", self._on_article_entry_focus_out)

        ttk.Label(
            article_frame,
            text="Kurz- und Text-ID werden automatisch im Hintergrund vergeben.",
            foreground="#5E6472",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(article_frame, text="Kunzer Produkt-URL").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        kunzer_url_entry = ttk.Entry(article_frame, textvariable=self.kunzer_product_url_var)
        kunzer_url_entry.grid(row=2, column=1, sticky="ew", pady=6)
        kunzer_url_entry.bind("<FocusOut>", self._on_live_field_focus_out)

        kunzer_row = ttk.Frame(article_frame)
        kunzer_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(kunzer_row, text="Aus Kunzer per Artikelnummer laden", command=self.load_from_kunzer_article_number).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(kunzer_row, text="Aus Kunzer per URL laden", command=self.load_from_kunzer_url).grid(row=0, column=1)

        options_row = ttk.Frame(article_frame)
        options_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(
            options_row,
            text="Nach dem Laden automatisch mit DeepL uebersetzen",
            variable=self.auto_translate_after_scrape_var,
        ).grid(row=0, column=0)

        button_row = ttk.Frame(article_frame)
        button_row.grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(button_row, text="Beispiel laden", command=self.load_demo_data).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(button_row, text="Artikelliste aktualisieren", command=self.refresh_preview).grid(row=0, column=1)

        export_frame = ttk.LabelFrame(self.project_tab, text="Export", padding=14)
        export_frame.grid(row=1, column=1, rowspan=2, sticky="nsew", pady=(14, 0), padx=(9, 0))
        export_frame.columnconfigure(0, weight=1)

        ttk.Label(
            export_frame,
            text="Du kannst die sieben Exportdateien entweder in einen neuen Zeitstempel-Unterordner schreiben oder immer direkt im Ausgabeordner aktualisieren.",
            wraplength=500,
            foreground="#5E6472",
        ).grid(row=0, column=0, sticky="w")

        ttk.Checkbutton(
            export_frame,
            text="Immer direkt in den Ausgabeordner schreiben (fester Importpfad)",
            variable=self.fixed_export_path_var,
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))

        ttk.Label(
            export_frame,
            text="Wenn aktiv, bleiben die Dateien bestehen: neue Artikel werden angehaengt und bestehende Zeilen derselben Artikelnummer ersetzt.",
            wraplength=500,
            foreground="#5E6472",
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        ttk.Button(export_frame, text="Export erstellen", style="Accent.TButton", command=self.export_current_bundle).grid(
            row=3, column=0, sticky="w", pady=(14, 0)
        )

        batch_frame = ttk.LabelFrame(export_frame, text="Listenimport", padding=10)
        batch_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(
            batch_frame,
            text="Waehle, welche Daten bei CSV/XLSX-Listen von Kunzer gescraped und direkt in die Output-Dateien geschrieben werden sollen.",
            wraplength=470,
            foreground="#5E6472",
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Checkbutton(batch_frame, text="Kurzbezeichnung", variable=self.batch_short_text_var).grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Checkbutton(batch_frame, text="Text", variable=self.batch_long_text_var).grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Checkbutton(batch_frame, text="Bilder", variable=self.batch_image_var).grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(batch_frame, text="Dokumente", variable=self.batch_document_var).grid(row=2, column=1, sticky="w", pady=(6, 0))
        ttk.Checkbutton(batch_frame, text="Videos", variable=self.batch_video_var).grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(batch_frame, text="Web Links", variable=self.batch_web_var).grid(row=3, column=1, sticky="w", pady=(6, 0))

        ttk.Label(
            batch_frame,
            text="Der Listenimport uebersetzt ausgewaehlte Texte automatisch und schreibt immer direkt in den festen Output-Pfad.",
            wraplength=470,
            foreground="#5E6472",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

        deepl_frame = ttk.LabelFrame(self.project_tab, text="DeepL", padding=14)
        deepl_frame.grid(row=2, column=0, sticky="ew", pady=(14, 0), padx=(0, 9))
        deepl_frame.columnconfigure(1, weight=1)

        ttk.Label(deepl_frame, text="API Key").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(deepl_frame, textvariable=self.deepl_api_key_var, show="*", width=60).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(deepl_frame, text="Base URL").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(deepl_frame, textvariable=self.deepl_base_url_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Label(
            deepl_frame,
            text="Pro: https://api.deepl.com  |  Free: https://api-free.deepl.com",
            foreground="#5E6472",
        ).grid(row=1, column=2, sticky="w", padx=(10, 0), pady=6)

        deepl_actions = ttk.Frame(deepl_frame)
        deepl_actions.grid(row=2, column=1, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(deepl_actions, text="Kurzbezeichnung uebersetzen", command=self.translate_short_texts).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(deepl_actions, text="Text uebersetzen", command=self.translate_long_texts).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(deepl_actions, text="Alles uebersetzen", command=self.translate_all_texts).grid(row=0, column=2)

        ttk.Label(
            deepl_frame,
            text="Die Uebersetzung laeuft jeweils aus dem deutschen Feld in EN, CZ, FR, IT und NL. UNI bleibt an Deutsch gekoppelt.",
            foreground="#5E6472",
            wraplength=1100,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

        browser_frame = ttk.LabelFrame(self.project_tab, text="Artikelverzeichnis", padding=14)
        browser_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(14, 0))
        browser_frame.columnconfigure(0, weight=2)
        browser_frame.columnconfigure(1, weight=1)
        browser_frame.rowconfigure(1, weight=1)

        browser_header = ttk.Frame(browser_frame)
        browser_header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        browser_header.columnconfigure(0, weight=1)
        ttk.Label(
            browser_header,
            text="Hier werden nur Artikel angezeigt, die bereits in den Output-Excel-Dateien vorhanden sind. Du kannst sie von dort erneut in die Maske laden.",
            foreground="#5E6472",
            wraplength=900,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(browser_header, text="Ausgewaehlten Artikel laden", command=self.load_selected_article_from_browser).grid(
            row=0, column=1, padx=(12, 8)
        )
        ttk.Button(browser_header, text="Liste aktualisieren", command=self.refresh_preview).grid(row=0, column=2)

        columns = ("article", "source", "short_id", "long_id", "images", "documents", "videos", "links")
        browser_table_frame = ttk.Frame(browser_frame)
        browser_table_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        browser_table_frame.columnconfigure(0, weight=1)
        browser_table_frame.rowconfigure(0, weight=1)

        self.article_browser_tree = ttk.Treeview(
            browser_table_frame,
            columns=columns,
            show="headings",
            height=10,
            selectmode="extended",
        )
        self.article_browser_tree.grid(row=0, column=0, sticky="nsew")
        self.article_browser_tree.heading("article", text="Artikelnummer")
        self.article_browser_tree.heading("source", text="Quelle")
        self.article_browser_tree.heading("short_id", text="Kurz-ID")
        self.article_browser_tree.heading("long_id", text="Text-ID")
        self.article_browser_tree.heading("images", text="Bilder")
        self.article_browser_tree.heading("documents", text="Dokumente")
        self.article_browser_tree.heading("videos", text="Videos")
        self.article_browser_tree.heading("links", text="Web Links")
        self.article_browser_tree.column("article", width=220, anchor="w")
        self.article_browser_tree.column("source", width=90, anchor="center")
        self.article_browser_tree.column("short_id", width=100, anchor="center")
        self.article_browser_tree.column("long_id", width=100, anchor="center")
        self.article_browser_tree.column("images", width=70, anchor="center")
        self.article_browser_tree.column("documents", width=90, anchor="center")
        self.article_browser_tree.column("videos", width=70, anchor="center")
        self.article_browser_tree.column("links", width=80, anchor="center")
        self.article_browser_tree.bind("<<TreeviewSelect>>", self._on_article_browser_select)
        self.article_browser_tree.bind("<Double-1>", self._on_article_browser_double_click)
        self.article_browser_tree.bind("<Button-3>", self._open_article_browser_context_menu)

        self.article_browser_context_menu = tk.Menu(browser_frame, tearoff=0)
        self.article_browser_context_menu.add_command(label="Zeilen kopieren", command=self.copy_selected_articles_from_browser)
        self.article_browser_context_menu.add_command(label="Zeilen loeschen", command=self.delete_selected_articles_from_browser)

        browser_scrollbar = ttk.Scrollbar(browser_table_frame, orient="vertical", command=self.article_browser_tree.yview)
        browser_scrollbar.grid(row=0, column=1, sticky="ns")
        self.article_browser_tree.configure(yscrollcommand=browser_scrollbar.set)

        detail_frame = ttk.LabelFrame(browser_frame, text="Exportdaten", padding=10)
        detail_frame.grid(row=1, column=1, sticky="nsew")
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(1, weight=1)

        ttk.Label(
            detail_frame,
            text="Die Detailansicht zeigt alle Daten, die aktuell in den Export-Excel-Dateien fuer den markierten Artikel vorhanden sind.",
            foreground="#5E6472",
            wraplength=380,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.article_browser_detail = ScrolledText(detail_frame, wrap="word", height=16, font=("Consolas", 9))
        self.article_browser_detail.grid(row=1, column=0, sticky="nsew")
        self.article_browser_detail.configure(state="disabled")

    def _build_short_tab(self) -> None:
        self.short_tab.columnconfigure(0, weight=1)
        self.short_tab.rowconfigure(0, weight=1)
        self.short_text_frame = SingleLineTranslationFrame(
            self.short_tab,
            "Kurzbezeichnung pro Sprache",
            on_change=lambda: self._write_live_section("short_text"),
        )
        self.short_text_frame.grid(row=0, column=0, sticky="nsew")

    def _build_long_tab(self) -> None:
        self.long_tab.columnconfigure(0, weight=1)
        self.long_tab.rowconfigure(0, weight=1)
        self.long_text_frame = MultiLineTranslationFrame(
            self.long_tab,
            "Langtext pro Sprache",
            on_change=lambda: self._write_live_section("long_text"),
        )
        self.long_text_frame.grid(row=0, column=0, sticky="nsew")

    def _build_media_tabs(self) -> None:
        self.image_tab.columnconfigure(0, weight=1)
        self.image_tab.rowconfigure(0, weight=1)
        self.document_tab.columnconfigure(0, weight=1)
        self.document_tab.rowconfigure(0, weight=1)
        self.links_tab.columnconfigure(0, weight=1)
        self.links_tab.rowconfigure(0, weight=1)
        self.links_tab.rowconfigure(1, weight=1)

        self.image_frame = MediaTableFrame(
            self.image_tab,
            title="Bilder",
            path_label="Bild-URL",
            path_dialog_title="Bilddateien auswaehlen",
            default_art="5",
            default_sprache="255",
            browse_filetypes=[("Bilddateien", "*.png *.jpg *.jpeg *.webp *.tif *.tiff"), ("Alle Dateien", "*.*")],
            preview_kind="image",
            on_change=lambda: self._write_live_section("images"),
        )
        self.image_frame.grid(row=0, column=0, sticky="nsew")

        self.document_frame = MediaTableFrame(
            self.document_tab,
            title="Dokumente",
            path_label="Dokument-URL",
            path_dialog_title="Dokumente auswaehlen",
            default_art="17",
            default_sprache="255",
            browse_filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")],
            infer_art=True,
            preview_kind="document",
            on_change=lambda: self._write_live_section("documents"),
        )
        self.document_frame.grid(row=0, column=0, sticky="nsew")

        self.video_frame = LinkTableFrame(
            self.links_tab,
            "Videos",
            "Video-Link",
            preview_kind="video",
            normalize_link_fn=normalize_youtube_url_to_public,
            on_change=lambda: self._write_live_section("videos"),
        )
        self.video_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        self.web_frame = LinkTableFrame(
            self.links_tab,
            "Web Links",
            "Web-Link",
            preview_kind="web",
            on_change=lambda: self._write_live_section("web_links"),
        )
        self.web_frame.grid(row=1, column=0, sticky="nsew")

    def choose_import_dir(self) -> None:
        selected = filedialog.askdirectory(title="Importordner waehlen", initialdir=self.import_dir_var.get())
        if selected:
            self.import_dir_var.set(selected)
            self._load_known_ids()
            self.refresh_preview()

    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="Ausgabeordner waehlen", initialdir=self.output_dir_var.get())
        if selected:
            self.output_dir_var.set(selected)
            self.status_var.set(f"Ausgabeordner gesetzt: {selected}")
            self._load_known_ids()
            self.refresh_preview()

    def choose_product_list_file(self) -> None:
        initial_dir = str(Path(self.product_list_path_var.get()).parent) if self.product_list_path_var.get().strip() else str(Path.cwd())
        selected = filedialog.askopenfilename(
            title="Produktliste CSV/XLSX waehlen",
            initialdir=initial_dir,
            filetypes=[("Importlisten", "*.csv *.xlsx"), ("CSV Dateien", "*.csv"), ("Excel Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if selected:
            self.product_list_path_var.set(selected)
            self.status_var.set(f"Produktliste gesetzt: {selected}")

    def _resolve_output_root(self) -> Path:
        output_root = Path(self.output_dir_var.get().strip())
        if not output_root:
            raise ValueError("Bitte einen Ausgabeordner angeben.")
        return output_root

    @contextmanager
    def _suspend_live_write(self):
        previous = self.live_write_suspended
        self.live_write_suspended = True
        try:
            yield
        finally:
            self.live_write_suspended = previous

    def _snapshot_from_bundle(self, bundle: ExportBundle, source_folder: Path) -> StoredArticleSnapshot:
        return StoredArticleSnapshot(
            article_number=bundle.article_number,
            source_label="Output",
            source_folder=source_folder,
            short_module_id=bundle.short_module_id,
            long_module_id=bundle.long_module_id,
            short_texts=bundle.short_texts,
            short_auto_uni=bundle.short_auto_uni,
            long_texts=bundle.long_texts,
            long_auto_uni=bundle.long_auto_uni,
            image_rows=bundle.image_rows,
            document_rows=bundle.document_rows,
            video_rows=bundle.video_rows,
            web_rows=bundle.web_rows,
        )

    def _copy_snapshot(self, snapshot: StoredArticleSnapshot) -> StoredArticleSnapshot:
        return StoredArticleSnapshot(
            article_number=snapshot.article_number,
            source_label=snapshot.source_label,
            source_folder=snapshot.source_folder,
            short_module_id=snapshot.short_module_id,
            long_module_id=snapshot.long_module_id,
            short_texts=TranslationSet(**vars(snapshot.short_texts)),
            short_auto_uni=snapshot.short_auto_uni,
            long_texts=TranslationSet(**vars(snapshot.long_texts)),
            long_auto_uni=snapshot.long_auto_uni,
            image_rows=[MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in snapshot.image_rows],
            document_rows=[MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in snapshot.document_rows],
            video_rows=[MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in snapshot.video_rows],
            web_rows=[MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in snapshot.web_rows],
        )

    def _render_article_browser(self, current_article: str | None = None) -> None:
        selected_article = current_article or (self.article_browser_tree.selection()[0] if self.article_browser_tree.selection() else "")
        selected_article = normalize_article_number(selected_article) or normalize_article_number(self.article_number_var.get())

        for item_id in self.article_browser_tree.get_children():
            self.article_browser_tree.delete(item_id)

        for article_number in sorted(self.article_browser_records):
            snapshot = self.article_browser_records[article_number]
            self.article_browser_tree.insert(
                "",
                "end",
                iid=article_number,
                values=(
                    snapshot.article_number,
                    snapshot.source_label,
                    snapshot.short_module_id,
                    snapshot.long_module_id,
                    len(snapshot.image_rows),
                    len(snapshot.document_rows),
                    len(snapshot.video_rows),
                    len(snapshot.web_rows),
                ),
            )

        if selected_article and selected_article in self.article_browser_records:
            self.article_browser_tree.selection_set(selected_article)
            self.article_browser_tree.focus(selected_article)
            self._update_article_browser_detail(self.article_browser_records[selected_article])
        else:
            self._update_article_browser_detail(None)

    def _upsert_article_browser_from_bundle(self, bundle: ExportBundle, output_root: Path) -> None:
        self.article_browser_records[bundle.article_number] = self._snapshot_from_bundle(bundle, output_root)
        self._render_article_browser(bundle.article_number)

    def _upsert_article_browser_section(self, section: str, bundle: ExportBundle, output_root: Path) -> None:
        existing_snapshot = self.article_browser_records.get(bundle.article_number)
        snapshot = self._copy_snapshot(existing_snapshot) if existing_snapshot is not None else StoredArticleSnapshot(
            article_number=bundle.article_number,
            source_label="Output",
            source_folder=output_root,
        )
        snapshot.article_number = bundle.article_number
        snapshot.source_label = "Output"
        snapshot.source_folder = output_root
        if bundle.short_module_id:
            snapshot.short_module_id = bundle.short_module_id
        if bundle.long_module_id:
            snapshot.long_module_id = bundle.long_module_id

        if section in {"short_text", "all"}:
            snapshot.short_texts = bundle.short_texts
            snapshot.short_auto_uni = bundle.short_auto_uni
            snapshot.short_module_id = bundle.short_module_id
        if section in {"long_text", "all"}:
            snapshot.long_texts = bundle.long_texts
            snapshot.long_auto_uni = bundle.long_auto_uni
            snapshot.long_module_id = bundle.long_module_id
        if section in {"images", "all"}:
            snapshot.image_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.image_rows]
        if section in {"documents", "all"}:
            snapshot.document_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.document_rows]
        if section in {"videos", "all"}:
            snapshot.video_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.video_rows]
        if section in {"web_links", "all"}:
            snapshot.web_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.web_rows]

        self.article_browser_records[bundle.article_number] = snapshot
        self._render_article_browser(bundle.article_number)

    def _write_short_text_live(self, bundle: ExportBundle, output_root: Path) -> None:
        written_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        article_keys = {bundle.article_number}
        write_workbook_with_upsert(
            output_root / SHORT_TEXT_FILE[0],
            SHORT_TEXT_FILE[1],
            SHORT_TEXT_HEADERS,
            [append_written_at(build_short_text_export_row(bundle), written_at)],
            replace_article_keys=article_keys,
        )
        write_workbook_with_upsert(
            output_root / SHORT_MAPPING_FILE[0],
            SHORT_MAPPING_FILE[1],
            SHORT_MAPPING_HEADERS,
            [append_written_at(build_short_mapping_export_row(bundle), written_at)],
            replace_article_keys=article_keys,
        )

    def _write_long_text_live(self, bundle: ExportBundle, output_root: Path) -> None:
        written_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        write_workbook_with_upsert(
            output_root / LONG_TEXT_FILE[0],
            LONG_TEXT_FILE[1],
            SHORT_TEXT_HEADERS,
            [append_written_at(build_long_text_export_row(bundle), written_at)],
            replace_article_keys={bundle.article_number},
        )

    def _write_media_section_live(
        self,
        bundle: ExportBundle,
        output_root: Path,
        file_spec: tuple[str, str],
        headers: list[str],
        rows: list[list[str]],
    ) -> None:
        replace_article_rows_preserving_timestamps(
            output_root / file_spec[0],
            file_spec[1],
            headers,
            bundle.article_number,
            rows,
        )

    def _write_live_section(self, section: str, status_message: str | None = None) -> bool:
        if self.live_write_suspended:
            return False

        try:
            output_root = self._resolve_output_root()
            bundle = self.collect_bundle()
        except ValueError:
            return False

        try:
            if section == "short_text":
                self._write_short_text_live(bundle, output_root)
            elif section == "long_text":
                self._write_long_text_live(bundle, output_root)
            elif section == "images":
                self._write_media_section_live(bundle, output_root, IMAGE_FILE, IMAGE_HEADERS, build_image_export_rows(bundle))
            elif section == "documents":
                self._write_media_section_live(bundle, output_root, DOCUMENT_FILE, DOCUMENT_HEADERS, build_document_export_rows(bundle))
            elif section == "videos":
                self._write_media_section_live(bundle, output_root, VIDEO_FILE, VIDEO_HEADERS, build_video_export_rows(bundle))
            elif section == "web_links":
                self._write_media_section_live(bundle, output_root, WEB_LINK_FILE, WEB_HEADERS, build_web_export_rows(bundle))
            else:
                raise ValueError(f"Unbekannter Live-Bereich: {section}")
        except Exception as exc:  # pragma: no cover - defensive UI feedback
            self.status_var.set(f"Live-Speichern fehlgeschlagen: {exc}")
            return False

        self.id_registry.remember_article_ids(bundle.article_number, bundle.short_module_id, bundle.long_module_id)
        self._upsert_article_browser_section(section, bundle, output_root)
        self.status_var.set(status_message or f"Live gespeichert: {bundle.article_number}")
        return True

    def _write_live_database(self, status_message: str | None = None) -> bool:
        if self.live_write_suspended:
            return False

        try:
            output_root = self._resolve_output_root()
            bundle = self.collect_bundle()
        except ValueError:
            return False

        try:
            export_bundle(bundle, output_root, use_timestamp_subdir=False)
        except Exception as exc:  # pragma: no cover - defensive UI feedback
            self.status_var.set(f"Live-Speichern fehlgeschlagen: {exc}")
            return False

        self.id_registry.remember_article_ids(bundle.article_number, bundle.short_module_id, bundle.long_module_id)
        self._upsert_article_browser_from_bundle(bundle, output_root)
        self.status_var.set(status_message or f"Live gespeichert: {bundle.article_number}")
        return True

    def _on_live_data_changed(self) -> None:
        self._write_live_database()

    def _on_article_entry_focus_out(self, _event: tk.Event[tk.Misc]) -> None:
        article_number = normalize_article_number(self.article_number_var.get())
        if not article_number:
            return
        self.article_number_var.set(article_number)
        self._ensure_ids_for_article(article_number)
        self._write_live_database()

    def _on_live_field_focus_out(self, _event: tk.Event[tk.Misc]) -> None:
        self._write_live_database()

    def _resolve_ids_for_article(self, article_number: str) -> tuple[str, str]:
        article_key = normalize_article_number(article_number)
        if not article_key:
            return "", ""

        short_id, long_id = self.id_registry.get_ids_for_article(article_key)
        if not short_id:
            short_id = self.id_registry.generate_unique()
        if not long_id:
            long_id = self.id_registry.generate_unique()

        self.id_registry.remember_article_ids(article_key, short_id, long_id)
        return short_id, long_id

    def _ensure_ids_for_article(self, article_number: str) -> tuple[str, str]:
        article_key = normalize_article_number(article_number)
        if not article_key:
            self.current_id_article_number = ""
            self.short_module_id_var.set("")
            self.long_module_id_var.set("")
            return "", ""

        current_short_id = self.short_module_id_var.get().strip()
        current_long_id = self.long_module_id_var.get().strip()
        if article_key == self.current_id_article_number and current_short_id and current_long_id:
            self.id_registry.remember_article_ids(article_key, current_short_id, current_long_id)
            return current_short_id, current_long_id

        short_id, long_id = self._resolve_ids_for_article(article_key)
        self.short_module_id_var.set(short_id)
        self.long_module_id_var.set(long_id)
        self.current_id_article_number = article_key
        return short_id, long_id

    def _get_batch_scrape_options(self) -> dict[str, bool]:
        options = {
            "short_text": self.batch_short_text_var.get(),
            "long_text": self.batch_long_text_var.get(),
            "images": self.batch_image_var.get(),
            "documents": self.batch_document_var.get(),
            "videos": self.batch_video_var.get(),
            "web_links": self.batch_web_var.get(),
        }
        if not any(options.values()):
            raise ValueError("Bitte fuer den Listenimport mindestens eine Datenart auswaehlen.")
        return options

    def _build_translation_set(self, german_text: str, client: DeepLClient | None) -> TranslationSet:
        values = {"de": german_text.strip(), "uni": german_text.strip()}
        if client and german_text.strip():
            values.update(client.translate_from_german(german_text))
        return TranslationSet(**values)

    def _build_bundle_from_kunzer_result(
        self,
        result: KunzerScrapeResult,
        client: DeepLClient | None = None,
        options: dict[str, bool] | None = None,
    ) -> ExportBundle:
        article_number = normalize_article_number(result.article_number)
        short_module_id, long_module_id = self._resolve_ids_for_article(article_number)
        scrape_options = options or {
            "short_text": True,
            "long_text": True,
            "images": True,
            "documents": True,
            "videos": True,
            "web_links": True,
        }
        short_texts = self._build_translation_set(result.short_text_de, client) if scrape_options["short_text"] else TranslationSet()
        long_texts = self._build_translation_set(result.long_text_de, client) if scrape_options["long_text"] else TranslationSet()
        return ExportBundle(
            article_number=article_number,
            short_module_id=short_module_id,
            long_module_id=long_module_id,
            short_texts=short_texts,
            short_auto_uni=True,
            long_texts=long_texts,
            long_auto_uni=True,
            image_rows=[MediaRow(link, art="5", sprache="255") for link in result.image_links] if scrape_options["images"] else [],
            document_rows=[MediaRow(link, art=infer_document_art(link), sprache="255") for link in result.document_links] if scrape_options["documents"] else [],
            video_rows=[MediaRow(link) for link in result.video_links] if scrape_options["videos"] else [],
            web_rows=[MediaRow(result.product_url)] if scrape_options["web_links"] else [],
            include_short_text=scrape_options["short_text"],
            include_long_text=scrape_options["long_text"],
            include_images=scrape_options["images"],
            include_documents=scrape_options["documents"],
            include_videos=scrape_options["videos"],
            include_web_links=scrape_options["web_links"],
        )

    def _apply_kunzer_result(
        self,
        result: KunzerScrapeResult,
        translate_after_load: bool = True,
        write_live: bool = True,
    ) -> None:
        with self._suspend_live_write():
            if not self.article_number_var.get().strip():
                self.article_number_var.set(result.article_number)
            else:
                self.article_number_var.set(normalize_article_number(result.article_number))
            self._ensure_ids_for_article(self.article_number_var.get())
            self.kunzer_product_url_var.set(result.product_url)

            self.short_text_frame.set_value(
                TranslationSet(de=result.short_text_de, uni=result.short_text_de),
                auto_uni=True,
            )
            self.long_text_frame.set_value(
                TranslationSet(de=result.long_text_de, uni=result.long_text_de),
                auto_uni=True,
            )

            image_rows = [MediaRow(link, art="5", sprache="255") for link in result.image_links]
            document_rows = [
                MediaRow(link, art=infer_document_art(link), sprache="255")
                for link in result.document_links
            ]
            video_rows = [MediaRow(link) for link in result.video_links]
            web_rows = [MediaRow(result.product_url)]

            self.image_frame.set_rows(image_rows)
            self.document_frame.set_rows(document_rows)
            self.video_frame.set_rows(video_rows)
            self.web_frame.set_rows(web_rows)

        if translate_after_load and self.auto_translate_after_scrape_var.get() and self.deepl_api_key_var.get().strip():
            self._translate_loaded_texts()

        self.refresh_preview()
        if write_live:
            self._write_live_database(status_message=f"Live gespeichert: {normalize_article_number(result.article_number)}")

    def _translate_loaded_texts(self) -> None:
        client = self._build_deepl_client()
        short_translations = client.translate_from_german(self.short_text_frame.get_german_text())
        long_translations = client.translate_from_german(self.long_text_frame.get_german_text())
        self.short_text_frame.apply_translations(short_translations)
        self.long_text_frame.apply_translations(long_translations)

    def _load_from_kunzer(self, article_number_or_url: str) -> None:
        try:
            self.status_var.set("Kunzer-Daten werden geladen ...")
            self.root.update_idletasks()
            result = self.kunzer_scraper.scrape_product(article_number_or_url)
            self._apply_kunzer_result(result)
            self.status_var.set(f"Kunzer-Daten geladen: {result.article_number}")
        except (KunzerScrapeError, ValueError, DeepLTranslationError) as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            self.status_var.set(f"Kunzer-Import fehlgeschlagen: {exc}")

    def load_from_kunzer_article_number(self) -> None:
        self._load_from_kunzer(self.article_number_var.get())

    def load_from_kunzer_url(self) -> None:
        self._load_from_kunzer(self.kunzer_product_url_var.get())

    def import_products_from_file(self) -> None:
        source_path = Path(self.product_list_path_var.get().strip())
        if not source_path:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst eine Produktliste als CSV oder XLSX auswaehlen.")
            return
        if not source_path.exists():
            messagebox.showwarning(APP_TITLE, f"Produktliste nicht gefunden:\n{source_path}")
            return

        try:
            output_root = self._resolve_output_root()
            options = self._get_batch_scrape_options()
            items = read_product_import_items(source_path)
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        deepl_client: DeepLClient | None = None
        try:
            if options["short_text"] or options["long_text"]:
                deepl_client = self._build_deepl_client()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        if output_root.exists():
            self.id_registry.load_from_folder(output_root, clear_existing=False)

        success_count = 0
        error_messages: list[str] = []
        last_result: KunzerScrapeResult | None = None
        last_bundle: ExportBundle | None = None
        total = len(items)

        for index, item in enumerate(items, start=1):
            identifier = item.product_url or item.article_number
            try:
                self.status_var.set(f"Listenimport {index}/{total}: {identifier}")
                self.root.update_idletasks()
                result = self.kunzer_scraper.scrape_product(identifier)
                bundle = self._build_bundle_from_kunzer_result(result, client=deepl_client, options=options)
                export_bundle(bundle, output_root, use_timestamp_subdir=False)
                success_count += 1
                last_result = result
                last_bundle = bundle
            except Exception as exc:  # pragma: no cover - user facing batch feedback
                error_messages.append(f"{identifier}: {exc}")

        if last_result and last_bundle:
            self._apply_kunzer_result(last_result, translate_after_load=False, write_live=False)
            self.short_module_id_var.set(last_bundle.short_module_id)
            self.long_module_id_var.set(last_bundle.long_module_id)
            if last_bundle.include_short_text:
                self.short_text_frame.set_value(last_bundle.short_texts, auto_uni=True)
            else:
                self.short_text_frame.set_value(TranslationSet(), auto_uni=True)
            if last_bundle.include_long_text:
                self.long_text_frame.set_value(last_bundle.long_texts, auto_uni=True)
            else:
                self.long_text_frame.set_value(TranslationSet(), auto_uni=True)
            if not last_bundle.include_images:
                self.image_frame.set_rows([])
            if not last_bundle.include_documents:
                self.document_frame.set_rows([])
            if not last_bundle.include_videos:
                self.video_frame.set_rows([])
            if not last_bundle.include_web_links:
                self.web_frame.set_rows([])

        self.refresh_preview()
        if error_messages:
            preview = "\n".join(error_messages[:8])
            if len(error_messages) > 8:
                preview += f"\n... und {len(error_messages) - 8} weitere"
            self.status_var.set(f"Listenimport abgeschlossen mit Fehlern: {success_count}/{total} erfolgreich")
            messagebox.showwarning(
                APP_TITLE,
                f"Listenimport abgeschlossen und direkt in den Output-Pfad geschrieben: {success_count} von {total} Produkten erfolgreich.\n\nFehler:\n{preview}",
            )
        else:
            self.status_var.set(f"Listenimport abgeschlossen: {success_count}/{total} Produkte erfolgreich")
            messagebox.showinfo(
                APP_TITLE,
                f"Listenimport erfolgreich abgeschlossen:\n{success_count} Produkte importiert und direkt in den Output-Pfad geschrieben.",
            )

    def _build_deepl_client(self) -> DeepLClient:
        auth_key = self.deepl_api_key_var.get().strip()
        if not auth_key:
            raise ValueError("Bitte einen DeepL API Key eingeben oder die Umgebungsvariable DEEPL_API_KEY setzen.")

        base_url = self.deepl_base_url_var.get().strip() or DEEPL_DEFAULT_BASE_URL
        return DeepLClient(auth_key=auth_key, base_url=base_url)

    def _translate_frame_from_german(self, frame: SingleLineTranslationFrame | MultiLineTranslationFrame, label: str) -> None:
        german_text = frame.get_german_text()
        if not german_text:
            raise ValueError(f"Bitte zuerst den deutschen Text fuer {label} eingeben.")

        client = self._build_deepl_client()
        translations = client.translate_from_german(german_text)
        frame.apply_translations(translations)
        self.refresh_preview()
        self._write_live_database(status_message=f"Live gespeichert: {normalize_article_number(self.article_number_var.get())}")
        self.status_var.set(f"{label} mit DeepL aus Deutsch uebersetzt.")

    def translate_short_texts(self) -> None:
        try:
            self._translate_frame_from_german(self.short_text_frame, "Kurzbezeichnung")
        except (ValueError, DeepLTranslationError) as exc:
            messagebox.showwarning(APP_TITLE, str(exc))

    def translate_long_texts(self) -> None:
        try:
            self._translate_frame_from_german(self.long_text_frame, "Text")
        except (ValueError, DeepLTranslationError) as exc:
            messagebox.showwarning(APP_TITLE, str(exc))

    def translate_all_texts(self) -> None:
        try:
            self._translate_loaded_texts()
            self.refresh_preview()
            self._write_live_database(status_message=f"Live gespeichert: {normalize_article_number(self.article_number_var.get())}")
            self.status_var.set("Kurzbezeichnung und Text mit DeepL aus Deutsch uebersetzt.")
        except (ValueError, DeepLTranslationError) as exc:
            messagebox.showwarning(APP_TITLE, str(exc))

    def _load_known_ids(self, initial: bool = False) -> None:
        import_folder = Path(self.import_dir_var.get().strip())
        output_folder = Path(self.output_dir_var.get().strip())
        warnings: list[str] = []
        loaded_from: list[Path] = []
        total_count = 0

        self.id_registry.used_ids.clear()
        self.id_registry.short_ids_by_article.clear()
        self.id_registry.long_ids_by_article.clear()

        if import_folder.exists():
            total_count, import_warnings = self.id_registry.load_from_folder(import_folder, clear_existing=True)
            warnings.extend(import_warnings)
            loaded_from.append(import_folder)

        if output_folder.exists() and output_folder != import_folder:
            total_count, output_warnings = self.id_registry.load_from_folder(output_folder, clear_existing=not loaded_from)
            warnings.extend(output_warnings)
            loaded_from.append(output_folder)

        if not loaded_from:
            self.known_id_count_var.set("0 IDs geladen")
            self.article_browser_records = {}
            for item_id in self.article_browser_tree.get_children():
                self.article_browser_tree.delete(item_id)
            self._update_article_browser_detail(None)
            if not initial:
                self.status_var.set(f"Keine ID-Quelle gefunden: {import_folder}")
            return

        current_article = normalize_article_number(self.article_number_var.get())
        if current_article:
            self.current_id_article_number = ""
            self._ensure_ids_for_article(current_article)

        self.known_id_count_var.set(f"{total_count} IDs geladen")
        if warnings:
            self.status_var.set(f"IDs geladen mit Hinweisen: {warnings[0]}")
        else:
            sources = ", ".join(str(path) for path in loaded_from)
            self.status_var.set(f"IDs erfolgreich geladen aus: {sources}")

    def load_demo_data(self) -> None:
        with self._suspend_live_write():
            self.article_number_var.set("WK DEMO-1000")
            self._ensure_ids_for_article(self.article_number_var.get())
            self.short_text_frame.set_value(
                TranslationSet(
                    de="Hydraulischer Demo-Heber 10 t",
                    en="Hydraulic demo jack 10 t",
                    cz="Hydraulicky demo zvedak 10 t",
                    fr="Cric hydraulique de demonstration 10 t",
                    it="Martinetto idraulico demo 10 t",
                    nl="Hydraulische demo krik 10 t",
                    uni="Hydraulischer Demo-Heber 10 t",
                ),
                auto_uni=True,
            )
            self.long_text_frame.set_value(
                TranslationSet(
                    de="- Demo-Produkt fuer die GUI\n- Zeigt den Ablauf fuer neue Artikel\n- Exportiert die benoetigten Excel-Dateien",
                    en="- Demo product for the GUI\n- Shows the workflow for new articles\n- Exports the required Excel files",
                    cz="- Demo produkt pro GUI\n- Ukazuje postup pro nove polozky\n- Exportuje potrebne Excel soubory",
                    fr="- Produit de demonstration pour l'interface\n- Montre le flux pour les nouveaux articles\n- Exporte les fichiers Excel necessaires",
                    it="- Prodotto demo per la GUI\n- Mostra il flusso per nuovi articoli\n- Esporta i file Excel necessari",
                    nl="- Demo-product voor de GUI\n- Toont de workflow voor nieuwe artikelen\n- Exporteert de benodigde Excel-bestanden",
                    uni="- Demo-Produkt fuer die GUI\n- Zeigt den Ablauf fuer neue Artikel\n- Exportiert die benoetigten Excel-Dateien",
                ),
                auto_uni=True,
            )
            self.image_frame.set_rows(
                [
                    MediaRow(r"S:\Apollo\WK DEMO-1000\wk-demo-1000-komplettansicht.png", art="5", sprache="255"),
                    MediaRow(r"S:\Apollo\WK DEMO-1000\wk-demo-1000-frontansicht.png", art="5", sprache="255"),
                ]
            )
            self.document_frame.set_rows(
                [
                    MediaRow(r"S:\dsp3\Web\WK DEMO-1000\WK DEMO-1000 Produktinfo.pdf", art="17", sprache="255"),
                    MediaRow(r"S:\dsp3\Web\WK DEMO-1000\WK DEMO-1000 Bedienungsanleitung.pdf", art="14", sprache="255"),
                ]
            )
            self.video_frame.set_rows([MediaRow("https://youtube.com/shorts/demo-video-1000")])
            self.web_frame.set_rows([MediaRow("https://www.kunzer.de/shop/p/WK%20DEMO-1000")])
        self.status_var.set("Beispieldaten geladen.")
        self.refresh_preview()
        self._write_live_database(status_message="Live gespeichert: WK DEMO-1000")

    def _refresh_article_browser(self) -> None:
        records: dict[str, StoredArticleSnapshot] = {}
        output_folder = Path(self.output_dir_var.get().strip())

        if output_folder.exists():
            records = load_article_snapshots_from_folder(output_folder, "Output")

        self.article_browser_records = records
        self._render_article_browser()

    def _update_article_browser_detail(self, snapshot: StoredArticleSnapshot | None) -> None:
        self.article_browser_detail.configure(state="normal")
        self.article_browser_detail.delete("1.0", "end")
        if snapshot is None:
            self.article_browser_detail.insert(
                "1.0",
                "Waehle einen Artikel im Verzeichnis aus, um alle vorhandenen Exportdaten anzuzeigen.\n\n"
                "Mit Doppelklick oder dem Button kannst du den Artikel anschliessend in die Bearbeitungsmaske laden.",
            )
        else:
            self.article_browser_detail.insert("1.0", format_article_snapshot(snapshot))
        self.article_browser_detail.configure(state="disabled")

    def _update_article_browser_multi_detail(self, article_numbers: list[str]) -> None:
        self.article_browser_detail.configure(state="normal")
        self.article_browser_detail.delete("1.0", "end")
        self.article_browser_detail.insert(
            "1.0",
            "Mehrere Artikel markiert.\n\n"
            f"Anzahl: {len(article_numbers)}\n"
            "Aktionen:\n"
            "- Rechtsklick fuer Kopieren oder Loeschen\n"
            "- Doppelklick oder 'Ausgewaehlten Artikel laden' ist fuer einen einzelnen Artikel gedacht\n\n"
            "Markierte Artikel:\n"
            + "\n".join(article_numbers),
        )
        self.article_browser_detail.configure(state="disabled")

    def _apply_article_snapshot(self, snapshot: StoredArticleSnapshot) -> None:
        self.article_number_var.set(snapshot.article_number)
        self.current_id_article_number = snapshot.article_number
        self.short_module_id_var.set(snapshot.short_module_id)
        self.long_module_id_var.set(snapshot.long_module_id)
        self.id_registry.remember_article_ids(snapshot.article_number, snapshot.short_module_id, snapshot.long_module_id)
        self.short_text_frame.set_value(snapshot.short_texts, auto_uni=snapshot.short_auto_uni)
        self.long_text_frame.set_value(snapshot.long_texts, auto_uni=snapshot.long_auto_uni)
        self.image_frame.set_rows(snapshot.image_rows)
        self.document_frame.set_rows(snapshot.document_rows)
        self.video_frame.set_rows(snapshot.video_rows)
        self.web_frame.set_rows(snapshot.web_rows)
        if snapshot.web_rows:
            self.kunzer_product_url_var.set(snapshot.web_rows[0].path_or_link)
        else:
            self.kunzer_product_url_var.set("")
        self.status_var.set(f"Artikel geladen: {snapshot.article_number} ({snapshot.source_label})")

    def load_selected_article_from_browser(self) -> None:
        selection = self.article_browser_tree.selection()
        if not selection:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst einen Artikel im Verzeichnis auswaehlen.")
            return
        if len(selection) > 1:
            messagebox.showwarning(APP_TITLE, "Bitte zum Laden genau einen Artikel im Verzeichnis auswaehlen.")
            return

        snapshot = self.article_browser_records.get(selection[0])
        if snapshot is None:
            messagebox.showwarning(APP_TITLE, "Der ausgewaehlte Artikel konnte nicht geladen werden.")
            return

        self._apply_article_snapshot(snapshot)
        self.refresh_preview()

    def _on_article_browser_select(self, _event: tk.Event[tk.Misc]) -> None:
        selection = self.article_browser_tree.selection()
        if not selection:
            self._update_article_browser_detail(None)
            return
        if len(selection) > 1:
            self._update_article_browser_multi_detail(list(selection))
            return
        snapshot = self.article_browser_records.get(selection[0])
        self._update_article_browser_detail(snapshot)

    def _on_article_browser_double_click(self, _event: tk.Event[tk.Misc]) -> None:
        self.load_selected_article_from_browser()

    def _open_article_browser_context_menu(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.article_browser_tree.identify_row(event.y)
        if item_id:
            if item_id not in self.article_browser_tree.selection():
                self.article_browser_tree.selection_set(item_id)
            self.article_browser_tree.focus(item_id)
            self._on_article_browser_select(event)

        has_selection = bool(self.article_browser_tree.selection())
        if self.article_browser_context_menu is None:
            return
        self.article_browser_context_menu.entryconfigure("Zeilen kopieren", state="normal" if has_selection else "disabled")
        self.article_browser_context_menu.entryconfigure("Zeilen loeschen", state="normal" if has_selection else "disabled")
        self.article_browser_context_menu.tk_popup(event.x_root, event.y_root)
        self.article_browser_context_menu.grab_release()

    def copy_selected_articles_from_browser(self) -> None:
        selection = list(self.article_browser_tree.selection())
        if not selection:
            return

        rows = []
        for item_id in selection:
            values = [str(value).strip() for value in self.article_browser_tree.item(item_id, "values")]
            rows.append("\t".join(values))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(rows))
        self.status_var.set(f"{len(selection)} Artikelzeile(n) in die Zwischenablage kopiert.")

    def delete_selected_articles_from_browser(self) -> None:
        selection = list(self.article_browser_tree.selection())
        if not selection:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst mindestens einen Artikel im Verzeichnis auswaehlen.")
            return

        article_numbers = [normalize_article_number(item_id) for item_id in selection if normalize_article_number(item_id)]
        if not article_numbers:
            return

        article_preview = "\n".join(article_numbers[:8])
        if len(article_numbers) > 8:
            article_preview += f"\n... und {len(article_numbers) - 8} weitere"
        confirmed = messagebox.askyesno(
            APP_TITLE,
            "Sollen die markierten Artikel wirklich aus allen Output-Dateien geloescht werden?\n\n" + article_preview,
        )
        if not confirmed:
            return

        try:
            output_root = self._resolve_output_root()
            for file_spec, headers in [
                (SHORT_TEXT_FILE, SHORT_TEXT_HEADERS),
                (SHORT_MAPPING_FILE, SHORT_MAPPING_HEADERS),
                (LONG_TEXT_FILE, SHORT_TEXT_HEADERS),
                (IMAGE_FILE, IMAGE_HEADERS),
                (DOCUMENT_FILE, DOCUMENT_HEADERS),
                (VIDEO_FILE, VIDEO_HEADERS),
                (WEB_LINK_FILE, WEB_HEADERS),
            ]:
                remove_article_rows_from_workbook(output_root / file_spec[0], file_spec[1], headers, set(article_numbers))
        except Exception as exc:  # pragma: no cover - defensive UI feedback
            messagebox.showerror(APP_TITLE, f"Artikel konnten nicht geloescht werden:\n{exc}")
            self.status_var.set(f"Loeschen fehlgeschlagen: {exc}")
            return

        for article_number in article_numbers:
            self.article_browser_records.pop(article_number, None)
        self._load_known_ids(initial=True)
        self._refresh_article_browser()
        self.status_var.set(f"{len(article_numbers)} Artikel aus den Output-Dateien geloescht.")

    def collect_bundle(self) -> ExportBundle:
        article_number = normalize_article_number(self.article_number_var.get())
        if not article_number:
            raise ValueError("Bitte eine Artikelnummer eingeben.")

        short_module_id, long_module_id = self._ensure_ids_for_article(article_number)

        return ExportBundle(
            article_number=article_number,
            short_module_id=short_module_id,
            long_module_id=long_module_id,
            short_texts=self.short_text_frame.get_value(),
            short_auto_uni=self.short_text_frame.auto_uni_var.get(),
            long_texts=self.long_text_frame.get_value(),
            long_auto_uni=self.long_text_frame.auto_uni_var.get(),
            image_rows=self.image_frame.get_rows(),
            document_rows=self.document_frame.get_rows(),
            video_rows=self.video_frame.get_rows(),
            web_rows=self.web_frame.get_rows(),
        )

    def refresh_preview(self) -> None:
        current_article = normalize_article_number(self.article_number_var.get())
        if current_article:
            self._ensure_ids_for_article(current_article)
        self._refresh_article_browser()

    def export_current_bundle(self) -> None:
        try:
            bundle = self.collect_bundle()
            output_root = self._resolve_output_root()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        try:
            export_dir = export_bundle(bundle, output_root, use_timestamp_subdir=not self.fixed_export_path_var.get())
        except Exception as exc:  # pragma: no cover - defensive user feedback
            messagebox.showerror(APP_TITLE, f"Export fehlgeschlagen:\n{exc}")
            self.status_var.set(f"Export fehlgeschlagen: {exc}")
            return

        self.refresh_preview()
        if self.fixed_export_path_var.get():
            self.status_var.set(f"Exportdateien im festen Pfad aktualisiert: {export_dir}")
            messagebox.showinfo(APP_TITLE, f"Exportdateien erfolgreich aktualisiert:\n{export_dir}")
        else:
            self.status_var.set(f"Export erstellt: {export_dir}")
            messagebox.showinfo(APP_TITLE, f"Export erfolgreich erstellt:\n{export_dir}")


def main() -> None:
    root = tk.Tk()
    app = ApolloImportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
