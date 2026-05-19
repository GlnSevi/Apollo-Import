from __future__ import annotations

import csv
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from html import unescape
import io
import json
import os
from pathlib import Path
import random
import re
import string
import threading
import sys
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
DEFAULT_GENART_SOURCE = Path(r"C:\Users\heimbuchner\Downloads\Genarten.xlsx")
DEEPL_DEFAULT_BASE_URL = "https://api.deepl.com"
ID_ALPHABET = string.ascii_uppercase + string.digits
ID_LENGTH = 6
SHORT_TEXT_MAX_LENGTH = 60
GENART_AUTO_APPLY_MIN_SCORE = 360
ATTACHMENT_FORMAT_TYPE_HEADER = "TecDoc Anhangsformattyp ID"
LAST_WRITTEN_HEADER = "Zuletzt geschrieben am"
SHORT_TEXT_UMLAUT_REPLACEMENTS = str.maketrans(
    {
        "ä": "ae",
        "ö": "oe",
        "ü": "ue",
        "Ä": "Ae",
        "Ö": "Oe",
        "Ü": "Ue",
        "ß": "ss",
    }
)
SESSION_STATE_FILE = Path.cwd() / "apollo_import_gui_state.json"
APP_ICON_PNG_RELATIVE_PATH = Path("assets") / "apollo_import_logo.png"
APP_ICON_ICO_RELATIVE_PATH = Path("assets") / "apollo_import_logo.ico"

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
GENART_FILE = ("GenArt_Artikel.xlsx", "GenArt")

IMAGE_HEADERS = ["Artikelnummer", "BILDPFAD", "Art", "Sprache", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
DOCUMENT_HEADERS = ["Artikelnummer", "Pfad zum Dokument", "Sprache", "Art", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
VIDEO_HEADERS = ["Produktnummer", "Link", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
WEB_HEADERS = ["Artikelnummer ", "Link", ATTACHMENT_FORMAT_TYPE_HEADER, LAST_WRITTEN_HEADER]
SHORT_MAPPING_HEADERS = ["Artikelnummer", "Text Modul ID", LAST_WRITTEN_HEADER]
GENART_HEADERS = ["Artikelnummer", "GenArt ID", "GenArt Bezeichnung", LAST_WRITTEN_HEADER]
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

GENART_FAMILY_TERMS = {
    "light": frozenset(
        {
            "lampe",
            "leuchte",
            "licht",
            "lamp",
            "light",
            "worklight",
            "flashlight",
            "headlamp",
            "inspection",
            "arbeitslampe",
            "meisterlampe",
            "prueflampe",
            "pruefleuchte",
            "stablampe",
            "taschenlampe",
            "stirnlampe",
            "inspectionlamp",
            "inspectionlight",
            "werkzeugleuchte",
        }
    ),
    "pin": frozenset(
        {
            "stift",
            "splint",
            "pin",
            "peg",
            "stiftschraube",
            "zylinderstift",
            "zentrierstift",
            "verriegelungsstift",
            "haltestift",
        }
    ),
}

GENART_COMPOUND_SUFFIXES = frozenset(
    {
        *{term for terms in GENART_FAMILY_TERMS.values() for term in terms},
        "zange",
        "hammer",
        "buerste",
        "schluessel",
    }
)

GENART_OPPOSING_FAMILY_PENALTIES = {
    ("light", "pin"): 180,
    ("pin", "light"): 180,
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


@dataclass(frozen=True)
class ImageSignature:
    average_hash: int
    difference_hash: int


@dataclass
class GenArtOption:
    id: str
    bezeichnung: str
    genart: str
    normalized_bezeichnung: str = ""
    normalized_genart: str = ""
    tokens: frozenset[str] = field(default_factory=frozenset)
    families: frozenset[str] = field(default_factory=frozenset)

    def display_label(self) -> str:
        if self.genart and self.bezeichnung and self.genart != self.bezeichnung:
            return f"{self.id} | {self.genart} | {self.bezeichnung}"
        if self.bezeichnung:
            return f"{self.id} | {self.bezeichnung}"
        return self.id


@dataclass
class GenArtSuggestion:
    option: GenArtOption
    total_score: float
    text_score: float = 0.0
    image_score: float = 0.0
    web_score: float = 0.0
    web_reason: str = ""


@dataclass
class GoogleLensWebResult:
    headline_lines: list[str] = field(default_factory=list)
    result_titles: list[str] = field(default_factory=list)
    result_snippets: list[str] = field(default_factory=list)
    page_urls: list[str] = field(default_factory=list)

    def as_context_text(self) -> str:
        return " ".join(
            part
            for part in [
                " ".join(self.headline_lines),
                " ".join(self.result_titles),
                " ".join(self.result_snippets),
                " ".join(self.page_urls),
            ]
            if part.strip()
        ).strip()


@dataclass
class ExportBundle:
    article_number: str
    short_module_id: str
    long_module_id: str
    short_texts: TranslationSet
    short_auto_uni: bool
    long_texts: TranslationSet
    long_auto_uni: bool
    genart_id: str = ""
    genart_bezeichnung: str = ""
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
    genart_id: str = ""
    genart_bezeichnung: str = ""
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


def replace_short_text_umlauts(value: str) -> str:
    return value.translate(SHORT_TEXT_UMLAUT_REPLACEMENTS)


def get_application_base_path() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", "")
        if meipass:
            return Path(meipass)
    return Path(__file__).resolve().parent


def resolve_application_asset_path(relative_path: str | Path) -> Path:
    return get_application_base_path() / Path(relative_path)


def sanitize_short_translation_set(translations: TranslationSet) -> TranslationSet:
    values: dict[str, str] = {}
    for code in [language_code for language_code, _label in UI_LANGUAGE_ORDER]:
        sanitized = replace_short_text_umlauts(getattr(translations, code))
        values[code] = sanitized[:SHORT_TEXT_MAX_LENGTH]
    return TranslationSet(**values)


def normalize_match_text(value: str) -> str:
    normalized = (
        value.casefold()
        .replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    return " ".join(normalized.split())


def build_match_tokens(value: str) -> frozenset[str]:
    normalized = normalize_match_text(value)
    base_tokens = {token for token in normalized.split() if token}
    expanded_tokens = set(base_tokens)

    for token in list(base_tokens):
        for suffix in GENART_COMPOUND_SUFFIXES:
            if token != suffix and len(token) > len(suffix) + 2 and token.endswith(suffix):
                expanded_tokens.add(suffix)

    return frozenset(expanded_tokens)


def infer_match_families_from_tokens(tokens: set[str] | frozenset[str]) -> frozenset[str]:
    token_set = set(tokens)
    families = set()
    for family_name, terms in GENART_FAMILY_TERMS.items():
        if token_set & terms:
            families.add(family_name)
    return frozenset(families)


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


def extract_match_text_from_path_or_link(path_or_link: str) -> str:
    value = path_or_link.strip()
    if not value:
        return ""

    parsed = urlparse(value)
    raw_path = unquote(parsed.path if parsed.scheme else value)
    path = Path(raw_path)
    parts: list[str] = []

    if parsed.netloc:
        parts.append(parsed.netloc.replace(".", " "))
    if path.stem:
        parts.append(path.stem)

    for part in path.parts[-4:-1]:
        cleaned = str(part).strip("/\\")
        if cleaned and cleaned not in {path.name, "."}:
            parts.append(cleaned)

    return normalize_match_text(" ".join(parts))


def _get_lanczos_resample() -> object:
    if Image is None:  # pragma: no cover - guarded by caller
        raise RuntimeError("Pillow ist nicht installiert.")
    if hasattr(Image, "Resampling"):
        return Image.Resampling.LANCZOS
    return Image.LANCZOS


def load_image_for_signature(path_or_link: str, timeout_seconds: int = 10) -> object:
    if Image is None:  # pragma: no cover - guarded by caller
        raise RuntimeError("Pillow ist nicht installiert.")

    parsed = urlparse(path_or_link.strip())
    if parsed.scheme in {"http", "https"}:
        request = urllib_request.Request(path_or_link, headers={"User-Agent": "ApolloImportGui/1.0"})
        with urllib_request.urlopen(request, timeout=timeout_seconds) as response:
            data = response.read()
        with Image.open(io.BytesIO(data)) as opened:
            return opened.convert("RGB")

    image_path = Path(path_or_link)
    if not image_path.exists():
        raise FileNotFoundError("Datei nicht gefunden")
    with Image.open(image_path) as opened:
        return opened.convert("RGB")


def build_image_signature(path_or_link: str) -> ImageSignature:
    image = load_image_for_signature(path_or_link)
    resample = _get_lanczos_resample()

    grayscale = image.convert("L")
    average_image = grayscale.resize((8, 8), resample)
    average_pixels = list(average_image.getdata())
    pixel_average = sum(average_pixels) / max(1, len(average_pixels))
    average_hash = 0
    for pixel in average_pixels:
        average_hash = (average_hash << 1) | int(pixel >= pixel_average)

    difference_image = grayscale.resize((9, 8), resample)
    difference_pixels = list(difference_image.getdata())
    difference_hash = 0
    for row_index in range(8):
        for column_index in range(8):
            left_pixel = difference_pixels[row_index * 9 + column_index]
            right_pixel = difference_pixels[row_index * 9 + column_index + 1]
            difference_hash = (difference_hash << 1) | int(left_pixel >= right_pixel)

    return ImageSignature(average_hash=average_hash, difference_hash=difference_hash)


def compare_image_signatures(first: ImageSignature, second: ImageSignature) -> float:
    average_distance = (first.average_hash ^ second.average_hash).bit_count()
    difference_distance = (first.difference_hash ^ second.difference_hash).bit_count()
    max_distance = 128
    return max(0.0, 1.0 - ((average_distance + difference_distance) / max_distance))


def build_google_lens_search_url(image_url: str) -> str:
    return f"https://lens.google.com/uploadbyurl?url={quote(image_url, safe=':/?&=%#')}&hl=de"


def prepare_local_image_for_lens(path_or_link: str) -> Path | None:
    value = path_or_link.strip()
    if not value:
        return None

    parsed = urlparse(value)
    if parsed.scheme in {"http", "https"}:
        return None

    file_path = Path(value)
    if not file_path.exists():
        return None

    if Image is not None:
        with Image.open(file_path) as opened:
            converted = opened.convert("RGB")
            converted.thumbnail((1400, 1400), _get_lanczos_resample())
            temp_path = Path(os.getenv("TEMP", str(Path.cwd()))) / f"apollo_lens_{random.randint(100000, 999999)}.jpg"
            converted.save(temp_path, format="JPEG", quality=88, optimize=True)
        return temp_path

    return file_path


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


def normalize_youtube_url_for_embed(url: str) -> str:
    value = url.strip()
    parsed = urlparse(value)
    host = parsed.netloc.lower()
    path = parsed.path.strip("/")
    query = parse_qs(parsed.query)
    video_id = ""

    if host in {"youtu.be", "www.youtu.be"}:
        video_id = path.split("/", 1)[0]
    elif host.endswith("youtube.com") or host.endswith("youtube-nocookie.com"):
        parts = [part for part in path.split("/") if part]
        if len(parts) >= 2 and parts[0] in {"embed", "shorts", "live"}:
            video_id = parts[1]
        elif parts[:1] == ["watch"]:
            video_id = (query.get("v") or [""])[0]

    if not video_id:
        return value

    passthrough: dict[str, str] = {}
    for key in ["start", "t"]:
        if query.get(key):
            passthrough[key] = query[key][0]

    if "t" in passthrough and "start" not in passthrough:
        passthrough["start"] = passthrough.pop("t")

    query_string = f"?{urlencode(passthrough)}" if passthrough else ""
    return f"https://www.youtube.com/embed/{video_id}{query_string}"


def extract_youtube_video_id(url: str) -> str:
    value = url.strip()
    parsed = urlparse(value)
    host = parsed.netloc.lower()
    path = parsed.path.strip("/")
    query = parse_qs(parsed.query)

    if host in {"youtu.be", "www.youtu.be"}:
        return path.split("/", 1)[0]
    if host.endswith("youtube.com") or host.endswith("youtube-nocookie.com"):
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


class GenArtRegistry:
    def __init__(self) -> None:
        self.options: list[GenArtOption] = []
        self.options_by_display: dict[str, GenArtOption] = {}
        self.options_by_id: dict[str, GenArtOption] = {}
        self.source_path: Path | None = None

    def load_from_workbook(self, path: Path) -> int:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            worksheet = workbook["Genarten"] if "Genarten" in workbook.sheetnames else workbook.active
            rows = worksheet.iter_rows(values_only=True)
            header = next(rows, None)
            if not header:
                raise ValueError("Die GenArt-Datei ist leer.")
            header_map = {str(value).strip(): index for index, value in enumerate(header) if value is not None}
            required_headers = {"ID", "Bezeichnung", "Genart"}
            if not required_headers.issubset(header_map):
                raise ValueError("Die GenArt-Datei braucht die Spalten ID, Bezeichnung und Genart.")

            options: list[GenArtOption] = []
            for row in rows:
                if row is None:
                    continue
                option_id = str(row[header_map["ID"]]).strip() if header_map["ID"] < len(row) and row[header_map["ID"]] is not None else ""
                bezeichnung = (
                    str(row[header_map["Bezeichnung"]]).strip()
                    if header_map["Bezeichnung"] < len(row) and row[header_map["Bezeichnung"]] is not None
                    else ""
                )
                genart = (
                    str(row[header_map["Genart"]]).strip()
                    if header_map["Genart"] < len(row) and row[header_map["Genart"]] is not None
                    else ""
                )
                if not option_id or not (bezeichnung or genart):
                    continue

                normalized_bezeichnung = normalize_match_text(bezeichnung)
                normalized_genart = normalize_match_text(genart)
                tokens = build_match_tokens(f"{bezeichnung} {genart}")
                families = infer_match_families_from_tokens(tokens)
                option = GenArtOption(
                    id=option_id,
                    bezeichnung=bezeichnung,
                    genart=genart,
                    normalized_bezeichnung=normalized_bezeichnung,
                    normalized_genart=normalized_genart,
                    tokens=tokens,
                    families=families,
                )
                options.append(option)
        finally:
            workbook.close()

        self.options = options
        self.options_by_display = {option.display_label(): option for option in options}
        self.options_by_id = {option.id: option for option in options}
        self.source_path = path
        return len(options)

    def list_display_values(self) -> list[str]:
        return [option.display_label() for option in self.options]

    def search_options(self, query: str, limit: int = 200) -> list[tuple[GenArtOption, float]]:
        if not query.strip():
            return [(option, 0.0) for option in self.options[:limit]]

        raw_query = query.strip()
        raw_query_casefold = raw_query.casefold()
        normalized_query = normalize_match_text(raw_query)
        query_tokens = set(build_match_tokens(raw_query))
        query_families = infer_match_families_from_tokens(query_tokens)
        matches: list[tuple[GenArtOption, float]] = []

        for option in self.options:
            score = 0.0
            searchable_fields = [
                option.id.casefold(),
                option.normalized_genart,
                option.normalized_bezeichnung,
            ]
            searchable_text = " ".join(part for part in searchable_fields if part)

            if raw_query_casefold == option.id.casefold():
                score += 1200
            elif option.id.casefold().startswith(raw_query_casefold):
                score += 850
            elif raw_query_casefold in option.id.casefold():
                score += 560

            if searchable_text and raw_query_casefold == searchable_text:
                score += 900

            if normalized_query:
                if option.normalized_bezeichnung.startswith(normalized_query):
                    score += 520
                elif normalized_query in option.normalized_bezeichnung:
                    score += 340

                if option.normalized_genart.startswith(normalized_query):
                    score += 500
                elif normalized_query in option.normalized_genart:
                    score += 320

                if normalized_query in searchable_text:
                    score += 220

                overlap = query_tokens & option.tokens
                if overlap:
                    score += len(overlap) * 46
                    score += (len(overlap) / max(1, len(query_tokens))) * 60

                family_overlap = query_families & option.families
                if family_overlap:
                    score += 180 * len(family_overlap)

                if option.normalized_bezeichnung:
                    score += SequenceMatcher(None, normalized_query, option.normalized_bezeichnung).ratio() * 55
                if option.normalized_genart:
                    score += SequenceMatcher(None, normalized_query, option.normalized_genart).ratio() * 50

            matched_terms = 0
            term_prefix_matches = 0
            for token in query_tokens:
                if token in searchable_text:
                    matched_terms += 1
                if any(field.startswith(token) for field in searchable_fields if field):
                    term_prefix_matches += 1

            if matched_terms:
                coverage = matched_terms / max(1, len(query_tokens))
                score += matched_terms * 85
                score += coverage * 260
                if len(query_tokens) > 1 and matched_terms == len(query_tokens):
                    score += 420
            if term_prefix_matches:
                score += term_prefix_matches * 70

            display_label = option.display_label().casefold()
            if raw_query_casefold and raw_query_casefold in display_label:
                score += 80

            if score > 0:
                matches.append((option, score))

        matches.sort(key=lambda item: (-item[1], item[0].display_label()))
        return matches[:limit]

    def search_display_values(self, query: str, limit: int = 200) -> list[str]:
        return [option.display_label() for option, _score in self.search_options(query, limit=limit)]

    def resolve(self, value: str) -> GenArtOption | None:
        query = value.strip()
        if not query:
            return None
        if query in self.options_by_display:
            return self.options_by_display[query]
        if query in self.options_by_id:
            return self.options_by_id[query]
        normalized_query = normalize_match_text(query)
        for option in self.options:
            if normalized_query in {option.normalized_bezeichnung, option.normalized_genart}:
                return option
        return None

    def suggest(
        self,
        short_text: str,
        long_text: str,
        image_context: str = "",
        category_context: str = "",
        web_context: str = "",
        limit: int = 5,
    ) -> list[tuple[GenArtOption, float]]:
        if not self.options:
            return []

        short_norm = normalize_match_text(short_text)
        long_norm = normalize_match_text(long_text)
        image_norm = normalize_match_text(image_context)
        category_norm = normalize_match_text(category_context)
        web_norm = normalize_match_text(web_context)
        text_query = " ".join(part for part in [short_norm, long_norm] if part).strip()
        query_text = " ".join(part for part in [short_norm, long_norm, image_norm, category_norm, web_norm] if part).strip()
        if not query_text:
            return []
        query_tokens = set(build_match_tokens(query_text))
        image_tokens = set(build_match_tokens(image_context))
        category_tokens = set(build_match_tokens(category_context))
        web_tokens = set(build_match_tokens(web_context))
        query_families = infer_match_families_from_tokens(query_tokens)
        suggestions: list[tuple[GenArtOption, float]] = []

        for option in self.options:
            score = 0.0
            if option.normalized_bezeichnung and option.normalized_bezeichnung in text_query:
                score += 220
            if option.normalized_genart and option.normalized_genart in text_query:
                score += 180
            if option.normalized_genart and any(option.normalized_genart == token or option.normalized_genart in token for token in query_tokens):
                score += 90
            if option.normalized_bezeichnung and any(
                option.normalized_bezeichnung == token or option.normalized_bezeichnung in token for token in query_tokens
            ):
                score += 115
            if short_norm and option.normalized_bezeichnung and short_norm in option.normalized_bezeichnung:
                score += 120
            if short_norm and option.normalized_genart and short_norm in option.normalized_genart:
                score += 95

            overlap = query_tokens & option.tokens
            if overlap:
                score += len(overlap) * 24
                score += (len(overlap) / max(1, len(option.tokens))) * 40

            family_overlap = query_families & option.families
            if family_overlap:
                score += 260 * len(family_overlap)
            elif query_families and option.families:
                for query_family in query_families:
                    for option_family in option.families:
                        score -= GENART_OPPOSING_FAMILY_PENALTIES.get((query_family, option_family), 0)

            extra_tokens = option.tokens - query_tokens
            if extra_tokens:
                score -= min(len(extra_tokens) * 14, 84)

            if image_norm:
                if option.normalized_bezeichnung and option.normalized_bezeichnung in image_norm:
                    score += 185
                if option.normalized_genart and option.normalized_genart in image_norm:
                    score += 165
                image_overlap = image_tokens & option.tokens
                if image_overlap:
                    score += len(image_overlap) * 38
                    score += (len(image_overlap) / max(1, len(option.tokens))) * 65

            if category_norm:
                if option.normalized_bezeichnung and option.normalized_bezeichnung in category_norm:
                    score += 160
                if option.normalized_genart and option.normalized_genart in category_norm:
                    score += 140
                category_overlap = category_tokens & option.tokens
                if category_overlap:
                    score += len(category_overlap) * 44
                    score += (len(category_overlap) / max(1, len(option.tokens))) * 80

            if web_norm:
                if option.normalized_bezeichnung and option.normalized_bezeichnung in web_norm:
                    score += 220
                if option.normalized_genart and option.normalized_genart in web_norm:
                    score += 200
                web_overlap = web_tokens & option.tokens
                if web_overlap:
                    score += len(web_overlap) * 58
                    score += (len(web_overlap) / max(1, len(option.tokens))) * 110

            if score > 0 or short_norm or image_norm:
                if option.normalized_bezeichnung:
                    score += SequenceMatcher(None, short_norm, option.normalized_bezeichnung).ratio() * 55
                if option.normalized_genart:
                    score += SequenceMatcher(None, short_norm, option.normalized_genart).ratio() * 45
                if image_norm and option.normalized_bezeichnung:
                    score += SequenceMatcher(None, image_norm, option.normalized_bezeichnung).ratio() * 70
                if image_norm and option.normalized_genart:
                    score += SequenceMatcher(None, image_norm, option.normalized_genart).ratio() * 60
                if category_norm and option.normalized_bezeichnung:
                    score += SequenceMatcher(None, category_norm, option.normalized_bezeichnung).ratio() * 75
                if category_norm and option.normalized_genart:
                    score += SequenceMatcher(None, category_norm, option.normalized_genart).ratio() * 65
                if web_norm and option.normalized_bezeichnung:
                    score += SequenceMatcher(None, web_norm, option.normalized_bezeichnung).ratio() * 110
                if web_norm and option.normalized_genart:
                    score += SequenceMatcher(None, web_norm, option.normalized_genart).ratio() * 95

            if score >= 70:
                suggestions.append((option, score))

        suggestions.sort(key=lambda item: (-item[1], item[0].display_label()))
        return suggestions[:limit]


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


class GoogleLensScrapeError(RuntimeError):
    pass


class GoogleLensScraper:
    def __init__(self, timeout_seconds: int = 60) -> None:
        self.timeout_seconds = timeout_seconds

    def detect_web(self, path_or_link: str) -> GoogleLensWebResult:
        browser_path = detect_browser_path()
        if browser_path is None:
            raise GoogleLensScrapeError("Kein kompatibler Browser fuer Google Lens gefunden.")

        local_upload_path = prepare_local_image_for_lens(path_or_link)
        target_url = ""
        is_remote = urlparse(path_or_link.strip()).scheme in {"http", "https"}
        if is_remote:
            target_url = build_google_lens_search_url(path_or_link.strip())

        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(executable_path=str(browser_path), headless=True)
                page = browser.new_page(viewport={"width": 1440, "height": 2200})
                try:
                    if is_remote:
                        page.goto(target_url, wait_until="domcontentloaded", timeout=self.timeout_seconds * 1000)
                    else:
                        page.goto("https://lens.google.com/", wait_until="domcontentloaded", timeout=self.timeout_seconds * 1000)
                        self._accept_google_cookies_if_present(page)
                        self._upload_local_image(page, local_upload_path)

                    self._accept_google_cookies_if_present(page)
                    page.wait_for_timeout(1800)
                    page.wait_for_load_state("networkidle", timeout=self.timeout_seconds * 1000)
                    page.wait_for_timeout(1800)
                    return self._extract_lens_result(page)
                finally:
                    browser.close()
        except PlaywrightTimeoutError as exc:
            raise GoogleLensScrapeError("Google Lens hat nicht rechtzeitig geantwortet.") from exc
        except Exception as exc:
            raise GoogleLensScrapeError(f"Google Lens konnte nicht gelesen werden: {exc}") from exc
        finally:
            if local_upload_path is not None and local_upload_path.exists() and local_upload_path.name.startswith("apollo_lens_"):
                try:
                    local_upload_path.unlink()
                except OSError:
                    pass

    def _accept_google_cookies_if_present(self, page: object) -> None:
        for label in ["Alle akzeptieren", "Accept all", "Ich stimme zu", "I agree"]:
            try:
                page.get_by_role("button", name=label).click(timeout=2500)
                page.wait_for_timeout(600)
                return
            except Exception:
                continue

    def _upload_local_image(self, page: object, upload_path: Path | None) -> None:
        if upload_path is None:
            raise GoogleLensScrapeError("Lokales Bild konnte nicht fuer Google Lens vorbereitet werden.")

        selectors = [
            "input[type='file']",
            "input[accept*='image']",
        ]
        for selector in selectors:
            try:
                locator = page.locator(selector).first
                if locator.count() > 0:
                    locator.set_input_files(str(upload_path))
                    return
            except Exception:
                continue
        raise GoogleLensScrapeError("Google Lens Upload-Feld wurde nicht gefunden.")

    def _extract_lens_result(self, page: object) -> GoogleLensWebResult:
        script = """
        () => {
            const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
            const body = document.body ? clean(document.body.innerText) : '';
            const lines = body.split('\\n').map(clean).filter(Boolean);
            const generic = new Set([
              'Alle', 'Shopping', 'Visuelle Treffer', 'Genaue Treffer', 'Ergebnisse',
              'Mehr', 'Google Apps', 'Anmelden', 'Bilder', 'Videos', 'Maps',
              'About this image', 'Zu diesem Bild'
            ]);

            const headlineLines = [];
            for (const line of lines) {
              if (line.length < 3 || line.length > 90 || generic.has(line)) continue;
              if (!headlineLines.includes(line)) headlineLines.push(line);
              if (headlineLines.length >= 12) break;
            }

            const titleCandidates = [];
            const snippetCandidates = [];
            const urls = [];
            const seenTitles = new Set();
            const seenUrls = new Set();

            for (const anchor of document.querySelectorAll('a[href]')) {
              const href = clean(anchor.href);
              const text = clean(anchor.textContent);
              if (!href || href.startsWith('javascript:') || href.startsWith('https://accounts.google.com')) continue;
              if (href.includes('/preferences?') || href.includes('/advanced_search')) continue;
              if (text && text.length >= 3 && text.length <= 120 && !seenTitles.has(text)) {
                titleCandidates.push(text);
                seenTitles.add(text);
              }
              const parentText = clean(anchor.parentElement ? anchor.parentElement.innerText : '');
              if (parentText && parentText.length > text.length && parentText.length <= 220 && !snippetCandidates.includes(parentText)) {
                snippetCandidates.push(parentText);
              }
              if (!seenUrls.has(href)) {
                urls.push(href);
                seenUrls.add(href);
              }
              if (titleCandidates.length >= 20 && urls.length >= 20) break;
            }

            return {
              headlineLines,
              resultTitles: titleCandidates.slice(0, 20),
              resultSnippets: snippetCandidates.slice(0, 16),
              pageUrls: urls.slice(0, 20),
            };
        }
        """
        raw = page.evaluate(script)
        return GoogleLensWebResult(
            headline_lines=[str(item).strip() for item in raw.get("headlineLines", []) if str(item).strip()],
            result_titles=[str(item).strip() for item in raw.get("resultTitles", []) if str(item).strip()],
            result_snippets=[str(item).strip() for item in raw.get("resultSnippets", []) if str(item).strip()],
            page_urls=[str(item).strip() for item in raw.get("pageUrls", []) if str(item).strip()],
        )


class KunzerScrapeError(RuntimeError):
    pass


@dataclass
class KunzerScrapeResult:
    article_number: str
    product_url: str
    title: str
    breadcrumb_text: str
    short_text_de: str
    long_text_de: str
    image_links: list[str]
    document_links: list[str]
    video_links: list[str]


class KunzerScraper:
    def __init__(self) -> None:
        self._sitemap_index: dict[str, str] | None = None
        self._result_cache: dict[str, KunzerScrapeResult] = {}

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
        cached = self._result_cache.get(target_url)
        if cached is not None:
            return self._clone_result(cached)
        browser_path = self._detect_browser_path()

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(browser_path) if browser_path else None,
                headless=True,
            )
            page = browser.new_page(viewport={"width": 1440, "height": 2400})
            try:
                page.goto(target_url, wait_until="domcontentloaded", timeout=90000)
                self._accept_cookies_if_present(page)
                try:
                    page.get_by_role("heading").first.wait_for(timeout=12000)
                except Exception:
                    pass
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                self._expand_all_documents(page)
                page.wait_for_timeout(250)

                title = self._clean_single_line(page.title().replace("| Kunzer", ""))
                heading = self._safe_first_text(page.get_by_role("heading").all_inner_texts())
                short_text_de = self._clean_single_line(heading or title)
                breadcrumb_text = self._extract_breadcrumb_text(page)

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

                result = KunzerScrapeResult(
                    article_number=article_number,
                    product_url=page.url,
                    title=short_text_de,
                    breadcrumb_text=breadcrumb_text,
                    short_text_de=short_text_de,
                    long_text_de=long_text_de,
                    image_links=image_links,
                    document_links=document_links,
                    video_links=video_links,
                )
                self._result_cache[target_url] = result
                self._result_cache[result.product_url] = result
                return self._clone_result(result)
            except PlaywrightTimeoutError as exc:
                raise KunzerScrapeError(f"Timeout beim Laden der Kunzer-Seite: {target_url}") from exc
            except Exception as exc:  # pragma: no cover - runtime GUI feedback
                raise KunzerScrapeError(f"Kunzer-Seite konnte nicht gelesen werden: {exc}") from exc
            finally:
                browser.close()

    def _clone_result(self, result: KunzerScrapeResult) -> KunzerScrapeResult:
        return KunzerScrapeResult(
            article_number=result.article_number,
            product_url=result.product_url,
            title=result.title,
            breadcrumb_text=result.breadcrumb_text,
            short_text_de=result.short_text_de,
            long_text_de=result.long_text_de,
            image_links=list(result.image_links),
            document_links=list(result.document_links),
            video_links=list(result.video_links),
        )

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

    def _extract_breadcrumb_text(self, page: object) -> str:
        script = """
        () => {
            const clean = (value) => (value || "").replace(/\\s+/g, " ").trim();
            const selectors = [
                '[aria-label*="breadcrumb" i]',
                '[class*="breadcrumb" i]',
                '[data-testid*="breadcrumb" i]',
                'nav',
            ];
            const candidates = [];
            const seen = new Set();

            for (const selector of selectors) {
                for (const node of document.querySelectorAll(selector)) {
                    const parts = [];
                    const descendants = node.querySelectorAll('a, span, li, p, div');
                    for (const descendant of descendants) {
                        const text = clean(descendant.textContent);
                        if (!text || text.length > 50) {
                            continue;
                        }
                        if (parts.length && parts[parts.length - 1] === text) {
                            continue;
                        }
                        parts.push(text);
                    }

                    if (parts.length < 2 || parts.length > 6) {
                        continue;
                    }

                    const text = parts.join(' > ');
                    if (!text || text.length > 140 || seen.has(text)) {
                        continue;
                    }
                    seen.add(text);

                    let score = 0;
                    if (/produkte/i.test(text)) score += 120;
                    if (/breadcrumb/i.test(selector)) score += 80;
                    if (/shop|downloads|kontakt|ueber uns/i.test(text)) score -= 60;
                    score += Math.max(0, 40 - parts.length * 6);
                    score += Math.max(0, 50 - text.length / 3);
                    candidates.push({ text, score });
                }
            }

            candidates.sort((left, right) => right.score - left.score);
            return candidates.length ? candidates[0].text : '';
        }
        """
        try:
            breadcrumb_text = str(page.evaluate(script)).strip()
        except Exception:
            return ""
        return self._clean_single_line(breadcrumb_text)

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
    short_texts = sanitize_short_translation_set(bundle.short_texts)
    return [
        bundle.article_number,
        bundle.short_module_id,
        "1",
        "1",
        *short_texts.export_values(bundle.short_auto_uni),
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


def build_genart_export_rows(bundle: ExportBundle) -> list[list[str]]:
    genart_id = bundle.genart_id.strip()
    genart_bezeichnung = bundle.genart_bezeichnung.strip()
    if not genart_id and not genart_bezeichnung:
        return []
    return [[bundle.article_number, genart_id, genart_bezeichnung]]


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
        normalized_video_link = normalize_youtube_url_for_embed(row.path_or_link)
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
    genart_rows = [append_written_at(row, written_at) for row in build_genart_export_rows(bundle)]
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
        if genart_rows:
            write_workbook(export_dir / GENART_FILE[0], GENART_FILE[1], GENART_HEADERS, genart_rows)
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
        write_workbook_with_upsert(
            export_dir / GENART_FILE[0],
            GENART_FILE[1],
            GENART_HEADERS,
            genart_rows,
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
        f"- GenArt: {bundle.genart_id or '(nicht gesetzt)'} {bundle.genart_bezeichnung}".rstrip(),
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
        f"- {GENART_FILE[0]}",
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
        f"GenArt: {(snapshot.genart_id + ' | ' + snapshot.genart_bezeichnung).strip(' |') or '-'}",
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

    genart_rows = read_workbook_rows(folder / GENART_FILE[0], GENART_FILE[1], len(GENART_HEADERS))
    for row in genart_rows:
        article_number = normalize_article_number(row[0])
        if not article_number:
            continue
        snapshot = ensure_snapshot(article_number)
        snapshot.genart_id = row[1].strip()
        snapshot.genart_bezeichnung = row[2].strip()

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
    def __init__(
        self,
        master: tk.Misc,
        title: str,
        on_change: Callable[[], None] | None = None,
        max_length: int | None = None,
    ) -> None:
        super().__init__(master, text=title, padding=14)
        self.on_change = on_change
        self.max_length = max_length
        self._suspend_limit_enforcement = False
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
        auto_uni_check.grid(row=0, column=0, sticky="w", pady=(0, 10))

        if self.max_length is not None:
            ttk.Label(
                self,
                text=f"Maximal {self.max_length} Zeichen. Umlaute werden automatisch ersetzt.",
                foreground="#5E6472",
            ).grid(row=0, column=1, sticky="e", pady=(0, 10))

        for row_index, (code, label) in enumerate(UI_LANGUAGE_ORDER, start=1):
            ttk.Label(self, text=label).grid(row=row_index, column=0, sticky="w", padx=(0, 10), pady=4)
            entry_state = "readonly" if code == "uni" else "normal"
            entry = ttk.Entry(self, textvariable=self.fields[code], width=90, state=entry_state)
            entry.grid(row=row_index, column=1, sticky="ew", pady=4)
            entry.bind("<FocusOut>", self._handle_focus_out)
            self.fields[code].trace_add("write", lambda *_args, code=code: self._sanitize_field_value(code))
            if code == "uni":
                self.uni_entry = entry
            if code == "de":
                self.fields["de"].trace_add("write", self._sync_uni)

        self._toggle_uni_state()

    def _sync_uni(self, *_args: object) -> None:
        if self.auto_uni_var.get():
            self.fields["uni"].set(self.fields["de"].get())

    def _sanitize_short_value(self, value: str) -> str:
        sanitized = replace_short_text_umlauts(value)
        if self.max_length is not None:
            sanitized = sanitized[: self.max_length]
        return sanitized

    def _sanitize_field_value(self, code: str) -> None:
        if self._suspend_limit_enforcement:
            return
        value = self.fields[code].get()
        sanitized = self._sanitize_short_value(value)
        if sanitized == value:
            return
        self._suspend_limit_enforcement = True
        try:
            self.fields[code].set(sanitized)
        finally:
            self._suspend_limit_enforcement = False

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
        values = {}
        for code, variable in self.fields.items():
            values[code] = self._sanitize_short_value(variable.get())
        return TranslationSet(**values)

    def get_german_text(self) -> str:
        return self.fields["de"].get().strip()

    def apply_translations(self, translations: dict[str, str]) -> None:
        for code, value in translations.items():
            if code in self.fields and code != "de":
                self.fields[code].set(self._sanitize_short_value(value))
        self._toggle_uni_state()

    def set_value(self, value: TranslationSet, auto_uni: bool | None = None) -> None:
        if auto_uni is not None:
            self.auto_uni_var.set(auto_uni)
        for code, variable in self.fields.items():
            variable.set(self._sanitize_short_value(getattr(value, code)))
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
        self.preview_bytes_cache: dict[str, bytes] = {}
        self.pdf_preview_cache: dict[str, bytes] = {}
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
        data = self._load_preview_bytes(path_or_link)
        return Image.open(io.BytesIO(data))

    def _load_preview_bytes(self, path_or_link: str) -> bytes:
        cache_key = path_or_link.strip()
        cached = self.preview_bytes_cache.get(cache_key)
        if cached is not None:
            return cached

        parsed = urlparse(path_or_link)
        if parsed.scheme in {"http", "https"}:
            request = urllib_request.Request(path_or_link, headers={"User-Agent": "ApolloImportGui/1.0"})
            with urllib_request.urlopen(request, timeout=12) as response:
                data = response.read()
                self.preview_bytes_cache[cache_key] = data
                return data

        file_path = Path(path_or_link)
        if not file_path.exists():
            raise FileNotFoundError("Datei nicht gefunden")
        data = file_path.read_bytes()
        self.preview_bytes_cache[cache_key] = data
        return data

    def _render_pdf_preview(self, path_or_link: str) -> object:
        cache_key = path_or_link.strip()
        cached = self.pdf_preview_cache.get(cache_key)
        if cached is not None:
            return Image.open(io.BytesIO(cached))

        pdf_bytes = self._load_preview_bytes(path_or_link)
        document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        try:
            page = document[0]
            matrix = pymupdf.Matrix(1.5, 1.5)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image = pixmap.pil_image()
        finally:
            document.close()

        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        self.pdf_preview_cache[cache_key] = buffer.getvalue()
        return image


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
        self.remote_image_cache: dict[str, bytes] = {}
        self.web_preview_cache: dict[str, tuple[bytes, str]] = {}
        self.page_title_cache: dict[str, str] = {}
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
        cached = self.remote_image_cache.get(url)
        if cached is None:
            request = urllib_request.Request(url, headers={"User-Agent": "ApolloImportGui/1.0"})
            with urllib_request.urlopen(request, timeout=10) as response:
                cached = response.read()
            self.remote_image_cache[url] = cached
        return Image.open(io.BytesIO(cached))

    def _render_web_preview(self, link: str) -> tuple[object, str]:
        parsed = urlparse(link)
        if parsed.scheme not in {"http", "https"}:
            raise ValueError("Nur http/https Links koennen als Webseite vorgeladen werden.")

        cached = self.web_preview_cache.get(link)
        if cached is not None:
            screenshot_bytes, title = cached
            return Image.open(io.BytesIO(screenshot_bytes)), title

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

        self.web_preview_cache[link] = (screenshot, title)
        return Image.open(io.BytesIO(screenshot)), title

    def _fetch_page_title(self, link: str) -> str:
        parsed = urlparse(link)
        if parsed.scheme not in {"http", "https"}:
            return ""

        cached = self.page_title_cache.get(link)
        if cached is not None:
            return cached

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
        title = " ".join(unescape(match.group(1)).split()).strip()
        self.page_title_cache[link] = title
        return title

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


class GenArtSearchDialog:
    def __init__(self, master: tk.Misc, registry: GenArtRegistry, initial_query: str = "") -> None:
        self.master = master
        self.registry = registry
        self.result: GenArtOption | None = None
        self.search_var = tk.StringVar(value=initial_query.strip())
        self.results_var = tk.StringVar(value="")

        self.window = tk.Toplevel(master)
        self.window.title("GenArt suchen")
        self.window.transient(master.winfo_toplevel())
        self.window.geometry("920x620")
        self.window.minsize(760, 460)
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)

        header = ttk.Frame(self.window, padding=14)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        ttk.Label(header, text="Suche", font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.search_entry = ttk.Entry(header, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(header, text="Auswaehlen", command=self._accept).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(header, text="Abbrechen", command=self._cancel).grid(row=0, column=3, padx=(8, 0))

        ttk.Label(
            header,
            text="Suche in ID, GenArt und Bezeichnung. Mehrere Begriffe sind moeglich.",
            foreground="#5E6472",
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Label(header, textvariable=self.results_var, foreground="#5E6472").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(6, 0)
        )

        table_frame = ttk.Frame(self.window, padding=(14, 0, 14, 14))
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("id", "genart", "bezeichnung", "score")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.heading("id", text="ID")
        self.tree.heading("genart", text="GenArt")
        self.tree.heading("bezeichnung", text="Bezeichnung")
        self.tree.heading("score", text="Treffer")
        self.tree.column("id", width=110, anchor="w")
        self.tree.column("genart", width=220, anchor="w")
        self.tree.column("bezeichnung", width=420, anchor="w")
        self.tree.column("score", width=80, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.search_var.trace_add("write", self._refresh_results)
        self.search_entry.bind("<Down>", self._focus_results)
        self.search_entry.bind("<Return>", self._accept_first_result)
        self.tree.bind("<Double-1>", lambda _event: self._accept())
        self.tree.bind("<Return>", lambda _event: self._accept())
        self.tree.bind("<Escape>", lambda _event: self._cancel())

        self._refresh_results()
        self.window.grab_set()
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")

    def show(self) -> GenArtOption | None:
        self.window.wait_window()
        return self.result

    def _refresh_results(self, *_args) -> None:
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        query = self.search_var.get().strip()
        results = self.registry.search_options(query, limit=300)
        if not query:
            results = [(option, 0.0) for option in self.registry.options[:300]]

        for option, score in results:
            self.tree.insert(
                "",
                "end",
                iid=option.id,
                values=(
                    option.id,
                    option.genart,
                    option.bezeichnung,
                    "" if score <= 0 else f"{score:.0f}",
                ),
            )

        shown_count = len(results)
        total_count = len(self.registry.options)
        if query:
            self.results_var.set(f"{shown_count} Treffer angezeigt, {total_count} GenArts insgesamt.")
        else:
            self.results_var.set(f"{shown_count} GenArts angezeigt. Tippe oben, um die Liste zu filtern.")

        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])

    def _focus_results(self, _event: tk.Event[tk.Misc]) -> str:
        children = self.tree.get_children()
        if children:
            self.tree.focus(children[0])
            self.tree.selection_set(children[0])
            self.tree.focus_set()
        return "break"

    def _accept_first_result(self, _event: tk.Event[tk.Misc]) -> str:
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
            self._accept()
        return "break"

    def _accept(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self.result = self.registry.resolve(selection[0])
        self.window.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.window.destroy()


class ApolloImportApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1360x920")
        self.root.minsize(1180, 780)
        self.app_icon_photo: object | None = None

        self.id_registry = IdRegistry()
        self.genart_registry = GenArtRegistry()
        self.kunzer_scraper = KunzerScraper()

        self.import_dir_var = tk.StringVar(value=str(DEFAULT_IMPORT_DIR))
        self.output_dir_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.genart_source_path_var = tk.StringVar(value=str(DEFAULT_GENART_SOURCE) if DEFAULT_GENART_SOURCE.exists() else "")
        self.product_list_path_var = tk.StringVar()
        self.deepl_api_key_var = tk.StringVar(value=os.getenv("DEEPL_API_KEY", ""))
        self.deepl_base_url_var = tk.StringVar(value=os.getenv("DEEPL_API_BASE_URL", DEEPL_DEFAULT_BASE_URL))
        self.kunzer_product_url_var = tk.StringVar()
        self.genart_display_var = tk.StringVar()
        self.genart_suggestion_var = tk.StringVar(value="Keine GenArt geladen.")
        self.auto_translate_after_scrape_var = tk.BooleanVar(value=True)
        self.google_lens_enabled_var = tk.BooleanVar(value=True)
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
        self.genart_count_var = tk.StringVar(value="0 GenArts geladen")
        self.current_id_article_number = ""
        self.article_browser_records: dict[str, StoredArticleSnapshot] = {}
        self.current_kunzer_category_context = ""
        self.google_lens_web_cache: dict[str, GoogleLensWebResult | None] = {}
        self.genart_image_signature_cache: dict[str, ImageSignature | None] = {}
        self.genart_image_reference_index: list[tuple[str, str, ImageSignature]] = []
        self.genart_image_index_dirty = True
        self.article_browser_cache_dir = ""
        self.article_browser_cache_signature: tuple[tuple[str, int, int], ...] | None = None
        self.live_write_suspended = False
        self.article_browser_context_menu: tk.Menu | None = None
        self.background_task_running = False
        self.pending_article_browser_selection = ""

        self._configure_style()
        self._configure_window_icon()
        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)
        self._restore_session_state()
        self._load_known_ids(initial=True)
        self._load_genart_catalog(initial=True)
        self.refresh_preview()
        self._apply_pending_session_selection()

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

    def _configure_window_icon(self) -> None:
        ico_path = resolve_application_asset_path(APP_ICON_ICO_RELATIVE_PATH)
        png_path = resolve_application_asset_path(APP_ICON_PNG_RELATIVE_PATH)

        if ico_path.exists():
            try:
                self.root.iconbitmap(default=str(ico_path))
            except tk.TclError:
                pass

        if png_path.exists():
            try:
                self.app_icon_photo = tk.PhotoImage(file=str(png_path))
                self.root.iconphoto(True, self.app_icon_photo)
            except tk.TclError:
                self.app_icon_photo = None

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

        self.main_notebook = ttk.Notebook(shell)
        self.main_notebook.grid(row=1, column=0, sticky="nsew")

        self.project_tab = ttk.Frame(self.main_notebook, padding=18)
        self.short_tab = ttk.Frame(self.main_notebook, padding=18)
        self.long_tab = ttk.Frame(self.main_notebook, padding=18)
        self.image_tab = ttk.Frame(self.main_notebook, padding=18)
        self.document_tab = ttk.Frame(self.main_notebook, padding=18)
        self.links_tab = ttk.Frame(self.main_notebook, padding=18)

        self.main_notebook.add(self.project_tab, text="Projekt")
        self.main_notebook.add(self.short_tab, text="Kurzbezeichnung")
        self.main_notebook.add(self.long_tab, text="Text")
        self.main_notebook.add(self.image_tab, text="Bilder")
        self.main_notebook.add(self.document_tab, text="Dokumente")
        self.main_notebook.add(self.links_tab, text="Links")

        self._build_project_tab()
        self._build_short_tab()
        self._build_long_tab()
        self._build_media_tabs()

        status_bar = ttk.Label(shell, textvariable=self.status_var, anchor="w", foreground="#5E6472")
        status_bar.grid(row=2, column=0, sticky="ew", pady=(12, 0))

    def _set_background_task_state(self, running: bool, status_message: str | None = None) -> None:
        self.background_task_running = running
        try:
            self.root.configure(cursor="watch" if running else "")
        except tk.TclError:
            pass
        if status_message:
            self.status_var.set(status_message)

    def _run_background_task(
        self,
        start_message: str,
        worker: Callable[[], object],
        on_success: Callable[[object], None],
        error_status_prefix: str,
    ) -> bool:
        if self.background_task_running:
            self.status_var.set("Es laeuft bereits ein Hintergrundvorgang. Bitte kurz warten.")
            return False

        self._set_background_task_state(True, start_message)

        def finish_success(result: object) -> None:
            self._set_background_task_state(False)
            try:
                on_success(result)
            except Exception as exc:  # pragma: no cover - defensive GUI feedback
                messagebox.showwarning(APP_TITLE, str(exc))
                self.status_var.set(f"{error_status_prefix}: {exc}")

        def finish_error(exc: Exception) -> None:
            self._set_background_task_state(False)
            messagebox.showwarning(APP_TITLE, str(exc))
            self.status_var.set(f"{error_status_prefix}: {exc}")

        def task() -> None:
            try:
                result = worker()
            except Exception as exc:  # pragma: no cover - depends on user/network/runtime
                self.root.after(0, lambda exc=exc: finish_error(exc))
                return
            self.root.after(0, lambda result=result: finish_success(result))

        threading.Thread(target=task, daemon=True).start()
        return True

    def _translation_set_to_state(self, translations: TranslationSet) -> dict[str, str]:
        return {
            code: getattr(translations, code)
            for code, _label in UI_LANGUAGE_ORDER
        }

    def _translation_set_from_state(self, payload: object) -> TranslationSet:
        if not isinstance(payload, dict):
            return TranslationSet()
        values = {
            code: str(payload.get(code, ""))
            for code, _label in UI_LANGUAGE_ORDER
        }
        return TranslationSet(**values)

    def _media_rows_to_state(self, rows: list[MediaRow]) -> list[dict[str, str]]:
        return [
            {
                "path_or_link": row.path_or_link,
                "art": row.art,
                "sprache": row.sprache,
            }
            for row in rows
        ]

    def _media_rows_from_state(self, payload: object) -> list[MediaRow]:
        if not isinstance(payload, list):
            return []
        rows: list[MediaRow] = []
        for item in payload:
            if not isinstance(item, dict):
                continue
            rows.append(
                MediaRow(
                    path_or_link=str(item.get("path_or_link", "")),
                    art=str(item.get("art", "")),
                    sprache=str(item.get("sprache", "")),
                )
            )
        return rows

    def _collect_session_state(self) -> dict[str, object]:
        selected_tab = 0
        try:
            selected_tab = self.main_notebook.index(self.main_notebook.select())
        except Exception:
            selected_tab = 0

        selected_browser_article = ""
        if hasattr(self, "article_browser_tree") and self.article_browser_tree.selection():
            selected_browser_article = str(self.article_browser_tree.selection()[0]).strip()

        return {
            "version": 1,
            "window_geometry": self.root.winfo_geometry(),
            "selected_tab": selected_tab,
            "paths": {
                "import_dir": self.import_dir_var.get(),
                "output_dir": self.output_dir_var.get(),
                "genart_source_path": self.genart_source_path_var.get(),
                "product_list_path": self.product_list_path_var.get(),
            },
            "api": {
                "deepl_api_key": self.deepl_api_key_var.get(),
                "deepl_base_url": self.deepl_base_url_var.get(),
            },
            "options": {
                "auto_translate_after_scrape": self.auto_translate_after_scrape_var.get(),
                "google_lens_enabled": self.google_lens_enabled_var.get(),
                "fixed_export_path": self.fixed_export_path_var.get(),
                "batch_short_text": self.batch_short_text_var.get(),
                "batch_long_text": self.batch_long_text_var.get(),
                "batch_image": self.batch_image_var.get(),
                "batch_document": self.batch_document_var.get(),
                "batch_video": self.batch_video_var.get(),
                "batch_web": self.batch_web_var.get(),
            },
            "article": {
                "article_number": self.article_number_var.get(),
                "short_module_id": self.short_module_id_var.get(),
                "long_module_id": self.long_module_id_var.get(),
                "kunzer_product_url": self.kunzer_product_url_var.get(),
                "genart_display": self.genart_display_var.get(),
                "genart_suggestion": self.genart_suggestion_var.get(),
                "current_kunzer_category_context": self.current_kunzer_category_context,
            },
            "translations": {
                "short_auto_uni": self.short_text_frame.auto_uni_var.get(),
                "short_values": self._translation_set_to_state(self.short_text_frame.get_value()),
                "long_auto_uni": self.long_text_frame.auto_uni_var.get(),
                "long_values": self._translation_set_to_state(self.long_text_frame.get_value()),
            },
            "media": {
                "images": {
                    "rows": self._media_rows_to_state(self.image_frame.get_rows()),
                    "form": {
                        "path_or_link": self.image_frame.path_var.get(),
                        "art": self.image_frame.art_var.get(),
                        "sprache": self.image_frame.sprache_var.get(),
                    },
                },
                "documents": {
                    "rows": self._media_rows_to_state(self.document_frame.get_rows()),
                    "form": {
                        "path_or_link": self.document_frame.path_var.get(),
                        "art": self.document_frame.art_var.get(),
                        "sprache": self.document_frame.sprache_var.get(),
                    },
                },
                "videos": {
                    "rows": self._media_rows_to_state(self.video_frame.get_rows()),
                    "form": {
                        "path_or_link": self.video_frame.link_var.get(),
                    },
                },
                "web_links": {
                    "rows": self._media_rows_to_state(self.web_frame.get_rows()),
                    "form": {
                        "path_or_link": self.web_frame.link_var.get(),
                    },
                },
            },
            "article_browser": {
                "selected_article": selected_browser_article,
            },
        }

    def _save_session_state(self) -> None:
        SESSION_STATE_FILE.write_text(
            json.dumps(self._collect_session_state(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _restore_session_state(self) -> None:
        if not SESSION_STATE_FILE.exists():
            return
        try:
            payload = json.loads(SESSION_STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        if not isinstance(payload, dict):
            return

        with self._suspend_live_write():
            geometry = payload.get("window_geometry")
            if isinstance(geometry, str) and geometry.strip():
                try:
                    self.root.geometry(geometry.strip())
                except tk.TclError:
                    pass

            paths = payload.get("paths")
            if isinstance(paths, dict):
                self.import_dir_var.set(str(paths.get("import_dir", self.import_dir_var.get())))
                self.output_dir_var.set(str(paths.get("output_dir", self.output_dir_var.get())))
                self.genart_source_path_var.set(str(paths.get("genart_source_path", self.genart_source_path_var.get())))
                self.product_list_path_var.set(str(paths.get("product_list_path", self.product_list_path_var.get())))

            api = payload.get("api")
            if isinstance(api, dict):
                self.deepl_api_key_var.set(str(api.get("deepl_api_key", self.deepl_api_key_var.get())))
                self.deepl_base_url_var.set(str(api.get("deepl_base_url", self.deepl_base_url_var.get())))

            options = payload.get("options")
            if isinstance(options, dict):
                self.auto_translate_after_scrape_var.set(bool(options.get("auto_translate_after_scrape", self.auto_translate_after_scrape_var.get())))
                self.google_lens_enabled_var.set(bool(options.get("google_lens_enabled", self.google_lens_enabled_var.get())))
                self.fixed_export_path_var.set(bool(options.get("fixed_export_path", self.fixed_export_path_var.get())))
                self.batch_short_text_var.set(bool(options.get("batch_short_text", self.batch_short_text_var.get())))
                self.batch_long_text_var.set(bool(options.get("batch_long_text", self.batch_long_text_var.get())))
                self.batch_image_var.set(bool(options.get("batch_image", self.batch_image_var.get())))
                self.batch_document_var.set(bool(options.get("batch_document", self.batch_document_var.get())))
                self.batch_video_var.set(bool(options.get("batch_video", self.batch_video_var.get())))
                self.batch_web_var.set(bool(options.get("batch_web", self.batch_web_var.get())))

            article = payload.get("article")
            if isinstance(article, dict):
                article_number = normalize_article_number(str(article.get("article_number", "")))
                self.article_number_var.set(article_number)
                self.short_module_id_var.set(str(article.get("short_module_id", "")))
                self.long_module_id_var.set(str(article.get("long_module_id", "")))
                self.kunzer_product_url_var.set(str(article.get("kunzer_product_url", "")))
                self.genart_display_var.set(str(article.get("genart_display", "")))
                self.genart_suggestion_var.set(str(article.get("genart_suggestion", self.genart_suggestion_var.get())))
                self.current_kunzer_category_context = str(article.get("current_kunzer_category_context", ""))
                self.current_id_article_number = article_number if article_number else ""

            translations = payload.get("translations")
            if isinstance(translations, dict):
                self.short_text_frame.set_value(
                    self._translation_set_from_state(translations.get("short_values")),
                    auto_uni=bool(translations.get("short_auto_uni", True)),
                )
                self.long_text_frame.set_value(
                    self._translation_set_from_state(translations.get("long_values")),
                    auto_uni=bool(translations.get("long_auto_uni", True)),
                )

            media = payload.get("media")
            if isinstance(media, dict):
                image_payload = media.get("images")
                if isinstance(image_payload, dict):
                    self.image_frame.set_rows(self._media_rows_from_state(image_payload.get("rows")))
                    form = image_payload.get("form")
                    if isinstance(form, dict):
                        self.image_frame.path_var.set(str(form.get("path_or_link", "")))
                        self.image_frame.art_var.set(str(form.get("art", self.image_frame.default_art)))
                        self.image_frame.sprache_var.set(str(form.get("sprache", self.image_frame.default_sprache)))

                document_payload = media.get("documents")
                if isinstance(document_payload, dict):
                    self.document_frame.set_rows(self._media_rows_from_state(document_payload.get("rows")))
                    form = document_payload.get("form")
                    if isinstance(form, dict):
                        self.document_frame.path_var.set(str(form.get("path_or_link", "")))
                        self.document_frame.art_var.set(str(form.get("art", self.document_frame.default_art)))
                        self.document_frame.sprache_var.set(str(form.get("sprache", self.document_frame.default_sprache)))

                video_payload = media.get("videos")
                if isinstance(video_payload, dict):
                    self.video_frame.set_rows(self._media_rows_from_state(video_payload.get("rows")))
                    form = video_payload.get("form")
                    if isinstance(form, dict):
                        self.video_frame.link_var.set(str(form.get("path_or_link", "")))

                web_payload = media.get("web_links")
                if isinstance(web_payload, dict):
                    self.web_frame.set_rows(self._media_rows_from_state(web_payload.get("rows")))
                    form = web_payload.get("form")
                    if isinstance(form, dict):
                        self.web_frame.link_var.set(str(form.get("path_or_link", "")))

            selected_tab = payload.get("selected_tab")
            if isinstance(selected_tab, int) and 0 <= selected_tab < len(self.main_notebook.tabs()):
                self.main_notebook.select(selected_tab)

            browser_state = payload.get("article_browser")
            if isinstance(browser_state, dict):
                self.pending_article_browser_selection = normalize_article_number(str(browser_state.get("selected_article", "")))

    def _apply_pending_session_selection(self) -> None:
        selected_article = normalize_article_number(self.pending_article_browser_selection)
        if not selected_article or selected_article not in self.article_browser_records:
            return
        self.article_browser_tree.selection_set(selected_article)
        self.article_browser_tree.focus(selected_article)
        self._update_article_browser_detail(self.article_browser_records[selected_article])

    def _handle_close(self) -> None:
        try:
            self._save_session_state()
        except Exception:
            pass
        self.root.destroy()

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

        ttk.Label(file_frame, text="GenArt XLSX").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(file_frame, textvariable=self.genart_source_path_var).grid(row=3, column=1, sticky="ew", pady=6)
        ttk.Button(file_frame, text="Datei waehlen", command=self.choose_genart_source_file).grid(row=3, column=2, padx=(8, 0), pady=6)
        ttk.Label(file_frame, textvariable=self.genart_count_var, foreground="#5E6472").grid(row=3, column=3, sticky="w", padx=(8, 0), pady=6)

        article_frame = ttk.LabelFrame(self.project_tab, text="Artikel", padding=14)
        article_frame.grid(row=1, column=0, sticky="new", pady=(14, 0), padx=(0, 9))
        article_frame.columnconfigure(1, weight=1)
        article_frame.columnconfigure(2, weight=0)

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

        ttk.Label(article_frame, text="GenArt").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=6)
        self.genart_combo = ttk.Combobox(article_frame, textvariable=self.genart_display_var)
        self.genart_combo.grid(row=3, column=1, sticky="ew", pady=6)
        self.genart_combo.bind("<<ComboboxSelected>>", lambda _event: self._write_live_section("genart"))
        self.genart_combo.bind("<KeyRelease>", self._on_genart_key_release)
        self.genart_combo.bind("<Return>", self._on_genart_return)
        self.genart_combo.bind("<FocusIn>", self._on_genart_focus_in)
        self.genart_combo.bind("<FocusOut>", self._on_genart_focus_out)
        self.genart_combo.bind("<Control-f>", self._open_genart_search_dialog_event)
        self.genart_combo.bind("<F4>", self._open_genart_search_dialog_event)
        ttk.Button(article_frame, text="Suchen...", command=self._open_genart_search_dialog).grid(
            row=3, column=2, sticky="w", padx=(8, 0), pady=6
        )

        genart_hint_row = ttk.Frame(article_frame)
        genart_hint_row.grid(row=4, column=0, columnspan=3, sticky="ew")
        genart_hint_row.columnconfigure(0, weight=1)
        ttk.Label(genart_hint_row, textvariable=self.genart_suggestion_var, foreground="#5E6472", wraplength=420).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(genart_hint_row, text="GenArt vorschlagen", command=self.suggest_genart_for_current_article).grid(
            row=0, column=1, padx=(8, 0)
        )

        kunzer_row = ttk.Frame(article_frame)
        kunzer_row.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(kunzer_row, text="Aus Kunzer per Artikelnummer laden", command=self.load_from_kunzer_article_number).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(kunzer_row, text="Aus Kunzer per URL laden", command=self.load_from_kunzer_url).grid(row=0, column=1)

        options_row = ttk.Frame(article_frame)
        options_row.grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Checkbutton(
            options_row,
            text="Nach dem Laden automatisch mit DeepL uebersetzen",
            variable=self.auto_translate_after_scrape_var,
        ).grid(row=0, column=0)

        button_row = ttk.Frame(article_frame)
        button_row.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))
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

        deepl_frame = ttk.LabelFrame(self.project_tab, text="APIs", padding=14)
        deepl_frame.grid(row=2, column=0, sticky="ew", pady=(14, 0), padx=(0, 9))
        deepl_frame.columnconfigure(1, weight=1)

        ttk.Label(deepl_frame, text="DeepL", font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w", pady=(0, 4))

        ttk.Label(deepl_frame, text="API Key").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(deepl_frame, textvariable=self.deepl_api_key_var, show="*", width=60).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(deepl_frame, text="Base URL").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(deepl_frame, textvariable=self.deepl_base_url_var).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Label(
            deepl_frame,
            text="Pro: https://api.deepl.com  |  Free: https://api-free.deepl.com",
            foreground="#5E6472",
        ).grid(row=2, column=2, sticky="w", padx=(10, 0), pady=6)

        deepl_actions = ttk.Frame(deepl_frame)
        deepl_actions.grid(row=3, column=1, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(deepl_actions, text="Kurzbezeichnung uebersetzen", command=self.translate_short_texts).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(deepl_actions, text="Text uebersetzen", command=self.translate_long_texts).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(deepl_actions, text="Alles uebersetzen", command=self.translate_all_texts).grid(row=0, column=2)

        ttk.Label(
            deepl_frame,
            text="Die Uebersetzung laeuft jeweils aus dem deutschen Feld in EN, CZ, FR, IT und NL. UNI bleibt an Deutsch gekoppelt.",
            foreground="#5E6472",
            wraplength=1100,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

        ttk.Separator(deepl_frame, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="ew", pady=(14, 12))

        ttk.Label(deepl_frame, text="Google Lens", font=("Segoe UI Semibold", 10)).grid(row=6, column=0, sticky="w", pady=(0, 4))
        ttk.Checkbutton(
            deepl_frame,
            text="Google Lens ohne API-Key fuer GenArt verwenden (inoffiziell)",
            variable=self.google_lens_enabled_var,
        ).grid(row=6, column=1, sticky="w", pady=(0, 4))

        ttk.Label(
            deepl_frame,
            text=(
                "Wenn aktiviert, nutzt die GenArt-Vorschlagslogik einen inoffiziellen Google-Lens-Browserabruf ohne API-Key. "
                "Dabei werden sichtbare Treffertexte und Seiten aus den Lens-Ergebnissen als zusaetzliches Signal fuer die GenArt-Suche verwendet. "
                "Das ist absichtlich als experimenteller Weg zu verstehen und kann sich durch Google-Aenderungen jederzeit veraendern."
            ),
            foreground="#5E6472",
            wraplength=1100,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))

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

        columns = ("article", "source", "short_id", "long_id", "genart", "images", "documents", "videos", "links")
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
        self.article_browser_tree.heading("genart", text="GenArt")
        self.article_browser_tree.heading("images", text="Bilder")
        self.article_browser_tree.heading("documents", text="Dokumente")
        self.article_browser_tree.heading("videos", text="Videos")
        self.article_browser_tree.heading("links", text="Web Links")
        self.article_browser_tree.column("article", width=220, anchor="w")
        self.article_browser_tree.column("source", width=90, anchor="center")
        self.article_browser_tree.column("short_id", width=100, anchor="center")
        self.article_browser_tree.column("long_id", width=100, anchor="center")
        self.article_browser_tree.column("genart", width=180, anchor="w")
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
            max_length=SHORT_TEXT_MAX_LENGTH,
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
            normalize_link_fn=normalize_youtube_url_for_embed,
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

    def choose_genart_source_file(self) -> None:
        initial_dir = (
            str(Path(self.genart_source_path_var.get()).parent)
            if self.genart_source_path_var.get().strip()
            else str(DEFAULT_GENART_SOURCE.parent)
        )
        selected = filedialog.askopenfilename(
            title="GenArt XLSX waehlen",
            initialdir=initial_dir,
            filetypes=[("Excel Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if selected:
            self.genart_source_path_var.set(selected)
            self._load_genart_catalog()

    def _load_genart_catalog(self, initial: bool = False) -> None:
        source_path = Path(self.genart_source_path_var.get().strip()) if self.genart_source_path_var.get().strip() else None
        if source_path is None or not source_path.exists():
            self.genart_registry = GenArtRegistry()
            self.genart_image_index_dirty = True
            self.genart_count_var.set("0 GenArts geladen")
            self.genart_suggestion_var.set("Keine GenArt-Datei geladen.")
            if hasattr(self, "genart_combo"):
                self.genart_combo.configure(values=[])
            if not initial and source_path is not None:
                self.status_var.set(f"GenArt-Datei nicht gefunden: {source_path}")
            return

        try:
            count = self.genart_registry.load_from_workbook(source_path)
        except Exception as exc:
            self.genart_image_index_dirty = True
            self.genart_count_var.set("0 GenArts geladen")
            self.genart_suggestion_var.set("GenArt-Datei konnte nicht geladen werden.")
            if hasattr(self, "genart_combo"):
                self.genart_combo.configure(values=[])
            if not initial:
                messagebox.showwarning(APP_TITLE, f"GenArt-Datei konnte nicht geladen werden:\n{exc}")
            return

        self.genart_count_var.set(f"{count} GenArts geladen")
        self.genart_image_index_dirty = True
        self.genart_suggestion_var.set(
            "GenArt-Katalog geladen. Du kannst im Feld tippen, ueber 'Suchen...' gezielt filtern oder Vorschlaege nutzen."
        )
        if hasattr(self, "genart_combo"):
            self._refresh_genart_combobox_values()
        if self.genart_display_var.get().strip():
            self._normalize_genart_selection()
        if not initial:
            self.status_var.set(f"GenArt-Katalog geladen: {count} Eintraege")

    def _refresh_genart_combobox_values(self) -> None:
        if not hasattr(self, "genart_combo"):
            return
        values = self.genart_registry.search_display_values(self.genart_display_var.get(), limit=250)
        self.genart_combo.configure(values=values)

    def _open_genart_search_dialog_event(self, _event: tk.Event[tk.Misc]) -> str:
        self._open_genart_search_dialog()
        return "break"

    def _open_genart_search_dialog(self) -> None:
        if not self.genart_registry.options:
            self._load_genart_catalog()
        if not self.genart_registry.options:
            messagebox.showinfo(APP_TITLE, "Es ist noch keine GenArt-Datei geladen.")
            return

        initial_query = self.genart_display_var.get().strip()
        dialog = GenArtSearchDialog(self.root, self.genart_registry, initial_query=initial_query)
        selected_option = dialog.show()
        if selected_option is None:
            return

        self._set_genart_option(selected_option)
        self.genart_suggestion_var.set(
            f"GenArt ausgewaehlt: {(selected_option.id + ' | ' + selected_option.bezeichnung).strip(' |')}"
        )
        self._write_live_section("genart")

    def _on_genart_focus_in(self, _event: tk.Event[tk.Misc]) -> None:
        self._refresh_genart_combobox_values()

    def _on_genart_key_release(self, event: tk.Event[tk.Misc]) -> None:
        if event.keysym in {"Return", "Escape", "Up", "Down", "Left", "Right", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"}:
            return
        self._refresh_genart_combobox_values()

    def _on_genart_return(self, _event: tk.Event[tk.Misc]) -> str:
        if self._normalize_genart_selection():
            self._write_live_section("genart")
            return "break"

        values_raw = self.genart_combo.cget("values")
        values = list(self.root.tk.splitlist(values_raw)) if isinstance(values_raw, str) else list(values_raw)
        if values:
            self.genart_display_var.set(str(values[0]))
            self._write_live_section("genart")
        return "break"

    def _normalize_genart_selection(self) -> bool:
        current_value = self.genart_display_var.get().strip()
        if not current_value:
            return True
        option = self.genart_registry.resolve(current_value)
        if option is None:
            self.genart_suggestion_var.set("GenArt nicht erkannt. Bitte aus der Liste waehlen oder Vorschlag nutzen.")
            return False
        self.genart_display_var.set(option.display_label())
        return True

    def _get_selected_genart(self) -> GenArtOption | None:
        return self.genart_registry.resolve(self.genart_display_var.get())

    def _get_selected_genart_values(self) -> tuple[str, str]:
        option = self._get_selected_genart()
        if option is not None:
            return option.id, option.bezeichnung
        raw = self.genart_display_var.get().strip()
        if not raw:
            return "", ""
        if self.genart_registry.options and "|" not in raw:
            return "", ""
        parts = [part.strip() for part in raw.split("|") if part.strip()]
        if not parts:
            return "", ""
        if len(parts) == 1:
            return parts[0], parts[0]
        return parts[0], parts[-1]

    def _set_genart_option(self, option: GenArtOption | None) -> None:
        if option is None:
            self.genart_display_var.set("")
            self._refresh_genart_combobox_values()
            return
        self.genart_display_var.set(option.display_label())
        self._refresh_genart_combobox_values()

    def _build_genart_image_context(self, image_rows: list[MediaRow]) -> str:
        parts = [extract_match_text_from_path_or_link(row.path_or_link) for row in image_rows[:3] if row.path_or_link.strip()]
        return " ".join(part for part in parts if part).strip()

    def _get_image_signature_for_path(self, path_or_link: str) -> ImageSignature | None:
        key = path_or_link.strip()
        if not key or Image is None:
            return None
        if key not in self.genart_image_signature_cache:
            try:
                self.genart_image_signature_cache[key] = build_image_signature(key)
            except Exception:  # pragma: no cover - depends on user files/network
                self.genart_image_signature_cache[key] = None
        return self.genart_image_signature_cache[key]

    def _ensure_genart_image_index(self) -> None:
        if not self.genart_image_index_dirty:
            return

        references: list[tuple[str, str, ImageSignature]] = []
        references_per_genart: dict[str, int] = {}
        for article_number in sorted(self.article_browser_records):
            if len(references) >= 400:
                break
            snapshot = self.article_browser_records[article_number]
            option = self.genart_registry.resolve(snapshot.genart_id) or self.genart_registry.resolve(snapshot.genart_bezeichnung)
            if option is None or not snapshot.image_rows:
                continue
            if references_per_genart.get(option.id, 0) >= 10:
                continue

            for row in snapshot.image_rows[:3]:
                signature = self._get_image_signature_for_path(row.path_or_link)
                if signature is None:
                    continue
                references.append((article_number, option.id, signature))
                references_per_genart[option.id] = references_per_genart.get(option.id, 0) + 1
                break

        self.genart_image_reference_index = references
        self.genart_image_index_dirty = False

    def _suggest_genart_from_images(self, image_rows: list[MediaRow], limit: int = 5) -> list[tuple[GenArtOption, float]]:
        if Image is None:
            return []

        current_signatures = [
            signature
            for signature in (self._get_image_signature_for_path(row.path_or_link) for row in image_rows[:3])
            if signature is not None
        ]
        if not current_signatures:
            return []

        self._ensure_genart_image_index()
        if not self.genart_image_reference_index:
            return []

        best_scores_by_id: dict[str, float] = {}
        matched_articles_by_id: dict[str, set[str]] = {}
        for current_signature in current_signatures:
            for article_number, option_id, reference_signature in self.genart_image_reference_index:
                similarity = compare_image_signatures(current_signature, reference_signature)
                if similarity < 0.58:
                    continue
                score = max(0.0, (similarity - 0.55) * 720)
                if score <= 0:
                    continue
                best_scores_by_id[option_id] = max(best_scores_by_id.get(option_id, 0.0), score)
                matched_articles_by_id.setdefault(option_id, set()).add(article_number)

        suggestions: list[tuple[GenArtOption, float]] = []
        for option_id, score in best_scores_by_id.items():
            option = self.genart_registry.resolve(option_id)
            if option is None:
                continue
            article_bonus = min(max(len(matched_articles_by_id.get(option_id, set())) - 1, 0) * 18, 54)
            suggestions.append((option, score + article_bonus))

        suggestions.sort(key=lambda item: (-item[1], item[0].display_label()))
        return suggestions[:limit]

    def _should_use_google_lens(self) -> bool:
        return self.google_lens_enabled_var.get()

    def _build_genart_source_label(self, suggestion: GenArtSuggestion) -> str:
        source_parts: list[str] = []
        if suggestion.web_score > 0:
            source_parts.append("Google Lens")
        if suggestion.text_score > 0:
            source_parts.append("Text")
        if suggestion.image_score > 0:
            source_parts.append("Bild")
        return " + ".join(source_parts) if source_parts else "Regel"

    def _get_google_lens_result_for_path(self, path_or_link: str) -> GoogleLensWebResult | None:
        key = path_or_link.strip()
        if not key:
            return None
        if key not in self.google_lens_web_cache:
            try:
                client = GoogleLensScraper()
                self.google_lens_web_cache[key] = client.detect_web(key)
            except Exception:  # pragma: no cover - depends on user files/network
                self.google_lens_web_cache[key] = None
        return self.google_lens_web_cache[key]

    def _collect_local_genart_suggestions(
        self,
        short_text: str,
        long_text: str,
        image_rows: list[MediaRow],
        category_context: str = "",
        limit: int = 5,
    ) -> list[GenArtSuggestion]:
        if not self.genart_registry.options:
            return []

        image_context = self._build_genart_image_context(image_rows)
        text_suggestions = self.genart_registry.suggest(
            short_text,
            long_text,
            image_context=image_context,
            category_context=category_context,
            limit=max(limit * 6, 24),
        )
        image_suggestions = self._suggest_genart_from_images(image_rows, limit=max(limit * 6, 24))

        combined_scores: dict[str, dict[str, object]] = {}
        for option, score in text_suggestions:
            combined_scores.setdefault(option.id, {"option": option, "text": 0.0, "image": 0.0})
            combined_scores[option.id]["text"] = score
        for option, score in image_suggestions:
            combined_scores.setdefault(option.id, {"option": option, "text": 0.0, "image": 0.0})
            combined_scores[option.id]["image"] = score

        suggestions: list[GenArtSuggestion] = []
        for data in combined_scores.values():
            option = data["option"]
            text_score = float(data["text"])
            image_score = float(data["image"])
            synergy_bonus = 55.0 if text_score > 0 and image_score > 0 else 0.0
            total_score = text_score + image_score + synergy_bonus
            suggestions.append(
                GenArtSuggestion(
                    option=option,
                    total_score=total_score,
                    text_score=text_score,
                    image_score=image_score,
                )
            )

        suggestions.sort(key=lambda item: (-item.total_score, -item.image_score, -item.text_score, item.option.display_label()))
        return suggestions[:limit]

    def _collect_google_lens_result(self, image_rows: list[MediaRow], limit: int = 2) -> GoogleLensWebResult | None:
        if not self._should_use_google_lens():
            return None

        aggregated = GoogleLensWebResult()
        for row in image_rows:
            result = self._get_google_lens_result_for_path(row.path_or_link)
            if result is None:
                continue

            for title in result.headline_lines:
                if title not in aggregated.headline_lines:
                    aggregated.headline_lines.append(title)
            for title in result.result_titles:
                if title not in aggregated.result_titles:
                    aggregated.result_titles.append(title)
            for snippet in result.result_snippets:
                if snippet not in aggregated.result_snippets:
                    aggregated.result_snippets.append(snippet)
            for page_url in result.page_urls:
                if page_url not in aggregated.page_urls:
                    aggregated.page_urls.append(page_url)

            if len(aggregated.headline_lines) >= 10 and len(aggregated.result_titles) >= 20:
                break
            limit -= 1
            if limit <= 0:
                break

        if not aggregated.as_context_text():
            return None
        return aggregated

    def _build_google_lens_reason(self, result: GoogleLensWebResult) -> str:
        parts: list[str] = []
        if result.headline_lines:
            parts.append("Lens: " + ", ".join(result.headline_lines[:3]))
        if result.result_titles:
            parts.append("Seiten: " + ", ".join(result.result_titles[:3]))
        return " | ".join(parts)

    def _collect_google_lens_genart_suggestions(
        self,
        short_text: str,
        long_text: str,
        image_rows: list[MediaRow],
        category_context: str = "",
        limit: int = 5,
    ) -> list[GenArtSuggestion]:
        web_result = self._collect_google_lens_result(image_rows)
        if web_result is None:
            return []

        web_context = web_result.as_context_text()
        suggestions = self.genart_registry.suggest(
            short_text="",
            long_text="",
            image_context="",
            category_context=category_context,
            web_context=web_context,
            limit=max(limit * 6, 24),
        )
        reason = self._build_google_lens_reason(web_result)
        return [
            GenArtSuggestion(
                option=option,
                total_score=score,
                web_score=score,
                web_reason=reason,
            )
            for option, score in suggestions[:limit]
        ]

    def _collect_genart_suggestions(
        self,
        short_text: str,
        long_text: str,
        image_rows: list[MediaRow],
        category_context: str = "",
        limit: int = 5,
    ) -> list[GenArtSuggestion]:
        local_suggestions = self._collect_local_genart_suggestions(
            short_text,
            long_text,
            image_rows,
            category_context=category_context,
            limit=max(limit * 8, 40),
        )
        suggestion_map: dict[str, GenArtSuggestion] = {
            suggestion.option.id: GenArtSuggestion(
                option=suggestion.option,
                total_score=suggestion.total_score,
                text_score=suggestion.text_score,
                image_score=suggestion.image_score,
                web_score=suggestion.web_score,
                web_reason=suggestion.web_reason,
            )
            for suggestion in local_suggestions
        }

        try:
            web_suggestions = self._collect_google_lens_genart_suggestions(
                short_text,
                long_text,
                image_rows,
                category_context=category_context,
                limit=max(limit * 8, 40),
            )
        except GoogleLensScrapeError as exc:
            self.status_var.set(f"Google Lens Fallback: {exc}")
            web_suggestions = []

        for web_suggestion in web_suggestions:
            existing = suggestion_map.get(web_suggestion.option.id)
            if existing is None:
                existing = web_suggestion
                suggestion_map[web_suggestion.option.id] = existing
            else:
                existing.web_score = max(existing.web_score, web_suggestion.web_score)
                if web_suggestion.web_reason:
                    existing.web_reason = web_suggestion.web_reason
                existing.total_score += web_suggestion.web_score

        suggestions = list(suggestion_map.values())
        suggestions.sort(
            key=lambda item: (
                -item.total_score,
                -item.web_score,
                -item.image_score,
                -item.text_score,
                item.option.display_label(),
            )
        )
        return suggestions[:limit]

    def _suggest_genart_option(self) -> tuple[GenArtOption | None, float]:
        if not self.genart_registry.options:
            return None, 0.0
        suggestions = self._collect_genart_suggestions(
            self.short_text_frame.get_german_text(),
            self.long_text_frame.get_german_text(),
            self.image_frame.get_rows(),
            category_context=self.current_kunzer_category_context,
        )
        if not suggestions:
            return None, 0.0
        return suggestions[0].option, suggestions[0].total_score

    def suggest_genart_for_current_article(self) -> None:
        if not self.genart_registry.options:
            self._load_genart_catalog()
        suggestions = self._collect_genart_suggestions(
            self.short_text_frame.get_german_text(),
            self.long_text_frame.get_german_text(),
            self.image_frame.get_rows(),
            category_context=self.current_kunzer_category_context,
        )
        if not suggestions:
            self.genart_suggestion_var.set("Keine passende GenArt gefunden.")
            messagebox.showinfo(APP_TITLE, "Es konnte keine passende GenArt vorgeschlagen werden.")
            return

        suggestion = suggestions[0]
        option = suggestion.option
        source_label = self._build_genart_source_label(suggestion)
        alternative_labels = ", ".join(candidate.option.display_label() for candidate in suggestions[1:3])

        self._set_genart_option(option)
        message = f"Vorschlag uebernommen: {option.display_label()} ({source_label}, Score {suggestion.total_score:.0f})"
        if suggestion.web_reason:
            message += f" | Google: {suggestion.web_reason}"
        if alternative_labels:
            message += f" | Alternativen: {alternative_labels}"
        self.genart_suggestion_var.set(message)
        self._write_live_section("genart")

    def _on_genart_focus_out(self, _event: tk.Event[tk.Misc]) -> None:
        if self._normalize_genart_selection():
            self._write_live_section("genart")

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
            genart_id=bundle.genart_id,
            genart_bezeichnung=bundle.genart_bezeichnung,
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
            genart_id=snapshot.genart_id,
            genart_bezeichnung=snapshot.genart_bezeichnung,
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
                    (snapshot.genart_bezeichnung or snapshot.genart_id),
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

    def _build_output_folder_signature(self, output_folder: Path) -> tuple[tuple[str, int, int], ...]:
        signature: list[tuple[str, int, int]] = []
        for file_name in [
            SHORT_TEXT_FILE[0],
            SHORT_MAPPING_FILE[0],
            LONG_TEXT_FILE[0],
            GENART_FILE[0],
            IMAGE_FILE[0],
            DOCUMENT_FILE[0],
            VIDEO_FILE[0],
            WEB_LINK_FILE[0],
        ]:
            path = output_folder / file_name
            if path.exists():
                stat = path.stat()
                signature.append((file_name, stat.st_mtime_ns, stat.st_size))
            else:
                signature.append((file_name, -1, -1))
        return tuple(signature)

    def _upsert_article_browser_from_bundle(self, bundle: ExportBundle, output_root: Path) -> None:
        self.article_browser_records[bundle.article_number] = self._snapshot_from_bundle(bundle, output_root)
        self.genart_image_index_dirty = True
        self.article_browser_cache_dir = str(output_root)
        self.article_browser_cache_signature = None
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
        snapshot.genart_id = bundle.genart_id
        snapshot.genart_bezeichnung = bundle.genart_bezeichnung

        if section in {"short_text", "all"}:
            snapshot.short_texts = bundle.short_texts
            snapshot.short_auto_uni = bundle.short_auto_uni
            snapshot.short_module_id = bundle.short_module_id
        if section in {"long_text", "all"}:
            snapshot.long_texts = bundle.long_texts
            snapshot.long_auto_uni = bundle.long_auto_uni
            snapshot.long_module_id = bundle.long_module_id
        if section in {"genart", "all"}:
            snapshot.genart_id = bundle.genart_id
            snapshot.genart_bezeichnung = bundle.genart_bezeichnung
        if section in {"images", "all"}:
            snapshot.image_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.image_rows]
        if section in {"documents", "all"}:
            snapshot.document_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.document_rows]
        if section in {"videos", "all"}:
            snapshot.video_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.video_rows]
        if section in {"web_links", "all"}:
            snapshot.web_rows = [MediaRow(row.path_or_link, art=row.art, sprache=row.sprache) for row in bundle.web_rows]

        self.article_browser_records[bundle.article_number] = snapshot
        self.genart_image_index_dirty = True
        self.article_browser_cache_dir = str(output_root)
        self.article_browser_cache_signature = None
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

    def _write_genart_live(self, bundle: ExportBundle, output_root: Path) -> None:
        written_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        write_workbook_with_upsert(
            output_root / GENART_FILE[0],
            GENART_FILE[1],
            GENART_HEADERS,
            [append_written_at(row, written_at) for row in build_genart_export_rows(bundle)],
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
            output_root.mkdir(parents=True, exist_ok=True)
            if section == "short_text":
                self._write_short_text_live(bundle, output_root)
            elif section == "long_text":
                self._write_long_text_live(bundle, output_root)
            elif section == "genart":
                self._write_genart_live(bundle, output_root)
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

    def _build_translation_set(
        self,
        german_text: str,
        client: DeepLClient | None,
        sanitize_short_text: bool = False,
    ) -> TranslationSet:
        values = {"de": german_text.strip(), "uni": german_text.strip()}
        if client and german_text.strip():
            values.update(client.translate_from_german(german_text))
        translations = TranslationSet(**values)
        if sanitize_short_text:
            return sanitize_short_translation_set(translations)
        return translations

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
        short_texts = self._build_translation_set(result.short_text_de, client, sanitize_short_text=True) if scrape_options["short_text"] else TranslationSet()
        long_texts = self._build_translation_set(result.long_text_de, client) if scrape_options["long_text"] else TranslationSet()
        image_rows = [MediaRow(link, art="5", sprache="255") for link in result.image_links] if scrape_options["images"] else []
        genart_suggestions = self._collect_genart_suggestions(
            result.short_text_de,
            result.long_text_de,
            image_rows,
            category_context=result.breadcrumb_text,
            limit=1,
        )
        suggested_genart = (
            genart_suggestions[0].option
            if genart_suggestions and genart_suggestions[0].total_score >= GENART_AUTO_APPLY_MIN_SCORE
            else None
        )
        return ExportBundle(
            article_number=article_number,
            short_module_id=short_module_id,
            long_module_id=long_module_id,
            short_texts=short_texts,
            short_auto_uni=True,
            long_texts=long_texts,
            long_auto_uni=True,
            genart_id=suggested_genart.id if suggested_genart else "",
            genart_bezeichnung=suggested_genart.bezeichnung if suggested_genart else "",
            image_rows=image_rows,
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
        short_translations: dict[str, str] | None = None,
        long_translations: dict[str, str] | None = None,
    ) -> None:
        with self._suspend_live_write():
            if not self.article_number_var.get().strip():
                self.article_number_var.set(result.article_number)
            else:
                self.article_number_var.set(normalize_article_number(result.article_number))
            self._ensure_ids_for_article(self.article_number_var.get())
            self.kunzer_product_url_var.set(result.product_url)
            self.current_kunzer_category_context = result.breadcrumb_text

            self.short_text_frame.set_value(
                TranslationSet(de=result.short_text_de, uni=result.short_text_de),
                auto_uni=True,
            )
            self.long_text_frame.set_value(
                TranslationSet(de=result.long_text_de, uni=result.long_text_de),
                auto_uni=True,
            )
            if short_translations:
                self.short_text_frame.apply_translations(short_translations)
            if long_translations:
                self.long_text_frame.apply_translations(long_translations)

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
            suggestions = self._collect_genart_suggestions(
                result.short_text_de,
                result.long_text_de,
                image_rows,
                category_context=result.breadcrumb_text,
                limit=3,
            )
            if suggestions and suggestions[0].total_score >= GENART_AUTO_APPLY_MIN_SCORE:
                suggestion = suggestions[0]
                self._set_genart_option(suggestion.option)
                source_label = self._build_genart_source_label(suggestion)
                self.genart_suggestion_var.set(
                    f"Automatisch vorgeschlagen: {suggestion.option.display_label()} ({source_label}, Score {suggestion.total_score:.0f})"
                )
            elif suggestions:
                suggestion = suggestions[0]
                source_label = self._build_genart_source_label(suggestion)
                self.genart_suggestion_var.set(
                    f"Vorschlag pruefen: {suggestion.option.display_label()} ({source_label}, Score {suggestion.total_score:.0f})"
                )
            else:
                self.genart_suggestion_var.set("Keine passende GenArt gefunden.")

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
        identifier = article_number_or_url.strip()
        if not identifier:
            messagebox.showwarning(APP_TITLE, "Bitte eine Artikelnummer oder Kunzer Produkt-URL eingeben.")
            return

        deepl_client: DeepLClient | None = None
        if self.auto_translate_after_scrape_var.get() and self.deepl_api_key_var.get().strip():
            try:
                deepl_client = self._build_deepl_client()
            except ValueError as exc:
                messagebox.showwarning(APP_TITLE, str(exc))
                return

        def worker() -> tuple[KunzerScrapeResult, dict[str, str] | None, dict[str, str] | None]:
            result = self.kunzer_scraper.scrape_product(identifier)
            if deepl_client is None:
                return result, None, None
            short_translations = deepl_client.translate_from_german(result.short_text_de)
            long_translations = deepl_client.translate_from_german(result.long_text_de)
            return result, short_translations, long_translations

        def on_success(payload: object) -> None:
            result, short_translations, long_translations = payload  # type: ignore[misc]
            self._apply_kunzer_result(
                result,
                translate_after_load=False,
                write_live=True,
                short_translations=short_translations,
                long_translations=long_translations,
            )
            self.status_var.set(f"Kunzer-Daten geladen: {result.article_number}")

        self._run_background_task(
            "Kunzer-Daten werden geladen ...",
            worker,
            on_success,
            "Kunzer-Import fehlgeschlagen",
        )

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

        def worker() -> dict[str, str]:
            return client.translate_from_german(german_text)

        def on_success(payload: object) -> None:
            translations = payload  # type: ignore[assignment]
            frame.apply_translations(translations)
            self.refresh_preview()
            self._write_live_database(status_message=f"Live gespeichert: {normalize_article_number(self.article_number_var.get())}")
            self.status_var.set(f"{label} mit DeepL aus Deutsch uebersetzt.")

        self._run_background_task(
            f"{label} wird mit DeepL uebersetzt ...",
            worker,
            on_success,
            f"{label} Uebersetzung fehlgeschlagen",
        )

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
            short_german = self.short_text_frame.get_german_text()
            long_german = self.long_text_frame.get_german_text()
            if not short_german and not long_german:
                raise ValueError("Bitte zuerst einen deutschen Kurztext oder Langtext eingeben.")
            client = self._build_deepl_client()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        def worker() -> tuple[dict[str, str], dict[str, str]]:
            short_translations = client.translate_from_german(short_german) if short_german else {}
            long_translations = client.translate_from_german(long_german) if long_german else {}
            return short_translations, long_translations

        def on_success(payload: object) -> None:
            short_translations, long_translations = payload  # type: ignore[misc]
            if short_translations:
                self.short_text_frame.apply_translations(short_translations)
            if long_translations:
                self.long_text_frame.apply_translations(long_translations)
            self.refresh_preview()
            self._write_live_database(status_message=f"Live gespeichert: {normalize_article_number(self.article_number_var.get())}")
            self.status_var.set("Kurzbezeichnung und Text mit DeepL aus Deutsch uebersetzt.")

        self._run_background_task(
            "Kurzbezeichnung und Text werden mit DeepL uebersetzt ...",
            worker,
            on_success,
            "DeepL Uebersetzung fehlgeschlagen",
        )

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
            self.current_kunzer_category_context = ""
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
        signature: tuple[tuple[str, int, int], ...] = ()

        if output_folder.exists():
            signature = self._build_output_folder_signature(output_folder)
            if (
                self.article_browser_cache_dir == str(output_folder)
                and self.article_browser_cache_signature == signature
            ):
                self._render_article_browser()
                return
            records = load_article_snapshots_from_folder(output_folder, "Output")

        self.article_browser_records = records
        self.article_browser_cache_dir = str(output_folder)
        self.article_browser_cache_signature = signature
        self.genart_image_index_dirty = True
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
        self.current_kunzer_category_context = ""
        self.article_number_var.set(snapshot.article_number)
        self.current_id_article_number = snapshot.article_number
        self.short_module_id_var.set(snapshot.short_module_id)
        self.long_module_id_var.set(snapshot.long_module_id)
        self.id_registry.remember_article_ids(snapshot.article_number, snapshot.short_module_id, snapshot.long_module_id)
        resolved_genart = self.genart_registry.resolve(snapshot.genart_id) or self.genart_registry.resolve(snapshot.genart_bezeichnung)
        if resolved_genart is not None:
            self._set_genart_option(resolved_genart)
        else:
            self.genart_display_var.set((snapshot.genart_id + " | " + snapshot.genart_bezeichnung).strip(" |"))
        if snapshot.genart_id or snapshot.genart_bezeichnung:
            self.genart_suggestion_var.set(
                f"Gespeicherte GenArt: {(snapshot.genart_id + ' | ' + snapshot.genart_bezeichnung).strip(' |')}"
            )
        else:
            self.genart_suggestion_var.set("Keine GenArt fuer diesen Artikel gespeichert.")
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
                (GENART_FILE, GENART_HEADERS),
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
        selected_genart_id, selected_genart_bezeichnung = self._get_selected_genart_values()

        return ExportBundle(
            article_number=article_number,
            short_module_id=short_module_id,
            long_module_id=long_module_id,
            short_texts=self.short_text_frame.get_value(),
            short_auto_uni=self.short_text_frame.auto_uni_var.get(),
            long_texts=self.long_text_frame.get_value(),
            long_auto_uni=self.long_text_frame.auto_uni_var.get(),
            genart_id=selected_genart_id,
            genart_bezeichnung=selected_genart_bezeichnung,
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
