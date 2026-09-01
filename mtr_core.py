# -*- coding: utf-8 -*-
from __future__ import annotations

import gzip
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

TOKEN_TAIL = r'(?:[^\s,;]|,(?=\S)|;(?=\S))*'
GENERIC_TEXT_ALIASES = {"оборудование", "мониторинг", "универсал", "источник", "монитор", "металлорукав", "контакт", "переход"}
# Explicitly confirmed removable brands / product families.  Keep this list
# separate from protected technical syntax so a future rule audit is simple.
CONFIRMED_GLOBAL_BRANDS = (
    "Унипол", "Гиперфлоу", "Метран", "Вэлан", "Ризур", "Рубеж",
    "Пензтяжпромарматура", "Волжский трубный завод",
)
CONFIRMED_GLOBAL_BRAND_RE = re.compile(
    r"(?i)(?<![\w])(?:" + "|".join(re.escape(x) for x in CONFIRMED_GLOBAL_BRANDS)
    + r")(?:-[A-Za-zА-Яа-яЁё0-9._/]+)?(?![\w])"
)
# Verified short PDF designations. They are allowed as manufacturer-free PDF
# fallbacks only with the strict patterns below.
PDF_SHORT_UNIQUE_TRIGGERS = {"КШГ", "ПМТД"}

# Absolute KEEP: survey-sheet / project references must never be removed.
OL_REF_RE = re.compile(r'(?i)(?:Комплектация\s+по\s+обосновывающему\s+документу\s+)?[0-9A-Za-zА-Яа-яЁё][0-9A-Za-zА-Яа-яЁё._/\-]{4,}(?:[-.]ОЛК?\d*(?:-\d+)?)')
OL_PHRASE_RE = re.compile(r'(?i)\bКомплектация\s+по\s+обосновывающему\s+документу\b')

def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name

def normalize_name(s: str) -> str:
    s = (s or "").lower().replace("ё", "е")
    s = re.sub(r'["«»“”„]', "", s)
    s = re.sub(
        r"\b(ооо|ао|зао|оао|пао|фгуп|нпо|нпп|нпф|нтк|нтц|"
        r"общество с ограниченной ответственностью|акционерное общество)\b",
        " ", s
    )
    return re.sub(r"\s+", " ", s).strip(" ,.")

def extract_inn(s: str) -> str:
    m = re.search(r"\bИНН\s*([0-9]{10,12})\b", s or "", re.I)
    return m.group(1) if m else ""

def normalize_code(value) -> str:
    s = str(value or "").strip()
    if re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    return s

def code_candidates(value) -> list[str]:
    """Return plausible Autodocs ids from plain or composite resource codes.

    Examples: 631-336532 -> 336532; 6410-456101-3 -> 456101.
    Short estimate/labour fragments such as 1-100-10 are deliberately ignored.
    """
    s = normalize_code(value)
    if not s:
        return []
    out = [s]
    # Autodocs ids in the compiled registry are 4-6 digits; allow up to 9
    # for forward compatibility, but never use 1-3 digit estimate fragments.
    groups = re.findall(r"(?<!\d)(\d{4,9})(?!\d)", s)
    groups.sort(key=len, reverse=True)
    for g in groups:
        if g not in out:
            out.append(g)
    return out

def _boundary_pattern(phrase: str, mode: str) -> str:
    esc = re.escape(phrase)
    left = r"(?<![\w])" if phrase and phrase[0].isalnum() else ""
    right = r"(?![\w])" if phrase and phrase[-1].isalnum() else ""
    if mode == "Точное совпадение":
        return left + esc + right
    if mode in ("Начало фирменного обозначения", "Слитное/дефисное обозначение", "Полный код КД"):
        # Aero IXIA model prefix occurs in both Latin and Cyrillic spelling.
        if phrase.upper() == "AI-":
            return left + r"(?:AI-|АИ-)[A-Za-zА-Яа-я0-9][A-Za-zА-Яа-я0-9._/+*\-,]*"
        # KD prefixes in estimates are sometimes written both with a dot and
        # with a space: "ЖНКЮ.464429.018" / "ЖНКЮ 464429.018".
        # Match only the designation itself so a closing bracket or quantity
        # such as "-1компл." is never swallowed together with the code.
        if phrase.endswith(".") and len(phrase) > 1:
            stem = re.escape(phrase[:-1])
            return left + stem + r"(?:\.\s*|\s+)[A-Za-zА-Яа-я0-9][A-Za-zА-Яа-я0-9._/\-]*"
        return left + esc + TOKEN_TAIL
    if mode == "Бренд + соседний индекс":
        if "Метран" in phrase:
            return r"(?<![\w])Метран(?:\s*-\s*|\s+)?[A-Za-zА-Яа-я0-9]" + TOKEN_TAIL
        return left + esc + r"(?:[\s\-]+[A-Za-zА-Яа-я0-9]" + TOKEN_TAIL + r")?"
    return left + esc + right

def _alias_pattern(alias: str) -> str:
    esc = re.escape(alias)
    left = r"(?<![\w])" if alias and alias[0].isalnum() else ""
    right = r"(?![\w])" if alias and alias[-1].isalnum() else ""
    return left + esc + right

# Technical fragments that must survive even when they are embedded inside a
# manufacturer model token. This is intentionally narrow: only unambiguous
# KEEP attributes are restored.
PROTECTED_FRAGMENT_RES = [
    re.compile(r"(?i)\bIP\d{2}[A-Z]?\b"),
    re.compile(r"(?i)(?<![\w])(?:УХЛ|ХЛ)(?:\d(?:\.\d)?)?(?![\w])"),
    re.compile(r"(?i)(?<![\w])(?:[012]?Ex[a-z]*II[ABC][TТ][1-6][A-Za-z0-9]*)(?![\w])"),
    re.compile(r"(?i)(?<![\w])II[ABC][TТ][1-6](?![\w])"),
    re.compile(r"(?i)(?<![\w])Exd(?![\w])"),
    re.compile(r"(?i)\bDN\s*[-=]?\s*\d+(?:[.,]\d+)?\b"),
    re.compile(r"(?i)\bPN\s*[-=]?\s*\d+(?:[.,]\d+)?\b"),
    re.compile(r"(?i)\bRAL\s*\d{3,4}\b"),
    OL_PHRASE_RE,
    OL_REF_RE,
]

def _protected_fragments(text: str) -> list[str]:
    value_text = text or ""
    found = []
    seen = set()
    # Project / survey-sheet references have the highest priority. A project
    # code can itself contain tokens such as "ХЛ"; do not restore those a
    # second time outside the full protected reference.
    absolute_spans = []
    for pat in (OL_PHRASE_RE, OL_REF_RE):
        for m in pat.finditer(value_text):
            absolute_spans.append((m.start(), m.end()))
            value = m.group(0).strip(" -_/.,;")
            key = value.casefold().replace(" ", "")
            if value and key not in seen:
                seen.add(key); found.append(value)
    for pat in PROTECTED_FRAGMENT_RES:
        if pat in (OL_PHRASE_RE, OL_REF_RE):
            continue
        for m in pat.finditer(value_text):
            if any(a <= m.start() and m.end() <= b for a, b in absolute_spans):
                continue
            value = m.group(0).strip(" -_/.,;")
            key = value.casefold().replace(" ", "")
            if value and key not in seen:
                seen.add(key); found.append(value)
    return found

def keep_signature(text: str) -> dict[str, list[str]]:
    """Normalized KEEP tokens used by regression audits.

    This is diagnostic only; it never drives deletion.
    """
    value = str(text or "")
    pats = {
        "IP": re.compile(r"(?i)\bIP\d{2}[A-Z]?\b"),
        "UHL": re.compile(r"(?i)(?<![\w])(?:УХЛ|ХЛ)(?:\d(?:[.]\d)?)?(?![\w])"),
        "EX": re.compile(r"(?i)(?<![\w])(?:[012]?Ex[a-z]*II[ABC][TТ][1-6][A-Za-z0-9]*|II[ABC][TТ][1-6]|Exd)(?![\w])"),
        "DN": re.compile(r"(?i)\bDN\s*[-=]?\s*\d+(?:[.,]\d+)?\b"),
        "PN": re.compile(r"(?i)\bPN\s*[-=]?\s*\d+(?:[.,]\d+)?\b"),
        "RAL": re.compile(r"(?i)\bRAL\s*\d{3,4}\b"),
        "SDR": re.compile(r"(?i)\bSDR\s*\d+(?:[.,]\d+)?\b"),
        "GOST": re.compile(r"(?i)\bГОСТ(?:\s+Р)?\s*[0-9][0-9.\-–—/]*"),
        "OL": OL_REF_RE,
        "INV": re.compile(r"(?i)\bинв[.]?\s*№?\s*[A-Za-zА-Яа-я0-9/._\-]+"),
    }
    out = {}
    for key, pat in pats.items():
        vals=[]
        for m in pat.finditer(value):
            v=re.sub(r"\s+", "", m.group(0)).casefold().rstrip(".,;:")
            vals.append(v)
        out[key]=sorted(vals)
    return out


def _rule_replacement(match: re.Match) -> str:
    keep = _protected_fragments(match.group(0))
    return (" " + " ".join(keep) + " ") if keep else " "

# PDF redaction has to be even more conservative than text replacement.
# These fragments are never allowed to fall inside a physical redaction box.
PDF_PROTECTED_RES = PROTECTED_FRAGMENT_RES + [
    re.compile(r"(?i)\bSDR\s*\d+(?:[.,]\d+)?\b"),
    re.compile(r"(?i)\bГОСТ(?:\s+Р)?\s*[0-9][0-9.\-–—/]*"),
    re.compile(r"(?i)\b(?:ОЛК?|опросный\s+лист)\b"),
    OL_PHRASE_RE,
    OL_REF_RE,
    re.compile(r"(?i)\bсталь\s+(?:марки\s+)?[A-Za-zА-Яа-я0-9.-]+\b"),
    re.compile(r"(?i)(?<![A-Za-zА-Яа-я0-9])Ст\.?\s*[0-9]{1,3}[A-Za-zА-Яа-я0-9.-]*\b"),
]

def _subtract_ranges(start: int, end: int, protected: list[tuple[int, int]]) -> list[tuple[int, int]]:
    parts = [(start, end)]
    for ps, pe in protected:
        if pe <= start or ps >= end:
            continue
        next_parts = []
        for a, b in parts:
            if pe <= a or ps >= b:
                next_parts.append((a, b))
                continue
            if a < ps:
                next_parts.append((a, ps))
            if pe < b:
                next_parts.append((pe, b))
        parts = next_parts
    return [(a, b) for a, b in parts if b > a]

# System/global-safe patterns.

# Service references from supplier letters / TKP after a protected OL link.
MONTHS = r"(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
LETTER_TAIL_RE = re.compile(
    rf"(?i)(?<![\w])№\s*[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._/\-–—]{{2,50}}"
    rf"\s+от\s+(?:\d{{1,2}}\s+{MONTHS}\s+\d{{4}}|\d{{1,2}}[.]\d{{1,2}}[.]\d{{2,4}})\s*г?[.]?"
    r"(?:\s*,?\s*(?:п|поз)[.]?\s*\d+(?:[.]\d+)*)?"
)
REVERSE_TU_RE = re.compile(
    r"\(\s*[A-Za-zА-Яа-я0-9][A-Za-zА-Яа-я0-9._/\- ]{2,}\s+ТУ\s*\)", re.I
)
ARTICLE_RE = re.compile(
    r'(?<!\w)(?:артикул|арт\.?|кат\.?\s*№|каталожный\s+номер)\s*[:№,\-]?\s*'
    r'[A-Za-zА-Яа-я0-9_./*"\'\-]+', re.I
)
VO_TU_RE = re.compile(r"(?<!\w)ВО\s+(?=ТУ\b)")
ATTACHED_TU_RE = re.compile(
    r"(?<!\w)[A-Za-zА-Яа-я0-9_.\-/]*[A-Za-zА-Яа-я0-9_.\-/]ТУ\b"
)
TU_WRAP_RE = re.compile(
    r"(?<!\w)ТУ\s*[-–—]?\s*"
    r"(?:(?!(?:DN|PN|IP|ГОСТ|SDR)\b)(?:[A-ZА-Я]{1,4}|\d{1,3})\s+)?"
    r"(?=[A-Za-zА-Яа-я0-9_.\-/]*\d)[A-Za-zА-Яа-я0-9_.\-/–—]*[-–—]"
    r"[ \t]*\n[ \t]*(?=[A-Za-zА-Яа-я0-9_.\-/–—]*\d)[A-Za-zА-Яа-я0-9_.\-/–—]+", re.I
)
TU_CODE_RE = re.compile(
    r"(?<!\w)ТУ\s*[-–—]?\s*"
    r"(?:(?!(?:DN|PN|IP|ГОСТ|SDR)\b)(?:[A-ZА-Я]{1,4}|\d{1,3})\s+)?"
    r"(?=[A-Za-zА-Яа-я0-9_.\-/]*\d)[A-Za-zА-Яа-я0-9_.\-/–—]+"
    # OCR / source tables may split one TU number with spaces around a
    # hyphen or between its numeric groups. Consume only further code-like
    # numeric groups, never following descriptive text or a protected OL tail.
    r"(?:\s*(?:[-–—]\s*)?(?=\d[0-9.\-/–—]*\b)\d[0-9.\-/–—]*)*", re.I
)
STANDALONE_TU_RE = re.compile(r"(?<!\w)ТУ(?!\w)", re.I)
CONTACT_RE = re.compile(
    r"(?i)(?:\bИНН\s*\d{10,12}\b|\bКПП\s*\d{9}\b|"
    r"\b[\w.+-]+@[\w.-]+\.[A-Za-zА-Яа-я]{2,}\b|"
    r"https?://\S+|www\.\S+)"
)
REQUEST_RE = re.compile(r"(?i)\bЗаявка\s*№\s*[A-Za-zА-Яа-я0-9._/\-]+")

LEGAL_FORM = r"(?:ООО|АО|ЗАО|ОАО|ПАО|НПО|НПП|ФГУП)"
# OCR often confuses visually identical Cyrillic/Latin letters in legal forms
# (ООО -> OOO, АО -> AO). Use the tolerant form only for PDF service-data
# detection, never as a general manufacturer heuristic.
PDF_LEGAL_FORM = r"(?:ООО|АО|ЗАО|ОАО|ПАО|НПО|НПП|ФГУП|[ОO]{3}|[ОO][АA][ОO]|З[АA][ОO]|П[АA][ОO]|[АA][ОO]|FGUP)"
PDF_ORG_WITH_INN_RE = re.compile(
    rf"(?i)\b{PDF_LEGAL_FORM}\b(?:(?!\bИНН\b)[\s\S]){{1,180}}\bИНН\b"
)
# Quoted organization name, including imperfect nested quotes commonly found
# in source estimates. The lookahead prevents stopping at an inner quote.
QUOTED_ORG = r'["«][^,\n]{1,140}?["»](?=\s*(?:г\.?\s*|Россия\b|№|от\b|[,;.\]]|\(|по\b|$))'
ROLE_ORG_RE = re.compile(
    rf"(?i)\b(?:производитель|изготовитель|поставщик|разработчик)\s*[:=–—-]?\s*"
    rf"(?:{LEGAL_FORM}\s+)?{QUOTED_ORG}(?:\s+[^,;.\n]{{1,50}}\s+филиал)?"
)
PO_ORG_RE = re.compile(
    rf"(?i)\bпо\s+{LEGAL_FORM}\s+{QUOTED_ORG}"
    rf"(?:\s+г\.?\s*[А-ЯЁA-Z][А-Яа-яЁёA-Za-z .-]{{1,45}})?"
)
KP_ORG_RE = re.compile(
    rf"(?i)\bКП\s+{LEGAL_FORM}\s+{QUOTED_ORG}"
)
ROLE_BARE_ORG_RE = re.compile(
    r"(?i)\b(?:производитель|изготовитель|поставщик|разработчик)\s*[:=–—-]?\s*"
    r"([A-Z][A-Za-z0-9&.+-]*(?:\s+[A-Z][A-Za-z0-9&.+-]*){0,4})(?=\s*[,;.]|$)"
)
PRODUCTION_ORG_RE = re.compile(
    rf"(?i)\b(?:производства|произв\.?)\s+{LEGAL_FORM}\s+{QUOTED_ORG}"
    rf"(?:\s+(?:Россия,?\s*)?г\.?\s*[А-ЯЁA-Z][А-Яа-яЁёA-Za-z .-]{{1,45}})?"
)
BRACKET_ORG_RE = re.compile(
    rf"(?i)\[\s*{LEGAL_FORM}\s+{QUOTED_ORG}\s*\]"
)
END_ORG_RE = re.compile(
    rf"(?i){LEGAL_FORM}\s+{QUOTED_ORG}(?=\s*[).;]*$)"
)
COMMERCIAL_ORG_RE = re.compile(
    rf"(?i)\b(Коммерческое\s+предложение|Прайс\s*лист)\s+{LEGAL_FORM}\s+"
    rf"(?:{QUOTED_ORG}|[А-ЯЁA-Z][А-Яа-яЁёA-Za-z0-9&.+-]*(?:\s+[А-ЯЁA-Z][А-Яа-яЁёA-Za-z0-9&.+-]*){{0,4}})"
)
COMMERCIAL_REFERENCE_RE = re.compile(
    rf"(?i)\b(?:Коммерческое\s+предложение|Прайс\s*лист)\s+{LEGAL_FORM}\s+"
    rf"(?:(?!\.\s|$).){{1,240}}(?:\.(?=\s|$)|$)"
)
COMMERCIAL_DANGLING_RE = re.compile(
    r"(?i)\b(?:Коммерческое\s+предложение|Прайс\s*лист)\s*[\"«»]?\s*"
    r"(?:№\s*[^,;.\s]+\s*)?(?:от\s+\d{1,2}\.\d{1,2}\.\d{2,4})?"
    r"(?:\s*,?\s*поз\.?\s*\d+)?\s*[\"«»]?"
    r"(?=[.;]|$)"
)

class Anonymizer:
    def __init__(self, data_file: str | os.PathLike | None = None):
        p = Path(data_file) if data_file else resource_path("mtr_data.json.gz")
        with gzip.open(p, "rt", encoding="utf-8") as f:
            data = json.load(f)
        self.version = data.get("version", "")
        self.meta = data.get("meta", {})
        self.global_unique_rules = list(data.get("global_unique_rules", []))
        self.global_unique_rules.sort(key=lambda x: len(str(x.get("trigger", x.get("regex", "")))), reverse=True)
        # Entries labelled "recovery" were promoted from observed output
        # residues without enough evidence that the token is never a technical
        # model. They remain useful for review, but are unsafe for deletion.
        self.conservative_model_aliases = {
            normalize_name(str(row.get("trigger", "")))
            for row in self.global_unique_rules
            if "recovery" in str(row.get("note", "")).casefold() and row.get("trigger")
        }
        self.registry = {str(k): v for k, v in data.get("registry", {}).items()}

        self.rules_by_inn = defaultdict(list)
        self.rules_by_name = defaultdict(list)
        for rr in data.get("rules", []):
            inn = str(rr.get("inn", "")).strip()
            if inn:
                self.rules_by_inn[inn].append(rr)
            self.rules_by_name[normalize_name(rr.get("manufacturer", ""))].append(rr)
        for d in (self.rules_by_inn, self.rules_by_name):
            for k in list(d):
                d[k].sort(
                    key=lambda x: (x.get("apply") == "Бренд + соседний индекс",
                                   len(x.get("trigger", ""))),
                    reverse=True
                )

        self.aliases_by_inn = defaultdict(list)
        self.alias_index = {}
        self.manufacturer_name_by_inn = {}
        for row in data.get("aliases", []):
            inn = str(row.get("inn", "")).strip()
            alias = str(row.get("alias", "")).strip()
            manufacturer = str(row.get("manufacturer", "")).strip()
            if inn and manufacturer:
                self.manufacturer_name_by_inn.setdefault(inn, manufacturer)
            if inn and alias:
                self.aliases_by_inn[inn].append(alias)
                n = normalize_name(alias)
                if len(n) >= 4:
                    # Only exact normalized alias lookup is automatic.
                    self.alias_index.setdefault(n, inn)
        for inn in list(self.aliases_by_inn):
            self.aliases_by_inn[inn] = sorted(set(self.aliases_by_inn[inn]), key=len, reverse=True)
        # Fast lookup for factory aliases embedded in MTR names. Index by the
        # first normalized word instead of scanning all ~4k aliases per row.
        self.alias_search_index = defaultdict(list)
        for alias_norm, alias_inn in self.alias_index.items():
            words = re.findall(r"[\w]+", alias_norm, flags=re.U)
            if words:
                self.alias_search_index[words[0]].append((alias_norm, alias_inn))
        for key in list(self.alias_search_index):
            self.alias_search_index[key].sort(key=lambda x: len(x[0]), reverse=True)

        self.pdf_markers = []
        for row in data.get("pdf_markers", []):
            marker = str(row.get("contains", "")).strip()
            inn = str(row.get("inn", "")).strip()
            manufacturer = str(row.get("manufacturer", "")).strip()
            if marker and inn:
                self.pdf_markers.append((marker, inn, manufacturer))
                if manufacturer:
                    self.manufacturer_name_by_inn.setdefault(inn, manufacturer)
        self.pdf_markers.sort(key=lambda x: len(x[0]), reverse=True)

        # In PDFs the Autodocs code / supplier column is often empty.  In that
        # case we may still safely remove a firm designation when the trigger is
        # unique to one manufacturer in the internal database.  Short/generic
        # triggers are deliberately excluded from this fallback.
        trigger_inns = defaultdict(set)
        trigger_rules = defaultdict(list)
        for inn_key, rows in self.rules_by_inn.items():
            for rr in rows:
                trigger = str(rr.get("trigger", "")).strip()
                if not trigger:
                    continue
                key = trigger.casefold()
                trigger_inns[key].add(inn_key)
                trigger_rules[key].append(rr)
        self.pdf_unique_rules = []
        for key, inns in trigger_inns.items():
            if len(inns) != 1:
                continue
            rr = trigger_rules[key][0]
            trigger = str(rr.get("trigger", "")).strip()
            compact = re.sub(r"[^A-Za-zА-Яа-яЁё0-9]", "", trigger)
            if len(compact) < 5 and trigger.upper() not in PDF_SHORT_UNIQUE_TRIGGERS:
                continue
            if normalize_name(trigger) in GENERIC_TEXT_ALIASES:
                continue
            self.pdf_unique_rules.append(rr)
        self.pdf_unique_rules.sort(key=lambda x: len(str(x.get("trigger", ""))), reverse=True)

        self.known_inns = set(self.manufacturer_name_by_inn) | set(self.rules_by_inn)

    def resolve_factory(self, code: str = "", factory: str = "") -> tuple[str, str]:
        factory = str(factory or "").strip()
        # Accept both a plain Autodocs id and composite estimate/resource codes.
        for candidate in code_candidates(code):
            if candidate in self.registry and self.registry[candidate]:
                factory = self.registry[candidate]
                break
        inn = extract_inn(factory)
        if inn and inn in self.known_inns:
            return factory, inn

        nf = normalize_name(factory)
        if not nf:
            return factory, ""
        if nf in self.alias_index:
            inn = self.alias_index[nf]
            return factory or self.manufacturer_name_by_inn.get(inn, ""), inn

        # Conservative substring check: only long confirmed aliases. Use the
        # word index instead of scanning all aliases for every row.
        best = ("", "")
        words = set(re.findall(r"[\w]+", nf, flags=re.U))
        candidates = []
        seen = set()
        for word in words:
            for item in self.alias_search_index.get(word, []):
                if item not in seen:
                    seen.add(item)
                    candidates.append(item)
        for alias_norm, alias_inn in candidates:
            if len(alias_norm) >= 7 and alias_norm in nf and len(alias_norm) > len(best[0]):
                best = (alias_norm, alias_inn)
        if best[1]:
            return factory, best[1]
        return factory, ""

    def _rules(self, factory: str, inn: str):
        if inn and inn in self.rules_by_inn:
            return self.rules_by_inn[inn]
        nf = normalize_name(factory)
        return self.rules_by_name.get(nf, [])

    @staticmethod
    def _cleanup(text: str) -> str:
        terminal_article = bool(re.search(ARTICLE_RE.pattern + r"\s*[.]?\s*$", text or "", re.I))
        # ГОСТ is a protected normative reference. In combined labels such as
        # "ГОСТ/ТУ NS-1" remove only the TU part, never ГОСТ itself.
        text = re.sub(r"(?i)\bГОСТ\s*/\s*ТУ\b", "ГОСТ", text)
        s = REVERSE_TU_RE.sub(_rule_replacement, text)
        for _ in range(3):
            s = ARTICLE_RE.sub(" ", s)
        s = VO_TU_RE.sub(" ", s)
        s = ATTACHED_TU_RE.sub(_rule_replacement, s)
        s = TU_CODE_RE.sub(_rule_replacement, s)
        s = STANDALONE_TU_RE.sub(" ", s)
        # Explicit producer/supplier contexts are safe to remove even when the
        # organization is not yet in the internal alias registry. Do not remove
        # arbitrary company mentions (for example customer/operator references).
        s = ROLE_ORG_RE.sub(" ", s)
        s = ROLE_BARE_ORG_RE.sub(" ", s)
        s = PRODUCTION_ORG_RE.sub(" ", s)
        s = PO_ORG_RE.sub(" ", s)
        s = BRACKET_ORG_RE.sub(" ", s)
        s = END_ORG_RE.sub(" ", s)
        s = COMMERCIAL_REFERENCE_RE.sub(" ", s)
        s = COMMERCIAL_ORG_RE.sub(r"\1 ", s)
        s = COMMERCIAL_DANGLING_RE.sub(" ", s)
        s = LETTER_TAIL_RE.sub(" ", s)
        s = REQUEST_RE.sub(" ", s)
        s = KP_ORG_RE.sub("КП ", s)
        s = CONTACT_RE.sub(" ", s)
        s = re.sub(r"(?i)\bпо\s+типу\s*(?=(?:комплектация|№|Масса\b|[,;.]|$))", " ", s)
        s = re.sub(r"(?i)\bпо\s*(?=(?:комплектация|[,;.]|$))", " ", s)
        # Remove a dangling role label when the producer/developer alias was
        # deleted immediately after it.
        s = re.sub(r"(?i)\b(?:производитель|изготовитель|поставщик|разработчик)\s*[:=–—-]?\s*(?=(?:№|[,;.]|$))", " ", s)

        s = re.sub(r"\s+", " ", s)
        s = re.sub(r"\s+([,;:.])", r"\1", s)
        s = re.sub(r"([,;])\s*([,;])+", r"\1", s)
        s = re.sub(r"\(\s*\)", " ", s)
        s = re.sub(r"\[\s*\]", " ", s)
        # Remove only empty quote pairs created by a deleted quoted identifier.
        s = re.sub(r'["«»]\s*["«»]', " ", s)
        # A period may belong to the technical designation immediately before
        # a removed TU tail. Preserve it; only discard a period belonging to a
        # terminal article expression that was itself removed.
        s = re.sub(r"\s+", " ", s).strip(" ,;-")
        if terminal_article:
            s = s.rstrip(".")
        return s.strip()

    def anonymize(self, name: str, code: str = "", factory: str = "") -> dict:
        original = str(name or "")
        fixed_factory, fixed_inn = self.resolve_factory(code, factory)
        resolved_factory, inn = fixed_factory, fixed_inn
        if not inn:
            _, inferred_inn = self.identify_in_text(original)
            if inferred_inn:
                inn = inferred_inn
                resolved_factory = self.manufacturer_name_by_inn.get(inn, resolved_factory)

        def apply_once(value: str, active_factory: str, active_inn: str):
            s = value
            applied_local = []
            s, brand_count = CONFIRMED_GLOBAL_BRAND_RE.subn(
                lambda match: (applied_local.append(match.group(0)) or " "), s
            )
            # Confirmed global firm identifiers are safe without manufacturer context.
            for rr in self.global_unique_rules:
                trigger = str(rr.get("trigger", "")).strip()
                rx = str(rr.get("regex", "")).strip()
                if not trigger and not rx:
                    continue
                note = str(rr.get("note", "")).casefold()
                if "recovery" in note or "model family" in note:
                    continue
                pat = rx if rx else _boundary_pattern(trigger, str(rr.get("apply", "Точное совпадение")))
                # Vehicle make plus a numeric designation describes the engine
                # or chassis configuration and must survive as a unit.
                if trigger.casefold() == "камаз" and re.search(
                    r"(?i)(?:двигатель\s*-?\s*|на\s+шасси\s+)КАМАЗ-\d", s
                ):
                    continue
                s2, n = re.subn(pat, _rule_replacement, s, flags=re.I)
                if n:
                    applied_local.append(trigger or rx)
                    s = s2
            for rr in self._rules(active_factory, active_inn):
                trigger = str(rr.get("trigger", "")).strip()
                if not trigger:
                    continue
                # Catalogue / design documentation codes and non-exact model
                # families are technical designations unless an independently
                # confirmed brand-only rule exists. Earlier data marked these
                # modes as removable and consequently truncated cable marks,
                # instrument scales and electrical models. Be conservative.
                rule_type = str(rr.get("type", "")).casefold()
                apply_mode = str(rr.get("apply", ""))
                if "код кд" in rule_type or (
                    "серия / модель" in rule_type and apply_mode != "Точное совпадение"
                ):
                    continue
                if "обозначение с префиксом" in rule_type and re.search(
                    _boundary_pattern(trigger, apply_mode) + r"\s+\d", s, re.I
                ):
                    continue
                pat = _boundary_pattern(trigger, str(rr.get("apply", "")))
                s2, n = re.subn(pat, _rule_replacement, s, flags=re.I)
                if n:
                    applied_local.append(trigger)
                    s = s2

            if active_inn:
                for alias in self.aliases_by_inn.get(active_inn, []):
                    if len(alias) < 3:
                        continue
                    alias_norm = normalize_name(alias)
                    if alias_norm in self.conservative_model_aliases:
                        continue
                    # A registered alias immediately followed by a numeric
                    # suffix is often the full technical grade/model rather
                    # than a standalone producer mention (e.g. GANK-4).
                    if re.search(_alias_pattern(alias) + r"\s*-?\s*\d", s, re.I):
                        continue
                    if re.search(r'["«]' + _alias_pattern(alias) + r'["»]', s, re.I):
                        continue
                    s2, n = re.subn(_alias_pattern(alias), " ", s, flags=re.I)
                    if n:
                        applied_local.append(alias)
                        s = s2

                    if len(alias_norm) >= 4 and alias_norm not in GENERIC_TEXT_ALIASES:
                        if re.search(r'["«]' + _alias_pattern(alias_norm) + r'["»]', s, re.I):
                            continue
                        base_pat = _alias_pattern(alias_norm)
                        model = (
                            r"(?:-(?:[A-Za-zА-Яа-я0-9][A-Za-zА-Яа-я0-9._/+*\-,]*))?"
                            r"(?-i:(?:\s+(?!(?:DN|PN|IP|ГОСТ|ТУ)\b)"
                            r"(?:[A-ZА-ЯЁ]{1,6}(?:[-./][A-ZА-ЯЁ0-9]+)+|[A-ZА-ЯЁ]{1,6}\d[A-ZА-ЯЁ0-9._/\-]*|[A-ZА-ЯЁ]{2,5}\b))?)"
                        )
                        s2, n2 = re.subn(base_pat + model, _rule_replacement, s, flags=re.I)
                        if n2:
                            applied_local.append(alias_norm)
                            s = s2
            return self._cleanup(s), applied_local

        # Converge inside one run. If the code/factory did not resolve a
        # manufacturer, a line may contain several confirmed organization
        # identifiers (supplier + producer). After removing the first one,
        # re-identify the next confirmed alias and continue. A known Autodocs
        # manufacturer stays fixed and cannot be replaced by incidental text.
        s = original
        applied = []
        ever_identified = bool(inn)
        for _ in range(8):
            s2, ap = apply_once(s, resolved_factory, inn)
            applied.extend(ap)
            changed = s2 != s
            s = s2

            next_inn = inn
            next_factory = resolved_factory
            if not fixed_inn:
                _, inferred = self.identify_in_text(s)
                if inferred:
                    next_inn = inferred
                    next_factory = self.manufacturer_name_by_inn.get(inferred, next_factory)
                    ever_identified = True
                else:
                    next_inn = ""
                    next_factory = ""

            if not changed and next_inn == inn:
                break
            inn, resolved_factory = next_inn, next_factory

        status = "ЗЕЛЁНЫЙ" if (fixed_inn or ever_identified or bool(applied)) else "ЖЁЛТЫЙ"
        reason = "" if status == "ЗЕЛЁНЫЙ" else "Производитель не определён: применены только безопасные системные правила"
        return {
            "text": s,
            "status": status,
            "reason": reason,
            "factory": fixed_factory or resolved_factory,
            "inn": fixed_inn or inn,
            "applied": applied,
            "changed": s != original,
        }

    def identify_in_text(self, text: str) -> tuple[str, str]:
        """Resolve a manufacturer from a line/block, conservatively."""
        inn = extract_inn(text)
        if inn and inn in self.known_inns:
            return text, inn
        nt = normalize_name(text)
        # Confirmed aliases only, with real phrase boundaries. Prefer longest.
        # This prevents substring false positives: "мониторинга" must not
        # resolve to ООО "Мониторинг" and "универсальный" to "Универсал".
        best = ("", "")
        words = set(re.findall(r"[\w]+", nt, flags=re.U))
        candidates = []
        seen = set()
        for word in words:
            for item in self.alias_search_index.get(word, []):
                if item not in seen:
                    seen.add(item)
                    candidates.append(item)
        # A few legacy aliases are ordinary Russian nouns. They are valid
        # when resolved by Autodocs code/INN, but must never identify a
        # manufacturer merely because that common word appears in a name.
        for alias_norm, alias_inn in candidates:
            if alias_norm in GENERIC_TEXT_ALIASES:
                continue
            if len(alias_norm) < 7 or alias_norm not in nt:
                continue
            pat = r"(?<![\w])" + re.escape(alias_norm) + r"(?![\w])"
            if re.search(pat, nt, flags=re.I) and len(alias_norm) > len(best[0]):
                best = (alias_norm, alias_inn)
        return text, best[1]

    def identify_pdf_marker(self, text: str) -> tuple[str, str]:
        """Resolve PDF manufacturer from a verified TU/standard marker.

        This fallback is intentionally PDF-only: a marker is used only when it
        was explicitly verified and stored in the compiled data.
        """
        value = str(text or "")
        for marker, inn, manufacturer in self.pdf_markers:
            if marker in value:
                return manufacturer or self.manufacturer_name_by_inn.get(inn, ""), inn
        return "", ""

    def redaction_spans(self, text: str, code: str = "", factory: str = "") -> list[tuple[int, int, str]]:
        """Return exact character ranges that are safe to physically redact.

        For PDF-only fallback, firm rules may be applied without a resolved
        manufacturer only when the trigger is unique to one manufacturer in
        the compiled database. Protected technical fragments are subtracted
        from every candidate range before it reaches the PDF renderer.
        """
        text = str(text or "")
        resolved_factory, inn = self.resolve_factory(code, factory)
        if not inn:
            _, inn2 = self.identify_in_text(text)
            inn = inn2
            if inn:
                resolved_factory = self.manufacturer_name_by_inn.get(inn, "")
        if not inn:
            marker_factory, marker_inn = self.identify_pdf_marker(text)
            if marker_inn:
                inn = marker_inn
                resolved_factory = marker_factory

        candidates: list[tuple[int, int, str]] = []

        for match in CONFIRMED_GLOBAL_BRAND_RE.finditer(text):
            candidates.append((match.start(), match.end(), "confirmed_brand:" + match.group(0)))

        def add_rule(rr, label_prefix="rule"):
            trigger = str(rr.get("trigger", "")).strip()
            if not trigger:
                return
            rule_type = str(rr.get("type", "")).casefold()
            apply_mode = str(rr.get("apply", ""))
            if "код кд" in rule_type or (
                "серия / модель" in rule_type and apply_mode != "Точное совпадение"
            ):
                return
            if "обозначение с префиксом" in rule_type and re.search(
                _boundary_pattern(trigger, apply_mode) + r"\s+\d", text, re.I
            ):
                return
            pattern = _boundary_pattern(trigger, str(rr.get("apply", "")))
            if label_prefix == "unique" and trigger.upper() == "КШГ":
                # Table extraction may insert unit / quantity columns between
                # "КШГ" and its catalogue number. Confirm the structured BROEN
                # code anywhere in the same row block, then redact both pieces.
                code_pat = re.compile(r"(?<![\w])(?:70|71|79)\.[0-9]{3}\.[0-9]{2,3}(?:\.[0-9A-Za-zА-Яа-я]+){1,5}", re.I)
                if not code_pat.search(text):
                    return
                for m in re.finditer(r"(?<![\w])КШГ(?![\w])", text, re.I):
                    candidates.append((m.start(), m.end(), f"{label_prefix}:{trigger}"))
                for m in code_pat.finditer(text):
                    candidates.append((m.start(), m.end(), f"{label_prefix}:{trigger}:code"))
                return
            # PDF-only fallback for a unique brand: short uppercase suffixes
            # such as "МУЛЬТИПАЙП ПРО RC" are part of the trade designation,
            # while technical tokens (ГАЗ, DN, PN, SDR, ГОСТ, etc.) stop it.
            if label_prefix == "unique" and "бренд" in str(rr.get("type", "")).lower():
                pattern += (
                    r"(?-i:(?:\s+(?!(?:ГАЗ|ВОДА|DN|PN|IP|ГОСТ|ТУ|SDR|ПЭ|СТАЛЬ|СТ\.|RAL)\b)"
                    r"[A-ZА-ЯЁ]{2,8}){0,2})"
                )
            pat = re.compile(pattern, re.I)
            for m in pat.finditer(text):
                candidates.append((m.start(), m.end(), f"{label_prefix}:{trigger}"))

        # Confirmed global firm identifiers apply to both text PDFs and OCR scans.
        for rr in self.global_unique_rules:
            trigger = str(rr.get("trigger", "")).strip()
            rx = str(rr.get("regex", "")).strip()
            note = str(rr.get("note", "")).casefold()
            if "recovery" in note or "model family" in note:
                continue
            if trigger.casefold() == "камаз" and re.search(
                r"(?i)(?:двигатель\s*-?\s*|на\s+шасси\s+)КАМАЗ-\d", text
            ):
                continue
            if rx:
                try:
                    for m in re.finditer(rx, text, re.I):
                        candidates.append((m.start(), m.end(), f"global:{trigger or rx}"))
                except re.error:
                    pass
            else:
                add_rule(rr, "global")

        if inn:
            for rr in self._rules(resolved_factory, inn):
                add_rule(rr)
        else:
            for rr in self.pdf_unique_rules:
                add_rule(rr, "unique")

        if inn:
            for alias in self.aliases_by_inn.get(inn, []):
                if len(alias) < 3:
                    continue
                alias_norm = normalize_name(alias)
                if alias_norm in self.conservative_model_aliases:
                    continue
                if re.search(_alias_pattern(alias) + r"\s*-?\s*\d", text, re.I):
                    continue
                if re.search(r'["«]' + _alias_pattern(alias) + r'["»]', text, re.I):
                    continue
                pat = re.compile(_alias_pattern(alias), re.I)
                for m in pat.finditer(text):
                    candidates.append((m.start(), m.end(), f"alias:{alias}"))

        # Global-safe redactions. These rules do not depend on manufacturer.
        for label, pat in (
            ("reverse_tu", REVERSE_TU_RE),
            ("article", ARTICLE_RE),
            ("vo_tu", VO_TU_RE),
            ("attached_tu", ATTACHED_TU_RE),
            ("tu_wrap", TU_WRAP_RE),
            ("tu", TU_CODE_RE),
            ("standalone_tu", STANDALONE_TU_RE),
            ("contact", CONTACT_RE),
            ("supplier_org_inn", PDF_ORG_WITH_INN_RE),
            ("letter_tail", LETTER_TAIL_RE),
            ("request", REQUEST_RE),
        ):
            for m in pat.finditer(text):
                candidates.append((m.start(), m.end(), label))

        protected = []
        for pat in PDF_PROTECTED_RES:
            for m in pat.finditer(text):
                protected.append((m.start(), m.end()))
        protected.sort()

        # Longest / broadest candidates first, then split away KEEP fragments.
        candidates.sort(key=lambda x: (x[0], -(x[1] - x[0])))
        pieces: list[tuple[int, int, str]] = []
        for a, b, label in candidates:
            for x, y in _subtract_ranges(a, b, protected):
                # Whitespace-only ranges do not need a physical annotation.
                if text[x:y].strip():
                    pieces.append((x, y, label))

        # Merge overlapping pieces. Keep a composite label for diagnostics.
        pieces.sort(key=lambda x: (x[0], x[1]))
        merged: list[list] = []
        for a, b, label in pieces:
            if merged and a <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], b)
                if label not in merged[-1][2]:
                    merged[-1][2].append(label)
            else:
                merged.append([a, b, [label]])
        return [(a, b, ";".join(labels)) for a, b, labels in merged]

    def redaction_matches(self, text: str, code: str = "", factory: str = "") -> list[str]:
        """Backward-compatible exact substrings for non-coordinate callers."""
        value = str(text or "")
        return [value[a:b] for a, b, _ in self.redaction_spans(value, code, factory)]
