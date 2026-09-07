"""String patterns extracted from PR #4; no legacy engine import."""
import re

TOKEN_TAIL = r'(?:[^\s,;]|,(?=\S)|;(?=\S))*'

GENERIC_TEXT_ALIASES = {"оборудование", "мониторинг", "универсал", "источник", "монитор", "металлорукав", "контакт", "переход"}

CONFIRMED_GLOBAL_BRANDS = (
    "Унипол", "Гиперфлоу", "Метран", "Вэлан", "Ризур", "Рубеж",
    "Пензтяжпромарматура", "Волжский трубный завод",
    "Армтел", "ИКСИА", "Аэро Иксиа", "Теплолюкс", "Аналитика",
    "Эрис", "Феррум", "Телта", "Тропик", "PERCo", "Durcon", "TOFF",
    "НПО ФСА", "ООО НПО ФСА", "НЗСП", "ООО НЗСП",
)

CONFIRMED_GLOBAL_BRAND_RE = re.compile(
    r"(?i)(?<![\w])(?:" + "|".join(re.escape(x) for x in CONFIRMED_GLOBAL_BRANDS)
    + r")(?:-[A-Za-zА-Яа-яЁё0-9._/]+)?(?![\w])"
)

OL_PHRASE_RE = re.compile(r'(?i)\bКомплектация\s+по\s+обосновывающему\s+документу\b')

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

MONTHS = r"(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"

LETTER_TAIL_RE = re.compile(
    rf"(?i)(?<![\w])(?:ВО\s*)?№\s*[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._/\-–—]{{2,50}}"
    rf"\s+от\s+(?:\d{{1,2}}\s+{MONTHS}\s+\d{{4}}|\d{{1,2}}[.]\d{{1,2}}[.]\d{{2,4}})\s*г?[.]?"
    r"(?:\s*,?\s*(?:п|поз)[.]?\s*\d+(?:[.]\d+)*)?"
)

REVERSE_TU_RE = re.compile(
    r"\(\s*[A-Za-zА-Яа-я0-9][A-Za-zА-Яа-я0-9._/\- ]{2,}\s+ТУ\s*\)", re.I
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
    # Source tables may split one TU number with spaces around a
    # hyphen or between its numeric groups. Consume only further code-like
    # numeric groups, never following descriptive text or a protected OL tail.
    r"(?:\s*(?:[-–—]\s*)?(?=\d[0-9.\-/–—]*\b)\d[0-9.\-/–—]*)*", re.I
)

CONTACT_RE = re.compile(
    r"(?i)(?:\bИНН\s*\d{10,12}\b|\bКПП\s*\d{9}\b|"
    r"\b[\w.+-]+@[\w.-]+\.[A-Za-zА-Яа-я]{2,}\b|"
    r"https?://\S+|www\.\S+)"
)
