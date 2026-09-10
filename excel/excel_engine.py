"""Independent, literal string rules for MTR Excel.

Find immutable technical ranges first, subtract them from every deletion, and
build the output once. No document-format logic belongs in this module.
"""
from collections import defaultdict
from functools import lru_cache
import re

from common.database import load_database
from excel import rules as r
from excel.supplemental_rules import EXTRA_RULES, EXTRA_GLOBAL_RULES
from excel.reviewed_technical import reviewed_ranges

LEGAL = r'(?:ООО|ПАО|ОАО|ЗАО|АО|НПО|НПП|ФГУП)'
LEGAL_RE = re.compile(rf'(?i)(?<!\w){LEGAL}(?!\w)')
TECH_START = r'(?:ГОСТ|ТУ|IP\s*\d|DN\s*\d|PN\s*\d|Ex|УХЛ|сталь|Ст\.|давление|температура|напряжение|размер|труба|кабель|диаметр)'
QUOTED_ORG_RE = re.compile(
    rf'(?i)(?<!\w){LEGAL}(?:[.\s]+{LEGAL})*[.\s]*'
    rf'(?:"[^;\n]{{1,180}}?"(?=\s*(?:[,;.]|{TECH_START}|[а-яё-]+\s+филиал\b|$))|«[^»\n]+»|“[^”\n]+”)'
    r'(?:\s+[а-яё-]+\s+филиал\b)?')
ROLE_RE = re.compile(r'(?i)\b(?:по\s+технологии|завод|производитель|изготовитель|поставщик|производства|разработчик)\s*[:=–—-]?\s*$')
# Literal technical words that also occur as company aliases in the database.
# Preserve these without an explicit company attribution; no name classifier.
AMBIGUOUS_ALIASES = r.GENERIC_TEXT_ALIASES | {
    'прибор', 'канат', 'никель', 'сплав', 'сенсор', 'сила', 'контур',
    'пульс', 'ресурс', 'метиз', 'волна', 'вектор', 'логика', 'система', 'знак',
    'агрегат', 'блок-бокс', 'элемент', 'звезда', 'квт', 'счетчик',
    'контактор', 'полюс', 'вст', 'эхз', 'сбп',
}


def explicit_alias_context(text, start, end):
    """Ambiguous dictionary words need attribution in THIS occurrence.

    A registry match elsewhere must never turn an engineering noun into a brand.
    The factory field uses a separate trusted lookup path.
    """
    prefix = text[:start]
    return bool(re.search(rf'(?i)(?<!\w){LEGAL}\s*["«“]?\s*$', prefix)
                or ROLE_RE.search(prefix.rstrip('"«“ ')))

DOTTED_ORG_RE = re.compile(rf'(?<!\w){LEGAL}(?:\.{LEGAL})*\.[А-ЯЁA-Z]{{2,}}(?!\w)')
# A legal form followed by title-case words is an explicit company attribution.
# Stop before a technical term; never eat arbitrary lower-case description.
BARE_ORG_RE = re.compile(
    rf'(?<!\w){LEGAL}(?:\s+{LEGAL})*\s+'
    r'(?!(?:ГОСТ|ТУ|IP|DN|PN|Ex|УХЛ)\b)[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z0-9&-]*'
    r'(?:\s+(?!(?:ГОСТ|ТУ|IP|DN|PN|Ex|УХЛ)\b)[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z0-9&-]*){0,6}')
REGISTER_RE = re.compile(r'(?i)\bРеестр\s+МТР\s+(?:ПАО\s+)?Газпром\s*№\s*[\w./-]+')
LETTER_RE = re.compile(
    rf'(?i)(?:(?<!\w)(?:служебн\w*\s+)?письм\w*\s*)?(?:ВО\s*)?№\s*'
    rf'[\w][\w./–—-]{{0,79}}\s+от\s+'
    rf'(?:\d{{1,2}}\s+{r.MONTHS}\s+\d{{4}}|\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}})'
    r'(?:\s*г(?:ода)?\.?)?')
REVERSE_LETTER_RE = re.compile(
    rf'(?i)\b(?:служебн\w*\s+)?письм\w*\s+от\s+'
    rf'(?:\d{{1,2}}\s+{r.MONTHS}\s+\d{{4}}|\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}})'
    r'(?:\s*г\.?)?\s*№\s*[\w][\w./–—-]*')
TU_PREFIX_RE = re.compile(r'(?i)(?<!\w)(?:(?:по\s+(?:типу\s+)?|тип\s+))?ТУ(?![A-Za-zА-Яа-яЁё])\s*№?')
INN_RE = re.compile(r'(?i)(?<!\w)(?:ИНН|КПП)(?!\w)\s*[:№]?\s*\d*')
RESIDUAL_RE = re.compile(rf'(?i)(?<!\w)(?:ТУ|{LEGAL}|ИНН|КПП)(?!\w)')

# Complete designations, including suffixes after ОЛ; source spelling is retained.
OL_RE = re.compile(r'(?i)(?<!\w)[\w][\w./-]*(?:[.-]Т?ОЛК?\d*[\w./-]*)(?!\w)')
OL_NUMBER_RE = re.compile(r'(?i)\b(?:ОЛК?|опросн\w*\s+лист)\s*№?\s*[\w][\w./-]*')
GOST_RE = re.compile(r'(?i)\bГОСТ(?:\s+Р)?(?:\s+(?:ИСО|МЭК|ISO|IEC)(?:/\w+)?)?\s*\d[\d. /–—-]*\d|\bГОСТ\b')
TECH_RES = [
    GOST_RE, OL_RE, OL_NUMBER_RE,
    re.compile(r'(?i)(?<!\w)IP\s*\d{2}[A-Z]?(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:УХЛ|ХЛ|У|Т)\s*\d(?:\.\d)?(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:[012]\s*)?Ex[\sa-z]{0,24}?II[ABC]?\s*[TТ][1-6](?:\s*(?:Ga|Gb|Gc|Da|Db|Dc))?(?:\s*[XХU])?(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:DN|PN|SDR|RAL)\s*[-=]?\s*\d+(?:[.,]\d+)?(?!\w)'),
    re.compile(r'(?i)\b(?:сталь\s+(?:марки\s+)?|ст\.?\s*)[\d][\w.-]*'),
    # Retain the legacy conservative steel safeguard; exact documented
    # non-technical catalog codes are excluded in protected_ranges below.
    re.compile(r'(?<!\w)(?:\d{2}[ХГНМСТЮФВБДКР]\w*|\d[ХГНМСТЮФВБДКР]\w+)(?!\w)'),
    re.compile(r'(?i)(?<!\w)[+-]?\d+(?:[.,]\d+)?(?:\s*[xх×*]\s*\d+(?:[.,]\d+)?){1,3}(?:\s*мм)?'),
    re.compile(r'(?i)(?<![\w.])[+-]?\d+(?:[.,]\d+)?\s*(?:МПа|кПа|Па|бар|кВ|мВ|В|кВт|Вт|мм|см|км|м|мА|А|Гц|кг|г|мл|л|kV|mV|V|kW|W|mA|A|Hz|kg|mm|°\s*[CС]|град\.?\s*[CС])(?!\w)'),
    re.compile(r'(?i)\b(?:давление|размеры?|температура|напряжение|диаметр)\s*[:=]?\s*(?:от\s*)?[+-]?\d+(?:[.,]\d+)?(?:\s*°?\s*[CС])?(?:\s*до\s*[+-]?\d+(?:[.,]\d+)?)?\s*(?:МПа|кПа|бар|кВ|В|мм|°\s*[CС])?'),
]

# These are engineering values, including forms found inside model strings.
TECH_RES += [
    # Units may precede a value in parameter labels, not only follow a number.
    re.compile(r'(?i)\b(?:мощность|напряжение|ток|частота|масса|длина|ширина|высота|давление|параметр)'
               r'(?:\s+[а-яё]+){0,4}\s*,?\s*'
               r'(?:кВт|кВА|кВАр|МВт|Вт|кВ|В|мА|А|Гц|кг|мм|см|МПа|кПа|Па|бар)'
               r'(?=\s*[:=–—-]\s*\d)'),
    re.compile(r'(?i)(?<!\w)(?:Масса\s*[-:=]?\s*)?[+-]?\d+(?:[.,]\d+)?\s*т\.(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:Масса\s*[-:=]?\s*)?[+-]?\d+(?:[.,]\d+)?\s*т(?!\w)'),
    re.compile(r'(?i)(?<!\w)\d+(?:[.,]\d+)?\s*(?:кВА|кВАр|МВт|кА|Ач|м3/ч(?:ас)?|м³/ч|м3|м²|м2)(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:ОСТ|СТО|DIN|ISO|IEC|EN|ASTM)(?:\s+Р)?\s+\d[\w./–—-]*(?:\s+\d[\w./–—-]*)?'),
    re.compile(r'(?i)(?<!\w)(?:[012I]\s*)?[EЕeе][XХxх][\sa-zа-я]{0,24}?II[ABCАВС]?\s*[TТ][1-6](?:\s*(?:Ga|Gb|Gc|Da|Db|Dc))?(?:\s*[XХU])?(?!\w)'),
    re.compile(r'(?i)(?<!\w)II(?:[123]?[GD]|G[abc])\s*II[ABCАВС]\s*[TТ][1-6](?:\s*[XХU])?(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:Pt|Cu|Ni)\s*\d+(?!\w)'),
    re.compile(r'(?i)(?<!\w)\d+\s*[-–]?\s*(?:шт\.?|компл\.?|водный|гранная|х-проводная)(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:\d+[xх×])?(?:\d+(?:/\d+)*(?:G|M)?BASE-[A-Z0-9]+|\d+GE|[QS]*SFP(?:28|56|\+)?|RJ-?45|RS-?485|RS-?232|USB(?:\s*\d\.\d)?|Bluetooth|GPS|HPL-пластик|Multi-mode|Yellow/Green)(?!\w)'),
    re.compile(r'(?i)(?<!\w)Т\d+К\d+(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:[012I]\s*)?[EЕ][XХ]\s*(?:ia|ib|ic|da|db|dc|d|e|ma|mb|mc|ta|tb|tc)(?!\w)'),
    re.compile(r'(?i)(?<!\w)изм\.\s*\d+(?!\w)'),
    re.compile(r'(?i)(?<!\w)(?:[DLHSДЛНШ]\s*=\s*\d+(?:[.,]\d+)?(?:\s*мм)?|[DLHSДЛНШ]\s*\d+(?:[.,]\d+)?\s*мм|[DД]\d+(?:[.,]\d+)?)(?!\w)'),
]

# REVIEW is deliberately broader than DELETE. A model-shaped token cannot
# certify its own anonymity merely because a manufacturer was identified.
REVIEW_TOKEN_RE = re.compile(r'(?<!\w)[\w]+(?:[-./+][\w]+)*(?!\w)')
TECH_WORDS = set('SFP QSFP RJ UTP STP GE GbE STM Serial Ethernet LAN WAN AC DC FM LC SC SM MM DDM TX RX PoE Wi Fi IP DN PN SDR RAL ГОСТ ОСТ ТУ УХЛ ХЛ LED LCD PVC PE HDPE ПВХ ПНД ПЭ МПа кПа мм см кг кВт кВ Гц USB GPS Bluetooth'.casefold().split())
TECH_WORDS.update('ЗРА САУ ППУ ОЦ УКРМ АВР ДЭС СОПТ ОПН ЭХЗ КТП БКТП ТНВД АБС'.casefold().split())


def review_tokens(text):
    masked = list(text)
    for a, b in protected_ranges(text):
        masked[a:b] = ' ' * (b - a)
    tokens = []
    for match in REVIEW_TOKEN_RE.finditer(''.join(masked)):
        token = match.group()
        if token.casefold() in TECH_WORDS or not re.search(r'[A-Za-zА-Яа-яЁё]', token):
            continue
        if (re.search(r'\d', token) or
                re.search(r'[A-ZА-ЯЁ]{3,}', token) or re.search(r'[A-Za-z]{3,}', token)):
            tokens.append(token)
    # A partially masked drive name is still a model if its vendor prefix
    # survived (for example, no matching manufacturer scope).
    tokens.extend(m.group() for m in re.finditer(r'(?i)(?<!\w)(?:[МM][VВBS](?:220|24)|КВМ(?:15|20|25)|КВБ(?:12|17)|КВБУ(?:14|18|22))(?!\w)', text))
    return list(dict.fromkeys(tokens))


def merge_ranges(ranges):
    merged = []
    for a, b in sorted(ranges):
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(b, merged[-1][1]))
        else:
            merged.append((a, b))
    return merged


def protected_ranges(text):
    # In a calendar date, "2025 г." means year, not a mass in grams.
    dates = [m.span() for m in re.finditer(
        rf'(?i)\b(?:\d{{1,2}}\s+{r.MONTHS}\s+\d{{4}}|\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}})\s*г(?:ода)?\.?', text)]
    # Manufacturer manual identifies this complete string as a meter code.
    # Neither 4ТМ nor .02М inside it is a steel grade or a physical length.
    catalog_codes = [m.span() for m in re.finditer(
        r'(?i)(?<!\w)СЭТ-4ТМ(?:\.\d{2}[МM]?(?:\.\d{2})?)?(?!\w)', text)]
    ranges = [(m.start(), m.end()) for rx in TECH_RES for m in rx.finditer(text)
              if not any(a <= m.start() and m.end() <= b for a, b in dates + catalog_codes)]
    ranges.extend((a, b) for a, b, _ in reviewed_ranges(text))
    for phrase in r.OL_PHRASE_RE.finditer(text):
        # A complete phrase extends through its OL designation, even on new lines.
        end = OL_RE.search(text, phrase.end())
        if end and not LETTER_RE.search(text[phrase.end():end.start()]):
            ranges.append((phrase.start(), end.end()))
        else:
            ranges.append(phrase.span())
    return merge_ranges(ranges)


def subtract(a, b, keeps):
    pieces = [(a, b)]
    for x, y in keeps:
        pieces = [(p, q) for s, e in pieces
                  for p, q in ((s, min(e, x)), (max(s, y), e)) if q > p] if any(
                      s < y and x < e for s, e in pieces) else pieces
    return pieces


class Anonymizer:
    def __init__(self, data_file=None):
        data = load_database(data_file)
        self.version = data['version']
        self.registry = data['registry']
        self.names = {}
        self.alias_index = defaultdict(list)
        self.rules_by_inn = defaultdict(list)
        self.rules_by_name = defaultdict(list)
        self.global_rules = []
        self.global_manufacturers = []
        rows = list(data['rules']) + list(EXTRA_RULES if data_file is None else ())
        self.rule_count = len(rows)
        alias_rows = list(data['aliases']) + [dict(row, alias=row['manufacturer']) for row in rows]
        for row in alias_rows:
            alias = r.normalize_name(row['alias'])
            inn = str(row.get('inn') or '') or 'name:' + r.normalize_name(row.get('manufacturer', ''))
            self.names.setdefault(inn, row.get('manufacturer', ''))
            if len(alias) < 3:
                continue
            words = re.findall(r'\w+', alias)
            if words:
                pattern = r._alias_pattern(alias).replace(r'\ ', r'[\s"«»“”]+')
                item = (alias, inn, re.compile(pattern, re.I))
                if not any(a == alias and i == inn for a, i, _ in self.alias_index[words[0]]):
                    self.alias_index[words[0]].append(item)
        for row in rows:
            pattern = self._compile_rule(row)
            if pattern:
                key = str(row.get('inn') or '') or 'name:' + r.normalize_name(row.get('manufacturer', ''))
                self.rules_by_inn[key].append(pattern)
                self.rules_by_name[r.normalize_name(row.get('manufacturer', ''))].append(pattern)
        for row in list(data['global_unique_rules']) + list(EXTRA_GLOBAL_RULES if data_file is None else ()):
            pattern = self._compile_rule(row)
            if pattern:
                self.global_rules.append(pattern)
                if row.get('manufacturer'):
                    self.global_manufacturers.append((pattern, row['manufacturer']))

    @staticmethod
    def _compile_rule(row):
        trigger = str(row.get('trigger', '')).strip()
        pattern = row.get('regex') or (r._boundary_pattern(trigger, row.get('apply', 'Точное совпадение')) if trigger else '')
        return re.compile(pattern, re.I) if pattern else None

    @lru_cache(maxsize=16384)
    def resolve_factory(self, code='', factory=''):
        for candidate in r.code_candidates(code):
            if self.registry.get(candidate):
                factory = self.registry[candidate]
                break
        inn = r.extract_inn(factory)
        if inn and inn in self.names:
            return factory, inn
        matches = self._aliases(factory, factory_field=True)
        if matches:
            longest = matches[0][1] - matches[0][0]
            identities = {i for a, b, i in matches if b - a == longest}
            if len(identities) == 1:
                key = next(iter(identities))
                return factory or self.names.get(key, ''), key
        return factory, ''

    def _aliases(self, text, factory_field=False):
        # Match against original characters to retain exact deletion offsets.
        found = []
        words = set(re.findall(r'\w+', text.lower().replace('ё', 'е')))
        for word in words:
            for alias, inn, pattern in self.alias_index.get(word, []):
                for m in pattern.finditer(text):
                    if (alias in AMBIGUOUS_ALIASES and not factory_field
                            and not explicit_alias_context(text, m.start(), m.end())):
                        continue
                    found.append((m.start(), m.end(), inn))
        return sorted(found, key=lambda x: -(x[1] - x[0]))

    def _candidates(self, text, code='', factory=''):
        resolved, inn = self.resolve_factory(code, factory)
        aliases = self._aliases(text)
        inns = {i for _, _, i in aliases if i}
        if inn:
            inns.add(inn)
        manufacturers = [resolved] if resolved else []
        manufacturers += [self.names[i] for i in sorted(inns) if self.names.get(i)]
        candidates = []
        for pattern, manufacturer in self.global_manufacturers:
            for match in pattern.finditer(text):
                manufacturers.append(manufacturer)
                role = ROLE_RE.search(text[:match.start()])
                if role:
                    candidates.append((role.start(), match.start(), 'обозначение производителя'))
        for a, b, _ in aliases:
            candidates.append((a, b, 'производитель'))
            # Extend confirmed aliases only through attached designation parts.
            # Brackets, commas and spaces terminate this expansion.
            tail = re.match(r'(?:-[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._/+\-]*)', text[b:])
            if tail:
                candidates.append((a, b + tail.end(), 'фирменное обозначение'))
        patterns = [
            ('производитель', QUOTED_ORG_RE), ('производитель', DOTTED_ORG_RE),
            ('производитель', BARE_ORG_RE), ('реестр', REGISTER_RE),
            ('реквизиты письма', LETTER_RE), ('реквизиты письма', REVERSE_LETTER_RE),
            ('реквизиты письма', r.LETTER_TAIL_RE), ('ИНН/КПП', INN_RE),
            ('ТУ', r.REVERSE_TU_RE), ('ТУ', r.VO_TU_RE), ('ТУ', r.ATTACHED_TU_RE),
            ('ТУ', r.TU_WRAP_RE), ('ТУ', r.TU_CODE_RE), ('ТУ', TU_PREFIX_RE),
            ('контакт', r.CONTACT_RE), ('бренд', r.CONFIRMED_GLOBAL_BRAND_RE),
        ]
        # TU № has the same code grammar as TU; preserve original offsets.
        for m in re.finditer(r'(?i)(?<!\w)ТУ\s*№\s*', text):
            tail = r.TU_CODE_RE.match('ТУ ' + text[m.end():])
            if tail:
                candidates.append((m.start(), m.end() + tail.end() - 3, 'ТУ'))
        for label, rx in patterns:
            for m in rx.finditer(text):
                candidates.append((*m.span(), label))
                if label in ('производитель', 'бренд'):
                    manufacturers.append(m.group().strip())
        # Remove legal-form remnants too, but record uncertain unquoted names.
        uncertain_org = any(not any(a <= m.start() and m.end() <= b and label == 'производитель'
                                     for a, b, label in candidates)
                            for m in LEGAL_RE.finditer(text))
        for a, b, label in list(candidates):
            if label in ('производитель', 'бренд'):
                role = ROLE_RE.search(text[:a])
                if role:
                    candidates.append((role.start(), a, 'обозначение производителя'))
        candidates.extend((*m.span(), 'юридическая форма') for m in LEGAL_RE.finditer(text))
        active = list(self.global_rules)
        for i in inns:
            active.extend(self.rules_by_inn.get(i, []))
        active.extend(self.rules_by_name.get(r.normalize_name(resolved), []))
        for rx in active:
            candidates.extend((*m.span(), 'признак из базы') for m in rx.finditer(text))
        return candidates, list(dict.fromkeys(filter(None, manufacturers))), uncertain_org

    def anonymize(self, name, code='', factory=''):
        original = str(name or '')
        keeps = protected_ranges(original)
        candidates, manufacturers, uncertain = self._candidates(original, code, factory)
        deleted = merge_ranges([piece for a, b, _ in candidates for piece in subtract(a, b, keeps)])
        # Keep characters never enter the removal log or a delete range.
        removed = []
        for a, b in deleted:
            fragment = original[a:b]
            if fragment.strip():
                labels = sorted({label for x, y, label in candidates if x < b and a < y})
                removed.append(f'{" / ".join(labels)}: {fragment}')
        # Clean punctuation only outside absolute KEEP spans.
        segments = []
        events = sorted([(a, b, '') for a, b in deleted] + [(a, b, original[a:b]) for a, b in keeps])
        pos = 0
        protected_values = []
        rendered_values = {}
        prefix_lengths = {'veza_actuator_voltage': 2, 'eridan_hose_diameter': 3,
                          'eridan_armour_diameter': 3, 'eridan_double_seal_diameter': 4}
        for a, b, row in reviewed_ranges(original):
            key = row['id']
            if key not in prefix_lengths or not any(x <= a - prefix_lengths[key] and y >= a for x, y in deleted):
                continue
            value = original[a:b]
            if key == 'veza_actuator_voltage':
                rendered_values[a, b] = value + ' В'
            elif key == 'eridan_hose_diameter':
                rendered_values[a, b] = 'кабельный ввод под металлорукав Ду' + value + ' мм'
            elif key == 'eridan_armour_diameter':
                rendered_values[a, b] = 'кабельный ввод для брони диаметром до ' + value + ' мм'
            else:
                lower = {'14': '10', '18': '14', '22': '18'}[value]
                rendered_values[a, b] = 'кабельный ввод с двойным уплотнением, диаметр кабеля ' + lower + '–' + value + ' мм'
        for a, b, value in events:
            if a < pos:
                raise AssertionError('Пересечение KEEP и DELETE')
            segments.append(original[pos:a])
            if value:
                marker = '\ue000' + str(len(protected_values)) + '\ue001'
                while marker in original:
                    marker = '\ue000' + marker
                protected_values.append((marker, rendered_values.get((a, b), value)))
                segments.append(marker)
            else:
                segments.append(' ')
            pos = b
        segments.append(original[pos:])
        text = ''.join(segments)
        text = re.sub(r'\[\s*\]|\(\s*\)|["«»]\s*["«»]', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip(' ,;')
        text = re.sub(r'\s+([,;])', r'\1', text)
        text = re.sub(r'([,;])(?:\s*[,;])+', r'\1', text)
        text = re.sub(r'[,;]\s*\.', '.', text)
        for marker, value in protected_values:
            text = text.replace(marker, value)
        residuals, _, _ = self._candidates(text, code, factory)
        failures = []
        # A documented technical KEEP can settle a collision. Unreviewed
        # overlaps (including brands inside an OL) still require inspection.
        reviewed_original = merge_ranges([(a, b) for a, b, _ in reviewed_ranges(original)])
        reviewed_output = merge_ranges([(a, b) for a, b, _ in reviewed_ranges(text)])
        residual_unknown = any(subtract(a, b, reviewed_output) for a, b, _ in residuals)
        collision_unknown = any(
            subtract(max(a, x), min(b, y), reviewed_original)
            for a, b, _ in candidates for x, y in keeps if a < y and x < b)
        if residual_unknown or RESIDUAL_RE.search(text):
            failures.append('Остался удаляемый признак; возможно пересечение с абсолютным KEEP')
        elif collision_unknown:
            failures.append('Часть удаляемого обозначения сохранена как абсолютный KEEP; проверить остаток')
        if uncertain:
            failures.append('Проверить границы названия организации')
        if any(original[a:b] not in text for a, b in keeps):
            raise AssertionError('Потеря абсолютного KEEP')
        if not text.strip():
            failures.append('Наименование стало пустым')
        if not manufacturers:
            failures.append('Производитель не определён')
        unresolved = review_tokens(text)
        if unresolved:
            failures.append('Проверить обозначения: ' + ', '.join(unresolved[:12]))
        _, identity = self.resolve_factory(code, factory)
        if manufacturers and not identity and not removed:
            failures.append('Нет подтверждённых правил для производителя')
        if text != original and not removed:
            # Whitespace-only formatting is not anonymization.
            text = original
        return {'text': text, 'factory': '; '.join(manufacturers),
                'status': 'ЖЁЛТЫЙ' if failures else 'ЗЕЛЁНЫЙ',
                'reason': '; '.join(failures), 'removed': removed,
                'changed': text != original}
