"""Explicit user decisions; these are not claims of catalog verification."""
import re

AERO_INN = '3257017280'
AERO_NAME_RE = re.compile(r'(?i)(?<!\w)(?:Аэро\s+Иксиа|Aero\s+IXIA|Иксиа)(?!\w)')
AERO_POLICY = {
    'id': 'aero_name_only_2026_09_10',
    'inn': AERO_INN,
    'decision': 'Удалять название АЭРО ИКСИА / Aero IXIA; сохранять обозначения оборудования.',
    'basis': 'Прямое решение пользователя от 10.09.2026, не подтверждение по каталогу.',
    'keep': ['AI-… / АИ-…', 'СКВ', 'BOX(S)', 'S-…', 'КАС-W / KAS-W / KAC-W',
             'CompactVolume', 'CrisperLine', 'RunAir', 'RunCool', 'RunRow'],
}
# Exact families covered by the decision. Other suppliers and arbitrary residual
# text do not become approved just because Aero IXIA occurs in the same row.
AERO_DESIGNATION_RE = re.compile(
    r'(?i)(?<!\w)(?:[AА][IИ]-[\w]+(?:[-./+*,][\w]+)*|'
    r'СКВ(?:\([\w]+\))?|BOX(?:\([\w]+\))?|S-\d+(?:[.,]\d+)?|'
    r'[КK][АA][СCS]-W(?:-[\w]+)*|CompactVolume|CrisperLine|RunAir|RunCool|RunRow)(?!\w)')


def aero_designation_ranges(text):
    return [m.span() for m in AERO_DESIGNATION_RE.finditer(text)]
