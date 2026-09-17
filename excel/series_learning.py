"""Scoped KEEP learning from explicit corrections, including old case journals.

No fuzzy whole-row replacements and no learned automatic DELETE. Series rules
are derived from live exact decisions and can be disabled through base export.
"""
from collections import defaultdict
import re

from excel.knowledge import contextual_events, digest, Snapshot
from excel.semantics import folded, factory_key, features, words, STOP

LETTER = 'A-Za-zА-Яа-яЁё'
DASH = '-‐‑‒–—−'
SERIES = re.compile(r'(?<!\w)([' + LETTER + r']{2,}(?:[' + DASH + r'][' + LETTER + r']{2,})*)(?=$|[^' + LETTER + r'])')
EXCLUDE = {'ООО', 'АО', 'ПАО', 'ОАО', 'ЗАО', 'НПО', 'ГОСТ', 'ТУ', 'DN', 'PN', 'IP', 'УХЛ', 'ОЛ'}


def families(source, require_upper=True):
    found = {}
    for match in SERIES.finditer(source):
        value = match.group()
        letters = re.sub('[^' + LETTER + ']', '', value)
        if not 2 <= len(letters) <= 20 or (require_upper and letters.upper() != letters) or value.upper() in EXCLUDE:
            continue
        found.setdefault(folded(value), value)
    return found


def family_spans(text, family, whole_model=False):
    escaped = r'[' + DASH + r']'
    pattern = escaped.join(re.escape(part) for part in folded(family).split('-'))
    suffix = r'(?:[' + DASH + r'./]?[0-9][' + DASH + r'\w./]*)?' if whole_model else ''
    return [m.span() for m in re.finditer(r'(?<!\w)' + pattern + r'(?![' + LETTER + r'])' + suffix, text, re.I)]


def scoped_occurrences(source, family, info):
    """Keep product and component decisions in separate, literal contexts."""
    result = defaultdict(list)
    scopes = {}
    for start, end in family_spans(source, family, whole_model=True):
        component = next((c for c in info.get('components', []) if c['start'] <= start < c['end']), None)
        context = component or info
        origin = component['start'] if component else 0
        head = next((w for w in words(source[origin:start]) if w not in STOP and w.isalpha()), '')
        category = context.get('category', 'не определено')
        if not head or category == 'не определено':
            continue
        factory = context.get('factory', '')
        scope = dict(factory=factory or 'Не указан', unknown_factory=not bool(factory),
                     category=category, role='комплектующее' if component else 'изделие', head=head)
        identity = digest([factory_key(factory), category, scope['role'], head])
        scopes[identity] = scope
        result[identity].append((start, end))
    return [(scopes[key], spans) for key, spans in result.items()]


def derived_events(events, admins):
    events = contextual_events(events)
    cases, controls = defaultdict(list), defaultdict(list)
    for event in events:
        if event['kind'] == 'case':
            cases[event['key']].append(event)
        elif event['kind'] == 'control' and event['user'] in admins:
            controls[event['key']].append(event)
    derived = []
    for key, history in cases.items():
        control = Snapshot.live(controls[key], cross_user=True)
        if any(e['value'] == 'DISABLED' for e in control):
            continue
        live = Snapshot.live(history, cross_user=True)
        # Disputed exact rows cannot supply a general rule in either direction.
        if len({folded(e['value']) for e in live}) != 1:
            continue
        for event in live:
            source, final = event['source'], event['value']
            automatic = event.get('automatic', source)
            info = event.get('classification')
            if not isinstance(info, dict):
                info = features(source, event.get('factory', ''))
            for normalized, family in families(source).items():
                count = len(family_spans(source, family))
                final_count = len(family_spans(final, family))
                auto_count = len(family_spans(automatic, family))
                if final_count == count and auto_count < count:
                    action = 'KEEP'
                elif final_count == 0:
                    action = 'DELETE'
                else:
                    continue
                scopes = scoped_occurrences(source, family, info)
                # If the same designation spans product and component contexts,
                # require a separate correction before generalizing it.
                if len(scopes) != 1:
                    continue
                scope, _ = scopes[0]
                rule_key = 'series:' + digest([normalized, factory_key(scope['factory']),
                                              scope['category'], scope['role'], scope['head'], scope['unknown_factory']])
                derived.append(dict(event, id=event['id'] + ':series:' + digest(rule_key)[:12],
                                    key=rule_key, kind='rule', value=action, fragment=family,
                                    scope=scope, series_rule=True, origin_event=event['id'],
                                    classification='серия', example=source, supersedes=[]))
    return derived


def matching_series(snapshot, source, info):
    matches = []
    for family_key in families(source, require_upper=False):
      for entry in getattr(snapshot, 'series_index', {}).get(family_key, []):
        sample = entry['sample']
        if not sample.get('series_rule') or entry['status'] == 'DISABLED':
            continue
        family = sample['fragment']
        for scope, spans in scoped_occurrences(source, family, info):
            expected = sample['scope']
            if (scope.get('unknown_factory') == expected.get('unknown_factory')
                    and factory_key(scope['factory']) == factory_key(expected['factory'])
                    and all(scope[k] == expected.get(k) for k in ('category', 'role', 'head'))):
                matches.append((entry, spans))
    return matches
