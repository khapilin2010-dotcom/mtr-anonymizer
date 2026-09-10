"""Expert layer over the unchanged legacy catalog. One snapshot per batch."""
from collections import Counter, defaultdict
import re

from excel.excel_engine import Anonymizer, protected_ranges, merge_ranges
from excel.knowledge import case_key, normalize, fragments
from excel.manufacturer_policy import AERO_INN, AERO_NAME_RE, aero_designation_ranges


def occurrences(text, fragment):
    return [m.span() for m in re.finditer(r'(?<!\w)' + re.escape(fragment) + r'(?!\w)', text)]


def protection_losses(source, final, base, code='', factory=''):
    spans = protected_ranges(source)
    _, identity = base.resolve_factory(code, factory)
    if identity == AERO_INN or AERO_NAME_RE.search(source):
        spans += aero_designation_ranges(source)
    values = Counter(source[a:b] for a, b in merge_ranges(spans))
    return [value for value, count in values.items() if final.count(value) < count]


class ExpertAnonymizer:
    def __init__(self, snapshot, base=None):
        self.snapshot = snapshot; self.base = base or Anonymizer()
        self.version = self.base.version
        self.component_index = defaultdict(list)
        for entries in snapshot.rules.values():
            for entry in entries:
                scope = entry['sample']['scope']
                if scope['role'] == 'комплектующее':
                    words = re.findall(r'\w+', scope['category'].casefold())
                    if words:
                        self.component_index[words[0]].append(entry)

    def anonymize(self, name, code='', factory=''):
        source = str(name or '')
        entry = self.snapshot.entries.get(case_key(code, source, factory))
        if entry and entry['status'] in ('ACTIVE', 'TRUSTED') and entry['value'].strip():
            if not protection_losses(source, entry['value'], self.base, code, factory):
                removed, _ = fragments(source, entry['value'])
                result = dict(text=entry['value'], factory=self.base.resolve_factory(code, factory)[0] or entry['sample'].get('factory', ''),
                              status='ЗЕЛЁНЫЙ', changed=entry['value'] != source,
                              reason='Проверенное решение; независимых инженеров: ' + str(entry['confirmations']),
                              removed=['Экспертное решение: ' + x for x in removed], trace=[], protected=[])
                self.describe(result, entry, 'Код Автодокс' if str(code).strip() else 'Точное наименование и завод')
                return result
        result = self.base.anonymize(source, code, factory)
        result['knowledge'] = {'version': self.snapshot.version, 'source': 'Статическая база', 'status': 'LEGACY', 'events': []}
        entry = self.snapshot.entries.get(case_key(code, source, factory))
        if entry and entry['status'] != 'DISABLED':
            self.describe(result, entry, 'Код Автодокс' if str(code).strip() else 'Точное наименование и завод')
            if entry['status'] == 'DISPUTED':
                return self.conflict(result, 'Инженеры сохранили разные решения; требуется проверка')
            if entry['status'] in ('ACTIVE', 'TRUSTED'):
                losses = protection_losses(source, entry['value'], self.base, code, factory)
                if losses:
                    result['knowledge']['proposed'] = entry['value']
                    return self.conflict(result, 'Решение удаляет защищённые признаки: ' + ', '.join(losses))
                if not entry['value'].strip():
                    return self.conflict(result, 'Экспертное наименование пустое')
                removed, restored = fragments(source, entry['value'])
                result.update(text=entry['value'], status='ЗЕЛЁНЫЙ', changed=entry['value'] != source,
                              removed=['Экспертное решение: ' + x for x in removed],
                              reason='Проверенное решение; независимых инженеров: ' + str(entry['confirmations']))
                return result
        # Factory must be resolved as the product factory. Aliases of components
        # elsewhere in the description do not enable a product-scoped rule.
        resolved, _ = self.base.resolve_factory(code, factory)
        scopes = {normalize(factory), normalize(resolved)} - {''}
        rules = [e for scope in scopes for e in self.snapshot.rules.get(scope, [])]
        component_entries = {e['key']: e for word in set(re.findall(r'\w+', source.casefold())) for e in self.component_index.get(word, [])}
        rules = list({e['key']: e for e in rules + list(component_entries.values())}.values())
        keeps, deletes, applied, conflicts = [], [], [], []
        for entry in rules:
            sample = entry['sample']; scope = sample['scope']
            # Literal category with boundaries; no fuzzy auto-application.
            if not occurrences(source.casefold(), scope['category'].casefold()):
                continue
            spans = occurrences(source, sample['fragment'])
            if scope['role'] == 'комплектующее':
                # Explicit component clauses only. The product's factory never
                # supplies missing evidence for a component manufacturer.
                contexts = []
                _, component_inn = self.base.resolve_factory('', scope['factory'])
                from excel.rules import normalize_name
                factory_name = normalize_name(scope['factory'])
                for clause in re.finditer(r'[^;\n]+', source):
                    content = clause.group()
                    marker = re.search(r'(?i)\b(?:привод\w*|электродвигател\w*|двигател\w*|комплектующ\w*|в\s+комплекте)\b', content)
                    if not marker:
                        continue
                    start = clause.start() + marker.start()
                    content = source[start:clause.end()]
                    attributed = (factory_name and factory_name in normalize_name(content)) or (component_inn and any(i == component_inn for _, _, i in self.base._aliases(content)))
                    if attributed and occurrences(content.casefold(), scope['category'].casefold()):
                        contexts.append((start, clause.end()))
                spans = [(a, b) for a, b in spans if any(x <= a and b <= y for x, y in contexts)]
            if not spans or entry['status'] == 'DISABLED':
                continue
            if entry['status'] == 'DISPUTED':
                conflicts.append('Спорное правило: ' + sample['fragment']); continue
            if entry['status'] not in ('ACTIVE', 'TRUSTED'):
                continue
            if entry['value'] == 'DELETE':
                protected = protected_ranges(source)
                _, identity = self.base.resolve_factory(code, factory)
                if identity == AERO_INN or AERO_NAME_RE.search(source):
                    protected += aero_designation_ranges(source)
                if any(a < y and x < b for a, b in spans for x, y in protected):
                    conflicts.append('DELETE пересекается с защитой: ' + sample['fragment']); continue
                deletes.extend((a, b, 'Экспертное DELETE: ' + entry['key']) for a, b in spans)
            else:
                keeps.extend(spans)
            applied.append(entry)
        if any(a < y and x < b for a, b, _ in deletes for x, y in keeps):
            conflicts.append('Одновременно применимы KEEP и DELETE')
            deletes = []  # Avoid a destructive resolution of a rule conflict.
        if applied:
            result = self.base.anonymize(source, code, factory, expert_keeps=keeps, expert_deletes=deletes)
            result['knowledge'] = {'version': self.snapshot.version, 'source': 'Контекстные правила инженеров',
                                   'status': 'ACTIVE', 'events': [h['id'] for e in applied for h in e['live']],
                                   'rules': [e['key'] for e in applied], 'confirmations': min(e['confirmations'] for e in applied)}
        if conflicts:
            return self.conflict(result, '; '.join(conflicts))
        return result

    def describe(self, result, entry, source):
        result['knowledge'] = {'version': self.snapshot.version, 'source': source, 'status': entry['status'],
                               'confirmations': entry['confirmations'], 'opposition': entry['opposition'],
                               'key': entry['key'], 'events': [e['id'] for e in entry['live']]}

    @staticmethod
    def conflict(result, reason):
        result.update(status='КРАСНЫЙ', reason=reason + '; автоматический безопасный результат оставлен для проверки')
        return result
