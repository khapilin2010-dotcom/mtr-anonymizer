"""Reviewed selection import: inspect first, then explicitly resolve conflicts."""
from collections import defaultdict
from pathlib import Path
import uuid

from excel.excel_engine import Anonymizer
from excel.expert_engine import protection_losses
from excel.knowledge import case_key, digest, file_hash, normalize, contextual_events, Snapshot
from excel.knowledge_edit_io import revision
from excel.mapped_import import _iter_mapped, guess_mapping


def best_candidate(candidates):
    def score(candidate):
        mapping = guess_mapping(candidate['headers'])
        return (int(mapping['source'] is not None and mapping['final'] is not None
                    and mapping['source'] != mapping['final']),
                sum(value is not None for value in mapping.values()), -candidate['row'])
    return max(range(len(candidates)), key=lambda i: score(candidates[i]))


def reject_xls_formulas(path):
    if Path(path).suffix.lower() != '.xls':
        return
    import struct
    from xlrd.compdoc import CompDoc
    doc = CompDoc(Path(path).read_bytes())
    stream = doc.get_named_stream('Workbook') or doc.get_named_stream('Book')
    position = 0
    while stream and position + 4 <= len(stream):
        kind, length = struct.unpack_from('<HH', stream, position)
        if kind in (0x0006, 0x0206, 0x0406):
            raise ValueError('В XLS есть формулы. Сохраните выборку как XLSX и повторите загрузку.')
        position += length + 4


def preview_selection(path, store, candidate, mapping):
    path = Path(path)
    indices = [value for value in mapping.values() if value is not None]
    if mapping.get('source') is None or mapping.get('final') is None:
        raise ValueError('Укажите исходное и обезличенное наименование. Одного кода недостаточно.')
    if len(set(indices)) != len(indices):
        raise ValueError('Для разных полей выберите разные столбцы.')
    if any(type(i) is not int or i < 0 or i >= len(candidate['headers']) for i in indices):
        raise ValueError('Выбранного столбца нет в таблице.')
    before = file_hash(path)
    reject_xls_formulas(path)
    store.sync(auto_backup=False)
    snapshot = store.snapshot()
    base = Anonymizer()
    grouped = defaultdict(list)
    issues = []
    rows = 0
    for sheet, number, values, formulas in _iter_mapped(path, candidate, mapping):
        def get(field):
            i = mapping.get(field)
            return '' if i is None or i >= len(values) or values[i] is None else str(values[i]).strip()
        source, final = get('source'), get('final')
        if not any(get(field) for field in mapping):
            continue
        rows += 1
        if set(indices) & formulas:
            issues.append((sheet, number, 'Формула или ошибка в наименовании, коде или заводе'))
            continue
        if not source or not final:
            issues.append((sheet, number, 'Нужны исходное и обезличенное наименования'))
            continue
        if all(get(field) == str(candidate['headers'][i]).strip() for field, i in mapping.items() if i is not None):
            issues.append((sheet, number, 'Повтор строки заголовков'))
            continue
        code, factory = get('code'), get('factory')
        losses = protection_losses(source, final, base, code, factory)
        if losses:
            issues.append((sheet, number, 'Потеряны защищённые признаки: ' + ', '.join(losses)))
            continue
        key = case_key(code, source, factory)
        grouped[key].append(dict(key=key, sheet=sheet, row=number, code=code,
                                 factory=factory, source=source, final=final))
    if not rows:
        raise ValueError('На выбранном листе после заголовков нет данных.')
    items, repeated = [], 0
    for key, copies in grouped.items():
        if len({normalize(item['final']) for item in copies}) != 1:
            issues.extend((item['sheet'], item['row'], 'В выборке разные результаты для одного и того же случая') for item in copies)
            continue
        item = copies[0]
        repeated += len(copies) - 1
        entry = snapshot.entries.get(key)
        if entry and entry['status'] in ('ACTIVE', 'TRUSTED') and normalize(entry['value']) == normalize(item['final']):
            repeated += 1
            continue
        items.append(dict(item, conflict=bool(entry), current=entry['value'] if entry else '',
                          current_status=entry['status'] if entry else '',
                          expected_revision=revision(entry) if entry else None))
    if file_hash(path) != before:
        raise ValueError('Выборка изменилась во время чтения. Повторите проверку.')
    return dict(path=str(path.resolve()), file_hash=before, store_id=store.store_id,
                version=snapshot.version, rows=rows, ready=len(items),
                new=sum(not item['conflict'] for item in items), repeated=repeated,
                conflicts=sum(item['conflict'] for item in items), issues=issues,
                plan=items, fingerprint=digest(items))


def commit_selection(path, store, plan, approved_conflicts=()):
    """Apply exactly the inspected plan; never silently reactivate existing cases."""
    if store.store_id != plan['store_id']:
        raise ValueError('Выбрана другая база. Повторите проверку выборки.')
    if str(Path(path).resolve()) != plan['path'] or file_hash(path) != plan['file_hash']:
        raise ValueError('Выборка изменилась после проверки. Нажмите «Проверить выборку» заново.')
    if digest(plan['plan']) != plan['fingerprint']:
        raise ValueError('Состав импорта изменился. Повторите проверку.')
    approved = set(approved_conflicts)
    if approved - {item['key'] for item in plan['plan'] if item['conflict']}:
        raise ValueError('Неизвестное подтверждение конфликта.')
    store.sync(auto_backup=False)
    snapshot = store.snapshot()
    if snapshot.version != plan['version']:
        raise ValueError('База изменилась после проверки. Нажмите «Проверить выборку» заново.')
    accepted = skipped = 0
    issues = list(plan['issues'])
    batch = 'selection-import:' + uuid.uuid4().hex
    controls = defaultdict(list)
    for event in contextual_events(snapshot.observed_events):
        if event['kind'] == 'control' and not event.get('usage'):
            controls[event['key']].append(event)
    for item in plan['plan']:
        if item['conflict'] and item['key'] not in approved:
            skipped += 1
            continue
        current = snapshot.entries.get(item['key'])
        record = dict(id=str(uuid.uuid4()), session=batch, batch_id=batch,
                      code=item['code'], source=item['source'], factory=item['factory'],
                      automatic=item['source'], classification=None, explicit_feedback=True,
                      provenance=dict(source='Загрузка ранее обезличенной выборки',
                                      file=Path(path).name, sheet=item['sheet'], row=item['row'],
                                      file_sha256=plan['file_hash'], reviewed=True))
        previous = current['history'] if current else []
        payload = store.decision_payload(record, item['final'], 'Импорт проверенной выборки', previous)
        store.emit(payload)
        if current and current['status'] == 'DISABLED':
            prior_controls = controls[item['key']]
            live = Snapshot.live(prior_controls, cross_user=True)
            inherited = ({k: live[0][k] for k in ('classification', 'scope', 'protected', 'settings', 'decision_override')
                          if k in live[0]} if len(live) == 1 else {})
            store.emit(dict(inherited, kind='control', key=item['key'], value='ENABLED',
                            requested_status=None, previous_status='DISABLED',
                            reason='Инженер подтвердил вариант проверенной выборки',
                            supersedes=[event['id'] for event in prior_controls]))
        accepted += 1
    sync = store.sync(auto_backup=False)
    after = store.snapshot()
    for item in plan['plan']:
        if item['conflict'] and item['key'] not in approved:
            continue
        entry = after.entries.get(item['key'])
        if entry and entry['status'] not in ('ACTIVE', 'TRUSTED'):
            issues.append((item['sheet'], item['row'], 'Есть одновременное решение другого инженера; требуется проверка'))
    return dict(accepted=accepted, repeated=plan['repeated'], skipped_conflicts=skipped,
                issues=issues, sync=sync)
