"""Controlled Excel export/import for knowledge maintenance.

The exported workbook is intentionally not the database itself. Engineers edit
only explicit management columns; import validates the current key/value before
creating a new immutable event. Audit history is preserved.
"""
from pathlib import Path
import uuid


HEADERS = [
    'Ключ', 'Тип', 'Статус', 'Текущее решение', 'Новое решение', 'Действие',
    'Независимых инженеров', 'Область', 'Пример', 'Причина изменения'
]
ACTION_ALIASES = {
    'CHANGE': 'CHANGE', 'ИЗМЕНИТЬ': 'CHANGE',
    'DISABLE': 'DISABLE', 'ОТКЛЮЧИТЬ': 'DISABLE',
    'ENABLE': 'ENABLE', 'ВКЛЮЧИТЬ': 'ENABLE',
}


def export_editable(path, store):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    store.sync()
    snapshot = store.snapshot()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Управление'
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1767A6')
        cell.alignment = Alignment(wrap_text=True)
    type_names = {'case': 'Точное решение', 'rule': 'Общее правило', 'control': 'Настройка'}
    for key, entry in snapshot.entries.items():
        sample = entry.get('sample', {})
        scope = sample.get('scope', {})
        example = sample.get('source', sample.get('example', ''))
        ws.append([key, type_names.get(entry.get('kind', ''), entry.get('kind', '')),
                   entry.get('status', ''), entry.get('value', ''), '', '',
                   entry.get('confirmations', 0), str(scope), example, ''])
    validation = DataValidation(
        type='list', formula1='"ИЗМЕНИТЬ,ОТКЛЮЧИТЬ,ВКЛЮЧИТЬ"', allow_blank=True)
    validation.errorTitle = 'Действие'
    validation.error = 'Выберите действие из списка.'
    validation.showErrorMessage = True
    ws.add_data_validation(validation)
    if ws.max_row >= 2:
        validation.add(f'F2:F{ws.max_row}')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    widths = [44, 18, 14, 55, 55, 16, 20, 38, 70, 55]
    for index, width in enumerate(widths, 1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(index)].width = width

    note = wb.create_sheet('Инструкция')
    note.append(['Как исправить базу решений'])
    note.append(['1. Найдите нужную строку по примеру, коду/наименованию или текущему решению.'])
    note.append(['2. Чтобы заменить решение: впишите правильный текст в «Новое решение», выберите «ИЗМЕНИТЬ» и укажите причину.'])
    note.append(['3. Если ошибочное правило вообще не должно работать: выберите «ОТКЛЮЧИТЬ» и укажите причину.'])
    note.append(['4. Чтобы вернуть ранее отключённую запись: выберите «ВКЛЮЧИТЬ» и укажите причину.'])
    note.append(['5. Не меняйте «Ключ» и «Текущее решение». Строки без действия при загрузке игнорируются.'])
    note.append(['6. Для общего правила новое решение при ИЗМЕНИТЬ должно быть KEEP или DELETE.'])
    note.append(['7. История не стирается: исправление записывается новым событием, старое остаётся в аудите.'])
    note.column_dimensions['A'].width = 120

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return path


def preview_editable(path, store):
    from openpyxl import load_workbook
    store.sync()
    snapshot = store.snapshot()
    wb = load_workbook(path, read_only=True, data_only=False)
    issues = []
    changes = []
    try:
        if 'Управление' not in wb.sheetnames:
            raise ValueError('Нет листа «Управление». Используйте выгрузку этой программы.')
        ws = wb['Управление']
        header = [str(c.value or '').strip() for c in next(ws.iter_rows())]
        missing = [name for name in HEADERS if name not in header]
        if missing:
            raise ValueError('Не хватает столбцов: ' + ', '.join(missing))
        pos = {name: header.index(name) for name in HEADERS}
        for row_no, cells in enumerate(ws.iter_rows(min_row=2), 2):
            values = [c.value for c in cells]
            raw_action = str(values[pos['Действие']] or '').strip().upper()
            if not raw_action:
                continue
            action = ACTION_ALIASES.get(raw_action)
            key = str(values[pos['Ключ']] or '').strip()
            current_file = str(values[pos['Текущее решение']] or '')
            new_value = str(values[pos['Новое решение']] or '')
            reason = str(values[pos['Причина изменения']] or '').strip()
            if not action:
                issues.append((row_no, 'Неизвестное действие: ' + raw_action)); continue
            entry = snapshot.entries.get(key)
            if not entry:
                issues.append((row_no, 'Ключ отсутствует в текущей базе')); continue
            if str(entry.get('value', '')) != current_file:
                issues.append((row_no, 'База изменилась после выгрузки; строка устарела')); continue
            if not reason:
                issues.append((row_no, 'Не указана причина изменения')); continue
            if action == 'CHANGE':
                if not new_value.strip():
                    issues.append((row_no, 'Для «ИЗМЕНИТЬ» заполните новое решение')); continue
                if entry.get('kind') == 'rule' and new_value.strip().upper() not in ('KEEP', 'DELETE'):
                    issues.append((row_no, 'Для общего правила допустимы только KEEP или DELETE')); continue
            changes.append(dict(row=row_no, key=key, action=action,
                                new_value=new_value, reason=reason, entry=entry))
    finally:
        wb.close()
    return dict(ready=len(changes), issues=issues, changes=changes)


def commit_editable(path, store):
    plan = preview_editable(path, store)
    applied = 0
    for item in plan['changes']:
        entry = item['entry']
        action = item['action']
        reason = item['reason']
        if action in ('DISABLE', 'ENABLE'):
            store.control(item['key'], action == 'DISABLE', reason)
            applied += 1
            continue
        sample = entry.get('sample', {})
        if entry.get('kind') == 'case':
            row = dict(id=sample.get('row_id') or str(uuid.uuid4()),
                       session=sample.get('session') or ('knowledge-edit:' + uuid.uuid4().hex),
                       code=sample.get('code', ''), source=sample.get('source', ''),
                       factory=sample.get('factory', ''),
                       automatic=sample.get('automatic', sample.get('source', '')),
                       classification=sample.get('classification'), explicit_feedback=True,
                       provenance={'source': 'Редактирование базы через Excel', 'reason': reason})
            store.decide(row, item['new_value'], 'Редактирование базы Excel')
            applied += 1
        elif entry.get('kind') == 'rule':
            scope = sample.get('scope', {})
            row = dict(id=sample.get('row_id') or str(uuid.uuid4()),
                       session=sample.get('session') or ('knowledge-edit:' + uuid.uuid4().hex),
                       source=sample.get('example', ''), factory=scope.get('factory', ''))
            store.propose_rule(row, sample.get('fragment', ''), item['new_value'].strip().upper(),
                               scope.get('category', ''), scope.get('role', 'изделие'))
            applied += 1
    store.sync()
    return dict(applied=applied, issues=plan['issues'])
