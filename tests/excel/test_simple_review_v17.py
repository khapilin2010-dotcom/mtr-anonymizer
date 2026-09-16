import uuid

from openpyxl import Workbook, load_workbook

from excel.file_io import HEADERS
from excel.review_io import REVIEW_HEADERS
from excel.output_update import apply_decisions
from excel.operator_engine import OperatorExpertAnonymizer
from excel.simple_store import SimpleKnowledgeStore


def test_engineer_keep_original_wins_on_next_run(tmp_path):
    store = SimpleKnowledgeStore(tmp_path / 'local', tmp_path / 'MTR_Knowledge', user='ENGINEER\\one')
    source = 'Модуль автоматизированной технологической обвязки скважин МОС-3/1 Чертеж МОС-29.М1'
    automatic = 'Модуль автоматизированной технологической обвязки скважин 3/1 Чертеж -29.М1'
    row = dict(
        id=str(uuid.uuid4()), session=str(uuid.uuid4()), row=12, sheet='Лист1',
        code='631-336536', source=source, factory='', automatic=automatic,
        status='ЖЁЛТЫЙ', classification=None, provenance={}, removed=['МОС']
    )
    store.decide(row, source, 'Оставить как в исходном')
    store.sync(auto_backup=False)

    result = OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=False).anonymize(
        source, row['code'], row['factory'])
    assert result['text'] == source
    assert result['status'] == 'ЗЕЛЁНЫЙ'
    assert result['knowledge']['status'] in ('ACTIVE', 'TRUSTED')


def test_engineer_correction_wins_on_next_run(tmp_path):
    store = SimpleKnowledgeStore(tmp_path / 'local', tmp_path / 'MTR_Knowledge', user='ENGINEER\\one')
    source = 'Модуль МОС-3/1 Чертеж МОС-29.М1 Масса-16,2 т.'
    automatic = 'Модуль 3/1 Чертеж -29.М1 Масса-16,2 т.'
    corrected = 'Модуль МОС-3/1 Чертеж МОС-29.М1 Масса-16,2 т.'
    row = dict(
        id=str(uuid.uuid4()), session=str(uuid.uuid4()), row=2, sheet='Лист1',
        code='631-336536', source=source, factory='', automatic=automatic,
        status='ЖЁЛТЫЙ', classification=None, provenance={}, removed=['МОС']
    )
    store.decide(row, corrected, 'Исправить')
    store.sync(auto_backup=False)
    result = OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=False).anonymize(
        source, row['code'], row['factory'])
    assert result['text'] == corrected


def test_apply_decision_updates_generated_xlsx(tmp_path):
    path = tmp_path / 'result.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Выборка'
    original_headers = ['Код Автодокс', 'Наименование', 'Производитель']
    ws.append(original_headers + list(HEADERS) + list(REVIEW_HEADERS))
    source = 'Модуль МОС-3/1 Чертеж МОС-29.М1 Масса-16,2 т.'
    automatic = 'Модуль 3/1 Чертеж -29.М1 Масса-16,2 т.'
    ws.append(['631-336536', source, '', '', automatic, 'ЖЁЛТЫЙ', 'МОС', '', '', '', '', 'row-id', 'session-id'])
    wb.save(path)
    wb.close()

    row = dict(row=2, sheet='Выборка', source=source, code='631-336536', factory='')
    apply_decisions(path, [{'row': row, 'final': source, 'action': 'Оставить как в исходном'}])

    wb = load_workbook(path)
    ws = wb['Выборка']
    headers = [cell.value for cell in ws[1]]
    final_col = headers.index(HEADERS[1]) + 1
    status_col = headers.index(HEADERS[2]) + 1
    action_col = headers.index(REVIEW_HEADERS[0]) + 1
    assert ws.cell(2, final_col).value == source
    assert ws.cell(2, status_col).value == 'ПРОВЕРЕНО ИНЖЕНЕРОМ: Оставить как в исходном'
    assert ws.cell(2, action_col).value == 'Оставить как в исходном'
    wb.close()
