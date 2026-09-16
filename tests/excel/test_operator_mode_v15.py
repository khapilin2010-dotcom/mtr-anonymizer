import uuid

from openpyxl import Workbook

from excel.knowledge import Snapshot, case_key
from excel.operator_engine import OperatorExpertAnonymizer
from excel.simple_store import SimpleKnowledgeStore
from excel.legacy_import import preview_legacy


def _case_event(source, final, user='ENGINEER\\one'):
    key = case_key('', source, '')
    return dict(
        schema=1,
        store_id=str(uuid.uuid4()),
        id=str(uuid.uuid4()),
        user=user,
        timestamp='2026-09-13T12:00:00+00:00',
        version='test',
        kind='case',
        key=key,
        value=final,
        source=source,
        code='',
        factory='',
        automatic=source,
        row_id=str(uuid.uuid4()),
        session=str(uuid.uuid4()),
        action='Исправить',
        supersedes=[],
    )


def test_exact_case_is_applied_immediately_without_checkbox():
    source = 'Насос ZZZTESTBRAND'
    final = 'Насос'
    snapshot = Snapshot([_case_event(source, final)])

    cautious = OperatorExpertAnonymizer(snapshot, auto_apply_confirmed=False)
    result = cautious.anonymize(source)
    assert result['text'] == final
    assert result['status'] == 'ЗЕЛЁНЫЙ'
    assert result['knowledge']['status'] == 'ACTIVE'

    automatic = OperatorExpertAnonymizer(snapshot, auto_apply_confirmed=True)
    auto_result = automatic.anonymize(source)
    assert auto_result['text'] == final


def test_exact_keep_original_is_applied_immediately():
    source = 'Модуль МОС-3/1 Чертеж МОС-29.М1'
    snapshot = Snapshot([_case_event(source, source)])
    result = OperatorExpertAnonymizer(snapshot, auto_apply_confirmed=False).anonymize(source)
    assert result['text'] == source
    assert result['status'] == 'ЗЕЛЁНЫЙ'


def test_simple_store_creates_adjacent_event_folder(tmp_path):
    shared = tmp_path / 'MTR_Knowledge'
    local = tmp_path / 'local'
    store = SimpleKnowledgeStore(local, shared, user='ENGINEER\\one')
    assert store.shared == shared
    assert (shared / 'mtr-knowledge.json').exists()


def test_old_reviewed_workbook_can_be_previewed(tmp_path):
    shared = tmp_path / 'MTR_Knowledge'
    store = SimpleKnowledgeStore(tmp_path / 'local', shared, user='ENGINEER\\one')
    path = tmp_path / 'old.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.append(['Код Автодокс', 'Исходное наименование', 'Завод', 'Обезличенное наименование'])
    ws.append(['1001', 'Насос ZZZTESTBRAND', 'Тестовый завод', 'Насос'])
    wb.save(path)
    wb.close()

    plan = preview_legacy(path, store)
    assert plan['ready'] == 1
    assert plan['conflicts'] == 0
    assert not plan['issues']
