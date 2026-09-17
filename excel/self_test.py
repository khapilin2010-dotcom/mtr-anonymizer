"""Portable EXE acceptance smoke test using synthetic data only."""
import csv
import hashlib
import json
from pathlib import Path
import sys
import tempfile

from common.database import database_path
from excel.excel_engine import Anonymizer
from excel.file_io import process_file
from excel.reviewed_technical import REVIEWED_KEEP
from excel.supplemental_rules import EXTRA_RULES, EXTRA_GLOBAL_RULES, supplemental_digest


def expert_smoke():
    from excel.knowledge import KnowledgeStore, restore_backup
    from excel.expert_engine import ExpertAnonymizer
    from excel.review_io import import_review
    from openpyxl import Workbook, load_workbook
    with tempfile.TemporaryDirectory(prefix='MTR_Knowledge_') as tmp:
        root = Path(tmp); share = root / 'Общая база'; share.mkdir()
        alice = KnowledgeStore(root / 'alice', share, 'alice')
        bob = KnowledgeStore(root / 'bob', share, 'bob')
        src = root / 'Проверка.xlsx'; wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Наименование', 'Производитель'])
        ws.append(['SYNTHETIC-EXPERT-001', 'Клапан XYZ-500 DN50', 'Тестовый завод']); wb.save(src)
        az = ExpertAnonymizer(alice.snapshot()); out, report = process_file(src, root, az, knowledge_store=alice)
        assert alice.sync()['online']
        wb = load_workbook(out); ws = wb.active; ws.cell(2, 5, 'Клапан DN50'); wb.save(out); wb.close()
        result = import_review(out, bob); assert result['accepted'] == 1 and not result['issues'], result
        assert import_review(out, bob)['accepted'] == 0
        alice.sync(); final = ExpertAnonymizer(alice.snapshot()).anonymize('Клапан XYZ-500 DN50', 'SYNTHETIC-EXPERT-001', 'Тестовый завод')
        assert final['text'] == 'Клапан DN50' and final['status'] == 'ЗЕЛЁНЫЙ', final
        row = bob.session_rows(report['session'])[0]
        alice.decide(row, row['source']); alice.sync(); bob.sync()
        assert ExpertAnonymizer(bob.snapshot()).anonymize(row['source'], row['code'], row['factory'])['status'] == 'КРАСНЫЙ'
        backup = root / 'backup.zip'; bob.backup(backup); restore_backup(backup, root / 'restored')
        assert list((share / 'backups').glob('*.zip'))
        if sys.platform == 'win32':
            from excel.expert_ui import gui
            assert gui(root / 'ui-smoke', smoke=True)
    return ['shared_events', 'excel_return', 'idempotent_import', 'expert_reuse', 'conflict', 'backup_restore', 'new_ui']


def run():
    import xlrd
    import xlwt
    from openpyxl import Workbook, load_workbook
    from excel.knowledge import VERSION
    az = Anonymizer()
    assert len(az.registry) > 100000, 'Incomplete registry'
    assert sum(map(len, az.rules_by_inn.values())) >= 500, 'Incomplete rules'
    assert az.anonymize('Клапан HAWLE-TEST9 DN50')['text'] == 'Клапан DN50'
    keep = 'IP66 УХЛ1 Ex d IIC T6 DN50 PN16 ГОСТ 8732-78'
    source = f'Клапан ООО "Тестовый завод" ТУ 1234-567 {keep} № TEST-123 от 4 апреля 2025 г.'
    cleaned = az.anonymize(source)
    assert cleaned['text'] == f'Клапан {keep}', cleaned
    assert az.resolve_factory('203-987654321-22')[0] == ''
    reviewed = az.anonymize('Автомат ВА47-29 C16 2P 220АС', factory='Неизвестный сборщик')
    assert reviewed['text'] == 'Автомат C16 2P 220АС', reviewed
    pressure = az.anonymize('Манометр МП4-УУХЛ1-25 MPa М20х1,5-8g', factory='ИНН 7021000501')
    assert all(value in pressure['text'] for value in ('УХЛ1', '25 MPa', 'М20х1,5-8g')), pressure
    aero = az.anonymize('Установка АЭРО ИКСИА СКВ BOX(S)-S-63 КАС-W AI-TEST123 IP54', factory='ИНН 3257017280')
    assert aero['text'] == 'Установка СКВ BOX(S)-S-63 КАС-W AI-TEST123 IP54', aero
    assert aero['status'] == 'ЗЕЛЁНЫЙ' and 'решение пользователя' in aero['reason'], aero
    for name in ('АЭРО ИКСИА', 'АЭРО ИКСИУ'):
        quoted = az.anonymize('Установка ООО «' + name + ' СКВ BOX(S)-S-63» IP54')
        assert quoted['text'] == 'Установка СКВ BOX(S)-S-63 IP54', quoted
        assert quoted['status'] == 'ЗЕЛЁНЫЙ', quoted
    tested = []
    with tempfile.TemporaryDirectory(prefix='MTR_Excel_') as tmp:
        folder = Path(tmp) / 'Русская папка с пробелами'
        folder.mkdir()
        for suffix in ('.xlsx', '.xlsm', '.xls', '.csv'):
            src = folder / ('синтетический вход' + suffix)
            rows = [['Код Автодокс','Наименование','Производитель'],['000001',source,'Исходный поставщик']]
            if suffix == '.csv':
                with src.open('w',encoding='utf-8-sig',newline='') as handle:
                    csv.writer(handle,delimiter=';').writerows(rows)
            elif suffix == '.xls':
                wb=xlwt.Workbook();ws=wb.add_sheet('Готово')
                for i,row in enumerate(rows):
                    for j,value in enumerate(row):ws.write(i,j,value)
                wb.save(str(src))
            else:
                wb=Workbook();ws=wb.active;ws.title='Готово'
                for row in rows:ws.append(row)
                wb.save(src)
            before=src.read_bytes()
            dst,report=process_file(src,folder,az)
            assert report['rows']==1 and report['changed']==1
            assert before==src.read_bytes()
            if suffix=='.csv':
                with dst.open(encoding='utf-8-sig',newline='') as handle:
                    values=list(csv.reader(handle,delimiter=';'))[1]
            elif suffix=='.xls':
                values=xlrd.open_workbook(str(dst)).sheet_by_index(0).row_values(1)
            else:
                output=load_workbook(dst,keep_vba=suffix=='.xlsm')
                values=[c.value for c in output.active[2]]
                output.close()
            assert values[:3]==rows[1]
            assert values[4]==f'Клапан {keep}',values[4]
            assert values[6]
            tested.append(suffix)
    expert_checks = expert_smoke()
    exact_checks = exact_decision_smoke()
    selection_checks = selection_import_smoke()
    learning_checks = learning_smoke()
    if sys.platform=='win32':
        import tkinter as tk
        root=tk.Tk();root.withdraw();root.update();root.destroy()
    assert not any(n in sys.modules for n in ('mtr_core','MTR_Obezlichivatel','fitz','pymupdf'))
    return {'result':'SELF_TEST_OK','version':VERSION,'frozen':bool(getattr(sys,'frozen',False)),
            'learning_checks':learning_checks, 'selection_checks':selection_checks, 'exact_decision_checks':exact_checks, 'expert_checks':expert_checks,'database':'mtr_data.json.gz','database_sha256':hashlib.sha256(database_path().read_bytes()).hexdigest(),
            'registry_count':len(az.registry),'formats':tested,'source_unchanged':True,
            'supplemental_sha256':supplemental_digest(),
            'reviewed_keep_count':len(REVIEWED_KEEP),
            'supplemental_rule_count':len(EXTRA_RULES) + len(EXTRA_GLOBAL_RULES)}

def exact_decision_smoke():
    """Exercise the production decision path inside the frozen Windows EXE."""
    from openpyxl import Workbook, load_workbook
    from excel.simple_store import SimpleKnowledgeStore
    from excel.operator_engine import OperatorExpertAnonymizer
    from excel.simple_review import build_review_queue
    from excel.knowledge import case_key
    from excel.knowledge_edit_io import export_editable, commit_editable
    from excel.file_io import HEADERS
    source = 'Модуль автоматизированной технологической обвязки скважин МОС-3/1 Чертеж МОС-29.М1'
    with tempfile.TemporaryDirectory(prefix='MTR_Exact_') as tmp:
        root = Path(tmp)
        store = SimpleKnowledgeStore(root / 'local', root / 'База', 'engineer')
        src = root / 'Исходный Excel.xlsx'; wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Наименование', 'Завод'])
        ws.append(['631-336536', source, ''])
        wb.save(src); wb.close(); before = src.read_bytes()
        out, report = process_file(src, root, OperatorExpertAnonymizer(store.snapshot()), knowledge_store=store)
        row = store.session_rows(report['session'])[0]
        store.decide(row, source, 'Оставить как в исходном')
        assert build_review_queue(store, [report['session']]) == []
        def final_value(path):
            wb = load_workbook(path); ws = wb.active
            value = ws.cell(2, [c.value for c in ws[1]].index(HEADERS[1]) + 1).value
            wb.close(); return value
        assert final_value(out) == source
        repeated, _ = process_file(src, root, OperatorExpertAnonymizer(store.snapshot()), knowledge_store=store)
        assert final_value(repeated) == source
        changed = OperatorExpertAnonymizer(store.snapshot()).anonymize(source + ' другой', row['code'])
        assert not changed['knowledge'].get('exact_applied') and changed['status'] == 'ЖЁЛТЫЙ'
        export = root / 'База.xlsx'; export_editable(export, store)
        wb = load_workbook(export); ws = wb['Управление']; cols = {c.value: c.column for c in ws[1]}
        key = case_key(row['code'], source)
        n = next(i for i in range(2, ws.max_row + 1) if ws.cell(i, cols['Ключ']).value == key)
        ws.cell(n, cols['Новое решение'], source + ' проверено')
        ws.cell(n, cols['Действие'], 'ИЗМЕНИТЬ'); ws.cell(n, cols['Причина изменения'], 'Проверка EXE')
        wb.save(export); wb.close()
        assert commit_editable(export, store)['applied'] == 1
        assert OperatorExpertAnonymizer(store.snapshot()).anonymize(source, row['code'])['text'] == source + ' проверено'
        store.control(key, True, 'Проверка отключения')
        assert not OperatorExpertAnonymizer(store.snapshot()).anonymize(source, row['code'])['knowledge'].get('exact_applied')
        assert src.read_bytes() == before
        if sys.platform == 'win32':
            from excel.simple_ui import run as simple_ui
            assert simple_ui(root / 'ui', root / 'UI база', smoke=True)
    return ['mos_keep_original', 'repeat_output', 'changed_source_review', 'excel_edit_import',
            'disable_decision', 'recover_before_skip', 'original_unchanged'] + (['simple_ui'] if sys.platform == 'win32' else [])


def selection_import_smoke():
    """Verify the new import path and its actual Tk dialog inside the EXE."""
    from openpyxl import Workbook
    from excel.simple_store import SimpleKnowledgeStore
    from excel.history_import import preview_selection, commit_selection
    from excel.mapped_import import header_candidates, guess_mapping
    from excel.operator_engine import OperatorExpertAnonymizer
    source = 'Модуль автоматизированной технологической обвязки скважин МОС-3/1 Чертеж МОС-29.М1'
    checks = ['selection_import', 'selection_repeat', 'selection_conflict', 'selection_source_unchanged']
    with tempfile.TemporaryDirectory(prefix='MTR_Selection_') as tmp:
        folder = Path(tmp)
        store = SimpleKnowledgeStore(folder / 'local', folder / 'shared', 'engineer')
        path = folder / 'Проверенная выборка.xlsx'
        wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Исходное наименование', 'Обезличенное наименование', 'Завод'])
        ws.append(['00001', source, source, '']); wb.save(path); wb.close()
        before = path.read_bytes()
        candidate = header_candidates(path)[0]; mapping = guess_mapping(candidate['headers'])
        plan = preview_selection(path, store, candidate, mapping)
        assert commit_selection(path, store, plan)['accepted'] == 1
        repeated = preview_selection(path, store, candidate, mapping)
        assert repeated['repeated'] == 1 and repeated['ready'] == 0
        assert OperatorExpertAnonymizer(store.snapshot()).anonymize(source, '00001')['text'] == source
        assert path.read_bytes() == before
        ws.cell(2, 3, source + ' в комплекте'); wb.save(path); wb.close()
        plan = preview_selection(path, store, candidate, mapping)
        assert plan['conflicts'] == 1
        assert commit_selection(path, store, plan)['accepted'] == 0
        if sys.platform == 'win32':
            import time
            import tkinter as tk
            from tkinter import ttk, messagebox
            from excel.theme import apply_theme, WHITE, BLUE
            from excel.history_import_ui import SelectionDialog
            root = tk.Tk(); root.withdraw()
            style = apply_theme(root)
            assert style.lookup('TFrame', 'background') == WHITE
            assert style.lookup('Primary.TButton', 'background') == BLUE
            dialog = SelectionDialog(root, store, path)
            def idle():
                deadline = time.monotonic() + 20
                while dialog.busy and time.monotonic() < deadline:
                    root.update(); time.sleep(0.01)
                assert not dialog.busy, 'Import dialog worker did not finish'
                root.update()
            old_info, old_error = messagebox.showinfo, messagebox.showerror
            errors = []
            messagebox.showinfo = lambda *a, **k: None
            messagebox.showerror = lambda *a, **k: errors.append(a)
            try:
                idle(); assert not errors, errors
                dialog.preview(); idle(); assert not errors, errors
                assert dialog.plan['conflicts'] == 1
                key = dialog.plan['plan'][0]['key']
                dialog.rows.selection_set(key); root.update(); dialog.show_selected()
                assert dialog.commit_button.instate(['disabled'])
                dialog.take.set(True); dialog.toggle()
                dialog.reviewed.set(True); dialog.update_commit()
                assert not dialog.commit_button.instate(['disabled'])
                dialog.commit(); idle(); assert not errors, errors
                assert dialog.report['accepted'] == 1
                assert OperatorExpertAnonymizer(store.snapshot()).anonymize(source, '00001')['text'] == source + ' в комплекте'
                dialog.close()
            finally:
                messagebox.showinfo, messagebox.showerror = old_info, old_error
                root.destroy()
            checks += ['white_blue_theme', 'selection_dialog_preview_and_commit']
    return checks


def learning_smoke():
    """Learn from live engineer corrections using production store and engine."""
    from excel.simple_store import SimpleKnowledgeStore
    from excel.operator_engine import OperatorExpertAnonymizer
    from excel.semantics import features
    from excel.simple_review import refresh_review_row
    from excel.review_display import change_spans, restored_spans
    from excel.sync_feedback import feedback
    with tempfile.TemporaryDirectory(prefix='MTR_Learning_') as tmp:
        folder = Path(tmp)
        store = SimpleKnowledgeStore(folder / 'local', folder / 'shared', 'engineer')
        for index, family in enumerate(('МОС', 'ПКМ-ТСТ')):
            source = 'Модуль ' + family + '-1'
            row = dict(id=str(index), session='smoke', source=source, automatic='Модуль',
                       code=str(index), factory='', classification=features(source))
            store.decide(row, source, 'Оставить как в исходном')
            following = dict(row, source='Модуль ' + family + '-2', code='new')
            got = refresh_review_row(store, following)
            assert got['automatic'] == following['source'], got
            assert got['provenance']['series_kept'] == [family], got
            assert change_spans(source, 'Модуль')[0]
            assert restored_spans(source, 'Модуль', source)
            store.sync(auto_backup=False)
        peer = SimpleKnowledgeStore(folder / 'peer', store.shared, 'peer')
        peer.sync(auto_backup=False)
        assert OperatorExpertAnonymizer(peer.snapshot()).anonymize('Модуль МОС-3')['knowledge']['series_kept'] == ['МОС']
        assert not feedback(store.shared, dict(online=False, pending=1, errors=['Нет доступа']))['ok']
    return ['mos_and_pkm_series_learning', 'next_row_refresh', 'shared_learning', 'color_diff_spans', 'sync_diagnostics']
