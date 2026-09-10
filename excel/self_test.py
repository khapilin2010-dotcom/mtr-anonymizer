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
    if sys.platform=='win32':
        import tkinter as tk
        root=tk.Tk();root.withdraw();root.update();root.destroy()
    assert not any(n in sys.modules for n in ('mtr_core','MTR_Obezlichivatel','fitz','pymupdf'))
    return {'result':'SELF_TEST_OK','version':'1.3 RC1','frozen':bool(getattr(sys,'frozen',False)),
            'expert_checks':expert_checks,'database':'mtr_data.json.gz','database_sha256':hashlib.sha256(database_path().read_bytes()).hexdigest(),
            'registry_count':len(az.registry),'formats':tested,'source_unchanged':True,
            'supplemental_sha256':supplemental_digest(),
            'reviewed_keep_count':len(REVIEWED_KEEP),
            'supplemental_rule_count':len(EXTRA_RULES) + len(EXTRA_GLOBAL_RULES)}
