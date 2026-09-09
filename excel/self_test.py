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
    if sys.platform=='win32':
        import tkinter as tk
        root=tk.Tk();root.withdraw();root.update();root.destroy()
    assert not any(n in sys.modules for n in ('mtr_core','MTR_Obezlichivatel','fitz','pymupdf'))
    return {'result':'SELF_TEST_OK','version':'1.2 RC4','frozen':bool(getattr(sys,'frozen',False)),
            'database':'mtr_data.json.gz','database_sha256':hashlib.sha256(database_path().read_bytes()).hexdigest(),
            'registry_count':len(az.registry),'formats':tested,'source_unchanged':True,
            'supplemental_sha256':supplemental_digest(),
            'reviewed_keep_count':len(REVIEWED_KEEP),
            'supplemental_rule_count':len(EXTRA_RULES) + len(EXTRA_GLOBAL_RULES)}
