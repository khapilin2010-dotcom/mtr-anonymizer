import unittest
from unittest.mock import patch
from pathlib import Path
from tempfile import TemporaryDirectory
from excel.excel_engine import Anonymizer, review_tokens
from excel.file_io import process_file

class ReviewedRC6Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.az = Anonymizer()

    def test_functional_options_are_preserved(self):
        for source in ('Вентилятор с ТШК', 'Щит с панелью ПЭСПЗ',
                       'Кран с КОФ', 'Клапан с МРЗ',
                       'Шкаф электропитания Unum- 230 B, 50 Гц',
                       'Опора +1д25х3, расположение теплоспутника сверху',
                       'Термопреобразователь -50...+200/А, ГП'):
            self.assertEqual(review_tokens(source), [], source)
            self.assertEqual(self.az.anonymize(source)['text'], source)

    def test_article_is_removed_while_dimensions_survive(self):
        z = self.az.anonymize('Светильник арт. TEST-987-IP65-УХЛ1 230В')
        for item in ('IP65', 'УХЛ1', '230В'):
            self.assertIn(item, z['text'])
        self.assertNotIn('TEST-987', z['text'])
        self.assertEqual(self.az.anonymize(z['text'])['text'], z['text'])

    def test_article_does_not_consume_next_word_or_ol(self):
        z = self.az.anonymize('Изделие артикул TEST987 с креплением Комплектация по обосновывающему документу TEST-ОВ.ОЛ1')
        self.assertEqual(z['text'], 'Изделие с креплением Комплектация по обосновывающему документу TEST-ОВ.ОЛ1')

    def test_article_with_spaced_groups_removed_once(self):
        source = 'Шкаф арт. ПИК 123 456 789 IP65'
        z = self.az.anonymize(source)
        self.assertEqual(z['text'], 'Шкаф IP65')
        self.assertEqual(self.az.anonymize(z['text'])['text'], z['text'])

    def test_article_with_no_number_is_not_guessed(self):
        source = 'Изделие артикул отсутствует, размер 20мм'
        self.assertEqual(self.az.anonymize(source)['text'], source)

    def test_rgt_size_and_options_survive_series_root(self):
        source = 'Регистр РГТ-1-108х3-2м-5-ОП-Э'
        z = self.az.anonymize(source, factory='ИНН 6452951387')
        self.assertEqual(z['text'], 'Регистр 1-108х3-2м-5-ОП-Э')
        self.assertEqual(z['status'], 'ЗЕЛЁНЫЙ')
        self.assertEqual(self.az.anonymize(z['text'], factory='ИНН 6452951387')['text'], z['text'])

    def test_explicit_company_is_not_hidden_by_keep(self):
        z = self.az.anonymize('Щит ООО "ПЭСПЗ" IP65')
        self.assertNotIn('ПЭСПЗ', z['text'])
        self.assertIn('IP65', z['text'])

    def test_drawing_does_not_remove_ol(self):
        z = self.az.anonymize('Модуль МОС-3/1 Чертеж МОС-29.М1 TEST-ОВ.ОЛ2', factory='ИНН 4501006138')
        self.assertIn('3/1', z['text'])
        self.assertIn('TEST-ОВ.ОЛ2', z['text'])
        self.assertNotIn('МОС', z['text'])

    def test_drawing_revision_is_not_square_metres(self):
        source = 'Модуль Чертеж МОС-28.М2 площадь 2м2 TEST-ОВ.ОЛ1'
        z = self.az.anonymize(source, factory='ИНН 4501006138')
        self.assertEqual(z['text'], 'Модуль площадь 2м2 TEST-ОВ.ОЛ1')
        self.assertEqual(self.az.anonymize(z['text'], factory='ИНН 4501006138')['text'], z['text'])

    def test_gland_series_decoded_without_losing_diameter(self):
        for token, expected in [('КВМ25', 'Ду25 мм'), ('КВБ17', 'до 17 мм'),
                                ('КВБУ14', '10–14 мм'), ('КВБУ18', '14–18 мм')]:
            z = self.az.anonymize('Извещатель с вводом ' + token + ' IP65')
            self.assertNotIn(token, z['text'])
            self.assertIn(expected, z['text'])
            self.assertIn('IP65', z['text'])
            self.assertEqual(self.az.anonymize(z['text'])['text'], z['text'])
        self.assertIn('КВМ99', self.az.anonymize('Извещатель КВМ99')['text'])

    def test_progress_does_not_rescan_sheet_per_row(self):
        from openpyxl import Workbook
        from openpyxl.worksheet.worksheet import Worksheet
        with TemporaryDirectory() as d:
            src = Path(d)/'source.xlsx'
            w=Workbook();w.active.append(['Код Autodocs','Наименование'])
            for i in range(300): w.active.append([str(i),'Кабель IP65'])
            w.save(src);w.close()
            calls=[];original=Worksheet.max_row.fget
            def counted(ws):
                calls.append(1)
                return original(ws)
            progress=[]
            with patch.object(Worksheet,'max_row',property(counted)):
                out, report=process_file(src,Path(d)/'out',self.az,progress.append)
            self.assertEqual(report['rows'],300)
            self.assertTrue(progress[-1].endswith('300 / 300'))
            self.assertLess(len(calls),10)

if __name__ == '__main__':unittest.main()
