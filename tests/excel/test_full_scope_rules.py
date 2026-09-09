"""RC5 boundaries and counterexamples, using synthetic descriptions."""
import unittest
from excel.excel_engine import Anonymizer, review_tokens

class FullScopeRulesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.az = Anonymizer()

    def test_voltage_decoded_without_losing_value(self):
        for token in ('MV220', 'МV220', 'МВ220', 'МS24'):
            with self.subTest(token=token):
                original = 'Клапан с электроприводом ' + token + ' IP54'
                z = self.az.anonymize(original, factory='ИНН 7720040225')
                value = '24' if token.endswith('24') else '220'
                self.assertEqual(z['text'], 'Клапан с электроприводом ' + value + ' В IP54')
                self.assertNotIn(value, '; '.join(z['removed']))
                self.assertEqual(self.az.anonymize(z['text'], factory='ИНН 7720040225')['text'], z['text'])

    def test_drive_without_factory_match_stays_reviewable(self):
        z = self.az.anonymize('Клапан MV220', factory='Другой завод')
        self.assertEqual(z['text'], 'Клапан MV220')
        self.assertEqual(z['status'], 'ЖЁЛТЫЙ')
        self.assertIn('MV220', review_tokens(z['text']))

    def test_embedded_opening_flanges_and_voltage_preserved(self):
        z = self.az.anonymize('Клапан КПУ-1Н-0-Н-160-2хф-МV220-СН-0', factory='ИНН 7720040225')
        for value in ('160', '2хф', '220 В'):
            self.assertIn(value, z['text'])
        self.assertNotIn('КПУ', z['text'])
        self.assertNotIn('МV', z['text'])

    def test_dotted_meter_index_is_not_a_length(self):
        z = self.az.anonymize('Счетчик СЭТ-4ТМ.02М.15 длина 2м 220 В')
        self.assertEqual(z['text'], 'Счетчик длина 2м 220 В')

    def test_real_alloy_grades_survive(self):
        for steel in ('09Г2С', '12Х18Н10Т', '10ХСНД', '17Г1С', '08Х18Н10', '30ХГСА', '20ГЛ', '30ХМЛ', '4К48'):
            z = self.az.anonymize('Деталь HAWLE-TEST9-' + steel)
            self.assertIn(steel, z['text'])

    def test_tol_and_full_completion_phrase_survive(self):
        text = 'Клапан Комплектация по обосновывающему документу TEST.001-ТХ.ТОЛ10-6'
        z = self.az.anonymize(text + ' Унипол')
        self.assertEqual(z['text'], text)
        self.assertNotIn('TEST.001-ТХ.ТОЛ10-6', review_tokens(z['text']))

    def test_vendor_inside_ol_is_not_certified(self):
        z = self.az.anonymize('Клапан Комплектация по обосновывающему документу HAWLE-ТХ.ОЛ1')
        self.assertIn('HAWLE-ТХ.ОЛ1', z['text'])
        self.assertEqual(z['status'], 'ЖЁЛТЫЙ')

    def test_technical_contexts_resolved(self):
        for text in ('Шкаф с УЗИП', 'Преобразователь избыточного давления М20',
                     'Пиктограмма ВЫХОД', 'Шкаф габариты 500hх400х220',
                     'Кабель Cat5e FRLS-ХЛ', 'Система управления Modbus RTU/TCP',
                     'Низковольтное устройство НКУ', 'Светильник потолок типа "Грильято" GRILIATO'):
            with self.subTest(text=text):
                self.assertEqual(review_tokens(text), [])

    def test_technical_acronym_does_not_hide_company(self):
        z = self.az.anonymize('Система управления ООО "АСУ" IP54')
        self.assertNotIn('АСУ', z['text'])
        self.assertIn('IP54', z['text'])

    def test_pictogram_catalog_code_removed_content_retained(self):
        z = self.az.anonymize('Пиктограмма "ВЫХОД" 335х165 ПЭУ 010', factory='ИНН 6229028102')
        self.assertEqual(z['text'], 'Пиктограмма "ВЫХОД" 335х165')
        other = self.az.anonymize('Пиктограмма ПЭУ 999', factory='ИНН 6229028102')
        self.assertIn('ПЭУ 999', other['text'])

    def test_unknown_factory_token_is_not_green_by_identity_alone(self):
        z = self.az.anonymize('Светильник ZZMODEL987 IP65', factory='ИНН 6229028102')
        self.assertEqual(z['status'], 'ЖЁЛТЫЙ')
        self.assertIn('ZZMODEL987', z['reason'])

if __name__ == '__main__':
    unittest.main()
