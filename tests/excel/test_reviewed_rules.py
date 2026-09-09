"""Documented RC4 decisions and counterexamples; synthetic descriptions only."""
import gzip
import json
import tempfile
import unittest
from pathlib import Path
from excel.excel_engine import Anonymizer, review_tokens
from excel.rules import code_candidates


class ReviewedRulesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.az = Anonymizer()

    def test_code_grammar_does_not_guess_from_estimate_fragments(self):
        for code in ('203-987654-22', '103-Д987654-80', 'text987654', '2026-987654-09'):
            with self.subTest(code=code):
                self.assertNotIn('987654', code_candidates(code))
        for code in ('987654', '631-987654', '6410-987654-3', '105-987654-С', '19-987654-C'):
            with self.subTest(code=code):
                self.assertIn('987654', code_candidates(code))

    def test_wrong_factory_is_not_applied_to_estimate(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'db.json.gz'
            db = dict(version='test', registry={'987654': 'Test Factory'}, aliases=[dict(alias='Test Factory', manufacturer='Test Factory', inn='')], rules=[dict(manufacturer='Test Factory', inn='', trigger='ZZSYNTHETIC', apply='Точное совпадение')],
                      global_unique_rules=[dict(trigger='ZZUNUSED', apply='Точное совпадение')])
            with gzip.open(path, 'wt', encoding='utf-8') as stream:
                json.dump(db, stream)
            az = Anonymizer(path)
            self.assertEqual(az.resolve_factory('203-987654-22')[0], '')
            self.assertEqual(az.resolve_factory('631-987654')[0], 'Test Factory')
            self.assertEqual(az.resolve_factory('105-987654-С')[0], 'Test Factory')

    def test_technical_parameters_survive_and_do_not_warn_as_models(self):
        for text in ('Манометр М20х1,5-8g 25 MPa', 'Прибор G3/4 10 psi', 'Автомат C16 2P',
                     'Светильник 5000K', 'Вентилятор режим работы Т80', 'Цвет RAL D2 DESIGN',
                     'Ввод кабельный ExdG В1.5', 'Сигнальное табло ВЫХОД'):
            with self.subTest(text=text):
                result = self.az.anonymize(text + ' Унипол')
                self.assertEqual(result['text'], text)
                self.assertEqual(review_tokens(text), [])
                self.assertEqual(self.az.anonymize(result['text'])['text'], text)

    def test_unknown_tokens_and_wrong_context_are_not_certified(self):
        for text in ('Изделие ZZQ987', 'Изделие C16', 'Изделие 5000K', 'Изделие Т80'):
            with self.subTest(text=text):
                self.assertTrue(review_tokens(text))
                self.assertEqual(self.az.anonymize(text + ' Унипол')['status'], 'ЖЁЛТЫЙ')

    def test_component_model_removed_with_electrical_parameters_retained(self):
        result = self.az.anonymize('Автомат ВА47-29 C16 2P IP54', factory='Неизвестный сборщик')
        self.assertEqual(result['text'], 'Автомат C16 2P IP54')
        self.assertNotIn('КЭАЗ', result['factory'])
        self.assertNotIn('IEK', result['factory'])
        self.assertIn('ВА47-290', self.az.anonymize('Автомат ВА47-290 C16')['text'])

    def test_rubezh_kd_and_model_removed_through_other_assembler(self):
        result = self.az.anonymize('Модуль МИ-R2 ПАСН.999999.888 IP20', factory='Неизвестный сборщик')
        self.assertEqual(result['text'], 'Модуль IP20')
        self.assertIn('Рубеж', result['factory'])

    def test_lighting_series_requires_relevant_factory(self):
        for model in ('LYRA', 'LHT', 'CLEAN', 'TOTEM', 'PROTON', 'GLOBUS', 'NORTH', 'COMP'):
            with self.subTest(model=model):
                result = self.az.anonymize('Светильник ' + model + ' LED 5000K', factory='ИНН 6229028102')
                self.assertEqual(result['text'], 'Светильник LED 5000K')
        self.assertIn('CLEAN', self.az.anonymize('Операция CLEAN', factory='Неизвестный завод')['text'])

    def test_velan_series_removed_without_gland_configuration_loss(self):
        result = self.az.anonymize('Ввод кабельный ВК-Л-ВЭЛ2БТ-G3/4 ExdG В1.5', factory='ИНН 2619000120')
        self.assertNotIn('ВЭЛ', result['text'])
        for value in ('ВК-Л-', '2БТ', 'G3/4', 'ExdG', 'В1.5'):
            self.assertIn(value, result['text'])
        self.assertEqual(self.az.anonymize(result['text'], factory='ИНН 2619000120')['text'], result['text'])

    def test_pressure_value_and_embedded_climate_survive_model_deletion(self):
        result = self.az.anonymize('Манометр МП4-УУХЛ1-25 MPa М20х1,5-8g', factory='ИНН 7021000501')
        self.assertNotIn('МП4', result['text'])
        for value in ('УХЛ1', '25 MPa', 'М20х1,5-8g'):
            self.assertIn(value, result['text'])

    def test_coating_brand_does_not_delete_generic_f100(self):
        result = self.az.anonymize('Покрытие АРМОКОТ F100 цвет RAL D2 DESIGN 100 мкм')
        self.assertNotIn('F100', result['text'])
        self.assertIn('RAL D2 DESIGN', result['text'])
        self.assertIn('F100', self.az.anonymize('Бетон F100')['text'])

    def test_only_documented_keep_can_settle_collision(self):
        result = self.az.anonymize('Прибор HAWLE-TEST9-М20х1,5-8g')
        self.assertIn('М20х1,5-8g', result['text'])
        self.assertEqual(result['status'], 'ЗЕЛЁНЫЙ')
        ip = self.az.anonymize('Прибор HAWLE-TEST9-IP66')
        self.assertEqual(ip['text'], 'Прибор IP66')
        self.assertEqual(ip['status'], 'ЗЕЛЁНЫЙ')
        unresolved = self.az.anonymize('Прибор HAWLE-TEST9-IP66 ZZQ987')
        self.assertEqual(unresolved['status'], 'ЖЁЛТЫЙ')

    def test_generic_kd_identifies_document_without_inventing_factory(self):
        result = self.az.anonymize('Изделие АБВГ.999999.888СБ IP54', factory='Неизвестный завод')
        self.assertEqual(result['text'], 'Изделие IP54')
        self.assertEqual(result['factory'], 'Неизвестный завод')
        self.assertIn('АБВГ.999999.888-ОЛ1', self.az.anonymize('Документ АБВГ.999999.888-ОЛ1')['text'])
        self.assertIn('ГОСТ 123456.001', self.az.anonymize('Изделие ГОСТ 123456.001')['text'])

    def test_ol_has_absolute_priority_and_collision_is_not_hidden(self):
        phrase = 'Комплектация по обосновывающему документу TEST-Hawle.ОЛ1'
        result = self.az.anonymize('Изделие ' + phrase)
        self.assertIn(phrase, result['text'])
        self.assertEqual(result['status'], 'ЖЁЛТЫЙ')


if __name__ == '__main__':
    unittest.main()
