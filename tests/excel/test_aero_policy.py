import unittest
from excel.excel_engine import Anonymizer


class AeroPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.az = Anonymizer()

    def test_name_only_keeps_complete_designations(self):
        tail = 'СКВ BOX(S)-S-6,3 КАС-W KAS-W KAC-W KAС-W AI-TEST-123 АИ-В3-660х350-1,4-3 IP54'
        for brand in ('АЭРО ИКСИА', 'Aero IXIA', 'ООО «АЭРО ИКСИА»', 'ООО АЭРО ИКСИА'):
            result = self.az.anonymize('Установка ' + brand + ' ' + tail, factory='ИНН 3257017280')
            self.assertEqual(result['text'], 'Установка ' + tail)
            self.assertEqual(result['status'], 'ЗЕЛЁНЫЙ')
            self.assertIn('решение пользователя', result['reason'])
            self.assertEqual(self.az.anonymize(result['text'], factory='ИНН 3257017280')['text'], result['text'])

    def test_legacy_series_kept(self):
        source = 'Установка CompactVolume CrisperLine RunAir RunCool RunRow AI-TEST123'
        result = self.az.anonymize(source, factory='ИНН 3257017280')
        self.assertEqual(result['text'], source)
        self.assertEqual(result['status'], 'ЗЕЛЁНЫЙ')

    def test_name_recognized_without_code(self):
        result = self.az.anonymize('Установка Aero IXIA СКВ BOX(S)-S-63 КАС-W')
        self.assertEqual(result['text'], 'Установка СКВ BOX(S)-S-63 КАС-W')
        self.assertEqual(result['status'], 'ЗЕЛЁНЫЙ')

    def test_attached_designation_not_swallowed(self):
        for brand in ('АЭРО ИКСИА', 'Aero IXIA', 'ИКСИА'):
            result = self.az.anonymize('Установка ' + brand + '-AI-TEST123')
            self.assertIn('AI-TEST123', result['text'])
            self.assertNotIn(brand, result['text'])

    def test_other_suppliers_not_implicitly_approved(self):
        result = self.az.anonymize('Установка СКВ BOX(S)-S-63 КАС-W', factory='ИНН 7720040225')
        self.assertEqual(result['status'], 'ЖЁЛТЫЙ')
        self.assertNotIn('решение пользователя', result['reason'])

    def test_other_brand_and_general_requisites_still_removed(self):
        result = self.az.anonymize('Установка АЭРО ИКСИА AI-TEST123 с приводом HAWLE-TEST9 ТУ 1234-567 № TEST-123 от 01.01.2025 IP54')
        self.assertEqual(result['text'], 'Установка AI-TEST123 с приводом IP54')

    def test_unrelated_unknown_stays_yellow(self):
        result = self.az.anonymize('Установка АЭРО ИКСИА AI-TEST123 XYZ-987')
        self.assertEqual(result['status'], 'ЖЁЛТЫЙ')
        self.assertIn('XYZ-987', result['reason'])
        self.assertNotIn('AI-TEST123', result['reason'])

    def test_brand_inside_ol_not_hidden(self):
        source = 'Установка TEST-АЭРО-ИКСИА-ОВ.ОЛ1'
        result = self.az.anonymize(source, factory='ИНН 3257017280')
        self.assertEqual(result['text'], source)
        self.assertEqual(result['status'], 'ЖЁЛТЫЙ')

    def test_code_scope_matches_explicit_factory(self):
        self.az.registry['999987654'] = 'ООО «АЭРО ИКСИА» ИНН 3257017280'
        try:
            source = 'Установка СКВ BOX(S)-S-63 AI-TEST123'
            self.assertEqual(self.az.anonymize(source, '999987654')['status'], 'ЗЕЛЁНЫЙ')
            self.assertEqual(self.az.anonymize(source, '999987654')['text'], source)
        finally:
            self.az.registry.pop('999987654')
            self.az.resolve_factory.cache_clear()

if __name__ == '__main__':
    unittest.main()
