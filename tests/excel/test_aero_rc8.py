import unittest
from excel.excel_engine import Anonymizer

class AeroRC8Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls): cls.az=Anonymizer()

    def test_shared_quotes_keep_designations_in_both_directions(self):
        for left,right in [('"','"'),('«','»'),('“','”')]:
            for body,expected in [
                ('АЭРО ИКСИА СКВ BOX(S)-S-63','СКВ BOX(S)-S-63'),
                ('AI-TEST123 АЭРО ИКСИА КАС-W','AI-TEST123 КАС-W'),
                ('СКВ BOX(S)-S-63 АЭРО ИКСИА','СКВ BOX(S)-S-63')]:
                source='Установка ООО '+left+body+right+' IP54'
                for factory in ('','ИНН 3257017280'):
                    z=self.az.anonymize(source,factory=factory)
                    self.assertEqual(z['text'],'Установка '+expected+' IP54',source)
                    self.assertEqual(z['status'],'ЗЕЛЁНЫЙ',source)
                    self.assertEqual(self.az.anonymize(z['text'],factory=factory)['text'],z['text'])

    def test_unknown_quoted_designation_survives_and_stays_yellow(self):
        z=self.az.anonymize('Установка ООО "АЭРО ИКСИА ZXQ-987" IP54')
        self.assertEqual(z['text'],'Установка ZXQ-987 IP54')
        self.assertEqual(z['status'],'ЖЁЛТЫЙ')
        self.assertIn('ZXQ-987',z['reason'])

    def test_typo_with_and_without_registry(self):
        for factory in ('','ИНН 3257017280'):
            for name in ('АЭРО ИКСИУ','аэро иксиу','Аэро\u00a0Иксиу'):
                z=self.az.anonymize('Установка '+name+' СКВ BOX(S)-S-63 IP54',factory=factory)
                self.assertEqual(z['text'],'Установка СКВ BOX(S)-S-63 IP54')
                self.assertEqual(z['status'],'ЗЕЛЁНЫЙ')

    def test_typo_inside_shared_quotes(self):
        z=self.az.anonymize('Установка ООО «АЭРО ИКСИУ AI-TEST123» IP54')
        self.assertEqual(z['text'],'Установка AI-TEST123 IP54')
        self.assertEqual(z['status'],'ЗЕЛЁНЫЙ')

    def test_typo_requires_full_name_and_word_boundary(self):
        for source in ('Установка ИКСИУ IP54','Установка АЭРО ИКСИУМ IP54'):
            self.assertEqual(self.az.anonymize(source)['text'],source)

    def test_other_brand_and_requisites_inside_shared_quotes(self):
        z=self.az.anonymize('Установка ООО "АЭРО ИКСИА AI-TEST123 HAWLE-TEST9 ТУ 1234-567" IP54')
        self.assertEqual(z['text'],'Установка AI-TEST123 IP54')

    def test_other_organization_and_absolute_keep_unaffected(self):
        z=self.az.anonymize('Клапан ООО «Тестовый Завод» DN50 PN16')
        self.assertEqual(z['text'],'Клапан DN50 PN16')
        source='Установка TEST-АЭРО-ИКСИА-ОВ.ОЛ1'
        z=self.az.anonymize(source,factory='ИНН 3257017280')
        self.assertEqual(z['text'],source)
        self.assertEqual(z['status'],'ЖЁЛТЫЙ')

    def test_supplier_role_removed_without_product(self):
        z=self.az.anonymize('Установка изготовитель ООО "АЭРО ИКСИА СКВ BOX(S)-S-63" IP54')
        self.assertEqual(z['text'],'Установка СКВ BOX(S)-S-63 IP54')

if __name__=='__main__':unittest.main()
