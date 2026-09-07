"""Synthetic regressions; no production resource names or Autodocs codes."""
import gzip
import json
import re
import subprocess
import sys

import pytest

from excel.excel_engine import Anonymizer
from excel.supplemental_rules import EXTRA_RULES, EXTRA_GLOBAL_RULES


@pytest.fixture(scope='module')
def az():
    return Anonymizer()


@pytest.fixture
def synthetic(tmp_path):
    db = dict(version='synthetic', registry={'999999991': 'компания "Альфа Синтетик", Россия, г. Тест'},
              aliases=[dict(alias='Альфа Синтетик', manufacturer='компания "Альфа Синтетик"', inn=''),
                       dict(alias='Бета Синтетик', manufacturer='АО "Бета Синтетик"', inn='')],
              rules=[dict(manufacturer='компания "Альфа Синтетик"', inn='', trigger='АЛФА.', apply='Полный код КД'),
                     dict(manufacturer='АО "Бета Синтетик"', inn='', trigger='БЕТА.', apply='Полный код КД')],
              global_unique_rules=[dict(trigger='ТЕСТБРЕНД', apply='Точное совпадение')])
    path = tmp_path / 'synthetic.json.gz'
    with gzip.open(path, 'wt', encoding='utf-8') as f:
        json.dump(db, f)
    return Anonymizer(path)


@pytest.mark.parametrize('context', [dict(code='999999991'),
                                    dict(factory='компания "Альфа Синтетик", Россия, г. Тест')])
def test_no_inn_rules_use_resolved_identity(synthetic, context):
    z = synthetic.anonymize('Прибор АЛФА.123456.007 IP66', **context)
    assert z['text'] == 'Прибор IP66'
    assert z['status'] == 'ЗЕЛЁНЫЙ'


def test_empty_inn_never_pools_other_factories(synthetic):
    z = synthetic.anonymize('Прибор БЕТА.123456.007', code='999999991')
    assert 'БЕТА.123456.007' in z['text']
    assert z['status'] == 'ЖЁЛТЫЙ'


def test_incomplete_ex_marking_cannot_hang_long_description():
    script = '''
from excel.excel_engine import Anonymizer
source = 'Датчик Exd с теплоизолирующим чехлом ' + 'и необогреваемым покрытием ' * 100
result = Anonymizer().anonymize(source)
assert result['text'] == source
'''
    subprocess.run([sys.executable, '-c', script], check=True, timeout=15)


def test_no_inn_alias_activates_own_rules(synthetic):
    z = synthetic.anonymize('Прибор Бета Синтетик БЕТА.123456.007 IP66')
    assert z['text'] == 'Прибор IP66'


def test_alias_attached_model_and_keep(synthetic):
    z = synthetic.anonymize('Прибор Альфа Синтетик-TEST9-IP66-2шт DN50')
    assert 'TEST9' not in z['text']
    assert 'IP66' in z['text'] and '2шт' in z['text'] and 'DN50' in z['text']


@pytest.mark.parametrize('model', ['ZXQ-987', 'Ultimate', 'UNKNOWN', 'КЗТЕСТр-80х1', 'МОДЕЛЬ-123'])
def test_unknown_model_never_green(synthetic, model):
    z = synthetic.anonymize('Прибор ' + model, code='999999991')
    assert model in z['text']
    assert z['status'] == 'ЖЁЛТЫЙ'


@pytest.mark.parametrize('keep', ['IIGbIIBT4', 'IExdIIСT5Gb', '1еxdIIBT4', '48x1000BASE-X',
                                  '6x40GBASE-R', 'Pt100', 'Т15К6', '4-водный', 'HPL-пластик',
                                  'GPS/Bluetooth', 'ОСТ 38 0183-75', '2-х-проводная'])
def test_extended_keep(az, keep):
    z = az.anonymize('Прибор Унипол ' + keep)
    assert keep in z['text']
    assert 'Унипол' not in z['text']


@pytest.mark.parametrize('keep', ['3шт', '5компл', '7шт', '9шт', '50г', '10мл', '4мА',
                                  '40*60*40', 'Exd', 'Ex db', '220V', 'изм.12', 'D630',
                                  '1ExdbIIBT4GbX', '1ExdIIC T6 Gb X', 'IIGbIIBT4X'])
def test_attached_brand_never_eats_engineering_values(az, keep):
    z = az.anonymize('Прибор HAWLE-TEST-' + keep)
    assert keep in z['text']
    assert 'HAWLE' not in z['text']


@pytest.mark.parametrize('text', ['Знак дорожный "Тупик"', 'Тумба2', 'Турбина-TEST9', 'Трубка'])
def test_tu_is_not_a_prefix_of_ordinary_words(az, text):
    z = az.anonymize(text)
    assert z['text'] == text
    assert az.anonymize(z['text'])['text'] == text


def test_catalog_model_is_not_a_bare_dimension(az):
    z = az.anonymize('Изделие S5700-TEST9 DN50', factory='ООО "Техкомпания Хуавэй"')
    assert z['text'] == 'Изделие DN50'


def test_model_keep_collision_stays_yellow_after_brand_is_gone(az):
    z = az.anonymize('Изделие HAWLE-TEST9-IP66')
    assert 'HAWLE' not in z['text'] and 'IP66' in z['text']
    assert z['status'] == 'ЖЁЛТЫЙ'
    assert 'KEEP' in z['reason']


@pytest.mark.parametrize('row', EXTRA_RULES + EXTRA_GLOBAL_RULES,
                         ids=lambda row: row['trigger'])
def test_confirmed_family_with_engineering_tail(az, row):
    token = row['trigger']
    if token == 'Dorado':
        token += '5000'
    if token == 'СМ5.':
        token += '999.999'
    z = az.anonymize('Изделие ' + token + ' IP66 DN50 220 В', factory=row['manufacturer'])
    assert not re.search(row['regex'], z['text'], re.I), z
    assert 'IP66 DN50 220 В' in z['text']
    assert row['source'].startswith('https://')
