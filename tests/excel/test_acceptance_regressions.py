"""Synthetic cases derived from defect classes, without customer rows or codes."""
import gzip
import json
import re

import pytest

from excel.excel_engine import Anonymizer, review_tokens


@pytest.fixture(scope='module')
def az():
    return Anonymizer()


@pytest.mark.parametrize('text', [
    'Агрегат газоперекачивающий', 'Блок-бокс электрообогрева',
    'Элемент Z-образный', 'центральный диэлектрический элемент',
    'Соединение обмоток: треугольник звезда с 0', 'Расходомер-счетчик газа',
    'Реле-контактор отключения аккумулятора', 'Выключатель 1 полюс',
    'Iн пл. вст. 20 А', 'Устройство ЭХЗ', 'Мощность, кВт - 7,25',
])
def test_engineering_nouns_are_not_company_names(az, text):
    result = az.anonymize(text + ' Унипол')
    assert result['text'] == text
    assert not any(re.search(r'производитель: (Агрегат|кВт|звезда|счетчик|элемент)', x, re.I)
                   for x in result['removed'])


@pytest.mark.parametrize('unit', ['кВт', 'МВт', 'кВА', 'кВАр', 'В', 'мм', 'МПа'])
@pytest.mark.parametrize('separator', [': ', ' - ', '—'])
def test_units_before_values_are_retained(az, unit, separator):
    text = 'Параметр, ' + unit + separator + '7,25'
    assert az.anonymize(text + ' Унипол')['text'] == text


@pytest.mark.parametrize('noun', ['Агрегат', 'Звезда', 'Элемент', 'Контактор', 'КВТ'])
def test_explicit_company_is_still_removed(az, noun):
    result = az.anonymize('Изделие ООО "' + noun + '" IP66')
    assert result['text'] == 'Изделие IP66'


def test_registry_identity_does_not_turn_nouns_into_brands(tmp_path):
    db = dict(version='synthetic', registry={'999999991': 'ООО "Элемент"'},
              aliases=[dict(alias='Элемент', manufacturer='ООО "Элемент"', inn='')],
              rules=[dict(manufacturer='ООО "Элемент"', inn='', trigger='QZS.', apply='Полный код КД')],
              global_unique_rules=[dict(trigger='UNUSED_SYNTHETIC_BRAND', apply='Точное совпадение')])
    path = tmp_path / 'database.json.gz'
    with gzip.open(path, 'wt', encoding='utf-8') as stream:
        json.dump(db, stream)
    result = Anonymizer(path).anonymize('Элемент QZS.999.777 IP66', code='999999991')
    assert result['text'] == 'Элемент IP66'


@pytest.mark.parametrize('prefix', ['ТУ 9999-111-22222222-2020', 'инв.№777777', 'IP66', ')'])
def test_letter_without_a_separating_space(az, prefix):
    source = 'Прибор ' + prefix + '№ TEST-987 от 7 сентября 2026 г. п. 9'
    result = az.anonymize(source)
    assert 'TEST-987' not in result['text']
    assert '2026' not in result['text']
    assert 'п. 9' not in result['text']
    assert az.anonymize(result['text'])['text'] == result['text']


def test_ol_survives_adjacent_letter(az):
    ol = '9999.999-АТХ.ОЛ7'
    source = 'Изделие Комплектация по обосновывающему документу ' + ol + '№ TEST-987 от 7 сентября 2026 г.'
    result = az.anonymize(source)
    assert 'Комплектация по обосновывающему документу ' + ol in result['text']
    assert 'TEST-987' not in result['text']


def test_quoted_organization_with_branch_has_no_orphan_quote(az):
    result = az.anonymize('Изделие, разработчик ООО "Синтетический разработчик" Северный филиал. IP66')
    assert 'разработчик' not in result['text'].lower()
    assert 'филиал' not in result['text'].lower()
    assert '"' not in result['text']
    assert 'IP66' in result['text']


@pytest.mark.parametrize('brand', ['IEK', 'ИЭК', 'Schneider Electric', 'OptiDin',
                                  'PRIZMA/R', 'PRIZMA/S', 'Astra Linux Special Edition'])
def test_component_brands_do_not_require_assembler_identity(az, brand):
    result = az.anonymize('Шкаф, производитель: ' + brand + ', IP66 220 В', factory='Неизвестный сборщик')
    assert brand not in result['text']
    assert 'производитель:' not in result['text']
    assert 'IP66 220 В' in result['text']


def test_ordinary_hyphenated_words_do_not_create_model_warnings():
    assert not review_tokens('блочно-комплектный воздушно-отопительный Масса-1,25 т.')
    assert review_tokens('ZZQ-987') == ['ZZQ-987']
