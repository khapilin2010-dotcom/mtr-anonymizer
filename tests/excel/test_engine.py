import gzip
import json
import re

import pytest

from excel.excel_engine import Anonymizer, protected_ranges


@pytest.fixture(scope='module')
def az():
    return Anonymizer()


TECHNICAL = [
    'IP66', 'IP 54', 'УХЛ1', 'УХЛ 4.2', 'Ex d IIC T6', '1ExdbIIBT4Gb',
    'DN50 PN16', 'DN 100 PN 25', 'сталь 12Х18Н10Т', 'сталь 09Г2С', 'Ст.20',
    '09Г2С', '159х5,0 мм', '3x2,5 мм', 'давление 16 МПа', 'напряжение 24В',
    'температура от -60°С до +100°С', '220 В', '0,66 кВ', 'ГОСТ 8732-78',
    'ГОСТ Р 51330.0-99', 'ГОСТ Р МЭК 60079-0-2011', 'SDR11', 'RAL 7035',
    'TEST.0001-АТТ.ОЛ1', 'TEST-ЭГ.ОЛ1-02', 'ОЛ № TEST-123',
    'Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1',
    'Комплектация по обосновывающему документу\nTEST.0001-АТТ.ОЛ1',
]


@pytest.mark.parametrize('technical', TECHNICAL)
@pytest.mark.parametrize('deletion', ['Унипол', 'ТУ 1234-567-89012345-2025'])
def test_absolute_keep(az, technical, deletion):
    source = f'Изделие {deletion} {technical}'
    result = az.anonymize(source)
    assert technical in result['text']
    assert deletion not in result['text']
    assert result['removed']


@pytest.mark.parametrize('tu', [
    'ТУ 16.К01-49-2005', 'по ТУ 3434-001-79740390-2007',
    'тип ТУ 1234-567', 'по типу ТУ 1234-567', 'ТУ №1234-567',
    'ТУ № 1234-567', 'ТУ', 'ТУ 1234-\n567-2025', 'ТУ 1234 - 567 - 2025',
    'ТУ 16 К01-49-2005', '(1234-567 ТУ)',
])
def test_tu_variants(az, tu):
    result = az.anonymize(f'Кабель {tu} IP66')
    assert result['text'] == 'Кабель IP66'
    assert az.anonymize(result['text'])['text'] == result['text']


@pytest.mark.parametrize('org', [
    'ООО "Газпром добыча Иркутск"', 'ООО "Сервисный центр СБМ"', 'ООО.НПО.ФСА',
    '[ООО ]', 'Реестр МТР ПАО Газпром № 3315', 'АО «Тестовый Завод»',
    'ПАО "Тестовый завод"', 'ОАО "Тестовый завод"', 'ЗАО "Тестовый завод"',
    'НПО "Тестовый завод"', 'НПП "Тестовый завод"', 'ООО Тестовый Завод',
    'ИНН 1234567890', 'ИНН: 123456789012',
])
def test_organizations(az, org):
    result = az.anonymize(f'Клапан {org} DN50 PN16')
    assert result['text'] == 'Клапан DN50 PN16'
    assert result['removed']


@pytest.mark.parametrize('letter', [
    '№ 50-01/АВ-013153 от 4 апреля 2025 г.',
    '№ 50-01/АВ-052742 от 1 декабря 2023 г.',
    '№ 50-021-027706 от 19 июля 2024 г.',
    'ВО№ TEST-123 от 12.03.2024', 'ВО № TEST-123 от 12.03.2024',
    'Письмо № TEST-123 от 12.03.2024',
    'служебное письмо от 12.03.2024 № TEST-123',
    '№ 1 от 01.01.25',
])
def test_letters(az, letter):
    keep = 'Комплектация по обосновывающему документу TEST-АТТ.ОЛ1'
    result = az.anonymize(f'{keep} {letter}')
    assert result['text'] == keep
    assert any('реквизиты письма' in item for item in result['removed'])


@pytest.mark.parametrize('brand', [
    'Унипол', 'Гиперфлоу', 'Метран', 'Вэлан', 'Ризур', 'Рубеж', 'Армтел',
    'Пензтяжпромарматура', 'Волжский трубный завод', 'ИКСИА', 'Аэро Иксиа',
    'Теплолюкс', 'Эрис', 'Феррум', 'Телта', 'Тропик', 'PERCo', 'Durcon', 'TOFF',
])
def test_existing_brands(az, brand):
    result = az.anonymize(f'Изделие {brand} IP66 УХЛ1')
    assert brand.casefold() not in result['text'].casefold()
    assert 'IP66 УХЛ1' in result['text']
    assert result['removed']


def test_no_generic_device_deleted(az):
    result = az.anonymize('Прибор ООО.НПО.ФСА ТУ 12-34 IP66')
    assert result['text'] == 'Прибор IP66'


def test_keep_beats_delete_and_disallows_green(az):
    keep = 'Комплектация по обосновывающему документу Унипол TEST-АТТ.ОЛ1'
    result = az.anonymize(keep)
    assert result['text'] == keep
    assert result['status'] == 'ЖЁЛТЫЙ'
    assert result['reason']


def test_unknown_organization_is_never_green(az):
    result = az.anonymize('Клапан неизвестного АО без названия IP66')
    assert result['status'] == 'ЖЁЛТЫЙ'
    assert 'АО' not in result['text']


def test_repeated_technical_values_preserved(az):
    result = az.anonymize('Клапан Унипол IP66; IP66; ГОСТ 8732-78; ГОСТ 8732-78')
    assert result['text'].count('IP66') == 2
    assert result['text'].count('ГОСТ 8732-78') == 2


def test_literal_database_rules_and_composite_code(tmp_path):
    data = {'version': 'synthetic', 'registry': {'123456': 'ООО "Синтетика" ИНН 1234567890'},
            'aliases': [{'manufacturer': 'ООО "Синтетика"', 'inn': '1234567890', 'alias': 'Синтетика'}],
            'rules': [{'inn': '1234567890', 'manufacturer': 'ООО "Синтетика"', 'trigger': 'ZZTEST-',
                       'type': 'Фирменная серия / модель', 'apply': 'Начало фирменного обозначения'}],
            'global_unique_rules': [{'trigger': 'TESTBRAND', 'note': 'recovery', 'apply': 'Точное совпадение'}]}
    path = tmp_path / 'synthetic.json.gz'
    with gzip.open(path, 'wt', encoding='utf-8') as handle:
        json.dump(data, handle)
    az = Anonymizer(path)
    result = az.anonymize('Клапан ZZTEST-123-IP66 TESTBRAND DN50', '631-123456')
    assert result['text'] == 'Клапан IP66 DN50'
    assert 'Синтетика' in result['factory']
    # Documented IP is a technical remainder, not an unresolved model.
    assert result['status'] == 'ЗЕЛЁНЫЙ'
    assert result['reason'] == ''


def test_missing_database_fails(tmp_path):
    with pytest.raises(FileNotFoundError):
        Anonymizer(tmp_path / 'absent.gz')


def test_no_manufacturer_is_not_assumed_safe(az):
    assert az.anonymize('Клапан неизвестной марки DN50')['status'] == 'ЖЁЛТЫЙ'


@pytest.mark.parametrize('brand', ['Хакель','ПТИМаш','БРОЕН','Сименс','АББ','ДКС','Роквул','Гусар'])
def test_short_database_manufacturers_removed(az,brand):
    result=az.anonymize(f'Клапан {brand} ТУ 123-45 IP66')
    assert result['text']=='Клапан IP66'
    assert result['factory']


@pytest.mark.parametrize('technical',['Прибор','канат','никель','сплав','сенсор'])
def test_technical_nouns_are_not_company_mentions(az,technical):
    result=az.anonymize(f'{technical} Унипол IP66')
    assert result['text']==f'{technical} IP66'


@pytest.mark.parametrize('org', ['ООО "НПО "Тест" Завод"','ООО "НПО "Тест""','ООО "Тестовый завод"'])
def test_nested_company_quotes(az,org):
    result=az.anonymize(f'Клапан производитель {org} DN50')
    assert result['text']=='Клапан DN50'
    assert '"' not in result['text']


def test_tu_removal_alone_does_not_prove_manufacturer_absent(az):
    result=az.anonymize('Клапан НЕИЗВЕСТНЫЙБРЕНД ТУ 123-45 IP66')
    assert result['status']=='ЖЁЛТЫЙ'
    assert 'ТУ' not in result['text']
