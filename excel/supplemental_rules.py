"""Excel-only confirmed catalog additions. No customer rows or frequencies.

Sources establish product family identity, not fitness for a particular project.
Attached model suffixes are removed subject to the engine's absolute KEEP.
"""
import re
import hashlib
import json


def family(manufacturer, inn, trigger, source, regex=None):
    tail = r'[A-Za-zА-Яа-яЁё0-9._/+\-]*'
    return dict(manufacturer=manufacturer, inn=inn, trigger=trigger,
                regex=regex or r'(?<!\w)' + re.escape(trigger) + tail,
                source=source, apply='Слитное/дефисное обозначение')


VEZA_CERT = 'https://www.veza.ru/api/v1/storage/get-file/Сертификаты%2FКлапаны%20и%20сетевые%20элементы%2F'
ELTEX = 'https://www.eltexalatau.kz/en/catalog/service_gateway_esr-1500.php'
VZLJOT = 'https://vzljot.ru/proektirovshikam/'
HAWLE = 'https://www.hawle.com/en/products/products-overview/gate-valves'
ARMTEL = 'https://armtel.com/wp-content/uploads/top-ec-ip2_rmlt.468366.009rje.pdf'
GORELTEX = 'https://exd.ru/o-kompanii/novosti/2019/zavod-goreltekh-poluchil-sertifikat-mek-ekh-na-vzryvozaschischyonnye-kabelnye-vvody-tipa-k.html'

EXTRA_RULES = [
    family('ООО "ВЕЗА"', '7720040225', 'ЕМАКС', VEZA_CERT + 'СС%20АД07.В.03425-21%20ТУ%20221%20тр%20тс%20012.pdf',
           r'(?<!\w)(?:ф-)?ЕМАКС[A-Za-zА-Яа-яЁё0-9._/+\-]*'),
    family('ООО "ВЕЗА"', '7720040225', 'ЭПВ', VEZA_CERT + 'СС%20АД07.В.03424-21%20ТУ%20198%20тр%20тс%20012.pdf'),
    family('ООО "ВЕЗА"', '7720040225', 'ВЕКТОР', 'https://www.veza.ru/produktsiya/teploenergeticheskoe-oborudovanie',
           r'(?<!\w)ВЕКТОР(?:\s*-?\s*\d[A-Za-zА-Яа-яЁё0-9._/+\-]*)?'),
    *[family('ООО "ЗАВОД ГОРЭЛТЕХ"', '7806155468', x, GORELTEX)
      for x in ('КОВ', 'КНВ', 'KOV', 'KNV')],
    *[family('ООО "Предприятие "ЭЛТЕКС"', '5410108110', x, ELTEX)
      for x in ('FH-S', 'FH-10SFP-', 'FH-DP', 'PM350-', 'РМ350-')],
    family('ЗАО "Взлет"', '7826013976', 'ТСР-М', VZLJOT),
    family('ЗАО "Взлет"', '7826013976', 'ТСР-027', VZLJOT),
    family('ЗАО "Взлет"', '7826013976', 'ПРОФИ', VZLJOT,
           r'(?<!\w)ПРОФИ(?:\s*-?\s*\d[\w.-]*(?:\s*М[АОИAОIO])?)?'),
    *[family('ООО "Хавле Индустриверке"', '4813010059', x, HAWLE,
             r'(?<!\w)' + x + r'(?:\+)?(?!\w)') for x in ('E1', 'E2', 'E3')],
    *[family('ООО "Армтел"', '7810803665', x, ARMTEL)
      for x in ('TOP-DIS-IP2', 'TOP-PAD-IP2', 'TOP-EC-IP2')],
    family('ЗАО "Арматэк"', '7804039762', 'ЭКСКЛЮЗИВ',
           'https://armatek.ru/katalog/zatvory-diskovye/zatvor-eksklyuziv/'),
    family('ООО "Техкомпания Хуавэй"', '7714186804', 'S5700-',
           'https://support.huawei.com/enterprise/en/knowledge/s5700-28x-pwr-li-ac-vid-PBI2-22347756'),
]

# Distinctive brand names can identify equipment even through a reseller.
EXTRA_GLOBAL_RULES = [
    family('Hawle', '', 'HAWLE', HAWLE),
    family('АО "Арктические технологии"', '', 'АРКТЕХ', 'https://arctex.ru/'),
    family('НПП ТЭК', '', 'РэмТЭК', 'https://old.npptec.ru/documents.php?id=969&uid=75'),
    family('ОМЗИТ', '', 'LAVART', 'https://omzit.ru/'),
    family('ОМЗИТ', '', 'ЛАВАРТ', 'https://omzit.ru/'),
]

EXTRA_RULES += [
    *[family('ООО "ВЕЗА"', '7720040225', trigger, 'https://www.veza.ru/produktsiya/' + path)
      for trigger, path in (
          ('Канал-ВЕНТ', 'kanalnoe-oborudovanie/dlya-kruglykh-kanalov/ventilyatory-dlya-kruglykh-kanalov/kanal-vent'),
          ('КПУ-1Н', 'klapany-i-setevye-elementy/protivopozharnye-klapany/kruglye-protivopozharnye-klapany/kpu-1n-kruglye'),
          ('ШСАУ-АВО', 'avtomatika/jbshchepromyshlennaya-avtomatika/shkafy-dlya-upravleniya-otopitelnymi-agregatami-i-zavesami/shsau-avo'),
          ('ПЕК-ОСА', 'ventilyatory/dopolnitelnoe-oborudovanie-ventilyatory/osevye-dopolnitelnoe-oborudovanie-ventilyatory/pek-osa'),
      )],
    family('ООО "Техкомпания Хуавэй"', '7714186804', 'S6730-',
           'https://e.huawei.com/en/products/switches/campus-switches/s6730-s'),
    family('ООО "Техкомпания Хуавэй"', '7714186804', 'Dorado',
           'https://e.huawei.com/en/documents/products/storage/f01ee520bde543ea970ce17d9471f3ef',
           r'(?<!\w)(?:OceanStor\s+)?Dorado\s*\d+[\w.-]*(?:\s+V\d+)?'),
]

EXTRA_GLOBAL_RULES += [
    *[family('Mitsubishi Electric', '', trigger,
             'https://www.mitsubishitech.co.uk/Data/Mr-Slim_Indoor/PEAD-RP/2016/PEAD-RP-JAQ/Leaflet/PEAD-RP_Power_Inverter_IPh/PEAD-RP125JAQ.pdf')
      for trigger in ('PUHZ-ZRP', 'PEAD-RP')],
]

SIMOS = 'https://simos.ru/production/product/m30ae-3u/'
EXTRA_RULES += [
    family('ООО "ВЕЗА"', '7720040225', 'AeroGuard',
           'https://www.veza.ru/produktsiya/otopitelnye-oborudovaniya/zavesy-vozdushnye/promyshlennye-zavesy-vozdushnye/aeroguard400'),
    *[family('ЗАО НТЦ "Симос"', '', token, SIMOS,
             r'(?<!\w)' + re.escape(token).replace(r'\-', '[-–]') + r'(?!\w)')
      for token in ('М30АЕ', 'ИП-12', 'ИП-13', 'ИП-14', 'ГС-01', 'ЛТ-02М', 'ЛТ-02М-01',
                    'ЛТ-02М-02', 'ЛТ-06', 'ВД-01', 'ДП-03', 'ДП-04', 'ДП-07М', 'ДП-08',
                    'ДП-09', 'СН-01', 'СН-02', 'СН-05', 'ПН-01', 'АК-02М', 'АК-03М',
                    'АР-01', 'АС-02М', 'АС-03М', 'ВЕ-01М', 'ДС-02', 'ДС-04', 'ДС-06',
                    'ДС-07', 'РТ-01', 'РТ-02', 'РТ-04', 'СА-01', 'СВ-01', 'СВ-01М',
                    'СК-01', 'СЦ-01', 'СЦ-02', 'СЧ-01', 'СЧ-03', 'ВК-03', 'ЕК-03',
                    'ЕК-04', 'КМ-15', 'КП-01', 'КС-02М', 'ОТ-07', 'ОТ-07М', 'ПК-03',
                    'УЛ-03', 'УЛ-04', 'ИП-03', 'ИП-04', 'ИП-03М', 'ИП-11', 'АК-02',
                    'АК-03', 'АС-02', 'АС-03', 'ДС-01', 'ВЕ-01')],
    family('ЗАО НТЦ "Симос"', '', 'СМ5.', SIMOS,
           r'(?<!\w)СМ5\.\d{3}\.\d{3}(?:-\d+)?(?!\w)'),
    *[family('ЗАО "ЭЛСИ Стальконструкция"', '5404173086', token,
             'https://www.elsi.ru/stalnye-opory-lep/opory-dlya-vl-10-kv/',
             r'(?<!\w)' + re.escape(token) + r'(?!\w)')
      for token in ('2АС10ПИ-1М', '2АСО10ПИ-1М', '2АУС10ПИ-1М', '2АУС10ПИ-2М',
                    '2ПС10ПИ-2М', '2ПС10ПИ-3М', 'АС10П-3М', 'АС10П-3УМ', 'АС10ПИ-1УМ',
                    'АСО10П-1М', 'АСО10П-4М', 'АСО10ПИ-1УМ', 'АУС10П-3М', 'АУС10П-3УМ',
                    'АУС10ПИ-1УМ', 'АУСО10П-1М')],
]


def supplemental_digest():
    return hashlib.sha256(json.dumps(EXTRA_RULES + EXTRA_GLOBAL_RULES,
                                     ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
