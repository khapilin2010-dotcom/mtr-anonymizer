"""Excel-only confirmed catalog additions. No customer rows or frequencies.

Sources establish product family identity, not fitness for a particular project.
Attached model suffixes are removed subject to the engine's absolute KEEP.
"""
import re
import hashlib
import json
from excel.reviewed_technical import REVIEWED_KEEP


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
    *[family('Световые технологии', '', 'PRIZMA/' + variant,
             'https://www.ltcompany.com/en/series/prizma-' + variant.lower(),
             r'(?<!\w)PRIZMA/' + variant + r'(?!\w)(?:\s+(?:300|595|1200)(?!\d)(?!\s*[xх×]))?(?:\s+(?:HFD\s+)?EM\b)?')
      for variant in ('R', 'S')],
    family('IEK GROUP', '', 'IEK', 'https://www.iek.ru/company/brandbook/',
           r'(?<!\w)IEK(?!\w)'),
    family('IEK GROUP', '', 'ИЭК', 'https://www.iek.ru/company/',
           r'(?<!\w)ИЭК(?!\w)'),
    family('Schneider Electric', '', 'Schneider Electric', 'https://www.se.com/ww/en/',
           r'(?<!\w)Schneider\s+Electric(?!\w)'),
    family('КЭАЗ', '', 'OptiDin', 'https://keaz.ru/catalog/ustroystva-na-din-reyku/modulnie-avtomaticheskie-vikluchateli',
           r'(?<!\w)OptiDin(?!\w)(?:\s+(?:[BВ][MМ]63|DМ?63|DM63)[\w./+\-]*)?'),
    family('Группа Астра', '', 'Astra Linux', 'https://astralinux.ru/os/',
           r'(?<!\w)Astra\s+Linux(?:\s+(?:Special|Common)\s+Edition)?(?!\w)'),
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
    family('ЭЛТЕКС', '', 'Eltex', 'https://eltex-co.com/'),
    family('Huawei', '', 'Huawei', 'https://e.huawei.com/en/products/switches/campus-switches/s6730-s'),
    family('РАМЭК', '', 'RAMEC', 'https://www.ramec.ru/produktyi-i-uslugi/kompyuternaya-texnika/kompyuteryi-ramec.html'),
    family('ВЕЗА', '', 'VEZA', 'https://www.veza.ru/'),
    family('Армтел', '', 'Armtel', 'https://armtel.com/ru/product/pult-dispetcherskij-top-dis-ip2/'),
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


# RC4: primary catalog review. An electrical series may be produced by several
# factories; recognizing the series must not invent a factory attribution.
EXTRA_GLOBAL_RULES += [
    family('', '', 'ВА47-29',
           'https://keaz.ru/catalog/ustroystva-na-din-reyku/modulnie-avtomaticheskie-vikluchateli/va47-29-modulnie-avtomaticheskie-vikluchateli-na-toki-do-63a-noviy',
           r'(?<!\w)ВА47-29(?!\w)'),
    family('', '', 'АВДТ32', 'https://keaz.ru/catalog', r'(?<!\w)АВДТ32(?!\w)'),
    family('', '', 'ВН-32',
           'https://keaz.ru/catalog/ustroystva-na-din-reyku/modulnie-vikluchateli-razediniteli-vikluchateli-nagruzki/vn32',
           r'(?<!\w)ВН-32(?!\w)'),
    family('Рубеж', '', 'ПАСН.',
           'https://products.rubezh.ru/download/file/a84b08ec-58d1-11f0-95f1-d4f5ef944508/',
           r'(?<!\w)ПАСН[. ]\s*\d{6}\.\d{3}(?:-\d+)?(?!\w)'),
]

RUBEZH_MODELS = (
    ('МИ-R2', 'mi_r2_1-7076'), ('МИБ-R2', 'mib_r2_1-7539'),
    ('МДУ-R2', 'mdu_r2_isp_220-3363'), ('МВП-R2', 'mvp_r2-3372'),
    ('МПО-PFM-R2', 'mpo_pfm_r2-3320'), ('МСВ-R2', 'msv_r2-3315'),
    ('МСП-R2', 'msp_r2-3314'), ('ОПОП 1-R2', 'opop_1_r2-3371'),
    ('ОПОП 124-R2', 'opop_124_r2-3370'), ('ОПОП 124Б-R2', 'opop_124b_r2-3312'),
    ('ОПОП 2-R2', 'opop_2_r2-3380'), ('РМ1М-R2', 'rm1m_r2-3325'),
    ('РМ2-R2', 'rm2_r2-3376'), ('РМ4-R2', 'rm4_r2-3373'),
    ('АМ1-R2', 'am1_r2-3361'), ('АМ4-R2', 'am4_r2-3378'),
    ('ШУЗ-R2', 'shuz_r2-3357'), ('ШУН/В-R2', 'shun_v_r2-3355'),
)
EXTRA_GLOBAL_RULES += [family('Рубеж', '', token, 'https://products.rubezh.ru/products/' + path + '/',
                              r'(?<!\w)' + re.escape(token).replace(r'\ ', r'\s+') + r'(?!\w)')
                       for token, path in RUBEZH_MODELS]
# AM names occur in Latin lettering in the catalog and Cyrillic in exports.
EXTRA_GLOBAL_RULES += [family('Рубеж', '', token,
    'https://products.rubezh.ru/download/file/a84b08ec-58d1-11f0-95f1-d4f5ef944508/',
    r'(?<!\w)' + token + r'(?!\w)') for token in ('AM1-R2', 'AM4-R2')]

EXTRA_RULES += [
    family('ООО "ВЕЗА"', '7720040225', 'Канал-ГКК',
           'https://www.veza.ru/produktsiya/kanalnoe-oborudovanie/dlya-kruglykh-kanalov/shumoglushiteli-dlya-kruglykh-kanalov/kanal-gkk',
           r'(?<!\w)Канал-ГКК(?!\w)'),
    family('ООО "ВЕЗА"', '7720040225', 'Канал-ГКП',
           'https://www.veza.ru/produktsiya/kanalnoe-oborudovanie/dlya-pryamougolnykh-kanalov/shumoglushiteli-dlya-pryamougolnykh-kanalov/kanal-gkp',
           r'(?<!\w)Канал-ГКП(?!\w)'),
    family('ООО "ВЕЗА"', '7720040225', 'ВКОП',
           'https://www.veza.ru/produktsiya/ventilyatory/protivopozharnye/osevye-protivopozharnye/vkop',
           r'(?<!\w)ВКОП(?:0)?(?!\w)'),
    family('ООО "ВЕЗА"', '7720040225', 'AIRMATE',
           'https://www.veza.ru/produktsiya/vozdukhoobrabatyvayushchie-ventilyatsionnye-agregaty/kompaktnye-ventilyatsionnye-ustanovki/podvesnye-ustanovki-airmate/airmate-4000',
           r'(?<!\w)AIRMATE-(?:800|1200|2000|4000|6000)(?!\w)'),
    family('ООО "ВЕЗА"', '7720040225', 'РОН110',
           'https://www.veza.ru/produktsiya/klapany-i-setevye-elementy/klapany-klapany-setevye/ustroystva-vozduhopriemnye/ron110-pryamougolnye',
           r'(?<!\w)РОН\s*110(?!\w)'),
    family('АО "Вэлан"', '2619000120', 'ВЭЛ',
           'https://velan.ru/product/vzryvozashchishchennoe/kabelnye-vvody/kabelnye-vvody-vk-vel/',
           r'(?<=ВК-Л-)ВЭЛ|(?<=ВК-Н-)ВЭЛ|(?<=ВК-С-)ВЭЛ'),
]


# Series names only: light source, colour temperature and dimensions survive.
EXTRA_RULES += [family('Световые технологии', inn, token,
                       'https://www.ltcompany.com/series/' + slug,
                       r'(?<!\w)' + token + r'(?!\w)')
                for inn in ('6229028102', '7715723321')
                for token, slug in (('LYRA', 'lyra-led'), ('LHT', 'lht'),
                                    ('CLEAN', 'clean'), ('TOTEM', 'totem'),
                                    ('GLOBUS', 'globus-led'), ('PROTON', 'proton-led-exd'),
                                    ('NORTH', 'north'), ('COMP', 'comp'))]
EXTRA_GLOBAL_RULES += [family('Морозовский химический завод', '', 'АРМОКОТ F100',
    'https://tdmhz.ru/wp-content/uploads/TI-Armokot-F100-metall.pdf',
    r'(?<!\w)АРМОКОТ\s*®?\s*F\s*100(?!\w)')]


EXTRA_RULES += [family('ООО "ЗАВОД ГОРЭЛТЕХ"', '7806155468', 'СГР01',
    'https://exd.ru/produciya/osvetitelnoe-oborudovanie/vzryvozaschischennye-svetilniki/svetilniki-perenosnye-sgr01.html',
    r'(?<!\w)СГР01(?!\w)')]

EXTRA_GLOBAL_RULES += [
    family('НТЦ Протей', '', 'ПАМР.',
           'https://docs.sgep-it.ru/upload/uf/018/606jvbsf3k8erwbrjumgdgepzpdx4uu9/Prikaz-Minpromtorga-Rossii-ot-03.03.2025-_1023-o-prisvoenii-i-podtverzhdenii-TKO-statusa-TORP.pdf',
           r'(?<!\w)ПАМР[. ]\s*\d{6}\.\d{3}(?:-\d+)?(?!\w)'),
    family('НПК Эталон', '', 'ЮВМА.420520.004',
           'https://npk-etalon.ru/upload/iblock/d61/6tuon4snnb64y6di54kwnjq4mznw45d2/izv_combo.pdf',
           r'(?<!\w)ЮВМА[. ]\s*420520\.004(?:-\d+)?(?!\w)'),
]

# ESKD product/KD identity: organization code + classification + registration.
# This establishes an identifying designation, not the name of its developer.
EXTRA_GLOBAL_RULES += [family('', '', 'Код изделия / КД по ГОСТ 2.201',
    'https://www.gostinfo.ru/Qa/Details/460',
    r'(?<!\w)(?!ГОСТ)(?-i:[А-ЯЁ]{4})[. ]\s*\d{6}\.\d{3}(?:-\d{1,3})?(?:(?:ПС|РЭ|СБ|ТУ|ИЭ)\d{0,2})?(?!\w)')]


# RC5: exact catalog families; no supplier inferred from shared electrical series.
EXTRA_GLOBAL_RULES += [
    family('', '', 'СЭТ-4ТМ', 'https://nzif.ru/uploads/sel/psch4tm03m/ruk_03_02_m.pdf',
           r'(?<!\w)СЭТ-4ТМ(?:\.\d{2}[МM]?(?:\.\d{2})?)?(?!\w)'),
    family('', '', 'ПМЛ', 'https://keaz.ru/catalog/kontaktor-pusk/kontaktori-puskateli-rele/pml-kontaktori-puskateli-s-katushkami-upravleniya-peremennim-i-postoyannim-tokom-na-toki-ot-10a-do-400',
           r'(?<!\w)ПМЛ(?:-\d[\w.-]*)?(?!\w)'),
    family('', '', 'РТЛ', 'https://keaz.ru/catalog/kontaktor-pusk/kontaktori-puskateli-rele/rtl-rele-peregruzki-teplovie-na-toki-ot-25a-do-500',
           r'(?<!\w)РТЛ-\d[\w.-]*(?!\w)'),
    family('', '', 'ВА57', 'https://keaz.ru/catalog/automat/avtomaticheskie-viklyuchateli-v-litom-korpuse/va57-blochnie-avtomaticheskie-vikluchateli-na-toki-ot-16a-do-630a',
           r'(?<!\w)ВА57-\d[\w.-]*(?!\w)'),
    family('Эридан', '', 'ИП535-07е', 'https://eridan.ru/catalog/ex-izveshateli/manual_ipr/',
           r'(?<!\w)ИП\s*535-07[еe][аa]?(?:-(?:RS|R2|R3|О|O))?(?!\w)'),
]
EXTRA_RULES += [
    family('ООО "ВЕЗА"', '7720040225', 'MV220',
           'https://ventilyatorov.ru/files/protizopozharniye-klapany-i-oborudovaniye-veza.pdf',
           r'(?<!\w)[МM][VВBS](?=(?:220|24)(?!\w))'),
    family('ООО "АЭРО ИКСИА"', '3257017280', 'СКВ',
           'https://nanocertifica.ru/news/provedena-sertifikatsiya-skv-aero-iksian/',
           r'(?<!\w)СКВ(?!\w)'),
    family('АО "Сантехпром"', '7718014490', 'РБС',
           'https://santexprom.ru/catalog/radiator-rbs-500-ch-95-a01.html',
           r'(?<!\w)РБС(?=-\d)'),
]
# Cyrillic spelling of the documented TOTEM series.
EXTRA_RULES += [family('Световые технологии', inn, 'ТОТЕМ',
                       'https://www.ltcompany.com/series/totem', r'(?<!\w)ТОТЕМ(?!\w)')
                for inn in ('6229028102', '7715723321')]


LT_CATALOG = 'https://rospolus.ru/doc/svet/svetovie_tech/LT_CATALOG.pdf'
EXTRA_RULES += [family('Световые технологии', inn, token, LT_CATALOG,
                       r'(?<!\w)' + re.escape(token) + r'(?!\w)')
                for inn in ('6229028102', '7715723321')
                for token in ('FREGAT', 'PIPE', 'MIZAR', 'LZ.OPL', 'ALD', 'URAN', 'INSEL')]
EXTRA_RULES += [family('Световые технологии', inn, token, LT_CATALOG, pattern)
                for inn in ('6229028102', '7715723321')
                for token, pattern in (
                    ('ПЭУ 010', r'(?<!\w)ПЭУ\s*(?:00[1-8]|010|011|012|09[1-4])(?!\w)'),
                    ('ППБ 0001', r'(?<!\w)ППБ\s*000[1-4](?!\w)'))]
EXTRA_RULES += [family('НПП ЭЛЕМЕР', '5044003551', token,
                       'https://www.elemer.ru/catalog/datchiki-temperatury/termometry-soprotivleniya/ts/',
                       r'(?<!\w)' + re.escape(token) + r'(?!\w)')
                for token in ('ТС-1088', 'ТС-1288', 'ТС-1388', 'ТС-0295')]

def supplemental_digest():
    return hashlib.sha256(json.dumps(EXTRA_RULES + EXTRA_GLOBAL_RULES + REVIEWED_KEEP,
                                     ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()
