"""Reviewed technical KEEP definitions, with evidence and narrow grammars.

These preserve source characters. They do not certify a manufacturer or hide
other unknown tokens. Context checks apply to the whole equipment description.
"""
import re

MANOTOM = 'https://manotom.ru/upload/docs/Перечень%20сокращений%20и%20схем%20условных%20обозначений.pdf'
VELAN = 'https://velan.ru/product/vzryvozashchishchennoe/kabelnye-vvody/kabelnye-vvody-vk-vel/'
KEAZ = 'https://keaz.ru/catalog/ustroystva-na-din-reyku/modulnie-avtomaticheskie-vikluchateli/va47-29-modulnie-avtomaticheskie-vikluchateli-na-toki-do-63a-noviy'


def keep(key, pattern, reason, source, context=''):
    return dict(id=key, regex=pattern, reason=reason, source=source, context=context, decision='KEEP')


REVIEWED_KEEP = [
    keep('thread_metric', r'(?<!\w)[МM]\d{1,3}\s*[xх×]\s*\d+(?:[.,]\d+)?(?:-[468][gGhH])?(?!\w)',
         'Размер метрической резьбы, шаг и поле допуска', MANOTOM),
    keep('thread_pipe', r'(?<!\w)(?:G|R|Rc|Rp|К|K)\s*(?:\d+\s+)?\d+/\d+(?!\w)',
         'Размер трубной резьбы', MANOTOM),
    keep('pressure_units', r'(?<!\w)(?:[+−-]?\d+(?:[.,]\d+)?\s*)?(?:MPa|kPa|kgf/cm[2²]|кгс/см[2²]|psi)(?!\w)',
         'Единица измерения давления', MANOTOM),
    keep('pressure_designation', r'(?<!\w)(?:Ду|Ру|[РP]N)\s*\d+(?:[.,]\d+)?(?:\s*(?:МПа|MPa|мм))?(?!\w)',
         'Условный проход или номинальное давление', 'Исходная запись: явная метка Ду/Ру/PN и значение'),
    keep('colour_temperature', r'(?<!\w)\d{4,5}\s*[KК](?!\w)',
         'Цветовая температура светильника в кельвинах', 'Исходная запись: светильник и числовая температура света',
         r'светильник|прожектор|ламп[аы]|цветов\w*\s+температур'),
    keep('voltage_kind', r'(?<!\w)\d+(?:[.,]\d+)?\s*(?:V?[AА][CС]|V?DC|В\s*(?:AC|DC))(?!\w)',
         'Напряжение и вид тока', VELAN),
    keep('engineering_units', r'(?<!\w)(?:\d+(?:[.,]\d+)?\s*)?(?:мм[2²]|см[2²]|м[2²]|м[3³]/ч(?:ас)?|мг/м[3³]|Дж/см[2²]|куб\.м/час|об/мин|кН|дБ|Ah)(?!\w)',
         'Единицы площади, расхода, концентрации, ударной вязкости, частоты вращения, силы, звука и ёмкости',
         'Исходная запись: общепринятая единица измерения'),
    keep('angular_value', r'(?<!\w)\d+(?:[.,]\d+)?\s*(?:град(?:\.(?:\s*[CС])?)?|гр\.[CС])(?!\w)',
         'Числовой угол или температура', 'Исходная запись: число и градусы'),
    keep('standard_sp', r'(?<!\w)СП\s*\d+(?:\.\d+)+(?!\w)',
         'Ссылка на свод правил', 'Исходная запись: СП и номер документа'),
    keep('ansi', r'(?<!\w)(?:ANSI(?:\s*(?:150|300|400|600|900|1500|2500))?|ASME)(?!\w)',
         'Обозначение системы стандартов или класса давления', 'Исходная запись: ссылка ANSI/ASME; не обозначение завода'),
    keep('chemical', r'(?<!\w)(?:CO2|H2S|PTFE|PEEK|ПММА)(?!\w)',
         'Химическая формула или материал', MANOTOM),
    keep('rack_height', r'(?<!\w)(?:[1-9]|[1-4]\d)U(?!\w)',
         'Высота оборудования в стойке', 'Исходная запись: шкаф/стойка и единица U', r'шкаф|стойк|панел|органайзер|сервер'),
    keep('electrical_poles', r'(?<!\w)[1-4][PР](?:\+N)?(?!\w)',
         'Количество полюсов электрического аппарата', KEAZ, r'автомат|выключател|щит|полюс|дифференциал|шкаф|коммутацион'),
    keep('breaker_curve', r'(?<!\w)(?:[1-4])?[BCDВСД](?:1|2|3|4|6|8|10|13|16|20|25|32|40|50|63|80|100|125)(?!\w)',
         'Характеристика расцепления и номинальный ток', KEAZ, r'автомат|выключател|расцепител|коммутацион|щит|шкаф'),
    keep('ex_dust', r'(?<!\w)(?:[012]\s*)?Ex\s*(?:t[abc]?|[dei]a?|[dei]b?)\s+III[ABC]\s*(?:T\d+\s*(?:°?[CС]|град\.[CС])\s*)?(?:D[abc])?(?!\w)',
         'Маркировка взрывозащиты для пылевой среды', VELAN),
    keep('ex_groups', r'(?<!\w)(?:III[ABC]|II[ABC](?:\+H2)?|[GD][abc])(?!\w)',
         'Группа смеси или уровень взрывозащиты', VELAN, r'\b[EЕ][xх]|взрывозащи'),
    keep('climate_extended', r'(?<!\w)(?:ОМ|В|Т)\s*[1-5](?:[.,][15])?(?!\w)',
         'Климатическое исполнение; только в явном климатическом контексте', VELAN,
         r'климатическ|кабельн\w*\s+ввод|ввод\w*\s+кабель|ВК-[ЛНС]-'),
    keep('quantity_ending', r'(?<!\w)\d+-(?:х|ти|ми|стоечный|мм)(?!\w)',
         'Количество или размер с русским окончанием', 'Исходная запись: числовое определение'),
    keep('nameplate', r'(?<!\w)(?:ВЫХОД|ПУСК|СТОП|ЗАГАЗОВАННОСТЬ|ПОЖАРНЫЙ\s+КРАН|ГАЗ\s+НЕ\s+ВХОДИ|ГАЗ\s+УХОДИ|АВТОМАТИКА\s+ОТКЛЮЧЕНА)(?!\w)',
         'Текст информационной надписи, не торговое наименование', 'Исходная запись: кнопка или сигнальное табло',
         r'табло|кноп|надпис|оповещател|сигнализ'),
    keep('metric_port', r'(?<!\w)[МM](?:8|10|12|16|20|25|32|40|50|63|75|90|110)(?!\w)',
         'Метрическая присоединительная резьба в описании прибора или кабельного ввода', MANOTOM,
         r'резьб|штуцер|кабельн\w*\s+ввод|ввод\w*\s+кабель|термопреобраз|датчик\s+давлен'),
    keep('ral_palette', r'(?<!\w)RAL\s+(?:D2\s+DESIGN|DESIGN(?:\s+SYSTEM(?:\s+plus)?)?|CLASSIC|EFFECT)(?!\w)',
         'Название палитры цвета; не марка покрытия', 'https://www.ral-farben.de/en/all-ral-colours'),
    keep('fan_temperature', r'(?<!\w)[ТT](?:80|100|120|200|400|600)(?!\w)',
         'Температурный режим вентилятора, явно обозначенный в исходном описании',
         'Исходная запись: режим работы Т80 и описание вентилятора', r'режим\s+работы\s+[ТT]\d'),
    keep('velan_ex_code', r'(?<!\w)Ex[de]G(?!\w)',
         'Вид взрывозащиты кабельного ввода',
         'https://velan.ru/product/vzryvozashchishchennoe/kabelnye-vvody/kabelnye-vvody-vk-vel/',
         r'кабельн\w*\s+ввод|ввод\w*\s+кабель|ВК-'),

    keep('manotom_embedded_climate', r'(?<=МП[234]-У)(?:УХЛ|У|Т)[1-5](?:[.,]\d)?(?!\w)',
         'Климатическое исполнение после обозначения манометра; сохранение слитной записи', MANOTOM),

    keep('dimensions_with_axis', r'(?<!\w)\d+(?:[.,]\d+)?(?:\([hwdlншд]\))?(?:\s*[xх×]\s*\d+(?:[.,]\d+)?(?:\([hwdlншд]\))?){1,3}(?!\w)',
         'Габаритные размеры с обозначением оси в скобках',
         'Исходная запись: размер/габарит и последовательность числовых размеров', r'размер|габарит'),
    keep('fiber_link', r'(?<!\w)ВОЛС(?!\w)',
         'Волоконно-оптическая линия связи', 'Исходная запись: линия связи/подключение оптического кабеля',
         r'кабел|подключени|оптическ|связ'),
    keep('surge_device', r'(?<!\w)УЗИП(?!\w)',
         'Устройство защиты от импульсных перенапряжений, без индекса модели',
         'Исходная запись: электрический щит и защита от перенапряжений', r'щит|перенапряж|импульс'),
    keep('spare_kit', r'(?<!\w)ЗИП(?!\w)',
         'Комплект запасных частей, инструмента и принадлежностей',
         'Исходная запись: комплект поставки', r'комплект\w*\s+ЗИП\b|\bЗИП\s+(?:в\s+)?комплект'),
    keep('diesel_unit', r'(?<!\w)ДГУ(?!\w)',
         'Дизель-генераторная установка, без индекса модели',
         'Исходная запись: дизельное или генераторное оборудование', r'дизель|генератор'),

]

COMPILED_KEEP = [(row, re.compile(row['regex'], re.I), re.compile(row['context'], re.I) if row['context'] else None)
                 for row in REVIEWED_KEEP]


def reviewed_ranges(text):
    for row, pattern, context in COMPILED_KEEP:
        if context is None or context.search(text):
            for match in pattern.finditer(text):
                if row['id'] in {'fiber_link', 'surge_device', 'spare_kit', 'diesel_unit'}:
                    prefix = text[max(0, match.start() - 50):match.start()]
                    if re.search(r'(?:ООО|ПАО|ОАО|ЗАО|АО)\s*[\"«]?\s*$', prefix, re.I):
                        continue
                yield match.start(), match.end(), row
