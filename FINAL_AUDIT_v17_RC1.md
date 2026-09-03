# FINAL AUDIT — v17 RC1

**Статус: RC1, не FINAL PASS.**

## Подтверждено локально

- Golden sample: 200/200 (200 строк; изменено 114; зелёных 176; жёлтых 24).
- KEEP-loss: 0.
- Idempotence failures: 0.
- Confirmed producer residual: 0.
- Синтетический текстовый PDF: колонка «Код продукции» идентична исходной; созданные redaction rectangles не пересекают её; поставщик удалён целиком; таблица и технические характеристики сохранены.
- Excel audit по контрольным 200 строкам: обработано 200, изменено 114, зелёных 176, жёлтых 24; changed rows with empty removal log — 0; KEEP-loss — 0; idempotence failures — 0; confirmed residual — 0. Дополнительный минимальный Unicode-path fixture: обработано 1, изменено 1.

## Исправленные реальные дефекты

- Защищены значения `4143086 (0)1)` и `2576244 (0)1)` в колонке «Код продукции».
- Запрещено частичное удаление фразы `Комплектация по обосновывающему документу ...` и номера ОЛ.
- Вместо удаления поставщика кусками очищается содержимое всей ячейки.
- Redaction rectangles ограничиваются геометрией колонок и не пересекают абсолютный KEEP.
- Excel-журнал формируется из реально удалённых фрагментов.
- Добавлены подтверждённые бренды из пользовательского тестирования.
- После Windows run #1 защищена структурная техническая модель `TEST-M1` от PDF-redaction.
- OCR table detection усилен устойчивым сопоставлением заголовков и резервным определением соседних колонок по форматам кода продукции и реквизитов поставщика.
- После Windows run #2 native KEEP сравнивается с исходным text layer после удаления extractor-only soft hyphen; существование модели сначала доказывается в исходном PDF.
- Для OCR абсолютный KEEP колонки «Код продукции» доказывается идентичностью пикселей геометрического crop до/после, отсутствием пересечений redaction и дополнительным OCR-контролем цифр кода; неоднозначность `Z`/`7` не подменяет проверку изображения.
- После Windows run #3 fixture сначала доказывает наличие полного `Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1` в исходном text layer; bbox каждого OCR KEEP-токена фиксируется до rasterization и затем проверяется по пикселям и пересечениям.
- После Windows run #4 synthetic table расширена с запасом под различия метрик Arial/DejaVu; каждый непустой `insert_textbox` проверяет код возврата и немедленно сообщает конкретную ячейку при переполнении. Полный набор KEEP-токенов по-прежнему валидируется в SOURCE до `process_pdf` и rasterization.
- После Windows run #5 SOURCE bbox ищется сначала по исходному написанию и вариантам дефиса (`-`, soft hyphen, Unicode/non-breaking hyphen), затем безопасным fallback по геометрии канонически совпавшей последовательности слов. KEEP-набор, source self-validation и pixel/geometry assertions не ослаблены.
- После Windows run #6 OCR table detector использует длинные растровые вертикальные/горизонтальные границы как основной структурный сигнал стандартной девятиколоночной СО. Колонки 4/5 выводятся из порядка границ независимо от OCR заголовков; отчёт содержит OCR header words с bbox, найденные границы, зоны кода/поставщика, режим и причину решения.
- После Windows run #7 добавлены контроль сохранения без redaction и контроль удалённой supplier-области через `PDF_REDACT_IMAGE_PIXELS`, а также полный pixel diff (число, bbox, max RGB delta) и отдельные diff четырёх кодовых glyph-bbox. Локально оба контроля дали `changed_pixels=0`; production supplier-redaction получила двухпунктовый защитный отступ от общей границы с колонкой кода для исключения округления image pixels.
- После Windows run #8 диагностика локализовала изменение на левой границе колонки 4. Для всех OCR-redaction абсолютный KEEP теперь расширяется на 2 pt с обеих сторон до геометрического split; отчёт показывает расстояние каждого rectangle и ближайший rectangle. Добавлены отдельные border/glyph diffs. Локальная имитация полного supplier+brand pixel-redaction с guard дала `changed_pixels=0` для всей code-column.
- После Windows run #9 устранён остаточный первый supplier glyph: диагностика assertion показывает source/residual bbox, границу, redaction rectangles и расстояние от кода. Проверка offsets показала изменение code-border при `0.8 pt` и строго `0` changed pixels начиная с `1.0 pt`; поэтому правый KEEP guard и начало supplier redaction согласованы на доказанном pixel-safe значении `1.0 pt`, сохраняя левый guard `2.0 pt`.
- После Windows run #10 security assertion отделяет OCR-артефакт линии (`in`) от остатка поставщика: для residual фиксируются bbox, доступная confidence, source OCR words, crop diff, пересечение source glyph/grid и число тёмных пикселей вне grid. Все 11 source supplier word-bbox обязаны стать физически белыми, чувствительные токены отсутствовать, а ядро пяти grid-линий и вся code-column — остаться pixel-identical.
- После Windows run #11 каждый тёмный residual-пиксель вне core grid получает точные page-координаты/bbox и pixel provenance. Артефакт допустим только как неизменный source anti-alias у grid, не принадлежащий ни одному source supplier glyph; изменённый off-grid пиксель или любой supplier-glyph remnant остаётся безусловным FAIL.
- После Windows run #12 верхний/нижний supplier inset увеличен с `0.8` до `2.0 pt`, не меняя доказанную горизонтальную границу code KEEP. Exact grid assertion сохранён. При FAIL отчёт содержит changed pixels/coordinates/bbox/max RGB, dark count, min grayscale, continuity, gaps, core hash, thickness и расстояния до supplier redaction; вертикальная граница CODE/SUPPLIER контролируется отдельно.
- После Windows run #13 диагностика доказала вариант B: horizontal core index/SHA, minimum grayscale, полная continuity `880/880`, `max_gap=0` и minimum thickness `4` совпали; менялись только 3–6 edge/пересечных anti-alias pixels из ~3570 dark pixels. Неверный byte-compare горизонтали заменён на бездопусковый structural invariant этих метрик; вертикальная CODE/SUPPLIER граница остаётся byte-identical.
- После Windows run #14 найден отдельный generic OCR span-redaction, расширяемый на `y±0.8` после cell geometry. Production теперь создаёт KEEP guards для всех raster vertical/horizontal boundaries и вычитает их из supplier и каждого generic OCR rectangle до `PDF_REDACT_IMAGE_PIXELS`. Диагностика хранит rule label, original/expanded/final bbox и расстояние до grid; уничтоженный guard-ом фрагмент переводит страницу в REVIEW.
- После Windows run #15 до redaction строятся OCR TECHNICAL KEEP glyph rectangles из production protected regex, технических word-patterns и fuzzy required phrase+OL. Supplier и generic rectangles последовательно вычитаются из CODE, TECH/OL и GRID KEEP. Диагностика связывает пересечение с rule/span, original glyph, expanded/final bbox; конфликт supplier/TECH переводится в REVIEW.
- После Windows run #16 устранено ложное TECH KEEP поставщика: required phrase ранее брала следующие 10 global OCR words и захватывала code/supplier (`АО "ТЕСТОВЫЙ ЗАВОД"`). Теперь продолжение phrase/ОЛ ограничено той же геометрической колонкой, а нормативные соседи — той же строкой, колонкой и допустимым gap. Weak OCR words внутри supplier не создают KEEP; strong technical token защищает только собственный bbox.
- После Windows run #17 TECH KEEP получил отдельный raster guard `2.0 pt`: геометрически непересекающийся соседний `PDF_REDACT_IMAGE_PIXELS` больше не затрагивает anti-alias/core glyph. Для каждого KEEP failure выводятся exact glyph/keep bbox, nearby rule/span original/expanded/final rectangles, расстояния, save-only/nearest-redaction controls и раздельный core/edge pixel diff. Локальный brand-only control дал IP66 `changed_pixels=0`.
- После Windows run #18 guard увеличен до следующего проверяемого минимального значения `2.5 pt`, поскольку Windows доказал недостаточность `2.0 pt`. Failure diagnostics автоматически воспроизводит save-only, nearest-brand-only и halo sweep `1.0/1.5/2.0/2.5 pt`. Локально при `2.5 pt` IP66 остаётся pixel-identical, а соседний бренд физически очищается (`1209 → 0` dark pixels).
- После пересмотра архитектуры таблица сначала преобразуется в cell permission model: `product_code=ABSOLUTE KEEP`, `supplier=DELETE ALLOWED`, колонки наименования/типа=`description DELETE ALLOWED`, остальные=`NEUTRAL`. Любой candidate сначала пересекается со своей разрешённой cell-zone и только затем из него вычитаются CODE, TECH/required phrase/OL и GRID KEEP. Финальный отчёт отдельно считает любой rectangle вне разрешённой зоны; OCR TECH preflight агрегирует ошибки всех токенов за один прогон.

## Ещё требуется

- Обязательный OCR fixture выполняется на Windows runner с реальным Tesseract RU+EN; локально Tesseract отсутствует.
- Windows EXE, ZIP и packaged self-test должны получить SUCCESS в GitHub Actions.
- После artifact пользователь повторно проверит семь закрытых производственных PDF локально.

FINAL PASS до этих проверок не присваивается.
