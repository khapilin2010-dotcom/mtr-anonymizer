# FINAL AUDIT v18 SIMPLIFIED

## Реализованные инварианты

- Physical PDF redaction создаётся только внутри разрешённой ячейки основной таблицы.
- Графа 4 `Код продукции` — абсолютный KEEP.
- Для обычной строки графа 5 очищается целиком с сохранением grid.
- Для графы 3 применяется allow-list: ГОСТ и полная конструкция `Комплектация … ОЛ`.
- Любое ТУ удаляется из граф 2 и 3.
- Строка `Инновационное оборудование` целиком относится к абсолютному KEEP.
- После Windows run #1 ГОСТ KEEP графы 3 восстанавливается по original word glyph bbox при platform-specific hyphen extraction; `search_for()` больше не является единственным источником геометрии.
- Redaction граф 2/3/5 перед `add_redact_annot` повторно clipping-уется против полной CODE KEEP зоны и проверяется assertion с нулевой площадью пересечения.

## Release gates

- Требуются полный pytest с реальным OCR RU+EN, Windows EXE, ZIP и packaged self-test.
- После artifact пользователь проверяет семь закрытых PDF и четыре Excel.

FINAL PASS до реальных acceptance-тестов не присваивается.
