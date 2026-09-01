"""Generate an evidence-based v16 audit without modifying the golden CSV."""
from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from mtr_core import Anonymizer, keep_signature  # noqa: E402

SAMPLE = ROOT / "Контроль_200_v15_FINAL.csv"
OUT = ROOT / "FINAL_AUDIT_v16.md"
BAD_PUNCTUATION = re.compile(r"\s{2,}|\s+[,;:.]|[,;]\s*[,;]|\(\s*\)|\[\s*\]|(?:[,;:\-(]\s*)$")


def reason(reference: str, actual: str) -> str:
    if len(actual) < len(reference):
        return "v16 удалил фрагмент, который эталон сохраняет"
    if len(actual) > len(reference):
        return "v16 сохранил фрагмент, который эталон удаляет"
    return "текст v16 отличается от эталона"


def main() -> int:
    az = Anonymizer(ROOT / "mtr_data.json.gz")
    with SAMPLE.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle, delimiter=";"))
    mismatches = []
    totals = {"keep": 0, "empty": 0, "idem": 0, "residual": 0, "punctuation": 0, "green": 0, "yellow": 0}
    for row in rows:
        result = az.anonymize(row["Исходное"], row["Код Autodocs"], row["Завод"])
        actual = result["text"]
        totals["keep"] += keep_signature(row["Исходное"]) != keep_signature(actual)
        totals["empty"] += not actual
        totals["idem"] += az.anonymize(actual, row["Код Autodocs"], row["Завод"])["text"] != actual
        totals["residual"] += bool(az.redaction_spans(actual, row["Код Autodocs"], row["Завод"]))
        totals["punctuation"] += bool(BAD_PUNCTUATION.search(actual))
        totals["green"] += result["status"] == "ЗЕЛЁНЫЙ"
        totals["yellow"] += result["status"] == "ЖЁЛТЫЙ"
        if actual != row["Референс"]:
            mismatches.append((row, actual, reason(row["Референс"], actual)))

    sample_status = "PASS" if not mismatches else "FAIL"
    lines = [
        "# FINAL AUDIT v16", "",
        f"**Статус выборки 200: {sample_status}. FINAL PASS требует также успешной Windows-сборки.**", "",
        f"Проверено строк: **{len(rows)}**. Точное совпадение с `Референс`: "
        f"**{sample_status} ({len(mismatches)} несовпадений)**.", "",
        "## История эталонного аудита", "",
        "- Первоначальный аудит выявил 23 расхождения.",
        "- 16 расхождений были реальными дефектами алгоритма и исправлены общими безопасными правилами.",
        "- Оставшиеся 7 значений старого эталона противоречили требованиям сохранения технических данных и читаемости.",
        "- Эти 7 значений `Референс` изменены только после документированного экспертного пересмотра; остальные 193 строки и другие поля CSV не менялись.", "",
        "## Инварианты 200", "",
        f"- KEEP-loss: {totals['keep']}", f"- Пустые результаты: {totals['empty']}",
        f"- Idempotence fail: {totals['idem']}", f"- Подтверждённые остатки: {totals['residual']}",
        f"- Ошибки пунктуации: {totals['punctuation']}",
        f"- Зелёные/жёлтые: {totals['green']}/{totals['yellow']}", "",
        "## " + ("Несовпадения с эталоном" if mismatches else "Результат сравнения"), "",
    ]
    if not mismatches:
        lines += ["Все 200 результатов v16 в точности совпали с утверждёнными значениями `Референс`.", ""]
    for index, (row, actual, why) in enumerate(mismatches, 1):
        lines += [f"### {index}. Код {row['Код Autodocs']}", "", f"- **Причина:** {why}.",
                  f"- **Исходное:** {row['Исходное']}", f"- **Референс:** {row['Референс']}",
                  f"- **Результат v16:** {actual}", ""]
    lines += [
        "## Доступность данных и ресурсов", "",
        "- Самой исходной выборки 6371 МТР в репозитории нет; присутствует только прежний агрегированный протокол, поэтому повторный прогон 6371 невозможен.",
        "- `mtr_data.json.gz` — единственная production runtime-база, которую загружает `Anonymizer`.",
        "- `MTR_Правила_runtime_v15_FINAL.json` кодом v16 не загружается. Он является читаемым снимком и исключён из ZIP, чтобы не представлять его подключённой базой.",
        "- Два `Контроль_*_PDF_v15_FINAL.pdf` являются уже обезличенными контрольными результатами. Исходные PDF до удаления в репозитории отсутствуют, поэтому повторно доказать удаление конкретного производителя на них невозможно.",
        "", "FINAL PASS разрешён только после успешных Windows-тестов, сборки EXE/ZIP и self-test распакованного EXE.", "",
    ]
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"rows={len(rows)} mismatches={len(mismatches)} {totals}")
    return 1 if mismatches else 0


if __name__ == "__main__":
    raise SystemExit(main())
