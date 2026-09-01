import csv
from pathlib import Path

import pytest

from mtr_core import Anonymizer, keep_signature


@pytest.fixture(scope="module")
def az():
    return Anonymizer()


@pytest.mark.parametrize("technical", [
    "IP66", "УХЛ1", "Ex d IIC T6", "сталь 12Х18Н10Т", "DN50 PN16",
    "12345-ОЛ2", "Комплектация по обосновывающему документу 12345-ОЛ2",
])
def test_protected_characteristics_survive(az, technical):
    result = az.anonymize(f"Клапан Унипол {technical}")["text"]
    assert technical in result
    assert "Унипол" not in result


def test_service_letter_removed_but_ol_kept(az):
    source = "Клапан Унипол DN50 12345-ОЛ2 № АБВ-123 от 12.03.2024"
    result = az.anonymize(source)["text"]
    assert "12345-ОЛ2" in result
    assert "АБВ-123" not in result and "12.03.2024" not in result


@pytest.mark.parametrize("brand", [
    "Унипол", "Гиперфлоу", "Метран", "Вэлан", "Ризур", "Рубеж",
    "Пензтяжпромарматура", "Волжский трубный завод",
])
def test_confirmed_brand_removed(az, brand):
    result = az.anonymize(f"Изделие {brand} IP66 УХЛ1 DN50 PN16")["text"]
    assert brand.casefold() not in result.casefold()
    assert all(value in result for value in ("IP66", "УХЛ1", "DN50", "PN16"))


@pytest.mark.parametrize("source", [
    "Кабель Унипол 3x2,5 мм 0,66 кВ УХЛ1 ГОСТ 31996",
    "Труба Волжский трубный завод 57x3,5 сталь 20 ГОСТ 8732 PN16",
    "Отвод Унипол DN100 PN25 сталь 09Г2С ГОСТ 17375",
])
def test_cable_pipe_and_fittings_are_readable(az, source):
    result = az.anonymize(source)["text"]
    assert result and keep_signature(source) == keep_signature(result)
    assert "  " not in result and not result.endswith((",", "-", "("))


def test_real_control_sample_200(az):
    path = Path("Контроль_200_v15_FINAL.csv")
    with path.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle, delimiter=";"))
    assert len(rows) >= 200
    for row in rows[:200]:
        result = az.anonymize(row["Исходное"], row["Код Autodocs"], row["Завод"])
        assert result["text"]
        assert keep_signature(row["Исходное"]) == keep_signature(result["text"])
        assert az.anonymize(result["text"], row["Код Autodocs"], row["Завод"])["text"] == result["text"]
