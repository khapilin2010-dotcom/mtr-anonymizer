"""Read the existing compiled manufacturer database without processing documents."""
import gzip
import json
import sys
from pathlib import Path


def database_path():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])) / "mtr_data.json.gz"


def load_database(path=None):
    source = Path(path) if path else database_path()
    with gzip.open(source, "rt", encoding="utf-8") as handle:
        data = json.load(handle)
    for key, kind in (("registry", dict), ("rules", list), ("aliases", list),
                      ("global_unique_rules", list)):
        if not isinstance(data.get(key), kind) or not data[key]:
            raise ValueError(f"База повреждена: отсутствует раздел {key}")
    return {key: data[key] for key in
            ("version", "registry", "rules", "aliases", "global_unique_rules")}
