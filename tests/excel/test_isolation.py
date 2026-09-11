import ast
from pathlib import Path
import subprocess
import sys


def test_only_independent_modules_imported():
    script = '''
import sys
from excel.excel_engine import Anonymizer
from excel.file_io import process_file
az=Anonymizer()
assert len(az.registry)>100000
assert not any(name in sys.modules for name in ('mtr_core','MTR_Obezlichivatel','fitz','pymupdf','pytesseract'))
'''
    subprocess.run([sys.executable, '-c', script], check=True)


def test_no_imports_of_legacy_or_document_rendering():
    banned={'mtr_core','MTR_Obezlichivatel','fitz','pymupdf','pytesseract'}
    for folder in ('excel','common'):
        for path in Path(folder).glob('*.py'):
            for node in ast.walk(ast.parse(path.read_text(encoding='utf-8'))):
                if isinstance(node,ast.Import):
                    assert not any(alias.name.split('.')[0] in banned for alias in node.names),path
                if isinstance(node,ast.ImportFrom):
                    assert (node.module or '').split('.')[0] not in banned,path
