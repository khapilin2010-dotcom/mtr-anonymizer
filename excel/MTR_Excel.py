"""MTR Excel desktop application and packaged self-test entry point."""
import argparse
import json
import os
from pathlib import Path
import sys

if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

APP_VERSION = '1.6 RC1'
APP_DIR = Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'MTR_Excel'
PROGRAM_DIR = (Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False)
               else Path(__file__).resolve().parents[1])
DEFAULT_KNOWLEDGE = PROGRAM_DIR / 'MTR_Knowledge'


def _set_runtime_version():
    import excel.knowledge as knowledge
    knowledge.VERSION = APP_VERSION


def self_test(marker):
    from excel.self_test import run
    report = run()
    Path(marker).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def _operator_store(shared=None):
    _set_runtime_version()
    from excel.simple_store import SimpleKnowledgeStore
    folder = Path(shared) if shared else DEFAULT_KNOWLEDGE
    folder.mkdir(parents=True, exist_ok=True)
    os.environ['MTR_KNOWLEDGE_DIR'] = str(folder)
    return SimpleKnowledgeStore(APP_DIR, folder)


def _install_operator_policy():
    """Patch only the classes imported by the existing mature UI.

    The old UI/review workflow stays intact.  We swap in the simplified store
    and conservative learned-decision policy before importing that UI module.
    """
    import excel.knowledge as knowledge
    import excel.expert_engine as expert_engine
    from excel.simple_store import SimpleKnowledgeStore
    from excel.operator_engine import OperatorExpertAnonymizer
    knowledge.VERSION = APP_VERSION
    knowledge.KnowledgeStore = SimpleKnowledgeStore
    expert_engine.ExpertAnonymizer = OperatorExpertAnonymizer


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--self-test', action='store_true')
    parser.add_argument('--marker', default='SELF_TEST_OK.json')
    parser.add_argument('--output')
    parser.add_argument('--shared', help='Папка базы знаний; по умолчанию MTR_Knowledge рядом с программой')
    parser.add_argument('--auto-confirmed', action='store_true',
                        help='Автоматически применять ранее подтверждённые решения')
    parser.add_argument('--import-review', metavar='FILE', help='Вернуть проверенный Excel этой программы')
    parser.add_argument('--import-history', metavar='FILE', help='Импорт старой проверенной/обезличенной выборки')
    parser.add_argument('--export-knowledge', metavar='FILE', help='Выгрузить базу знаний для редактирования')
    parser.add_argument('--import-knowledge', metavar='FILE', help='Загрузить отредактированную выгрузку базы знаний')
    parser.add_argument('--manager-report', metavar='FILE', help='Сформировать отчёт руководителю XLSX')
    parser.add_argument('files', nargs='*')
    args = parser.parse_args()

    if args.self_test:
        self_test(args.marker)
        return

    _set_runtime_version()
    os.environ['MTR_AUTO_APPLY_CONFIRMED'] = '1' if args.auto_confirmed else '0'
    store = _operator_store(args.shared)

    if args.import_review:
        from excel.review_io import import_review
        print(json.dumps(import_review(args.import_review, store), ensure_ascii=False, default=str))
        return
    if args.import_history:
        from excel.legacy_import import commit_legacy
        print(json.dumps(commit_legacy(args.import_history, store, accept_conflicts=False),
                         ensure_ascii=False, default=str))
        return
    if args.export_knowledge:
        from excel.knowledge_edit_io import export_editable
        export_editable(args.export_knowledge, store)
        print(args.export_knowledge)
        return
    if args.import_knowledge:
        from excel.knowledge_edit_io import commit_editable
        print(json.dumps(commit_editable(args.import_knowledge, store), ensure_ascii=False, default=str))
        return
    if args.manager_report:
        from excel.quality_tools import export_manager_report
        export_manager_report(args.manager_report, store)
        print(args.manager_report)
        return

    if args.files:
        if not args.output:
            parser.error('Для командной строки требуется --output ПАПКА')
        from excel.operator_engine import OperatorExpertAnonymizer
        from excel.file_io import process_file
        store.sync()
        az = OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=args.auto_confirmed)
        errors = []
        for file in args.files:
            try:
                process_file(file, args.output, az, knowledge_store=store)
            except Exception as exc:
                errors.append(f'{file}: {exc}')
        store.sync()
        if errors:
            raise RuntimeError('\n'.join(errors))
        return
    gui()


def gui():
    _set_runtime_version()
    from excel.launcher import run as launcher

    def open_main(selected_folder, auto_apply):
        os.environ['MTR_KNOWLEDGE_DIR'] = str(selected_folder)
        os.environ['MTR_AUTO_APPLY_CONFIRMED'] = '1' if auto_apply else '0'
        _install_operator_policy()
        from excel.expert_ui import gui as expert_gui
        expert_gui(APP_DIR)

    launcher(APP_DIR, DEFAULT_KNOWLEDGE, open_main)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        import traceback
        (APP_DIR / 'startup_error.log').write_text(traceback.format_exc(), encoding='utf-8')
        if '--self-test' not in sys.argv and not any(x in sys.argv for x in ('--output',)):
            try:
                from tkinter import messagebox
                messagebox.showerror('Ошибка запуска MTR Excel',
                                     f'Не удалось запустить программу. Журнал: {APP_DIR / "startup_error.log"}')
            except Exception:
                pass
        raise
