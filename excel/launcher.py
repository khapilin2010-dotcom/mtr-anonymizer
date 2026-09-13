"""Simple start screen for the Excel module.

The engineer normally needs only two choices: where the knowledge folder lives
and whether previously confirmed decisions may be applied automatically.
Everything else stays available as explicit maintenance buttons.
"""
import os
from pathlib import Path

from excel.simple_store import SimpleKnowledgeStore
from excel.legacy_import import preview_legacy, commit_legacy
from excel.knowledge_edit_io import export_editable, preview_editable, commit_editable


def run(app_dir, default_knowledge, open_main):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    app_dir = Path(app_dir)
    default_knowledge = Path(default_knowledge)
    app_dir.mkdir(parents=True, exist_ok=True)
    default_knowledge.mkdir(parents=True, exist_ok=True)

    root = tk.Tk()
    root.title('MTR Excel — запуск')
    root.geometry('780x560')
    root.minsize(720, 520)
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TButton', padding=8)

    current = default_knowledge
    config = app_dir / 'knowledge_config.json'
    if config.exists():
        try:
            import json
            configured = json.loads(config.read_text('utf-8')).get('folder')
            if configured and Path(configured).exists():
                current = Path(configured)
        except Exception:
            pass

    knowledge_dir = tk.StringVar(value=str(current))
    auto_apply = tk.BooleanVar(value=False)
    status = tk.StringVar(value='По умолчанию накопленные решения не удаляют спорные элементы автоматически.')

    top = ttk.Frame(root, padding=16)
    top.pack(fill='x')
    ttk.Label(top, text='MTR Excel', font=('Segoe UI', 23, 'bold')).pack(anchor='w')
    ttk.Label(top, text='Обезличивание МТР с последним решением за инженером',
              font=('Segoe UI', 11)).pack(anchor='w', pady=(2, 0))

    box = ttk.LabelFrame(root, text='База знаний', padding=12)
    box.pack(fill='x', padx=16, pady=8)
    ttk.Label(box, text='Папка базы. Если программу положить на D: или в общую папку, база может лежать рядом с ней.').pack(anchor='w')
    row = ttk.Frame(box)
    row.pack(fill='x', pady=(6, 0))
    ttk.Entry(row, textvariable=knowledge_dir).pack(side='left', fill='x', expand=True)

    def choose_folder():
        value = filedialog.askdirectory(title='Папка базы знаний', initialdir=knowledge_dir.get())
        if value:
            knowledge_dir.set(value)
    ttk.Button(row, text='Выбрать…', command=choose_folder).pack(side='left', padx=(6, 0))

    mode = ttk.LabelFrame(root, text='Ответственность за накопленные решения', padding=12)
    mode.pack(fill='x', padx=16, pady=8)
    ttk.Checkbutton(mode,
                    text='Автоматически применять ранее подтверждённые решения',
                    variable=auto_apply).pack(anchor='w')
    ttk.Label(mode,
              text='Галочка по умолчанию снята. Без неё накопленные DELETE/исправления используются как подсказка и идут на проверку. '
                   'Если инженер включает режим сам, ранее подтверждённые решения могут применяться без повторного просмотра.',
              wraplength=720).pack(anchor='w', pady=(4, 0))

    tools = ttk.LabelFrame(root, text='Импорт и управление базой', padding=12)
    tools.pack(fill='x', padx=16, pady=8)
    tools.columnconfigure(0, weight=1)
    tools.columnconfigure(1, weight=1)

    def store():
        folder = Path(knowledge_dir.get().strip())
        if not folder:
            raise ValueError('Укажите папку базы знаний.')
        folder.mkdir(parents=True, exist_ok=True)
        os.environ['MTR_KNOWLEDGE_DIR'] = str(folder)
        return SimpleKnowledgeStore(app_dir, folder)

    def import_old():
        path = filedialog.askopenfilename(title='Старая проверенная выборка',
                                          filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.xls *.csv')])
        if not path:
            return
        try:
            s = store()
            plan = preview_legacy(path, s)
            text = (f"Готово к импорту: {plan['ready']}\n"
                    f"Уже есть в базе: {plan['repeated']}\n"
                    f"Конфликтов с текущей базой: {plan['conflicts']}\n"
                    f"Проблемных строк: {len(plan['issues'])}")
            accept_conflicts = False
            if plan['conflicts']:
                accept_conflicts = messagebox.askyesno(
                    'Есть конфликты', text + '\n\nДобавить конфликтующие решения как новые голоса инженера?')
                if not accept_conflicts:
                    text += '\nКонфликтующие строки будут пропущены.'
            if not messagebox.askyesno('Импорт проверенной выборки', text + '\n\nПродолжить?'):
                return
            report = commit_legacy(path, s, accept_conflicts)
            status.set(f"Импортировано решений: {report['accepted']}. Проблем: {len(report['issues'])}.")
            if report['issues']:
                messagebox.showwarning('Импорт завершён с замечаниями',
                                       status.get() + '\nПервые проблемы:\n' + '\n'.join(
                                           f'{a} / {b}: {c}' for a, b, c in report['issues'][:12]))
            else:
                messagebox.showinfo('Готово', status.get())
        except Exception as exc:
            messagebox.showerror('Не удалось импортировать', str(exc))

    def export_base():
        path = filedialog.asksaveasfilename(title='Выгрузить базу для редактирования',
                                            defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')],
                                            initialfile='MTR_База_знаний_редактирование.xlsx')
        if not path:
            return
        try:
            export_editable(path, store())
            status.set('База выгружена. Меняйте только столбцы «Новое решение», «Действие» и «Причина изменения».')
            messagebox.showinfo('Готово', status.get())
        except Exception as exc:
            messagebox.showerror('Не удалось выгрузить', str(exc))

    def import_base():
        path = filedialog.askopenfilename(title='Загрузить отредактированную базу',
                                          filetypes=[('Excel', '*.xlsx')])
        if not path:
            return
        try:
            s = store()
            plan = preview_editable(path, s)
            text = f"Изменений к применению: {plan['ready']}\nПроблемных строк: {len(plan['issues'])}"
            if plan['issues']:
                text += '\n\nПроблемные строки не будут применены.'
            if not messagebox.askyesno('Изменение базы знаний', text + '\n\nПрименить изменения?'):
                return
            report = commit_editable(path, s)
            status.set(f"Применено изменений базы: {report['applied']}. Проблем: {len(report['issues'])}.")
            messagebox.showinfo('Готово', status.get())
        except Exception as exc:
            messagebox.showerror('Не удалось загрузить', str(exc))

    ttk.Button(tools, text='Импорт старых проверенных Excel…', command=import_old).grid(row=0, column=0, sticky='ew', padx=(0, 4), pady=4)
    ttk.Button(tools, text='Выгрузить базу для редактирования…', command=export_base).grid(row=0, column=1, sticky='ew', padx=(4, 0), pady=4)
    ttk.Button(tools, text='Загрузить отредактированную базу…', command=import_base).grid(row=1, column=0, columnspan=2, sticky='ew', pady=4)

    ttk.Label(root, textvariable=status, wraplength=735).pack(fill='x', padx=18, pady=(4, 0))

    def launch():
        try:
            s = store()
            s.sync()
            os.environ['MTR_AUTO_APPLY_CONFIRMED'] = '1' if auto_apply.get() else '0'
            selected = knowledge_dir.get().strip()
        except Exception as exc:
            messagebox.showerror('Не удалось открыть базу', str(exc))
            return
        root.destroy()
        open_main(selected, auto_apply.get())

    bottom = ttk.Frame(root, padding=16)
    bottom.pack(fill='x', side='bottom')
    ttk.Label(bottom, text='Разработал Виктор Хапилин').pack(side='left')
    ttk.Button(bottom, text='Открыть Excel-модуль', command=launch).pack(side='right')

    root.mainloop()
