"""White and light-blue palette for the engineer application."""
WHITE = '#FFFFFF'
PALE_BLUE = '#EAF4FD'
BLUE = '#1675BD'
DARK_BLUE = '#125A91'
TEXT = '#16334D'
MUTED = '#4D6C83'
BORDER = '#BDD8EE'


def apply_theme(root):
    from tkinter import ttk
    root.configure(background=WHITE)
    root.option_add('*Font', ('Segoe UI', 10))
    root.option_add('*TCombobox*Listbox.background', WHITE)
    root.option_add('*TCombobox*Listbox.foreground', TEXT)
    root.option_add('*TCombobox*Listbox.selectBackground', BLUE)
    root.option_add('*TCombobox*Listbox.selectForeground', WHITE)
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('.', background=WHITE, foreground=TEXT, bordercolor=BORDER,
                    lightcolor=WHITE, darkcolor=BORDER, font=('Segoe UI', 10))
    style.configure('TFrame', background=WHITE)
    style.configure('TLabel', background=WHITE, foreground=TEXT)
    style.configure('Header.TFrame', background=PALE_BLUE)
    style.configure('Header.TLabel', background=PALE_BLUE, foreground=DARK_BLUE)
    style.configure('TLabelframe', background=WHITE, bordercolor=BORDER)
    style.configure('TLabelframe.Label', background=WHITE, foreground=DARK_BLUE)
    style.configure('TButton', background=PALE_BLUE, foreground=DARK_BLUE,
                    bordercolor=BORDER, padding=(10, 7))
    style.map('TButton', background=[('disabled', '#F0F5F9'), ('pressed', '#BDDEF7'), ('active', '#D6EBFC')],
              foreground=[('disabled', '#8297A8')], bordercolor=[('focus', BLUE)])
    style.configure('Big.TButton', padding=(12, 11), font=('Segoe UI', 10, 'bold'))
    style.configure('Primary.TButton', background=BLUE, foreground=WHITE,
                    padding=(16, 12), font=('Segoe UI', 11, 'bold'))
    style.map('Primary.TButton', background=[('disabled', '#C6DCEB'), ('pressed', DARK_BLUE), ('active', '#2587D0')],
              foreground=[('disabled', '#627D90'), ('!disabled', WHITE)])
    style.configure('TEntry', fieldbackground=WHITE, foreground=TEXT, padding=5)
    style.configure('TCombobox', fieldbackground=WHITE, foreground=TEXT, padding=5)
    style.map('TCombobox', fieldbackground=[('readonly', WHITE)], foreground=[('readonly', TEXT)])
    style.configure('TCheckbutton', background=WHITE, foreground=TEXT)
    style.map('TCheckbutton', background=[('active', PALE_BLUE)])
    style.configure('Treeview', background=WHITE, fieldbackground=WHITE, foreground=TEXT, rowheight=28)
    style.configure('Treeview.Heading', background=PALE_BLUE, foreground=DARK_BLUE, padding=7)
    style.map('Treeview', background=[('selected', '#CEE8FB')], foreground=[('selected', TEXT)])
    style.configure('Horizontal.TProgressbar', background=BLUE, troughcolor=PALE_BLUE,
                    bordercolor=BORDER, lightcolor=BLUE, darkcolor=BLUE)
    return style


def text_colors():
    return dict(background=WHITE, foreground=TEXT, insertbackground=BLUE,
                selectbackground='#CEE8FB', selectforeground=TEXT,
                highlightbackground=BORDER, highlightcolor=BLUE, highlightthickness=1,
                borderwidth=0)
