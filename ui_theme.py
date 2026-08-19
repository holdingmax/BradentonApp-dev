"""
Premium Tkinter theme and reusable UI components for BradentonApp.
Compact dark workspace layout with dual-column tab shells.
"""

import tkinter as tk
from tkinter import ttk

FONT = "Segoe UI"

# Compact window bounds
WINDOW_GEOMETRY = "1000x650"
WINDOW_MINSIZE = (900, 580)
PAD_INNER = 4
PAD_TAB = 4


class Theme:
    """Base shell palette — deep charcoal / slate dark mode."""

    BG = "#1E1E24"
    BG_DEEP = "#14141A"
    SURFACE = "#282A36"
    SURFACE_ALT = "#3C3F41"
    BORDER = "#4B5563"
    BORDER_FOCUS = "#60A5FA"

    TEXT = "#E8EAED"
    TEXT_SOFT = "#B0B8C4"
    TEXT_ON_DARK = "#F3F4F6"
    TEXT_MUTED_ON_DARK = "#9CA3AF"

    BTN_SECONDARY = "#3C3F41"
    BTN_SECONDARY_HOVER = "#4B5563"
    BTN_SECONDARY_TEXT = "#E8EAED"

    SUCCESS = "#34D399"
    ERROR = "#F87171"
    WARNING = "#FBBF24"

    SHADOW = "#14141A"
    TAB_IDLE = "#1E1E24"
    TAB_ACTIVE_BG = "#282A36"
    TAB_HOVER = "#252830"


class SectionTheme:
    """Per-module accent palette."""

    def __init__(
        self,
        accent,
        accent_hover,
        accent_soft,
        card_tint,
        border_focus,
        tab_selected_fg=None,
    ):
        self.accent = accent
        self.accent_hover = accent_hover
        self.accent_soft = accent_soft
        self.card_tint = card_tint
        self.border_focus = border_focus
        self.tab_selected_fg = tab_selected_fg or accent


THEME = Theme()

EFT_THEME = SectionTheme(
    accent="#3B82F6",
    accent_hover="#2563EB",
    accent_soft="#1E3A5F",
    card_tint="#252830",
    border_focus="#60A5FA",
    tab_selected_fg="#60A5FA",
)

CHASE_THEME = SectionTheme(
    accent="#10B981",
    accent_hover="#059669",
    accent_soft="#064E3B",
    card_tint="#252830",
    border_focus="#34D399",
    tab_selected_fg="#34D399",
)

CMV_THEME = SectionTheme(
    accent="#8B5CF6",
    accent_hover="#7C3AED",
    accent_soft="#4C1D95",
    card_tint="#2A2540",
    border_focus="#A78BFA",
    tab_selected_fg="#A78BFA",
)

SALES_THEME = SectionTheme(
    accent="#F97316",
    accent_hover="#EA580C",
    accent_soft="#7C2D12",
    card_tint="#2D261F",
    border_focus="#FB923C",
    tab_selected_fg="#FB923C",
)

REPORTE_DIARIO_THEME = SectionTheme(
    accent="#06B6D4",
    accent_hover="#0891B2",
    accent_soft="#164E63",
    card_tint="#1F2A32",
    border_focus="#22D3EE",
    tab_selected_fg="#22D3EE",
)


def apply_root_style(root):
    root.configure(bg=THEME.BG)


def apply_notebook_style(style, sections=None):
    sections = sections or [
        EFT_THEME,
        CHASE_THEME,
        REPORTE_DIARIO_THEME,
        CMV_THEME,
        SALES_THEME,
    ]
    style.theme_use("clam")
    style.configure(
        "Premium.TNotebook",
        background=THEME.BG,
        borderwidth=0,
        tabmargins=[4, 4, 4, 0],
    )
    style.configure(
        "Premium.TNotebook.Tab",
        font=(FONT, 9, "bold"),
        padding=[12, 7],
        background=THEME.TAB_IDLE,
        foreground=THEME.TEXT_SOFT,
        borderwidth=0,
        focuscolor=THEME.BG,
    )
    style.map(
        "Premium.TNotebook.Tab",
        background=[
            ("selected", THEME.SURFACE),
            ("active", THEME.TAB_HOVER),
        ],
        foreground=[
            ("selected", THEME.TEXT),
            ("active", THEME.TEXT),
        ],
        expand=[("selected", [1, 1, 1, 0])],
    )


def create_header_banner(parent):
    banner = tk.Frame(parent, bg=THEME.BG_DEEP, padx=12, pady=10)
    banner.pack(fill=tk.X)

    top = tk.Frame(banner, bg=THEME.BG_DEEP)
    top.pack(fill=tk.X)

    tk.Label(
        top,
        text="Financial Automation Suite",
        font=(FONT, 16, "bold"),
        fg=THEME.TEXT_ON_DARK,
        bg=THEME.BG_DEEP,
        anchor=tk.W,
    ).pack(side=tk.LEFT, anchor=tk.W)

    badges = tk.Frame(top, bg=THEME.BG_DEEP)
    badges.pack(side=tk.RIGHT)
    for label, color in (
        ("EFT", EFT_THEME.accent),
        ("Chase", CHASE_THEME.accent),
        ("Diario", REPORTE_DIARIO_THEME.accent),
        ("CMV", CMV_THEME.accent),
        ("Ventas", SALES_THEME.accent),
    ):
        tk.Label(
            badges,
            text=label,
            font=(FONT, 7, "bold"),
            fg="#FFFFFF",
            bg=color,
            padx=6,
            pady=2,
        ).pack(side=tk.LEFT, padx=(4, 0))

    gradient = tk.Frame(banner, bg=THEME.BG_DEEP)
    gradient.pack(fill=tk.X, pady=(8, 0))
    for color in (
        EFT_THEME.accent,
        CHASE_THEME.accent,
        REPORTE_DIARIO_THEME.accent,
        CMV_THEME.accent,
        SALES_THEME.accent,
    ):
        tk.Frame(gradient, bg=color, height=2).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
    return banner


def create_scrollable_body(parent):
    body = tk.Frame(parent, bg=THEME.BG, padx=PAD_INNER, pady=PAD_INNER)
    body.pack(fill=tk.BOTH, expand=True)
    return body


def create_dual_column_tab(parent):
    """
    Two-column tab shell: header row spans both columns; row 1 = ops | rules.
    Returns (header_frame, left_ops_column, right_rules_column).
    """
    parent.grid_columnconfigure(0, weight=1, uniform="tabcol")
    parent.grid_columnconfigure(1, weight=1, uniform="tabcol")
    parent.grid_rowconfigure(1, weight=1)

    header = tk.Frame(parent, bg=THEME.BG)
    header.grid(
        row=0,
        column=0,
        columnspan=2,
        sticky="ew",
        padx=PAD_TAB,
        pady=(PAD_TAB, 2),
    )

    left = tk.Frame(parent, bg=THEME.BG)
    left.grid(row=1, column=0, sticky="nsew", padx=(PAD_TAB, 2), pady=PAD_TAB)
    left.grid_columnconfigure(0, weight=1)

    right = tk.Frame(parent, bg=THEME.BG)
    right.grid(row=1, column=1, sticky="nsew", padx=(2, PAD_TAB), pady=PAD_TAB)
    right.grid_columnconfigure(0, weight=1)
    right.grid_rowconfigure(0, weight=1)

    return header, left, right


def create_compact_section_header(parent, title, description, section_theme):
    wrap = tk.Frame(parent, bg=THEME.BG)
    wrap.pack(fill=tk.X)

    title_row = tk.Frame(wrap, bg=THEME.BG)
    title_row.pack(fill=tk.X, anchor=tk.W)

    stripe = tk.Frame(title_row, bg=section_theme.accent, width=4, height=18)
    stripe.pack(side=tk.LEFT, padx=(0, 6))

    tk.Label(
        title_row,
        text=title,
        font=(FONT, 11, "bold"),
        fg=THEME.TEXT,
        bg=THEME.BG,
        anchor=tk.W,
    ).pack(side=tk.LEFT, anchor=tk.W)

    tk.Label(
        wrap,
        text=description,
        font=(FONT, 8),
        fg=THEME.TEXT_SOFT,
        bg=THEME.BG,
        anchor=tk.W,
        justify=tk.LEFT,
        wraplength=900,
    ).pack(anchor=tk.W, pady=(2, 0), padx=(10, 0))

    return wrap


def create_panel_label(parent, text, section_theme=None):
    section_theme = section_theme or EFT_THEME
    tk.Label(
        parent,
        text=text,
        font=(FONT, 9, "bold"),
        fg=THEME.TEXT,
        bg=section_theme.card_tint,
        anchor=tk.W,
    ).pack(fill=tk.X, pady=(0, 4))


def create_card(parent, section_theme=None, padx=8, pady=8, fill=tk.X, expand=False):
    section_theme = section_theme or EFT_THEME
    outer = tk.Frame(parent, bg=THEME.SHADOW, padx=0, pady=0)
    outer.pack(fill=fill, expand=expand, pady=(0, 4))

    inner = tk.Frame(
        outer,
        bg=section_theme.card_tint,
        highlightbackground=section_theme.accent_soft,
        highlightthickness=1,
        padx=padx,
        pady=pady,
    )
    inner.pack(fill=fill)
    return inner


def create_primary_button(parent, text, command, section_theme=None):
    section_theme = section_theme or EFT_THEME
    btn = tk.Button(
        parent,
        text=text,
        font=(FONT, 9, "bold"),
        fg="#FFFFFF",
        bg=section_theme.accent,
        activebackground=section_theme.accent_hover,
        activeforeground="#FFFFFF",
        relief=tk.FLAT,
        overrelief=tk.FLAT,
        cursor="hand2",
        padx=16,
        pady=9,
        borderwidth=0,
        highlightthickness=0,
        command=command,
    )
    bind_hover(btn, section_theme.accent, section_theme.accent_hover)
    return btn


def create_secondary_button(
    parent, text, command, section_theme=None, width=None
):
    section_theme = section_theme or EFT_THEME
    btn = tk.Button(
        parent,
        text=text,
        font=(FONT, 8, "bold"),
        fg=THEME.TEXT,
        bg=THEME.BTN_SECONDARY,
        activebackground=section_theme.accent,
        activeforeground="#FFFFFF",
        relief=tk.FLAT,
        overrelief=tk.FLAT,
        cursor="hand2",
        padx=10,
        pady=6,
        borderwidth=0,
        highlightthickness=0,
        command=command,
    )
    if width:
        btn.configure(width=width)
    bind_hover_soft(btn, section_theme)
    return btn


def bind_hover_soft(button, section_theme):
    normal = THEME.BTN_SECONDARY
    hover = section_theme.accent
    fg_normal = THEME.TEXT
    fg_hover = "#FFFFFF"

    def on_enter(_event):
        button.configure(bg=hover, fg=fg_hover)

    def on_leave(_event):
        button.configure(bg=normal, fg=fg_normal)

    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)


def create_file_row(
    parent,
    label,
    textvariable,
    browse_command,
    section_theme=None,
    label_width=11,
    browse_label="Browse",
):
    section_theme = section_theme or EFT_THEME
    row = tk.Frame(parent, bg=section_theme.card_tint)
    row.pack(fill=tk.X, pady=(0, 6))

    tk.Label(
        row,
        text=label,
        font=(FONT, 8, "bold"),
        fg=THEME.TEXT,
        bg=section_theme.card_tint,
        width=label_width,
        anchor=tk.W,
    ).pack(side=tk.LEFT)

    field_wrap = tk.Frame(
        row,
        bg=THEME.SURFACE_ALT,
        highlightbackground=section_theme.accent_soft,
        highlightthickness=1,
    )
    field_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 6))

    entry = tk.Entry(
        field_wrap,
        textvariable=textvariable,
        font=(FONT, 9),
        fg=THEME.TEXT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground=THEME.SURFACE_ALT,
        highlightcolor=section_theme.border_focus,
        insertbackground=section_theme.accent,
        readonlybackground=THEME.SURFACE_ALT,
    )
    entry.pack(fill=tk.X, expand=True, padx=6, pady=4)

    create_secondary_button(
        row, browse_label, browse_command, section_theme=section_theme
    ).pack(side=tk.RIGHT)
    return entry


def create_compact_entry(parent, textvariable, section_theme=None, label=None):
    section_theme = section_theme or EFT_THEME
    wrap = tk.Frame(parent, bg=section_theme.card_tint)
    wrap.pack(fill=tk.X, pady=(0, 4))
    if label:
        tk.Label(
            wrap,
            text=label,
            font=(FONT, 8, "bold"),
            fg=THEME.TEXT_SOFT,
            bg=section_theme.card_tint,
            anchor=tk.W,
        ).pack(anchor=tk.W)
    entry = tk.Entry(
        wrap,
        textvariable=textvariable,
        font=(FONT, 9),
        fg=THEME.TEXT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        insertbackground=section_theme.accent,
    )
    entry.pack(fill=tk.X, pady=(2, 0), ipady=4)
    return entry


def create_scrolled_listbox(parent, section_theme=None, height=8):
    """Bounded vertical list with scrollbar for rule managers."""
    section_theme = section_theme or CHASE_THEME
    container = tk.Frame(
        parent,
        bg=THEME.SURFACE_ALT,
        highlightbackground=section_theme.accent_soft,
        highlightthickness=1,
    )
    container.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)

    scrollbar = tk.Scrollbar(container, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

    listbox = tk.Listbox(
        container,
        font=(FONT, 9),
        fg=THEME.TEXT,
        bg=THEME.SURFACE_ALT,
        selectbackground=section_theme.accent,
        selectforeground="#FFFFFF",
        activestyle="none",
        relief=tk.FLAT,
        highlightthickness=0,
        yscrollcommand=scrollbar.set,
        height=height,
    )
    listbox.grid(row=0, column=0, sticky="nsew")
    scrollbar.config(command=listbox.yview)
    return listbox, container


def create_status_bar(parent, textvariable, section_theme=None, wraplength=420):
    section_theme = section_theme or EFT_THEME
    bar = tk.Frame(
        parent,
        bg=THEME.SURFACE,
        highlightbackground=section_theme.accent_soft,
        highlightthickness=1,
        padx=8,
        pady=6,
    )
    bar.pack(fill=tk.X, pady=(4, 0))

    dot = tk.Label(
        bar, text="●", font=(FONT, 8), fg=section_theme.accent, bg=THEME.SURFACE
    )
    dot.pack(side=tk.LEFT)

    label = tk.Label(
        bar,
        textvariable=textvariable,
        font=(FONT, 8),
        fg=THEME.TEXT_SOFT,
        bg=THEME.SURFACE,
        anchor=tk.W,
        wraplength=wraplength,
        justify=tk.LEFT,
    )
    label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
    return label, dot


def create_info_panel(parent, title, lines, section_theme=None):
    """Compact read-only notes block for the right column."""
    section_theme = section_theme or EFT_THEME
    card = create_card(
        parent,
        section_theme=section_theme,
        padx=PAD_TAB,
        pady=PAD_TAB,
        fill=tk.BOTH,
        expand=True,
    )
    create_panel_label(card, title, section_theme)
    for line in lines:
        tk.Label(
            card,
            text=line,
            font=(FONT, 8),
            fg=THEME.TEXT_SOFT,
            bg=section_theme.card_tint,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=380,
        ).pack(anchor=tk.W, pady=(0, 2))
    return card


def create_log_panel(parent, title, section_theme=None, height=10):
    """Scrollable summary log for the right column."""
    section_theme = section_theme or EFT_THEME
    card = create_card(
        parent,
        section_theme=section_theme,
        padx=PAD_TAB,
        pady=PAD_TAB,
        fill=tk.BOTH,
        expand=True,
    )
    create_panel_label(card, title, section_theme)

    container = tk.Frame(
        card,
        bg=THEME.SURFACE_ALT,
        highlightbackground=section_theme.accent_soft,
        highlightthickness=1,
    )
    container.pack(fill=tk.BOTH, expand=True)
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)

    scrollbar = tk.Scrollbar(container, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

    log = tk.Text(
        container,
        font=(FONT, 8),
        fg=THEME.TEXT_SOFT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        highlightthickness=0,
        wrap=tk.WORD,
        height=height,
        padx=6,
        pady=6,
        yscrollcommand=scrollbar.set,
        state=tk.DISABLED,
    )
    log.grid(row=0, column=0, sticky="nsew")
    scrollbar.config(command=log.yview)
    return log, card


def set_status_style(label, dot, message, section_theme=None, is_error=False, completed=False):
    section_theme = section_theme or EFT_THEME
    if is_error:
        color = THEME.ERROR
    elif completed or "completed" in message.lower() or "updated" in message.lower():
        color = THEME.SUCCESS
    else:
        color = section_theme.accent
    label.configure(fg=THEME.TEXT if not is_error and not completed else color)
    dot.configure(fg=color)


def bind_hover(button, normal_color, hover_color):
    button.bind(
        "<Enter>",
        lambda _e, b=button, h=hover_color: b.configure(bg=h),
    )
    button.bind(
        "<Leave>",
        lambda _e, b=button, n=normal_color: b.configure(bg=n),
    )


EXCEL_FILETYPES = [
    ("Excel Files", "*.xlsx"),
    ("Excel Files", "*.xls"),
    ("All Files", "*.*"),
]

EXCEL_FILETYPES_MASTER = [
    ("Excel Files", "*.xlsx"),
    ("Excel Files", "*.xls"),
    ("All Files", "*.*"),
]

DEPT_FILETYPES = [
    ("Department CSV", "*.csv"),
    ("Department Excel", "*.xlsx"),
    ("Department Excel", "*.xls"),
    ("All Files", "*.*"),
]

SALES_FILETYPES = [
    ("Sales Reports", "*.csv"),
    ("Sales Reports", "*.xlsx"),
    ("Sales Reports", "*.xls"),
    ("Sales Reports", "*.xlsm"),
    ("All Files", "*.*"),
]

SALES_CSV_FILETYPES = SALES_FILETYPES

PDF_DAILY_FILETYPES = [
    ("Elistar Daily PDF", "*.pdf"),
    ("All Files", "*.*"),
]

ANALISIS_MASTER_FILETYPES = [
    ("Bradenton Analisis C-Store", "*.xlsx"),
    ("Bradenton Analisis C-Store", "*.xlsm"),
    ("All Files", "*.*"),
]
