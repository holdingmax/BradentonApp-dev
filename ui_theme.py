"""
Tkinter theme and reusable UI components for BradentonApp.
Clean, professional light workspace with a distinct color per module.
"""

import tkinter as tk
from tkinter import ttk

FONT = "Segoe UI"

WINDOW_GEOMETRY = "1180x760"
WINDOW_MINSIZE = (1040, 680)
PAD_INNER = 14
PAD_TAB = 12


class Theme:
    """Base shell palette — light, professional workspace."""

    BG = "#F4F6F9"
    BG_DEEP = "#16233F"
    SURFACE = "#FFFFFF"
    SURFACE_ALT = "#F6F7FA"
    BORDER = "#E1E5EB"
    BORDER_FOCUS = "#3B5BDB"

    TEXT = "#1F2A3D"
    TEXT_SOFT = "#5B6472"
    TEXT_ON_DARK = "#FFFFFF"
    TEXT_MUTED_ON_DARK = "#8A93A3"

    BTN_SECONDARY = "#EDF0F5"
    BTN_SECONDARY_HOVER = "#DEE3EB"
    BTN_SECONDARY_TEXT = "#1F2A3D"

    SUCCESS = "#15803D"
    ERROR = "#DC2626"
    WARNING = "#B45309"

    SHADOW = "#D8DCE3"
    TAB_IDLE = "#EAEDF2"
    TAB_ACTIVE_BG = "#FFFFFF"
    TAB_HOVER = "#E1E6ED"

    HEADER_SUBTITLE = "#AAB6CC"


class SectionTheme:
    """Per-module accent palette — one distinct color identity per tab."""

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

# Cupones y EFT — indigo
EFT_THEME = SectionTheme(
    accent="#3B5BDB",
    accent_hover="#2C46B5",
    accent_soft="#DDE3FA",
    card_tint="#FFFFFF",
    border_focus="#3B5BDB",
)

# Gettel / Toyota — teal (its own identity, no longer borrowing EFT's blue)
GETTEL_THEME = SectionTheme(
    accent="#0D9488",
    accent_hover="#0B7A70",
    accent_soft="#D6F1EE",
    card_tint="#FFFFFF",
    border_focus="#0D9488",
)

# Chase Bank — green
CHASE_THEME = SectionTheme(
    accent="#16A34A",
    accent_hover="#128038",
    accent_soft="#DCF3E3",
    card_tint="#FFFFFF",
    border_focus="#16A34A",
)

# Reporte Diario — cyan
REPORTE_DIARIO_THEME = SectionTheme(
    accent="#0284C7",
    accent_hover="#0369A1",
    accent_soft="#D7EFFB",
    card_tint="#FFFFFF",
    border_focus="#0284C7",
)

# CMV Costo — purple
CMV_THEME = SectionTheme(
    accent="#7C3AED",
    accent_hover="#6425D1",
    accent_soft="#E9E0FC",
    card_tint="#FFFFFF",
    border_focus="#7C3AED",
)

# CMV Ventas — orange
SALES_THEME = SectionTheme(
    accent="#EA580C",
    accent_hover="#C2440A",
    accent_soft="#FCE4D6",
    card_tint="#FFFFFF",
    border_focus="#EA580C",
)

# Proveedores — pink
PROVEEDORES_THEME = SectionTheme(
    accent="#DB2777",
    accent_hover="#B91C63",
    accent_soft="#FBD9EA",
    card_tint="#FFFFFF",
    border_focus="#DB2777",
)


def apply_root_style(root):
    root.configure(bg=THEME.BG)


def apply_notebook_style(style, sections=None):
    """
    ttk.Notebook can't render a different hue per tab in its shared style —
    each module's color identity instead lives in that tab's own content
    (section badge, cards, buttons; see create_tab_icon for a small color
    swatch next to the tab label). This just gives the strip itself one
    clean, consistent look.
    """
    style.theme_use("clam")
    style.configure(
        "Premium.TNotebook",
        background=THEME.BG,
        borderwidth=0,
        tabmargins=[6, 8, 6, 0],
    )
    style.configure(
        "Premium.TNotebook.Tab",
        font=(FONT, 10, "bold"),
        padding=[14, 10],
        background=THEME.TAB_IDLE,
        foreground=THEME.TEXT_SOFT,
        borderwidth=0,
        focuscolor=THEME.BG,
    )
    style.map(
        "Premium.TNotebook.Tab",
        background=[
            ("selected", THEME.TAB_ACTIVE_BG),
            ("active", THEME.TAB_HOVER),
        ],
        foreground=[
            ("selected", THEME.TEXT),
            ("active", THEME.TEXT),
        ],
        expand=[("selected", [1, 1, 1, 0])],
    )


def make_tab_icon(color, size=10):
    """
    A tiny solid-color square PhotoImage used as each notebook tab's leading
    icon — the one place a per-tab hue can actually render in a ttk
    Notebook strip. Caller must keep a reference (Tkinter drops images with
    no live reference), e.g. self._tab_icons.append(...).
    """
    image = tk.PhotoImage(width=size, height=size)
    image.put(color, to=(0, 0, size, size))
    return image


def create_header_banner(parent, subtitle=None):
    banner = tk.Frame(parent, bg=THEME.BG_DEEP, padx=20, pady=14)
    banner.pack(fill=tk.X)

    top = tk.Frame(banner, bg=THEME.BG_DEEP)
    top.pack(fill=tk.X)

    left = tk.Frame(top, bg=THEME.BG_DEEP)
    left.pack(side=tk.LEFT, anchor=tk.W)

    badge = tk.Frame(left, bg=EFT_THEME.accent, width=34, height=34)
    badge.pack(side=tk.LEFT, padx=(0, 10))
    badge.pack_propagate(False)
    tk.Label(
        badge,
        text="B",
        font=(FONT, 14, "bold"),
        fg="#FFFFFF",
        bg=EFT_THEME.accent,
    ).pack(expand=True)

    text_col = tk.Frame(left, bg=THEME.BG_DEEP)
    text_col.pack(side=tk.LEFT)
    tk.Label(
        text_col,
        text="Bradenton App",
        font=(FONT, 15, "bold"),
        fg=THEME.TEXT_ON_DARK,
        bg=THEME.BG_DEEP,
        anchor=tk.W,
    ).pack(anchor=tk.W)
    tk.Label(
        text_col,
        text=subtitle or "Suite de automatización contable",
        font=(FONT, 9),
        fg=THEME.HEADER_SUBTITLE,
        bg=THEME.BG_DEEP,
        anchor=tk.W,
    ).pack(anchor=tk.W)

    return banner


def create_scrollable_body(parent):
    body = tk.Frame(parent, bg=THEME.BG, padx=PAD_INNER, pady=PAD_INNER)
    body.pack(fill=tk.BOTH, expand=True)
    return body


def create_scrollable_tab_frame(notebook):
    """
    Wrap one notebook tab's content in a vertically scrollable canvas.

    Only the tabs whose content can exceed the window's minimum height need
    this — most tabs fit comfortably as a plain frame. Returns (outer,
    inner): add `outer` to the notebook, build the tab's real content into
    `inner` exactly as if it were a normal tab frame.
    """
    outer = tk.Frame(notebook, bg=THEME.BG)
    canvas = tk.Canvas(outer, bg=THEME.BG, highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    inner = tk.Frame(canvas, bg=THEME.BG)
    inner_window = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _sync_scrollregion(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _sync_inner_width(event):
        canvas.itemconfigure(inner_window, width=event.width)

    inner.bind("<Configure>", _sync_scrollregion)
    canvas.bind("<Configure>", _sync_inner_width)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_mousewheel(_event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind_mousewheel(_event):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _bind_mousewheel)
    canvas.bind("<Leave>", _unbind_mousewheel)

    return outer, inner


def create_dual_column_tab(parent):
    """
    Two-column tab shell: header row spans both columns; row 1 = ops | rules.
    Returns (header_frame, left_ops_column, right_rules_column).
    """
    parent.grid_columnconfigure(0, weight=3, uniform="tabcol")
    parent.grid_columnconfigure(1, weight=2, uniform="tabcol")
    parent.grid_rowconfigure(1, weight=1)

    header = tk.Frame(parent, bg=THEME.BG)
    header.grid(
        row=0,
        column=0,
        columnspan=2,
        sticky="ew",
        padx=PAD_TAB,
        pady=(PAD_TAB, 8),
    )

    left = tk.Frame(parent, bg=THEME.BG)
    left.grid(row=1, column=0, sticky="nsew", padx=(PAD_TAB, 6), pady=(0, PAD_TAB))
    left.grid_columnconfigure(0, weight=1)

    right = tk.Frame(parent, bg=THEME.BG)
    right.grid(row=1, column=1, sticky="nsew", padx=(6, PAD_TAB), pady=(0, PAD_TAB))
    right.grid_columnconfigure(0, weight=1)
    right.grid_rowconfigure(0, weight=1)

    return header, left, right


def create_compact_section_header(parent, title, description, section_theme):
    wrap = tk.Frame(parent, bg=THEME.BG)
    wrap.pack(fill=tk.X)

    title_row = tk.Frame(wrap, bg=THEME.BG)
    title_row.pack(fill=tk.X, anchor=tk.W)

    badge = tk.Frame(title_row, bg=section_theme.accent, width=28, height=28)
    badge.pack(side=tk.LEFT, padx=(0, 10))
    badge.pack_propagate(False)
    tk.Label(
        badge,
        text="●",
        font=(FONT, 10),
        fg="#FFFFFF",
        bg=section_theme.accent,
    ).pack(expand=True)

    tk.Label(
        title_row,
        text=title,
        font=(FONT, 15, "bold"),
        fg=THEME.TEXT,
        bg=THEME.BG,
        anchor=tk.W,
    ).pack(side=tk.LEFT, anchor=tk.W)

    tk.Label(
        wrap,
        text=description,
        font=(FONT, 9),
        fg=THEME.TEXT_SOFT,
        bg=THEME.BG,
        anchor=tk.W,
        justify=tk.LEFT,
        wraplength=1000,
    ).pack(anchor=tk.W, pady=(4, 0), padx=(38, 0))

    return wrap


def create_panel_label(parent, text, section_theme=None):
    section_theme = section_theme or EFT_THEME
    tk.Label(
        parent,
        text=text,
        font=(FONT, 10, "bold"),
        fg=THEME.TEXT,
        bg=section_theme.card_tint,
        anchor=tk.W,
    ).pack(fill=tk.X, pady=(0, 8))


def create_card(parent, section_theme=None, padx=14, pady=14, fill=tk.X, expand=False):
    section_theme = section_theme or EFT_THEME
    outer = tk.Frame(parent, bg=THEME.SHADOW, padx=0, pady=0)
    outer.pack(fill=fill, expand=expand, pady=(0, 10))

    inner = tk.Frame(
        outer,
        bg=section_theme.card_tint,
        highlightbackground=THEME.BORDER,
        highlightthickness=1,
        padx=padx,
        pady=pady,
    )
    inner.pack(fill=fill, padx=(0, 0), pady=(0, 1))
    return inner


def create_primary_button(parent, text, command, section_theme=None):
    section_theme = section_theme or EFT_THEME
    btn = tk.Button(
        parent,
        text=text,
        font=(FONT, 10, "bold"),
        fg="#FFFFFF",
        bg=section_theme.accent,
        activebackground=section_theme.accent_hover,
        activeforeground="#FFFFFF",
        relief=tk.FLAT,
        overrelief=tk.FLAT,
        cursor="hand2",
        padx=18,
        pady=10,
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
        font=(FONT, 9, "bold"),
        fg=THEME.BTN_SECONDARY_TEXT,
        bg=THEME.BTN_SECONDARY,
        activebackground=section_theme.accent,
        activeforeground="#FFFFFF",
        relief=tk.FLAT,
        overrelief=tk.FLAT,
        cursor="hand2",
        padx=12,
        pady=7,
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
    fg_normal = THEME.BTN_SECONDARY_TEXT
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
    browse_label="Examinar",
):
    section_theme = section_theme or EFT_THEME
    row = tk.Frame(parent, bg=section_theme.card_tint)
    row.pack(fill=tk.X, pady=(0, 8))

    tk.Label(
        row,
        text=label,
        font=(FONT, 9, "bold"),
        fg=THEME.TEXT_SOFT,
        bg=section_theme.card_tint,
        width=label_width,
        anchor=tk.W,
    ).pack(side=tk.LEFT)

    field_wrap = tk.Frame(
        row,
        bg=THEME.SURFACE_ALT,
        highlightbackground=THEME.BORDER,
        highlightthickness=1,
    )
    field_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))

    entry = tk.Entry(
        field_wrap,
        textvariable=textvariable,
        font=(FONT, 10),
        fg=THEME.TEXT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground=THEME.SURFACE_ALT,
        highlightcolor=section_theme.border_focus,
        insertbackground=section_theme.accent,
        readonlybackground=THEME.SURFACE_ALT,
    )
    entry.pack(fill=tk.X, expand=True, padx=8, pady=6)

    create_secondary_button(
        row, browse_label, browse_command, section_theme=section_theme
    ).pack(side=tk.RIGHT)
    return entry


def create_compact_entry(parent, textvariable, section_theme=None, label=None):
    section_theme = section_theme or EFT_THEME
    wrap = tk.Frame(parent, bg=section_theme.card_tint)
    wrap.pack(fill=tk.X, pady=(0, 6))
    if label:
        tk.Label(
            wrap,
            text=label,
            font=(FONT, 9, "bold"),
            fg=THEME.TEXT_SOFT,
            bg=section_theme.card_tint,
            anchor=tk.W,
        ).pack(anchor=tk.W)
    entry = tk.Entry(
        wrap,
        textvariable=textvariable,
        font=(FONT, 10),
        fg=THEME.TEXT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground=THEME.BORDER,
        highlightcolor=section_theme.border_focus,
        insertbackground=section_theme.accent,
    )
    entry.pack(fill=tk.X, pady=(3, 0), ipady=5)
    return entry


def create_status_bar(parent, textvariable, section_theme=None, wraplength=460):
    section_theme = section_theme or EFT_THEME
    bar = tk.Frame(
        parent,
        bg=THEME.SURFACE_ALT,
        highlightbackground=THEME.BORDER,
        highlightthickness=1,
        padx=10,
        pady=8,
    )
    bar.pack(fill=tk.X, pady=(8, 0))

    dot = tk.Label(
        bar, text="●", font=(FONT, 9), fg=section_theme.accent, bg=THEME.SURFACE_ALT
    )
    dot.pack(side=tk.LEFT)

    label = tk.Label(
        bar,
        textvariable=textvariable,
        font=(FONT, 9),
        fg=THEME.TEXT_SOFT,
        bg=THEME.SURFACE_ALT,
        anchor=tk.W,
        wraplength=wraplength,
        justify=tk.LEFT,
    )
    label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
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
        row = tk.Frame(card, bg=section_theme.card_tint)
        row.pack(fill=tk.X, pady=(0, 6), anchor=tk.W)
        tk.Label(
            row,
            text="›",
            font=(FONT, 9, "bold"),
            fg=section_theme.accent,
            bg=section_theme.card_tint,
        ).pack(side=tk.LEFT, anchor=tk.N, padx=(0, 6))
        tk.Label(
            row,
            text=line,
            font=(FONT, 9),
            fg=THEME.TEXT_SOFT,
            bg=section_theme.card_tint,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=380,
        ).pack(side=tk.LEFT, anchor=tk.W)
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
        highlightbackground=THEME.BORDER,
        highlightthickness=1,
    )
    container.pack(fill=tk.BOTH, expand=True)
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)

    scrollbar = tk.Scrollbar(container, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

    log = tk.Text(
        container,
        font=(FONT, 9),
        fg=THEME.TEXT_SOFT,
        bg=THEME.SURFACE_ALT,
        relief=tk.FLAT,
        highlightthickness=0,
        wrap=tk.WORD,
        height=height,
        padx=8,
        pady=8,
        yscrollcommand=scrollbar.set,
        state=tk.DISABLED,
    )
    log.grid(row=0, column=0, sticky="nsew")
    scrollbar.config(command=log.yview)
    return log, card


def set_status_style(label, dot, message, section_theme=None, is_error=False, completed=False):
    section_theme = section_theme or EFT_THEME
    lowered = message.lower()
    is_success = completed or "completed" in lowered or "updated" in lowered
    if is_error:
        color = THEME.ERROR
        prefix = "✗ "
    elif is_success:
        color = THEME.SUCCESS
        prefix = "✓ "
    else:
        color = section_theme.accent
        prefix = "● "
    dot.configure(text=prefix.strip(), fg=color)
    label.configure(fg=THEME.TEXT if not is_error and not is_success else color)


def bind_hover(button, normal_color, hover_color):
    button.bind(
        "<Enter>",
        lambda _e, b=button, h=hover_color: b.configure(bg=h),
    )
    button.bind(
        "<Leave>",
        lambda _e, b=button, n=normal_color: b.configure(bg=n),
    )


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

PDF_DAILY_FILETYPES = [
    ("Elistar Daily PDF", "*.pdf"),
    ("All Files", "*.*"),
]

ANALISIS_MASTER_FILETYPES = [
    ("Bradenton Analisis C-Store", "*.xlsx"),
    ("Bradenton Analisis C-Store", "*.xlsm"),
    ("All Files", "*.*"),
]
