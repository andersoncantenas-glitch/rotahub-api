import logging
from tkinter import ttk


def apply_style(root):
    """Aplica estilos Tkinter (tema moderno, sem afetar a lógica do sistema)"""
    style = ttk.Style(root)

    # Em alguns Windows, "clam" funciona melhor e é mais consistente
    try:
        style.theme_use("clam")
    except Exception:
        logging.debug("Falha ignorada")

    # Paleta moderna
    PRIMARY = "#2B2F8F"
    PRIMARY_DARK = "#1F246F"
    BG = "#F4F6FB"
    CARD = "#FFFFFF"
    TEXT = "#1F2937"
    MUTED = "#6B7280"
    BORDER = "#E5E7EB"
    DANGER = "#B42318"
    DANGER_HOVER = "#D92D20"
    WARN = "#F79009"
    WARN_HOVER = "#FDB022"
    GHOST = "#EEF2F7"
    GHOST_HOVER = "#E5EAF3"

    # Tabelas (Treeview): linhas/colunas mais visíveis
    try:
        style.configure(
            "Treeview",
            rowheight=26,
            borderwidth=1,
            relief="solid",
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER,
        )
        style.map(
            "Treeview",
            background=[("selected", "#DDE6F8")],
            foreground=[("selected", TEXT)],
        )
        style.configure(
            "Treeview.Heading",
            borderwidth=1,
            relief="solid",
            background=BG,
            foreground=TEXT,
            font=("Segoe UI", 8, "bold"),
        )
    except Exception:
        logging.debug("Falha ignorada")

    # Base
    style.configure(".", font=("Segoe UI", 10))
    style.configure("Sidebar.TFrame", background=PRIMARY)
    style.configure("Content.TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="flat", borderwidth=0)
    style.configure("CardInset.TFrame", background=CARD, relief="flat", borderwidth=0)

    # Labels
    style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 15, "bold"))
    style.configure("CardLabel.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 8, "bold"))
    style.configure("CardStrong.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI Semibold", 10))
    style.configure("CardValue.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI Semibold", 22))
    style.configure("SidebarLogo.TLabel", background=PRIMARY, foreground="white", font=("Segoe UI", 16, "bold"))
    style.configure("SidebarSmall.TLabel", background=PRIMARY, foreground="#DDE3FF", font=("Segoe UI", 9))

    # Entradas
    style.configure("Field.TEntry", padding=(10, 8))
    style.map(
        "Field.TEntry",
        bordercolor=[("focus", PRIMARY), ("!focus", BORDER)],
        lightcolor=[("focus", PRIMARY), ("!focus", BORDER)],
        darkcolor=[("focus", PRIMARY), ("!focus", BORDER)],
    )

    # Botões (mantendo seus nomes de style)
    style.configure(
        "Side.TButton",
        background=PRIMARY,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
        padding=(12, 10),
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Side.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "SideActive.TButton",
        background=PRIMARY_DARK,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
        padding=(12, 10),
        focusthickness=0,
        focuscolor="none",
    )

    style.configure(
        "SideSub.TButton",
        background=PRIMARY,
        foreground="#EEF2FF",
        borderwidth=0,
        font=("Segoe UI", 8, "bold"),
        anchor="w",
        padding=(12, 8),
        focusthickness=0,
        focuscolor="none",
    )
    style.map("SideSub.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "SideSubActive.TButton",
        background=PRIMARY_DARK,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 8, "bold"),
        anchor="w",
        padding=(12, 8),
        focusthickness=0,
        focuscolor="none",
    )

    style.configure(
        "Primary.TButton",
        background=PRIMARY,
        foreground="white",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Primary.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "Ghost.TButton",
        background=GHOST,
        foreground=TEXT,
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Ghost.TButton", background=[("active", GHOST_HOVER)])

    style.configure(
        "Warn.TButton",
        background=WARN,
        foreground="black",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Warn.TButton", background=[("active", WARN_HOVER)])

    style.configure(
        "Danger.TButton",
        background=DANGER,
        foreground="white",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Danger.TButton", background=[("active", DANGER_HOVER)])

    # Separators
    style.configure("TSeparator", background=BORDER)

    # Treeview modernizado
    style.configure(
        "Treeview",
        background="white",
        foreground=TEXT,
        fieldbackground="white",
        rowheight=28,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        borderwidth=1,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Treeview.Heading",
        background="#F3F4F6",
        foreground=TEXT,
        relief="flat",
        font=("Segoe UI", 8, "bold"),
        padding=(8, 6),
    )
    style.map("Treeview", background=[("selected", "#DDE7FF")], foreground=[("selected", TEXT)])

    # Ajustes finais de padronizacao visual com acabamento mais liso.
    style.configure(".", background=BG, foreground=TEXT)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("Content.TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="flat", borderwidth=0)
    style.configure("CardInset.TFrame", background="#F5F8FE", relief="flat", borderwidth=0)
    style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI Semibold", 16))
    style.configure("CardLabel.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI Semibold", 9))
    style.configure("InsetTitle.TLabel", background="#F5F8FE", foreground=TEXT, font=("Segoe UI Semibold", 10))
    style.configure("InsetBody.TLabel", background="#F5F8FE", foreground=MUTED, font=("Segoe UI", 9))
    style.configure("InsetStrong.TLabel", background="#F5F8FE", foreground=TEXT, font=("Segoe UI Semibold", 10))
    style.configure("InsetValue.TLabel", background="#F5F8FE", foreground=TEXT, font=("Segoe UI Semibold", 21))
    style.configure("InsetValueSmall.TLabel", background="#F5F8FE", foreground="#1D4ED8", font=("Segoe UI Semibold", 16))
    style.configure("SidebarLogo.TLabel", background=PRIMARY, foreground="white", font=("Segoe UI Semibold", 17))
    style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI Semibold", 20))
    style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
    style.configure("Value.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
    style.configure("TCombobox", padding=(8, 6))
    style.configure("Side.TButton", font=("Segoe UI Semibold", 10), padding=(16, 12), borderwidth=0)
    style.configure("SideActive.TButton", font=("Segoe UI Semibold", 10), padding=(16, 12), borderwidth=0)
    style.configure("SideSub.TButton", font=("Segoe UI Semibold", 8), padding=(14, 8), borderwidth=0)
    style.configure("SideSubActive.TButton", font=("Segoe UI Semibold", 8), padding=(14, 8), borderwidth=0)
    style.configure("Primary.TButton", font=("Segoe UI Semibold", 10), padding=(16, 11), borderwidth=0)
    style.configure("Ghost.TButton", font=("Segoe UI Semibold", 10), padding=(16, 11), borderwidth=0)
    style.configure("CompactGhost.TButton", font=("Segoe UI Semibold", 9), padding=(8, 6), borderwidth=0)
    style.configure(
        "Toggle.TButton",
        background="#DBEAFE",
        foreground="#1D4ED8",
        font=("Segoe UI Semibold", 9),
        padding=(10, 6),
        borderwidth=0,
    )
    style.map("Toggle.TButton", background=[("active", "#BFDBFE")], foreground=[("active", "#1E40AF")])
    style.configure("Warn.TButton", foreground="white", font=("Segoe UI Semibold", 10), padding=(16, 11), borderwidth=0)
    style.configure("Danger.TButton", font=("Segoe UI Semibold", 10), padding=(16, 11), borderwidth=0)
    style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.configure("TNotebook.Tab", background=GHOST, foreground=TEXT, padding=(16, 8), font=("Segoe UI Semibold", 9), borderwidth=0)
    style.map("TNotebook.Tab", background=[("selected", CARD), ("active", GHOST_HOVER)])
    style.configure(
        "Treeview",
        background="white",
        foreground=TEXT,
        fieldbackground="white",
        rowheight=30,
        bordercolor="#D7E2F1",
        lightcolor="#D7E2F1",
        darkcolor="#D7E2F1",
        borderwidth=0,
        relief="flat",
        font=("Segoe UI", 9),
    )
    style.configure(
        "Treeview.Heading",
        background="#EAF2FF",
        foreground=TEXT,
        relief="flat",
        borderwidth=0,
        font=("Segoe UI Semibold", 9),
        padding=(10, 8),
    )


__all__ = ["apply_style"]
