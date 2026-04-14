import logging
import tkinter as tk
from tkinter import ttk


class PageBase(ttk.Frame):
    """Classe base para todas as páginas da aplicação"""

    def __init__(self, parent, app, page_name, title=None):
        super().__init__(parent, style="Content.TFrame")
        self.app = app
        self.page_name = page_name
        self.page_title_override = title

        # Estrutura
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self, style="Content.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)
        self.header = header

        self.lbl_title = ttk.Label(header, text=self._resolve_header_title(), style="Title.TLabel")
        self.lbl_title.grid(row=0, column=0, sticky="w")

        self.lbl_status = ttk.Label(
            header,
            text="STATUS: -",
            style="Subtitle.TLabel",
        )
        self.lbl_status.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.header_right = ttk.Frame(header, style="Content.TFrame")
        self.header_right.grid(row=0, column=1, rowspan=2, sticky="e")

        ttk.Separator(self).grid(row=1, column=0, sticky="ew")

        body_shell = ttk.Frame(self, style="Content.TFrame")
        body_shell.grid(row=2, column=0, sticky="nsew")
        body_shell.grid_rowconfigure(0, weight=1)
        body_shell.grid_columnconfigure(0, weight=1)
        self.body_shell = body_shell

        body_bg = ttk.Style(self).lookup("Content.TFrame", "background") or "#F4F6FB"
        self.body_canvas = tk.Canvas(
            body_shell,
            highlightthickness=0,
            bd=0,
            relief="flat",
            background=body_bg,
        )
        self.body_canvas.grid(row=0, column=0, sticky="nsew")
        self.body_canvas._wheel_scroll_target = True

        self.body_scrollbar = ttk.Scrollbar(body_shell, orient="vertical", command=self.body_canvas.yview)
        self.body_scrollbar.grid(row=0, column=1, sticky="ns")
        self.body_xscrollbar = ttk.Scrollbar(body_shell, orient="horizontal", command=self.body_canvas.xview)
        self.body_xscrollbar.grid(row=1, column=0, sticky="ew")
        self.body_canvas.configure(
            yscrollcommand=self.body_scrollbar.set,
            xscrollcommand=self.body_xscrollbar.set,
        )

        self.body = ttk.Frame(self.body_canvas, style="Content.TFrame", padding=(22, 18))
        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(1, weight=1)
        self._body_window = self.body_canvas.create_window((0, 0), window=self.body, anchor="nw")

        def _sync_body_scrollregion(_event=None):
            try:
                self.body_canvas.configure(scrollregion=self.body_canvas.bbox("all"))
            except Exception:
                logging.debug("Falha ignorada")

        def _fit_body_size(event):
            try:
                req_w = max(self.body.winfo_reqwidth(), event.width)
                req_h = max(self.body.winfo_reqheight(), event.height)
                self.body_canvas.itemconfigure(self._body_window, width=req_w, height=req_h)
            except Exception:
                logging.debug("Falha ignorada")

        self.body.bind("<Configure>", _sync_body_scrollregion, add="+")
        self.body_canvas.bind("<Configure>", _fit_body_size, add="+")

        ttk.Separator(self).grid(row=3, column=0, sticky="ew")

        footer = ttk.Frame(self, style="Content.TFrame", padding=(22, 14))
        footer.grid(row=4, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)

        self.footer_left = ttk.Frame(footer, style="Content.TFrame")
        self.footer_left.grid(row=0, column=0, sticky="w")

        self.footer_right = ttk.Frame(footer, style="Content.TFrame")
        self.footer_right.grid(row=0, column=1, sticky="e")

    def _resolve_header_title(self) -> str:
        if self.page_title_override:
            return str(self.page_title_override)
        if hasattr(self.app, "get_routine_title"):
            return self.app.get_routine_title(self.page_name)
        return str(self.page_name)

    def refresh_header_title(self):
        """Reaplica o titulo visual da rotina quando houver renumeracao central."""
        self.lbl_title.config(text=self._resolve_header_title())

    def set_status(self, txt):
        """Atualiza texto de status na página"""
        self.lbl_status.config(text=txt)

    def on_show(self):
        """Método chamado quando a página é exibida (pode ser sobrescrito)"""
        pass


__all__ = ["PageBase"]
