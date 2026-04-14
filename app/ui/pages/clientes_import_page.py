# -*- coding: utf-8 -*-
import logging
from tkinter import filedialog, messagebox, ttk

from app.services.cliente_service import (
    fetch_clientes_rows,
    importar_clientes_excel,
    salvar_clientes_linhas,
    sync_cliente_upsert_api,
)
from app.ui.components.tree_helpers import enable_treeview_sorting, tree_insert_aligned
from app.utils.async_ui import run_async_ui
from app.utils.excel_helpers import (
    require_excel_support,
    require_pandas,
    upper,
)


def configure_clientes_import_page_dependencies(dependencies=None, **kwargs):
    merged = {}
    if dependencies:
        merged.update(dict(dependencies))
    if kwargs:
        merged.update(kwargs)
    globals().update(merged)


def is_desktop_api_sync_enabled():
    # Fallback seguro: sem injeção explícita, mantém modo local.
    return False


def ensure_system_api_binding(context: str = "Operacao", parent=None, force_probe: bool = False) -> bool:
    # Fallback seguro: impede operação remota quando a injeção não foi configurada.
    logging.error("Dependencia nao configurada para ClientesImportPage: ensure_system_api_binding (%s)", context)
    return False


class ClientesImportPage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, style="Content.TFrame")
        self.app = app
        self._editing = None
        self._load_seq = 0
        self._import_seq = 0
        self._save_seq = 0

        card = ttk.Frame(self, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        ttk.Label(card, text="Clientes (Base do Wibi)", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        actions.grid_columnconfigure(4, weight=1)

        self.btn_importar = ttk.Button(actions, text="IMPORTAR CLIENTES (EXCEL)", style="Warn.TButton",
                                       command=self.importar_clientes_excel)
        self.btn_importar.grid(row=0, column=0, padx=6)

        self.btn_atualizar = ttk.Button(actions, text="ATUALIZAR", style="Ghost.TButton",
                                        command=self.carregar)
        self.btn_atualizar.grid(row=0, column=1, padx=6)

        self.btn_inserir = ttk.Button(actions, text="\u2795 INSERIR LINHA", style="Ghost.TButton",
                                      command=self.inserir_linha)
        self.btn_inserir.grid(row=0, column=2, padx=6)

        self.btn_salvar = ttk.Button(actions, text="SALVAR ALTERAÇÕES", style="Primary.TButton",
                                     command=self.salvar_alteracoes)
        self.btn_salvar.grid(row=0, column=3, padx=6)

        self.lbl_info = ttk.Label(
            actions,
            text="Dica: duplo clique na célula para editar. ENTER salva a célula. ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_info.grid(row=0, column=4, padx=(12, 0), sticky="e")

        self.pb_loading = ttk.Progressbar(actions, mode="indeterminate", length=150)
        self.pb_loading.grid(row=1, column=0, columnspan=2, padx=6, pady=(8, 0), sticky="w")
        self.lbl_loading = ttk.Label(
            actions,
            text="",
            background="white",
            foreground="#2563EB",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_loading.grid(row=1, column=2, columnspan=3, padx=(8, 6), pady=(8, 0), sticky="w")

        cols = ["CÓD CLIENTE", "NOME CLIENTE", "ENDEREÇO", "TELEFONE", "VENDEDOR"]
        table_wrap = ttk.Frame(card, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=14)
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=190, minwidth=120, anchor="w")

        self.tree.column("ENDEREÇO", width=420, minwidth=200)
        self.tree.bind("<Double-1>", self._start_edit_cell)

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CÓD CLIENTE"},
            money_cols=set(),
            date_cols=set()
        )

        # Se rolar o tree durante edição, fecha/commita corretamente
        self.tree.bind("<MouseWheel>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-4>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-5>", self._on_tree_scroll, add=True)

        self.carregar()

    def _set_busy(self, busy, text=""):
        state = "disabled" if busy else "normal"
        for btn in (self.btn_importar, self.btn_atualizar, self.btn_inserir, self.btn_salvar):
            try:
                btn.configure(state=state)
            except Exception:
                logging.debug("Falha ao ajustar estado de botao da tela de clientes")
        if busy:
            self.lbl_loading.configure(text=text or "Processando...")
            try:
                self.pb_loading.start(10)
            except Exception:
                logging.debug("Falha ao iniciar progressbar da tela de clientes")
        else:
            try:
                self.pb_loading.stop()
            except Exception:
                logging.debug("Falha ao parar progressbar da tela de clientes")
            self.lbl_loading.configure(text=text or "")

    def _get_row_values(self, iid):
        vals = self.tree.item(iid, "values") or ("", "", "", "", "")
        vals = list(vals) + [""] * (5 - len(vals))
        return [str(v or "").strip() for v in vals[:5]]

    def _is_blank_row(self, cod, nome, endereco, telefone, vendedor):
        return not (cod or nome or endereco or telefone or vendedor)

    def _normalize_cod(self, cod):
        return str(cod or "").strip()

    def _apply_clientes_rows(self, seq, rows):
        if seq != self._load_seq or not self.winfo_exists():
            return
        if isinstance(rows, dict):
            if not bool(rows.get("ok", False)):
                self._handle_clientes_load_error(seq, RuntimeError(rows.get("error") or "Falha ao carregar clientes."))
                return
            rows = rows.get("data") or []
        self._set_busy(False, f"Lista carregada: {len(rows)} clientes")
        self.tree.delete(*self.tree.get_children())

        for r in rows:
            tree_insert_aligned(self.tree, "", "end", (r[0] or "", r[1] or "", r[2] or "", r[3] or "", r[4] or ""))

        if self.app and hasattr(self.app, "refresh_programacao_comboboxes"):
            self.app.refresh_programacao_comboboxes()

    def _handle_clientes_load_error(self, seq, exc):
        if seq != self._load_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Falha ao carregar clientes")
        logging.error("Falha ao carregar clientes", exc_info=(type(exc), exc, exc.__traceback__))

    def carregar(self, async_load=True):
        self._load_seq += 1
        seq = self._load_seq
        self._set_busy(True, "Carregando lista de clientes...")
        if not async_load:
            try:
                rows = fetch_clientes_rows(is_desktop_api_sync_enabled=is_desktop_api_sync_enabled)
            except Exception as exc:
                self._handle_clientes_load_error(seq, exc)
                return
            self._apply_clientes_rows(seq, rows)
            return

        run_async_ui(
            self,
            lambda: fetch_clientes_rows(is_desktop_api_sync_enabled=is_desktop_api_sync_enabled),
            lambda rows, seq=seq: self._apply_clientes_rows(seq, rows),
            lambda exc, seq=seq: self._handle_clientes_load_error(seq, exc),
        )

    def inserir_linha(self):
        tree_insert_aligned(self.tree, "", "end", ("", "", "", "", ""))
        items = self.tree.get_children()
        if items:
            self.tree.see(items[-1])

    def _on_tree_scroll(self, event=None):
        if self._editing:
            self._commit_edit()

    def _start_edit_cell(self, event=None):
        if self._editing:
            self._commit_edit()

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1
        vals = self._get_row_values(row_id)
        if col_index < 0 or col_index >= len(vals):
            return

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return

        x, y, w, h = bbox
        value = vals[col_index]

        entry = ttk.Entry(self.tree, style="Field.TEntry")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus_set()
        entry.selection_range(0, "end")

        entry.bind("<Return>", lambda e: self._commit_edit())
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        entry.bind("<FocusOut>", lambda e: self._commit_edit())

        self._editing = (row_id, col_index, entry)

    def _commit_edit(self):
        if not self._editing:
            return

        row_id, col_index, entry = self._editing
        try:
            new_value = entry.get()
        except Exception:
            new_value = ""

        vals = self._get_row_values(row_id)
        if 0 <= col_index < len(vals):
            vals[col_index] = str(new_value).strip()
            self.tree.item(row_id, values=tuple(vals))

        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    def _cancel_edit(self):
        if not self._editing:
            return
        _, _, entry = self._editing
        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    def salvar_alteracoes(self):
        if not ensure_system_api_binding(context="Salvar alteracoes de clientes", parent=self):
            return
        if self._editing:
            self._commit_edit()

        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("ATENÇÃO", "Não há dados para salvar.")
            return

        linhas = []
        cods_seen = set()

        for iid in items:
            cod, nome, endereco, telefone, vendedor = self._get_row_values(iid)

            if self._is_blank_row(cod, nome, endereco, telefone, vendedor):
                continue

            cod_norm = self._normalize_cod(cod)
            nome_norm = str(nome or "").strip()

            if not cod_norm or not nome_norm:
                messagebox.showwarning("ATENÇÃO", "Todas as linhas precisam ter pelo menos CÓD CLIENTE e NOME CLIENTE.")
                return

            cod_key = upper(cod_norm)
            if cod_key in cods_seen:
                messagebox.showwarning("ATENÇÃO", f"CÓD CLIENTE duplicado na tabela: {cod_norm}\n\nCorrija antes de salvar.")
                return
            cods_seen.add(cod_key)

            linhas.append((upper(cod_norm), upper(nome_norm), upper(endereco), upper(telefone), upper(vendedor)))

        if not linhas:
            messagebox.showwarning("ATENÇÃO", "Nenhuma linha válida para salvar.")
            return

        self._save_seq += 1
        seq = self._save_seq
        self._set_busy(True, "Salvando clientes...")
        run_async_ui(
            self,
            lambda linhas=tuple(linhas): salvar_clientes_linhas(
                linhas,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            ),
            lambda msg, seq=seq: self._on_salvar_clientes_done(seq, msg),
            lambda exc, seq=seq: self._on_salvar_clientes_error(seq, exc),
        )

    def _on_salvar_clientes_done(self, seq, msg):
        if seq != self._save_seq or not self.winfo_exists():
            return
        if isinstance(msg, dict):
            if not bool(msg.get("ok", False)):
                self._on_salvar_clientes_error(seq, RuntimeError(msg.get("error") or "Falha ao salvar clientes."))
                return
            msg = msg.get("data") or "Clientes salvos"
        self._set_busy(False, "Clientes salvos")
        messagebox.showinfo("OK", msg)
        self.carregar()

    def _on_salvar_clientes_error(self, seq, exc):
        if seq != self._save_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Falha ao salvar clientes")
        logging.error("Falha ao salvar clientes", exc_info=(type(exc), exc, exc.__traceback__))
        messagebox.showerror("ERRO", str(exc))

    def _on_importar_clientes_done(self, seq, msg):
        if seq != self._import_seq or not self.winfo_exists():
            return
        if isinstance(msg, dict):
            if not bool(msg.get("ok", False)):
                self._on_importar_clientes_error(seq, RuntimeError(msg.get("error") or "Falha ao importar clientes."))
                return
            msg = msg.get("data") or "Importação concluída"
        messagebox.showinfo("OK", msg)
        self.carregar()

    def _on_importar_clientes_error(self, seq, exc):
        if seq != self._import_seq or not self.winfo_exists():
            return
        logging.error("Falha ao importar clientes via Excel", exc_info=(type(exc), exc, exc.__traceback__))
        messagebox.showerror("ERRO", str(exc))

    def importar_clientes_excel(self):
        if not ensure_system_api_binding(context="Importar clientes via planilha", parent=self):
            return
        path = filedialog.askopenfilename(
            title="IMPORTAR CLIENTES (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return None

        if not (require_pandas() and require_excel_support(path)):
            return

        self._import_seq += 1
        seq = self._import_seq
        self._set_busy(True, "Importando clientes do Excel...")
        run_async_ui(
            self,
            lambda path=path: importar_clientes_excel(
                path,
                is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            ),
            lambda msg, seq=seq: self._on_importar_clientes_done(seq, msg),
            lambda exc, seq=seq: self._on_importar_clientes_error(seq, exc),
        )
        return


__all__ = ["ClientesImportPage", "configure_clientes_import_page_dependencies"]
