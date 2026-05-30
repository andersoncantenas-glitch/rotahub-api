# -*- coding: utf-8 -*-
import logging
from tkinter import filedialog, messagebox, ttk

from app.services.cliente_service import (
    fetch_cliente_historico,
    fetch_cliente_localizacoes,
    fetch_clientes_dashboard,
    fetch_clientes_lookup,
    fetch_clientes_rows,
    importar_clientes_excel,
    salvar_clientes_linhas,
)
from app.ui.components.tree_helpers import enable_treeview_sorting, tree_insert_aligned
from app.utils.async_ui import run_async_ui
from app.utils.excel_helpers import (
    require_excel_support,
    require_pandas,
    upper,
)
from app.utils.formatters import safe_float, safe_int


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
        self._clientes_lookup = []

        card = ttk.Frame(self, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(3, weight=1)

        ttk.Label(card, text="Clientes", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.dashboard = ttk.Frame(card, style="Card.TFrame")
        self.dashboard.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        for i in range(3):
            self.dashboard.grid_columnconfigure(i, weight=1)
        self._build_dashboard_tiles()

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        actions.grid_columnconfigure(4, weight=1)
        self.actions = actions

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
        table_wrap.grid(row=3, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)
        self.table_wrap = table_wrap

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

        self.detail_wrap = ttk.Frame(card, style="Card.TFrame")
        self.detail_wrap.grid(row=3, column=0, sticky="nsew")
        self.detail_wrap.grid_columnconfigure(0, weight=1)
        self.detail_wrap.grid_rowconfigure(1, weight=1)

        self.actions.grid_remove()
        self.table_wrap.grid_remove()
        self._show_historico()
        self._refresh_dashboard()

    def _build_dashboard_tiles(self):
        self.lbl_dash_total = self._tile(
            self.dashboard, 0, "Lista de clientes", "0 clientes", "Inserir, editar e importar base do Wibi",
            self._show_lista_clientes,
        )
        self.lbl_dash_hist = self._tile(
            self.dashboard, 1, "Historico", "0 com historico", "Entregas, cancelamentos, alteracoes, KG e mortalidade",
            self._show_historico,
        )
        self.lbl_dash_loc = self._tile(
            self.dashboard, 2, "Localizacao", "0 com localizacao", "Pontos registrados pelo app motorista na entrega",
            self._show_localizacoes,
        )

    def _tile(self, parent, col, title, value, hint, command):
        box = ttk.Frame(parent, style="Card.TFrame", padding=10)
        box.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0 if col == 2 else 6))
        box.grid_columnconfigure(0, weight=1)
        ttk.Label(box, text=title, background="white", foreground="#0F172A", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        lbl_value = ttk.Label(box, text=value, background="white", foreground="#0891B2", font=("Segoe UI", 16, "bold"))
        lbl_value.grid(row=1, column=0, sticky="w", pady=(4, 2))
        ttk.Label(box, text=hint, background="white", foreground="#64748B", font=("Segoe UI", 8, "bold"), wraplength=330).grid(row=2, column=0, sticky="w")
        ttk.Button(box, text="ACESSAR", style="Ghost.TButton", command=command).grid(row=3, column=0, sticky="ew", pady=(10, 0))
        return lbl_value

    def _refresh_dashboard(self):
        result = fetch_clientes_dashboard()
        if not isinstance(result, dict) or not result.get("ok"):
            return
        data = result.get("data") or {}
        self.lbl_dash_total.configure(text=f"{safe_int(data.get('total_clientes'), 0)} clientes")
        self.lbl_dash_hist.configure(text=f"{safe_int(data.get('clientes_com_historico'), 0)} com historico")
        self.lbl_dash_loc.configure(text=f"{safe_int(data.get('clientes_com_localizacao'), 0)} clientes / {safe_int(data.get('amostras_localizacao'), 0)} pontos")

    def _clear_detail(self):
        self.actions.grid_remove()
        self.table_wrap.grid_remove()
        self.detail_wrap.grid()
        for child in self.detail_wrap.winfo_children():
            child.destroy()

    def _show_lista_clientes(self):
        self.detail_wrap.grid_remove()
        self.actions.grid()
        self.table_wrap.grid()
        if not self.tree.get_children():
            self.carregar()

    def _fmt_kg(self, value):
        return f"{safe_float(value, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _load_cliente_options(self):
        if not self._clientes_lookup:
            result = fetch_clientes_lookup()
            if isinstance(result, dict) and result.get("ok"):
                self._clientes_lookup = list(result.get("data") or [])
        return [f"{cod} - {nome}" for cod, nome in self._clientes_lookup]

    def _selected_cod(self, combo):
        raw = str(combo.get() or "").strip()
        if " - " in raw:
            return raw.split(" - ", 1)[0].strip()
        return raw.strip()

    def _build_cliente_selector(self, parent, title, command):
        head = ttk.Frame(parent, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        head.grid_columnconfigure(1, weight=1)
        ttk.Label(head, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12))
        combo = ttk.Combobox(head, values=self._load_cliente_options(), state="normal")
        combo.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(head, text="ATUALIZAR", style="Ghost.TButton", command=lambda: command(combo)).grid(row=0, column=2, sticky="e")
        if combo["values"]:
            combo.set(combo["values"][0])
        combo.bind("<Return>", lambda _e: command(combo))
        return combo

    def _show_historico(self):
        self._clear_detail()
        combo = self._build_cliente_selector(self.detail_wrap, "Historico do cliente", self._carregar_historico)
        resumo = ttk.Frame(self.detail_wrap, style="Card.TFrame")
        resumo.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        for i in range(7):
            resumo.grid_columnconfigure(i, weight=1)
        self.hist_resumo_labels = {}
        for col, key, label in [
            (0, "total_programacoes", "Programacoes"),
            (1, "entregues", "Entregues"),
            (2, "canceladas", "Canceladas"),
            (3, "alteradas", "Alteradas"),
            (4, "mortalidade_aves", "Mortalidade"),
            (5, "kg_recebidos", "KG recebidos"),
            (6, "kg_descontados", "KG descontados"),
        ]:
            box = ttk.Frame(resumo, style="Card.TFrame", padding=8)
            box.grid(row=0, column=col, sticky="nsew", padx=3)
            ttk.Label(box, text=label, background="white", foreground="#64748B", font=("Segoe UI", 8, "bold")).pack(anchor="w")
            lbl = ttk.Label(box, text="0", background="white", foreground="#0F172A", font=("Segoe UI", 12, "bold"))
            lbl.pack(anchor="w")
            self.hist_resumo_labels[key] = lbl

        cols = ("data", "programacao", "pedido", "status", "cx", "kg_rec", "kg_desc", "mort", "valor", "motorista")
        self.hist_tree = ttk.Treeview(self.detail_wrap, columns=cols, show="headings", height=12)
        self.hist_tree.grid(row=2, column=0, sticky="nsew")
        self.detail_wrap.grid_rowconfigure(2, weight=1)
        labels = {
            "data": "DATA", "programacao": "PROGRAMACAO", "pedido": "PEDIDO", "status": "STATUS",
            "cx": "CX", "kg_rec": "KG RECEB.", "kg_desc": "KG DESC.", "mort": "MORT.", "valor": "VALOR REC.", "motorista": "MOTORISTA",
        }
        for c in cols:
            self.hist_tree.heading(c, text=labels[c])
            self.hist_tree.column(c, width=120, minwidth=80, anchor="w")
        self.hist_tree.column("motorista", width=180)
        vsb = ttk.Scrollbar(self.detail_wrap, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=2, column=1, sticky="ns")
        self._carregar_historico(combo)

    def _carregar_historico(self, combo):
        cod = self._selected_cod(combo)
        if not cod:
            return
        result = fetch_cliente_historico(cod)
        if not isinstance(result, dict) or not result.get("ok"):
            messagebox.showerror("ERRO", (result or {}).get("error") or "Falha ao carregar historico.")
            return
        data = result.get("data") or {}
        resumo = data.get("resumo") or {}
        for key, lbl in self.hist_resumo_labels.items():
            val = resumo.get(key, 0)
            lbl.configure(text=self._fmt_kg(val) if key.startswith("kg_") else str(safe_int(val, 0)))
        self.hist_tree.delete(*self.hist_tree.get_children())
        for r in data.get("rows") or []:
            self.hist_tree.insert("", "end", values=(
                r.get("data_ref", ""),
                r.get("codigo_programacao", ""),
                r.get("pedido", ""),
                r.get("status_pedido", ""),
                f"{safe_int(r.get('caixas_atuais'), 0)}/{safe_int(r.get('caixas_programadas'), 0)}",
                self._fmt_kg(r.get("kg_recebido")),
                self._fmt_kg(r.get("kg_descontado")),
                safe_int(r.get("mortalidade_aves"), 0),
                self._fmt_kg(r.get("valor_recebido")),
                r.get("motorista", ""),
            ))

    def _show_localizacoes(self):
        self._clear_detail()
        combo = self._build_cliente_selector(self.detail_wrap, "Localizacao do cliente", self._carregar_localizacoes)
        self.loc_info = ttk.Label(
            self.detail_wrap,
            text="",
            background="white",
            foreground="#0F172A",
            font=("Segoe UI", 9, "bold"),
        )
        self.loc_info.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        cols = ("quando", "programacao", "pedido", "status", "latitude", "longitude", "endereco", "motorista", "origem")
        self.loc_tree = ttk.Treeview(self.detail_wrap, columns=cols, show="headings", height=13)
        self.loc_tree.grid(row=2, column=0, sticky="nsew")
        self.detail_wrap.grid_rowconfigure(2, weight=1)
        labels = {
            "quando": "REGISTRADO EM", "programacao": "PROGRAMACAO", "pedido": "PEDIDO", "status": "STATUS",
            "latitude": "LATITUDE", "longitude": "LONGITUDE", "endereco": "ENDERECO", "motorista": "MOTORISTA", "origem": "ORIGEM",
        }
        for c in cols:
            self.loc_tree.heading(c, text=labels[c])
            self.loc_tree.column(c, width=130, minwidth=80, anchor="w")
        self.loc_tree.column("endereco", width=260)
        vsb = ttk.Scrollbar(self.detail_wrap, orient="vertical", command=self.loc_tree.yview)
        self.loc_tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=2, column=1, sticky="ns")
        self._carregar_localizacoes(combo)

    def _carregar_localizacoes(self, combo):
        cod = self._selected_cod(combo)
        if not cod:
            return
        result = fetch_cliente_localizacoes(cod)
        if not isinstance(result, dict) or not result.get("ok"):
            messagebox.showerror("ERRO", (result or {}).get("error") or "Falha ao carregar localizacoes.")
            return
        data = result.get("data") or {}
        resumo = data.get("resumo") or {}
        ultima = resumo.get("ultima") or {}
        self.loc_info.configure(
            text=(
                f"Amostras: {safe_int(resumo.get('amostras'), 0)} | "
                f"Com coordenada: {safe_int(resumo.get('com_coordenada'), 0)} | "
                f"Ultima: {ultima.get('registrado_em') or '-'}"
            )
        )
        self.loc_tree.delete(*self.loc_tree.get_children())
        for r in data.get("rows") or []:
            endereco = " - ".join([str(r.get(k) or "").strip() for k in ("endereco", "bairro", "cidade") if str(r.get(k) or "").strip()])
            self.loc_tree.insert("", "end", values=(
                r.get("registrado_em", ""),
                r.get("codigo_programacao", ""),
                r.get("pedido", ""),
                r.get("status_pedido", ""),
                "" if r.get("latitude") is None else r.get("latitude"),
                "" if r.get("longitude") is None else r.get("longitude"),
                endereco,
                r.get("motorista_nome") or r.get("motorista_codigo") or "",
                r.get("origem", ""),
            ))

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
        self._clientes_lookup = []
        self._refresh_dashboard()
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
