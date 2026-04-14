# -*- coding: utf-8 -*-
import logging
import os
import sqlite3
import tkinter as tk
import urllib
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except Exception:
    pd = None

from app.db.connection import get_db
from app.db.migrations import table_has_column
from app.security.passwords import hash_password_pbkdf2, verify_password_pbkdf2
from app.services.api_client import SyncError, _call_api
from app.services.cliente_service import sync_cliente_upsert_api
from app.services.motorista_service import (
    extract_motorista_seq,
    fetch_motoristas_rows,
    next_motorista_codigo,
    sync_motorista_upsert_api,
)
from app.services.vendedor_service import (
    fetch_vendedores_rows,
    sync_vendedor_upsert_api,
    update_vendedor_password_hash_local,
)
from app.repositories.motorista_repository import update_motorista_status_local
from app.repositories.programacao_repository import db_has_column
from app.services.api_binding import ensure_system_api_binding
from app.services.runtime_flags import can_read_from_api, is_desktop_api_sync_enabled
from app.ui.components.entry_helpers import bind_entry_smart
from app.ui.components.tree_helpers import enable_treeview_sorting, tree_insert_aligned
from app.utils.async_ui import run_async_ui
from app.utils.excel_helpers import (
    excel_engine_for,
    guess_col,
    require_excel_support,
    require_pandas,
    upper,
)
from app.utils.formatters import safe_int
from app.utils.validators import (
    is_valid_cpf,
    is_valid_motorista_codigo,
    is_valid_motorista_senha,
    is_valid_phone,
    normalize_cpf,
    normalize_phone,
)


class CadastroCRUD(ttk.Frame):
    """Componente CRUD reutilizável para cadastros (com validações por tabela)"""
    def __init__(self, parent, titulo, table, fields, app=None):
        super().__init__(parent, style="Content.TFrame")
        self.table = table
        self.fields = fields
        self.app = app
        self._body_window = None
        self.selected_id = None
        self._is_admin = bool(getattr(self.app, "user", {}).get("is_admin")) if self.app else False
        self._edit_mode = "view"  # view | novo | status
        self._has_status_field = any(col == "status" for col, _ in fields)
        self._load_seq = 0

        # Card
        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True)

        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(3, weight=1)
        card.grid_rowconfigure(6, weight=1)

        ttk.Label(card, text=titulo, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        # FORM responsivo
        form = ttk.Frame(card, style="Card.TFrame")
        form.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        form.grid_columnconfigure(0, weight=1)

        self.entries = {}

        max_cols = 4
        for idx, (col, label) in enumerate(fields):
            r = (idx // max_cols) * 2
            c = idx % max_cols

            form.grid_columnconfigure(c, weight=1, uniform="formcols")

            ttk.Label(form, text=label, style="CardLabel.TLabel").grid(
                row=r, column=c, sticky="w", padx=8, pady=(0, 6)
            )

            if col == "status" and self.table in {"ajudantes", "motoristas", "vendedores"}:
                if self.table == "motoristas":
                    ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "INATIVO"])
                elif self.table == "vendedores":
                    ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "DESATIVADO"])
                else:
                    ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "DESATIVADO"])
                ent.set("ATIVO")
            elif col == "permissoes" and self.table == "usuarios":
                ent = ttk.Combobox(form, state="readonly", values=["ADMIN", "GERENTE", "OPERADOR", "VISUALIZADOR"])
                ent.set("OPERADOR")
            else:
                ent = ttk.Entry(form, style="Field.TEntry")
            ent.grid(row=r + 1, column=c, sticky="ew", padx=8, pady=(0, 12))
            self.entries[col] = ent
            if isinstance(ent, ttk.Entry):
                kind, precision = self._infer_entry_kind(col, label)
                bind_entry_smart(ent, kind, precision=precision)

        # BOTOES
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(6, 14))
        for idx in range(0, 8):
            actions.grid_columnconfigure(idx, weight=1)
        actions.grid_columnconfigure(20, weight=1)

        self.btn_novo = ttk.Button(actions, text="NOVO", style="Ghost.TButton", command=self.novo)
        self.btn_novo.grid(row=0, column=0, padx=4, sticky="ew")
        self.btn_alterar = ttk.Button(actions, text="ALTERAR", style="Ghost.TButton", command=self.alterar)
        self.btn_alterar.grid(row=0, column=1, padx=4, sticky="ew")
        self.btn_salvar = ttk.Button(actions, text="SALVAR", style="Primary.TButton", command=self.salvar)
        self.btn_salvar.grid(row=0, column=2, padx=4, sticky="ew")
        cmd_liberar = (lambda: self._set_status_rapido(True))
        cmd_bloquear = (lambda: self._set_status_rapido(False))
        if self.table == "motoristas":
            cmd_liberar = (lambda: self._toggle_motorista_access(True))
            cmd_bloquear = (lambda: self._toggle_motorista_access(False))

        self.btn_liberar = ttk.Button(actions, text="LIBERAR", style="Primary.TButton", command=cmd_liberar)
        self.btn_liberar.grid(row=0, column=3, padx=4, sticky="ew")
        self.btn_bloquear = ttk.Button(actions, text="BLOQUEAR", style="Danger.TButton", command=cmd_bloquear)
        self.btn_bloquear.grid(row=0, column=4, padx=4, sticky="ew")

        if self.table in {"usuarios", "motoristas", "vendedores"}:
            if self.table == "motoristas":
                cmd_senha = self.alterar_senha_motorista
            elif self.table == "vendedores":
                cmd_senha = self.alterar_senha_vendedor
            else:
                cmd_senha = self.alterar_senha
            self.btn_senha = ttk.Button(
                actions,
                text="ALTERAR SENHA",
                style="Warn.TButton",
                command=cmd_senha,
            )
            self.btn_senha.grid(row=0, column=5, padx=4, sticky="ew")

        if self.table == "ajudantes":
            self._ajudantes_status_filter = tk.StringVar(value="TODOS")
            ttk.Label(actions, text="FILTRO STATUS:", style="CardLabel.TLabel").grid(row=0, column=10, padx=(18, 6), sticky="e")
            cb_status_filter = ttk.Combobox(
                actions,
                state="readonly",
                width=14,
                textvariable=self._ajudantes_status_filter,
                values=["TODOS", "ATIVO", "DESATIVADO"],
            )
            cb_status_filter.grid(row=0, column=11, padx=6, sticky="w")
            cb_status_filter.bind("<<ComboboxSelected>>", lambda _e: self.carregar())

        if self.table == "clientes":
            self.btn_importar_clientes = ttk.Button(
                actions,
                text="IMPORTAR CLIENTES (EXCEL)",
                style="Warn.TButton",
                command=self.importar_clientes_excel,
            )
            self.btn_importar_clientes.grid(row=0, column=6, padx=6)

        # Tabela
        cols = ["ID"] + [label for _, label in self.fields]
        table_wrap = ttk.Frame(card, style="Card.TFrame")
        table_wrap.grid(row=3, column=0, sticky="nsew", pady=(0, 6))
        card.grid_rowconfigure(3, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=12)
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew")

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=160, minwidth=90, anchor="w")
        self.tree.column("ID", width=70, minwidth=60, anchor="center")

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ID"},
            money_cols=set(),
            date_cols=set()
        )

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.carregar()
        self.limpar()
        self._set_form_mode("view")
        self._update_password_controls()

    def _legacy_vincular_selecionadas_programacao_api_only(self):
        # Mantido apenas por compatibilidade com binds legados.
        return self.vincular_selecionadas_programacao()

    # -------------------------
    # Helpers
    # -------------------------
    def _legacy_vincular_selecionadas_programacao_sem_caixas(self):
        # Mantido apenas por compatibilidade com binds legados.
        return self.vincular_selecionadas_programacao()

    def _norm(self, v):
        return upper(str(v or "").strip())

    def _extract_motorista_seq(self, codigo: str) -> int:
        return extract_motorista_seq(codigo)

    def _next_motorista_codigo(self, cur=None) -> str:
        result = next_motorista_codigo(can_read_from_api=can_read_from_api, cur=cur)
        if isinstance(result, dict):
            codigo = str(result.get("data") or "").strip().upper()
            if codigo:
                return codigo
            return "MOT-01"
        return str(result or "MOT-01")

    def _set_widget_enabled(self, w, enabled: bool):
        try:
            if isinstance(w, ttk.Combobox):
                if str(getattr(w, "cget")("state") or "") == "readonly":
                    w.configure(state="readonly" if enabled else "disabled")
                else:
                    w.configure(state="normal" if enabled else "disabled")
            else:
                w.configure(state="normal" if enabled else "disabled")
        except Exception:
            logging.debug("Falha ignorada")

    def _set_form_mode(self, mode: str):
        self._edit_mode = str(mode or "view").lower()
        if self._edit_mode not in {"view", "novo", "status"}:
            self._edit_mode = "view"

        for col, _ in self.fields:
            w = self.entries.get(col)
            if not w:
                continue
            if self._edit_mode == "novo":
                self._set_widget_enabled(w, True)
            elif self._edit_mode == "status":
                self._set_widget_enabled(w, col == "status")
            else:
                self._set_widget_enabled(w, False)

        if self._edit_mode == "novo" and self.table == "motoristas" and "codigo" in self.entries:
            self._set("codigo", self._next_motorista_codigo())

        try:
            has_sel = bool(self.selected_id)
            status_ops = self._has_status_field and has_sel
            if hasattr(self, "btn_alterar"):
                self.btn_alterar.config(state="normal" if status_ops else "disabled")
            if hasattr(self, "btn_liberar"):
                self.btn_liberar.config(state="normal" if status_ops else "disabled")
            if hasattr(self, "btn_bloquear"):
                self.btn_bloquear.config(state="normal" if status_ops else "disabled")
        except Exception:
            logging.debug("Falha ignorada")

    def novo(self):
        self.selected_id = None
        for col, _ in self.fields:
            self._set(col, "")
        if self.table == "ajudantes" and "status" in self.entries:
            self._set("status", "ATIVO")
        if self.table == "motoristas" and "status" in self.entries:
            self._set("status", "ATIVO")
        if self.table == "vendedores" and "status" in self.entries:
            self._set("status", "ATIVO")
        self._set_form_mode("novo")
        self._update_password_controls()

    def _set_status_rapido(self, liberar: bool):
        if not self._has_status_field:
            messagebox.showwarning("ATENCAO", "Este cadastro nao possui campo STATUS.")
            return
        if not self.selected_id:
            messagebox.showwarning("ATENCAO", "Selecione um cadastro na tabela.")
            return
        status_on = "ATIVO"
        status_off = "INATIVO" if self.table == "motoristas" else "DESATIVADO"
        self._set("status", status_on if liberar else status_off)
        self._edit_mode = "status"
        self.salvar()

    def _toggle_motorista_access(self, liberar: bool):
        if self.table != "motoristas":
            self._set_status_rapido(liberar)
            return
        if not self.definir_acesso_app_motorista(liberar):
            return
        ref = self._selected_motorista_ref()
        if not ref:
            return
        try:
            update_motorista_status_local(int(ref.get("id")), "ATIVO" if liberar else "INATIVO")
        except Exception:
            logging.debug("Falha ao atualizar status local do motorista apos alterar acesso", exc_info=True)
        self.carregar()
        self.limpar()

    def _selected_motorista_ref(self):
        if self.table != "motoristas":
            return None
        sel = self.tree.selection()
        if not sel:
            return None
        vals = list(self.tree.item(sel[0], "values") or [])
        if not vals:
            return None
        cols = [upper(c) for c in (self.tree["columns"] or [])]
        def _val(col_name: str, default=""):
            try:
                idx = cols.index(upper(col_name))
                return str(vals[idx] if idx < len(vals) else default)
            except Exception:
                return default
        return {
            "id": safe_int(_val("ID", "0"), 0),
            "codigo": self._norm(_val("CODIGO")),
            "nome": self._norm(_val("NOME")),
            "cpf": self._norm(_val("CPF")),
            "telefone": self._norm(_val("TELEFONE")),
            "status": self._norm(_val("STATUS", "ATIVO")),
        }

    def _upsert_motorista_acesso_na_api(self, codigo: str, liberar: bool, admin_nome: str, motivo: str, desktop_secret: str, ref=None) -> bool:
        ref = ref or {}
        nome = self._norm(ref.get("nome") or "")
        if not nome and self.selected_id:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT COALESCE(nome,'') FROM motoristas WHERE id=? LIMIT 1", (self.selected_id,))
                    rr = cur.fetchone()
                    nome = self._norm(rr[0] if rr else "")
            except Exception:
                logging.debug("Falha ao buscar nome do motorista para upsert de acesso", exc_info=True)
        if not codigo or not nome:
            return False

        status_ui = upper(self._norm(ref.get("status") or "") or "")
        status_api = "DESATIVADO" if status_ui in {"INATIVO", "DESATIVADO", "BLOQUEADO"} else "ATIVO"

        payload = {
            "codigo": codigo,
            "nome": nome,
            "telefone": self._norm(ref.get("telefone") or self._get("telefone")),
            "cpf": self._norm(ref.get("cpf") or self._get("cpf")),
            "status": status_api,
            "acesso_liberado": bool(liberar),
            "acesso_liberado_por": admin_nome,
            "acesso_obs": motivo or "Sincronizado automaticamente ao alterar acesso",
        }
        senha_ui = self._norm(self._get("senha"))
        if senha_ui and "*" not in senha_ui:
            payload["senha"] = senha_ui
        try:
            _call_api(
                "POST",
                "desktop/cadastros/motoristas/upsert",
                payload=payload,
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            return True
        except Exception:
            logging.exception("Falha ao sincronizar motorista na API antes de definir acesso")
            return False

    def _infer_entry_kind(self, col: str, label: str):
        c = str(col or "").strip().lower()
        l = upper(label or "")

        if any(k in c for k in ("hora",)) or "HORA" in l:
            return "time", 2
        if any(k in c for k in ("data",)) or "DATA" in l:
            return "date", 2
        if any(k in c for k in ("valor", "preco", "custo", "adiantamento", "diaria")):
            return "money", 2
        if "cpf" in c:
            return "cpf", 2
        if "telefone" in c:
            return "phone", 2
        if any(k in c for k in ("capacidade", "qtd", "quantidade", "caixa", "caixas")):
            return "int", 2
        if any(k in c for k in ("kg", "km", "media", "litros")):
            return "decimal", 2
        if c in {"senha", "codigo", "nome", "sobrenome", "status", "veiculo", "motorista", "equipe"}:
            return "text", 2
        if any(k in l for k in ("VALOR", "PREÇO", "CUSTO", "ADIANTAMENTO", "DIÃÂÂRIA")):
            return "money", 2
        if "CPF" in l:
            return "cpf", 2
        if "TELEFONE" in l:
            return "phone", 2
        if any(k in l for k in ("CAPACIDADE", "CAIXA", "QTD")):
            return "int", 2
        if any(k in l for k in ("KG", "KM", "MÉDIA", "MEDIA", "LITROS")):
            return "decimal", 2
        return "text", 2

    def _get(self, key):
        ent = self.entries.get(key)
        return ent.get().strip() if ent else ""

    def _set(self, key, value):
        ent = self.entries.get(key)
        if not ent:
            return
        txt = "" if value is None else str(value)
        try:
            if isinstance(ent, ttk.Combobox):
                ent.set(txt)
            else:
                ent.delete(0, "end")
                ent.insert(0, txt)
        except Exception:
            logging.debug("Falha ignorada")

    def _user_can_change_password(self, user_id):
        if not self.app:
            return False
        if self._is_admin:
            return True
        return str(user_id) == str(self.app.user.get("id"))

    def _update_password_controls(self):
        if self.table not in {"usuarios", "motoristas", "vendedores"}:
            return
        if self.table == "usuarios":
            can_change = self.selected_id and self._user_can_change_password(self.selected_id)
        else:
            can_change = bool(self.selected_id) and self._is_admin
        if hasattr(self, "btn_senha"):
            self.btn_senha.config(state="normal" if can_change else "disabled")
        senha_ent = self.entries.get("senha")
        if senha_ent:
            if self.table == "usuarios":
                if self._is_admin and not self.selected_id:
                    senha_ent.config(state="normal")
                else:
                    senha_ent.config(state="disabled")
            else:
                if self._is_admin and not self.selected_id:
                    senha_ent.config(state="normal")
                else:
                    senha_ent.config(state="disabled")

    def _require_fields(self, fields_required):
        for f in fields_required:
            if not str(self._get(f)).strip():
                return False, f
        return True, ""

    def _dup_exists(self, cur, colname, value_norm, ignore_id=None):
        """
        Verifica duplicidade case-insensitive:
        SELECT id FROM table WHERE UPPER(col)=UPPER(?) AND id<>?
        """
        if not value_norm:
            return False
        try:
            if ignore_id:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE UPPER({colname})=UPPER(?) AND id<>? LIMIT 1",
                    (value_norm, ignore_id)
                )
            else:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE UPPER({colname})=UPPER(?) LIMIT 1",
                    (value_norm,)
                )
            return cur.fetchone() is not None
        except Exception:
            # Se por algum motivo a coluna não existir no DB, não trava o sistema.
            return False

    def _dup_exists_exact(self, cur, colname, value, ignore_id=None):
        """
        Verifica duplicidade exata (útil para CPF/telefone quando você normaliza para dÃÂÂgitos).
        """
        if not value:
            return False
        try:
            if ignore_id:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE {colname}=? AND id<>? LIMIT 1",
                    (value, ignore_id)
                )
            else:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE {colname}=? LIMIT 1",
                    (value,)
                )
            return cur.fetchone() is not None
        except Exception:
            return False

    def _sync_motorista_upsert_api(self, data: dict):
        if self.table != "motoristas":
            return
        result = sync_motorista_upsert_api(
            data,
            is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            norm=self._norm,
        )
        if isinstance(result, dict) and not bool(result.get("ok", False)):
            raise RuntimeError(result.get("error") or "Falha ao sincronizar motorista.")

    def _sync_veiculo_upsert_api(self, data: dict):
        if self.table != "veiculos":
            return
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        placa = self._norm(data.get("placa"))
        modelo = self._norm(data.get("modelo"))
        if not placa or not modelo:
            return
        payload = {
            "placa": placa,
            "modelo": modelo,
            "capacidade_cx": int(safe_int(data.get("capacidade_cx"), 0)),
        }
        _call_api(
            "POST",
            "desktop/cadastros/veiculos/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def _sync_cliente_upsert_api(self, cod, nome, endereco, telefone, vendedor):
        if self.table != "clientes":
            return
        result = sync_cliente_upsert_api(
            cod,
            nome,
            endereco,
            telefone,
            vendedor,
            is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
        )
        if isinstance(result, dict) and not bool(result.get("ok", False)):
            raise RuntimeError(result.get("error") or "Falha ao sincronizar cliente.")

    def _sync_ajudante_upsert_api(self, data: dict):
        if self.table != "ajudantes":
            return
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        nome = self._norm(data.get("nome"))
        sobrenome = self._norm(data.get("sobrenome"))
        if not nome or not sobrenome:
            return
        payload = {
            "nome": nome,
            "sobrenome": sobrenome,
            "telefone": self._norm(data.get("telefone")),
            "status": self._norm(data.get("status") or "ATIVO"),
        }
        _call_api(
            "POST",
            "desktop/cadastros/ajudantes/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def _sync_vendedor_upsert_api(self, data: dict):
        if self.table != "vendedores":
            return
        result = sync_vendedor_upsert_api(
            data,
            is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
            norm=self._norm,
            normalize_phone=normalize_phone,
        )
        if isinstance(result, dict) and not bool(result.get("ok", False)):
            raise RuntimeError(result.get("error") or "Falha ao sincronizar vendedor.")

    def _has_delete_references(self, cur) -> str:
        if not self.selected_id:
            return ""
        try:
            if self.table == "motoristas":
                cod = self._norm(self._get("codigo"))
                nome = self._norm(self._get("nome"))
                if not cod and not nome:
                    return ""
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                wh = []
                params = []
                if "motorista_codigo" in cols and cod:
                    wh.append("UPPER(COALESCE(motorista_codigo,''))=UPPER(?)")
                    params.append(cod)
                if "codigo_motorista" in cols and cod:
                    wh.append("UPPER(COALESCE(codigo_motorista,''))=UPPER(?)")
                    params.append(cod)
                if "motorista" in cols:
                    if cod:
                        wh.append("UPPER(COALESCE(motorista,''))=UPPER(?)")
                        params.append(cod)
                    if nome:
                        wh.append("UPPER(COALESCE(motorista,''))=UPPER(?)")
                        params.append(nome)
                if wh:
                    cur.execute(
                        "SELECT COUNT(*) FROM programacoes WHERE " + " OR ".join(wh),
                        tuple(params),
                    )
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Motorista vinculado a programação/rota."
            elif self.table == "vendedores":
                cod = self._norm(self._get("codigo"))
                nome = self._norm(self._get("nome"))
                if cod and db_has_column(cur, "clientes", "vendedor"):
                    cur.execute("SELECT COUNT(*) FROM clientes WHERE UPPER(COALESCE(vendedor,''))=UPPER(?)", (cod,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Vendedor vinculado ao cadastro de clientes."
                if nome and db_has_column(cur, "clientes", "vendedor"):
                    cur.execute("SELECT COUNT(*) FROM clientes WHERE UPPER(COALESCE(vendedor,''))=UPPER(?)", (nome,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Vendedor vinculado ao cadastro de clientes."
                if cod and table_has_column(cur, "programacoes", "usuario_criacao"):
                    cur.execute("SELECT COUNT(*) FROM programacoes WHERE UPPER(COALESCE(usuario_criacao,''))=UPPER(?)", (cod,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Vendedor vinculado a programação/rota."
            elif self.table == "veiculos":
                placa = self._norm(self._get("placa"))
                if placa:
                    cur.execute("SELECT COUNT(*) FROM programacoes WHERE UPPER(COALESCE(veiculo,''))=UPPER(?)", (placa,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Veiculo vinculado a programacao/rota."
            elif self.table == "clientes":
                cod = self._norm(self._get("cod_cliente"))
                if cod:
                    cur.execute("SELECT COUNT(*) FROM programacao_itens WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)", (cod,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Cliente vinculado a programação."
                    try:
                        cur.execute("SELECT COUNT(*) FROM recebimentos WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)", (cod,))
                        if int((cur.fetchone() or [0])[0] or 0) > 0:
                            return "Cliente vinculado a recebimentos."
                    except Exception:
                        logging.debug("Falha ignorada")
            elif self.table == "ajudantes":
                aid = str(self.selected_id)
                nome = self._norm(self._get("nome"))
                sobrenome = self._norm(self._get("sobrenome"))
                alvo = f"{nome} {sobrenome}".strip()
                cur.execute("PRAGMA table_info(equipes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                conds = []
                params = []
                for c in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2"):
                    if c in cols:
                        conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                        params.append(aid)
                        if nome:
                            conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                            params.append(nome)
                        if alvo:
                            conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                            params.append(alvo)
                if conds:
                    cur.execute("SELECT COUNT(*) FROM equipes WHERE " + " OR ".join(conds), tuple(params))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Ajudante vinculado a equipe."
        except Exception:
            logging.debug("Falha na verificação de vÃÂnculos para exclusão", exc_info=True)
        return ""

    # -------------------------
    # Ações padrão
    # -------------------------
    def limpar(self):
        self.selected_id = None
        for col, _ in self.fields:
            self._set(col, "")
        if self.table == "ajudantes" and "status" in self.entries:
            self._set("status", "ATIVO")
        if self.table == "motoristas" and "status" in self.entries:
            self._set("status", "ATIVO")
            if "codigo" in self.entries:
                self._set("codigo", self._next_motorista_codigo())
        if self.table == "vendedores" and "status" in self.entries:
            self._set("status", "ATIVO")
        self._set_form_mode("view")
        self._update_password_controls()

    def alterar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENCAO", "Selecione um item na tabela para alterar.")
            return
        self._on_select()
        if not self._has_status_field:
            messagebox.showwarning("ATENCAO", "Este cadastro nao possui campo STATUS.")
            return
        self._set_form_mode("status")
        try:
            ent = self.entries.get("status")
            if ent:
                ent.focus_set()
        except Exception:
            logging.debug("Falha ignorada")

    def _fetch_cadastro_rows(self, status_filter="TODOS"):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and can_read_from_api():
            try:
                api_rows = None
                if self.table == "motoristas":
                    res_motoristas = fetch_motoristas_rows(fields=self.fields, can_read_from_api=can_read_from_api)
                    if isinstance(res_motoristas, dict):
                        if bool(res_motoristas.get("ok", False)) and isinstance(res_motoristas.get("data"), list):
                            return res_motoristas.get("data")
                        if res_motoristas.get("error"):
                            logging.debug(
                                "Falha ao carregar motoristas via service; usando fallback local: %s",
                                res_motoristas.get("error"),
                            )
                    elif res_motoristas is not None:
                        return res_motoristas
                elif self.table == "vendedores":
                    res_vendedores = fetch_vendedores_rows(fields=self.fields, can_read_from_api=can_read_from_api)
                    if isinstance(res_vendedores, dict):
                        if bool(res_vendedores.get("ok", False)) and isinstance(res_vendedores.get("data"), list):
                            return res_vendedores.get("data")
                        if res_vendedores.get("error"):
                            logging.debug(
                                "Falha ao carregar vendedores via service; usando fallback local: %s",
                                res_vendedores.get("error"),
                            )
                    elif res_vendedores is not None:
                        return res_vendedores
                elif self.table == "veiculos":
                    api_rows = _call_api(
                        "GET",
                        "desktop/cadastros/veiculos",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                elif self.table == "ajudantes":
                    api_rows = _call_api(
                        "GET",
                        "desktop/cadastros/ajudantes",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )

                if isinstance(api_rows, list):
                    rows = []
                    for r in api_rows:
                        if not isinstance(r, dict):
                            continue
                        row_map = {}
                        if self.table == "veiculos":
                            row_map = {
                                "placa": str(r.get("placa") or ""),
                                "modelo": str(r.get("modelo") or ""),
                                "capacidade_cx": str(r.get("capacidade_cx") or 0),
                            }
                        elif self.table == "ajudantes":
                            row_map = {
                                "nome": str(r.get("nome_base") or r.get("nome") or ""),
                                "sobrenome": str(r.get("sobrenome") or ""),
                                "telefone": str(r.get("telefone") or ""),
                                "status": str(r.get("status") or "ATIVO"),
                            }
                        row = [int(safe_int(r.get("id"), 0))]
                        for c, _ in self.fields:
                            row.append(row_map.get(c, ""))
                        rows.append(tuple(row))
                    return rows
            except Exception:
                logging.debug("Falha ao carregar cadastro via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute(f"PRAGMA table_info({self.table})")
                cols_exist = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                cols_exist = set()

            select_cols = []
            for c, _ in self.fields:
                if str(c).lower() in cols_exist:
                    select_cols.append(c)
                else:
                    select_cols.append(f"'' AS {c}")
            cols_db = ", ".join(select_cols)
            if self.table == "ajudantes" and status_filter in {"ATIVO", "DESATIVADO"}:
                cur.execute(
                    f"SELECT id, {cols_db} FROM {self.table} "
                    "WHERE UPPER(COALESCE(status, 'ATIVO'))=? ORDER BY id DESC",
                    (status_filter,),
                )
            else:
                cur.execute(f"SELECT id, {cols_db} FROM {self.table} ORDER BY id DESC")
            return cur.fetchall() or []

    def _apply_cadastro_rows(self, seq, rows):
        if seq != self._load_seq or not self.winfo_exists():
            return
        self.tree.delete(*self.tree.get_children())
        senha_pos = None
        if self.table in {"usuarios", "motoristas", "vendedores"}:
            try:
                senha_pos = 1 + [c for c, _ in self.fields].index("senha")
            except Exception:
                senha_pos = None
        for row in rows:
            if senha_pos is not None:
                row = list(row)
                row[senha_pos] = "******" if row[senha_pos] else ""
                row = tuple(row)
            tree_insert_aligned(self.tree, "", "end", row)

    def _handle_cadastro_load_error(self, seq, exc):
        if seq != self._load_seq or not self.winfo_exists():
            return
        logging.error("Falha ao carregar cadastro %s", self.table, exc_info=(type(exc), exc, exc.__traceback__))

    def carregar(self, async_load=True):
        status_filter = "TODOS"
        if self.table == "ajudantes":
            try:
                status_filter = upper(self._ajudantes_status_filter.get())
            except Exception:
                status_filter = "TODOS"
        self._load_seq += 1
        seq = self._load_seq
        if not async_load:
            try:
                rows = self._fetch_cadastro_rows(status_filter=status_filter)
            except Exception as exc:
                self._handle_cadastro_load_error(seq, exc)
                return
            self._apply_cadastro_rows(seq, rows)
            return

        run_async_ui(
            self,
            lambda status_filter=status_filter: self._fetch_cadastro_rows(status_filter=status_filter),
            lambda rows, seq=seq: self._apply_cadastro_rows(seq, rows),
            lambda exc, seq=seq: self._handle_cadastro_load_error(seq, exc),
        )

    def salvar(self):
        """
        Salva/atualiza com validações por cadastro:
        - motoristas: valida NOME/CPF/TELEFONE; exige codigo e senha (manual); codigo único; CPF único (se houver coluna)
        - usuarios: exige nome, senha; nome único; sem codigo
        - veiculos: exige placa, modelo, capacidade_cx; placa única; capacidade int >= 0

        Obs: ID é a chave primária autogerada (sequência do SQLite). Você já vê o ID na tabela.
        """
        if not ensure_system_api_binding(context=f"Salvar cadastro ({self.table})", parent=self):
            return

        if self._edit_mode not in {"novo", "status"}:
            messagebox.showwarning("ATENCAO", "Clique em NOVO ou ALTERAR antes de salvar.")
            return

        data = {col: self.entries[col].get().strip() for col, _ in self.fields}

        # Normalizações por tabela
        if self.table in {"motoristas", "vendedores", "usuarios", "veiculos", "equipes", "ajudantes", "clientes"}:
            for k in list(data.keys()):
                data[k] = self._norm(data[k])

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_sync_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())

        if self._edit_mode == "status":
            if not self.selected_id:
                messagebox.showwarning("ATENCAO", "Selecione um cadastro para alterar o status.")
                return
            if "status" not in data:
                messagebox.showwarning("ATENCAO", "Este cadastro nao possui campo STATUS.")
                return
            novo_status = self._norm(data.get("status"))
            if self.table == "motoristas":
                if novo_status not in {"ATIVO", "INATIVO"}:
                    messagebox.showwarning("ATENCAO", "Status invalido. Use ATIVO ou INATIVO.")
                    return
                data["status"] = novo_status
            elif self.table == "vendedores":
                if novo_status not in {"ATIVO", "DESATIVADO"}:
                    messagebox.showwarning("ATENCAO", "Status invalido. Use ATIVO ou DESATIVADO.")
                    return
                data["status"] = novo_status
            elif self.table == "ajudantes":
                if novo_status not in {"ATIVO", "DESATIVADO"}:
                    messagebox.showwarning("ATENCAO", "Status invalido. Use ATIVO ou DESATIVADO.")
                    return
                data["status"] = novo_status

            if api_sync_enabled and self.table in {"motoristas", "vendedores", "ajudantes"}:
                try:
                    if self.table == "motoristas":
                        self._sync_motorista_upsert_api(data)
                    elif self.table == "vendedores":
                        self._sync_vendedor_upsert_api(data)
                    elif self.table == "ajudantes":
                        self._sync_ajudante_upsert_api(data)
                except Exception as e:
                    messagebox.showerror(
                        "ERRO",
                        "Falha ao salvar status na API central.\n"
                        "Nenhuma alteração local foi aplicada.\n\n"
                        f"Detalhe: {e}",
                    )
                    return

            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute(f"UPDATE {self.table} SET status=? WHERE id=?", (data["status"], self.selected_id))
            except Exception as e:
                messagebox.showerror("ERRO", str(e))
                return

            self.carregar()
            self.limpar()
            if self.app:
                self.app.refresh_programacao_comboboxes()
            return

        saved_via_api = False
        # Validar obrigatórios + duplicidade + tipos
        try:
            with get_db() as conn:
                cur = conn.cursor()

                # -------------------------
                # MOTORISTAS
                # -------------------------
                if self.table == "motoristas":
                    # Exige NOME, CÓDIGO e SENHA (manuais) + TELEFONE
                    # CPF é opcional, porém deve ser válido e único quando informado.
                    req = ["nome", "senha", "telefone"] if not self.selected_id else ["nome", "telefone"]
                    ok, f = self._require_fields(req)
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    # Nome (básico)
                    nome = self._norm(data.get("nome"))
                    if len(nome) < 3:
                        messagebox.showwarning("ATENÇÃO", "NOME deve ter pelo menos 3 caracteres.")
                        return
                    data["nome"] = nome

                    # Código (login flutter)  manual, mas pode usar GERAR
                    cod = self._norm(data.get("codigo"))
                    if not self.selected_id:
                        cod = self._next_motorista_codigo(cur)
                    if not is_valid_motorista_codigo(cod):
                        messagebox.showwarning(
                            "ATENÇÃO",
                            "CÓDIGO inválido. Use apenas letras/números/._- e 3 a 24 caracteres."
                        )
                        return
                    if self._dup_exists(cur, "codigo", cod, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE MOTORISTA COM ESTE CÓDIGO: {cod}")
                        return
                    data["codigo"] = cod

                    # Senha (login flutter)  manual
                    senha = (self.entries.get("senha").get().strip() if self.entries.get("senha") else "").strip()
                    if self.selected_id:
                        if senha:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha do motorista.")
                                return
                            if not is_valid_motorista_senha(senha):
                                messagebox.showwarning(
                                    "ATENÇÃO",
                                    "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                                )
                                return
                            data["senha"] = hash_password_pbkdf2(senha)
                        else:
                            data.pop("senha", None)
                    else:
                        if not is_valid_motorista_senha(senha):
                            messagebox.showwarning(
                                "ATENÇÃO",
                                "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                            )
                            return
                        data["senha"] = hash_password_pbkdf2(senha)

                    # CPF (opcional)
                    cpf_raw = self.entries.get("cpf").get().strip() if self.entries.get("cpf") else ""
                    cpf = normalize_cpf(cpf_raw)
                    if cpf and not is_valid_cpf(cpf):
                        messagebox.showwarning("ATENÇÃO", "CPF inválido.")
                        return
                    data["cpf"] = cpf  # salva só dÃÂÂgitos

                    # Telefone
                    tel_raw = self.entries.get("telefone").get().strip() if self.entries.get("telefone") else ""
                    tel = normalize_phone(tel_raw)
                    if not is_valid_phone(tel):
                        messagebox.showwarning(
                            "ATENÇÃO",
                            "TELEFONE inválido. Informe DDD+Número (10 ou 11 dÃÂÂgitos)."
                        )
                        return
                    data["telefone"] = tel  # salva só dÃÂÂgitos

                    # Status do motorista
                    status_m = self._norm(data.get("status") or "ATIVO")
                    if status_m not in {"ATIVO", "INATIVO"}:
                        status_m = "ATIVO"
                    data["status"] = status_m

                    # CPF único quando informado (se a coluna existir)
                    if cpf and db_has_column(cur, "motoristas", "cpf"):
                        if self._dup_exists_exact(cur, "cpf", cpf, ignore_id=self.selected_id):
                            messagebox.showerror("ERRO", "J EXISTE MOTORISTA COM ESTE CPF.")
                            return

                    # Fluxo natural: cadastro novo de motorista já nasce apto ao app mobile.
                    # Mantém compatibilidade com bases que ainda não possuem essas colunas.
                    if not self.selected_id:
                        try:
                            cur.execute("PRAGMA table_info(motoristas)")
                            cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                        except Exception:
                            cols_m = set()
                        if "acesso_liberado" in cols_m and "acesso_liberado" not in data:
                            data["acesso_liberado"] = 1
                        if "acesso_liberado_por" in cols_m and not data.get("acesso_liberado_por"):
                            data["acesso_liberado_por"] = "AUTO_DESKTOP"
                        if "acesso_liberado_em" in cols_m and not data.get("acesso_liberado_em"):
                            data["acesso_liberado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        if "acesso_obs" in cols_m and not data.get("acesso_obs"):
                            data["acesso_obs"] = "Liberacao automatica no cadastro desktop"

                # -------------------------
                # USURIOS (login por nome + senha)
                # -------------------------
                if self.table == "usuarios":
                    nome = self._norm(data.get("nome"))
                    senha_plana =""
                    try:
                        senha_plana = (self.entries.get("senha").get() or "").strip()
                    except Exception:
                        senha_plana = str(data.get("senha") or "").strip()

                    if not nome:
                        messagebox.showwarning("ANTENÇÃO", "PREENCHA O CAMPO: NOME.")
                        return
                    if self._dup_exists(cur, "nome", nome, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE USURIO COM ESTE NOME:{nome}")
                        return

                    data["nome"] = nome
                    if self.selected_id:
                        if senha_plana:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha aqui.")
                                return
                            if len(senha_plana) < 6:
                                messagebox.showwarning("ANTENÇÃO", "SENHA deve ter pelo menos 6 caracteres.")
                                return
                            data["senha"] = hash_password_pbkdf2(senha_plana)
                        else:
                            data.pop("senha", None)
                    else:
                        if not senha_plana:
                            messagebox.showwarning("ANTENÇÃO", "PREENCHA O CAMPO: SENHA.")
                            return
                        if len(senha_plana) < 6:
                            messagebox.showwarning("ANTENÇÃO", "SENHA deve ter pelo menos 6 caracteres.")
                            return
                        data["senha"] = hash_password_pbkdf2(senha_plana)

                # -------------------------
                # VENDEDORES (login do app vendedor)
                # -------------------------
                if self.table == "vendedores":
                    req = ["codigo", "nome", "status"] if self.selected_id else ["codigo", "nome", "senha", "status"]
                    ok, f = self._require_fields(req)
                    if not ok:
                        messagebox.showwarning("ATENÃ‡ÃƒO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    codigo = self._norm(data.get("codigo"))
                    nome = self._norm(data.get("nome"))
                    cidade_base = self._norm(data.get("cidade_base"))
                    telefone = normalize_phone(data.get("telefone"))
                    status_v = self._norm(data.get("status") or "ATIVO")

                    if not is_valid_motorista_codigo(codigo):
                        messagebox.showwarning(
                            "ATENÃ‡ÃƒO",
                            "CODIGO invÃ¡lido. Use apenas letras/nÃºmeros/._- e 3 a 24 caracteres."
                        )
                        return
                    if len(nome) < 3:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "NOME deve ter pelo menos 3 caracteres.")
                        return
                    if telefone and not is_valid_phone(telefone):
                        messagebox.showwarning(
                            "ATENÃ‡ÃƒO",
                            "TELEFONE invÃ¡lido. Informe DDD+NÃºmero (10 ou 11 dÃ­gitos) ou deixe vazio."
                        )
                        return
                    if status_v not in {"ATIVO", "DESATIVADO"}:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "STATUS invÃ¡lido para vendedor.")
                        return
                    if self._dup_exists(cur, "codigo", codigo, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"JÃ EXISTE VENDEDOR COM ESTE CÃ“DIGO: {codigo}")
                        return

                    senha_plana = (self.entries.get("senha").get().strip() if self.entries.get("senha") else "").strip()
                    if self.selected_id:
                        if senha_plana:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÃ‡ÃƒO", "Somente ADMIN pode alterar senha do vendedor.")
                                return
                            if not is_valid_motorista_senha(senha_plana):
                                messagebox.showwarning(
                                    "ATENÃ‡ÃƒO",
                                    "SENHA invÃ¡lida. Use 4 a 24 caracteres."
                                )
                                return
                            data["senha"] = hash_password_pbkdf2(senha_plana)
                        else:
                            data.pop("senha", None)
                    else:
                        if not is_valid_motorista_senha(senha_plana):
                            messagebox.showwarning(
                                "ATENÃ‡ÃƒO",
                                "SENHA invÃ¡lida. Use 4 a 24 caracteres."
                            )
                            return
                        data["senha"] = hash_password_pbkdf2(senha_plana)

                    data["codigo"] = codigo
                    data["nome"] = nome
                    data["telefone"] = telefone
                    data["cidade_base"] = cidade_base
                    data["status"] = status_v

                # -------------------------
                # AJUDANTES (nome, sobrenome, telefone)
                # -------------------------
                if self.table == "ajudantes":
                    ok, f = self._require_fields(["nome", "sobrenome", "telefone", "status"])
                    if not ok:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    nome = self._norm(data.get("nome"))
                    sobrenome = self._norm(data.get("sobrenome"))
                    telefone = normalize_phone(data.get("telefone"))
                    status = self._norm(data.get("status"))

                    if len(nome) < 2:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", "NOME deve ter pelo menos 2 caracteres.")
                        return
                    if len(sobrenome) < 2:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", "SOBRENOME deve ter pelo menos 2 caracteres.")
                        return
                    if not is_valid_phone(telefone):
                        messagebox.showwarning(
                            "ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O",
                            "TELEFONE inválido. Informe DDD+Número (10 ou 11 dÃÂÂÂgitos)."
                        )
                        return
                    if status not in {"ATIVO", "DESATIVADO"}:
                        messagebox.showwarning("ATENÇÃO", "STATUS inválido para ajudante.")
                        return

                    # Evita duplicidade de pessoa com mesmo nome completo.
                    try:
                        if self.selected_id:
                            cur.execute(
                                """
                                SELECT id FROM ajudantes
                                WHERE UPPER(nome)=UPPER(?) AND UPPER(sobrenome)=UPPER(?) AND id<>?
                                LIMIT 1
                                """,
                                (nome, sobrenome, self.selected_id),
                            )
                        else:
                            cur.execute(
                                """
                                SELECT id FROM ajudantes
                                WHERE UPPER(nome)=UPPER(?) AND UPPER(sobrenome)=UPPER(?)
                                LIMIT 1
                                """,
                                (nome, sobrenome),
                            )
                        if cur.fetchone():
                            messagebox.showerror("ERRO", "JÃ EXISTE AJUDANTE COM ESTE NOME/SOBRENOME.")
                            return
                    except Exception:
                        logging.debug("Falha ignorada")

                    data["nome"] = nome
                    data["sobrenome"] = sobrenome
                    data["telefone"] = telefone
                    data["status"] = status

                # -------------------------
                # VECULOS (placa, modelo, capacidade_cx)
                # -------------------------
                if self.table == "veiculos":
                    ok, f = self._require_fields(["placa", "modelo", "capacidade_cx"])
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    placa = self._norm(data.get("placa"))
                    if self._dup_exists(cur, "placa", placa, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE VECULO COM ESTA PLACA: {placa}")
                        return

                    cap = safe_int(data.get("capacidade_cx"), -1)
                    if cap < 0:
                        messagebox.showwarning("ATENÇÃO", "CAPACIDADE (CX) deve ser um número inteiro >= 0.")
                        return
                    data["capacidade_cx"] = cap

                # -------------------------
                # SALVAR (UPDATE/INSERT)
                # -------------------------
                if self.table in {"motoristas", "vendedores", "veiculos", "ajudantes"} and api_sync_enabled:
                    try:
                        if self.table == "motoristas":
                            self._sync_motorista_upsert_api(data)
                        elif self.table == "vendedores":
                            self._sync_vendedor_upsert_api(data)
                        elif self.table == "veiculos":
                            self._sync_veiculo_upsert_api(data)
                        elif self.table == "ajudantes":
                            self._sync_ajudante_upsert_api(data)
                        saved_via_api = True
                    except Exception as e:
                        raise RuntimeError(
                            "Falha ao salvar cadastro na API central. "
                            "Nenhuma alteração local foi aplicada."
                        ) from e

                if not saved_via_api:
                    if self.selected_id:
                        sets = ", ".join([f"{c}=?" for c in data.keys()])
                        values = list(data.values()) + [self.selected_id]
                        cur.execute(f"UPDATE {self.table} SET {sets} WHERE id=?", values)
                    else:
                        cols = ", ".join(data.keys())
                        qs = ", ".join(["?"] * len(data))
                        cur.execute(f"INSERT INTO {self.table} ({cols}) VALUES ({qs})", list(data.values()))
                    
                    # Se for um novo usuário, atribuir permissões do perfil
                    if self.table == "usuarios" and not self.selected_id:
                        try:
                            from app.services.permissions_service import atribuir_permissoes_por_perfil
                            cur.execute("SELECT id FROM usuarios WHERE nome=?", (data.get("nome"),))
                            novo_usuario = cur.fetchone()
                            if novo_usuario:
                                usuario_id = novo_usuario[0]
                                perfil = data.get("permissoes", "OPERADOR")
                                usuario_admin = getattr(self.app.user if self.app else {}, "get", lambda x: "ADMIN")("nome") or "ADMIN"
                                atribuir_permissoes_por_perfil(usuario_id, perfil, usuario_admin)
                        except Exception as e:
                            logging.warning(f"Erro ao atribuir permissões ao novo usuário: {e}")
                    
                    # Se for atualização de usuário, atualizar permissões do perfil
                    elif self.table == "usuarios" and self.selected_id:
                        try:
                            from app.services.permissions_service import atribuir_permissoes_por_perfil
                            perfil = data.get("permissoes", "OPERADOR")
                            usuario_admin = getattr(self.app.user if self.app else {}, "get", lambda x: "ADMIN")("nome") or "ADMIN"
                            atribuir_permissoes_por_perfil(self.selected_id, perfil, usuario_admin)
                        except Exception as e:
                            logging.warning(f"Erro ao atualizar permissões do usuário: {e}")

        except sqlite3.IntegrityError as e:
            messagebox.showerror("ERRO", f"REGISTRO DUPLICADO OU INVLIDO.\n\n{e}")
            return
        except Exception as e:
            messagebox.showerror("ERRO", str(e))
            return

        if (not saved_via_api) and self.table == "motoristas":
            try:
                self._sync_motorista_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "Motorista salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )
        elif (not saved_via_api) and self.table == "vendedores":
            try:
                self._sync_vendedor_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "SincronizaÃ§Ã£o",
                    "Vendedor salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )
        elif (not saved_via_api) and self.table == "veiculos":
            try:
                self._sync_veiculo_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "Veiculo salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )
        elif (not saved_via_api) and self.table == "ajudantes":
            try:
                self._sync_ajudante_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "Ajudante salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def excluir(self):
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM ITEM NA TABELA ANTES DE EXCLUIR.")
            return

        if not messagebox.askyesno("CONFIRMAR", "DESEJA EXCLUIR ESTE REGISTRO?"):
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                endpoint = ""
                if self.table == "motoristas":
                    cod = self._norm(self._get("codigo"))
                    if cod:
                        endpoint = f"desktop/cadastros/motoristas/{urllib.parse.quote(cod)}"
                elif self.table == "vendedores":
                    cod = self._norm(self._get("codigo"))
                    if cod:
                        endpoint = f"desktop/cadastros/vendedores/{urllib.parse.quote(cod)}"
                elif self.table == "veiculos":
                    placa = self._norm(self._get("placa"))
                    if placa:
                        endpoint = f"desktop/cadastros/veiculos/{urllib.parse.quote(placa)}"
                elif self.table == "ajudantes":
                    endpoint = f"desktop/cadastros/ajudantes/{int(safe_int(self.selected_id, 0))}"
                elif self.table == "clientes":
                    cod_cli = self._norm(self._get("cod_cliente"))
                    if cod_cli:
                        endpoint = f"desktop/cadastros/clientes/{urllib.parse.quote(cod_cli)}"

                if endpoint:
                    _call_api(
                        "DELETE",
                        endpoint,
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    self.carregar()
                    self.limpar()
                    if self.app:
                        self.app.refresh_programacao_comboboxes()
                    return
            except SyncError as e:
                msg = str(e)
                if msg.startswith("409 "):
                    messagebox.showwarning("ATENÇÃO", msg.split(":", 1)[-1].strip() or msg)
                else:
                    messagebox.showerror(
                        "ERRO",
                        "Nao foi possivel excluir o cadastro na API central.\n\n"
                        "A exclusao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                        f"Detalhe: {msg}",
                    )
                return
            except Exception:
                logging.debug("Falha ao excluir cadastro via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel excluir o cadastro na API central.\n\n"
                    "A exclusao local foi bloqueada para evitar divergencia entre servidor e desktop.",
                )
                return

        with get_db() as conn:
            cur = conn.cursor()
            bloqueio = self._has_delete_references(cur)
            if bloqueio:
                messagebox.showwarning("ATENÇÃO", f"Exclusão bloqueada.\n\n{bloqueio}\nUse status DESATIVADO.")
                return
            cur.execute(f"DELETE FROM {self.table} WHERE id=?", (self.selected_id,))

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def editar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENÇÃO", "Selecione um item na tabela para editar.")
            return
        self._on_select()
        try:
            first_col = self.fields[0][0] if self.fields else ""
            ent = self.entries.get(first_col)
            if ent:
                ent.focus_set()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.tree.item(sel[0], "values")
        self.selected_id = row[0]
        for i, (col, _) in enumerate(self.fields, start=1):
            if self.table in {"usuarios", "motoristas", "vendedores"} and col == "senha":
                self._set(col, "")
                continue
            self._set(col, row[i] if i < len(row) else "")
        self._set_form_mode("view")
        self._update_password_controls()

        # Mantive seu importar_clientes_excel como estava (se você ainda usa em algum ponto)
    def _importar_clientes_worker(self, path):
        df = pd.read_excel(path, engine=excel_engine_for(path))

        col_cod = guess_col(df.columns, ["cod", "cÃ³d", "codigo", "cliente", "cod cliente"])
        col_nome = guess_col(df.columns, ["nome", "cliente"])
        col_end = guess_col(df.columns, ["endereco", "endereÃ§o", "rua", "logradouro"])
        col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
        col_vendedor = guess_col(df.columns, ["vendedor", "vend", "representante"])

        if not col_cod or not col_nome:
            cols = list(df.columns or [])
            if len(cols) >= 2:
                col_cod = col_cod or cols[0]
                col_nome = col_nome or cols[1]
            else:
                raise ValueError("NÃƒO IDENTIFIQUEI AS COLUNAS DE CÃ“DIGO E NOME DO CLIENTE NO EXCEL.")

        total = 0
        sync_falhas = 0
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_mode = bool(desktop_secret and is_desktop_api_sync_enabled())
        if api_mode:
            for _, r in df.iterrows():
                cod = str(r.get(col_cod, "")).strip()
                nome = str(r.get(col_nome, "")).strip()
                if not cod or not nome:
                    continue

                endereco = str(r.get(col_end, "")).strip() if col_end else ""
                telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                vendedor = str(r.get(col_vendedor, "")).strip() if col_vendedor else ""
                try:
                    self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                    total += 1
                except Exception:
                    sync_falhas += 1
        else:
            with get_db() as conn:
                cur = conn.cursor()
                for _, r in df.iterrows():
                    cod = str(r.get(col_cod, "")).strip()
                    nome = str(r.get(col_nome, "")).strip()
                    if not cod or not nome:
                        continue

                    endereco = str(r.get(col_end, "")).strip() if col_end else ""
                    telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                    vendedor = str(r.get(col_vendedor, "")).strip() if col_vendedor else ""

                    cur.execute("""
                        INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                        VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(cod_cliente) DO UPDATE SET
                            nome_cliente=excluded.nome_cliente,
                            endereco=COALESCE(NULLIF(excluded.endereco,''), clientes.endereco),
                            telefone=COALESCE(NULLIF(excluded.telefone,''), clientes.telefone),
                            vendedor=COALESCE(NULLIF(excluded.vendedor,''), clientes.vendedor)
                    """, (upper(cod), upper(nome), upper(endereco), upper(telefone), upper(vendedor)))
                    total += 1
                    try:
                        self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                    except Exception:
                        sync_falhas += 1

        msg = f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}"
        if sync_falhas:
            msg += f"\nFalhas de sincronizaÃ§Ã£o API: {sync_falhas}"
        return msg

    def _on_importar_clientes_done(self, seq, msg):
        if seq != self._import_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Importação concluída")
        messagebox.showinfo("OK", msg)
        self.carregar()

    def _on_importar_clientes_error(self, seq, exc):
        if seq != self._import_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Falha na importação")
        logging.error("Falha ao importar clientes via Excel", exc_info=(type(exc), exc, exc.__traceback__))
        messagebox.showerror("ERRO", str(exc))

    def importar_clientes_excel(self):
        if not ensure_system_api_binding(context="Importar clientes (cadastros)", parent=self):
            return
        path = filedialog.askopenfilename(
            title="IMPORTAR CLIENTES (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return None

        if not (require_pandas() and require_excel_support(path)):
            return

        try:
            df = pd.read_excel(path, engine=excel_engine_for(path))

            col_cod = guess_col(df.columns, ["cod", "cód", "codigo", "cliente", "cod cliente"])
            col_nome = guess_col(df.columns, ["nome", "cliente"])
            col_end = guess_col(df.columns, ["endereco", "endereço", "rua", "logradouro"])
            col_bairro = guess_col(df.columns, ["bairro"])
            col_cidade = guess_col(df.columns, ["cidade"])
            col_uf = guess_col(df.columns, ["uf", "estado"])
            col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
            col_rota = guess_col(df.columns, ["rota"])

            if not col_cod or not col_nome:
                cols = list(df.columns or [])
                if len(cols) >= 2:
                    col_cod = col_cod or cols[0]
                    col_nome = col_nome or cols[1]
                else:
                    messagebox.showerror("ERRO", "NÃO IDENTIFIQUEI AS COLUNAS DE CÓDIGO E NOME DO CLIENTE NO EXCEL.")
                    return

            total = 0
            with get_db() as conn:
                cur = conn.cursor()
                for _, r in df.iterrows():
                    cod = str(r.get(col_cod, "")).strip()
                    nome = str(r.get(col_nome, "")).strip()
                    if not cod or not nome:
                        continue

                    endereco = str(r.get(col_end, "")).strip() if col_end else ""
                    bairro = str(r.get(col_bairro, "")).strip() if col_bairro else ""
                    cidade = str(r.get(col_cidade, "")).strip() if col_cidade else ""
                    uf = str(r.get(col_uf, "")).strip() if col_uf else ""
                    telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                    rota = str(r.get(col_rota, "")).strip() if col_rota else ""

                    cur.execute("""
                        INSERT INTO clientes (cod_cliente, nome_cliente, endereco, bairro, cidade, uf, telefone, rota)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(cod_cliente) DO UPDATE SET
                            nome_cliente=excluded.nome_cliente,
                            endereco=COALESCE(NULLIF(excluded.endereco,''), clientes.endereco),
                            bairro=COALESCE(NULLIF(excluded.bairro,''), clientes.bairro),
                            cidade=COALESCE(NULLIF(excluded.cidade,''), clientes.cidade),
                            uf=COALESCE(NULLIF(excluded.uf,''), clientes.uf),
                            telefone=COALESCE(NULLIF(excluded.telefone,''), clientes.telefone),
                            rota=COALESCE(NULLIF(excluded.rota,''), clientes.rota)
                    """, (
                        upper(cod),
                        upper(nome),
                        upper(endereco),
                        upper(bairro),
                        upper(cidade),
                        upper(uf),
                        upper(telefone),
                        upper(rota)
                    ))
                    total += 1

            self.carregar()
            messagebox.showinfo("OK", f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}")

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

    def alterar_senha(self):
        if self.table != "usuarios":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM USURIO NA TABELA.")
            return
        if not self._user_can_change_password(self.selected_id):
            messagebox.showwarning("ATENÇÃO", "Você não tem permissão para alterar esta senha.")
            return

        require_current = not self._is_admin

        win = tk.Toplevel(self)
        win.title("Alterar Senha")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        row = 0
        ent_atual = None
        if require_current:
            ttk.Label(frm, text="Senha atual:").grid(row=row, column=0, sticky="w", pady=2)
            ent_atual = ttk.Entry(frm, show="*")
            ent_atual.grid(row=row, column=1, sticky="ew", pady=2)
            row += 1

        ttk.Label(frm, text="Nova senha:").grid(row=row, column=0, sticky="w", pady=2)
        ent_nova = ttk.Entry(frm, show="*")
        ent_nova.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1

        ttk.Label(frm, text="Confirmar:").grid(row=row, column=0, sticky="w", pady=2)
        ent_conf = ttk.Entry(frm, show="*")
        ent_conf.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1

        frm.grid_columnconfigure(1, weight=1)

        def _salvar():
            senha_atual = ent_atual.get().strip() if ent_atual else ""
            senha_nova = ent_nova.get().strip()
            senha_conf = ent_conf.get().strip()

            if not senha_nova:
                messagebox.showwarning("ATENÇÃO", "Informe a nova senha.")
                return
            if len(senha_nova) < 6:
                messagebox.showwarning("ATENÇÃO", "A nova senha deve ter pelo menos 6 caracteres.")
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÇÃO", "Confirmação não confere.")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT senha FROM usuarios WHERE id=?", (self.selected_id,))
                row_db = cur.fetchone()
                if not row_db:
                    messagebox.showerror("ERRO", "Usuário não encontrado.")
                    return
                senha_db = row_db[0] or ""

                if require_current:
                    ok = False
                    if str(senha_db).startswith("pbkdf2_sha256$"):
                        ok = verify_password_pbkdf2(senha_atual, senha_db)
                    else:
                        ok = (senha_db == senha_atual)
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", "Senha atual inválida.")
                        return

                nova_hash = hash_password_pbkdf2(senha_nova)
                cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (nova_hash, self.selected_id))

            messagebox.showinfo("OK", "Senha atualizada com sucesso.")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def alterar_senha_motorista(self):
        if self.table != "motoristas":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM MOTORISTA NA TABELA.")
            return
        if not self._is_admin:
            messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha do motorista.")
            return

        win = tk.Toplevel(self)
        win.title("Alterar Senha do Motorista")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Nova senha:").grid(row=0, column=0, sticky="w", pady=2)
        ent_nova = ttk.Entry(frm, show="*")
        ent_nova.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Confirmar:").grid(row=1, column=0, sticky="w", pady=2)
        ent_conf = ttk.Entry(frm, show="*")
        ent_conf.grid(row=1, column=1, sticky="ew", pady=2)

        frm.grid_columnconfigure(1, weight=1)

        def _salvar():
            senha_nova = ent_nova.get().strip()
            senha_conf = ent_conf.get().strip()
            if not senha_nova:
                messagebox.showwarning("ATENÇÃO", "Informe a nova senha.")
                return
            if not is_valid_motorista_senha(senha_nova):
                messagebox.showwarning(
                    "ATENÇÃO",
                    "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                )
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÇÃO", "Confirmação não confere.")
                return

            codigo = self._norm(self._get("codigo"))
            if not codigo:
                messagebox.showwarning("ATENÇÃO", "Código do motorista não encontrado.")
                return

            if not ensure_system_api_binding(context=f"Alterar senha do motorista {codigo}", parent=self):
                return

            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            admin_nome = upper((getattr(self.app, "user", {}) or {}).get("nome", "")) or "ADMIN_DESKTOP"
            payload = {
                "nova_senha": senha_nova,
                "admin": admin_nome,
                "motivo": "Alteracao de senha no cadastro de motoristas (desktop)",
            }
            try:
                _call_api(
                    "POST",
                    f"admin/motoristas/senha/{urllib.parse.quote(codigo)}",
                    payload=payload,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
            except Exception as exc:
                messagebox.showerror("ERRO", f"Falha ao sincronizar senha do motorista na API:\n{exc}")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("UPDATE motoristas SET senha=? WHERE id=?", (self._norm(senha_nova), self.selected_id))

            messagebox.showinfo("OK", "Senha do motorista atualizada com sucesso (LOCAL + API).")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def alterar_senha_vendedor(self):
        if self.table != "vendedores":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM VENDEDOR NA TABELA.")
            return
        if not self._is_admin:
            messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha do vendedor.")
            return

        win = tk.Toplevel(self)
        win.title("Alterar Senha do Vendedor")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Nova senha:").grid(row=0, column=0, sticky="w", pady=2)
        ent_nova = ttk.Entry(frm, show="*")
        ent_nova.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Confirmar:").grid(row=1, column=0, sticky="w", pady=2)
        ent_conf = ttk.Entry(frm, show="*")
        ent_conf.grid(row=1, column=1, sticky="ew", pady=2)

        frm.grid_columnconfigure(1, weight=1)

        def _salvar():
            senha_nova = ent_nova.get().strip()
            senha_conf = ent_conf.get().strip()
            if not senha_nova:
                messagebox.showwarning("ATENÇÃO", "Informe a nova senha.")
                return
            if not is_valid_motorista_senha(senha_nova):
                messagebox.showwarning("ATENÇÃO", "SENHA inválida. Use 4 a 24 caracteres.")
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÇÃO", "Confirmação não confere.")
                return

            codigo = self._norm(self._get("codigo"))
            nome = self._norm(self._get("nome"))
            if not codigo or not nome:
                messagebox.showwarning("ATENÇÃO", "Código ou nome do vendedor não encontrado.")
                return

            if not ensure_system_api_binding(context=f"Alterar senha do vendedor {codigo}", parent=self):
                return

            sync_data = {
                "codigo": codigo,
                "nome": nome,
                "telefone": normalize_phone(self._get("telefone")),
                "cidade_base": self._norm(self._get("cidade_base")),
                "status": self._norm(self._get("status") or "ATIVO"),
                "senha": senha_nova,
            }
            try:
                self._sync_vendedor_upsert_api(sync_data)
            except Exception as exc:
                messagebox.showerror("ERRO", f"Falha ao sincronizar senha do vendedor na API:\n{exc}")
                return

            local_res = update_vendedor_password_hash_local(self.selected_id, hash_password_pbkdf2(senha_nova))
            if isinstance(local_res, dict) and not bool(local_res.get("ok", False)):
                messagebox.showerror(
                    "ERRO",
                    f"Senha sincronizada na API, mas falhou ao atualizar localmente:\n{local_res.get('error')}",
                )
                return

            messagebox.showinfo("OK", "Senha do vendedor atualizada com sucesso (LOCAL + API).")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def definir_acesso_app_motorista(self, liberar: bool):
        if self.table != "motoristas":
            return False
        if not self._is_admin:
            messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar acesso ao app.")
            return False
        ref = self._selected_motorista_ref()
        if not ref:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM MOTORISTA NA TABELA.")
            return False
        self.selected_id = safe_int(ref.get("id"), 0) or self.selected_id

        codigo = self._norm(ref.get("codigo"))
        if not codigo:
            messagebox.showwarning("ATENÇÃO", "Código do motorista não encontrado no item selecionado.")
            return False

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            messagebox.showerror(
                "ERRO",
                "ROTA_SECRET não configurada no Desktop.\nDefina a variável de ambiente e tente novamente.",
            )
            return False

        acao = "LIBERAR" if liberar else "BLOQUEAR"
        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja {acao} o acesso do motorista {codigo} ao app?",
        ):
            return False

        admin_nome = upper((getattr(self.app, "user", {}) or {}).get("nome", "")) or "ADMIN_DESKTOP"
        payload = {
            "liberado": bool(liberar),
            "admin": admin_nome,
            "motivo": "Definido no cadastro de motoristas (desktop)",
        }
        if not ensure_system_api_binding(context=f"Alterar acesso do motorista {codigo}", parent=self):
            return False
        try:
            _call_api(
                "POST",
                f"admin/motoristas/acesso/{urllib.parse.quote(codigo)}",
                payload=payload,
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            messagebox.showinfo("OK", f"Acesso ao app atualizado para {codigo}.")
            return True
        except Exception as exc:
            msg = str(exc or "")
            msg_upper = upper(msg)
            if "404" in msg and ("MOTORISTA NAO ENCONTRADO" in msg_upper or "MOTORISTA NÃO ENCONTRADO" in msg_upper):
                ok_sync = self._upsert_motorista_acesso_na_api(
                    codigo=codigo,
                    liberar=liberar,
                    admin_nome=admin_nome,
                    motivo=payload.get("motivo", ""),
                    desktop_secret=desktop_secret,
                    ref=ref,
                )
                if ok_sync:
                    messagebox.showinfo("OK", f"Motorista {codigo} sincronizado na API e acesso atualizado.")
                    self.carregar()
                    return True
            messagebox.showerror("ERRO", f"Falha ao atualizar acesso do motorista:\n{exc}")
            return False


__all__ = ["CadastroCRUD"]
