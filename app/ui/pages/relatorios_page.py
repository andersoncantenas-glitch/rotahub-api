# -*- coding: utf-8 -*-
import datetime
import logging
import os
import re
import urllib.parse
import urllib.request
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
except Exception:
    canvas = None
    A4 = None

from app.db.connection import get_db
from app.services.api_client import _call_api
from app.context import AppContext
from app.ui.components.page_base import PageBase
from app.utils.excel_helpers import (
    _require_openpyxl as require_openpyxl,
    require_pandas,
    require_excel_support,
    upper,
)
from app.utils.formatters import (
    fmt_money,
    normalize_date,
    normalize_date_time_components,
    safe_float,
    safe_int,
)
from app.utils.text_fix import fix_mojibake_text
from app.utils.validators import validate_required


class RelatoriosPage(PageBase):
    def __init__(self, parent, app, context=None):
        self.context = context if isinstance(context, AppContext) else AppContext(
            config=None,
            hooks={},
        )
        cfg = self.context.config
        hooks = self.context.hooks or {}
        self.APP_CONFIG = cfg

        def _fallback_require_reportlab() -> bool:
            if canvas is None or A4 is None:
                messagebox.showerror("ERRO", "Biblioteca ReportLab não está disponível neste ambiente.")
                return False
            return True

        self.build_folha_retorno_operacional = hooks.get(
            "build_folha_retorno_operacional",
            lambda prog: f"Folha de retorno nao disponivel para {prog or '-'} (hook nao configurado).",
        )
        self.can_read_from_api = hooks.get("can_read_from_api", lambda: False)
        self.db_has_column = hooks.get("db_has_column", lambda *_args, **_kwargs: False)
        self.ensure_system_api_binding = hooks.get(
            "ensure_system_api_binding",
            lambda context="Operacao", parent=None, force_probe=False: False,
        )
        self.fetch_programacao_itens = hooks.get("fetch_programacao_itens", lambda *_args, **_kwargs: [])
        self.is_desktop_api_sync_enabled = hooks.get("is_desktop_api_sync_enabled", lambda: False)
        self.normalize_date_column = hooks.get("normalize_date_column", lambda series: series)
        self.normalize_datetime_column = hooks.get("normalize_datetime_column", lambda series: series)
        self.resolve_equipe_nomes = hooks.get("resolve_equipe_nomes", lambda equipe_raw: str(equipe_raw or ""))
        self.require_reportlab = hooks.get("require_reportlab", _fallback_require_reportlab)

        super().__init__(parent, app, "Relatorios")
        self.body.grid_rowconfigure(2, weight=1)
        self.body.grid_columnconfigure(0, weight=1)

        card = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ttk.Label(card, text="Relatorios operacionais", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            card,
            text="Filtre a programacao, gere o resumo textual e use exportacao, PDF e controle de status no mesmo modulo.",
            style="CardLabel.TLabel",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        filters_wrap = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        filters_wrap.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        for idx in range(6):
            filters_wrap.grid_columnconfigure(idx, weight=1 if idx < 4 else 0)

        ttk.Label(filters_wrap, text="Filtros", style="InsetTitle.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w"
        )

        ttk.Label(filters_wrap, text="Tipo de Relatorio", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.cb_tipo_rel = ttk.Combobox(
            filters_wrap,
            state="readonly",
            values=[
                "Programacoes",
                "Prestacao de Contas",
                "Mortalidade Motorista",
                "Rotina Motorista/Ajudantes",
                "KM de Veiculos",
                "Despesas",
            ],
            width=28,
        )
        self.cb_tipo_rel.grid(row=2, column=0, sticky="ew", padx=(0, 8))
        self.cb_tipo_rel.set("Programacoes")

        ttk.Label(filters_wrap, text="Codigo", style="CardLabel.TLabel").grid(row=1, column=1, sticky="w", pady=(8, 0))
        self.ent_filtro_codigo = ttk.Entry(filters_wrap, style="Field.TEntry", width=16)
        self.ent_filtro_codigo.grid(row=2, column=1, sticky="ew", padx=8)

        ttk.Label(filters_wrap, text="Motorista", style="CardLabel.TLabel").grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.ent_filtro_motorista = ttk.Entry(filters_wrap, style="Field.TEntry", width=24)
        self.ent_filtro_motorista.grid(row=2, column=2, sticky="ew", padx=8)

        ttk.Label(filters_wrap, text="Data", style="CardLabel.TLabel").grid(row=1, column=3, sticky="w", pady=(8, 0))
        self.ent_filtro_data = ttk.Entry(filters_wrap, style="Field.TEntry", width=12)
        self.ent_filtro_data.grid(row=2, column=3, sticky="ew", padx=8)
        self._bind_date_mask_relatorio(self.ent_filtro_data)

        ttk.Button(filters_wrap, text="\U0001F50D BUSCAR", style="Primary.TButton", command=self._buscar_programacoes_relatorio).grid(
            row=2, column=4, padx=(8, 4), sticky="ew"
        )
        ttk.Button(filters_wrap, text="\U0001F9F9 LIMPAR", style="Ghost.TButton", command=self._limpar_filtros_relatorio).grid(
            row=2, column=5, padx=(4, 0), sticky="ew"
        )

        actions_wrap = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        actions_wrap.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        for idx in range(4):
            actions_wrap.grid_columnconfigure(idx, weight=1 if idx == 0 else 0)

        ttk.Label(actions_wrap, text="Saida e controle", style="InsetTitle.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w"
        )
        ttk.Label(actions_wrap, text="Programacao", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.cb_prog = ttk.Combobox(actions_wrap, state="readonly")
        self.cb_prog.grid(row=2, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(actions_wrap, text="\U0001F4CA GERAR RESUMO", style="Primary.TButton", command=self.gerar_resumo).grid(
            row=2, column=1, padx=4, sticky="ew"
        )
        ttk.Button(actions_wrap, text="\U0001F4E4 EXPORTAR EXCEL", style="Warn.TButton", command=self.exportar_excel).grid(
            row=2, column=2, padx=4, sticky="ew"
        )
        ttk.Button(actions_wrap, text="\U0001F4D1 GERAR PDF", style="Primary.TButton", command=self.abrir_previsualizacao_relatorio).grid(
            row=2, column=3, padx=(4, 0), sticky="ew"
        )

        ttk.Button(actions_wrap, text="\U0001F504 ATUALIZAR", style="Ghost.TButton", command=self.refresh_comboboxes).grid(
            row=3, column=1, padx=4, pady=(8, 0), sticky="ew"
        )
        ttk.Button(actions_wrap, text="\U0001F3C1 FINALIZAR ROTA", style="Danger.TButton", command=self.finalizar_rota).grid(
            row=3, column=2, padx=4, pady=(8, 0), sticky="ew"
        )
        ttk.Button(actions_wrap, text="\u21A9 REABRIR ROTA", style="Warn.TButton", command=self.reabrir_rota).grid(
            row=3, column=3, padx=(4, 0), pady=(8, 0), sticky="ew"
        )

        self.var_show_receb_detalhe = tk.BooleanVar(value=True)
        self.var_show_desp_detalhe = tk.BooleanVar(value=True)

        details_frame = ttk.Frame(actions_wrap, style="CardInset.TFrame", padding=(10, 8))
        details_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Label(details_frame, text="Blocos do resumo:", style="InsetTitle.TLabel").pack(side="left", padx=(0, 8))
        ttk.Checkbutton(
            details_frame,
            text="Recebimentos detalhados",
            variable=self.var_show_receb_detalhe,
            command=self._refresh_resumo_if_ready,
        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(
            details_frame,
            text="Despesas detalhadas",
            variable=self.var_show_desp_detalhe,
            command=self._refresh_resumo_if_ready,
        ).pack(side="left")

        dash = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        dash.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        dash.grid_columnconfigure(0, weight=1)
        ttk.Label(dash, text="Resumo do recorte", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        kpi_row = ttk.Frame(dash, style="Card.TFrame")
        kpi_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        for idx in range(4):
            kpi_row.grid_columnconfigure(idx, weight=1, uniform="rel_kpi")

        kpi_1 = ttk.Frame(kpi_row, style="CardInset.TFrame", padding=(10, 8))
        kpi_1.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Label(kpi_1, text="Registros", style="InsetTitle.TLabel").pack(anchor="w")
        self.lbl_kpi_1 = ttk.Label(kpi_1, text="Registros: 0", style="InsetStrong.TLabel", justify="left")
        self.lbl_kpi_1.pack(anchor="w", pady=(6, 0))

        kpi_2 = ttk.Frame(kpi_row, style="CardInset.TFrame", padding=(10, 8))
        kpi_2.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(kpi_2, text="Total", style="InsetTitle.TLabel").pack(anchor="w")
        self.lbl_kpi_2 = ttk.Label(kpi_2, text="Total: R$ 0,00", style="InsetStrong.TLabel", justify="left")
        self.lbl_kpi_2.pack(anchor="w", pady=(6, 0))

        kpi_3 = ttk.Frame(kpi_row, style="CardInset.TFrame", padding=(10, 8))
        kpi_3.grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Label(kpi_3, text="Media", style="InsetTitle.TLabel").pack(anchor="w")
        self.lbl_kpi_3 = ttk.Label(kpi_3, text="Media: 0,00", style="InsetStrong.TLabel", justify="left")
        self.lbl_kpi_3.pack(anchor="w", pady=(6, 0))

        kpi_4 = ttk.Frame(kpi_row, style="CardInset.TFrame", padding=(10, 8))
        kpi_4.grid(row=0, column=3, sticky="ew", padx=(4, 0))
        ttk.Label(kpi_4, text="Destaque", style="InsetTitle.TLabel").pack(anchor="w")
        self.lbl_kpi_4 = ttk.Label(kpi_4, text="Destaque: -", style="InsetStrong.TLabel", justify="left")
        self.lbl_kpi_4.pack(anchor="w", pady=(6, 0))

        chart_frame = ttk.Frame(dash, style="CardInset.TFrame", padding=(10, 8))
        chart_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        chart_frame.grid_columnconfigure(0, weight=1)
        ttk.Label(chart_frame, text="Distribuicao visual", style="InsetTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.cv_chart = tk.Canvas(chart_frame, height=140, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.cv_chart.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self._chart_bars = []
        self._chart_tip_items = ()
        self.cv_chart.bind("<Motion>", self._on_chart_motion)
        self.cv_chart.bind("<Leave>", lambda _e: self._hide_chart_tooltip())

        output_card = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        output_card.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        output_card.grid_columnconfigure(0, weight=1)
        output_card.grid_rowconfigure(1, weight=1)
        ttk.Label(output_card, text="Saida do relatorio", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        output_wrap = ttk.Frame(output_card, style="CardInset.TFrame", padding=(10, 10))
        output_wrap.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        output_wrap.grid_columnconfigure(0, weight=1)
        output_wrap.grid_rowconfigure(0, weight=1)

        self.txt = tk.Text(
            output_wrap,
            height=18,
            wrap="none",
            font=("Consolas", 10),
            background="white",
            relief="flat",
            bd=0,
            padx=8,
            pady=8,
        )
        self.txt.grid(row=0, column=0, sticky="nsew")
        txt_vsb = ttk.Scrollbar(output_wrap, orient="vertical", command=self.txt.yview)
        txt_vsb.grid(row=0, column=1, sticky="ns")
        txt_hsb = ttk.Scrollbar(output_wrap, orient="horizontal", command=self.txt.xview)
        txt_hsb.grid(row=1, column=0, sticky="ew")
        self.txt.configure(yscrollcommand=txt_vsb.set, xscrollcommand=txt_hsb.set)

        self.cb_tipo_rel.bind("<<ComboboxSelected>>", lambda _e: self._buscar_programacoes_relatorio())
        self.cb_prog.bind("<<ComboboxSelected>>", lambda _e: self._refresh_resumo_if_ready())
        self.ent_filtro_codigo.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())
        self.ent_filtro_motorista.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())
        self.ent_filtro_data.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())

        self.refresh_comboboxes()

    # ----------------------------
    # Helpers de regra/segurança
    # ----------------------------
    def _get_prog_status_info(self, prog: str):
        """
        Retorna dict com status/prestacao_status quando existir.
        CompatÃÂÂvel com bases antigas (sem coluna prestacao_status).
        """
        info = {"status": "", "prestacao_status": "", "status_operacional": "", "finalizada_no_app": 0}
        prog = upper(str(prog or "").strip())
        if not prog:
            return info

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and self.can_read_from_api():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(prog)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
                if isinstance(rota, dict):
                    status = upper(str(rota.get("status") or "").strip())
                    status_op = upper(str(rota.get("status_operacional") or "").strip())
                    fin_app = safe_int(rota.get("finalizada_no_app"), 0)
                    info["status_operacional"] = status_op
                    info["finalizada_no_app"] = fin_app
                    info["prestacao_status"] = upper(str(rota.get("prestacao_status") or "").strip())
                    info["status"] = status_op or status
                    if not info["status"] and fin_app == 1:
                        info["status"] = "FINALIZADA"
                    return info
            except Exception:
                logging.debug("Falha ao consultar status da programacao via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {str(c[1]).lower() for c in cur.fetchall()}
            status_op_expr = "COALESCE(status_operacional,'')" if "status_operacional" in cols else "''"
            prest_expr = "COALESCE(prestacao_status,'')" if "prestacao_status" in cols else "''"
            fin_app_expr = "COALESCE(finalizada_no_app,0)" if "finalizada_no_app" in cols else "0"
            cur.execute(
                f"""
                SELECT
                    COALESCE(status,'') AS status,
                    {status_op_expr} AS status_operacional,
                    {prest_expr} AS prestacao_status,
                    {fin_app_expr} AS finalizada_no_app
                FROM programacoes
                WHERE codigo_programacao=?
                ORDER BY id DESC
                LIMIT 1
                """,
                (prog,),
            )
            row = cur.fetchone()
            if row:
                status = upper(row[0] if row[0] is not None else "")
                status_op = upper(row[1] if row[1] is not None else "")
                fin_app = safe_int(row[3], 0)
                info["status_operacional"] = status_op
                info["prestacao_status"] = upper(row[2] if row[2] is not None else "")
                info["finalizada_no_app"] = fin_app
                info["status"] = status_op or status
                if not info["status"] and fin_app == 1:
                    info["status"] = "FINALIZADA"
        return info

    def _is_prestacao_fechada(self, prog: str) -> bool:
        info = self._get_prog_status_info(prog)
        return info.get("prestacao_status", "") == "FECHADA"

    def _tipo_exige_programacao(self) -> bool:
        tipo = upper(self.cb_tipo_rel.get().strip())
        return tipo in ("PROGRAMACOES", "PRESTACAO DE CONTAS")

    def refresh_comboboxes(self):
        self._buscar_programacoes_relatorio()

    def _limpar_filtros_relatorio(self):
        self.cb_tipo_rel.set("Programacoes")
        self.ent_filtro_codigo.delete(0, "end")
        self.ent_filtro_motorista.delete(0, "end")
        self.ent_filtro_data.delete(0, "end")
        self._buscar_programacoes_relatorio()

    def _buscar_programacoes_relatorio(self):
        tipo_rel = upper(self.cb_tipo_rel.get().strip()) if hasattr(self, "cb_tipo_rel") else "PROGRAMACOES"
        filtro_cod = upper(self.ent_filtro_codigo.get().strip()) if hasattr(self, "ent_filtro_codigo") else ""
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        filtro_data_raw = str(self.ent_filtro_data.get().strip()) if hasattr(self, "ent_filtro_data") else ""

        data_patterns = []
        if filtro_data_raw:
            n = normalize_date(filtro_data_raw)
            if n:
                data_patterns.append(n)
                data_patterns.append(f"{n[8:10]}/{n[5:7]}/{n[0:4]}")
            data_patterns.append(filtro_data_raw)
            data_patterns = [p for p in dict.fromkeys(data_patterns) if p]

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and self.is_desktop_api_sync_enabled():
            try:
                modo = "finalizadas_prestacao" if "PRESTACAO" in tipo_rel else "todas"
                resp = _call_api(
                    "GET",
                    f"desktop/programacoes?modo={urllib.parse.quote(modo)}&limit=400",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                encontrados = []
                for r in (arr or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    motorista = upper((r or {}).get("motorista") or "")
                    data_ref = str((r or {}).get("data_referencia") or (r or {}).get("data_criacao") or "")
                    if filtro_cod and filtro_cod not in codigo:
                        continue
                    if filtro_mot and filtro_mot not in motorista:
                        continue
                    if data_patterns:
                        txt = upper(data_ref)
                        ok_data = any(upper(p) in txt for p in data_patterns)
                        if not ok_data:
                            continue
                    if codigo:
                        encontrados.append(codigo)
                atual = upper(self.cb_prog.get().strip())
                self.cb_prog["values"] = encontrados
                if atual not in encontrados:
                    self.cb_prog.set("")
                self.set_status(f"STATUS: {len(encontrados)} programacoes encontradas para {self.cb_tipo_rel.get()}.")
                return
            except Exception:
                logging.debug("Falha ao buscar programacoes de relatorio via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols = [str(c[1]).lower() for c in cur.fetchall()]

                sql = "SELECT codigo_programacao FROM programacoes WHERE 1=1"
                params = []

                if filtro_cod:
                    sql += " AND UPPER(COALESCE(codigo_programacao,'')) LIKE ?"
                    params.append(f"%{filtro_cod}%")

                if filtro_mot and "motorista" in cols:
                    sql += " AND UPPER(COALESCE(motorista,'')) LIKE ?"
                    params.append(f"%{filtro_mot}%")

                if data_patterns:
                    date_cols = [c for c in ("data_criacao", "data", "data_saida") if c in cols]
                    if date_cols:
                        clauses = []
                        for dc in date_cols:
                            for pat in data_patterns:
                                clauses.append(f"COALESCE({dc},'') LIKE ?")
                                params.append(f"%{pat}%")
                        sql += " AND (" + " OR ".join(clauses) + ")"

                if "PRESTACAO" in tipo_rel:
                    if "prestacao_status" in cols:
                        sql += " AND (UPPER(COALESCE(prestacao_status,'')) IN ('PENDENTE','FECHADA') OR UPPER(COALESCE(status,''))='FINALIZADA')"
                    elif "status" in cols:
                        sql += " AND UPPER(COALESCE(status,''))='FINALIZADA'"

                if "id" in cols:
                    if "status_operacional" in cols and "status" in cols:
                        sql += (
                            " ORDER BY CASE "
                            "WHEN UPPER(COALESCE(NULLIF(status_operacional,''), status, '')) IN ('ATIVA','EM_ROTA','CARREGADA','INICIADA') THEN 0 "
                            "WHEN UPPER(COALESCE(NULLIF(status_operacional,''), status, '')) IN ('FINALIZADA','FINALIZADO') THEN 1 "
                            "ELSE 2 END, id DESC LIMIT 400"
                        )
                    elif "status" in cols:
                        sql += " ORDER BY CASE WHEN UPPER(COALESCE(status,''))='ATIVA' THEN 0 ELSE 1 END, id DESC LIMIT 400"
                    else:
                        sql += " ORDER BY id DESC LIMIT 400"
                else:
                    sql += " LIMIT 400"

                cur.execute(sql, tuple(params))
            except Exception:
                cur.execute("SELECT codigo_programacao FROM programacoes ORDER BY id DESC LIMIT 400")

            encontrados = [r[0] for r in cur.fetchall()]
            atual = upper(self.cb_prog.get().strip())
            self.cb_prog["values"] = encontrados
            if atual not in encontrados:
                self.cb_prog.set("")

        self.set_status(f"STATUS: {len(encontrados)} programacoes encontradas para {self.cb_tipo_rel.get()}.")

    def _refresh_resumo_if_ready(self):
        try:
            if (not self._tipo_exige_programacao()) or upper(self.cb_prog.get().strip()):
                self.gerar_resumo()
        except Exception:
            logging.debug("Falha ignorada")

    def _api_bundle_relatorio(self, prog: str):
        prog = upper(str(prog or "").strip())
        self._last_bundle_api_state = "disabled"
        if not prog:
            self._last_bundle_api_state = "invalid"
            return None
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not self.is_desktop_api_sync_enabled():
            self._last_bundle_api_state = "disabled"
            return None
        try:
            resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(prog)}/bundle",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            rota = resp.get("rota") if isinstance(resp, dict) else None
            clientes = resp.get("clientes") if isinstance(resp, dict) else []
            receb = resp.get("recebimentos") if isinstance(resp, dict) else []
            desp = resp.get("despesas") if isinstance(resp, dict) else []
            if not isinstance(rota, dict):
                self._last_bundle_api_state = "not_found"
                return None
            self._last_bundle_api_state = "ok"
            return {
                "rota": rota,
                "clientes": clientes if isinstance(clientes, list) else [],
                "recebimentos": receb if isinstance(receb, list) else [],
                "despesas": desp if isinstance(desp, list) else [],
            }
        except Exception:
            self._last_bundle_api_state = "failed"
            logging.debug("Falha ao montar bundle de relatorio via API", exc_info=True)
            return None

    def _fmt_rel_money(self, v):
        return f"R$ {safe_float(v, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _draw_chart(self, labels, values, color="#1D4ED8"):
        try:
            cv = self.cv_chart
            cv.delete("all")
            self._chart_bars = []
            self._chart_tip_items = ()
            w = max(cv.winfo_width(), 10)
            h = max(cv.winfo_height(), 10)
            if (not labels) or (not values):
                cv.create_text(w / 2, h / 2, text="Sem dados para gráfico", fill="#6B7280", font=("Segoe UI", 10))
                return

            n = min(len(labels), len(values), 10)
            labels = labels[:n]
            vals = [max(safe_float(v, 0.0), 0.0) for v in values[:n]]
            vmax = max(vals) if vals else 1.0
            if vmax <= 0:
                vmax = 1.0

            # Paleta no estilo da referência (barras coloridas + legenda)
            palette = [
                "#F68D56", "#F2C96D", "#D7E64A", "#63D04A", "#67D08D",
                "#58C6C4", "#5E9AC8", "#A78BFA", "#F472B6", "#22C55E"
            ]

            # ÃÂrea do gráfico
            left = 52
            right = 16
            top = 46
            bottom = 26
            plot_w = max(w - left - right, 40)
            plot_h = max(h - top - bottom, 30)
            x0 = left
            y0 = top
            x1 = left + plot_w
            y1 = top + plot_h

            # Fundo
            cv.create_rectangle(0, 0, w, h, fill="#FFFFFF", outline="")

            # Grade horizontal + escala Y
            def _nice_step(v):
                if v <= 5:
                    return 1.0
                if v <= 10:
                    return 2.0
                if v <= 25:
                    return 5.0
                if v <= 50:
                    return 10.0
                if v <= 100:
                    return 20.0
                return max(round(v / 5.0), 1.0)

            step = _nice_step(vmax)
            y_top = step * max(int((vmax / step) + 0.999), 1)
            ticks = 5
            for i in range(ticks + 1):
                v = (y_top / ticks) * i
                yy = y1 - (v / y_top) * plot_h
                cv.create_line(x0, yy, x1, yy, fill="#D1D5DB", width=1)
                txt = f"{v:.0f}" if abs(v - round(v)) < 1e-9 else f"{v:.1f}"
                cv.create_text(x0 - 8, yy, text=txt, anchor="e", fill="#4B5563", font=("Segoe UI", 8, "bold"))

            # Eixos
            cv.create_line(x0, y0, x0, y1, fill="#374151", width=1)
            cv.create_line(x0, y1, x1, y1, fill="#374151", width=2)

            # Barras
            gap = 10
            bw = max((plot_w - gap * (n + 1)) / n, 10)
            for i, (lb, vv) in enumerate(zip(labels, vals)):
                c = palette[i % len(palette)] if n > 1 else color
                bx1 = x0 + gap + i * (bw + gap)
                bx2 = bx1 + bw
                bh = (vv / y_top) * plot_h
                by1 = y1 - bh
                cv.create_rectangle(bx1, by1, bx2, y1, fill=c, outline="")
                cv.create_text((bx1 + bx2) / 2, y1 + 11, text=str(lb)[:10], fill="#111827", font=("Segoe UI", 8))
                self._chart_bars.append({
                    "x1": bx1, "y1": by1, "x2": bx2, "y2": y1,
                    "label": str(lb), "value": vv, "color": c
                })

            # Legenda superior (duas linhas quando necessário)
            legend_x = x0 + 4
            legend_y = 10
            max_w = x1 - x0 - 8
            cx = legend_x
            cy = legend_y
            for i, lb in enumerate(labels):
                c = palette[i % len(palette)] if n > 1 else color
                entry_text = str(lb)[:16]
                txt_w = max(len(entry_text) * 6, 32)
                block_w = 14 + 6 + txt_w + 12
                if (cx - legend_x + block_w) > max_w:
                    cx = legend_x
                    cy += 18
                cv.create_rectangle(cx, cy + 3, cx + 12, cy + 11, fill=c, outline="")
                cv.create_text(cx + 16, cy + 7, text=entry_text, anchor="w", fill="#374151", font=("Segoe UI", 8, "bold"))
                cx += block_w
        except Exception:
            logging.debug("Falha ignorada")

    def _hide_chart_tooltip(self):
        try:
            cv = self.cv_chart
            for iid in self._chart_tip_items:
                try:
                    cv.delete(iid)
                except Exception:
                    logging.debug("Falha ignorada")
            self._chart_tip_items = ()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_chart_motion(self, event=None):
        try:
            if event is None:
                return
            hit = None
            for bar in (self._chart_bars or []):
                if bar["x1"] <= event.x <= bar["x2"] and bar["y1"] <= event.y <= bar["y2"]:
                    hit = bar
                    break
            if not hit:
                self._hide_chart_tooltip()
                return

            self._hide_chart_tooltip()
            cv = self.cv_chart
            txt = f"{hit['label']}: {safe_float(hit['value'], 0.0):.2f}"
            tx = min(max(event.x + 10, 10), max(cv.winfo_width() - 150, 10))
            ty = max(event.y - 24, 8)
            rect = cv.create_rectangle(tx, ty, tx + 140, ty + 20, fill="#111827", outline="")
            label = cv.create_text(tx + 6, ty + 10, text=txt, anchor="w", fill="white", font=("Segoe UI", 8, "bold"))
            self._chart_tip_items = (rect, label)
        except Exception:
            logging.debug("Falha ignorada")

    def _set_dashboard(self, k1, k2, k3, k4, labels=None, values=None, color="#1D4ED8"):
        self.lbl_kpi_1.config(text=k1)
        self.lbl_kpi_2.config(text=k2)
        self.lbl_kpi_3.config(text=k3)
        self.lbl_kpi_4.config(text=k4)
        self._draw_chart(labels or [], values or [], color=color)

    def _bind_date_mask_relatorio(self, ent: tk.Entry):
        digits_state = {"value": ""}

        def _is_partial_date_digits_valid(digits: str) -> bool:
            n = len(digits)
            if n == 0:
                return True
            if n >= 1 and int(digits[0]) > 3:
                return False
            if n >= 2:
                dd = int(digits[:2])
                if dd < 1 or dd > 31:
                    return False
            if n >= 3 and int(digits[2]) > 1:
                return False
            if n >= 4:
                mm = int(digits[2:4])
                if mm < 1 or mm > 12:
                    return False
            if n == 8:
                try:
                    dd = int(digits[:2])
                    mm = int(digits[2:4])
                    yy = int(digits[4:8])
                    datetime(yy, mm, dd)
                except Exception:
                    return False
            return True

        def _to_masked(digits: str) -> str:
            if len(digits) <= 2:
                return digits
            if len(digits) <= 4:
                return f"{digits[:2]}/{digits[2:]}"
            return f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"

        def _apply_mask():
            try:
                raw = str(ent.get() or "")
                digits = re.sub(r"\D", "", raw)[:8]
                while digits and (not _is_partial_date_digits_valid(digits)):
                    digits = digits[:-1]
                ent.delete(0, "end")
                ent.insert(0, _to_masked(digits))
                ent.icursor("end")
                digits_state["value"] = digits
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out():
            try:
                digits = digits_state["value"]
                if digits and len(digits) != 8:
                    ent.delete(0, "end")
                    digits_state["value"] = ""
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<KeyRelease>", lambda _e: _apply_mask())
        ent.bind("<FocusOut>", lambda _e: (_apply_mask(), _on_focus_out()))

    def on_show(self):
        self.refresh_comboboxes()
        self.set_status("STATUS: Relatórios e exportação.")

    def abrir_previsualizacao_relatorio(self):
        tipo_rel = upper(self.cb_tipo_rel.get().strip())
        prog = upper(self.cb_prog.get().strip())
        if self._tipo_exige_programacao() and not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao para visualizar.")
            return
        self.gerar_resumo()
        txt = self.txt.get("1.0", "end").strip()
        top = tk.Toplevel(self)
        top.title(f"Pré-visualização - {self.cb_tipo_rel.get()}")
        top.geometry("1180x760")
        top.minsize(900, 600)
        toolbar = ttk.Frame(top, padding=(8, 8, 8, 0))
        toolbar.pack(fill="x")
        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        footer = ttk.Frame(top, padding=(8, 0, 8, 8))
        footer.pack(fill="x")
        preview_zoom = tk.IntVar(value=100)
        preview_mode = tk.StringVar(value="fit_page")

        ttk.Button(toolbar, text="Atualizar", style="Ghost.TButton", command=lambda: self.abrir_previsualizacao_relatorio()).pack(side="left")
        ttk.Button(toolbar, text="\U0001F5A8 Imprimir", style="Primary.TButton", command=self.gerar_pdf).pack(side="left", padx=(8, 0))
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Label(toolbar, text="Folha:", style="CardLabel.TLabel").pack(side="left", padx=(0, 4))
        ttk.Radiobutton(toolbar, text="Inteira", value="fit_page", variable=preview_mode).pack(side="left")
        ttk.Radiobutton(toolbar, text="Largura", value="fit_width", variable=preview_mode).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(toolbar, text="Real", value="actual", variable=preview_mode).pack(side="left", padx=(6, 0))
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Label(toolbar, text="Zoom:", style="CardLabel.TLabel").pack(side="left", padx=(12, 4))

        def _set_zoom(v):
            try:
                novo = max(50, min(200, int(v)))
            except Exception:
                novo = 100
            preview_zoom.set(novo)
            try:
                cb_zoom.set(str(novo))
            except Exception:
                logging.debug("Falha ignorada")

        ttk.Button(toolbar, text="-", width=3, command=lambda: _set_zoom(preview_zoom.get() - 10)).pack(side="left")
        cb_zoom = ttk.Combobox(toolbar, state="readonly", width=6, values=["75", "90", "100", "110", "125", "150"])
        cb_zoom.pack(side="left", padx=4)
        cb_zoom.set("100")
        cb_zoom.bind("<<ComboboxSelected>>", lambda _e: _set_zoom(cb_zoom.get()))
        ttk.Button(toolbar, text="+", width=3, command=lambda: _set_zoom(preview_zoom.get() + 10)).pack(side="left")
        ttk.Label(toolbar, text="%", style="CardLabel.TLabel").pack(side="left", padx=(2, 0))
        ttk.Button(toolbar, text="Fechar", style="Ghost.TButton", command=top.destroy).pack(side="right")

        if "PROGRAMACOES" in tipo_rel and prog:
            self._create_programacao_canvas_preview_tab(nb, prog, preview_zoom, preview_mode)

        if "PRESTACAO" in tipo_rel and prog:
            self._create_a4_preview_tab(nb, "Folha Prestação", self._build_preview_folha_prestacao(prog))
            self._create_a4_preview_tab(nb, "Folha Retorno", self.build_folha_retorno_operacional(prog))

        tab_resumo = ttk.Frame(nb)
        nb.add(tab_resumo, text="Resumo")
        t = tk.Text(tab_resumo, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("1.0", txt)
        t.configure(state="disabled")

        if "PRESTACAO" in tipo_rel and prog:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
            bundle = self._api_bundle_relatorio(prog)
            api_state = str(getattr(self, "_last_bundle_api_state", ""))
            rec_lines = []
            if bundle:
                rows = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("valor"), 0.0),
                        str((r or {}).get("forma_pagamento") or ""),
                        str((r or {}).get("observacao") or ""),
                    )
                    for r in (bundle.get("recebimentos") or [])
                    if isinstance(r, dict)
                ]
            elif api_enabled and api_state == "not_found":
                rows = []
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), COALESCE(valor,0), COALESCE(forma_pagamento,''), COALESCE(observacao,'')
                        FROM recebimentos WHERE codigo_programacao=? ORDER BY id DESC
                    """, (prog,))
                    rows = cur.fetchall() or []
            rec_lines.append(f"FOLHA DE RECEBIMENTOS - {prog}")
            rec_lines.append("=" * 90)
            rec_lines.append("")
            if not rows:
                rec_lines.append("Sem recebimentos.")
            else:
                rec_lines.append("COD | CLIENTE | VALOR | FORMA | OBS")
                rec_lines.append("-" * 90)
                for r in rows:
                    rec_lines.append(f"{r[0]} | {r[1]} | {self._fmt_rel_money(r[2])} | {r[3] or '-'} | {r[4] or '-'}")
            self._create_a4_preview_tab(nb, "Folha Recebimentos", "\n".join(rec_lines))

            desp_lines = []
            if bundle:
                rows = [
                    (
                        str((r or {}).get("categoria") or "OUTROS"),
                        str((r or {}).get("descricao") or ""),
                        safe_float((r or {}).get("valor"), 0.0),
                        str((r or {}).get("observacao") or ""),
                    )
                    for r in (bundle.get("despesas") or [])
                    if isinstance(r, dict)
                ]
            elif api_enabled and api_state == "not_found":
                rows = []
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT COALESCE(categoria,'OUTROS'), COALESCE(descricao,''), COALESCE(valor,0), COALESCE(observacao,'')
                        FROM despesas WHERE codigo_programacao=? ORDER BY id DESC
                    """, (prog,))
                    rows = cur.fetchall() or []
            desp_lines.append(f"FOLHA DE DESPESAS - {prog}")
            desp_lines.append("=" * 90)
            desp_lines.append("")
            if not rows:
                desp_lines.append("Sem despesas.")
            else:
                desp_lines.append("CATEGORIA | DESCRICAO | VALOR | OBS")
                desp_lines.append("-" * 90)
                for r in rows:
                    desp_lines.append(f"{r[0] or 'OUTROS'} | {r[1]} | {self._fmt_rel_money(r[2])} | {r[3] or '-'}")
            self._create_a4_preview_tab(nb, "Folha Despesas", "\n".join(desp_lines))

    def _create_a4_preview_tab(self, notebook, title: str, content: str):
        """
        Cria aba de preview em formato de folha A4 (largura fixa), com rolagem.
        """
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=title)
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        container = ttk.Frame(tab, style="Content.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg="#ECEFF4", highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # A4 aproximado em tela: largura fixa e aspecto de folha.
        sheet_width_px = 860
        sheet = ttk.Frame(canvas, style="Card.TFrame", padding=18)
        text = tk.Text(
            sheet,
            wrap="none",
            font=("Consolas", 10),
            width=116,
            height=60,
            relief="flat",
            borderwidth=0,
            background="white",
            foreground="#111827",
        )
        text.pack(fill="both", expand=True)
        text.insert("1.0", content or "")
        text.configure(state="disabled")

        win_id = canvas.create_window((20, 20), window=sheet, anchor="nw", width=sheet_width_px)

        def _on_configure(_e=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                logging.debug("Falha ignorada")

        def _on_resize(event=None):
            try:
                cw = max(int(canvas.winfo_width()), 0)
                x = max(int((cw - sheet_width_px) / 2), 20)
                canvas.coords(win_id, x, 20)
                _on_configure()
            except Exception:
                logging.debug("Falha ignorada")

        sheet.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_resize)
        _on_resize()

    def _collect_programacao_payload(self, prog: str):
        programacao_page = None
        if hasattr(self, "app") and hasattr(self.app, "pages"):
            programacao_page = self.app.pages.get("Programacao")

        meta = {}
        if programacao_page and hasattr(programacao_page, "_buscar_meta_programacao"):
            try:
                meta = programacao_page._buscar_meta_programacao(prog) or {}
            except Exception:
                logging.debug("Falha ao buscar metadados da programacao para preview.", exc_info=True)

        itens_raw = self.fetch_programacao_itens(prog) or []
        itens = []
        for item in itens_raw:
            if not isinstance(item, dict):
                continue
            itens.append(
                {
                    "cod_cliente": upper(item.get("cod_cliente", "")),
                    "nome_cliente": upper(item.get("nome_cliente", "")),
                    "endereco": upper(item.get("endereco", "")),
                    "qnt_caixas": safe_int(item.get("caixas_atual"), safe_int(item.get("qnt_caixas"), 0)),
                    "preco": safe_float(item.get("preco_atual"), safe_float(item.get("preco"), 0.0)),
                    "vendedor": upper(item.get("vendedor", "")),
                    "pedido": upper(item.get("pedido", "")),
                    "obs": str(item.get("obs") or item.get("alteracao_detalhe") or "").strip(),
                }
            )
        itens.sort(key=lambda r: (upper(r.get("nome_cliente", "")), upper(r.get("cod_cliente", ""))))
        return programacao_page, meta, itens

    def _create_programacao_canvas_preview_tab(self, notebook, prog: str, zoom_var=None, mode_var=None):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text="Folha Programação")
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        container = ttk.Frame(tab, style="Content.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg="#ECEFF4", highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        page, meta, itens = self._collect_programacao_payload(prog)
        equipe_raw = str(meta.get("equipe") or "")
        equipe_txt = (
            page._resolve_equipe_ajudantes(equipe_raw)
            if page and hasattr(page, "_resolve_equipe_ajudantes")
            else (equipe_raw or "-")
        )
        motorista = upper(meta.get("motorista") or "-")
        veiculo = upper(meta.get("veiculo") or "-")
        usuario_criacao = upper(meta.get("usuario_criacao") or "-")
        usuario_edicao = upper(meta.get("usuario_ultima_edicao") or "-")
        tipo_estimativa = upper(meta.get("tipo_estimativa") or "KG")
        kg_estimado = safe_float(meta.get("kg_estimado"), 0.0)
        caixas_estimado = safe_int(meta.get("caixas_estimado"), 0)

        def _render(_e=None):
            canvas.delete("all")
            cw = max(canvas.winfo_width(), 980)
            ch = max(canvas.winfo_height(), 720)
            pad = 24
            ratio = 210.0 / 297.0
            zoom = 1.0
            try:
                if zoom_var is not None:
                    zoom = max(0.5, min(2.0, float(zoom_var.get()) / 100.0))
            except Exception:
                zoom = 1.0
            mode = "fit_page"
            try:
                if mode_var is not None:
                    mode = str(mode_var.get() or "fit_page").strip()
            except Exception:
                mode = "fit_page"
            avail_w = max(400, cw - 2 * pad)
            avail_h = max(500, ch - 2 * pad)
            if mode == "fit_width":
                page_w = avail_w
                page_h = page_w / ratio
            elif mode == "actual":
                page_h = max(580, avail_h * zoom)
                page_w = page_h * ratio
            else:
                page_h = avail_h
                page_w = page_h * ratio
                if page_w > avail_w:
                    page_w = avail_w
                    page_h = page_w / ratio
                page_h *= zoom
                page_w *= zoom
            if page_w > avail_w and mode != "actual":
                page_w = avail_w
                page_h = page_w / ratio
            x0 = max(20, (cw - page_w) / 2)
            y0 = 20
            x1 = x0 + page_w
            y1 = y0 + page_h
            canvas.create_rectangle(x0, y0, x1, y1, fill="white", outline="#C7CDD4", width=1)
            canvas.configure(scrollregion=(0, 0, max(cw, x1 + 24), max(ch, y1 + 24)))

            def px(rx):
                return x0 + (rx * page_w)

            def py(ry):
                return y0 + (ry * page_h)

            # Mantem o preview proporcional ao PDF real gerado em A4.
            pdf_w = 595.0
            left_margin = 40.0
            col_cx_pdf = 340.0
            col_preco_pdf = 410.0
            col_vendedor_pdf = 430.0
            col_pedido_pdf = 520.0

            def ppdf(x_pdf):
                return x0 + ((x_pdf / pdf_w) * page_w)

            canvas.create_text(px(0.10), py(0.07), text=f"PROGRAMACAO: {prog}", anchor="w", font=("Segoe UI", 19, "bold"))
            canvas.create_text(px(0.10), py(0.12), text=f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}", anchor="w", font=("Segoe UI", 11))
            canvas.create_text(
                px(0.10),
                py(0.15),
                text=f"Motorista: {motorista}  |  Veiculo: {veiculo}  |  Equipe: {equipe_txt or '-'}",
                anchor="w",
                font=("Segoe UI", 11),
            )
            estimado_txt = (
                f"Estimado (FOB): {caixas_estimado} CX" if tipo_estimativa == "CX" else f"Estimado (CIF): {kg_estimado:.2f} KG"
            )
            canvas.create_text(px(0.10), py(0.18), text=estimado_txt, anchor="w", font=("Segoe UI", 11))
            canvas.create_text(
                px(0.10),
                py(0.21),
                text=f"Criado por: {usuario_criacao}  |  Ultima edicao: {usuario_edicao}",
                anchor="w",
                font=("Segoe UI", 11),
            )

            header_y = py(0.27)
            x_cliente = ppdf(left_margin)
            x_cx = ppdf(col_cx_pdf)
            x_preco = ppdf(col_preco_pdf)
            x_vendedor = ppdf(col_vendedor_pdf)
            x_pedido = ppdf(col_pedido_pdf)

            canvas.create_text(x_cliente, header_y, text="CLIENTE / ENDERECO", anchor="w", font=("Segoe UI", 11, "bold"))
            canvas.create_text(ppdf(320), header_y, text="CX", anchor="center", font=("Segoe UI", 10, "bold"))
            canvas.create_text(ppdf(370), header_y, text="PRECO", anchor="center", font=("Segoe UI", 10, "bold"))
            canvas.create_text(ppdf(470), header_y, text="VENDEDOR", anchor="center", font=("Segoe UI", 10, "bold"))
            canvas.create_text(ppdf(548), header_y, text="PEDIDO", anchor="center", font=("Segoe UI", 10, "bold"))

            line_y = py(0.285)
            canvas.create_line(ppdf(40), line_y, ppdf(555), line_y, fill="#222")

            y = py(0.315)
            row_gap = page_h * 0.060
            max_rows = 10
            for idx, item in enumerate(itens[:max_rows]):
                cidade = (
                    page._extrair_cidade_do_endereco(item["endereco"])
                    if page and hasattr(page, "_extrair_cidade_do_endereco")
                    else ""
                )
                linha_cliente = f"{cidade} - {item['cod_cliente']} - {item['nome_cliente']}" if cidade else f"{item['cod_cliente']} - {item['nome_cliente']}"
                linha_cliente = linha_cliente[:34] + "..." if len(linha_cliente) > 37 else linha_cliente
                vendedor_txt = item["vendedor"][:11]
                pedido_txt = str(item["pedido"])[:8]

                canvas.create_text(x_cliente, y, text=linha_cliente, anchor="w", font=("Segoe UI", 8))
                canvas.create_text(ppdf(320), y, text=str(item["qnt_caixas"]), anchor="center", font=("Segoe UI", 8))
                canvas.create_text(ppdf(370), y, text=f"{item['preco']:.2f}", anchor="center", font=("Segoe UI", 8))
                canvas.create_text(ppdf(470), y, text=vendedor_txt, anchor="center", font=("Segoe UI", 8))
                canvas.create_text(ppdf(548), y, text=pedido_txt, anchor="center", font=("Segoe UI", 8))

                obs_y = y + page_h * 0.018
                canvas.create_text(ppdf(50), obs_y, text="OBS:", anchor="w", font=("Segoe UI", 8, "italic"))
                canvas.create_line(ppdf(82), obs_y + 2, ppdf(285), obs_y + 2, fill="#222")
                y += row_gap

            if len(itens) > max_rows:
                canvas.create_text(
                    px(0.10),
                    y,
                    text=f"... e mais {len(itens) - max_rows} cliente(s) na programacao.",
                    anchor="w",
                    font=("Segoe UI", 10, "italic"),
                    fill="#6B7280",
                )

        canvas.bind("<Configure>", _render)
        if zoom_var is not None:
            zoom_var.trace_add("write", lambda *_: _render())
        if mode_var is not None:
            mode_var.trace_add("write", lambda *_: _render())
        _render()

    def _gerar_pdf_programacao_relatorio(self, prog: str):
        if not self.require_reportlab():
            return
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        page, meta, itens = self._collect_programacao_payload(prog)
        if not itens:
            messagebox.showwarning("ATENCAO", "Sem itens na programacao.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar PDF da Programacao",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"PROGRAMACAO_{prog}.pdf"
        )
        if not path:
            return

        try:
            c = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            y = h - 60
            to_txt = lambda v: fix_mojibake_text(str(v or ""))

            motorista = upper(meta.get("motorista") or "")
            veiculo = upper(meta.get("veiculo") or "")
            equipe_raw = str(meta.get("equipe") or "")
            equipe_txt = (
                page._resolve_equipe_ajudantes(equipe_raw)
                if page and hasattr(page, "_resolve_equipe_ajudantes")
                else (equipe_raw or "-")
            )
            kg_estimado = safe_float(meta.get("kg_estimado"), 0.0)
            tipo_estimativa = upper(meta.get("tipo_estimativa") or "KG")
            caixas_estimado = safe_int(meta.get("caixas_estimado"), 0)
            usuario_criacao = upper(meta.get("usuario_criacao") or "")
            usuario_edicao = upper(meta.get("usuario_ultima_edicao") or "")

            c.setFont("Helvetica-Bold", 14)
            c.drawString(40, y, f"PROGRAMACAO: {to_txt(prog)}")
            y -= 22

            c.setFont("Helvetica", 10)
            c.drawString(40, y, f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            y -= 16
            c.drawString(40, y, f"Motorista: {to_txt(motorista)}  |  Veiculo: {to_txt(veiculo)}  |  Equipe: {to_txt(equipe_txt)}")
            y -= 16
            if tipo_estimativa == "CX":
                c.drawString(40, y, f"Estimado (FOB): {safe_int(caixas_estimado, 0)} CX")
            else:
                c.drawString(40, y, f"Estimado (CIF): {safe_float(kg_estimado, 0.0):.2f} KG")
            y -= 16
            c.drawString(40, y, f"Criado por: {to_txt(usuario_criacao or '-')}  |  Ultima edicao: {to_txt(usuario_edicao or '-')}")
            y -= 24

            c.setFont("Helvetica-Bold", 9)
            c.drawString(40, y, "CLIENTE / ENDERECO")
            c.drawString(320, y, "CX")
            c.drawString(370, y, "PRECO")
            c.drawString(430, y, "VENDEDOR")
            c.drawString(520, y, "PEDIDO")
            y -= 10
            c.line(40, y, w - 40, y)
            y -= 14
            c.setFont("Helvetica", 8)

            for item in itens:
                if y < 95:
                    c.showPage()
                    y = h - 60
                    c.setFont("Helvetica", 8)

                cidade = (
                    page._extrair_cidade_do_endereco(item["endereco"])
                    if page and hasattr(page, "_extrair_cidade_do_endereco")
                    else ""
                )
                linha_cliente = f"{cidade} - {item['cod_cliente']} - {item['nome_cliente']}" if cidade else f"{item['cod_cliente']} - {item['nome_cliente']}"
                if len(linha_cliente) > 78:
                    linha_cliente = linha_cliente[:78] + "..."

                c.drawString(40, y, to_txt(linha_cliente))
                c.drawRightString(340, y, str(item["qnt_caixas"]))
                c.drawRightString(410, y, f"{safe_float(item['preco'], 0.0):.2f}")
                c.drawString(430, y, to_txt(item["vendedor"])[:12])
                c.drawString(520, y, to_txt(item["pedido"])[:18])
                y -= 12

                obs_line = f"OBS: {to_txt(item['obs'])}" if item["obs"] else "OBS: ________________________________________________"
                if len(obs_line) > 110:
                    obs_line = obs_line[:110] + "..."
                c.setFont("Helvetica-Oblique", 8)
                c.drawString(50, y, obs_line)
                c.setFont("Helvetica", 8)
                y -= 14

            if y < 40:
                c.showPage()
                y = h - 60

            c.setFont("Helvetica-Oblique", 9)
            c.drawCentredString(w / 2, 26, '"Tudo posso naquele que me fortalece." (Filipenses 4:13)')
            c.save()
            messagebox.showinfo("OK", "PDF gerado com sucesso! (A4 pronto para impressao)")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF: {str(e)}")

    def _build_preview_folha_programacao(self, prog: str) -> str:
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                clientes_api = [r for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                kg_estimado = safe_float(rota.get("kg_estimado"), 0.0)
                status = str(rota.get("status") or "")
                local = str(rota.get("local_rota") or rota.get("local") or "")
                tipo_estimativa = str(rota.get("tipo_estimativa") or "KG")
                caixas_estimado = safe_int(rota.get("caixas_estimado"), 0)
                usuario_criacao = str(
                    rota.get("usuario_criacao")
                    or rota.get("usuario")
                    or rota.get("criado_por")
                    or rota.get("created_by")
                    or rota.get("autor")
                    or ""
                )
                usuario_edicao = str(rota.get("usuario_ultima_edicao") or "")
                itens = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0)),
                        str((r or {}).get("vendedor") or ""),
                        str((r or {}).get("cidade") or ""),
                        str((r or {}).get("pedido") or ""),
                        safe_int((r or {}).get("caixas_atual"), safe_int((r or {}).get("qnt_caixas"), 0)),
                    )
                    for r in clientes_api
                ]
                itens.sort(key=lambda r: (upper(r[1] or ""), upper(r[0] or "")))
                equipe_txt = self.resolve_equipe_nomes(equipe)
                total_prev = sum(safe_float(r[2], 0.0) for r in itens)
                tipo_estimativa = upper(tipo_estimativa or "KG")
                if tipo_estimativa == "CX":
                    estimativa_txt = f"FOB / CX ESTIMADO: {safe_int(caixas_estimado, 0)}"
                else:
                    estimativa_txt = f"CIF / KG ESTIMADO: {safe_float(kg_estimado, 0.0):.2f}"
                lines = []
                lines.append("=" * 118)
                lines.append(f"{'FOLHA DE PROGRAMAÃ‡ÃƒO':^118}")
                lines.append("=" * 118)
                lines.append(f"CÃ“DIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}")
                lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   LOCAL: {upper(local or '-')}")
                lines.append(f"EQUIPE: {equipe_txt or '-'}")
                lines.append(
                    f"{estimativa_txt}   CLIENTES: {len(itens)}   TOTAL ESTIMADO: {self._fmt_rel_money(total_prev)}"
                )
                lines.append(
                    f"CRIADO POR: {upper(usuario_criacao or '-')}" +
                    f"   ÃšLTIMA EDIÃ‡ÃƒO: {upper(usuario_edicao or '-')}"
                )
                lines.append("-" * 118)
                lines.append(f"{'COD':<6} {'CLIENTE / CIDADE':<48} {'CX':>4} {'PREÃ‡O':>12} {'VENDEDOR':<22} {'PEDIDO':>12}")
                lines.append("-" * 118)
                for cod, nome, preco, vendedor, cidade, pedido, caixas in itens:
                    cli = f"{upper(nome or '-')}"
                    if cidade:
                        cli += f" - {upper(cidade)}"
                    lines.append(
                        f"{str(cod)[:6]:<6} {cli[:48]:<48} {safe_int(caixas,0):>4} {self._fmt_rel_money(preco):>12} "
                        f"{upper(vendedor or '-')[:22]:<22} {str(pedido or '-')[:12]:>12}"
                    )
                lines.append("-" * 118)
                lines.append("ObservaÃ§Ã£o: _______________________________________________________________________________________________")
                lines.append("Assinatura responsÃ¡vel: _________________________________________")
                return "\n".join(lines)
            except Exception:
                logging.debug("Falha ao montar folha de programacao via API bundle", exc_info=True)
                if api_enabled:
                    return (
                        f"FOLHA DE PROGRAMACAO - {prog}\n"
                        + "=" * 90
                        + "\n\nFalha ao processar dados retornados pelo servidor."
                    )

        if api_enabled and api_state == "not_found":
            return (
                f"FOLHA DE PROGRAMACAO - {prog}\n"
                + "=" * 90
                + "\n\nProgramacao nao encontrada no servidor."
            )

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                local_expr = "COALESCE(local_rota,'')" if "local_rota" in cols_p else ("COALESCE(local,'')" if "local" in cols_p else "''")
                kg_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols_p else "0"
                status_expr = "COALESCE(status,'')" if "status" in cols_p else "''"
                tipo_estim_expr = "COALESCE(tipo_estimativa,'KG')" if "tipo_estimativa" in cols_p else "'KG'"
                caixas_estim_expr = "COALESCE(caixas_estimado,0)" if "caixas_estimado" in cols_p else "0"
                user_criacao_expr = "COALESCE(usuario_criacao,'')" if "usuario_criacao" in cols_p else "''"
                user_edicao_expr = "COALESCE(usuario_ultima_edicao,'')" if "usuario_ultima_edicao" in cols_p else "''"

                cur.execute(f"""
                    SELECT COALESCE(data_criacao,''), COALESCE(motorista,''), COALESCE(veiculo,''),
                           COALESCE(equipe,''), {kg_expr}, {status_expr}, {local_expr},
                           {tipo_estim_expr}, {caixas_estim_expr}, {user_criacao_expr}, {user_edicao_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", 0, "", "", "KG", 0, "", "")

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
                if cur.fetchone():
                    cur.execute("PRAGMA table_info(programacao_itens)")
                    cols_i = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                    cidade_expr = "COALESCE(cidade,'')" if "cidade" in cols_i else "''"
                    vendedor_expr = "COALESCE(vendedor,'')" if "vendedor" in cols_i else "''"
                    pedido_expr = "COALESCE(pedido,'')" if "pedido" in cols_i else "''"
                    caixas_expr = "COALESCE(qnt_caixas,0)" if "qnt_caixas" in cols_i else "0"
                    preco_expr = "COALESCE(preco,0)" if "preco" in cols_i else "0"

                    cur.execute(f"""
                        SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), {preco_expr},
                               {vendedor_expr}, {cidade_expr}, {pedido_expr}, {caixas_expr}
                        FROM programacao_itens
                        WHERE codigo_programacao=?
                        ORDER BY nome_cliente ASC, cod_cliente ASC
                    """, (prog,))
                    itens = cur.fetchall() or []
                else:
                    itens = []
        except Exception as e:
            return (
                f"ERRO AO MONTAR FOLHA DE PROGRAMAÇÃO ({prog})\n"
                + "=" * 90
                + f"\n\n{str(e)}\n\n"
                + "Verifique a estrutura do banco e tente novamente."
            )

        data_criacao, motorista, veiculo, equipe, kg_estimado, status, local, tipo_estimativa, caixas_estimado, usuario_criacao, usuario_edicao = meta
        equipe_txt = self.resolve_equipe_nomes(equipe)
        total_prev = sum(safe_float(r[2], 0.0) for r in itens)
        tipo_estimativa = upper(tipo_estimativa or "KG")
        if tipo_estimativa == "CX":
            estimativa_txt = f"FOB / CX ESTIMADO: {safe_int(caixas_estimado, 0)}"
        else:
            estimativa_txt = f"CIF / KG ESTIMADO: {safe_float(kg_estimado, 0.0):.2f}"
        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PROGRAMAÇÃO':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   LOCAL: {upper(local or '-')}")
        lines.append(f"EQUIPE: {equipe_txt or '-'}")
        lines.append(
            f"{estimativa_txt}   CLIENTES: {len(itens)}   TOTAL ESTIMADO: {self._fmt_rel_money(total_prev)}"
        )
        lines.append(
            f"CRIADO POR: {upper(usuario_criacao or '-')}" +
            f"   ÚLTIMA EDIÇÃO: {upper(usuario_edicao or '-')}"
        )
        lines.append("-" * 118)
        lines.append(f"{'COD':<6} {'CLIENTE / CIDADE':<48} {'CX':>4} {'PREÇO':>12} {'VENDEDOR':<22} {'PEDIDO':>12}")
        lines.append("-" * 118)
        for cod, nome, preco, vendedor, cidade, pedido, caixas in itens:
            cli = f"{upper(nome or '-')}"
            if cidade:
                cli += f" - {upper(cidade)}"
            lines.append(
                f"{str(cod)[:6]:<6} {cli[:48]:<48} {safe_int(caixas,0):>4} {self._fmt_rel_money(preco):>12} "
                f"{upper(vendedor or '-')[:22]:<22} {str(pedido or '-')[:12]:>12}"
            )
        lines.append("-" * 118)
        lines.append("Observação: _______________________________________________________________________________________________")
        lines.append("Assinatura responsável: _________________________________________")
        return "\n".join(lines)

    def _build_preview_folha_prestacao(self, prog: str) -> str:
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                receb_api = [r for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp_api = [r for r in (bundle.get("despesas") or []) if isinstance(r, dict)]

                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                status = str(rota.get("status") or "")
                prest = str(rota.get("prestacao_status") or "PENDENTE")
                km_rodado = safe_float(rota.get("km_rodado"), 0.0)
                media_km_l = safe_float(rota.get("media_km_l"), 0.0)
                custo_km = safe_float(rota.get("custo_km"), 0.0)
                adiantamento = safe_float(rota.get("adiantamento"), 0.0)
                mort_aves = safe_int(rota.get("mortalidade_transbordo_aves"), 0)
                mort_kg = safe_float(rota.get("mortalidade_transbordo_kg"), 0.0)
                obs_transb = str(rota.get("obs_transbordo") or "")
                nf_kg_base = safe_float(rota.get("nf_kg"), 0.0)

                total_receb = sum(safe_float(r.get("valor"), 0.0) for r in receb_api)
                total_desp = sum(safe_float(r.get("valor"), 0.0) for r in desp_api)
                receb = [
                    (
                        str(r.get("cod_cliente") or ""),
                        str(r.get("nome_cliente") or ""),
                        safe_float(r.get("valor"), 0.0),
                        str(r.get("forma_pagamento") or ""),
                    )
                    for r in receb_api[:12]
                ]
                desp = [
                    (
                        str(r.get("categoria") or "OUTROS"),
                        str(r.get("descricao") or ""),
                        safe_float(r.get("valor"), 0.0),
                    )
                    for r in desp_api[:12]
                ]

                equipe_txt = self.resolve_equipe_nomes(equipe)
                entradas = total_receb + safe_float(adiantamento, 0.0)
                saidas = total_desp
                resultado = entradas - saidas
                log_info = self._collect_logistica_rastreio(prog)
                kg_nf_util = max(nf_kg_base - safe_float(mort_kg, 0.0), 0.0)

                lines = []
                lines.append("=" * 118)
                lines.append(f"{'FOLHA DE PRESTACAO DE CONTAS':^118}")
                lines.append("=" * 118)
                lines.append(f"CODIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}   PRESTACAO: {upper(prest or '-')}")
                lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}")
                lines.append("-" * 118)
                lines.append(
                    f"ENTRADAS (RECEB+ADIANT.): {self._fmt_rel_money(entradas)}   "
                    f"SAIDAS (DESPESAS): {self._fmt_rel_money(saidas)}   "
                    f"RESULTADO: {self._fmt_rel_money(resultado)}"
                )
                lines.append(
                    f"KM RODADO: {safe_float(km_rodado,0.0):.2f}   "
                    f"MEDIA KM/L: {safe_float(media_km_l,0.0):.2f}   "
                    f"CUSTO/KM: {safe_float(custo_km,0.0):.2f}"
                )
                lines.append("-" * 118)
                lines.append("[RECEBIMENTOS - ULTIMOS LANCAMENTOS]")
                lines.append(f"{'COD':<6} {'CLIENTE':<52} {'VALOR':>14} {'FORMA':<16}")
                for cod, nome, valor, forma in receb:
                    lines.append(f"{str(cod)[:6]:<6} {upper(nome or '-')[:52]:<52} {self._fmt_rel_money(valor):>14} {upper(forma or '-')[:16]:<16}")
                if not receb:
                    lines.append("Sem recebimentos.")
                lines.append("-" * 118)
                lines.append("[DESPESAS - ULTIMOS LANCAMENTOS]")
                lines.append(f"{'CATEGORIA':<20} {'DESCRICAO':<72} {'VALOR':>14}")
                for cat, desc, valor in desp:
                    lines.append(f"{upper(cat or 'OUTROS')[:20]:<20} {upper(desc or '-')[:72]:<72} {self._fmt_rel_money(valor):>14}")
                if not desp:
                    lines.append("Sem despesas.")
                lines.append("-" * 118)
                lines.append("[RASTREABILIDADE LOGISTICA]")
                for ln in (log_info.get("resumo") or []):
                    lines.append(str(ln))
                lines.append(
                    "Conciliacao de caixas: "
                    + ("OK" if bool(log_info.get("itens_ok", True)) else "DIVERGENTE")
                )
                lines.append("-" * 118)
                lines.append("[OCORRENCIAS DE TRANSBORDO / MORTALIDADE]")
                lines.append(f"Mortalidade (aves): {safe_int(mort_aves, 0)}")
                lines.append(
                    f"Mortalidade (KG): {safe_float(mort_kg, 0.0):.2f}".replace(".", ",")
                )
                lines.append(
                    f"KG util NF (NF - mortalidade): {safe_float(kg_nf_util, 0.0):.2f}".replace(".", ",")
                )
                lines.append(f"Obs transbordo: {str(obs_transb or '-').strip() or '-'}")
                lines.append("-" * 118)
                lines.append("Conferido por: ____________________________________   Data: ____/____/________")
                return "\n".join(lines)
            except Exception:
                logging.debug("Falha ao montar folha de prestacao via API bundle", exc_info=True)
                if api_enabled:
                    return (
                        f"FOLHA DE PRESTACAO - {prog}\n"
                        + "=" * 90
                        + "\n\nFalha ao processar dados retornados pelo servidor."
                    )

        if api_enabled and api_state == "not_found":
            return (
                f"FOLHA DE PRESTACAO - {prog}\n"
                + "=" * 90
                + "\n\nProgramacao nao encontrada no servidor."
            )

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                status_expr = "COALESCE(status,'')" if "status" in cols_p else "''"
                prest_expr = "COALESCE(prestacao_status,'')" if "prestacao_status" in cols_p else "''"
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols_p else "0"
                media_expr = "COALESCE(media_km_l,0)" if "media_km_l" in cols_p else "0"
                custo_expr = "COALESCE(custo_km,0)" if "custo_km" in cols_p else "0"
                adiant_expr = "COALESCE(adiantamento,0)" if "adiantamento" in cols_p else "0"
                mort_aves_expr = "COALESCE(mortalidade_transbordo_aves,0)" if "mortalidade_transbordo_aves" in cols_p else "0"
                mort_kg_expr = "COALESCE(mortalidade_transbordo_kg,0)" if "mortalidade_transbordo_kg" in cols_p else "0"
                obs_transb_expr = "COALESCE(obs_transbordo,'')" if "obs_transbordo" in cols_p else "''"
                data_criacao_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols_p else ("COALESCE(data,'')" if "data" in cols_p else "''")
                cur.execute(f"""
                    SELECT {data_criacao_expr}, COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           {status_expr}, {prest_expr}, {km_expr}, {media_expr}, {custo_expr}, {adiant_expr},
                           {mort_aves_expr}, {mort_kg_expr}, {obs_transb_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", "", "", 0, 0, 0, 0, 0, 0.0, "")

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
                if cur.fetchone():
                    cur.execute("PRAGMA table_info(recebimentos)")
                    cols_rec = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                    cur.execute("SELECT COALESCE(SUM(valor),0) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                    total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)
                    cod_rec_col = "COALESCE(cod_cliente,'')" if "cod_cliente" in cols_rec else "''"
                    nome_rec_col = "COALESCE(nome_cliente,'')" if "nome_cliente" in cols_rec else "''"
                    valor_rec_col = "COALESCE(valor,0)" if "valor" in cols_rec else "0"
                    forma_rec_col = "COALESCE(forma_pagamento,'')" if "forma_pagamento" in cols_rec else "''"
                    order_rec_col = "id DESC" if "id" in cols_rec else "1 DESC"
                    cur.execute(
                        f"""
                        SELECT {cod_rec_col}, {nome_rec_col}, {valor_rec_col}, {forma_rec_col}
                        FROM recebimentos
                        WHERE codigo_programacao=?
                        ORDER BY {order_rec_col}
                        LIMIT 12
                        """,
                        (prog,),
                    )
                    receb = cur.fetchall() or []
                else:
                    total_receb = 0.0
                    receb = []

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                if cur.fetchone():
                    cur.execute("PRAGMA table_info(despesas)")
                    cols_desp = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                    cur.execute("SELECT COALESCE(SUM(valor),0) FROM despesas WHERE codigo_programacao=?", (prog,))
                    total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)
                    categoria_desp_col = "COALESCE(categoria,'OUTROS')" if "categoria" in cols_desp else "'OUTROS'"
                    descricao_desp_col = "COALESCE(descricao,'')" if "descricao" in cols_desp else "''"
                    valor_desp_col = "COALESCE(valor,0)" if "valor" in cols_desp else "0"
                    order_desp_col = "id DESC" if "id" in cols_desp else "1 DESC"
                    cur.execute(
                        f"""
                        SELECT {categoria_desp_col}, {descricao_desp_col}, {valor_desp_col}
                        FROM despesas
                        WHERE codigo_programacao=?
                        ORDER BY {order_desp_col}
                        LIMIT 12
                        """,
                        (prog,),
                    )
                    desp = cur.fetchall() or []
                else:
                    total_desp = 0.0
                    desp = []
        except Exception as e:
            return (
                f"ERRO AO MONTAR FOLHA DE PRESTAÇÃO ({prog})\n"
                + "=" * 90
                + f"\n\n{str(e)}\n\n"
                + "Verifique a estrutura do banco e tente novamente."
            )

        (
            data_criacao,
            motorista,
            veiculo,
            equipe,
            status,
            prest,
            km_rodado,
            media_km_l,
            custo_km,
            adiantamento,
            mort_aves,
            mort_kg,
            obs_transb,
        ) = meta
        equipe_txt = self.resolve_equipe_nomes(equipe)
        entradas = total_receb + safe_float(adiantamento, 0.0)
        saidas = total_desp
        resultado = entradas - saidas
        log_info = self._collect_logistica_rastreio(prog)
        kg_nf_util = 0.0
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                nf_kg_expr = "COALESCE(nf_kg,0)" if "nf_kg" in cols_prog else "0"
                cur.execute(
                    f"""
                    SELECT {nf_kg_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (prog,),
                )
                row_nf = cur.fetchone()
                nf_kg_base = safe_float((row_nf[0] if row_nf else 0.0), 0.0)
                kg_nf_util = max(nf_kg_base - safe_float(mort_kg, 0.0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")

        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PRESTAÇÃO DE CONTAS':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}   PRESTAÇÃO: {upper(prest or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}")
        lines.append("-" * 118)
        lines.append(
            f"ENTRADAS (RECEB+ADIANT.): {self._fmt_rel_money(entradas)}   "
            f"SAIDAS (DESPESAS): {self._fmt_rel_money(saidas)}   "
            f"RESULTADO: {self._fmt_rel_money(resultado)}"
        )
        lines.append(
            f"KM RODADO: {safe_float(km_rodado,0.0):.2f}   "
            f"MÉDIA KM/L: {safe_float(media_km_l,0.0):.2f}   "
            f"CUSTO/KM: {safe_float(custo_km,0.0):.2f}"
        )
        lines.append("-" * 118)
        lines.append("[RECEBIMENTOS - ÚLTIMOS LANÇAMENTOS]")
        lines.append(f"{'COD':<6} {'CLIENTE':<52} {'VALOR':>14} {'FORMA':<16}")
        for cod, nome, valor, forma in receb:
            lines.append(f"{str(cod)[:6]:<6} {upper(nome or '-')[:52]:<52} {self._fmt_rel_money(valor):>14} {upper(forma or '-')[:16]:<16}")
        if not receb:
            lines.append("Sem recebimentos.")
        lines.append("-" * 118)
        lines.append("[DESPESAS - ÚLTIMOS LANÇAMENTOS]")
        lines.append(f"{'CATEGORIA':<20} {'DESCRIÇÃO':<72} {'VALOR':>14}")
        for cat, desc, valor in desp:
            lines.append(f"{upper(cat or 'OUTROS')[:20]:<20} {upper(desc or '-')[:72]:<72} {self._fmt_rel_money(valor):>14}")
        if not desp:
            lines.append("Sem despesas.")
        lines.append("-" * 118)
        lines.append("[RASTREABILIDADE LOGISTICA]")
        for ln in (log_info.get("resumo") or []):
            lines.append(str(ln))
        lines.append(
            "Conciliação de caixas: "
            + ("OK" if bool(log_info.get("itens_ok", True)) else "DIVERGENTE")
        )
        lines.append("-" * 118)
        lines.append("[OCORRÊNCIAS DE TRANSBORDO / MORTALIDADE]")
        lines.append(f"Mortalidade (aves): {safe_int(mort_aves, 0)}")
        lines.append(
            f"Mortalidade (KG): {safe_float(mort_kg, 0.0):.2f}".replace(".", ",")
        )
        lines.append(
            f"KG útil NF (NF - mortalidade): {safe_float(kg_nf_util, 0.0):.2f}".replace(".", ",")
        )
        lines.append(f"Obs transbordo: {str(obs_transb or '-').strip() or '-'}")
        lines.append("-" * 118)
        lines.append("Conferido por: ____________________________________   Data: ____/____/________")
        return "\n".join(lines)

    def _gerar_relatorio_rotina_motoristas_ajudantes(self):
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and self.is_desktop_api_sync_enabled():
            try:
                q = urllib.parse.quote(filtro_mot, safe="")
                resp = _call_api(
                    "GET",
                    f"desktop/relatorios/rotina-motoristas?motorista_like={q}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("codigo_programacao") or ""),
                            str((r or {}).get("motorista") or ""),
                            str((r or {}).get("equipe") or ""),
                            str((r or {}).get("status") or ""),
                            safe_float((r or {}).get("kg_vendido"), 0.0),
                            safe_float((r or {}).get("km_rodado"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio de rotina via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                if "nf_kg_vendido" in cols and "kg_vendido" in cols:
                    kg_expr = "COALESCE(nf_kg_vendido, kg_vendido, 0)"
                elif "nf_kg_vendido" in cols:
                    kg_expr = "COALESCE(nf_kg_vendido,0)"
                elif "kg_vendido" in cols:
                    kg_expr = "COALESCE(kg_vendido,0)"
                else:
                    kg_expr = "0"
                cur.execute(f"""
                    SELECT COALESCE(codigo_programacao,''), COALESCE(motorista,''), COALESCE(equipe,''), COALESCE(status,''), {kg_expr}, {km_expr}
                    FROM programacoes
                    ORDER BY id DESC
                """)
                rows = cur.fetchall() or []

        mot = {}
        aju = {}
        for codigo, motorista, equipe, status, kg_vendido, km_rodado in rows:
            m = upper(motorista or "SEM MOTORISTA")
            if filtro_mot and filtro_mot not in m:
                continue
            d = mot.setdefault(m, {"viagens": 0, "kg": 0.0, "km": 0.0, "em_rota": 0})
            d["viagens"] += 1
            d["kg"] += safe_float(kg_vendido, 0.0)
            d["km"] += safe_float(km_rodado, 0.0)
            if upper(status) in ("EM_ROTA", "ATIVA", "CARREGADA", "INICIADA"):
                d["em_rota"] += 1
            for nm in re.split(r"[|,;/]+", self.resolve_equipe_nomes(equipe) or ""):
                an = upper((nm or "").strip())
                if not an or an in ("-", "NAN", "NONE", "SEM EQUIPE"):
                    continue
                a = aju.setdefault(an, {"viagens": 0, "km": 0.0})
                a["viagens"] += 1
                a["km"] += safe_float(km_rodado, 0.0)

        rank_m = sorted(mot.items(), key=lambda kv: (-kv[1]["viagens"], kv[0]))
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE ROTINA - MOTORISTAS E AJUDANTES\n")
        self.txt.insert("end", "=" * 100 + "\n\n")
        self.txt.insert("end", "[MOTORISTAS]\n")
        self.txt.insert("end", "MOTORISTA | VIAGENS | KG ENTREGUES | KM RODADO | MEDIA KM/VIAGEM | EM ROTA\n")
        self.txt.insert("end", "-" * 100 + "\n")
        for nome, d in rank_m:
            media_km = d["km"] / max(d["viagens"], 1)
            self.txt.insert("end", f"{nome} | {d['viagens']} | {d['kg']:.2f} | {d['km']:.2f} | {media_km:.2f} | {d['em_rota']}\n")
        self.txt.insert("end", "\n[AJUDANTES]\n")
        self.txt.insert("end", "AJUDANTE | VIAGENS | KM RODADO | MEDIA KM/VIAGEM\n")
        self.txt.insert("end", "-" * 100 + "\n")
        for nome, d in sorted(aju.items(), key=lambda kv: (-kv[1]["viagens"], kv[0])):
            media_km = d["km"] / max(d["viagens"], 1)
            self.txt.insert("end", f"{nome} | {d['viagens']} | {d['km']:.2f} | {media_km:.2f}\n")

        top_nome = rank_m[0][0] if rank_m else "-"
        top_viagens = rank_m[0][1]["viagens"] if rank_m else 0
        self._set_dashboard(
            f"Registros: {len(rank_m)} motoristas",
            f"Total viagens: {sum(d['viagens'] for _n, d in rank_m)}",
            f"Total KG: {sum(d['kg'] for _n, d in rank_m):.2f}",
            f"Destaque: {top_nome} ({top_viagens} viagens)",
            [n for n, _d in rank_m[:8]],
            [d["viagens"] for _n, d in rank_m[:8]],
            color="#0EA5E9",
        )
        self.set_status(f"STATUS: Relatório de rotina gerado ({len(rank_m)} motorista(s)).")

    def _gerar_relatorio_km_veiculos(self):
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and self.is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/relatorios/km-veiculos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("veiculo") or ""),
                            safe_int((r or {}).get("viagens"), 0),
                            safe_float((r or {}).get("km_rodado"), 0.0),
                            safe_float((r or {}).get("media_km_l"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio KM via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                media_expr = "COALESCE(media_km_l,0)" if "media_km_l" in cols else "0"
                cur.execute(f"""
                    SELECT COALESCE(veiculo,''), COUNT(1), COALESCE(SUM({km_expr}),0), COALESCE(AVG(NULLIF({media_expr},0)),0)
                    FROM programacoes
                    GROUP BY COALESCE(veiculo,'')
                    ORDER BY 3 DESC, 1 ASC
                """)
                rows = cur.fetchall() or []
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE KM POR VEICULO\n")
        self.txt.insert("end", "=" * 90 + "\n\n")
        self.txt.insert("end", "VEICULO | VIAGENS | KM RODADO | MELHOR MEDIA KM/L\n")
        self.txt.insert("end", "-" * 90 + "\n")
        for veic, viagens, km, media in rows:
            self.txt.insert("end", f"{upper(veic or '-')} | {safe_int(viagens, 0)} | {safe_float(km, 0.0):.2f} | {safe_float(media, 0.0):.2f}\n")
        total_km = sum(safe_float(r[2], 0.0) for r in rows)
        top = rows[0][0] if rows else "-"
        self._set_dashboard(
            f"Veiculos: {len(rows)}",
            f"KM total: {total_km:.2f}",
            f"Media KM/veiculo: {(total_km / max(len(rows), 1)):.2f}",
            f"Destaque: {upper(top or '-')}",
            [upper((r[0] or "-")) for r in rows[:8]],
            [safe_float(r[2], 0.0) for r in rows[:8]],
            color="#2563EB",
        )
        self.set_status(f"STATUS: Relatorio de KM por veiculo gerado ({len(rows)} veiculo(s)).")

    def _gerar_relatorio_despesas_geral(self):
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and self.is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/relatorios/despesas-categorias",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("categoria") or "OUTROS"),
                            safe_int((r or {}).get("qtd"), 0),
                            safe_float((r or {}).get("total"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio de despesas via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                if cur.fetchone():
                    cur.execute("PRAGMA table_info(despesas)")
                    cols_desp = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                    categoria_expr = "COALESCE(categoria,'OUTROS')" if "categoria" in cols_desp else "'OUTROS'"
                    valor_expr = "COALESCE(valor,0)" if "valor" in cols_desp else "0"
                    cur.execute(
                        f"""
                        SELECT {categoria_expr} AS categoria, COUNT(1) AS qtd, COALESCE(SUM({valor_expr}),0) AS total
                        FROM despesas
                        GROUP BY {categoria_expr}
                        ORDER BY total DESC, categoria ASC
                        """
                    )
                    rows = cur.fetchall() or []
                else:
                    rows = []
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO GERAL DE DESPESAS\n")
        self.txt.insert("end", "=" * 90 + "\n\n")
        self.txt.insert("end", "CATEGORIA | QTD LANÇAMENTOS | TOTAL\n")
        self.txt.insert("end", "-" * 90 + "\n")
        for cat, qtd, total in rows:
            self.txt.insert("end", f"{upper(cat or 'OUTROS')} | {safe_int(qtd, 0)} | {self._fmt_rel_money(total)}\n")
        total_geral = sum(safe_float(r[2], 0.0) for r in rows)
        top = upper(rows[0][0]) if rows else "-"
        self._set_dashboard(
            f"Categorias: {len(rows)}",
            f"Total despesas: {self._fmt_rel_money(total_geral)}",
            f"Média/categoria: {self._fmt_rel_money(total_geral / max(len(rows), 1))}",
            f"Maior categoria: {top}",
            [upper((r[0] or "OUTROS")) for r in rows[:8]],
            [safe_float(r[2], 0.0) for r in rows[:8]],
            color="#DC2626",
        )
        self.set_status(f"STATUS: Relatório de despesas gerado ({len(rows)} categoria(s)).")

    def gerar_resumo(self):
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Programacoes"
        is_mortalidade = "MORTALIDADE" in upper(tipo_rel)
        is_rotina = "ROTINA" in upper(tipo_rel)
        is_km_veic = "KM DE VEICULOS" in upper(tipo_rel)
        is_despesas = upper(tipo_rel) == "DESPESAS"
        if is_mortalidade:
            self._gerar_resumo_mortalidade_motorista()
            return
        if is_rotina:
            self._gerar_relatorio_rotina_motoristas_ajudantes()
            return
        if is_km_veic:
            self._gerar_relatorio_km_veiculos()
            return
        if is_despesas:
            self._gerar_relatorio_despesas_geral()
            return

        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        is_prestacao = "PRESTACAO" in upper(tipo_rel)

        def fmt_money(v):
            vv = safe_float(v, 0.0)
            return f"R$ {vv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        def normalize_unit_price(v):
            vv = safe_float(v, 0.0)
            if abs(vv) >= 100 and abs(vv - round(vv)) < 1e-9:
                return vv / 100.0
            return vv

        motorista = veiculo = equipe = ""
        usuario_criacao = ""
        status = prestacao = ""
        data_criacao = ""
        nf = local_rota = local_carreg = ""
        data_saida = hora_saida = data_chegada = hora_chegada = ""
        kg_estimado = 0.0
        nf_kg = nf_caixas = nf_kg_carregado = nf_kg_vendido = nf_saldo = 0.0
        km_inicial = km_final = litros = km_rodado = media_km_l = custo_km = 0.0
        ced_qtd = {200: 0, 100: 0, 50: 0, 20: 0, 10: 0, 5: 0, 2: 0}
        valor_dinheiro = 0.0
        adiantamento = 0.0
        total_entregas = 0
        total_receb = 0.0
        total_desp = 0.0
        recebimentos = []
        despesas = []
        clientes_programacao = []

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if api_enabled and api_state == "not_found":
            messagebox.showwarning("ATENCAO", f"Programacao nao encontrada no servidor: {prog}")
            return
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                clientes_api = [r for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                receb_api = [r for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp_api = [r for r in (bundle.get("despesas") or []) if isinstance(r, dict)]

                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                status = upper(str(rota.get("status") or ""))
                prestacao = upper(str(rota.get("prestacao_status") or ""))
                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                usuario_criacao = str(
                    rota.get("usuario")
                    or rota.get("criado_por")
                    or rota.get("usuario_criacao")
                    or rota.get("created_by")
                    or rota.get("autor")
                    or rota.get("responsavel")
                    or ""
                )
                kg_estimado = safe_float(rota.get("kg_estimado"), 0.0)
                nf = str(rota.get("num_nf") or rota.get("nf_numero") or "")
                local_rota = str(rota.get("local_rota") or rota.get("tipo_rota") or "")
                local_carreg = str(rota.get("local_carregamento") or rota.get("granja_carregada") or "")
                data_saida = str(rota.get("data_saida") or "")
                hora_saida = str(rota.get("hora_saida") or "")
                data_chegada = str(rota.get("data_chegada") or "")
                hora_chegada = str(rota.get("hora_chegada") or "")
                nf_kg = safe_float(rota.get("nf_kg"), 0.0)
                nf_caixas = safe_int(rota.get("nf_caixas"), 0)
                nf_kg_carregado = safe_float(rota.get("nf_kg_carregado"), 0.0)
                nf_kg_vendido = safe_float(rota.get("nf_kg_vendido"), 0.0)
                nf_saldo = safe_float(rota.get("nf_saldo"), 0.0)
                km_inicial = safe_float(rota.get("km_inicial"), 0.0)
                km_final = safe_float(rota.get("km_final"), 0.0)
                litros = safe_float(rota.get("litros"), 0.0)
                km_rodado = safe_float(rota.get("km_rodado"), 0.0)
                media_km_l = safe_float(rota.get("media_km_l"), 0.0)
                custo_km = safe_float(rota.get("custo_km"), 0.0)
                ced_qtd[200] = safe_int(rota.get("ced_200_qtd"), 0)
                ced_qtd[100] = safe_int(rota.get("ced_100_qtd"), 0)
                ced_qtd[50] = safe_int(rota.get("ced_50_qtd"), 0)
                ced_qtd[20] = safe_int(rota.get("ced_20_qtd"), 0)
                ced_qtd[10] = safe_int(rota.get("ced_10_qtd"), 0)
                ced_qtd[5] = safe_int(rota.get("ced_5_qtd"), 0)
                ced_qtd[2] = safe_int(rota.get("ced_2_qtd"), 0)
                valor_dinheiro = safe_float(rota.get("valor_dinheiro"), 0.0)
                adiantamento = safe_float(rota.get("adiantamento"), 0.0)

                data_saida, hora_saida = normalize_date_time_components(data_saida, hora_saida)
                data_chegada, hora_chegada = normalize_date_time_components(data_chegada, hora_chegada)

                clientes_programacao = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0)),
                        str((r or {}).get("vendedor") or ""),
                    )
                    for r in clientes_api
                ]
                total_entregas = len(clientes_programacao)

                if is_prestacao:
                    recebimentos = [
                        (
                            str((r or {}).get("cod_cliente") or ""),
                            str((r or {}).get("nome_cliente") or ""),
                            safe_float((r or {}).get("valor"), 0.0),
                            str((r or {}).get("forma_pagamento") or ""),
                            str((r or {}).get("observacao") or ""),
                            str((r or {}).get("data_registro") or ""),
                        )
                        for r in receb_api
                    ]
                    despesas = [
                        (
                            str((r or {}).get("descricao") or ""),
                            safe_float((r or {}).get("valor"), 0.0),
                            str((r or {}).get("categoria") or "OUTROS"),
                            str((r or {}).get("observacao") or ""),
                            str((r or {}).get("data_registro") or ""),
                        )
                        for r in desp_api
                    ]
                    total_receb = sum(safe_float(r[2], 0.0) for r in recebimentos)
                    total_desp = sum(safe_float(r[1], 0.0) for r in despesas)
            except Exception:
                logging.debug("Falha ao carregar resumo via API.", exc_info=True)
                if api_enabled:
                    messagebox.showwarning(
                        "Relatorios",
                        "Falha ao processar dados retornados pelo servidor para esta programacao.\n"
                        "Tente novamente apos atualizar o servidor.",
                    )
                    return
                bundle = None

        if not bundle:
            with get_db() as conn:
                cur = conn.cursor()
                cols = []

                def has_col(col):
                    try:
                        return self.db_has_column(cur, "programacoes", col)
                    except Exception:
                        return False

                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols = [str(c[1]).lower() for c in (cur.fetchall() or [])]
                except Exception:
                    cols = []

                usuario_col_name = ""
                for cand in ("usuario", "criado_por", "usuario_criacao", "user", "created_by", "autor", "responsavel"):
                    if cand in cols:
                        usuario_col_name = cand
                        break

                nf_col = "num_nf" if has_col("num_nf") else ("nf_numero" if has_col("nf_numero") else "'' as num_nf")
                local_rota_col = "local_rota" if has_col("local_rota") else ("tipo_rota" if has_col("tipo_rota") else "'' as local_rota")
                local_carreg_col = "local_carregamento" if has_col("local_carregamento") else ("granja_carregada" if has_col("granja_carregada") else "'' as local_carregamento")

                if has_col("adiantamento") and has_col("adiantamento_rota"):
                    adiant_col = (
                        "CASE WHEN COALESCE(adiantamento, 0) <> 0 THEN adiantamento "
                        "ELSE COALESCE(adiantamento_rota, 0) END as adiantamento"
                    )
                elif has_col("adiantamento"):
                    adiant_col = "adiantamento"
                elif has_col("adiantamento_rota"):
                    adiant_col = "adiantamento_rota as adiantamento"
                else:
                    adiant_col = "0 as adiantamento"

                status_col = "COALESCE(status,'') as status" if has_col("status") else "'' as status"
                prest_col = "COALESCE(prestacao_status,'') as prestacao_status" if has_col("prestacao_status") else "'' as prestacao_status"
                data_col = "COALESCE(data_criacao,'') as data_criacao" if has_col("data_criacao") else ("COALESCE(data,'') as data_criacao" if has_col("data") else "'' as data_criacao")
                usuario_col = f"COALESCE({usuario_col_name},'') as usuario_criacao" if usuario_col_name else "'' as usuario_criacao"
                kg_estimado_col = "COALESCE(kg_estimado,0)" if has_col("kg_estimado") else "0"
                data_saida_col = "COALESCE(data_saida,'')" if has_col("data_saida") else "''"
                hora_saida_col = "COALESCE(hora_saida,'')" if has_col("hora_saida") else "''"
                data_chegada_col = "COALESCE(data_chegada,'')" if has_col("data_chegada") else "''"
                hora_chegada_col = "COALESCE(hora_chegada,'')" if has_col("hora_chegada") else "''"
                nf_kg_col = "COALESCE(nf_kg,0)" if has_col("nf_kg") else "0"
                nf_caixas_col = "COALESCE(nf_caixas,0)" if has_col("nf_caixas") else "0"
                nf_kg_carregado_col = "COALESCE(nf_kg_carregado,0)" if has_col("nf_kg_carregado") else ("COALESCE(kg_carregado,0)" if has_col("kg_carregado") else "0")
                nf_kg_vendido_col = "COALESCE(nf_kg_vendido,0)" if has_col("nf_kg_vendido") else ("COALESCE(kg_vendido,0)" if has_col("kg_vendido") else "0")
                nf_saldo_col = "COALESCE(nf_saldo,0)" if has_col("nf_saldo") else "0"
                km_inicial_col = "COALESCE(km_inicial,0)" if has_col("km_inicial") else "0"
                km_final_col = "COALESCE(km_final,0)" if has_col("km_final") else "0"
                litros_col = "COALESCE(litros,0)" if has_col("litros") else "0"
                km_rodado_col = "COALESCE(km_rodado,0)" if has_col("km_rodado") else "0"
                media_km_l_col = "COALESCE(media_km_l,0)" if has_col("media_km_l") else "0"
                custo_km_col = "COALESCE(custo_km,0)" if has_col("custo_km") else "0"
                ced_200_col = "COALESCE(ced_200_qtd,0)" if has_col("ced_200_qtd") else "0"
                ced_100_col = "COALESCE(ced_100_qtd,0)" if has_col("ced_100_qtd") else "0"
                ced_50_col = "COALESCE(ced_50_qtd,0)" if has_col("ced_50_qtd") else "0"
                ced_20_col = "COALESCE(ced_20_qtd,0)" if has_col("ced_20_qtd") else "0"
                ced_10_col = "COALESCE(ced_10_qtd,0)" if has_col("ced_10_qtd") else "0"
                ced_5_col = "COALESCE(ced_5_qtd,0)" if has_col("ced_5_qtd") else "0"
                ced_2_col = "COALESCE(ced_2_qtd,0)" if has_col("ced_2_qtd") else "0"
                valor_dinheiro_col = "COALESCE(valor_dinheiro,0)" if has_col("valor_dinheiro") else "0"

                cur.execute(f"""
                    SELECT
                        motorista, veiculo, equipe,
                        {status_col}, {prest_col}, {data_col},
                        {usuario_col},
                        {kg_estimado_col},
                        {nf_col} as num_nf,
                        {local_rota_col} as local_rota,
                        {local_carreg_col} as local_carreg,
                        {data_saida_col}, {hora_saida_col},
                        {data_chegada_col}, {hora_chegada_col},
                        {nf_kg_col}, {nf_caixas_col}, {nf_kg_carregado_col},
                        {nf_kg_vendido_col}, {nf_saldo_col},
                        {km_inicial_col}, {km_final_col}, {litros_col},
                        {km_rodado_col}, {media_km_l_col}, {custo_km_col},
                        {ced_200_col}, {ced_100_col}, {ced_50_col},
                        {ced_20_col}, {ced_10_col}, {ced_5_col}, {ced_2_col},
                        {valor_dinheiro_col},
                        {adiant_col}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                row = cur.fetchone()
                if not row:
                    messagebox.showwarning("ATENCAO", f"Programacao nao encontrada: {prog}")
                    return

                motorista, veiculo, equipe = row[0] or "", row[1] or "", row[2] or ""
                status, prestacao = upper(row[3] or ""), upper(row[4] or "")
                data_criacao = row[5] or ""
                usuario_criacao = row[6] or ""
                kg_estimado = safe_float(row[7], 0.0)
                nf, local_rota, local_carreg = row[8] or "", row[9] or "", row[10] or ""
                data_saida, hora_saida = row[11] or "", row[12] or ""
                data_chegada, hora_chegada = row[13] or "", row[14] or ""
                nf_kg, nf_caixas, nf_kg_carregado = safe_float(row[15], 0.0), safe_int(row[16], 0), safe_float(row[17], 0.0)
                nf_kg_vendido, nf_saldo = safe_float(row[18], 0.0), safe_float(row[19], 0.0)
                km_inicial, km_final = safe_float(row[20], 0.0), safe_float(row[21], 0.0)
                litros, km_rodado = safe_float(row[22], 0.0), safe_float(row[23], 0.0)
                media_km_l, custo_km = safe_float(row[24], 0.0), safe_float(row[25], 0.0)
                ced_qtd[200], ced_qtd[100], ced_qtd[50] = safe_int(row[26], 0), safe_int(row[27], 0), safe_int(row[28], 0)
                ced_qtd[20], ced_qtd[10], ced_qtd[5], ced_qtd[2] = safe_int(row[29], 0), safe_int(row[30], 0), safe_int(row[31], 0), safe_int(row[32], 0)
                valor_dinheiro = safe_float(row[33], 0.0)
                adiantamento = safe_float(row[34], 0.0)

                data_saida, hora_saida = normalize_date_time_components(data_saida, hora_saida)
                data_chegada, hora_chegada = normalize_date_time_components(data_chegada, hora_chegada)

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
                has_itens = bool(cur.fetchone())
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
                has_recebimentos = bool(cur.fetchone())
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                has_despesas = bool(cur.fetchone())

                if has_itens:
                    cols_itens = {str(c[1]).lower() for c in (cur.execute("PRAGMA table_info(programacao_itens)").fetchall() or [])}
                    try:
                        cur.execute("SELECT COUNT(*) FROM programacao_itens WHERE codigo_programacao=?", (prog,))
                        total_entregas = safe_int((cur.fetchone() or [0])[0], 0)
                    except Exception:
                        total_entregas = 0

                    cod_cliente_item_col = "COALESCE(cod_cliente,'')" if "cod_cliente" in cols_itens else "''"
                    nome_cliente_item_col = "COALESCE(nome_cliente,'')" if "nome_cliente" in cols_itens else "''"
                    preco_col = "COALESCE(preco,0)" if "preco" in cols_itens else "0"
                    vendedor_item_col = "COALESCE(vendedor,'')" if "vendedor" in cols_itens else "''"
                    cur.execute(
                        f"""
                        SELECT
                            {cod_cliente_item_col},
                            {nome_cliente_item_col},
                            {preco_col},
                            {vendedor_item_col}
                        FROM programacao_itens
                        WHERE codigo_programacao=?
                        ORDER BY nome_cliente ASC, cod_cliente ASC
                        """,
                        (prog,),
                    )
                    clientes_programacao = cur.fetchall() or []
                else:
                    total_entregas = 0
                    clientes_programacao = []

                if is_prestacao:
                    if has_recebimentos:
                        cols_receb = {str(c[1]).lower() for c in (cur.execute("PRAGMA table_info(recebimentos)").fetchall() or [])}
                        cod_cliente_receb_col = "COALESCE(cod_cliente,'')" if "cod_cliente" in cols_receb else "''"
                        nome_cliente_receb_col = "COALESCE(nome_cliente,'')" if "nome_cliente" in cols_receb else "''"
                        valor_receb_col = "COALESCE(valor,0)" if "valor" in cols_receb else "0"
                        forma_receb_col = "COALESCE(forma_pagamento,'')" if "forma_pagamento" in cols_receb else "''"
                        obs_receb_col = "COALESCE(observacao,'')" if "observacao" in cols_receb else "''"
                        data_receb_col = "COALESCE(data_registro,'')" if "data_registro" in cols_receb else "''"
                        cur.execute("SELECT COALESCE(SUM(valor),0) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                        total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)
                        cur.execute(
                            f"""
                            SELECT
                                {cod_cliente_receb_col},
                                {nome_cliente_receb_col},
                                {valor_receb_col},
                                {forma_receb_col},
                                {obs_receb_col},
                                {data_receb_col}
                            FROM recebimentos
                            WHERE codigo_programacao=?
                            ORDER BY data_registro DESC, id DESC
                            """,
                            (prog,),
                        )
                        recebimentos = cur.fetchall() or []
                    else:
                        total_receb = 0.0
                        recebimentos = []

                    if has_despesas:
                        cols_desp = {str(c[1]).lower() for c in (cur.execute("PRAGMA table_info(despesas)").fetchall() or [])}
                        descricao_rel_desp_col = "COALESCE(descricao,'')" if "descricao" in cols_desp else "''"
                        valor_rel_desp_col = "COALESCE(valor,0)" if "valor" in cols_desp else "0"
                        categoria_rel_desp_col = "COALESCE(categoria,'OUTROS')" if "categoria" in cols_desp else "'OUTROS'"
                        obs_rel_desp_col = "COALESCE(observacao,'')" if "observacao" in cols_desp else "''"
                        data_rel_desp_col = "COALESCE(data_registro,'')" if "data_registro" in cols_desp else "''"
                        cur.execute("SELECT COALESCE(SUM(valor),0) FROM despesas WHERE codigo_programacao=?", (prog,))
                        total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)
                        cur.execute(
                            f"""
                            SELECT
                                {descricao_rel_desp_col},
                                {valor_rel_desp_col},
                                {categoria_rel_desp_col},
                                {obs_rel_desp_col},
                                {data_rel_desp_col}
                            FROM despesas
                            WHERE codigo_programacao=?
                            ORDER BY data_registro DESC, id DESC
                            """,
                            (prog,),
                        )
                        despesas = cur.fetchall() or []
                    else:
                        total_desp = 0.0
                        despesas = []

        equipe_txt = self.resolve_equipe_nomes(equipe)
        data_criacao_n = normalize_date(data_criacao)
        data_criacao_show = data_criacao_n if data_criacao_n is not None else (data_criacao or "")

        self.txt.delete("1.0", "end")

        if not is_prestacao:
            valor_total_clientes = sum(normalize_unit_price(r[2]) for r in clientes_programacao)

            self.txt.insert("end", f"RELATORIO DE PROGRAMACAO - {prog}\n")
            self.txt.insert("end", "=" * 90 + "\n\n")

            self.txt.insert("end", "[IDENTIFICACAO]\n")
            self.txt.insert("end", f"Codigo: {prog}\n")
            self.txt.insert("end", f"Status: {status or '-'}\n")
            self.txt.insert("end", f"Data: {data_criacao_show or '-'}\n")
            self.txt.insert("end", f"Usuario criacao: {usuario_criacao or '-'}\n")
            self.txt.insert("end", f"Motorista: {motorista or '-'}\n")
            self.txt.insert("end", f"Equipe: {equipe_txt or '-'}\n")
            self.txt.insert("end", f"Veiculo: {veiculo or '-'}\n")
            self.txt.insert("end", f"KG estimado: {kg_estimado:.2f}\n")
            self.txt.insert("end", f"Clientes na programacao: {len(clientes_programacao)}\n")
            self.txt.insert("end", f"Total estimado (clientes): {fmt_money(valor_total_clientes)}\n\n")

            self.txt.insert("end", "[CLIENTES / PRECO / VENDEDOR]\n")
            if not clientes_programacao:
                self.txt.insert("end", "Sem clientes cadastrados na programacao.\n")
            else:
                self.txt.insert("end", "COD | CLIENTE | PRECO | VENDEDOR\n")
                self.txt.insert("end", "-" * 90 + "\n")
                for cod_cli, nome_cli, preco_cli, vendedor in clientes_programacao:
                    self.txt.insert(
                        "end",
                        f"{(cod_cli or '-')} | {(nome_cli or '-')} | {fmt_money(normalize_unit_price(preco_cli))} | {(vendedor or '-')}\n"
                    )

            self.set_status("STATUS: Resumo de programacao gerado.")
            self._set_dashboard(
                f"Clientes: {len(clientes_programacao)}",
                f"Total estimado: {fmt_money(valor_total_clientes)}",
                f"Preço médio: {fmt_money(valor_total_clientes / max(len(clientes_programacao), 1))}",
                f"Motorista: {motorista or '-'}",
                [str(r[0]) for r in clientes_programacao[:8]],
                [normalize_unit_price(r[2]) for r in clientes_programacao[:8]],
                color="#1D4ED8",
            )
            return

        ced_total = sum(float(ced) * safe_int(qtd, 0) for ced, qtd in ced_qtd.items())
        dinheiro_total = valor_dinheiro if safe_float(valor_dinheiro, 0.0) > 0 else ced_total
        total_entradas = total_receb + adiantamento
        total_saidas = total_desp + ced_total
        valor_final_caixa = total_entradas - total_desp
        diferenca = valor_final_caixa - ced_total
        resultado = total_entradas - total_saidas

        self.txt.insert("end", f"RELATORIO DE PRESTACAO DE CONTAS - PROGRAMACAO {prog}\n")
        self.txt.insert("end", f"Tipo: {tipo_rel}\n")
        self.txt.insert("end", "=" * 90 + "\n\n")

        self.txt.insert("end", "[IDENTIFICACAO]\n")
        self.txt.insert("end", f"Status: {status or '-'}\n")
        self.txt.insert("end", f"Prestacao: {prestacao or '-'}\n")
        self.txt.insert("end", f"Data criacao: {data_criacao_show or '-'}\n")
        self.txt.insert("end", f"Motorista: {motorista or '-'}\n")
        self.txt.insert("end", f"Veiculo: {veiculo or '-'}\n")
        self.txt.insert("end", f"Equipe: {equipe_txt or '-'}\n")
        self.txt.insert("end", f"Entregas (itens): {total_entregas}\n")
        self.txt.insert("end", f"KG estimado: {kg_estimado:.2f}\n\n")

        self.txt.insert("end", "[DADOS DA ROTA]\n")
        self.txt.insert("end", f"NF: {nf or '-'}\n")
        self.txt.insert("end", f"Local da rota: {local_rota or '-'}\n")
        self.txt.insert("end", f"Local carregamento: {local_carreg or '-'}\n")
        self.txt.insert("end", f"Saida: {(data_saida + ' ' + hora_saida).strip() or '-'}\n")
        self.txt.insert("end", f"Chegada: {(data_chegada + ' ' + hora_chegada).strip() or '-'}\n\n")

        self.txt.insert("end", "[NOTA FISCAL / CARREGAMENTO]\n")
        self.txt.insert("end", f"NF KG: {nf_kg:.2f}\n")
        self.txt.insert("end", f"NF caixas: {safe_int(nf_caixas, 0)}\n")
        self.txt.insert("end", f"KG carregado: {nf_kg_carregado:.2f}\n")
        self.txt.insert("end", f"KG vendido: {nf_kg_vendido:.2f}\n")
        self.txt.insert("end", f"Saldo (KG): {nf_saldo:.2f}\n\n")

        self.txt.insert("end", "[ROTA / KM]\n")
        self.txt.insert("end", f"KM inicial: {km_inicial:.2f}\n")
        self.txt.insert("end", f"KM final: {km_final:.2f}\n")
        self.txt.insert("end", f"Litros: {litros:.2f}\n")
        self.txt.insert("end", f"KM rodado: {km_rodado:.2f}\n")
        self.txt.insert("end", f"Media km/l: {media_km_l:.2f}\n")
        self.txt.insert("end", f"Custo por KM: {custo_km:.2f}\n\n")

        self.txt.insert("end", "[CONTAGEM DE CEDULAS]\n")
        for ced in [200, 100, 50, 20, 10, 5, 2]:
            qtd = safe_int(ced_qtd.get(ced, 0), 0)
            self.txt.insert("end", f"R$ {ced:>3},00 -> QTD {qtd:>4} -> TOTAL {fmt_money(qtd * ced)}\n")
        self.txt.insert("end", f"Total cedulas: {fmt_money(ced_total)}\n")
        self.txt.insert("end", f"Total dinheiro (campo): {fmt_money(dinheiro_total)}\n\n")

        self.txt.insert("end", "[RESUMO FINANCEIRO]\n")
        self.txt.insert("end", f"Recebimentos: {fmt_money(total_receb)}\n")
        self.txt.insert("end", f"Adiantamento: {fmt_money(adiantamento)}\n")
        self.txt.insert("end", f"Despesas: {fmt_money(total_desp)}\n")
        self.txt.insert("end", f"Cedulas: {fmt_money(ced_total)}\n")
        self.txt.insert("end", f"Total entradas: {fmt_money(total_entradas)}\n")
        self.txt.insert("end", f"Total saidas: {fmt_money(total_saidas)}\n")
        self.txt.insert("end", f"Valor final caixa: {fmt_money(valor_final_caixa)}\n")
        self.txt.insert("end", f"Diferenca caixa x cedulas: {fmt_money(diferenca)}\n")
        self.txt.insert("end", f"Resultado liquido: {fmt_money(resultado)}\n\n")

        if bool(self.var_show_receb_detalhe.get()):
            self.txt.insert("end", "[RECEBIMENTOS DETALHADOS]\n")
            if not recebimentos:
                self.txt.insert("end", "Sem recebimentos registrados.\n\n")
            else:
                for cod_cli, nome_cli, valor, forma, obs, data_reg in recebimentos:
                    data_show = (data_reg or "")[:19]
                    self.txt.insert("end", f"{data_show} | {cod_cli} | {nome_cli} | {fmt_money(valor)} | {forma or '-'} | {obs or '-'}\n")
                self.txt.insert("end", "\n")

        if bool(self.var_show_desp_detalhe.get()):
            self.txt.insert("end", "[DESPESAS DETALHADAS]\n")
            if not despesas:
                self.txt.insert("end", "Sem despesas registradas.\n")
            else:
                for desc, valor, cat, obs, data_reg in despesas:
                    data_show = (data_reg or "")[:19]
                    self.txt.insert("end", f"{data_show} | {cat or 'OUTROS'} | {desc or '-'} | {fmt_money(valor)} | {obs or '-'}\n")

        if prestacao == "FECHADA":
            self.txt.insert("end", "\n[ALERTA] Prestacao FECHADA: alteracoes financeiras estao bloqueadas.\n")

        self.set_status("STATUS: Resumo detalhado gerado.")
        self._set_dashboard(
            f"Recebimentos: {fmt_money(total_receb)}",
            f"Despesas: {fmt_money(total_desp)}",
            f"Resultado: {fmt_money(resultado)}",
            f"Prestação: {prestacao or '-'}",
            ["Entradas", "SaÃÂdas", "Resultado"],
            [total_entradas, total_saidas, abs(resultado)],
            color="#0EA5E9",
        )

    def _gerar_resumo_mortalidade_motorista(self):
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Mortalidade Motorista"
        filtro_cod = upper(self.ent_filtro_codigo.get().strip()) if hasattr(self, "ent_filtro_codigo") else ""
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        filtro_data_raw = str(self.ent_filtro_data.get().strip()) if hasattr(self, "ent_filtro_data") else ""

        data_patterns = []
        if filtro_data_raw:
            n = normalize_date(filtro_data_raw)
            if n:
                data_patterns.append(n)
                data_patterns.append(f"{n[8:10]}/{n[5:7]}/{n[0:4]}")
            data_patterns.append(filtro_data_raw)
            data_patterns = [p for p in dict.fromkeys(data_patterns) if p]

        rows = []
        used_api = False
        api_failed = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
        if api_enabled:
            try:
                q_cod = urllib.parse.quote(filtro_cod, safe="")
                q_mot = urllib.parse.quote(filtro_mot, safe="")
                q_data = urllib.parse.quote("|".join(data_patterns), safe="")
                resp = _call_api(
                    "GET",
                    f"desktop/relatorios/mortalidade-motorista?codigo_like={q_cod}&motorista_like={q_mot}&data_like={q_data}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("codigo_programacao") or ""),
                            str((r or {}).get("motorista") or ""),
                            str((r or {}).get("data_ref") or ""),
                            str((r or {}).get("status_ref") or ""),
                            safe_int((r or {}).get("mortalidade_total"), 0),
                            safe_int((r or {}).get("clientes_com_mortalidade"), 0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    used_api = True
                else:
                    messagebox.showwarning(
                        "Relatorios",
                        "Resposta invalida do servidor para o relatorio de mortalidade.",
                    )
                    return
            except Exception:
                api_failed = True
                logging.debug("Falha no relatorio de mortalidade via API; usando fallback local.", exc_info=True)

        if not used_api:
            if api_enabled and (not api_failed):
                messagebox.showwarning(
                    "Relatorios",
                    "Nao foi possivel validar os dados do relatorio de mortalidade no servidor.",
                )
                return
            with get_db() as conn:
                cur = conn.cursor()

                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = [str(c[1]).lower() for c in (cur.fetchall() or [])]
                has_status = "status" in cols_p
                has_data_criacao = "data_criacao" in cols_p
                has_data = "data" in cols_p
                data_expr = (
                    "COALESCE(p.data_criacao,'')"
                    if has_data_criacao
                    else ("COALESCE(p.data,'')" if has_data else "''")
                )
                status_expr = "COALESCE(p.status,'')" if has_status else "''"
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens_controle'")
                has_controle = bool(cur.fetchone())
                mort_sum_expr = "COALESCE(SUM(COALESCE(pc.mortalidade_aves, 0)), 0)" if has_controle else "0"
                mort_count_expr = "COUNT(CASE WHEN COALESCE(pc.mortalidade_aves,0) > 0 THEN 1 END)" if has_controle else "0"
                join_controle = (
                    """
                    LEFT JOIN programacao_itens_controle pc
                      ON UPPER(COALESCE(pc.codigo_programacao,'')) = UPPER(COALESCE(p.codigo_programacao,''))
                    """
                    if has_controle
                    else ""
                )

                sql = f"""
                    SELECT
                        COALESCE(p.codigo_programacao,'') as codigo_programacao,
                        COALESCE(p.motorista,'') as motorista,
                        {data_expr} as data_ref,
                        {status_expr} as status_ref,
                        {mort_sum_expr} as mortalidade_total,
                        {mort_count_expr} as clientes_com_mortalidade
                    FROM programacoes p
                    {join_controle}
                    WHERE 1=1
                """
                params = []

                if filtro_cod:
                    sql += " AND UPPER(COALESCE(p.codigo_programacao,'')) LIKE ?"
                    params.append(f"%{filtro_cod}%")
                if filtro_mot:
                    sql += " AND UPPER(COALESCE(p.motorista,'')) LIKE ?"
                    params.append(f"%{filtro_mot}%")
                if data_patterns:
                    clauses = []
                    for pat in data_patterns:
                        clauses.append(f"{data_expr} LIKE ?")
                        params.append(f"%{pat}%")
                    sql += " AND (" + " OR ".join(clauses) + ")"

                sql += """
                    GROUP BY p.codigo_programacao, p.motorista, data_ref, status_ref
                    ORDER BY mortalidade_total ASC, p.codigo_programacao DESC
                """
                cur.execute(sql, tuple(params))
                rows = cur.fetchall() or []

        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE MORTALIDADE POR MOTORISTA\n")
        self.txt.insert("end", f"Tipo: {tipo_rel}\n")
        self.txt.insert("end", "=" * 95 + "\n\n")

        if not rows:
            self.txt.insert("end", "Nenhum dado encontrado para os filtros informados.\n")
            self.set_status("STATUS: Nenhum dado de mortalidade encontrado.")
            return

        # Ranking por rota (menor mortalidade primeiro)
        self.txt.insert("end", "[RANKING POR ROTA - MENOR MORTALIDADE]\n")
        self.txt.insert("end", "POS | PROGRAMAÇÃO | MOTORISTA | MORTALIDADE | CLIENTES C/ MORT. | DATA | STATUS\n")
        self.txt.insert("end", "-" * 95 + "\n")
        for i, (codigo, motorista, data_ref, status_ref, mort_total, cli_mort) in enumerate(rows, start=1):
            self.txt.insert(
                "end",
                f"{i:>3} | {codigo or '-'} | {(motorista or '-')} | {safe_int(mort_total, 0):>11} | {safe_int(cli_mort, 0):>16} | {(data_ref or '-')[:19]} | {upper(status_ref or '-')}\n",
            )

        # Consolidado por motorista
        resumo_mot = {}
        for codigo, motorista, data_ref, status_ref, mort_total, cli_mort in rows:
            mot = upper(motorista or "SEM MOTORISTA")
            item = resumo_mot.setdefault(
                mot,
                {"rotas": 0, "mort_total": 0, "melhor": None, "pior": 0},
            )
            mort = safe_int(mort_total, 0)
            item["rotas"] += 1
            item["mort_total"] += mort
            item["pior"] = max(item["pior"], mort)
            if item["melhor"] is None or mort < item["melhor"]:
                item["melhor"] = mort

        self.txt.insert("end", "\n[CONSOLIDADO POR MOTORISTA]\n")
        self.txt.insert("end", "MOTORISTA | ROTAS | MORTALIDADE TOTAL | MEDIA/ROTA | MELHOR ROTA | PIOR ROTA\n")
        self.txt.insert("end", "-" * 95 + "\n")
        ranking_motorista = sorted(
            resumo_mot.items(),
            key=lambda kv: (
                (kv[1]["mort_total"] / max(kv[1]["rotas"], 1)),
                kv[1]["mort_total"],
                kv[0],
            ),
        )
        for mot, d in ranking_motorista:
            media = d["mort_total"] / max(d["rotas"], 1)
            self.txt.insert(
                "end",
                f"{mot} | {d['rotas']} | {d['mort_total']} | {media:.2f} | {safe_int(d['melhor'], 0)} | {safe_int(d['pior'], 0)}\n",
            )

        melhor_mot, melhor_data = ranking_motorista[0]
        melhor_media = melhor_data["mort_total"] / max(melhor_data["rotas"], 1)
        self.txt.insert("end", "\n[DESTAQUE]\n")
        self.txt.insert(
            "end",
            f"Motorista com menor média de mortalidade por rota: {melhor_mot} "
            f"(média {melhor_media:.2f} aves/rota em {melhor_data['rotas']} rota(s)).\n",
        )
        self._set_dashboard(
            f"Rotas analisadas: {len(rows)}",
            f"Motoristas: {len(resumo_mot)}",
            f"Melhor média: {melhor_media:.2f}",
            f"Destaque: {melhor_mot}",
            [m for m, _d in ranking_motorista[:8]],
            [(_d["mort_total"] / max(_d["rotas"], 1)) for _m, _d in ranking_motorista[:8]],
            color="#7C3AED",
        )
        self.set_status(f"STATUS: Relatório de mortalidade gerado ({len(rows)} rota(s)).")

    def exportar_excel(self):
        prog = upper(self.cb_prog.get())
        if not (require_pandas() and require_openpyxl()):
            return
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        path = filedialog.asksaveasfilename(
            title="Exportar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"RELATORIO_{prog}.xlsx"
        )
        if not path:
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            api_enabled = bool(desktop_secret and self.is_desktop_api_sync_enabled())
            bundle = self._api_bundle_relatorio(prog)
            api_state = str(getattr(self, "_last_bundle_api_state", ""))
            if api_enabled and api_state == "not_found":
                messagebox.showwarning("ATENCAO", f"Programacao nao encontrada no servidor: {prog}")
                return
            if bundle:
                itens = [dict(r) for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                rec = [dict(r) for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp = [dict(r) for r in (bundle.get("despesas") or []) if isinstance(r, dict)]
                df_itens = pd.DataFrame(itens)
                df_rec = pd.DataFrame(rec)
                df_desp = pd.DataFrame(desp)
            else:
                if api_enabled and api_state not in ("failed", "disabled"):
                    messagebox.showwarning(
                        "ATENCAO",
                        "Falha ao processar dados da programacao no servidor. Tente novamente.",
                    )
                    return
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
                    if cur.fetchone():
                        cur.execute("SELECT * FROM programacao_itens WHERE codigo_programacao=?", (prog,))
                        itens = cur.fetchall()
                        cols_itens = [d[0] for d in cur.description]
                    else:
                        itens, cols_itens = [], []

                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
                    if cur.fetchone():
                        cur.execute("SELECT * FROM recebimentos WHERE codigo_programacao=?", (prog,))
                        rec = cur.fetchall()
                        cols_rec = [d[0] for d in cur.description]
                    else:
                        rec, cols_rec = [], []

                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                    if cur.fetchone():
                        cur.execute("SELECT * FROM despesas WHERE codigo_programacao=?", (prog,))
                        desp = cur.fetchall()
                        cols_desp = [d[0] for d in cur.description]
                    else:
                        desp, cols_desp = [], []

                df_itens = pd.DataFrame(itens, columns=cols_itens)
                df_rec = pd.DataFrame(rec, columns=cols_rec)
                df_desp = pd.DataFrame(desp, columns=cols_desp)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:

                if "data_registro" in df_rec.columns:
                    df_rec["data_registro"] = self.normalize_datetime_column(df_rec["data_registro"])
                if "data_registro" in df_desp.columns:
                    df_desp["data_registro"] = self.normalize_datetime_column(df_desp["data_registro"])

                # normaliza datas no ITENS, se existirem
                if "data_venda" in df_itens.columns:
                    df_itens["data_venda"] = self.normalize_date_column(df_itens["data_venda"])
                if "data_saida" in df_itens.columns:
                    df_itens["data_saida"] = self.normalize_date_column(df_itens["data_saida"])
                if "data_chegada" in df_itens.columns:
                    df_itens["data_chegada"] = self.normalize_date_column(df_itens["data_chegada"])
                if "saida_dt" in df_itens.columns:
                    df_itens["saida_dt"] = self.normalize_datetime_column(df_itens["saida_dt"])
                if "chegada_dt" in df_itens.columns:
                    df_itens["chegada_dt"] = self.normalize_datetime_column(df_itens["chegada_dt"])

                df_itens.to_excel(writer, index=False, sheet_name="ITENS")
                df_rec.to_excel(writer, index=False, sheet_name="RECEBIMENTOS")
                df_desp.to_excel(writer, index=False, sheet_name="DESPESAS")

            messagebox.showinfo("OK", "Excel exportado com sucesso!")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao exportar Excel: {str(e)}")

    def gerar_pdf(self):
        if not self.require_reportlab():
            return
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Programacoes"
        is_mortalidade = "MORTALIDADE" in upper(tipo_rel)
        is_prestacao = "PRESTACAO" in upper(tipo_rel)
        prog = upper(self.cb_prog.get())

        # Prestacao de contas: reutiliza exatamente o mesmo gerador de PDF da tela de Despesas.
        if is_prestacao:
            if not prog:
                messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
                return
            try:
                despesas_page = None
                if hasattr(self, "app") and hasattr(self.app, "pages"):
                    despesas_page = self.app.pages.get("Despesas")
                if not despesas_page or not hasattr(despesas_page, "imprimir_resumo"):
                    messagebox.showerror("ERRO", "Tela de Despesas indisponivel para gerar PDF da prestacao.")
                    return

                prev_prog = getattr(despesas_page, "_current_programacao", "")
                despesas_page._current_programacao = prog
                despesas_page.imprimir_resumo()
                despesas_page._current_programacao = prev_prog
            except Exception as e:
                messagebox.showerror("ERRO", f"Erro ao gerar PDF da prestacao: {str(e)}")
            return

        if "PROGRAMACOES" in upper(tipo_rel):
            self._gerar_pdf_programacao_relatorio(prog)
            return

        if (not is_mortalidade) and (not prog):
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        nome_base = f"RELATORIO_{prog}" if prog else "RELATORIO_MORTALIDADE_MOTORISTA"

        path = filedialog.asksaveasfilename(
            title="Salvar PDF do Relatorio",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"{nome_base}.pdf"
        )
        if not path:
            return

        self.gerar_resumo()
        txt = self.txt.get("1.0", "end").strip().splitlines()

        try:
            c = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            y = h - 60

            c.setFont("Helvetica-Bold", 14)
            titulo_pdf = f"RELATORIO - {prog}" if prog else "RELATORIO - MORTALIDADE POR MOTORISTA"
            c.drawString(40, y, titulo_pdf)
            y -= 24

            c.setFont("Helvetica", 10)
            for line in txt:
                c.drawString(40, y, line[:110])
                y -= 14
                if y < 60:
                    c.showPage()
                    y = h - 60
                    c.setFont("Helvetica", 10)

            c.save()
            messagebox.showinfo("OK", "PDF gerado com sucesso!")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF: {str(e)}")

    def finalizar_rota(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return
        if not self.ensure_system_api_binding(context=f"Finalizar rota ({prog})", parent=self):
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"A rota {prog} esta com a prestacao FECHADA.\n\n"
                "Ela ja esta travada para alteracoes."
            )
            return

        if st == "FINALIZADA":
            messagebox.showinfo("Info", f"A rota {prog} ja esta FINALIZADA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja FINALIZAR a rota {prog}?\n\n"
            "Ela deixara de aparecer nas Rotas Ativas."
        ):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/status",
                payload={
                    "status": "FINALIZADA",
                    "status_operacional": "FINALIZADA",
                    "finalizada_no_app": 1,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            messagebox.showinfo("OK", f"Rota FINALIZADA: {prog}")

            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                self.app.refresh_programacao_comboboxes()

            self.set_status(f"STATUS: Rota finalizada: {prog}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao finalizar rota: {str(e)}")

    def reabrir_rota(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return
        if not self.ensure_system_api_binding(context=f"Reabrir rota ({prog})", parent=self):
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"Nao e permitido REABRIR a rota {prog} pois a prestacao esta FECHADA.\n\n"
                "Se precisar reabrir, primeiro reabra a prestacao."
            )
            return

        if st == "ATIVA":
            messagebox.showinfo("Info", f"A rota {prog} ja esta ATIVA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja REABRIR a rota {prog}?\n\n"
            "Ela voltara a aparecer nas Rotas Ativas."
        ):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/status",
                payload={
                    "status": "ATIVA",
                    "status_operacional": "",
                    "finalizada_no_app": 0,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            messagebox.showinfo("OK", f"Rota REABERTA: {prog}")

            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                self.app.refresh_programacao_comboboxes()

            self.set_status(f"STATUS: Rota reaberta: {prog}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao reabrir rota: {str(e)}")
    

# ==========================
# ===== FIM DA PARTE 8 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 9 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== FIM DA PARTE 9 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 10 (ATUALIZADA) =====
# ==========================



__all__ = ["RelatoriosPage"]
