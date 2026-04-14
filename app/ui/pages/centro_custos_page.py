# -*- coding: utf-8 -*-
import logging
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk

from app.db.connection import get_db
from app.services.api_client import _call_api
from app.ui.components.page_base import PageBase
from app.ui.components.tree_helpers import enable_treeview_sorting, tree_insert_aligned
from app.utils.excel_helpers import upper
from app.utils.formatters import safe_float, safe_int


def configure_centro_custos_page_dependencies(dependencies=None, **kwargs):
    merged = {}
    if dependencies:
        merged.update(dict(dependencies))
    if kwargs:
        merged.update(kwargs)
    globals().update(merged)


def is_desktop_api_sync_enabled() -> bool:
    # Fallback seguro quando a injeção explicita não foi aplicada.
    return False


class CentroCustosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "CentroCustos")
        self.body.grid_rowconfigure(3, weight=1)
        self.body.grid_columnconfigure(0, weight=1)
        self._chart_labels = []
        self._chart_values = []

        filtros = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        filtros.grid(row=0, column=0, sticky="ew")
        filtros.grid_columnconfigure(8, weight=1)

        ttk.Label(filtros, text="Periodo (dias)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.var_periodo = tk.StringVar(value="30")
        self.cb_periodo = ttk.Combobox(
            filtros, textvariable=self.var_periodo, state="readonly", width=10,
            values=["7", "15", "30", "60", "90", "180", "TODAS"]
        )
        self.cb_periodo.grid(row=1, column=0, sticky="w", padx=(0, 10))

        ttk.Label(filtros, text="Veiculo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.var_veiculo = tk.StringVar(value="TODOS")
        self.cb_veiculo = ttk.Combobox(
            filtros, textvariable=self.var_veiculo, state="readonly", width=16, values=["TODOS"]
        )
        self.cb_veiculo.grid(row=1, column=1, sticky="w", padx=(0, 10))

        ttk.Button(filtros, text="\U0001F504 ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
            row=1, column=2, sticky="w", padx=(0, 6)
        )

        ttk.Label(filtros, text="Metrica do grafico", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.var_chart_metric = tk.StringVar(value="CUSTO_KM")
        self.cb_chart_metric = ttk.Combobox(
            filtros,
            textvariable=self.var_chart_metric,
            state="readonly",
            width=16,
            values=["CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"],
        )
        self.cb_chart_metric.grid(row=1, column=3, sticky="w", padx=(0, 10))

        self.lbl_resumo = ttk.Label(
            self.body,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=980,
        )
        self.lbl_resumo.grid(row=1, column=0, sticky="ew", pady=(8, 8))

        chart_wrap = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        chart_wrap.grid(row=2, column=0, sticky="ew")
        chart_wrap.grid_columnconfigure(0, weight=1)
        self.lbl_chart_title = ttk.Label(chart_wrap, text="Custo por Veiculo (Custo/KM)", style="CardTitle.TLabel")
        self.lbl_chart_title.grid(row=0, column=0, sticky="w")
        self.cv_chart = tk.Canvas(chart_wrap, height=220, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.cv_chart.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.cv_chart.bind("<Configure>", lambda _e: self._draw_chart(self._chart_labels, self._chart_values))

        table_wrap = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        table_wrap.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        cols = ("VEICULO", "ROTAS", "KM_RODADO", "KG_CARREGADO", "DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA")
        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        for c in cols:
            self.tree.heading(c, text=c)
            if c == "VEICULO":
                w = 160
            elif c in ("DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA"):
                w = 130
            else:
                w = 110
            self.tree.column(c, anchor="w" if c == "VEICULO" else "center", width=w)
        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ROTAS", "KM_RODADO", "KG_CARREGADO", "DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA"},
        )

        self.tree.tag_configure("bom", foreground="#14532D", background="#ECFDF3")
        self.tree.tag_configure("medio", foreground="#7C2D12", background="#FFEDD5")
        self.tree.tag_configure("alto", foreground="#7F1D1D", background="#FEE2E2")

        self.cb_periodo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.cb_veiculo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.cb_chart_metric.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())

    def on_show(self):
        self.refresh_comboboxes()
        self.refresh_data()

    def _parse_data(self, raw: str):
        txt = str(raw or "").strip()
        if not txt:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d/%m/%y %H:%M:%S", "%d/%m/%y"):
            try:
                return datetime.strptime(txt[:19], fmt) if "H" in fmt else datetime.strptime(txt[:10], fmt)
            except Exception:
                pass
        try:
            return datetime.fromisoformat(txt.replace(" ", "T"))
        except Exception:
            return None

    def _draw_chart(self, labels, values):
        try:
            cv = self.cv_chart
            cv.delete("all")
            w = max(cv.winfo_width(), 10)
            h = max(cv.winfo_height(), 10)
            if (not labels) or (not values):
                cv.create_text(w / 2, h / 2, text="Sem dados para grafico", fill="#6B7280", font=("Segoe UI", 10))
                return

            n = min(len(labels), len(values), 10)
            labels = labels[:n]
            vals = [max(safe_float(v, 0.0), 0.0) for v in values[:n]]
            vmax = max(vals) if vals else 1.0
            if vmax <= 0:
                vmax = 1.0

            metric = "CUSTO_KM"
            try:
                metric = upper(str(self.var_chart_metric.get() or "").strip()) or "CUSTO_KM"
            except Exception:
                metric = "CUSTO_KM"

            def _fmt_axis(v: float) -> str:
                vv = safe_float(v, 0.0)
                if metric == "DESPESA_TOTAL":
                    return f"R$ {vv:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{vv:.3f}"

            def _fmt_bar(v: float) -> str:
                vv = safe_float(v, 0.0)
                if metric == "DESPESA_TOTAL":
                    return f"R$ {vv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{vv:.3f}"

            palette = ["#F68D56", "#F2C96D", "#D7E64A", "#63D04A", "#67D08D", "#58C6C4", "#5E9AC8", "#A78BFA", "#F472B6", "#22C55E"]
            left, right, top, bottom = 50, 16, 30, 28
            plot_w = max(w - left - right, 40)
            plot_h = max(h - top - bottom, 30)
            x0, y0 = left, top
            x1, y1 = left + plot_w, top + plot_h

            ticks = 5
            for i in range(ticks + 1):
                v = (vmax / ticks) * i
                yy = y1 - (v / vmax) * plot_h
                cv.create_line(x0, yy, x1, yy, fill="#E5E7EB")
                cv.create_text(x0 - 6, yy, text=_fmt_axis(v), anchor="e", fill="#4B5563", font=("Segoe UI", 8, "bold"))

            cv.create_line(x0, y0, x0, y1, fill="#374151")
            cv.create_line(x0, y1, x1, y1, fill="#374151", width=2)

            gap = 10
            bw = max((plot_w - gap * (n + 1)) / n, 10)
            for i, (lb, vv) in enumerate(zip(labels, vals)):
                c = palette[i % len(palette)]
                bx1 = x0 + gap + i * (bw + gap)
                bx2 = bx1 + bw
                bh = (vv / vmax) * plot_h
                by1 = y1 - bh
                cv.create_rectangle(bx1, by1, bx2, y1, fill=c, outline="")
                cv.create_text((bx1 + bx2) / 2, y1 + 11, text=str(lb)[:10], fill="#111827", font=("Segoe UI", 8))
                if bh > 16:
                    cv.create_text((bx1 + bx2) / 2, by1 - 8, text=_fmt_bar(vv), fill="#111827", font=("Segoe UI", 7, "bold"))
        except Exception:
            logging.debug("Falha ignorada")

    def refresh_comboboxes(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=1200",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                vals = sorted(
                    {
                        upper((r or {}).get("veiculo") or "")
                        for r in (arr or [])
                        if isinstance(r, dict) and str((r or {}).get("veiculo") or "").strip()
                    }
                )
                values = ["TODOS"] + vals
                self.cb_veiculo.configure(values=values)
                if self.var_veiculo.get() not in values:
                    self.var_veiculo.set("TODOS")
                return
            except Exception:
                logging.debug("Falha ao carregar filtro de veiculos via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT DISTINCT upper(trim(COALESCE(veiculo,'')))
                    FROM programacoes
                    WHERE trim(COALESCE(veiculo,'')) <> ''
                    ORDER BY 1
                """)
                vals = [r[0] for r in (cur.fetchall() or []) if str(r[0] or "").strip()]
            except Exception:
                vals = []
        values = ["TODOS"] + vals
        self.cb_veiculo.configure(values=values)
        if self.var_veiculo.get() not in values:
            self.var_veiculo.set("TODOS")

    def refresh_data(self):
        periodo = upper(self.var_periodo.get().strip())
        veiculo_sel = upper(self.var_veiculo.get().strip())
        cutoff = None
        if periodo != "TODAS":
            try:
                cutoff = datetime.now() - timedelta(days=int(periodo))
            except Exception:
                cutoff = None

        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/centro-custos/rows?limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("rows") if isinstance(resp, dict) else []
                rows = [
                    (
                        (r or {}).get("codigo_programacao", ""),
                        (r or {}).get("veiculo", ""),
                        (r or {}).get("data_ref", ""),
                        (r or {}).get("km_rodado", 0),
                        (r or {}).get("kg_carregado", 0),
                        (r or {}).get("total_desp", 0),
                    )
                    for r in (arr or [])
                    if isinstance(r, dict)
                ]
            except Exception:
                logging.debug("Falha ao carregar centro de custos via API", exc_info=True)

        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

                if "data_saida" in cols and "data_criacao" in cols:
                    data_expr = "COALESCE(data_saida,data_criacao,'')"
                elif "data_saida" in cols:
                    data_expr = "COALESCE(data_saida,'')"
                elif "data_criacao" in cols:
                    data_expr = "COALESCE(data_criacao,'')"
                elif "data" in cols:
                    data_expr = "COALESCE(data,'')"
                else:
                    data_expr = "''"
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                kg_expr = "COALESCE(nf_kg_carregado, kg_carregado, 0)" if "nf_kg_carregado" in cols else ("COALESCE(kg_carregado,0)" if "kg_carregado" in cols else "0")
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                has_despesas = bool(cur.fetchone())
                despesas_join = (
                    """
                 LEFT JOIN (
                        SELECT codigo_programacao, COALESCE(SUM(valor),0) as total_desp
                          FROM despesas
                         GROUP BY codigo_programacao
                    ) d ON d.codigo_programacao = p.codigo_programacao
                    """
                    if has_despesas
                    else ""
                )
                total_desp_expr = "COALESCE(d.total_desp,0)" if has_despesas else "0"

                cur.execute(
                    f"""
                    SELECT p.codigo_programacao,
                           upper(trim(COALESCE(p.veiculo,''))) as veiculo,
                           {data_expr} as data_ref,
                           {km_expr} as km_rodado,
                           {kg_expr} as kg_carregado,
                           {total_desp_expr} as total_desp
                      FROM programacoes p
                    {despesas_join}
                     WHERE trim(COALESCE(p.veiculo,'')) <> ''
                    """
                )
                rows = cur.fetchall() or []

        agg = {}
        for r in rows:
            veiculo = upper(str(r[1] or "").strip())
            dt_ref = self._parse_data(r[2])
            km = safe_float(r[3], 0.0)
            kg = safe_float(r[4], 0.0)
            desp = safe_float(r[5], 0.0)

            if cutoff is not None:
                if dt_ref is None or dt_ref < cutoff:
                    continue
            if veiculo_sel and veiculo_sel != "TODOS" and veiculo != veiculo_sel:
                continue
            if not veiculo:
                continue

            d = agg.setdefault(veiculo, {"rotas": 0, "km": 0.0, "kg": 0.0, "desp": 0.0})
            d["rotas"] += 1
            d["km"] += km
            d["kg"] += kg
            d["desp"] += desp

        self.tree.delete(*self.tree.get_children())
        rows_out = []
        for veic, d in sorted(agg.items(), key=lambda kv: (-safe_float(kv[1].get("desp", 0), 0.0), kv[0])):
            rotas = safe_int(d["rotas"], 0)
            km = safe_float(d["km"], 0.0)
            kg = safe_float(d["kg"], 0.0)
            desp = safe_float(d["desp"], 0.0)
            custo_km = (desp / km) if km > 0 else 0.0
            custo_kg = (desp / kg) if kg > 0 else 0.0
            ticket = (desp / rotas) if rotas > 0 else 0.0
            rows_out.append((veic, rotas, km, kg, desp, custo_km, custo_kg, ticket))

        media_custo_km = (
            sum(r[5] for r in rows_out) / float(len(rows_out))
            if rows_out else 0.0
        )
        for veic, rotas, km, kg, desp, custo_km, custo_kg, ticket in rows_out:
            if media_custo_km <= 0:
                tag = "medio"
            elif custo_km > media_custo_km * 1.2:
                tag = "alto"
            elif custo_km < media_custo_km * 0.9:
                tag = "bom"
            else:
                tag = "medio"
            tree_insert_aligned(
                self.tree, "", "end",
                (
                    veic,
                    rotas,
                    round(km, 1),
                    round(kg, 2),
                    round(desp, 2),
                    round(custo_km, 3),
                    round(custo_kg, 3),
                    round(ticket, 2),
                ),
                tags=(tag,),
            )

        total_rotas = sum(r[1] for r in rows_out)
        total_km = sum(r[2] for r in rows_out)
        total_kg = sum(r[3] for r in rows_out)
        total_desp = sum(r[4] for r in rows_out)
        custo_km_global = (total_desp / total_km) if total_km > 0 else 0.0
        custo_kg_global = (total_desp / total_kg) if total_kg > 0 else 0.0
        metric = upper(self.var_chart_metric.get().strip())
        if metric == "CUSTO_KG":
            idx = 6
            title = "Custo por Veiculo (Custo/KG)"
        elif metric == "DESPESA_TOTAL":
            idx = 4
            title = "Custo por Veiculo (Despesa Total)"
        else:
            idx = 5
            title = "Custo por Veiculo (Custo/KM)"
        self.lbl_chart_title.config(text=title)
        chart_rows = sorted(rows_out, key=lambda r: safe_float(r[idx], 0.0), reverse=True)
        self._chart_labels = [r[0] for r in chart_rows[:10]]
        self._chart_values = [r[idx] for r in chart_rows[:10]]
        self._draw_chart(self._chart_labels, self._chart_values)

        self.lbl_resumo.config(
            text=(
                f"Veiculos: {len(rows_out)} | Rotas: {total_rotas} | "
                f"KM: {total_km:.1f} | KG carregado: {total_kg:.2f} | "
                f"Despesas: R$ {total_desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                + f" | Custo/KM global: {custo_km_global:.3f} | Custo/KG global: {custo_kg_global:.3f}"
            )
        )
        self.set_status(f"STATUS: Centro de Custos atualizado ({self.var_periodo.get()} dias / veiculo {self.var_veiculo.get()}).")


__all__ = ["CentroCustosPage", "configure_centro_custos_page_dependencies"]
