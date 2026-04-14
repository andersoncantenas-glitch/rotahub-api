# -*- coding: utf-8 -*-
import logging
import sqlite3
import tkinter as tk
from tkinter import ttk


def _align_values_to_tree_columns(tree: ttk.Treeview, values):
    try:
        cols = list(tree.cget("columns") or [])
    except Exception:
        cols = []
    ncols = len(cols)

    if isinstance(values, (list, tuple)):
        vals = list(values)
    elif isinstance(values, sqlite3.Row):
        # sqlite3.Row precisa virar lista de valores; caso contrario cai como objeto unico.
        vals = list(values)
    elif values is None:
        vals = []
    else:
        vals = [values]

    if ncols > 0:
        if len(vals) < ncols:
            vals.extend([""] * (ncols - len(vals)))
        elif len(vals) > ncols:
            vals = vals[:ncols]
    return tuple(vals)


def tree_insert_aligned(tree: ttk.Treeview, parent: str, index: str, values, **kwargs):
    aligned = _align_values_to_tree_columns(tree, values)
    # Zebra leve + melhor leitura visual entre linhas
    try:
        if not getattr(tree, "_zebra_ready", False):
            tree.tag_configure("row_even", background="#FFFFFF")
            tree.tag_configure("row_odd", background="#F8FAFC")
            tree._zebra_ready = True
        existing_tags = list(kwargs.get("tags", ()))
        pos = len(tree.get_children(parent or ""))
        zebra_tag = "row_even" if (pos % 2 == 0) else "row_odd"
        if zebra_tag not in existing_tags:
            existing_tags.append(zebra_tag)
        kwargs["tags"] = tuple(existing_tags)
    except Exception:
        logging.debug("Falha ao aplicar estilo zebrado na tabela")
    return tree.insert(parent, index, values=aligned, **kwargs)


# =========================================================
# ORDENACAO UNIVERSAL PARA TREEVIEW (CLIQUE NO CABECALHO)
# =========================================================
def enable_treeview_sorting(tree: ttk.Treeview, numeric_cols=None, money_cols=None, date_cols=None):
    """
    Habilita recursos de tabela no Treeview:
    - Ordenacao por cabecalho
    - Indicador de filtro ao passar mouse no cabecalho
    - Filtro por coluna (valores especificos)

    Uso do filtro:
    - Passe o mouse no cabecalho para ver o indicador "\u23F7"
    - Clique no indicador (lado direito do cabecalho) ou botao direito no cabecalho
    """
    numeric_cols = set(numeric_cols or [])
    money_cols = set(money_cols or [])
    date_cols = set(date_cols or [])

    if not hasattr(tree, "_sort_state"):
        tree._sort_state = {"col": None, "reverse": False}

    if not hasattr(tree, "_filter_state"):
        tree._filter_state = {}

    if not hasattr(tree, "_filter_all_iids"):
        tree._filter_all_iids = list(tree.get_children(""))

    if not hasattr(tree, "_base_heading_text"):
        tree._base_heading_text = {
            c: (tree.heading(c, "text") or c)
            for c in tree["columns"]
        }

    tree._hover_filter_col = None

    def _clean_header_text(t: str) -> str:
        t = (t or "").strip()
        for suffix in (" \u2191", " \u2193", " \u23F7", " \u23F7*"):
            if t.endswith(suffix):
                t = t[: -len(suffix)].strip()
        return t

    def _to_float(v):
        try:
            if v is None:
                return 0.0
            s = str(v).strip()
            if not s or s in {"", "-", "None"}:
                return 0.0

            s = s.replace("R$", "").replace(" ", "")

            neg = False
            if s.startswith("(") and s.endswith(")"):
                neg = True
                s = s[1:-1].strip()

            s = s.replace(".", "").replace(",", ".")
            val = float(s)
            return -val if neg else val
        except Exception:
            return 0.0

    def _to_date_key(v):
        if v is None:
            return (0, 0, 0)

        s = str(v).strip()
        if not s or s in {"", "-", "None"}:
            return (0, 0, 0)

        if "-" in s and len(s) >= 10:
            try:
                y, m, d = s[:10].split("-")
                return (int(y), int(m), int(d))
            except Exception:
                logging.debug("Falha ignorada")

        if "/" in s and len(s) >= 10:
            try:
                d, m, y = s[:10].split("/")
                return (int(y), int(m), int(d))
            except Exception:
                logging.debug("Falha ignorada")

        return (0, 0, 0)

    def _format_header(col):
        base = _clean_header_text(tree._base_heading_text.get(col) or tree.heading(col, "text") or col)
        suffix = ""

        if tree._sort_state.get("col") == col:
            suffix += " \u2193" if tree._sort_state.get("reverse") else " \u2191"

        if tree._hover_filter_col == col:
            suffix += " \u23F7*" if col in tree._filter_state else " \u23F7"
        elif col in tree._filter_state:
            suffix += " \u23F7*"

        return f"{base}{suffix}"

    def _refresh_headers():
        for c in tree["columns"]:
            tree.heading(c, text=_format_header(c))

    def _value_for_compare(v):
        return str(v or "").strip()

    def _row_matches_filters(iid):
        for col, allowed in tree._filter_state.items():
            current = _value_for_compare(tree.set(iid, col))
            if current not in allowed:
                return False
        return True

    def _apply_filters():
        all_iids = [iid for iid in tree._filter_all_iids if tree.exists(iid)]
        if not tree._filter_state:
            for idx, iid in enumerate(all_iids):
                tree.reattach(iid, "", idx)
            _refresh_headers()
            return

        pos = 0
        for iid in all_iids:
            if _row_matches_filters(iid):
                tree.reattach(iid, "", pos)
                pos += 1
            else:
                tree.detach(iid)
        _refresh_headers()

    def _iter_values_for_col(col):
        vals = []
        for iid in tree._filter_all_iids:
            if tree.exists(iid):
                vals.append(_value_for_compare(tree.set(iid, col)))
        return vals

    def _open_filter_popup(col):
        top = tk.Toplevel(tree)
        top.title(f"Filtrar: {col}")
        top.transient(tree.winfo_toplevel())
        top.resizable(False, False)
        top.grab_set()

        frm = ttk.Frame(top, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.grid_columnconfigure(0, weight=1)

        ttk.Label(frm, text=f"Coluna: {col}", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")

        search_var = tk.StringVar(value="")
        ent = ttk.Entry(frm, textvariable=search_var)
        ent.grid(row=1, column=0, sticky="ew", pady=(6, 6))

        listbox = tk.Listbox(frm, selectmode="multiple", exportselection=False, height=12)
        listbox.grid(row=2, column=0, sticky="nsew")

        unique_vals = sorted(set(_iter_values_for_col(col)), key=lambda s: s.upper())
        current_allowed = set(tree._filter_state.get(col, set()))

        def _fill_list():
            needle = search_var.get().strip().upper()
            listbox.delete(0, "end")
            for v in unique_vals:
                show = "(vazio)" if v == "" else v
                if needle and needle not in show.upper():
                    continue
                listbox.insert("end", show)
                if (v in current_allowed) or (not current_allowed):
                    listbox.selection_set("end")

        def _selected_values():
            selected = set()
            for idx in listbox.curselection():
                txt = listbox.get(idx)
                selected.add("" if txt == "(vazio)" else txt)
            return selected

        def _apply_and_close():
            vals = _selected_values()
            if vals and len(vals) < len(unique_vals):
                tree._filter_state[col] = vals
            else:
                tree._filter_state.pop(col, None)
            _apply_filters()
            top.destroy()

        def _clear_col_filter():
            tree._filter_state.pop(col, None)
            _apply_filters()
            top.destroy()

        def _clear_all_filters():
            tree._filter_state.clear()
            _apply_filters()
            top.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="Aplicar", style="Primary.TButton", command=_apply_and_close).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btns, text="Limpar coluna", style="Ghost.TButton", command=_clear_col_filter).grid(row=0, column=1, padx=6)
        ttk.Button(btns, text="Limpar tudo", style="Ghost.TButton", command=_clear_all_filters).grid(row=0, column=2, padx=6)

        search_var.trace_add("write", lambda *_: _fill_list())
        _fill_list()
        ent.focus_set()

    def sort_by(col):
        if tree._sort_state["col"] == col:
            tree._sort_state["reverse"] = not tree._sort_state["reverse"]
        else:
            tree._sort_state["col"] = col
            tree._sort_state["reverse"] = False

        reverse = tree._sort_state["reverse"]
        data = [(tree.set(iid, col), iid) for iid in tree.get_children("")]

        if col in money_cols or col in numeric_cols:
            data.sort(key=lambda x: _to_float(x[0]), reverse=reverse)
        elif col in date_cols:
            data.sort(key=lambda x: _to_date_key(x[0]), reverse=reverse)
        else:
            data.sort(key=lambda x: str(x[0]).strip().upper(), reverse=reverse)

        for idx, (_, iid) in enumerate(data):
            tree.move(iid, "", idx)

        _refresh_headers()

    def _column_right_edge(col_name):
        x = 0
        for c in tree["columns"]:
            w = int(tree.column(c, "width") or 0)
            x += w
            if c == col_name:
                return x
        return None

    def _on_motion(event):
        region = tree.identify("region", event.x, event.y)
        col_id = tree.identify_column(event.x)
        hover_col = None
        if region == "heading" and col_id and col_id.startswith("#"):
            try:
                idx = int(col_id[1:]) - 1
                if 0 <= idx < len(tree["columns"]):
                    hover_col = tree["columns"][idx]
            except Exception:
                hover_col = None

        if tree._hover_filter_col != hover_col:
            tree._hover_filter_col = hover_col
            _refresh_headers()

    def _on_leave(_event):
        if tree._hover_filter_col is not None:
            tree._hover_filter_col = None
            _refresh_headers()

    def _on_right_click(event):
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return
        col_id = tree.identify_column(event.x)
        if not col_id.startswith("#"):
            return
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(tree["columns"]):
            return
        col = tree["columns"][idx]
        _open_filter_popup(col)
        return "break"

    def _on_left_click(event):
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return
        col_id = tree.identify_column(event.x)
        if not col_id.startswith("#"):
            return
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(tree["columns"]):
            return
        col = tree["columns"][idx]

        col_right = _column_right_edge(col)
        if col_right is None:
            return

        if event.x >= (col_right - 20):
            _open_filter_popup(col)
            return "break"

    if not hasattr(tree, "_filter_wrapped_io"):
        tree._filter_wrapped_io = True
        _orig_insert = tree.insert
        _orig_delete = tree.delete

        def _insert_wrapped(parent, index, iid=None, **kw):
            new_iid = _orig_insert(parent, index, iid=iid, **kw)
            if parent == "" and new_iid not in tree._filter_all_iids:
                tree._filter_all_iids.append(new_iid)
            return new_iid

        def _delete_wrapped(*items):
            for iid in items:
                try:
                    tree._filter_all_iids.remove(iid)
                except Exception:
                    pass
            return _orig_delete(*items)

        tree.insert = _insert_wrapped
        tree.delete = _delete_wrapped

    for c in tree["columns"]:
        tree.heading(c, command=lambda col=c: sort_by(col))

    tree.bind("<Motion>", _on_motion, add="+")
    tree.bind("<Leave>", _on_leave, add="+")
    tree.bind("<Button-3>", _on_right_click, add="+")
    tree.bind("<Button-1>", _on_left_click, add="+")

    _refresh_headers()


__all__ = [
    "enable_treeview_sorting",
    "tree_insert_aligned",
    "_align_values_to_tree_columns",
]
