# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from app.services import saas_admin_service
from app.ui.components.page_base import PageBase


class SaaSAdminPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "SaaSAdmin", title="Admin SaaS")
        self.company_id = None
        self.plans_by_name = {}
        self._build()

    def on_show(self):
        if not getattr(self.app, "user", {}).get("is_admin", False):
            messagebox.showerror(
                "Acesso negado",
                "A função Admin SaaS está disponível somente para o administrador do sistema.",
            )
            return
        self.set_status("STATUS: Administracao SaaS, plano, cobranca e auditoria.")
        self.refresh_data()

    def _build(self):
        self.header_right.grid_columnconfigure(0, weight=0)
        ttk.Button(self.header_right, text="Atualizar", style="Ghost.TButton", command=self.refresh_data).grid(row=0, column=0, sticky="e")

        shell = ttk.Frame(self.body, style="Content.TFrame")
        shell.grid(row=0, column=0, sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        summary = ttk.Frame(shell, style="Content.TFrame")
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        for idx in range(4):
            summary.grid_columnconfigure(idx, weight=1, uniform="saas_summary")

        self.lbl_company = self._summary_box(summary, 0, "Empresa", "-")
        self.lbl_plan = self._summary_box(summary, 1, "Plano", "-")
        self.lbl_vehicles = self._summary_box(summary, 2, "Veiculos", "-")
        self.lbl_billing = self._summary_box(summary, 3, "Cobranca", "-")

        self.nb = ttk.Notebook(shell)
        self.nb.grid(row=1, column=0, sticky="nsew")
        self._build_plan_tab()
        self._build_payments_tab()
        self._build_audit_tab()

    def _summary_box(self, parent, col, title, value):
        box = ttk.LabelFrame(parent, text=title, padding=10)
        box.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 0))
        label = ttk.Label(box, text=value, style="CardTitle.TLabel")
        label.pack(anchor="w")
        return label

    def _build_plan_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)
        self.nb.add(tab, text="Empresa e Plano")

        left = ttk.LabelFrame(tab, text="Plano", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.grid_columnconfigure(1, weight=1)

        ttk.Label(left, text="Plano atual:").grid(row=0, column=0, sticky="w", pady=4)
        self.lbl_plan_detail = ttk.Label(left, text="-")
        self.lbl_plan_detail.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(left, text="Novo plano:").grid(row=1, column=0, sticky="w", pady=4)
        self.var_plan = tk.StringVar()
        self.cb_plan = ttk.Combobox(left, textvariable=self.var_plan, state="readonly")
        self.cb_plan.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(left, text="Motivo:").grid(row=2, column=0, sticky="w", pady=4)
        self.ent_plan_reason = ttk.Entry(left)
        self.ent_plan_reason.grid(row=2, column=1, sticky="ew", pady=4)

        ttk.Button(left, text="Aplicar plano", style="Primary.TButton", command=self.change_plan).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        right = ttk.LabelFrame(tab, text="Status e Uso", padding=10)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)

        self.tree_usage = ttk.Treeview(right, columns=("metric", "valor"), show="headings", height=8)
        self.tree_usage.heading("metric", text="Indicador")
        self.tree_usage.heading("valor", text="Valor")
        self.tree_usage.column("metric", width=140, anchor="w")
        self.tree_usage.column("valor", width=180, anchor="w")
        self.tree_usage.grid(row=0, column=0, sticky="nsew")

        btns = ttk.Frame(right)
        btns.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)
        ttk.Button(btns, text="Ativar", command=lambda: self.set_company_status("active")).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(btns, text="Suspender", style="Warn.TButton", command=lambda: self.set_company_status("suspended")).grid(row=0, column=1, sticky="ew", padx=(4, 0))

    def _build_payments_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        self.nb.add(tab, text="Pagamentos")

        form = ttk.LabelFrame(tab, text="Nova cobranca", padding=10)
        form.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        form.grid_columnconfigure(5, weight=1)

        ttk.Label(form, text="Valor").grid(row=0, column=0, sticky="w")
        self.ent_amount = ttk.Entry(form, width=12)
        self.ent_amount.grid(row=0, column=1, sticky="w", padx=(4, 12))
        ttk.Label(form, text="Vencimento").grid(row=0, column=2, sticky="w")
        self.ent_due = ttk.Entry(form, width=14)
        self.ent_due.grid(row=0, column=3, sticky="w", padx=(4, 12))
        ttk.Label(form, text="Obs").grid(row=0, column=4, sticky="w")
        self.ent_payment_notes = ttk.Entry(form)
        self.ent_payment_notes.grid(row=0, column=5, sticky="ew", padx=(4, 12))
        ttk.Button(form, text="Criar", style="Primary.TButton", command=self.create_payment).grid(row=0, column=6, sticky="e")

        list_box = ttk.LabelFrame(tab, text="Pagamentos", padding=8)
        list_box.grid(row=1, column=0, sticky="nsew")
        list_box.grid_columnconfigure(0, weight=1)
        list_box.grid_rowconfigure(0, weight=1)

        cols = ("id", "status", "amount", "due", "paid", "reference")
        self.tree_payments = ttk.Treeview(list_box, columns=cols, show="headings", height=12)
        headings = {
            "id": "ID",
            "status": "Status",
            "amount": "Valor",
            "due": "Vencimento",
            "paid": "Pago em",
            "reference": "Referencia",
        }
        for col, text in headings.items():
            self.tree_payments.heading(col, text=text)
            self.tree_payments.column(col, width=100, anchor="w")
        self.tree_payments.column("id", width=60, anchor="center")
        self.tree_payments.grid(row=0, column=0, sticky="nsew")
        payments_vsb = ttk.Scrollbar(list_box, orient="vertical", command=self.tree_payments.yview)
        payments_vsb.grid(row=0, column=1, sticky="ns")
        self.tree_payments.configure(yscrollcommand=payments_vsb.set)

        pay_btns = ttk.Frame(tab)
        pay_btns.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(pay_btns, text="Registrar pagamento selecionado", command=self.register_selected_payment).pack(side="left")
        ttk.Button(pay_btns, text="Rodar suspensao por atraso", style="Warn.TButton", command=self.run_overdue_check).pack(side="left", padx=(8, 0))

    def _build_audit_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        self.nb.add(tab, text="Auditoria")

        cols = ("id", "action", "severity", "entity", "created")
        self.tree_audit = ttk.Treeview(tab, columns=cols, show="headings", height=16)
        for col, text in {
            "id": "ID",
            "action": "Acao",
            "severity": "Severidade",
            "entity": "Entidade",
            "created": "Criado em",
        }.items():
            self.tree_audit.heading(col, text=text)
            self.tree_audit.column(col, width=120, anchor="w")
        self.tree_audit.column("id", width=60, anchor="center")
        self.tree_audit.grid(row=0, column=0, sticky="nsew")
        audit_vsb = ttk.Scrollbar(tab, orient="vertical", command=self.tree_audit.yview)
        audit_vsb.grid(row=0, column=1, sticky="ns")
        self.tree_audit.configure(yscrollcommand=audit_vsb.set)

    def refresh_data(self):
        res = saas_admin_service.get_dashboard(self.company_id)
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha ao carregar dados.")
            return
        data = res.get("data") or {}
        company = data.get("company") or {}
        subscription = data.get("subscription") or {}
        usage = data.get("usage") or {}
        vehicles = (usage.get("vehicles") or {})
        self.company_id = int(company.get("id") or 0)

        self.lbl_company.config(text=f"{company.get('name') or '-'}\n{company.get('status') or '-'}")
        self.lbl_plan.config(text=f"{subscription.get('plan_name') or '-'}\n{subscription.get('status') or '-'}")
        limit = vehicles.get("vehicle_limit")
        limit_txt = "sem limite" if limit is None else str(limit)
        self.lbl_vehicles.config(text=f"{vehicles.get('vehicle_count', 0)} / {limit_txt}")
        self.lbl_billing.config(text=f"Venc.: {subscription.get('next_due_date') or '-'}")
        self.lbl_plan_detail.config(text=f"{subscription.get('plan_name') or '-'} ({subscription.get('plan_code') or '-'})")

        self.plans_by_name = {}
        plan_names = []
        for plan in data.get("plans") or []:
            label = f"{plan.get('name')} ({plan.get('code')})"
            self.plans_by_name[label] = plan
            plan_names.append(label)
        self.cb_plan.configure(values=plan_names)
        if plan_names and not self.var_plan.get():
            self.var_plan.set(plan_names[0])

        self._fill_usage(usage)
        self._fill_payments(data.get("payments") or [])
        self._fill_audit(data.get("audit_logs") or [])

    def _fill_usage(self, usage):
        for item in self.tree_usage.get_children():
            self.tree_usage.delete(item)
        rows = [
            ("Veiculos", (usage.get("vehicles") or {}).get("vehicle_count", 0)),
            ("Usuarios", usage.get("users", 0)),
            ("Motoristas", usage.get("motoristas", 0)),
            ("Vendedores", usage.get("vendedores", 0)),
            ("Clientes", usage.get("clientes", 0)),
            ("Programacoes", usage.get("programacoes", 0)),
        ]
        for name, value in rows:
            self.tree_usage.insert("", "end", values=(name, value))

    def _fill_payments(self, payments):
        for item in self.tree_payments.get_children():
            self.tree_payments.delete(item)
        for payment in payments:
            self.tree_payments.insert(
                "",
                "end",
                iid=str(payment.get("id")),
                values=(
                    payment.get("id"),
                    payment.get("status") or "",
                    f"{float(payment.get('amount') or 0):.2f}",
                    payment.get("due_date") or "",
                    payment.get("paid_at") or "",
                    payment.get("reference") or "",
                ),
            )

    def _fill_audit(self, logs):
        for item in self.tree_audit.get_children():
            self.tree_audit.delete(item)
        for log in logs:
            entity = f"{log.get('entity_type') or ''}:{log.get('entity_id') or ''}".strip(":")
            self.tree_audit.insert(
                "",
                "end",
                values=(
                    log.get("id"),
                    log.get("action") or "",
                    log.get("severity") or "",
                    entity,
                    log.get("created_at") or "",
                ),
            )

    def change_plan(self):
        if not self.company_id:
            return
        plan = self.plans_by_name.get(self.var_plan.get())
        if not plan:
            messagebox.showwarning("Admin SaaS", "Selecione um plano.")
            return
        actor = (self.app.user or {}).get("nome", "ADMIN") if self.app else "ADMIN"
        res = saas_admin_service.change_company_plan(
            self.company_id,
            str(plan.get("code") or ""),
            actor=actor,
            reason=self.ent_plan_reason.get(),
        )
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha ao alterar plano.")
            return
        messagebox.showinfo("Admin SaaS", "Plano alterado.")
        self.refresh_data()

    def set_company_status(self, status):
        if not self.company_id:
            return
        actor = (self.app.user or {}).get("nome", "ADMIN") if self.app else "ADMIN"
        res = saas_admin_service.set_company_status(self.company_id, status, actor=actor)
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha ao alterar status.")
            return
        self.refresh_data()

    def create_payment(self):
        if not self.company_id:
            return
        try:
            amount = float(str(self.ent_amount.get() or "0").replace(",", "."))
        except Exception:
            messagebox.showwarning("Admin SaaS", "Valor invalido.")
            return
        res = saas_admin_service.create_payment(
            self.company_id,
            amount,
            self.ent_due.get(),
            notes=self.ent_payment_notes.get(),
        )
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha ao criar cobranca.")
            return
        self.ent_amount.delete(0, "end")
        self.ent_payment_notes.delete(0, "end")
        self.refresh_data()

    def register_selected_payment(self):
        selected = self.tree_payments.selection()
        if not selected:
            messagebox.showwarning("Admin SaaS", "Selecione um pagamento.")
            return
        payment_id = int(selected[0])
        actor = (self.app.user or {}).get("nome", "ADMIN") if self.app else "ADMIN"
        res = saas_admin_service.register_payment(payment_id, actor=actor)
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha ao registrar pagamento.")
            return
        self.refresh_data()

    def run_overdue_check(self):
        res = saas_admin_service.run_overdue_check()
        if not res.get("ok"):
            messagebox.showerror("Admin SaaS", res.get("error") or "Falha na automacao.")
            return
        data = res.get("data") or {}
        messagebox.showinfo("Admin SaaS", f"Assinaturas suspensas: {data.get('suspended', 0)}")
        self.refresh_data()


__all__ = ["SaaSAdminPage"]
