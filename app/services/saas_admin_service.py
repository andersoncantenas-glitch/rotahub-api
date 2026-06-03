from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path

from app.repositories import audit_repository, company_repository, payment_repository, plan_repository, subscription_repository
from app.repositories.base_repository import ensure_saas_ready, get_db
from app.services.billing_automation_service import suspend_overdue_subscriptions
from app.services.saas_result import error_message, service_result
from app.services.usage_service import get_company_usage
from app.services.vehicle_limit_service import vehicle_usage_snapshot


def get_dashboard(company_id: int | None = None) -> dict:
    try:
        company = company_repository.get_company(company_id) if company_id else company_repository.get_default_company()
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        cid = int(company["id"])
        subscription = subscription_repository.get_active_subscription(cid)
        usage = get_company_usage(cid)
        payments = payment_repository.list_payments(company_id=cid, limit=20)
        audits = audit_repository.list_audit_logs(company_id=cid, limit=50)
        return service_result(
            ok=True,
            data={
                "company": company,
                "subscription": subscription,
                "usage": usage.get("data") if isinstance(usage, dict) else {},
                "payments": payments,
                "audit_logs": audits,
                "plans": plan_repository.list_plans(include_inactive=False, limit=100),
            },
        )
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao carregar dashboard SaaS."))


def change_company_plan(company_id: int, plan_code: str, *, actor: str = "ADMIN", reason: str = "") -> dict:
    try:
        company = company_repository.get_company(company_id)
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        plan = plan_repository.get_plan_by_code(plan_code)
        if not plan:
            return service_result(ok=False, data=None, error="Plano nao encontrado.")
        with get_db() as conn:
            ensure_saas_ready(conn)
            usage = vehicle_usage_snapshot(conn, int(company_id))
        vehicle_count = int((usage or {}).get("vehicle_count") or 0)
        vehicle_limit = plan.get("vehicle_limit")
        usage_result = get_company_usage(company_id)
        usage_data = (usage_result.get("data") if isinstance(usage_result, dict) else {}) or {}
        user_count = int(usage_data.get("users") or 0)
        user_limit = plan.get("user_limit")
        if vehicle_limit is not None and vehicle_count > int(vehicle_limit):
            return service_result(
                ok=False,
                data={"usage": usage, "plan": plan},
                error=(
                    f"Downgrade bloqueado: empresa possui {vehicle_count} veiculos, "
                    f"mas o plano {plan.get('name')} permite {int(vehicle_limit)}."
                ),
            )
        if user_limit is not None and user_count > int(user_limit):
            return service_result(
                ok=False,
                data={"usage": usage_data, "plan": plan},
                error=(
                    f"Downgrade bloqueado: empresa possui {user_count} usuarios, "
                    f"mas o plano {plan.get('name')} permite {int(user_limit)}."
                ),
            )
        subscription = subscription_repository.change_company_plan(company_id, int(plan["id"]))
        audit_repository.create_audit_log(
            {
                "company_id": int(company_id),
                "actor_type": "admin",
                "action": "plano_alterado",
                "entity_type": "subscription",
                "entity_id": str(subscription.get("id") or ""),
                "severity": "info",
                "metadata": {
                    "actor": actor,
                    "new_plan_code": plan.get("code"),
                    "reason": reason,
                    "vehicle_count": vehicle_count,
                    "user_count": user_count,
                },
            }
        )
        return service_result(ok=True, data=subscription)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao alterar plano."))


def set_company_status(company_id: int, status: str, *, actor: str = "ADMIN", reason: str = "") -> dict:
    try:
        company = company_repository.update_company(company_id, {"status": str(status or "").strip().lower()})
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")
        audit_repository.create_audit_log(
            {
                "company_id": int(company_id),
                "actor_type": "admin",
                "action": "empresa_status_alterado",
                "entity_type": "company",
                "entity_id": str(company_id),
                "severity": "warning",
                "metadata": {"actor": actor, "status": status, "reason": reason},
            }
        )
        return service_result(ok=True, data=company)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao alterar status."))


def create_payment(company_id: int, amount: float, due_date: str = "", *, notes: str = "") -> dict:
    try:
        subscription = subscription_repository.get_active_subscription(company_id)
        payment = payment_repository.create_payment(
            {
                "company_id": int(company_id),
                "subscription_id": (subscription or {}).get("id"),
                "amount": float(amount or 0),
                "due_date": str(due_date or "").strip() or None,
                "notes": notes,
            }
        )
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao criar pagamento."))


def generate_boleto(payment_id: int, *, actor: str = "ADMIN") -> dict:
    try:
        payment = payment_repository.get_payment(payment_id)
        if not payment:
            return service_result(ok=False, data=None, error="Pagamento nao encontrado.")
        if str(payment.get("status") or "").strip().lower() == "paid":
            return service_result(ok=False, data=None, error="Nao e possivel gerar boleto para pagamento ja quitado.")
        company = company_repository.get_company(int(payment.get("company_id") or 0)) or {}
        if not company:
            return service_result(ok=False, data=None, error="Empresa nao encontrada.")

        generated_at = datetime.now().replace(microsecond=0)
        our_number = _boleto_our_number(payment)
        digitable_line = _boleto_digitable_line(our_number, float(payment.get("amount") or 0), str(payment.get("due_date") or ""))
        pdf_url = f"/owner/generated/boletos/boleto-{int(payment_id):06d}.pdf"
        pdf_path = _owner_boleto_dir() / f"boleto-{int(payment_id):06d}.pdf"
        _render_boleto_pdf(
            pdf_path,
            payment=payment,
            company=company,
            our_number=our_number,
            digitable_line=digitable_line,
            generated_at=generated_at,
            pdf_url=pdf_url,
        )
        updated = payment_repository.update_boleto(
            payment_id,
            {
                "reference": our_number,
                "boleto_our_number": our_number,
                "boleto_digitable_line": digitable_line,
                "boleto_pdf_url": pdf_url,
                "boleto_pdf_path": str(pdf_path),
                "boleto_generated_at": generated_at.isoformat(sep=" "),
            },
        )
        if not updated:
            return service_result(ok=False, data=None, error="Pagamento nao encontrado.")
        audit_repository.create_audit_log(
            {
                "company_id": int(updated.get("company_id") or 0),
                "actor_type": "admin",
                "action": "boleto_gerado",
                "entity_type": "payment",
                "entity_id": str(payment_id),
                "severity": "info",
                "metadata": {"actor": actor, "reference": our_number, "pdf_url": pdf_url},
            }
        )
        return service_result(ok=True, data=updated)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao gerar boleto."))


def register_payment(payment_id: int, *, method: str = "manual", reference: str = "", notes: str = "", actor: str = "ADMIN") -> dict:
    try:
        payment = payment_repository.register_payment(payment_id, method=method, reference=reference, notes=notes)
        if not payment:
            return service_result(ok=False, data=None, error="Pagamento nao encontrado.")
        audit_repository.create_audit_log(
            {
                "company_id": int(payment.get("company_id") or 0),
                "actor_type": "admin",
                "action": "pagamento_registrado",
                "entity_type": "payment",
                "entity_id": str(payment.get("id") or payment_id),
                "severity": "info",
                "metadata": {"actor": actor, "method": method, "reference": reference},
            }
        )
        return service_result(ok=True, data=payment)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao registrar pagamento."))


def run_overdue_check(grace_days: int = 0) -> dict:
    return suspend_overdue_subscriptions(grace_days=int(grace_days or 0))


def format_features(features_json: str | None) -> str:
    try:
        data = json.loads(str(features_json or "{}"))
    except Exception:
        data = {}
    if not isinstance(data, dict):
        return ""
    enabled = [key for key, value in sorted(data.items()) if bool(value)]
    return ", ".join(enabled)


def _owner_boleto_dir() -> Path:
    path = Path(__file__).resolve().parents[2] / "backend" / "web_owner" / "generated" / "boletos"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _boleto_our_number(payment: dict) -> str:
    payment_id = int(payment.get("id") or 0)
    company_id = int(payment.get("company_id") or 0)
    seed = f"ROT-{company_id:04d}-{payment_id:08d}"
    check = int(hashlib.sha1(seed.encode("ascii", "ignore")).hexdigest()[:6], 16) % 97
    return f"RH{company_id:04d}{payment_id:08d}{check:02d}"


def _only_digits(value: object) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _boleto_digitable_line(our_number: str, amount: float, due_date: str) -> str:
    amount_cents = max(0, int(round(float(amount or 0) * 100)))
    due_digits = _only_digits(due_date)[-8:].rjust(8, "0")
    base = (_only_digits(our_number) + f"{amount_cents:010d}" + due_digits).ljust(44, "0")[:44]
    checksum = sum((idx + 2) * int(ch) for idx, ch in enumerate(base)) % 10
    full = base + str(checksum)
    return " ".join([full[0:5], full[5:15], full[15:25], full[25:35], full[35:45]])


def _render_boleto_pdf(
    pdf_path: Path,
    *,
    payment: dict,
    company: dict,
    our_number: str,
    digitable_line: str,
    generated_at: datetime,
    pdf_url: str,
) -> None:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    pdf = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4
    left = 18 * mm
    top = height - 18 * mm

    def draw(label: str, value: object, y: float, *, bold: bool = False, size: int = 10) -> None:
        pdf.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        pdf.drawString(left, y, f"{label}: {str(value or '-')[:110]}")

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(left, top, "RotaHub - Boleto de cobranca")
    pdf.setFont("Helvetica", 9)
    pdf.drawRightString(width - left, top, f"Gerado em {generated_at.strftime('%d/%m/%Y %H:%M')}")
    pdf.line(left, top - 8 * mm, width - left, top - 8 * mm)

    y = top - 18 * mm
    draw("Beneficiario", "RotaHub SaaS", y, bold=True)
    y -= 8 * mm
    draw("Pagador", company.get("legal_name") or company.get("name") or company.get("nome"), y, bold=True)
    y -= 7 * mm
    draw("Documento do pagador", company.get("document"), y)
    y -= 7 * mm
    draw("Email/telefone", f"{company.get('email') or '-'} / {company.get('phone') or '-'}", y)
    y -= 10 * mm
    draw("Nosso numero", our_number, y, bold=True)
    y -= 8 * mm
    draw("Pagamento", f"#{payment.get('id')}", y)
    y -= 8 * mm
    draw("Valor", f"R$ {float(payment.get('amount') or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), y, bold=True, size=12)
    y -= 8 * mm
    draw("Vencimento", payment.get("due_date") or "Sem vencimento informado", y, bold=True)

    y -= 16 * mm
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(left, y, "Linha digitavel")
    y -= 8 * mm
    pdf.setFont("Courier-Bold", 14)
    pdf.drawString(left, y, digitable_line)

    y -= 18 * mm
    pdf.setFont("Helvetica", 9)
    notes = str(payment.get("notes") or "").strip()
    if notes:
        pdf.drawString(left, y, f"Observacao: {notes[:120]}")
        y -= 7 * mm
    pdf.drawString(left, y, f"Link do boleto: {pdf_url}")
    y -= 12 * mm
    pdf.setFont("Helvetica-Bold", 9)
    pdf.drawString(left, y, "Instrucao importante")
    y -= 6 * mm
    pdf.setFont("Helvetica", 8)
    pdf.drawString(left, y, "Documento gerado pelo controle SaaS do RotaHub. Para compensacao bancaria automatica, configure um provedor/banco integrado.")

    pdf.line(left, 42 * mm, width - left, 42 * mm)
    pdf.setFont("Courier", 22)
    pdf.drawCentredString(width / 2, 28 * mm, _only_digits(digitable_line).ljust(44, "0")[:44])
    pdf.setFont("Helvetica", 8)
    pdf.drawCentredString(width / 2, 18 * mm, "Representacao numerica para conferencia manual")
    pdf.showPage()
    pdf.save()
