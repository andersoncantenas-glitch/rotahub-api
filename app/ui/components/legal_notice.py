import logging
import tkinter as tk
from tkinter import ttk


DEVELOPER_COMPANY = "Nexlume Tecnologia"
DEVELOPER_TAGLINE = "Solucoes em desenvolvimento de sistemas e automacao inteligente"
DEVELOPER_FOOTER_TEXT = f"Desenvolvido por {DEVELOPER_COMPANY}"


def build_contact_line(support_whatsapp: str = "", support_email: str = "") -> str:
    canais = []
    if str(support_whatsapp or "").strip():
        canais.append(f"WhatsApp {str(support_whatsapp).strip()}")
    if str(support_email or "").strip():
        canais.append(f"E-mail {str(support_email).strip()}")
    if not canais:
        return "Contato: nao informado"
    return "Contato: " + " | ".join(canais)


def build_developer_summary_text(support_whatsapp: str = "", support_email: str = "") -> str:
    return (
        f"Sistema desenvolvido por {DEVELOPER_COMPANY}\n"
        f"{DEVELOPER_TAGLINE}\n"
        f"{build_contact_line(support_whatsapp, support_email)}"
    )


def build_legal_notice_text(
    app_name: str = "Sistema",
    app_version: str = "",
    support_whatsapp: str = "",
    support_email: str = "",
) -> str:
    app_label = str(app_name or "Sistema").strip() or "Sistema"
    version_label = f" | versao {app_version.strip()}" if str(app_version or "").strip() else ""
    contato = build_contact_line(support_whatsapp, support_email)
    return "\n".join(
        [
            f"TERMOS DE USO, CONDICOES E TITULARIDADE - {app_label}{version_label}",
            "",
            "1. TITULARIDADE INTELECTUAL",
            (
                f"Este sistema e uma obra intelectual original, concebida e desenvolvida integralmente por "
                f"{DEVELOPER_COMPANY}."
            ),
            (
                f"{DEVELOPER_COMPANY} detem 100% da autoria, criacao, arquitetura, interfaces, fluxos, "
                "regras de negocio, automacoes, documentacao tecnica e codigo-fonte deste sistema, "
                "ressalvados apenas componentes de terceiros utilizados sob suas respectivas licencas."
            ),
            "",
            "2. LICENCA DE USO",
            (
                "O uso do sistema concede apenas permissao operacional de uso dentro do escopo autorizado "
                "pelo titular, sem qualquer transferencia de propriedade intelectual, cessao de codigo-fonte "
                "ou compartilhamento automatico de direitos patrimoniais."
            ),
            "",
            "3. RESTRICOES",
            (
                "E vedado copiar, revender, sublicenciar, distribuir, desmontar, fazer engenharia reversa, "
                "reutilizar total ou parcialmente o codigo, remover a identificacao de autoria ou reivindicar "
                "para si a criacao deste sistema sem autorizacao previa e expressa do titular."
            ),
            "",
            "4. USO RESPONSAVEL",
            (
                "O usuario e responsavel pelas informacoes inseridas, pelo uso adequado das credenciais de "
                "acesso e pela observancia das normas internas, legais e regulatorias aplicaveis ao seu negocio."
            ),
            "",
            "5. DADOS, OPERACAO E LIMITACOES",
            (
                "A disponibilidade do sistema pode depender de infraestrutura, conectividade, servicos de "
                "terceiros, configuracoes locais e procedimentos corretos de backup e operacao. O uso do sistema "
                "nao elimina a necessidade de validacao humana, controle interno e boas praticas de seguranca."
            ),
            "",
            "6. ATUALIZACOES E EVOLUCOES",
            (
                f"{DEVELOPER_COMPANY} podera realizar manutencoes, melhorias, correcoes, ajustes visuais, "
                "alteracoes de fluxo e novas funcionalidades, preservando a titularidade intelectual sobre "
                "cada evolucao implementada."
            ),
            "",
            "7. CREDITOS E AUTORIA",
            (
                "A identificacao de desenvolvimento exibida no sistema integra o registro de autoria e deve ser "
                "preservada como referencia de origem tecnica e intelectual da solucao."
            ),
            "",
            "8. ACEITE",
            (
                "Ao acessar ou utilizar o sistema, o usuario declara ciencia e concordancia com estes termos, "
                "reconhecendo expressamente a autoria e a titularidade exclusiva de "
                f"{DEVELOPER_COMPANY}."
            ),
            "",
            "DESENVOLVEDOR",
            f"Sistema desenvolvido por {DEVELOPER_COMPANY}",
            DEVELOPER_TAGLINE,
            contato,
        ]
    )


def open_legal_notice(
    owner,
    *,
    app_name: str = "Sistema",
    app_version: str = "",
    support_whatsapp: str = "",
    support_email: str = "",
    apply_window_icon=None,
):
    host = owner.winfo_toplevel() if hasattr(owner, "winfo_toplevel") else owner
    win = tk.Toplevel(host)
    win.title("Termos de uso e autoria")
    if callable(apply_window_icon):
        try:
            apply_window_icon(win)
        except Exception:
            logging.debug("Falha ao aplicar icone na janela legal.", exc_info=True)
    win.geometry("860x620")
    win.minsize(700, 520)
    try:
        win.transient(host)
    except Exception:
        logging.debug("Falha ao definir janela transitoria.", exc_info=True)
    try:
        win.grab_set()
    except Exception:
        logging.debug("Falha ao aplicar grab_set.", exc_info=True)

    root = ttk.Frame(win, style="Content.TFrame", padding=14)
    root.pack(fill="both", expand=True)
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=1)

    header = ttk.Frame(root, style="Card.TFrame", padding=(12, 10, 12, 10))
    header.grid(row=0, column=0, sticky="ew")
    header.grid_columnconfigure(0, weight=1)

    ttk.Label(header, text="Termos de uso, condicoes e autoria", style="CardTitle.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Label(
        header,
        text=build_developer_summary_text(support_whatsapp, support_email),
        background="white",
        foreground="#111827",
        justify="left",
        wraplength=760,
        font=("Segoe UI", 9),
    ).grid(row=1, column=0, sticky="w", pady=(6, 0))

    text_wrap = ttk.Frame(root, style="Content.TFrame")
    text_wrap.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
    text_wrap.grid_rowconfigure(0, weight=1)
    text_wrap.grid_columnconfigure(0, weight=1)

    txt = tk.Text(
        text_wrap,
        wrap="word",
        font=("Segoe UI", 10),
        bg="white",
        fg="#111827",
        relief="solid",
        bd=1,
        padx=10,
        pady=10,
    )
    txt.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(text_wrap, orient="vertical", command=txt.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    txt.configure(yscrollcommand=vsb.set)

    txt.insert(
        "1.0",
        build_legal_notice_text(
            app_name=app_name,
            app_version=app_version,
            support_whatsapp=support_whatsapp,
            support_email=support_email,
        ),
    )
    txt.config(state="disabled")

    actions = ttk.Frame(root, style="Content.TFrame")
    actions.grid(row=2, column=0, sticky="e", pady=(10, 0))
    ttk.Button(actions, text="Fechar", style="Primary.TButton", command=win.destroy).pack(side="right")

    return win


__all__ = [
    "DEVELOPER_COMPANY",
    "DEVELOPER_TAGLINE",
    "DEVELOPER_FOOTER_TEXT",
    "build_contact_line",
    "build_developer_summary_text",
    "build_legal_notice_text",
    "open_legal_notice",
]
