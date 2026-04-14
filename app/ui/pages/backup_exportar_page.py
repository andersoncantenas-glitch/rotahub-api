# -*- coding: utf-8 -*-
import logging
import os
import shutil
from datetime import datetime
from tkinter import filedialog, messagebox
from tkinter import ttk

from database_runtime import validate_database_identity
from app.db.connection import get_db
from app.ui.components.page_base import PageBase
from app.utils.excel_helpers import require_pandas, _require_openpyxl as require_openpyxl

try:
    import pandas as pd
except Exception:
    pd = None

# ---------------------------------------------------------------------------
# Módulo-level globals injetados via configure_backup_exportar_page_dependencies
# ---------------------------------------------------------------------------
DB_PATH: str = ""
APP_CONFIG = None


def normalize_date_column(series):
    """Fallback seguro: passthrough sem transformação."""
    return series


def configure_backup_exportar_page_dependencies(dependencies=None, **kwargs):
    """Injeta dependências de runtime (DB_PATH, APP_CONFIG, normalize_date_column)."""
    merged = {}
    if dependencies:
        merged.update(dict(dependencies))
    if kwargs:
        merged.update(kwargs)
    globals().update(merged)


# ---------------------------------------------------------------------------

class BackupExportarPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "BackupExportar")

        self.app = app

        card = ttk.Frame(self.body, style="Card.TFrame", padding=18)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ttk.Label(card, text="Ferramentas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

        ttk.Button(card, text="\U0001F4BE FAZER BACKUP DO BANCO", style="Primary.TButton", command=self.backup_db)\
            .grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="RESTAURAR BANCO (IMPORTAR .DB)", style="Warn.TButton", command=self.restore_db)\
            .grid(row=2, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="\U0001F4E4 EXPORTAR VENDAS IMPORTADAS (EXCEL)", style="Ghost.TButton", command=self.exportar_vendas)\
            .grid(row=3, column=0, sticky="ew", pady=6)

        self.lbl = ttk.Label(card, text="Dica: Faça backup diariamente.", background="white", foreground="#444")
        self.lbl.grid(row=4, column=0, sticky="w", pady=(12, 0))

    def on_show(self):
        self.set_status("STATUS: Backup e exportações.")

    # ----------------------------
    # Helpers de segurança
    # ----------------------------
    def _is_sqlite_db_file(self, path: str) -> bool:
        """Valida assinatura do arquivo SQLite (primeiros 16 bytes)."""
        try:
            with open(path, "rb") as f:
                header = f.read(16)
            return header == b"SQLite format 3\x00"
        except Exception:
            return False

    def _make_safe_copy(self, src: str, dst: str):
        """Cópia binária simples (mantém compatibilidade)."""
        with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
            shutil.copyfileobj(fsrc, fdst)

    def _sqlite_backup_copy(self, src: str, dst: str):
        """Cópia consistente de banco SQLite, inclusive com WAL ativo."""
        import sqlite3

        src_conn = sqlite3.connect(src)
        dst_conn = sqlite3.connect(dst)
        try:
            with dst_conn:
                src_conn.backup(dst_conn)
        finally:
            try:
                dst_conn.close()
            except Exception:
                logging.debug("Falha ignorada")
            try:
                src_conn.close()
            except Exception:
                logging.debug("Falha ignorada")

    def _cleanup_sqlite_wal(self, db_path: str):
        for suffix in ("-wal", "-shm"):
            p = f"{db_path}{suffix}"
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    logging.debug("Falha ignorada")

    def _backup_current_db_automatic(self) -> str:
        """Cria um backup automático no mesmo diretório do DB (antes do restore)."""
        if not os.path.exists(DB_PATH):
            return ""
        try:
            base_dir = os.path.dirname(DB_PATH) or "."
            auto_name = f"auto_backup_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            auto_path = os.path.join(base_dir, auto_name)
            try:
                self._sqlite_backup_copy(DB_PATH, auto_path)
            except Exception:
                self._make_safe_copy(DB_PATH, auto_path)
            return auto_path
        except Exception:
            return ""

    def backup_db(self):
        if not os.path.exists(DB_PATH):
            messagebox.showerror("ERRO", "Banco não encontrado.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar backup do banco",
            defaultextension=".db",
            filetypes=[("DB", "*.db")],
            initialfile=f"backup_rota_granja_{datetime.now().strftime('%Y%m%d_%H%M')}.db"
        )
        if not path:
            return

        try:
            # melhor pratica: usar SQLite backup API quando possivel
            try:
                self._sqlite_backup_copy(DB_PATH, path)
            except Exception:
                # fallback: cópia binária (mantém funcionamento se der algo no backup API)
                self._make_safe_copy(DB_PATH, path)

            messagebox.showinfo("OK", "Backup criado com sucesso!")
            self.set_status(f"STATUS: Backup criado: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao criar backup: {str(e)}")

    def restore_db(self):
        path = filedialog.askopenfilename(
            title="Selecionar banco (.db)",
            filetypes=[("DB", "*.db")]
        )
        if not path:
            return

        # âÅâ€œââ‚¬¦ valida se é sqlite de verdade
        if not self._is_sqlite_db_file(path):
            messagebox.showerror(
                "ERRO",
                "O arquivo selecionado não parece ser um banco SQLite válido.\n\n"
                "Selecione um .db gerado pelo sistema (backup)."
            )
            return

        # confirmação mais clara (alto risco)
        if not messagebox.askyesno(
            "CONFIRMAR RESTAURAÇÃO",
            "Isso vai SUBSTITUIR seu banco atual.\n\n"
            "Recomendado: fechar telas e não estar com operações em andamento.\n\n"
            "Deseja continuar?"
        ):
            return

        try:
            try:
                validate_database_identity(path, APP_CONFIG)
            except Exception as ident_exc:
                messagebox.showerror(
                    "ERRO",
                    "O banco selecionado nao e compativel com este ambiente/tenant/company.\n\n"
                    f"Detalhes: {str(ident_exc)}"
                )
                return
            # âÅâ€œââ‚¬¦ cria backup automático do DB atual
            auto_backup = self._backup_current_db_automatic()

            # âÅâââ€š¬Åâ€œââââ‚¬Å¡¬¦ tenta fechar conexÃÆââ‚¬â„¢µes âââââ€š¬Å¡¬Åâââ€š¬Åâ€œconhecidasâââââ€š¬Å¡¬Â (melhor esforÃÆââ‚¬â„¢§o, sem quebrar)
            try:
                # se existir algum método no app para reabrir/fechar conexões, chamamos
                if hasattr(self.app, "close_db_connections"):
                    self.app.close_db_connections()
            except Exception:
                logging.debug("Falha ignorada")

            # limpa WAL/SHM antigos (melhor esforço)
            try:
                self._cleanup_sqlite_wal(DB_PATH)
            except Exception:
                logging.debug("Falha ignorada")

            # âÅâ€œââ‚¬¦ restaura por cópia binária (simples e compatÃÂÂvel)
            # Observação: se houver conexão aberta, pode falhar no Windows.
            try:
                self._sqlite_backup_copy(path, DB_PATH)
            except Exception:
                self._make_safe_copy(path, DB_PATH)
            try:
                self._cleanup_sqlite_wal(DB_PATH)
            except Exception:
                logging.debug("Falha ignorada")

            msg = "Banco restaurado! Reinicie o sistema."
            if auto_backup:
                msg += f"\n\nBackup automático do banco anterior:\n{os.path.basename(auto_backup)}"

            messagebox.showinfo("OK", msg)
            self.set_status("STATUS: Banco restaurado. Reinicie o sistema.")

        except PermissionError:
            messagebox.showerror(
                "ERRO",
                "Nao foi possivel substituir o banco (arquivo em uso).\n\n"
                "Feche o sistema e tente novamente, ou reinicie o computador."
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao restaurar banco: {str(e)}")

    def exportar_vendas(self):
        if not (require_pandas() and require_openpyxl()):
            return
        path = filedialog.asksaveasfilename(
            title="Exportar vendas importadas",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"VENDAS_IMPORTADAS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not path:
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='vendas_importadas'")
                if not cur.fetchone():
                    messagebox.showwarning("ATENÇÃO", "A tabela de vendas importadas não existe nesta base.")
                    return
                df = pd.read_sql_query("SELECT * FROM vendas_importadas ORDER BY id DESC", conn)

            if df.empty:
                messagebox.showwarning("ATENÇÃO", "Não há vendas importadas para exportar.")
                return

            try:
                if "data_venda" in df.columns:
                    df["data_venda"] = normalize_date_column(df["data_venda"])
                df.to_excel(path, index=False)
            except Exception as e:
                messagebox.showerror(
                    "ERRO",
                    "Falha ao exportar para Excel.\n\n"
                    "Verifique se o pacote 'openpyxl' está instalado.\n\n"
                    f"Detalhes: {str(e)}"
                )
                return

            messagebox.showinfo("OK", "Exportação feita com sucesso!")
            self.set_status(f"STATUS: Vendas exportadas: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao exportar vendas: {str(e)}")


__all__ = ["BackupExportarPage", "configure_backup_exportar_page_dependencies"]
