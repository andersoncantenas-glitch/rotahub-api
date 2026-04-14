"""
Página de ferramentas do sistema.
Acesso a backup, logs, verificação de integridade e informações do sistema.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import logging
from datetime import datetime

from app.services.system_tools_service import (
    fazer_backup_banco,
    listar_backups,
    restaurar_backup,
    listar_logs_sistema,
    limpar_logs_sistema,
    verificar_integridade_banco,
    informacoes_sistema,
)


class SystemToolsPage(ttk.Frame):
    """Página de ferramentas e manutenção do sistema."""
    
    def __init__(self, parent, app=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.app = app
        self.context = getattr(app, "context", None) if app else None
    
    def on_show(self):
        """Chamado quando a página é exibida."""
        if hasattr(self.app, 'set_status'):
            self.app.set_status("STATUS: Ferramentas do sistema, backup e manutenção.")
        
        # Layout
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self._build_header()
        self._build_tabs()
    
    def _build_header(self):
        """Cria a seção de header."""
        header = ttk.Frame(self, style="Card.TFrame", padding=12)
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header.grid_columnconfigure(0, weight=1)
        
        ttk.Label(header, text="⚙ FERRAMENTAS DO SISTEMA", 
                 style="CardLabel.TLabel", font=("Segoe UI", 11, "bold")).grid(
                     row=0, column=0, sticky="w"
                 )
        
        ttk.Label(header, text="Backup, logs, verificação de integridade e informações", 
                 style="CardLabel.TLabel").grid(row=1, column=0, sticky="w")
    
    def _build_tabs(self):
        """Cria o notebook com abas de ferramentas."""
        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        
        # Aba 1: Backup
        self._build_tab_backup(notebook)
        
        # Aba 2: Logs
        self._build_tab_logs(notebook)
        
        # Aba 3: Verificação
        self._build_tab_verificacao(notebook)
        
        # Aba 4: Informações
        self._build_tab_informacoes(notebook)
    
    def _build_tab_backup(self, notebook):
        """Aba de gerenciamento de backups."""
        tab = ttk.Frame(notebook, padding=12)
        notebook.add(tab, text="Backup e Restauração")
        
        tab.grid_rowconfigure(2, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        
        # Seção superior: Criar backup
        frm_criar = ttk.LabelFrame(tab, text="Criar Backup", padding=12)
        frm_criar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        
        txt_desc = ttk.Label(frm_criar, text="Salve uma cópia do banco para restauração futura.", 
                            wraplength=400, justify="left")
        txt_desc.pack(pady=(0, 8))
        
        def criar_backup():
            nome_usuario = getattr(self.app.user if self.app else {}, "get", lambda x: "ADMIN")("nome") or "ADMIN"
            res = fazer_backup_banco(usuario=nome_usuario)
            
            if res.get("ok"):
                dados = res.get("data") or {}
                mensagem = f"Backup criado com sucesso!\n\nArquivo: {dados.get('arquivo')}\nTamanho: {dados.get('tamanho_kb')} KB"
                messagebox.showinfo("Sucesso", mensagem)
                self._refresh_backups()
            else:
                messagebox.showerror("Erro", f"Erro ao criar backup:\n{res.get('error')}")
        
        ttk.Button(frm_criar, text="📦 Criar Backup Agora", style="Primary.TButton",
                  command=criar_backup).pack(fill="x")
        
        # Seção média: Lista de backups
        frm_lista = ttk.LabelFrame(tab, text="Backups Disponíveis", padding=12)
        frm_lista.grid(row=1, column=0, sticky="ew")
        frm_lista.grid_columnconfigure(0, weight=1)
        frm_lista.grid_rowconfigure(0, weight=1)
        
        self.tree_backups = ttk.Treeview(frm_lista, height=8, columns=("tamanho", "data"))
        self.tree_backups.column("#0", width=250, anchor="w")
        self.tree_backups.column("tamanho", width=80, anchor="e")
        self.tree_backups.column("data", width=150, anchor="w")
        self.tree_backups.heading("#0", text="Arquivo")
        self.tree_backups.heading("tamanho", text="Tamanho (KB)")
        self.tree_backups.heading("data", text="Data de Criação")
        
        scrollbar = ttk.Scrollbar(frm_lista, orient="vertical", command=self.tree_backups.yview)
        self.tree_backups.configure(yscrollcommand=scrollbar.set)
        
        self.tree_backups.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Seção inferior: Restaurar
        frm_restaurar = ttk.LabelFrame(tab, text="Restaurar Backup", padding=12)
        frm_restaurar.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        
        ttk.Label(frm_restaurar, text="⚠ ATENÇÃO: Restaurar sobrescreverá os dados atuais!",
                 foreground="red", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        
        def restaurar_backup():
            selecionado = self.tree_backups.selection()
            if not selecionado:
                messagebox.showwarning("Atenção", "Selecione um backup para restaurar")
                return
            
            # Obter dados do backup
            item = selecionado[0]
            valores = self.tree_backups.item(item)
            arquivo = valores.get("values", [None])[3] if hasattr(self.tree_backups, "_backup_files") else None
            
            if not arquivo:
                messagebox.showerror("Erro", "Não foi possível obter o arquivo do backup")
                return
            
            if not messagebox.askyesno("Confirmar", 
                                      "Tem certeza que deseja restaurar este backup?\n"
                                      "This will replace current data with backup data."):
                return
            
            nome_usuario = getattr(self.app.user if self.app else {}, "get", lambda x: "ADMIN")("nome") or "ADMIN"
            res = restaurar_backup(arquivo, usuario=nome_usuario)
            
            if res.get("ok"):
                messagebox.showinfo("Sucesso", "Backup restaurado com sucesso!\nO aplicativo será reiniciado.")
                if self.app:
                    self.app.root.quit()
            else:
                messagebox.showerror("Erro", f"Erro ao restaurar:\n{res.get('error')}")
        
        ttk.Button(frm_restaurar, text="⬅ Restaurar Backup Selecionado", 
                  style="Warn.TButton", command=restaurar_backup).pack(fill="x")
        
        # Carregar backups iniciais
        self._refresh_backups()
    
    def _refresh_backups(self):
        """Recarrega lista de backups."""
        try:
            for item in self.tree_backups.get_children():
                self.tree_backups.delete(item)
            
            res = listar_backups()
            if not res.get("ok"):
                return
            
            backups = res.get("data") or []
            self._backup_files = backups  # Guardar para acesso depois
            
            for backup in backups:
                arquivo = backup.get("arquivo", "")
                tamanho = backup.get("tamanho_kb", 0)
                data = backup.get("data_criacao", "")
                caminho = backup.get("caminho", "")
                
                # Exibir timestamp no formato amigável
                try:
                    dt = datetime.fromisoformat(data)
                    data_fmt = dt.strftime("%d/%m/%Y %H:%M:%S")
                except:
                    data_fmt = data
                
                self.tree_backups.insert("", "end", iid=len(self._backup_files)-1,
                                        values=(tamanho, data_fmt, "", caminho))
                self.tree_backups.item(self.tree_backups.get_children()[-1], text=arquivo)
        except Exception as e:
            logging.exception("Erro ao listar backups: %s", e)
    
    def _build_tab_logs(self, notebook):
        """Aba de gerenciamento de logs."""
        tab = ttk.Frame(notebook, padding=12)
        notebook.add(tab, text="Logs do Sistema")
        
        tab.grid_rowconfigure(1, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        
        # Controles
        frm_ctrl = ttk.Frame(tab)
        frm_ctrl.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        
        ttk.Button(frm_ctrl, text="Recarregar Logs", style="Ghost.TButton",
                  command=self._refresh_logs).pack(side="left", padx=(0, 4))
        
        ttk.Label(frm_ctrl, text="Limpar logs com mais de").pack(side="left", padx=(0, 4))
        
        self.spinbox_dias = ttk.Spinbox(frm_ctrl, from_=1, to=365, width=5)
        self.spinbox_dias.set(30)
        self.spinbox_dias.pack(side="left", padx=(0, 4))
        
        ttk.Label(frm_ctrl, text="dias:").pack(side="left", padx=(0, 4))
        
        def limpar():
            dias = int(self.spinbox_dias.get())
            if messagebox.askyesno("Confirmar", f"Limpar logs com mais de {dias} dias?"):
                res = limpar_logs_sistema(dias)
                if res.get("ok"):
                    msgbox_show = messagebox.showinfo("Sucesso", 
                        f"Logs removidos: {res.get('data', {}).get('linhas_deletadas', 0)}")
                    self._refresh_logs()
                else:
                    messagebox.showerror("Erro", res.get("error"))
        
        ttk.Button(frm_ctrl, text="Limpar", style="Danger.TButton",
                  command=limpar).pack(side="left")
        
        # Logs
        frm_logs = ttk.LabelFrame(tab, text="Últimas Ações", padding=8)
        frm_logs.grid(row=1, column=0, sticky="nsew")
        frm_logs.grid_rowconfigure(0, weight=1)
        frm_logs.grid_columnconfigure(0, weight=1)
        
        self.tree_logs = ttk.Treeview(frm_logs, height=15, columns=("tipo", "usuario", "status", "data"))
        self.tree_logs.column("#0", width=150, anchor="w")
        self.tree_logs.column("tipo", width=100, anchor="w")
        self.tree_logs.column("usuario", width=80, anchor="w")
        self.tree_logs.column("status", width=60, anchor="center")
        self.tree_logs.column("data", width=150, anchor="w")
        
        self.tree_logs.heading("#0", text="Descrição")
        self.tree_logs.heading("tipo", text="Tipo de Ação")
        self.tree_logs.heading("usuario", text="Usuário")
        self.tree_logs.heading("status", text="Status")
        self.tree_logs.heading("data", text="Data/Hora")
        
        scrollbar = ttk.Scrollbar(frm_logs, orient="vertical", command=self.tree_logs.yview)
        self.tree_logs.configure(yscrollcommand=scrollbar.set)
        
        self.tree_logs.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self._refresh_logs()
    
    def _refresh_logs(self):
        """Recarrega logs do sistema."""
        try:
            for item in self.tree_logs.get_children():
                self.tree_logs.delete(item)
            
            res = listar_logs_sistema(100)
            if not res.get("ok"):
                return
            
            logs = res.get("data") or []
            for log in logs:
                self.tree_logs.insert("", "end",
                                     text=log.get("descricao", ""),
                                     values=(
                                         log.get("tipo_acao", ""),
                                         log.get("usuario", ""),
                                         log.get("status", ""),
                                         log.get("executado_em", "")
                                     ))
        except Exception as e:
            logging.exception("Erro ao carregar logs: %s", e)
    
    def _build_tab_verificacao(self, notebook):
        """Aba de verificação de integridade."""
        tab = ttk.Frame(notebook, padding=12)
        notebook.add(tab, text="Verificação de Integridade")
        
        ttk.Label(tab, text="Verifique a integridade e consistência do banco de dados.",
                 wraplength=400, justify="left").pack(pady=(0, 12))
        
        def fazer_verificacao():
            res = verificar_integridade_banco()
            
            status = "✓ OK" if res.get("ok") else "✗ PROBLEMAS"
            cor = "green" if res.get("ok") else "red"
            
            txt_resultado.configure(state="normal")
            txt_resultado.delete("1.0", tk.END)
            txt_resultado.insert("1.0", 
                f"Status: {status}\n\n"
                f"Resultado: {res.get('data', {}).get('integridade', res.get('error'))}\n\n"
                f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            )
            txt_resultado.configure(state="disabled")
            
            if res.get("ok"):
                messagebox.showinfo("Sucesso", "✓ Banco de dados íntegro e consistente!")
            else:
                messagebox.showwarning("Atenção", f"Problemas encontrados:\n{res.get('error')}")
        
        ttk.Button(tab, text="▶ Verificar Integridade", style="Primary.TButton",
                  command=fazer_verificacao).pack(fill="x", pady=(0, 12))
        
        ttk.Label(tab, text="Resultado da Verificação:", style="CardLabel.TLabel").pack(
            anchor="w", pady=(0, 4)
        )
        
        txt_resultado = scrolledtext.ScrolledText(tab, height=10, wrap=tk.WORD)
        txt_resultado.pack(fill="both", expand=True)
        txt_resultado.insert("1.0", "Clique em 'Verificar Integridade' para testar o banco")
        txt_resultado.configure(state="disabled")
    
    def _build_tab_informacoes(self, notebook):
        """Aba de informações do sistema."""
        tab = ttk.Frame(notebook, padding=12)
        notebook.add(tab, text="Informações do Sistema")
        
        ttk.Button(tab, text="🔄 Atualizar Informações", style="Ghost.TButton",
                  command=self._refresh_info).pack(anchor="e", pady=(0, 8))
        
        txt_info = scrolledtext.ScrolledText(tab, wrap=tk.WORD, height=20)
        txt_info.pack(fill="both", expand=True)
        txt_info.configure(state="disabled")
        
        self.txt_info = txt_info
        self._refresh_info()
    
    def _refresh_info(self):
        """Recarrega informações do sistema."""
        try:
            res = informacoes_sistema()
            if not res.get("ok"):
                texto = f"Erro: {res.get('error')}"
            else:
                dados = res.get("data", {})
                registros = dados.get("registros_por_tabela", {})
                
                texto = "INFORMAÇÕES DO SISTEMA\n"
                texto += "=" * 50 + "\n\n"
                
                texto += f"Data/Hora Atual: {dados.get('data_hora_atual', '-')}\n"
                texto += f"Tamanho do Banco: {dados.get('tamanho_banco_kb', 0)} KB\n"
                texto += f"Última Ação: {dados.get('ultima_acao_em', 'Sem ações registradas')}\n\n"
                
                texto += "REGISTROS POR TABELA\n"
                texto += "-" * 50 + "\n"
                for tabela, qtd in sorted(registros.items()):
                    texto += f"{tabela:.<30} {qtd:>6} registros\n"
            
            self.txt_info.configure(state="normal")
            self.txt_info.delete("1.0", tk.END)
            self.txt_info.insert("1.0", texto)
            self.txt_info.configure(state="disabled")
        except Exception as e:
            logging.exception("Erro ao carregar informações: %s", e)


__all__ = ["SystemToolsPage"]
