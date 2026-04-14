"""
Página de gerenciamento de permissões de usuários.
Permite ao admin conceder/revogar permissões por módulo.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import logging

from app.db.connection import get_db
from app.services.permissions_service import (
    listar_permissoes_disponiveis,
    listar_permissoes_usuario,
    conceder_permissao,
    revogar_permissao,
)


class PermissionsPage(ttk.Frame):
    """Página de gerenciamento de permissões do sistema."""
    
    def __init__(self, parent, app=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.app = app
        self.context = getattr(app, "context", None) if app else None    
    def on_show(self):
        """Chamado quando a página é exibida."""
        if hasattr(self.app, 'set_status'):
            self.app.set_status("STATUS: Gerenciamento de permissões de usuários.")        
        # Layout
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self._build_header()
        self._build_main_content()
    
    def _build_header(self):
        """Cria a seção de header."""
        header = ttk.Frame(self, style="Card.TFrame", padding=12)
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header.grid_columnconfigure(1, weight=1)
        
        ttk.Label(header, text="✓ GERENCIAMENTO DE PERMISSÕES", 
                 style="CardLabel.TLabel", font=("Segoe UI", 11, "bold")).grid(
                     row=0, column=0, sticky="w"
                 )
        
        ttk.Label(header, text="Atribua e revogue permissões de usuários por módulo", 
                 style="CardLabel.TLabel").grid(row=1, column=0, sticky="w")
    
    def _build_main_content(self):
        """Cria o conteúdo principal."""
        container = ttk.Frame(self)
        container.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)
        
        # Painel esquerdo: Lista de usuários
        left_frame = ttk.LabelFrame(container, text="Usuários", padding=8)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        left_frame.grid_rowconfigure(1, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        ttk.Label(left_frame, text="Selecione um usuário:", style="CardLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        
        self.tree_usuarios = ttk.Treeview(left_frame, height=15, columns=("nome", "admin"))
        self.tree_usuarios.column("#0", width=30, anchor="center")
        self.tree_usuarios.column("nome", width=120, anchor="w")
        self.tree_usuarios.column("admin", width=50, anchor="center")
        self.tree_usuarios.heading("#0", text="ID")
        self.tree_usuarios.heading("nome", text="Nome")
        self.tree_usuarios.heading("admin", text="Admin")
        
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree_usuarios.yview)
        self.tree_usuarios.configure(yscrollcommand=scrollbar.set)
        
        self.tree_usuarios.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")
        
        self.tree_usuarios.bind("<<TreeviewSelect>>", self._on_usuario_selected)
        
        # Painel direito: Permissões
        right_frame = ttk.LabelFrame(container, text="Permissões", padding=8)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        right_frame.grid_rowconfigure(2, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)
        
        ttk.Label(right_frame, text="Disponíveis:", style="CardLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        
        # Filtro de módulo
        frm_filtro = ttk.Frame(right_frame)
        frm_filtro.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        frm_filtro.grid_columnconfigure(1, weight=1)
        
        ttk.Label(frm_filtro, text="Filtro:", style="CardLabel.TLabel").pack(side="left", padx=(0, 4))
        self.var_filtro_modulo = tk.StringVar(value="")
        combo_modulos = ttk.Combobox(frm_filtro, textvariable=self.var_filtro_modulo, 
                                     state="readonly", width=20)
        combo_modulos.pack(side="left", fill="x", expand=True)
        combo_modulos.bind("<<ComboboxSelected>>", lambda e: self._refresh_permissoes_disponiveis())
        
        self.combo_modulos = combo_modulos
        
        # Listbox de permissões disponíveis
        self.listbox_disponiveis = tk.Listbox(right_frame, height=7, font=("Segoe UI", 9))
        self.listbox_disponiveis.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        scrollbar2 = ttk.Scrollbar(right_frame, orient="vertical", command=self.listbox_disponiveis.yview)
        self.listbox_disponiveis.configure(yscrollcommand=scrollbar2.set)
        scrollbar2.grid(row=2, column=1, sticky="ns")
        
        # Botões de ação
        frm_botoes = ttk.Frame(right_frame)
        frm_botoes.grid(row=3, column=0, sticky="ew")
        frm_botoes.grid_columnconfigure(0, weight=1)
        frm_botoes.grid_columnconfigure(1, weight=1)
        
        ttk.Button(frm_botoes, text="Conceder →", style="Primary.TButton",
                  command=self._conceder_permissao).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(frm_botoes, text="← Revogar", style="Warn.TButton",
                  command=self._revogar_permissao).grid(row=0, column=1, sticky="ew")
        
        # Painel inferior: Permissões do usuário selecionado
        bottom_frame = ttk.LabelFrame(container, text="Permissões Concedidas ao Usuário", padding=8)
        bottom_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)
        
        self.text_permissoes_usuario = scrolledtext.ScrolledText(bottom_frame, height=6, wrap=tk.WORD)
        self.text_permissoes_usuario.grid(row=0, column=0, sticky="nsew")
        self.text_permissoes_usuario.configure(state="disabled")
        
        # Refresh inicial
        self._refresh_usuarios()
        self._refresh_permissoes_disponiveis()
    
    def _refresh_usuarios(self):
        """Recarrega lista de usuários."""
        try:
            # Limpar
            for item in self.tree_usuarios.get_children():
                self.tree_usuarios.delete(item)
            
            # Carregar
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT id, nome, permissoes
                    FROM usuarios
                    ORDER BY nome
                """)
                
                for user_id, nome, permissoes in cur.fetchall():
                    is_admin = "ADMIN" in (permissoes or "").upper() or nome.upper() == "ADMIN"
                    admin_txt = "✓" if is_admin else ""
                    self.tree_usuarios.insert("", "end", iid=user_id, values=(nome, admin_txt))
        except Exception as e:
            logging.exception("Erro ao carregar usuários: %s", e)
    
    def _refresh_permissoes_disponiveis(self):
        """Recarrega lista de permissões disponíveis."""
        try:
            self.listbox_disponiveis.delete(0, tk.END)
            
            res = listar_permissoes_disponiveis()
            if not res.get("ok"):
                return
            
            permissoes = res.get("data") or []
            modulo_filtro = self.var_filtro_modulo.get() or ""
            
            # Extrair módulos únicos
            modulos = sorted(set(p["modulo"] for p in permissoes))
            self.combo_modulos["values"] = [""] + modulos
            
            # Filtrar e exibir
            for perm in permissoes:
                if modulo_filtro and perm["modulo"] != modulo_filtro:
                    continue
                
                label = f"[{perm['modulo']}] {perm['nome']}"
                self.listbox_disponiveis.insert(tk.END, label)
                self.listbox_disponiveis.itemconfig(tk.END, {"data": perm["id"]})
        except Exception as e:
            logging.exception("Erro ao carregar permissões: %s", e)
    
    def _refresh_permissoes_usuario(self):
        """Recarrega permissões do usuário selecionado."""
        try:
            usuario_id = self._get_usuario_selecionado()
            if not usuario_id:
                self.text_permissoes_usuario.configure(state="normal")
                self.text_permissoes_usuario.delete("1.0", tk.END)
                self.text_permissoes_usuario.insert("1.0", "Selecione um usuário")
                self.text_permissoes_usuario.configure(state="disabled")
                return
            
            res = listar_permissoes_usuario(usuario_id)
            if not res.get("ok"):
                messagebox.showerror("Erro", res.get("error", "Erro ao carregar permissões"))
                return
            
            permissoes = res.get("data") or []
            
            # Agrupar por módulo
            por_modulo = {}
            for perm in permissoes:
                modulo = perm["modulo"]
                if modulo not in por_modulo:
                    por_modulo[modulo] = []
                por_modulo[modulo].append(perm["nome"])
            
            # Exibir
            texto = ""
            for modulo in sorted(por_modulo.keys()):
                texto += f"\n[{modulo.upper()}]\n"
                for nome in sorted(por_modulo[modulo]):
                    texto += f"  ✓ {nome}\n"
            
            self.text_permissoes_usuario.configure(state="normal")
            self.text_permissoes_usuario.delete("1.0", tk.END)
            self.text_permissoes_usuario.insert("1.0", texto.strip() or "Nenhuma permissão concedida")
            self.text_permissoes_usuario.configure(state="disabled")
        except Exception as e:
            logging.exception("Erro ao carregar permissões do usuário: %s", e)
    
    def _on_usuario_selected(self, event=None):
        """Evento de seleção de usuário."""
        self._refresh_permissoes_usuario()
    
    def _get_usuario_selecionado(self):
        """Retorna o ID do usuário selecionado."""
        selecionado = self.tree_usuarios.selection()
        return int(selecionado[0]) if selecionado else None
    
    def _conceder_permissao(self):
        """Concede a permissão selecionada ao usuário."""
        usuario_id = self._get_usuario_selecionado()
        selecionado = self.listbox_disponiveis.curselection()
        
        if not usuario_id:
            messagebox.showwarning("Atenção", "Selecione um usuário")
            return
        
        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione uma permissão")
            return
        
        # Obter ID da permissão
        try:
            res = listar_permissoes_disponiveis()
            permissoes = res.get("data") or []
            indice_listbox = selecionado[0]
            
            # Contar permissões visíveis
            modulo_filtro = self.var_filtro_modulo.get() or ""
            permissoes_filtradas = [
                p for p in permissoes 
                if not modulo_filtro or p["modulo"] == modulo_filtro
            ]
            
            if indice_listbox >= len(permissoes_filtradas):
                return
            
            permissao_id = permissoes_filtradas[indice_listbox]["id"]
            
            # Conceder
            res = conceder_permissao(usuario_id, permissao_id)
            if res.get("ok"):
                messagebox.showinfo("Sucesso", "Permissão concedida!")
                self._refresh_permissoes_usuario()
            else:
                messagebox.showerror("Erro", res.get("error", "Erro ao conceder permissão"))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {str(e)}")
            logging.exception("Erro ao conceder permissão: %s", e)
    
    def _revogar_permissao(self):
        """Revoga uma permissão do usuário selecionado."""
        usuario_id = self._get_usuario_selecionado()
        if not usuario_id:
            messagebox.showwarning("Atenção", "Selecione um usuário")
            return
        
        # Obter permissões do usuário
        res = listar_permissoes_usuario(usuario_id)
        permissoes = res.get("data") or []
        
        if not permissoes:
            messagebox.showinfo("Info", "Usuário não tem permissões para revogar")
            return
        
        # Selecionar qual permissão revogar (simples: usar popup)
        top = tk.Toplevel(self)
        top.title("Revogar Permissão")
        top.geometry("400x300")
        
        ttk.Label(top, text="Selecione a permissão a revogar:", style="CardLabel.TLabel").pack(
            pady=(8, 4), padx=8, anchor="w"
        )
        
        listbox = tk.Listbox(top, height=10)
        listbox.pack(fill="both", expand=True, padx=8, pady=4)
        
        perm_para_id = {}
        for perm in permissoes:
            label = f"[{perm['modulo']}] {perm['nome']}"
            listbox.insert(tk.END, label)
            perm_para_id[len(perm_para_id)] = perm["id"]
        
        def revogar():
            selecionado = listbox.curselection()
            if not selecionado:
                messagebox.showwarning("Atenção", "Selecione uma permissão")
                return
            
            perm_id = perm_para_id[selecionado[0]]
            res = revogar_permissao(usuario_id, perm_id)
            
            if res.get("ok"):
                messagebox.showinfo("Sucesso", "Permissão revogada!")
                self._refresh_permissoes_usuario()
                top.destroy()
            else:
                messagebox.showerror("Erro", res.get("error", "Erro ao revogar permissão"))
        
        ttk.Button(top, text="Revogar", style="Danger.TButton", command=revogar).pack(
            pady=8, padx=8, fill="x"
        )


__all__ = ["PermissionsPage"]
