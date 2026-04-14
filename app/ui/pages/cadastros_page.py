# -*- coding: utf-8 -*-
from tkinter import ttk

from app.ui.components.cadastro_crud import CadastroCRUD
from app.ui.components.page_base import PageBase
from app.ui.pages.clientes_import_page import ClientesImportPage


class CadastrosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Cadastros")

        card = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        nb = ttk.Notebook(card)
        nb.grid(row=0, column=0, sticky="nsew")

        frm_motoristas = ttk.Frame(nb, style="Content.TFrame")
        crud_motoristas = CadastroCRUD(
            frm_motoristas,
            "Motoristas",
            "motoristas",
            [
                ("nome", "NOME"),
                ("codigo", "CODIGO"),
                ("senha", "SENHA"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app,
        )
        crud_motoristas.pack(fill="both", expand=True)
        nb.add(frm_motoristas, text="Motoristas")

        frm_vendedores = ttk.Frame(nb, style="Content.TFrame")
        crud_vendedores = CadastroCRUD(
            frm_vendedores,
            "Vendedores",
            "vendedores",
            [
                ("codigo", "CODIGO"),
                ("nome", "NOME"),
                ("senha", "SENHA"),
                ("telefone", "TELEFONE"),
                ("cidade_base", "CIDADE BASE"),
                ("status", "STATUS"),
            ],
            app=app,
        )
        crud_vendedores.pack(fill="both", expand=True)
        nb.add(frm_vendedores, text="Vendedores")

        frm_usuarios = ttk.Frame(nb, style="Content.TFrame")
        crud_usuarios = CadastroCRUD(
            frm_usuarios,
            "Usuários",
            "usuarios",
            [
                ("nome", "NOME"),
                ("senha", "SENHA"),
                ("permissoes", "PERMISSÕES"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
            ],
            app=app,
        )
        crud_usuarios.pack(fill="both", expand=True)
        nb.add(frm_usuarios, text="Usuários")

        frm_veiculos = ttk.Frame(nb, style="Content.TFrame")
        crud_veiculos = CadastroCRUD(
            frm_veiculos,
            "Veiculos",
            "veiculos",
            [
                ("placa", "PLACA"),
                ("modelo", "MODELO"),
                ("capacidade_cx", "CAPACIDADE (CX)"),
            ],
            app=app,
        )
        crud_veiculos.pack(fill="both", expand=True)
        nb.add(frm_veiculos, text="Veiculos")

        frm_ajudantes = ttk.Frame(nb, style="Content.TFrame")
        crud_ajudantes = CadastroCRUD(
            frm_ajudantes,
            "Ajudantes",
            "ajudantes",
            [
                ("nome", "NOME"),
                ("sobrenome", "SOBRENOME"),
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app,
        )
        crud_ajudantes.pack(fill="both", expand=True)
        nb.add(frm_ajudantes, text="Ajudantes")

        frm_clientes = ttk.Frame(nb, style="Content.TFrame")
        clientes_page = ClientesImportPage(frm_clientes, app=app)
        clientes_page.pack(fill="both", expand=True)
        nb.add(frm_clientes, text="Clientes")

    def on_show(self):
        self.set_status("STATUS: Cadastros (CRUD + Base de Clientes via Wibi).")


__all__ = ["CadastrosPage"]
