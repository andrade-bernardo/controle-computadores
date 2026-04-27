import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
import sys
import shutil

import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


# =========================
# CONFIGURAÇÕES GERAIS
# =========================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DB_PATH = BASE_DIR / "computadores.db"
BACKUP_DIR = BASE_DIR / "backups"
RELATORIOS_DIR = BASE_DIR / "relatorios"

STATUS = ["Em uso", "Estoque", "Manutenção", "Baixado"]


# =========================
# BANCO DE DADOS
# =========================

def conectar():
    return sqlite3.connect(DB_PATH)


def criar_banco():
    with conectar() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS computadores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                patrimonio TEXT UNIQUE NOT NULL,
                marca TEXT,
                modelo TEXT,
                serial TEXT,
                responsavel TEXT,
                setor TEXT,
                localizacao TEXT,
                status TEXT,
                observacoes TEXT,
                atualizado_em TEXT
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS historico (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                computador_id INTEGER,
                data_hora TEXT,
                descricao TEXT
            )
        """)


def fazer_backup_automatico():
    if not DB_PATH.exists():
        return

    BACKUP_DIR.mkdir(exist_ok=True)

    hoje = datetime.now().strftime("%Y-%m-%d")
    backup_path = BACKUP_DIR / f"backup_computadores_{hoje}.db"

    if backup_path.exists():
        return

    try:
        shutil.copy2(DB_PATH, backup_path)
    except Exception as erro:
        messagebox.showwarning(
            "Aviso de backup",
            f"Não foi possível criar o backup automático.\n\nErro: {erro}"
        )


# =========================
# APLICAÇÃO
# =========================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Controle de Ativos de TI")
        self.geometry("1280x720")
        self.minsize(1100, 650)

        criar_banco()
        fazer_backup_automatico()

        self.id_atual = None
        self.campos = {}

        self.configurar_estilos()
        self.montar_tela()
        self.carregar()
        self.atualizar_resumo()

    # =========================
    # ESTILO
    # =========================

    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#1F2937",
            rowheight=34,
            fieldbackground="#FFFFFF",
            borderwidth=0,
            font=("Segoe UI", 10)
        )

        style.configure(
            "Treeview.Heading",
            background="#E5E7EB",
            foreground="#111827",
            font=("Segoe UI", 10, "bold"),
            relief="flat"
        )

        style.map(
            "Treeview",
            background=[("selected", "#2563EB")],
            foreground=[("selected", "#FFFFFF")]
        )

    # =========================
    # INTERFACE
    # =========================

    def montar_tela(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(
            self,
            width=230,
            corner_radius=0,
            fg_color="#111827"
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        self.main = ctk.CTkFrame(
            self,
            fg_color="#F3F4F6",
            corner_radius=0
        )
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(3, weight=1)

        self.montar_menu_lateral()
        self.montar_cabecalho()
        self.montar_cards()
        self.montar_area_busca()
        self.montar_conteudo()

    def montar_menu_lateral(self):
        titulo = ctk.CTkLabel(
            self.sidebar,
            text="Ativos de TI",
            font=("Segoe UI", 24, "bold"),
            text_color="#FFFFFF"
        )
        titulo.pack(anchor="w", padx=24, pady=(28, 4))

        subtitulo = ctk.CTkLabel(
            self.sidebar,
            text="Controle patrimonial",
            font=("Segoe UI", 13),
            text_color="#9CA3AF"
        )
        subtitulo.pack(anchor="w", padx=24, pady=(0, 28))

        itens = [
            ("Computadores", "●"),
            ("Relatórios", "■"),
            ("Backups", "◆"),
            ("Configurações", "○"),
        ]

        for texto, icone in itens:
            cor = "#1F2937" if texto == "Computadores" else "#111827"

            item = ctk.CTkFrame(
                self.sidebar,
                fg_color=cor,
                corner_radius=10
            )
            item.pack(fill="x", padx=14, pady=4)

            label = ctk.CTkLabel(
                item,
                text=f"{icone}  {texto}",
                font=("Segoe UI", 14, "bold" if texto == "Computadores" else "normal"),
                text_color="#FFFFFF" if texto == "Computadores" else "#D1D5DB",
                anchor="w"
            )
            label.pack(fill="x", padx=14, pady=10)

        rodape = ctk.CTkLabel(
            self.sidebar,
            text="Sistema local\nSQLite • Excel • Backup",
            font=("Segoe UI", 12),
            text_color="#6B7280",
            justify="left"
        )
        rodape.pack(side="bottom", anchor="w", padx=24, pady=24)

    def montar_cabecalho(self):
        header = ctk.CTkFrame(
            self.main,
            fg_color="#F3F4F6",
            corner_radius=0
        )
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 8))
        header.grid_columnconfigure(0, weight=1)

        bloco_titulo = ctk.CTkFrame(header, fg_color="transparent")
        bloco_titulo.grid(row=0, column=0, sticky="w")

        titulo = ctk.CTkLabel(
            bloco_titulo,
            text="Controle de Computadores",
            font=("Segoe UI", 26, "bold"),
            text_color="#111827"
        )
        titulo.pack(anchor="w")

        subtitulo = ctk.CTkLabel(
            bloco_titulo,
            text="Gerencie patrimônio, responsáveis, setores, localização e histórico dos equipamentos.",
            font=("Segoe UI", 13),
            text_color="#6B7280"
        )
        subtitulo.pack(anchor="w", pady=(2, 0))

        self.label_data = ctk.CTkLabel(
            header,
            text=datetime.now().strftime("%d/%m/%Y"),
            font=("Segoe UI", 13, "bold"),
            text_color="#374151"
        )
        self.label_data.grid(row=0, column=1, sticky="e")

    def montar_cards(self):
        cards_frame = ctk.CTkFrame(
            self.main,
            fg_color="#F3F4F6",
            corner_radius=0
        )
        cards_frame.grid(row=1, column=0, sticky="ew", padx=24, pady=(8, 12))

        for i in range(4):
            cards_frame.grid_columnconfigure(i, weight=1)

        self.card_total = self.criar_card(cards_frame, "Total", "0", "#2563EB", 0)
        self.card_uso = self.criar_card(cards_frame, "Em uso", "0", "#059669", 1)
        self.card_estoque = self.criar_card(cards_frame, "Estoque", "0", "#7C3AED", 2)
        self.card_manutencao = self.criar_card(cards_frame, "Manutenção", "0", "#D97706", 3)

    def criar_card(self, parent, titulo, valor, cor, coluna):
        card = ctk.CTkFrame(
            parent,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5E7EB"
        )
        card.grid(row=0, column=coluna, sticky="ew", padx=6)

        indicador = ctk.CTkFrame(
            card,
            width=5,
            height=60,
            fg_color=cor,
            corner_radius=8
        )
        indicador.pack(side="left", fill="y", padx=(12, 10), pady=14)

        texto_frame = ctk.CTkFrame(card, fg_color="transparent")
        texto_frame.pack(side="left", fill="both", expand=True, pady=12)

        label_titulo = ctk.CTkLabel(
            texto_frame,
            text=titulo,
            font=("Segoe UI", 13),
            text_color="#6B7280",
            anchor="w"
        )
        label_titulo.pack(anchor="w")

        label_valor = ctk.CTkLabel(
            texto_frame,
            text=valor,
            font=("Segoe UI", 26, "bold"),
            text_color="#111827",
            anchor="w"
        )
        label_valor.pack(anchor="w")

        return label_valor

    def montar_area_busca(self):
        busca_frame = ctk.CTkFrame(
            self.main,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5E7EB"
        )
        busca_frame.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 12))
        busca_frame.grid_columnconfigure(0, weight=1)

        self.busca = tk.StringVar()

        campo_busca = ctk.CTkEntry(
            busca_frame,
            textvariable=self.busca,
            placeholder_text="Pesquisar por patrimônio, marca, modelo, responsável, setor, local ou status...",
            height=42,
            corner_radius=10,
            border_color="#D1D5DB",
            font=("Segoe UI", 13)
        )
        campo_busca.grid(row=0, column=0, sticky="ew", padx=(16, 8), pady=14)
        campo_busca.bind("<KeyRelease>", lambda e: self.carregar())

        btn_limpar = ctk.CTkButton(
            busca_frame,
            text="Limpar",
            height=42,
            width=90,
            corner_radius=10,
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#111827",
            command=self.limpar_busca
        )
        btn_limpar.grid(row=0, column=1, padx=8, pady=14)

        btn_excel = ctk.CTkButton(
            busca_frame,
            text="Exportar Excel",
            height=42,
            width=140,
            corner_radius=10,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            text_color="#FFFFFF",
            command=self.exportar_excel
        )
        btn_excel.grid(row=0, column=2, padx=(8, 16), pady=14)

    def montar_conteudo(self):
        conteudo = ctk.CTkFrame(
            self.main,
            fg_color="#F3F4F6",
            corner_radius=0
        )
        conteudo.grid(row=3, column=0, sticky="nsew", padx=24, pady=(0, 24))
        conteudo.grid_columnconfigure(0, weight=2)
        conteudo.grid_columnconfigure(1, weight=1)
        conteudo.grid_rowconfigure(0, weight=1)

        tabela_card = ctk.CTkFrame(
            conteudo,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5E7EB"
        )
        tabela_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        tabela_card.grid_columnconfigure(0, weight=1)
        tabela_card.grid_rowconfigure(1, weight=1)

        titulo_tabela = ctk.CTkLabel(
            tabela_card,
            text="Equipamentos cadastrados",
            font=("Segoe UI", 17, "bold"),
            text_color="#111827",
            anchor="w"
        )
        titulo_tabela.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))

        tabela_container = ctk.CTkFrame(tabela_card, fg_color="#FFFFFF")
        tabela_container.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        tabela_container.grid_columnconfigure(0, weight=1)
        tabela_container.grid_rowconfigure(0, weight=1)

        colunas = (
            "patrimonio",
            "marca",
            "modelo",
            "responsavel",
            "setor",
            "localizacao",
            "status",
        )

        self.tabela = ttk.Treeview(
            tabela_container,
            columns=colunas,
            show="headings",
            selectmode="browse"
        )

        titulos = {
            "patrimonio": "Patrimônio",
            "marca": "Marca",
            "modelo": "Modelo",
            "responsavel": "Responsável",
            "setor": "Setor",
            "localizacao": "Localização",
            "status": "Status",
        }

        larguras = {
            "patrimonio": 120,
            "marca": 110,
            "modelo": 130,
            "responsavel": 170,
            "setor": 120,
            "localizacao": 150,
            "status": 110,
        }

        for coluna in colunas:
            self.tabela.heading(coluna, text=titulos[coluna])
            self.tabela.column(coluna, width=larguras[coluna], minwidth=80)

        scroll_y = ttk.Scrollbar(
            tabela_container,
            orient="vertical",
            command=self.tabela.yview
        )
        self.tabela.configure(yscrollcommand=scroll_y.set)

        self.tabela.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")

        self.tabela.bind("<<TreeviewSelect>>", self.selecionar)

        self.montar_formulario(conteudo)

    def montar_formulario(self, parent):
        form = ctk.CTkFrame(
            parent,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5E7EB"
        )
        form.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        form.grid_columnconfigure(0, weight=1)
        form.grid_rowconfigure(4, weight=1)

        titulo = ctk.CTkLabel(
            form,
            text="Dados do equipamento",
            font=("Segoe UI", 17, "bold"),
            text_color="#111827",
            anchor="w"
        )
        titulo.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 4))

        subtitulo = ctk.CTkLabel(
            form,
            text="Cadastre ou atualize as informações do computador selecionado.",
            font=("Segoe UI", 12),
            text_color="#6B7280",
            anchor="w"
        )
        subtitulo.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))

        campos_frame = ctk.CTkFrame(form, fg_color="transparent")
        campos_frame.grid(row=2, column=0, sticky="ew", padx=18)
        campos_frame.grid_columnconfigure(0, weight=1)
        campos_frame.grid_columnconfigure(1, weight=1)

        campos = [
            ("patrimonio", "Patrimônio *", 0, 0),
            ("marca", "Marca", 0, 1),
            ("modelo", "Modelo", 1, 0),
            ("serial", "Serial", 1, 1),
            ("responsavel", "Responsável", 2, 0),
            ("setor", "Setor", 2, 1),
            ("localizacao", "Localização", 3, 0),
        ]

        for chave, label, linha, coluna in campos:
            bloco = ctk.CTkFrame(campos_frame, fg_color="transparent")
            bloco.grid(row=linha, column=coluna, sticky="ew", padx=5, pady=5)
            bloco.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                bloco,
                text=label,
                font=("Segoe UI", 12, "bold"),
                text_color="#374151",
                anchor="w"
            ).grid(row=0, column=0, sticky="ew")

            var = tk.StringVar()
            entrada = ctk.CTkEntry(
                bloco,
                textvariable=var,
                height=36,
                corner_radius=8,
                border_color="#D1D5DB",
                font=("Segoe UI", 12)
            )
            entrada.grid(row=1, column=0, sticky="ew", pady=(3, 0))

            self.campos[chave] = var

        status_bloco = ctk.CTkFrame(campos_frame, fg_color="transparent")
        status_bloco.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        status_bloco.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            status_bloco,
            text="Status",
            font=("Segoe UI", 12, "bold"),
            text_color="#374151",
            anchor="w"
        ).grid(row=0, column=0, sticky="ew")

        self.status = tk.StringVar(value="Em uso")

        self.status_menu = ctk.CTkOptionMenu(
            status_bloco,
            variable=self.status,
            values=STATUS,
            height=36,
            corner_radius=8,
            fg_color="#FFFFFF",
            button_color="#2563EB",
            button_hover_color="#1D4ED8",
            text_color="#111827",
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color="#111827"
        )
        self.status_menu.grid(row=1, column=0, sticky="ew", pady=(3, 0))

        obs_frame = ctk.CTkFrame(form, fg_color="transparent")
        obs_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=(8, 4))
        obs_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            obs_frame,
            text="Observações",
            font=("Segoe UI", 12, "bold"),
            text_color="#374151",
            anchor="w"
        ).grid(row=0, column=0, sticky="ew")

        self.obs = ctk.CTkTextbox(
            obs_frame,
            height=70,
            corner_radius=8,
            border_width=1,
            border_color="#D1D5DB",
            font=("Segoe UI", 12)
        )
        self.obs.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        historico_frame = ctk.CTkFrame(form, fg_color="transparent")
        historico_frame.grid(row=4, column=0, sticky="nsew", padx=18, pady=(8, 4))
        historico_frame.grid_columnconfigure(0, weight=1)
        historico_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            historico_frame,
            text="Histórico de movimentações",
            font=("Segoe UI", 12, "bold"),
            text_color="#374151",
            anchor="w"
        ).grid(row=0, column=0, sticky="ew")

        self.historico = ctk.CTkTextbox(
            historico_frame,
            height=120,
            corner_radius=8,
            border_width=1,
            border_color="#D1D5DB",
            font=("Segoe UI", 12),
            state="disabled"
        )
        self.historico.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

        botoes = ctk.CTkFrame(form, fg_color="transparent")
        botoes.grid(row=5, column=0, sticky="ew", padx=18, pady=(10, 18))
        botoes.grid_columnconfigure((0, 1, 2), weight=1)

        btn_novo = ctk.CTkButton(
            botoes,
            text="Novo",
            height=40,
            corner_radius=10,
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#111827",
            command=self.novo
        )
        btn_novo.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        btn_salvar = ctk.CTkButton(
            botoes,
            text="Salvar",
            height=40,
            corner_radius=10,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            text_color="#FFFFFF",
            command=self.salvar
        )
        btn_salvar.grid(row=0, column=1, sticky="ew", padx=5)

        btn_excluir = ctk.CTkButton(
            botoes,
            text="Excluir",
            height=40,
            corner_radius=10,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            text_color="#FFFFFF",
            command=self.excluir
        )
        btn_excluir.grid(row=0, column=2, sticky="ew", padx=(5, 0))

    # =========================
    # FUNÇÕES
    # =========================

    def atualizar_resumo(self):
        with conectar() as conn:
            total = conn.execute("SELECT COUNT(*) FROM computadores").fetchone()[0]
            em_uso = conn.execute(
                "SELECT COUNT(*) FROM computadores WHERE status = ?",
                ("Em uso",)
            ).fetchone()[0]
            estoque = conn.execute(
                "SELECT COUNT(*) FROM computadores WHERE status = ?",
                ("Estoque",)
            ).fetchone()[0]
            manutencao = conn.execute(
                "SELECT COUNT(*) FROM computadores WHERE status = ?",
                ("Manutenção",)
            ).fetchone()[0]

        self.card_total.configure(text=str(total))
        self.card_uso.configure(text=str(em_uso))
        self.card_estoque.configure(text=str(estoque))
        self.card_manutencao.configure(text=str(manutencao))

    def limpar_busca(self):
        self.busca.set("")
        self.carregar()

    def carregar(self):
        busca = f"%{self.busca.get()}%"

        with conectar() as conn:
            dados = conn.execute("""
                SELECT id,
                       patrimonio,
                       marca,
                       modelo,
                       responsavel,
                       setor,
                       localizacao,
                       status
                FROM computadores
                WHERE patrimonio LIKE ?
                   OR marca LIKE ?
                   OR modelo LIKE ?
                   OR responsavel LIKE ?
                   OR setor LIKE ?
                   OR localizacao LIKE ?
                   OR status LIKE ?
                ORDER BY patrimonio
            """, [busca] * 7).fetchall()

        self.tabela.delete(*self.tabela.get_children())

        for indice, item in enumerate(dados):
            tag = "par" if indice % 2 == 0 else "impar"
            self.tabela.insert("", "end", iid=item[0], values=item[1:], tags=(tag,))

        self.tabela.tag_configure("par", background="#FFFFFF")
        self.tabela.tag_configure("impar", background="#F9FAFB")

        self.atualizar_resumo()

    def selecionar(self, event=None):
        item = self.tabela.selection()

        if not item:
            return

        self.id_atual = int(item[0])

        with conectar() as conn:
            dados = conn.execute("""
                SELECT patrimonio,
                       marca,
                       modelo,
                       serial,
                       responsavel,
                       setor,
                       localizacao,
                       status,
                       observacoes
                FROM computadores
                WHERE id = ?
            """, (self.id_atual,)).fetchone()

            historico = conn.execute("""
                SELECT data_hora,
                       descricao
                FROM historico
                WHERE computador_id = ?
                ORDER BY id DESC
            """, (self.id_atual,)).fetchall()

        if not dados:
            return

        chaves = [
            "patrimonio",
            "marca",
            "modelo",
            "serial",
            "responsavel",
            "setor",
            "localizacao",
        ]

        for i, chave in enumerate(chaves):
            self.campos[chave].set(dados[i] or "")

        self.status.set(dados[7] or "Em uso")

        self.obs.delete("1.0", "end")
        self.obs.insert("1.0", dados[8] or "")

        texto = ""

        for data, desc in historico:
            texto += f"{data} - {desc}\n"

        self.historico.configure(state="normal")
        self.historico.delete("1.0", "end")
        self.historico.insert("1.0", texto)
        self.historico.configure(state="disabled")

    def novo(self):
        self.id_atual = None

        for var in self.campos.values():
            var.set("")

        self.status.set("Em uso")
        self.obs.delete("1.0", "end")

        self.historico.configure(state="normal")
        self.historico.delete("1.0", "end")
        self.historico.configure(state="disabled")

        self.tabela.selection_remove(self.tabela.selection())

    def salvar(self):
        dados = {chave: var.get().strip() for chave, var in self.campos.items()}

        dados["status"] = self.status.get()
        dados["observacoes"] = self.obs.get("1.0", "end").strip()

        if not dados["patrimonio"]:
            messagebox.showwarning("Atenção", "Informe o patrimônio.")
            return

        agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        try:
            with conectar() as conn:
                if self.id_atual is None:
                    cursor = conn.execute("""
                        INSERT INTO computadores (
                            patrimonio,
                            marca,
                            modelo,
                            serial,
                            responsavel,
                            setor,
                            localizacao,
                            status,
                            observacoes,
                            atualizado_em
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        dados["patrimonio"],
                        dados["marca"],
                        dados["modelo"],
                        dados["serial"],
                        dados["responsavel"],
                        dados["setor"],
                        dados["localizacao"],
                        dados["status"],
                        dados["observacoes"],
                        agora,
                    ))

                    self.id_atual = cursor.lastrowid

                    conn.execute("""
                        INSERT INTO historico (
                            computador_id,
                            data_hora,
                            descricao
                        ) VALUES (?, ?, ?)
                    """, (
                        self.id_atual,
                        agora,
                        f"Cadastrado: responsável {dados['responsavel'] or '-'}, local {dados['localizacao'] or '-'}, status {dados['status']}"
                    ))

                else:
                    conn.execute("""
                        UPDATE computadores
                        SET patrimonio = ?,
                            marca = ?,
                            modelo = ?,
                            serial = ?,
                            responsavel = ?,
                            setor = ?,
                            localizacao = ?,
                            status = ?,
                            observacoes = ?,
                            atualizado_em = ?
                        WHERE id = ?
                    """, (
                        dados["patrimonio"],
                        dados["marca"],
                        dados["modelo"],
                        dados["serial"],
                        dados["responsavel"],
                        dados["setor"],
                        dados["localizacao"],
                        dados["status"],
                        dados["observacoes"],
                        agora,
                        self.id_atual,
                    ))

                    conn.execute("""
                        INSERT INTO historico (
                            computador_id,
                            data_hora,
                            descricao
                        ) VALUES (?, ?, ?)
                    """, (
                        self.id_atual,
                        agora,
                        f"Atualizado: responsável {dados['responsavel'] or '-'}, local {dados['localizacao'] or '-'}, status {dados['status']}"
                    ))

            self.carregar()
            self.atualizar_resumo()
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso.")

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Já existe um computador com esse patrimônio.")

    def exportar_excel(self):
        try:
            with conectar() as conn:
                dados = conn.execute("""
                    SELECT patrimonio,
                           marca,
                           modelo,
                           serial,
                           responsavel,
                           setor,
                           localizacao,
                           status,
                           observacoes,
                           atualizado_em
                    FROM computadores
                    ORDER BY patrimonio
                """).fetchall()

            if not dados:
                messagebox.showwarning(
                    "Atenção",
                    "Não existem computadores cadastrados para exportar."
                )
                return

            RELATORIOS_DIR.mkdir(exist_ok=True)

            agora_arquivo = datetime.now().strftime("%Y-%m-%d_%H-%M")
            caminho = RELATORIOS_DIR / f"relatorio_computadores_{agora_arquivo}.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.title = "Computadores"

            cabecalhos = [
                "Patrimônio",
                "Marca",
                "Modelo",
                "Serial",
                "Responsável",
                "Setor",
                "Localização",
                "Status",
                "Observações",
                "Atualizado em",
            ]

            ws.append(cabecalhos)

            for linha in dados:
                ws.append(list(linha))

            cor_cabecalho = "D9EAF7"
            cor_texto = "1F2937"
            cor_borda = "D1D5DB"

            fonte_cabecalho = Font(bold=True, color=cor_texto)
            preenchimento_cabecalho = PatternFill("solid", fgColor=cor_cabecalho)
            borda_fina = Border(
                left=Side(style="thin", color=cor_borda),
                right=Side(style="thin", color=cor_borda),
                top=Side(style="thin", color=cor_borda),
                bottom=Side(style="thin", color=cor_borda),
            )

            for cell in ws[1]:
                cell.font = fonte_cabecalho
                cell.fill = preenchimento_cabecalho
                cell.border = borda_fina
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = borda_fina
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

            larguras = {
                "A": 15,
                "B": 18,
                "C": 22,
                "D": 20,
                "E": 25,
                "F": 18,
                "G": 25,
                "H": 16,
                "I": 35,
                "J": 18,
            }

            for coluna, largura in larguras.items():
                ws.column_dimensions[coluna].width = largura

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            wb.save(caminho)

            messagebox.showinfo(
                "Relatório gerado",
                f"Relatório Excel criado com sucesso em:\n\n{caminho}"
            )

        except Exception as erro:
            messagebox.showerror(
                "Erro ao exportar",
                f"Não foi possível gerar o relatório Excel.\n\nErro: {erro}"
            )

    def excluir(self):
        if self.id_atual is None:
            messagebox.showwarning("Atenção", "Selecione um computador.")
            return

        if not messagebox.askyesno("Confirmar", "Deseja excluir este computador?"):
            return

        with conectar() as conn:
            conn.execute("DELETE FROM computadores WHERE id = ?", (self.id_atual,))
            conn.execute("DELETE FROM historico WHERE computador_id = ?", (self.id_atual,))

        self.novo()
        self.carregar()
        self.atualizar_resumo()
        messagebox.showinfo("Sucesso", "Registro excluído.")


if __name__ == "__main__":
    app = App()
    app.mainloop()