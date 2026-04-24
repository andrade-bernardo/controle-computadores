import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DB_PATH = BASE_DIR / "computadores.db"

STATUS = ["Em uso", "Estoque", "Manutenção", "Baixado"]


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Controle de Computadores")
        self.geometry("1000x600")

        criar_banco()

        self.id_atual = None

        self.montar_tela()
        self.carregar()

    def montar_tela(self):
        self.columnconfigure(0, weight=2)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        topo = ttk.Frame(self, padding=10)
        topo.grid(row=0, column=0, columnspan=2, sticky="ew")
        topo.columnconfigure(1, weight=1)

        ttk.Label(topo, text="Pesquisar:").grid(row=0, column=0, padx=5)

        self.busca = tk.StringVar()
        campo_busca = ttk.Entry(topo, textvariable=self.busca)
        campo_busca.grid(row=0, column=1, sticky="ew", padx=5)
        campo_busca.bind("<KeyRelease>", lambda e: self.carregar())

        ttk.Button(topo, text="Limpar", command=self.limpar_busca).grid(row=0, column=2, padx=5)

        colunas = (
            "patrimonio",
            "marca",
            "modelo",
            "responsavel",
            "setor",
            "localizacao",
            "status",
        )

        self.tabela = ttk.Treeview(self, columns=colunas, show="headings")

        titulos = {
            "patrimonio": "Patrimônio",
            "marca": "Marca",
            "modelo": "Modelo",
            "responsavel": "Responsável",
            "setor": "Setor",
            "localizacao": "Localização",
            "status": "Status",
        }

        for coluna in colunas:
            self.tabela.heading(coluna, text=titulos[coluna])
            self.tabela.column(coluna, width=120)

        self.tabela.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.tabela.bind("<<TreeviewSelect>>", self.selecionar)

        form = ttk.LabelFrame(self, text="Cadastro", padding=10)
        form.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        form.columnconfigure(1, weight=1)

        self.campos = {}

        labels = [
            ("patrimonio", "Patrimônio *"),
            ("marca", "Marca"),
            ("modelo", "Modelo"),
            ("serial", "Serial"),
            ("responsavel", "Responsável"),
            ("setor", "Setor"),
            ("localizacao", "Localização"),
        ]

        for linha, (chave, texto) in enumerate(labels):
            ttk.Label(form, text=texto).grid(row=linha, column=0, sticky="w", pady=4)

            var = tk.StringVar()
            ttk.Entry(form, textvariable=var).grid(row=linha, column=1, sticky="ew", pady=4)

            self.campos[chave] = var

        ttk.Label(form, text="Status").grid(row=7, column=0, sticky="w", pady=4)

        self.status = tk.StringVar(value="Em uso")
        ttk.Combobox(
            form,
            textvariable=self.status,
            values=STATUS,
            state="readonly"
        ).grid(row=7, column=1, sticky="ew", pady=4)

        ttk.Label(form, text="Observações").grid(row=8, column=0, sticky="nw", pady=4)

        self.obs = tk.Text(form, height=5)
        self.obs.grid(row=8, column=1, sticky="ew", pady=4)

        botoes = ttk.Frame(form)
        botoes.grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)
        botoes.columnconfigure((0, 1, 2), weight=1)

        ttk.Button(botoes, text="Novo", command=self.novo).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=3
        )

        ttk.Button(botoes, text="Salvar", command=self.salvar).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=3
        )

        ttk.Button(botoes, text="Excluir", command=self.excluir).grid(
            row=0,
            column=2,
            sticky="ew",
            padx=3
        )

        ttk.Label(form, text="Histórico").grid(row=10, column=0, columnspan=2, sticky="w")

        self.historico = tk.Text(form, height=8, state="disabled")
        self.historico.grid(row=11, column=0, columnspan=2, sticky="nsew", pady=5)

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

        for item in dados:
            self.tabela.insert("", "end", iid=item[0], values=item[1:])

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

        self.historico.config(state="normal")
        self.historico.delete("1.0", "end")
        self.historico.insert("1.0", texto)
        self.historico.config(state="disabled")

    def novo(self):
        self.id_atual = None

        for var in self.campos.values():
            var.set("")

        self.status.set("Em uso")

        self.obs.delete("1.0", "end")

        self.historico.config(state="normal")
        self.historico.delete("1.0", "end")
        self.historico.config(state="disabled")

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
                        f"Cadastrado para {dados['responsavel']} em {dados['localizacao']}"
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
                        f"Atualizado: responsável {dados['responsavel']}, local {dados['localizacao']}, status {dados['status']}"
                    ))

            self.carregar()
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso.")

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Já existe um computador com esse patrimônio.")

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
        messagebox.showinfo("Sucesso", "Registro excluído.")


if __name__ == "__main__":
    app = App()
    app.mainloop()