import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import pandas as pd
from datetime import datetime
from fuzzywuzzy import process
import matplotlib.pyplot as plt
import smtplib
from email.mime.text import MIMEText

# Caminho da planilha Excel
EXCEL_PATH = "C:\\Oficina\\dados.xlsx"

# Verificar e criar planilha e campos necessários
def verificar_criar_planilha():
    if not os.path.exists(EXCEL_PATH):
        with pd.ExcelWriter(EXCEL_PATH) as writer:
            df_clientes = pd.DataFrame(columns=["ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço"])
            df_clientes.to_excel(writer, sheet_name="Gerenciar Clientes", index=False)
            df_produtos = pd.DataFrame(columns=["ID", "Nome", "Preço", "Estoque", "Ponto de Pedido"])
            df_produtos.to_excel(writer, sheet_name="Gerenciar Produtos", index=False)
            df_ordens = pd.DataFrame(columns=["ID", "Cliente", "Modelo Moto", "Descrição de Serviços", "Peças Utilizadas", "Quantidade Utilizada", "Valor", "Valor da Mão de Obra", "Status"])
            df_ordens.to_excel(writer, sheet_name="Ordens de Serviço", index=False)
    else:
        with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='overlay') as writer:
            if "Gerenciar Clientes" not in writer.sheets:
                df_clientes = pd.DataFrame(columns=["ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço"])
                df_clientes.to_excel(writer, sheet_name="Gerenciar Clientes", index=False)
            if "Gerenciar Produtos" not in writer.sheets:
                df_produtos = pd.DataFrame(columns=["ID", "Nome", "Preço", "Estoque", "Ponto de Pedido"])
                df_produtos.to_excel(writer, sheet_name="Gerenciar Produtos", index=False)
            if "Ordens de Serviço" not in writer.sheets:
                df_ordens = pd.DataFrame(columns=["ID", "Cliente", "Modelo Moto", "Descrição de Serviços", "Peças Utilizadas", "Quantidade Utilizada", "Valor", "Valor da Mão de Obra", "Status"])
                df_ordens.to_excel(writer, sheet_name="Ordens de Serviço", index=False)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestão JM")
        self.root.geometry("1200x700")

        # Definir atributos iniciais
        self.tree_clientes = None
        self.tree_produtos = None
        self.logo_image = None
        self.frame_produtos = None
        self.frame_relatorios = None
        self.frame_ordens = None
        self.tree_ordens = None
        self.frame_configuracoes = None
        self.frame_email = None
        self.valor_peca_var = tk.StringVar() # Adicionado
        self.peca_var = tk.StringVar() # Adicionado

        # Verificar e criar planilha
        verificar_criar_planilha()

        # Configuração inicial
        self.sidebar()
        self.setup_main_frame()
        self.carregar_wallpaper()

        # Fechar conexão ao fechar o app
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def sidebar(self):
        # Criar a sidebar
        self.sidebar_frame = tk.Frame(self.root, bg="#2c3e50", width=250, height=700)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Logo
        self.logo_label = tk.Label(self.sidebar_frame, bg="#2c3e50", fg="white", text="LOGO", font=("Arial", 16, "bold"))
        self.logo_label.pack(pady=20)
        self.carregar_logo()

        # Botões da Sidebar
        buttons = [
            ("Tela Inicial", self.tela_inicial),
            ("Gerenciar Clientes", self.gerenciar_clientes),
            ("Gerenciar Produtos", self.gerenciar_produtos),
            ("Gerenciar Relatórios", self.gerenciar_relatorios),
            ("Ordens de Serviço", self.gerenciar_ordens),
            ("Configurações", self.configuracoes)
        ]

        for text, command in buttons:
            btn = tk.Button(self.sidebar_frame, text=text, command=command, bg="#34495e", fg="white", relief="flat", font=("Arial", 12))
            btn.pack(fill=tk.X, padx=10, pady=5)

    def setup_main_frame(self):
        # Frame principal
        self.main_frame = tk.Frame(self.root, bg="white", width=950, height=700)
        self.main_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def carregar_wallpaper(self):
        try:
            wallpaper_path = os.path.join("C:\\Temp", "wallpaper.jpg")
            if os.path.exists(wallpaper_path):
                img = Image.open(wallpaper_path)
                img = img.resize((950, 700), Image.LANCZOS)
                self.wallpaper_image = ImageTk.PhotoImage(img)
                label = tk.Label(self.main_frame, image=self.wallpaper_image)
                label.place(x=0, y=0, relwidth=1, relheight=1)
            else:
                raise FileNotFoundError(f"Wallpaper não encontrado em {wallpaper_path}")
        except Exception as e:
            print("Erro ao carregar wallpaper:", e)
            messagebox.showerror("Erro", f"Erro ao carregar wallpaper: {e}")

    def salvar_wallpaper(self):
        file_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
        if file_path:
            try:
                wallpaper_path = os.path.join("C:\\Temp", "wallpaper.jpg")
                img = Image.open(file_path)
                img.save(wallpaper_path)
                messagebox.showinfo("Sucesso", "Wallpaper salvo com sucesso!")
                self.carregar_wallpaper()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar wallpaper: {e}")

    def carregar_logo(self):
        try:
            logo_path = os.path.join("C:\\Temp", "logo.jpg")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((200, 100), Image.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img)
                self.logo_label.config(image=self.logo_image)
            else:
                raise FileNotFoundError(f"Logo não encontrado em {logo_path}")
        except Exception as e:
            print("Erro ao carregar logo:", e)
            messagebox.showerror("Erro", f"Erro ao carregar logo: {e}")

    def salvar_logo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
        if file_path:
            try:
                logo_path = os.path.join("C:\\Temp", "logo.jpg")
                img = Image.open(file_path)
                img.save(logo_path)
                messagebox.showinfo("Sucesso", "Logo salvo com sucesso!")
                self.carregar_logo()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar logo: {e}")

    def tela_inicial(self):
        self.clear_main_frame()
        self.carregar_wallpaper()
        now = datetime.now()
        data_hora_str = now.strftime("%d/%m/%Y %H:%M:%S")
        tk.Label(self.main_frame, text="Bem-vindo à JM", font=("Arial", 24), bg="white").pack(pady=20)
        tk.Label(self.main_frame, text=data_hora_str, font=("Arial", 24), bg="white").pack(pady=20)

    def gerenciar_clientes(self):
        self.clear_main_frame()
        self.tree_clientes = tk.Frame(self.main_frame)
        self.tree_clientes.pack(fill=tk.BOTH, expand=True)

        # Formulário de inclusão de cliente
        tk.Label(self.tree_clientes, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
        nome_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=nome_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.tree_clientes, text="CPF/CNPJ:").grid(row=1, column=0, padx=5, pady=5)
        cpf_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=cpf_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.tree_clientes, text="Telefone:").grid(row=2, column=0, padx=5, pady=5)
        telefone_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=telefone_var).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.tree_clientes, text="Email:").grid(row=3, column=0, padx=5, pady=5)
        email_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=email_var).grid(row=3, column=1, padx=5, pady=5)
        tk.Label(self.tree_clientes, text="Endereço:").grid(row=4, column=0, padx=5, pady=5)
        endereco_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=endereco_var).grid(row=4, column=1, padx=5, pady=5)

        tk.Button(
            self.tree_clientes,
            text="Salvar",
            command=lambda: self.salvar_cliente(nome_var, cpf_var, telefone_var, email_var, endereco_var)
        ).grid(row=5, column=0, columnspan=2, pady=20)

        # Campo de busca avançada
        tk.Label(self.tree_clientes, text="Buscar Cliente:").grid(row=6, column=0, padx=5, pady=5)
        busca_var = tk.StringVar()
        tk.Entry(self.tree_clientes, textvariable=busca_var).grid(row=6, column=1, padx=5, pady=5)
        tk.Button(
            self.tree_clientes,
            text="Buscar",
            command=lambda: self.buscar_cliente_avancada(busca_var.get(), self.tree_clientes)
        ).grid(row=6, column=2, padx=5, pady=5)

        # Tabela com os clientes cadastrados
        self.tree_clientes = self.create_client_table()
        self.tree_clientes.grid(row=7, column=0, columnspan=3, padx=10, pady=10)

        self.load_client_data()

    def create_client_table(self):
        columns = ("ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço")
        tree = ttk.Treeview(self.tree_clientes, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        tree.bind("<Button-3>", self.show_client_menu)
        return tree

    def load_client_data(self):
        for row in self.tree_clientes.get_children():
            self.tree_clientes.delete(row)

        df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
        for _, row in df.iterrows():
            self.tree_clientes.insert("", tk.END, values=row.tolist())

    def buscar_cliente_avancada(self, termo, treeview):
        try:
            # Carregar os dados da planilha
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
            
            # Usar fuzzywuzzy para encontrar correspondências
            nomes_clientes = df["Nome"].tolist()
            resultados = process.extract(termo, nomes_clientes, limit=5)

            # Limpar o treeview antes de exibir os resultados
            treeview.delete(*treeview.get_children())

            # Adicionar os resultados no treeview
            for resultado in resultados:
                cliente = df[df["Nome"] == resultado[0]].iloc[0]
                treeview.insert("", tk.END, values=cliente.tolist())
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar cliente: {e}")

    def show_client_menu(self, event):
        item = self.tree_clientes.selection()[0]
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Editar", command=lambda: self.editar_cliente(item))
        menu.add_command(label="Excluir", command=lambda: self.excluir_cliente(item))
        menu.post(event.x_root, event.y_root)

    def editar_cliente(self, item):
        cliente_id = int(self.tree_clientes.item(item, "values")[0])
        df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
        cliente = df[df["ID"] == cliente_id]

        if cliente.empty:
            messagebox.showerror("Erro", "Cliente não encontrado.")
            return

        cliente = cliente.iloc[0]

        janela_editar = tk.Toplevel(self.root)
        janela_editar.title("Editar Cliente")
        janela_editar.geometry("400x300")

        nome_var = tk.StringVar(value=cliente["Nome"])
        cpf_var = tk.StringVar(value=cliente["CPF/CNPJ"])
        telefone_var = tk.StringVar(value=cliente["Telefone"])
        email_var = tk.StringVar(value=cliente["Email"])
        endereco_var = tk.StringVar(value=cliente["Endereço"])

        tk.Label(janela_editar, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=nome_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="CPF/CNPJ:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=cpf_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Telefone:").grid(row=2, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=telefone_var).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Email:").grid(row=3, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=email_var).grid(row=3, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Endereço:").grid(row=4, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=endereco_var).grid(row=4, column=1, padx=5, pady=5)

        def salvar_edicao():
            df.loc[df["ID"] == cliente_id, ["Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço"]] = [
                nome_var.get(), cpf_var.get(), telefone_var.get(), email_var.get(), endereco_var.get()
            ]
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Clientes", index=False)
            janela_editar.destroy()
            self.load_client_data()

        tk.Button(janela_editar, text="Salvar", command=salvar_edicao).grid(row=5, column=0, columnspan=2, pady=20)

    def excluir_cliente(self, item):
        cliente_id = int(self.tree_clientes.item(item, "values")[0])
        resposta = messagebox.askyesno("Excluir", "Tem certeza que deseja excluir este cliente?")
        if resposta:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
            if cliente_id not in df["ID"].values:
                messagebox.showerror("Erro", "Cliente não encontrado.")
                return
            df = df[df["ID"] != cliente_id]
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Clientes", index=False)
            self.load_client_data()

    def salvar_cliente(self, nome_var, cpf_var, telefone_var, email_var, endereco_var):
        try:
            if os.path.exists(EXCEL_PATH):
                df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
            else:
                df = pd.DataFrame(columns=["ID", "Nome", "CPF/CNPJ", "Telefone", "Email", "Endereço"])

            novo_id = df["ID"].max() + 1 if not df.empty else 1
            novo_cliente = pd.DataFrame({
                "ID": [novo_id],
                "Nome": [nome_var.get()],
                "CPF/CNPJ": [cpf_var.get()],
                "Telefone": [telefone_var.get()],
                "Email": [email_var.get()],
                "Endereço": [endereco_var.get()]
            })
            df = pd.concat([df, novo_cliente], ignore_index=True)
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Clientes", index=False)
            messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
            self.load_client_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar cliente:\n{e}")

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def gerenciar_produtos(self):
        self.clear_main_frame()
        self.frame_produtos = tk.Frame(self.main_frame)
        self.frame_produtos.pack(fill=tk.BOTH, expand=True)

        # Formulário de inclusão de produto
        tk.Label(self.frame_produtos, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
        nome_var = tk.StringVar()
        tk.Entry(self.frame_produtos, textvariable=nome_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.frame_produtos, text="Preço:").grid(row=1, column=0, padx=5, pady=5)
        preco_var = tk.StringVar()
        tk.Entry(self.frame_produtos, textvariable=preco_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.frame_produtos, text="Estoque:").grid(row=2, column=0, padx=5, pady=5)
        estoque_var = tk.StringVar()
        tk.Entry(self.frame_produtos, textvariable=estoque_var).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.frame_produtos, text="Ponto de Pedido:").grid(row=3, column=0, padx=5, pady=5)
        ponto_pedido_var = tk.StringVar()
        tk.Entry(self.frame_produtos, textvariable=ponto_pedido_var).grid(row=3, column=1, padx=5, pady=5)

        tk.Button(
            self.frame_produtos,
            text="Salvar",
            command=lambda: self.salvar_produto(nome_var, preco_var, estoque_var, ponto_pedido_var)
        ).grid(row=4, column=0, columnspan=2, pady=20)

        # Campo de busca avançada
        tk.Label(self.frame_produtos, text="Buscar Produto:").grid(row=5, column=0, padx=5, pady=5)
        busca_var = tk.StringVar()
        tk.Entry(self.frame_produtos, textvariable=busca_var).grid(row=5, column=1, padx=5, pady=5)
        tk.Button(
            self.frame_produtos,
            text="Buscar",
            command=lambda: self.buscar_produto_avancada(busca_var.get())
        ).grid(row=5, column=2, padx=5, pady=5)

        # Tabela com os produtos cadastrados
        self.tree_produtos = self.create_produto_table()
        self.tree_produtos.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

        self.load_produto_data()

    def create_produto_table(self):
        columns = ("ID", "Nome", "Preço", "Estoque", "Ponto de Pedido")
        tree = ttk.Treeview(self.frame_produtos, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        tree.bind("<Button-3>", self.show_produto_menu)
        return tree

    def load_produto_data(self):
        for row in self.tree_produtos.get_children():
            self.tree_produtos.delete(row)

        df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
        for _, row in df.iterrows():
            self.tree_produtos.insert("", tk.END, values=row.tolist())

    def buscar_produto_avancada(self, termo):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
            produtos = df["Nome"].tolist()
            resultados = process.extract(termo, produtos, limit=5)
            self.tree_produtos.delete(*self.tree_produtos.get_children())
            for resultado in resultados:
                produto = df[df["Nome"] == resultado[0]].iloc[0]
                self.tree_produtos.insert("", tk.END, values=produto.tolist())
        except Exception as e:
            print(f"Erro ao buscar produto: {e}")
            messagebox.showerror("Erro", f"Erro ao buscar produto: {e}")

    def show_produto_menu(self, event):
        item = self.tree_produtos.selection()[0]
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Editar", command=lambda: self.editar_produto(item))
        menu.add_command(label="Excluir", command=lambda: self.excluir_produto(item))
        menu.post(event.x_root, event.y_root)

    def editar_produto(self, item):
        produto_id = int(self.tree_produtos.item(item, "values")[0])
        df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
        produto = df[df["ID"] == produto_id]

        if produto.empty:
            messagebox.showerror("Erro", "Produto não encontrado.")
            return

        produto = produto.iloc[0]

        janela_editar = tk.Toplevel(self.root)
        janela_editar.title("Editar Produto")
        janela_editar.geometry("400x300")

        nome_var = tk.StringVar(value=produto["Nome"])
        preco_var = tk.StringVar(value=produto["Preço"])
        estoque_var = tk.StringVar(value=produto["Estoque"])
        ponto_pedido_var = tk.StringVar(value=produto["Ponto de Pedido"])

        tk.Label(janela_editar, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=nome_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Preço:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=preco_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Estoque:").grid(row=2, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=estoque_var).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(janela_editar, text="Ponto de Pedido:").grid(row=3, column=0, padx=5, pady=5)
        tk.Entry(janela_editar, textvariable=ponto_pedido_var).grid(row=3, column=1, padx=5, pady=5)

        def salvar_edicao():
            df.loc[df["ID"] == produto_id, ["Nome", "Preço", "Estoque", "Ponto de Pedido"]] = [
                nome_var.get(), float(preco_var.get()), int(estoque_var.get()), int(ponto_pedido_var.get())
            ]
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Produtos", index=False)
            janela_editar.destroy()
            self.load_produto_data()

        tk.Button(janela_editar, text="Salvar", command=salvar_edicao).grid(row=4, column=0, columnspan=2, pady=20)

    def excluir_produto(self, item):
        produto_id = int(self.tree_produtos.item(item, "values")[0])
        resposta = messagebox.askyesno("Excluir", "Tem certeza que deseja excluir este produto?")
        if resposta:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
            if produto_id not in df["ID"].values:
                messagebox.showerror("Erro", "Produto não encontrado.")
                return
            df = df[df["ID"] != produto_id]
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Produtos", index=False)
            self.load_produto_data()

    def salvar_produto(self, nome_var, preco_var, estoque_var, ponto_pedido_var):
        try:
            if os.path.exists(EXCEL_PATH):
                df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
            else:
                df = pd.DataFrame(columns=["ID", "Nome", "Preço", "Estoque", "Ponto de Pedido"])

            novo_id = df["ID"].max() + 1 if not df.empty else 1
            novo_produto = pd.DataFrame({
                "ID": [novo_id],
                "Nome": [nome_var.get()],
                "Preço": [float(preco_var.get())],
                "Estoque": [int(estoque_var.get())],
                "Ponto de Pedido": [int(ponto_pedido_var.get())]
            })
            df = pd.concat([df, novo_produto], ignore_index=True)
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Gerenciar Produtos", index=False)
            messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!")
            self.load_produto_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar produto:\n{e}")

    def gerenciar_relatorios(self):
        self.clear_main_frame()
        self.frame_relatorios = tk.Frame(self.main_frame)
        self.frame_relatorios.pack(fill=tk.BOTH, expand=True)

        tk.Button(self.frame_relatorios, text="Gerar Relatório de Clientes", command=self.gerar_relatorio_clientes).pack(pady=10)
        tk.Button(self.frame_relatorios, text="Gerar Relatório de Produtos", command=self.gerar_relatorio_produtos).pack(pady=10)
        tk.Button(self.frame_relatorios, text="Gerar Relatório de Ordens de Serviço", command=self.gerar_relatorio_ordens).pack(pady=10)
        tk.Button(self.frame_relatorios, text="Gerar Gráfico de Ordens de Serviço", command=self.gerar_grafico_ordens).pack(pady=10)
        tk.Button(self.frame_relatorios, text="Gerar Gráfico Financeiro", command=self.gerar_grafico_financeiro).pack(pady=10)

    def gerar_relatorio_clientes(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Sucesso", "Relatório de clientes gerado com sucesso!")

    def gerar_relatorio_produtos(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Sucesso", "Relatório de produtos gerado com sucesso!")

    def gerar_relatorio_ordens(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Ordens de Serviço")
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Sucesso", "Relatório de ordens de serviço gerado com sucesso!")

    def gerar_grafico_ordens(self):
        df = pd.read_excel(EXCEL_PATH, sheet_name="Ordens de Serviço")
        status_counts = df["Status"].value_counts()
        status_counts.plot(kind='bar', title='Ordens de Serviço Concluídas e em Aberto')
        plt.xlabel('Status')
        plt.ylabel('Quantidade')
        plt.show()

    def gerar_grafico_financeiro(self):
        df_ordens = pd.read_excel(EXCEL_PATH, sheet_name="Ordens de Serviço")
        df_produtos = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
        df_ordens['Data'] = pd.to_datetime(df_ordens['Data'])
        df_ordens.set_index('Data', inplace=True)
        df_produtos['Data'] = pd.to_datetime(df_produtos['Data'])
        df_produtos.set_index('Data', inplace=True)

        df_ordens['Valor'].resample('M').sum().plot(label='Receita', legend=True)
        df_produtos['Preço'].resample('M').sum().plot(label='Despesas', legend=True)
        plt.title('Gráfico Financeiro')
        plt.xlabel('Data')
        plt.ylabel('Valor')
        plt.show()

    def gerenciar_ordens(self):
        self.clear_main_frame()
        self.frame_ordens = tk.Frame(self.main_frame)
        self.frame_ordens.pack(fill=tk.BOTH, expand=True)

        # Campo de busca para cliente
        tk.Label(self.frame_ordens, text="Buscar Cliente:").grid(row=0, column=0, padx=5, pady=5)
        cliente_busca_var = tk.StringVar()
        tk.Entry(self.frame_ordens, textvariable=cliente_busca_var).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(
            self.frame_ordens,
            text="Buscar",
            command=lambda: self.buscar_cliente_para_ordem(cliente_busca_var.get())
        ).grid(row=0, column=2, padx=5, pady=5)

        # Informações do cliente (aparecem após busca)
        self.cliente_info_frame = tk.Frame(self.frame_ordens, bg="#f9f9f9", relief="solid", borderwidth=1)
        self.cliente_info_frame.grid(row=1, column=0, columnspan=3, pady=10, padx=5, sticky="nsew")

        tk.Label(self.cliente_info_frame, text="Cliente:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.cliente_nome_label = tk.Label(self.cliente_info_frame, text="N/A", anchor="w", bg="#f9f9f9")
        self.cliente_nome_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Label(self.cliente_info_frame, text="CPF/CNPJ:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cliente_cpf_label = tk.Label(self.cliente_info_frame, text="N/A", anchor="w", bg="#f9f9f9")
        self.cliente_cpf_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(self.cliente_info_frame, text="Telefone:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.cliente_telefone_label = tk.Label(self.cliente_info_frame, text="N/A", anchor="w", bg="#f9f9f9")
        self.cliente_telefone_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        tk.Label(self.cliente_info_frame, text="Endereço:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.cliente_endereco_label = tk.Label(self.cliente_info_frame, text="N/A", anchor="w", bg="#f9f9f9")
        self.cliente_endereco_label.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Descrição do Serviço
        tk.Label(self.frame_ordens, text="Descrição do Serviço:").grid(row=2, column=0, padx=5, pady=5)
        descricao_servico_var = tk.StringVar()
        tk.Entry(self.frame_ordens, textvariable=descricao_servico_var).grid(row=2, column=1, padx=5, pady=5)

        # Seleção de Peças Utilizadas
        tk.Label(self.frame_ordens, text="Peça Utilizada:").grid(row=3, column=0, padx=5, pady=5)
        self.peca_var = tk.StringVar()
        tk.Entry(self.frame_ordens, textvariable=self.peca_var).grid(row=3, column=1, padx=5, pady=5)
        tk.Button(
            self.frame_ordens,
            text="Buscar",
            command=self.mostrar_lista_pecas
        ).grid(row=3, column=2, padx=5, pady=5)

        # Valor da Peça
        tk.Label(self.frame_ordens, text="Valor da Peça:").grid(row=4, column=0, padx=5, pady=5)
        tk.Entry(self.frame_ordens, textvariable=self.valor_peca_var, state="readonly").grid(row=4, column=1, padx=5, pady=5)

        # Valor da Mão de Obra
        tk.Label(self.frame_ordens, text="Valor da Mão de Obra:").grid(row=5, column=0, padx=5, pady=5)
        mao_obra_var = tk.StringVar()
        tk.Entry(self.frame_ordens, textvariable=mao_obra_var).grid(row=5, column=1, padx=5, pady=5)

        # Status
        tk.Label(self.frame_ordens, text="Status da Ordem:").grid(row=6, column=0, padx=5, pady=5)
        status_var = tk.StringVar(value="Aberto")
        status_combobox = ttk.Combobox(
            self.frame_ordens,
            textvariable=status_var,
            values=["Aberto", "Em andamento", "Finalizado"]
        )
        status_combobox.grid(row=6, column=1, padx=5, pady=5)

        # Botão para Adicionar ao Grid
        tk.Button(
            self.frame_ordens,
            text="Adicionar ao Grid",
            command=lambda: self.adicionar_ao_grid(
                descricao_servico_var.get(),
                self.peca_var.get(),
                self.valor_peca_var.get(),
                mao_obra_var.get(),
                status_var.get()
            )
        ).grid(row=7, column=0, columnspan=2, pady=10)

        # Grid de Ordens com o campo de status
        self.grid_ordens = ttk.Treeview(
            self.frame_ordens,
            columns=(
                "Descrição do Serviço",
                "Peça Utilizada",
                "Valor da Peça",
                "Valor da Mão de Obra",
                "Status"
            ),
            show="headings"
        )
        for col in (
            "Descrição do Serviço",
            "Peça Utilizada",
            "Valor da Peça",
            "Valor da Mão de Obra",
            "Status"
        ):
            self.grid_ordens.heading(col, text=col)
            self.grid_ordens.column(col, width=150)
        self.grid_ordens.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

        # Totalizador abaixo do grid
        self.total_label = tk.Label(self.frame_ordens, text="Total Peças: R$ 0.00 \n Total Mão de Obra: R$ 0.00", font=("Arial", 12, "bold"))
        self.total_label.grid(row=9, column=0, columnspan=3, pady=10)

        # Botões de Ação
        tk.Button(
            self.frame_ordens,
            text="Salvar",
            command=self.salvar_ordens
        ).grid(row=10, column=0, padx=5, pady=10)

        tk.Button(
            self.frame_ordens,
            text="Alterar",
            command=self.alterar_ordem
        ).grid(row=10, column=1, padx=5, pady=10)

        tk.Button(
            self.frame_ordens,
            text="Excluir",
            command=self.excluir_ordem
        ).grid(row=10, column=2, padx=5, pady=10)

        self.load_ordem_data()

    def mostrar_lista_pecas(self):
        janela_lista = tk.Toplevel(self.root)
        janela_lista.title("Selecionar Peça")
        janela_lista.geometry("400x300")

        lista_pecas = tk.Listbox(janela_lista)
        lista_pecas.pack(fill=tk.BOTH, expand=True)

        df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
        for _, row in df.iterrows():
            lista_pecas.insert(tk.END, row["Nome"])

        def selecionar_peca(event):
            selecionada = lista_pecas.get(lista_pecas.curselection())
            peca = df[df["Nome"] == selecionada].iloc[0]
            self.peca_var.set(selecionada)
            self.valor_peca_var.set(peca["Preço"])
            janela_lista.destroy()

        lista_pecas.bind("<<ListboxSelect>>", selecionar_peca)

    def buscar_cliente_para_ordem(self, termo):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Clientes")
            cliente = df[df["Nome"].str.contains(termo, case=False)].iloc[0]

            # Atualizar os labels com as informações do cliente
            self.cliente_nome_label.config(text=cliente["Nome"])
            self.cliente_cpf_label.config(text=cliente["CPF/CNPJ"])
            self.cliente_telefone_label.config(text=cliente["Telefone"])
            self.cliente_endereco_label.config(text=cliente["Endereço"])

        except IndexError:
            messagebox.showwarning("Aviso", "Cliente não encontrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar cliente: {e}")

    def adicionar_ao_grid(self, descricao_servico, peca, valor_peca, mao_obra, status):
        if not descricao_servico or not peca or not valor_peca or not mao_obra or not status:
            messagebox.showwarning("Aviso", "Preencha todos os campos!")
            return

        self.grid_ordens.insert(
            "",
            "end",
            values=(descricao_servico, peca, valor_peca, mao_obra, status)
        )
        self.atualizar_totalizador()  # Atualiza o totalizador

    def atualizar_totalizador(self):
        total_pecas = 0
        total_mao_obra = 0

        for item in self.grid_ordens.get_children():
            valores = self.grid_ordens.item(item, "values")
            total_pecas += float(valores[2])  # Índice 2 é o Valor da Peça
            total_mao_obra += float(valores[3])  # Índice 3 é o Valor da Mão de Obra

        self.total_label.config(
            text=f"Total Peças: R$ {total_pecas:.2f} \n Total Mão de Obra: R$ {total_mao_obra:.2f}"
        )

    def load_ordem_data(self):
        for row in self.grid_ordens.get_children():
            self.grid_ordens.delete(row)

        df = pd.read_excel(EXCEL_PATH, sheet_name="Ordens de Serviço")
        for _, row in df.iterrows():
            self.grid_ordens.insert("", tk.END, values=row.tolist())

    def salvar_ordens(self):
        try:
            ordens = []
            for item in self.grid_ordens.get_children():
                valores = self.grid_ordens.item(item, "values")
                ordens.append(valores)

            colunas = ["Descrição do Serviço", "Peça Utilizada", "Valor da Peça", "Valor da Mão de Obra", "Status"]
            df = pd.DataFrame(ordens, columns=colunas)

            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name="Ordens de Serviço", index=False)

            messagebox.showinfo("Sucesso", "Ordens de serviço salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar ordens de serviço: {e}")

    def alterar_ordem(self):
        try:
            item_selecionado = self.grid_ordens.selection()
            if not item_selecionado:
                messagebox.showwarning("Aviso", "Nenhuma ordem selecionada para alterar.")
                return

            valores = self.grid_ordens.item(item_selecionado, "values")

            # Janela para edição
            janela_editar = tk.Toplevel(self.root)
            janela_editar.title("Alterar Ordem de Serviço")
            janela_editar.geometry("400x300")

            # Campos para edição
            descricao_var = tk.StringVar(value=valores[0])
            peca_var = tk.StringVar(value=valores[1])
            valor_peca_var = tk.StringVar(value=valores[2])
            mao_obra_var = tk.StringVar(value=valores[3])
            status_var = tk.StringVar(value=valores[4])

            tk.Label(janela_editar, text="Descrição do Serviço:").grid(row=0, column=0, padx=5, pady=5)
            tk.Entry(janela_editar, textvariable=descricao_var).grid(row=0, column=1, padx=5, pady=5)

            tk.Label(janela_editar, text="Peça Utilizada:").grid(row=1, column=0, padx=5, pady=5)
            tk.Entry(janela_editar, textvariable=peca_var).grid(row=1, column=1, padx=5, pady=5)

            tk.Label(janela_editar, text="Valor da Peça:").grid(row=2, column=0, padx=5, pady=5)
            tk.Entry(janela_editar, textvariable=valor_peca_var).grid(row=2, column=1, padx=5, pady=5)

            tk.Label(janela_editar, text="Valor da Mão de Obra:").grid(row=3, column=0, padx=5, pady=5)
            tk.Entry(janela_editar, textvariable=mao_obra_var).grid(row=3, column=1, padx=5, pady=5)

            tk.Label(janela_editar, text="Status:").grid(row=4, column=0, padx=5, pady=5)
            ttk.Combobox(
                janela_editar, textvariable=status_var, values=["Aberto", "Em andamento", "Finalizado"]
            ).grid(row=4, column=1, padx=5, pady=5)

            # Botão para salvar as alterações
            tk.Button(
                janela_editar,
                text="Salvar",
                command=lambda: self.confirmar_alteracao(
                    item_selecionado, descricao_var.get(), peca_var.get(),
                    valor_peca_var.get(), mao_obra_var.get(), status_var.get(),
                    janela_editar
                )
            ).grid(row=5, column=0, columnspan=2, pady=20)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao alterar ordem: {e}")

    def confirmar_alteracao(self, item, descricao, peca, valor_peca, mao_obra, status, janela_editar):
        self.grid_ordens.item(item, values=(descricao, peca, valor_peca, mao_obra, status))
        self.atualizar_totalizador()
        janela_editar.destroy()

    def excluir_ordem(self):
        try:
            item_selecionado = self.grid_ordens.selection()
            if not item_selecionado:
                messagebox.showwarning("Aviso", "Nenhuma ordem selecionada para excluir.")
                return

            self.grid_ordens.delete(item_selecionado)
            self.atualizar_totalizador()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir ordem: {e}")

    def configuracoes(self):
        self.clear_main_frame()
        self.frame_configuracoes = tk.Frame(self.main_frame)
        self.frame_configuracoes.pack(fill=tk.BOTH, expand=True)

        # Botões de configuração
        tk.Button(self.frame_configuracoes, text="Configuração de Email", command=self.configuracao_email, bg="#16a085", fg="white", font=("Arial", 12)).pack(fill=tk.X, padx=10, pady=5)
        tk.Button(self.frame_configuracoes, text="Alterar Wallpaper", command=self.salvar_wallpaper, bg="#16a085", fg="white", font=("Arial", 12)).pack(fill=tk.X, padx=10, pady=5)
        tk.Button(self.frame_configuracoes, text="Alterar Logo", command=self.salvar_logo, bg="#16a085", fg="white", font=("Arial", 12)).pack(fill=tk.X, padx=10, pady=5)

    def configuracao_email(self):
        self.clear_main_frame()
        self.frame_email = tk.Frame(self.main_frame)
        self.frame_email.pack(fill=tk.BOTH, expand=True)

        # Campos de configuração de email
        tk.Label(self.frame_email, text="Email Remetente:").grid(row=0, column=0, padx=5, pady=5)
        email_remetente_var = tk.StringVar()
        tk.Entry(self.frame_email, textvariable=email_remetente_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.frame_email, text="Senha:").grid(row=1, column=0, padx=5, pady=5)
        senha_var = tk.StringVar()
        tk.Entry(self.frame_email, textvariable=senha_var, show="*").grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.frame_email, text="Servidor SMTP:").grid(row=2, column=0, padx=5, pady=5)
        servidor_smtp_var = tk.StringVar(value="smtp.gmail.com")
        tk.Entry(self.frame_email, textvariable=servidor_smtp_var).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.frame_email, text="Porta SMTP:").grid(row=3, column=0, padx=5, pady=5)
        porta_smtp_var = tk.StringVar(value="587")
        tk.Entry(self.frame_email, textvariable=porta_smtp_var).grid(row=3, column=1, padx=5, pady=5)

        tk.Label(self.frame_email, text="Destinatário para Notificações:").grid(row=4, column=0, padx=5, pady=5)
        destinatario_var = tk.StringVar()
        tk.Entry(self.frame_email, textvariable=destinatario_var).grid(row=4, column=1, padx=5, pady=5)

        tk.Button(
            self.frame_email,
            text="Salvar Configurações",
            command=lambda: self.salvar_configuracoes(email_remetente_var, senha_var, servidor_smtp_var, porta_smtp_var, destinatario_var)
        ).grid(row=5, column=0, columnspan=2, pady=20)

    def salvar_configuracoes(self, email_remetente_var, senha_var, servidor_smtp_var, porta_smtp_var, destinatario_var):
        try:
            df_config = pd.DataFrame({
                "Configuração": ["Email Remetente", "Senha", "Servidor SMTP", "Porta SMTP", "Destinatário para Notificações"],
                "Valor": [email_remetente_var.get(), senha_var.get(), servidor_smtp_var.get(), porta_smtp_var.get(), destinatario_var.get()]
            })
            with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
                df_config.to_excel(writer, sheet_name="Configurações", index=False)
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configurações:\n{e}")

    def on_closing(self):
        self.root.destroy()

    def buscar_pecas(self, termo):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name="Gerenciar Produtos")
            pecas = df["Nome"].tolist()
            resultados = process.extract(termo, pecas, limit=5)
            # Aqui você pode adicionar lógica para mostrar os resultados da busca
            print(resultados)
        except Exception as e:
            print(f"Erro ao buscar peças: {e}")
            messagebox.showerror("Erro", f"Erro ao buscar peças: {e}")

    def enviar_email(self, destinatario, assunto, mensagem):
        try:
            df_config = pd.read_excel(EXCEL_PATH, sheet_name="Configurações")
            email_remetente = df_config.loc[df_config["Configuração"] == "Email Remetente", "Valor"].values[0]
            senha = df_config.loc[df_config["Configuração"] == "Senha", "Valor"].values[0]
            servidor_smtp = df_config.loc[df_config["Configuração"] == "Servidor SMTP", "Valor"].values[0]
            porta_smtp = df_config.loc[df_config["Configuração"] == "Porta SMTP", "Valor"].values[0]

            msg = MIMEText(mensagem)
            msg["Subject"] = assunto
            msg["From"] = email_remetente
            msg["To"] = destinatario

            with smtplib.SMTP(servidor_smtp, porta_smtp) as server:
                server.starttls()
                server.login(email_remetente, senha)
                server.sendmail(email_remetente, destinatario, msg.as_string())

            messagebox.showinfo("Sucesso", "Email enviado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar email: {e}")

# Iniciar o aplicativo
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
