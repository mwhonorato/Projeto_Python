# -*- coding: utf-8 -*-
"""
Created on Wed Nov 13 08:07:50 2024

@author: tistahl
"""
import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from PIL import Image, ImageTk

# Conectar ao banco de dados SQL Server
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=localhost;'
    'DATABASE=Oficina;'
    'UID=sa;'
    'PWD=Kkh@501350'
)
cursor = conn.cursor()

# Função para salvar dados no banco de dados
def save_to_db(table, data):
    placeholders = ', '.join(['?'] * len(data))
    cursor.execute(f'INSERT INTO {table} VALUES ({placeholders})', data)
    conn.commit()

# Função para atualizar dados no banco de dados
def update_db(table, data, primary_key, key_value):
    set_clause = ', '.join([f'{col} = ?' for col in data.keys()])
    cursor.execute(f'UPDATE {table} SET {set_clause} WHERE {primary_key} = ?', list(data.values()) + [key_value])
    conn.commit()

# Função para excluir dados no banco de dados
def delete_from_db(table, primary_key, key_value):
    cursor.execute(f'DELETE FROM {table} WHERE {primary_key} = ?', (key_value,))
    conn.commit()

# Função para validar os campos antes de salvar os dados
def validate_entries(entries):
    for label, entry in entries.items():
        if not entry.get():
            messagebox.showerror("Erro", f"O campo '{label}' não pode estar vazio.")
            return False
    return True

# Cadastro de Produtos
labels_produtos = ["Código do Produto", "Tipo do Produto", "Nome da Peça", "Quantidade", "Data da Compra", "Fornecedor do Produto", "Número da Nota Fiscal", "Ponto de Pedido"]
entries_produtos = {}
for i, label in enumerate(labels_produtos):
    tk.Label(frames['Cadastro de Produtos'], text=label, anchor='w').grid(row=i, column=0, sticky='w')
    entry = tk.Entry(frames['Cadastro de Produtos'])
    entry.grid(row=i, column=1, sticky='w')
    entries_produtos[label] = entry

def save_produto():
    if validate_entries(entries_produtos):
        data = (
            entries_produtos["Código do Produto"].get(),
            entries_produtos["Tipo do Produto"].get(),
            entries_produtos["Nome da Peça"].get(),
            entries_produtos["Quantidade"].get(),
            entries_produtos["Data da Compra"].get(),
            entries_produtos["Fornecedor do Produto"].get(),
            entries_produtos["Número da Nota Fiscal"].get(),
            entries_produtos["Ponto de Pedido"].get()
        )
        save_to_db('Produtos', data)
        messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso!")
        update_product_list()

def update_produto():
    if validate_entries(entries_produtos):
        data = {
            "TipoProduto": entries_produtos["Tipo do Produto"].get(),
            "NomePeca": entries_produtos["Nome da Peça"].get(),
            "Quantidade": entries_produtos["Quantidade"].get(),
            "DataCompra": entries_produtos["Data da Compra"].get(),
            "FornecedorProduto": entries_produtos["Fornecedor do Produto"].get(),
            "NumeroNotaFiscal": entries_produtos["Número da Nota Fiscal"].get(),
            "PontoPedido": entries_produtos["Ponto de Pedido"].get()
        }
        update_db('Produtos', data, "CodigoProduto", entries_produtos["Código do Produto"].get())
        messagebox.showinfo("Sucesso", "Produto atualizado com sucesso!")
        update_product_list()

def delete_produto():
    delete_from_db('Produtos', "CodigoProduto", entries_produtos["Código do Produto"].get())
    messagebox.showinfo("Sucesso", "Produto excluído com sucesso!")
    update_product_list()

def on_product_select(event):
    selected_item = product_table.selection()[0]
    values = product_table.item(selected_item, 'values')
    for i, label in enumerate(labels_produtos):
        entries_produtos[label].delete(0, tk.END)
        entries_produtos[label].insert(0, values[i])

tk.Button(frames['Cadastro de Produtos'], text="Salvar", command=save_produto, bg='blue', fg='white').grid(row=len(labels_produtos), column=0, sticky='w', pady=10)
tk.Button(frames['Cadastro de Produtos'], text="Alterar", command=update_produto, bg='green', fg='white').grid(row=len(labels_produtos), column=1, sticky='w', pady=10)
tk.Button(frames['Cadastro de Produtos'], text="Excluir", command=delete_produto, bg='red', fg='white').grid(row=len(labels_produtos), column=2, sticky='w', pady=10)

product_table = ttk.Treeview(frames['Cadastro de Produtos'], columns=labels_produtos, show='headings')
for label in labels_produtos:
    product_table.heading(label, text=label)
    product_table.column(label, anchor='center')
product_table.grid(row=len(labels_produtos)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
product_table.bind('<ButtonRelease-1>', on_product_select)

def update_product_list():
    for item in product_table.get_children():
        product_table.delete(item)
    cursor.execute('SELECT * FROM Produtos')
    for row in cursor.fetchall():
        product_table.insert('', 'end', values=row)

update_product_list()

# Repita o mesmo processo para as outras tabelas (Ferramentas, Clientes, Ordens de Serviço, Estoque, Notas Fiscais)

# Fechar a conexão ao banco de dados ao finalizar
root.protocol("WM_DELETE_WINDOW", lambda: (conn.close(), root.destroy()))

# Resto do código da interface gráfica permanece o mesmo
# Cadastro de Ferramentas
labels_ferramentas = ["Tipo de Ferramenta", "Modelo da Ferramenta", "Data da Compra", "Fornecedor", "Quantidade"]
entries_ferramentas = {}
for i, label in enumerate(labels_ferramentas):
    tk.Label(frames['Cadastro de Ferramentas'], text=label, anchor='w').grid(row=i, column=0, sticky='w')
    entry = tk.Entry(frames['Cadastro de Ferramentas'])
    entry.grid(row=i, column=1, sticky='w')
    entries_ferramentas[label] = entry

def save_ferramenta():
    if validate_entries(entries_ferramentas):
        data = (
            entries_ferramentas["Tipo de Ferramenta"].get(),
            entries_ferramentas["Modelo da Ferramenta"].get(),
            entries_ferramentas["Data da Compra"].get(),
            entries_ferramentas["Fornecedor"].get(),
            entries_ferramentas["Quantidade"].get()
        )
        save_to_db('Ferramentas', data)
        messagebox.showinfo("Sucesso", "Ferramenta cadastrada com sucesso!")
        update_tool_list()

def update_ferramenta():
    if validate_entries(entries_ferramentas):
        data = {
            "DataCompra": entries_ferramentas["Data da Compra"].get(),
            "Fornecedor": entries_ferramentas["Fornecedor"].get(),
            "Quantidade": entries_ferramentas["Quantidade"].get()
        }
        update_db('Ferramentas', data, "TipoFerramenta", entries_ferramentas["Tipo de Ferramenta"].get())
        messagebox.showinfo("Sucesso", "Ferramenta atualizada com sucesso!")
        update_tool_list()

def delete_ferramenta():
    delete_from_db('Ferramentas', "TipoFerramenta", entries_ferramentas["Tipo de Ferramenta"].get())
    messagebox.showinfo("Sucesso", "Ferramenta excluída com sucesso!")
    update_tool_list()

def on_tool_select(event):
    selected_item = tool_table.selection()[0]
    values = tool_table.item(selected_item, 'values')
    for i, label in enumerate(labels_ferramentas):
        entries_ferramentas[label].delete(0, tk.END)
        entries_ferramentas[label].insert(0, values[i])

tk.Button(frames['Cadastro de Ferramentas'], text="Salvar", command=save_ferramenta, bg='blue', fg='white').grid(row=len(labels_ferramentas), column=0, sticky='w', pady=10)
tk.Button(frames['Cadastro de Ferramentas'], text="Alterar", command=update_ferramenta, bg='green', fg='white').grid(row=len(labels_ferramentas), column=1, sticky='w', pady=10)
tk.Button(frames['Cadastro de Ferramentas'], text="Excluir", command=delete_ferramenta, bg='red', fg='white').grid(row=len(labels_ferramentas), column=2, sticky='w', pady=10)

tool_table = ttk.Treeview(frames['Cadastro de Ferramentas'], columns=labels_ferramentas, show='headings')
for label in labels_ferramentas:
    tool_table.heading(label, text=label)
    tool_table.column(label, anchor='center')
tool_table.grid(row=len(labels_ferramentas)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
tool_table.bind('<ButtonRelease-1>', on_tool_select)

def update_tool_list():
    for item in tool_table.get_children():
        tool_table.delete(item)
    cursor.execute('SELECT * FROM Ferramentas')
    for row in cursor.fetchall():
        tool_table.insert('', 'end', values=row)

update_tool_list()

# Cadastro de Clientes
cliente_codigo = 1
labels_clientes = ["Código do Cliente", "Nome do Cliente", "CPF/CNPJ", "Endereço", "Número do Telefone"]
entries_clientes = {}
for i, label in enumerate(labels_clientes):
    tk.Label(frames['Cadastro de Clientes'], text=label).grid(row=i, column=0, sticky='w')
    if label == "Código do Cliente":
        entry = tk.Label(frames['Cadastro de Clientes'], text=f"{cliente_codigo:07d}")
    else:
        entry = tk.Entry(frames['Cadastro de Clientes'])
        entry.grid(row=i, column=1, sticky='w')
    entries_clientes[label] = entry

def save_cliente():
    if validate_entries(entries_clientes):
        global cliente_codigo
        data = (
            cliente_codigo,
            entries_clientes["Nome do Cliente"].get(),
            entries_clientes["CPF/CNPJ"].get(),
            entries_clientes["Endereço"].get(),
            entries_clientes["Número do Telefone"].get()
        )
        save_to_db('Clientes', data)
        messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
        cliente_codigo += 1
        entries_clientes["Código do Cliente"].config(text=f"{cliente_codigo:07d}")
        update_client_list()

def update_cliente():
    if validate_entries(entries_clientes):
        data = {
            "NomeCliente": entries_clientes["Nome do Cliente"].get(),
            "CPFCNPJ": entries_clientes["CPF/CNPJ"].get(),
            "Endereco": entries_clientes["Endereço"].get(),
            "NumeroTelefone": entries_clientes["Número do Telefone"].get()
        }
        update_db('Clientes', data, "CodigoCliente", entries_clientes["Código do Cliente"].cget("text"))
        messagebox.showinfo("Sucesso", "Cliente atualizado com sucesso!")
        update_client_list()

def delete_cliente():
    delete_from_db('Clientes', "CodigoCliente", entries_clientes["Código do Cliente"].cget("text"))
    messagebox.showinfo("Sucesso", "Cliente excluído com sucesso!")
    update_client_list()

def on_client_select(event):
    selected_item = client_table.selection()[0]
    values = client_table.item(selected_item, 'values')
    for i, label in enumerate(labels_clientes):
        if label == "Código do Cliente":
            entries_clientes[label].config(text=values[i])
        else:
            entries_clientes[label].delete(0, tk.END)
            entries_clientes[label].insert(0, values[i])

tk.Button(frames['Cadastro de Clientes'], text="Salvar", command=save_cliente, bg='blue', fg='white').grid(row=len(labels_clientes), column=0, sticky='w', pady=10)
tk.Button(frames['Cadastro de Clientes'], text="Alterar", command=update_cliente, bg='green', fg='white').grid(row=len(labels_clientes), column=1, sticky='w', pady=10)
tk.Button(frames['Cadastro de Clientes'], text="Excluir", command=delete_cliente, bg='red', fg='white').grid(row=len(labels_clientes), column=2, sticky='w', pady=10)

client_table = ttk.Treeview(frames['Cadastro de Clientes'], columns=labels_clientes, show='headings')
for label in labels_clientes:
    client_table.heading(label, text=label)
    client_table.column(label, anchor='center')
client_table.grid(row=len(labels_clientes)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
client_table.bind('<ButtonRelease-1>', on_client_select)

def update_client_list():
    for item in client_table.get_children():
        client_table.delete(item)
    cursor.execute('SELECT * FROM Clientes')
    for row in cursor.fetchall():
        client_table.insert('', 'end', values=row)

update_client_list()

# Inclusão de Ordens de Serviço
labels_ordens_servico = ["Número da Ordem de Serviço", "Código do Cliente", "Nome do Cliente", "CPF/CNPJ", "Endereço", "Número de Telefone", "Serviço Realizado", "Código do Produto", "Quantidade Utilizada"]
entries_ordens_servico = {}
for i, label in enumerate(labels_ordens_servico):
    tk.Label(frames['Inclusão de Ordens de Serviço'], text=label, anchor='w').grid(row=i, column=0, sticky='w')
    entry = tk.Entry(frames['Inclusão de Ordens de Serviço'])
    entry.grid(row=i, column=1, sticky='w')
    entries_ordens_servico[label] = entry

def buscar_cliente():
    codigo = entries_ordens_servico["Código do Cliente"].get()
    cursor.execute('SELECT * FROM Clientes WHERE CodigoCliente = ?', (codigo,))
    row = cursor.fetchone()
    if row:
        entries_ordens_servico["Nome do Cliente"].delete(0, tk.END)
        entries_ordens_servico["Nome do Cliente"].insert(0, row[1])
        entries_ordens_servico["CPF/CNPJ"].delete(0, tk.END)
        entries_ordens_servico["CPF/CNPJ"].insert(0, row[2])
        entries_ordens_servico["Endereço"].delete(0, tk.END)
        entries_ordens_servico["Endereço"].insert(0, row[3])
        entries_ordens_servico["Número de Telefone"].delete(0, tk.END)
        entries_ordens_servico["Número de Telefone"].insert(0, row[4])

entries_ordens_servico["Código do Cliente"].bind("<FocusOut>", lambda e: buscar_cliente())


def atualizar_quantidade_produto(codigo, quantidade_usada):
    cursor.execute('SELECT Quantidade FROM Produtos WHERE CodigoProduto = ?', (codigo,))
    row = cursor.fetchone()
    if row:
        nova_quantidade = row[0] - quantidade_usada
        cursor.execute('UPDATE Produtos SET Quantidade = ? WHERE CodigoProduto = ?', (nova_quantidade, codigo))
        conn.commit()

def save_ordem_servico():
    if validate_entries(entries_ordens_servico):
        data = (
            entries_ordens_servico["Número da Ordem de Serviço"].get(),
            entries_ordens_servico["Código do Cliente"].get(),
            entries_ordens_servico["Nome do Cliente"].get(),
            entries_ordens_servico["CPF/CNPJ"].get(),
            entries_ordens_servico["Endereço"].get(),
            entries_ordens_servico["Número de Telefone"].get(),
            entries_ordens_servico["Serviço Realizado"].get(),
            entries_ordens_servico["Código do Produto"].get(),
            entries_ordens_servico["Quantidade Utilizada"].get()
        )
        save_to_db('OrdensServico', data)
        atualizar_quantidade_produto(entries_ordens_servico["Código do Produto"].get(), int(entries_ordens_servico["Quantidade Utilizada"].get()))
        messagebox.showinfo("Sucesso", "Ordem de Serviço cadastrada com sucesso!")
        update_service_order_list()

def update_ordem_servico():
    if validate_entries(entries_ordens_servico):
        data = {
            "CodigoCliente": entries_ordens_servico["Código do Cliente"].get(),
            "NomeCliente": entries_ordens_servico["Nome do Cliente"].get(),
            "CPFCNPJ": entries_ordens_servico["CPF/CNPJ"].get(),
            "Endereco": entries_ordens_servico["Endereço"].get(),
            "NumeroTelefone": entries_ordens_servico["Número de Telefone"].get(),
            "ServicoRealizado": entries_ordens_servico["Serviço Realizado"].get(),
            "CodigoProduto": entries_ordens_servico["Código do Produto"].get(),
            "QuantidadeUtilizada": entries_ordens_servico["Quantidade Utilizada"].get()
        }
        update_db('OrdensServico', data, "NumeroOrdemServico", entries_ordens_servico["Número da Ordem de Serviço"].get())
        messagebox.showinfo("Sucesso", "Ordem de Serviço atualizada com sucesso!")
        update_service_order_list()

def delete_ordem_servico():
    delete_from_db('OrdensServico', "NumeroOrdemServico", entries_ordens_servico["Número da Ordem de Serviço"].get())
    messagebox.showinfo("Sucesso", "Ordem de Serviço excluída com sucesso!")
    update_service_order_list()

def on_service_order_select(event):
    selected_item = service_order_table.selection()[0]
    values = service_order_table.item(selected_item, 'values')
    for i, label in enumerate(labels_ordens_servico):
        entries_ordens_servico[label].delete(0, tk.END)
        entries_ordens_servico[label].insert(0, values[i])

tk.Button(frames['Inclusão de Ordens de Serviço'], text="Salvar", command=save_ordem_servico, bg='blue', fg='white').grid(row=len(labels_ordens_servico), column=0, sticky='w', pady=10)
tk.Button(frames['Inclusão de Ordens de Serviço'], text="Alterar", command=update_ordem_servico, bg='green', fg='white').grid(row=len(labels_ordens_servico), column=1, sticky='w', pady=10)
tk.Button(frames['Inclusão de Ordens de Serviço'], text="Excluir", command=delete_ordem_servico, bg='red', fg='white').grid(row=len(labels_ordens_servico), column=2, sticky='w', pady=10)

service_order_table = ttk.Treeview(frames['Inclusão de Ordens de Serviço'], columns=labels_ordens_servico, show='headings')
for label in labels_ordens_servico:
    service_order_table.heading(label, text=label)
    service_order_table.column(label, anchor='center')
service_order_table.grid(row=len(labels_ordens_servico)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
service_order_table.bind('<ButtonRelease-1>', on_service_order_select)

def update_service_order_list():
    for item in service_order_table.get_children():
        service_order_table.delete(item)
    cursor.execute('SELECT * FROM OrdensServico')
    for row in cursor.fetchall():
        service_order_table.insert('', 'end', values=row)

update_service_order_list()

# Controle de Estoque
labels_estoque = ["Código do Produto", "Nome do Produto", "Quantidade em Estoque", "Ponto de Pedido"]
entries_estoque = {}
for i, label in enumerate(labels_estoque):
    tk.Label(frames['Controle de Estoque'], text=label, anchor='w').grid(row=i, column=0, sticky='w')
    entry = tk.Entry(frames['Controle de Estoque'])
    entry.grid(row=i, column=1, sticky='w')
    entries_estoque[label] = entry

def save_estoque():
    if validate_entries(entries_estoque):
        data = (
            entries_estoque["Código do Produto"].get(),
            entries_estoque["Nome do Produto"].get(),
            entries_estoque["Quantidade em Estoque"].get(),
            entries_estoque["Ponto de Pedido"].get()
        )
        save_to_db('Estoque', data)
        messagebox.showinfo("Sucesso", "Estoque atualizado com sucesso!")
        update_stock_list()

def update_estoque():
    if validate_entries(entries_estoque):
        data = {
            "NomeProduto": entries_estoque["Nome do Produto"].get(),
            "QuantidadeEstoque": entries_estoque["Quantidade em Estoque"].get(),
            "PontoPedido": entries_estoque["Ponto de Pedido"].get()
        }
        update_db('Estoque', data, "CodigoProduto", entries_estoque["Código do Produto"].get())
        messagebox.showinfo("Sucesso", "Estoque atualizado com sucesso!")
        update_stock_list()

def delete_estoque():
    delete_from_db('Estoque', "CodigoProduto", entries_estoque["Código do Produto"].get())
    messagebox.showinfo("Sucesso", "Estoque excluído com sucesso!")
    update_stock_list()

def on_stock_select(event):
    selected_item = stock_table.selection()[0]
    values = stock_table.item(selected_item, 'values')
    for i, label in enumerate(labels_estoque):
        entries_estoque[label].delete(0, tk.END)
        entries_estoque[label].insert(0, values[i])

tk.Button(frames['Controle de Estoque'], text="Salvar", command=save_estoque, bg='blue', fg='white').grid(row=len(labels_estoque), column=0, sticky='w', pady=10)
tk.Button(frames['Controle de Estoque'], text="Alterar", command=update_estoque, bg='green', fg='white').grid(row=len(labels_estoque), column=1, sticky='w', pady=10)
tk.Button(frames['Controle de Estoque'], text="Excluir", command=delete_estoque, bg='red', fg='white').grid(row=len(labels_estoque), column=2, sticky='w', pady=10)

stock_table = ttk.Treeview(frames['Controle de Estoque'], columns=labels_estoque, show='headings')
for label in labels_estoque:
    stock_table.heading(label, text=label)
    stock_table.column(label, anchor='center')
stock_table.grid(row=len(labels_estoque)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
stock_table.bind('<ButtonRelease-1>', on_stock_select)

def update_stock_list():
    for item in stock_table.get_children():
        stock_table.delete(item)
    cursor.execute('SELECT * FROM Estoque')
    for row in cursor.fetchall():
        stock_table.insert('', 'end', values=row)

update_stock_list()

# Notas Fiscais
labels_notas_fiscais = ["Número da Nota Fiscal", "Data de Emissão", "Código do Cliente", "Nome do Cliente", "Valor Total"]
entries_notas_fiscais = {}
for i, label in enumerate(labels_notas_fiscais):
    tk.Label(frames['Notas Fiscais'], text=label, anchor='w').grid(row=i, column=0, sticky='w')
    entry = tk.Entry(frames['Notas Fiscais'])
    entry.grid(row=i, column=1, sticky='w')
    entries_notas_fiscais[label] = entry

def save_nota_fiscal():
    if validate_entries(entries_notas_fiscais):
        data = (
            entries_notas_fiscais["Número da Nota Fiscal"].get(),
            entries_notas_fiscais["Data de Emissão"].get(),
            entries_notas_fiscais["Código do Cliente"].get(),
            entries_notas_fiscais["Nome do Cliente"].get(),
            entries_notas_fiscais["Valor Total"].get()
        )
        save_to_db('NotasFiscais', data)
        messagebox.showinfo("Sucesso", "Nota Fiscal cadastrada com sucesso!")
        update_invoice_list()

def update_nota_fiscal():
    if validate_entries(entries_notas_fiscais):
        data = {
            "DataEmissao": entries_notas_fiscais["Data de Emissão"].get(),
            "CodigoCliente": entries_notas_fiscais["Código do Cliente"].get(),
            "NomeCliente": entries_notas_fiscais["Nome do Cliente"].get(),
            "ValorTotal": entries_notas_fiscais["Valor Total"].get()
        }
        update_db('NotasFiscais', data, "NumeroNotaFiscal", entries_notas_fiscais["Número da Nota Fiscal"].get())
        messagebox.showinfo("Sucesso", "Nota Fiscal atualizada com sucesso!")
        update_invoice_list()

def delete_nota_fiscal():
    delete_from_db('NotasFiscais', "NumeroNotaFiscal", entries_notas_fiscais["Número da Nota Fiscal"].get())
    messagebox.showinfo("Sucesso", "Nota Fiscal excluída com sucesso!")
    update_invoice_list()

def on_invoice_select(event):
    selected_item = invoice_table.selection()[0]
    values = invoice_table.item(selected_item, 'values')
    for i, label in enumerate(labels_notas_fiscais):
        entries_notas_fiscais[label].delete(0, tk.END)
        entries_notas_fiscais[label].insert(0, values[i])

tk.Button(frames['Notas Fiscais'], text="Salvar", command=save_nota_fiscal, bg='blue', fg='white').grid(row=len(labels_notas_fiscais), column=0, sticky='w', pady=10)
tk.Button(frames['Notas Fiscais'], text="Alterar", command=update_nota_fiscal, bg='green', fg='white').grid(row=len(labels_notas_fiscais), column=1, sticky='w', pady=10)
tk.Button(frames['Notas Fiscais'], text="Excluir", command=delete_nota_fiscal, bg='red', fg='white').grid(row=len(labels_notas_fiscais), column=2, sticky='w', pady=10)

invoice_table = ttk.Treeview(frames['Notas Fiscais'], columns=labels_notas_fiscais, show='headings')
for label in labels_notas_fiscais:
    invoice_table.heading(label, text=label)
    invoice_table.column(label, anchor='center')
invoice_table.grid(row=len(labels_notas_fiscais)+1, column=0, columnspan=3, pady=10, padx=10, sticky='nsew')
invoice_table.bind('<ButtonRelease-1>', on_invoice_select)

def update_invoice_list():
    for item in invoice_table.get_children():
        invoice_table.delete(item)
    cursor.execute('SELECT * FROM NotasFiscais')
    for row in cursor.fetchall():
        invoice_table.insert('', 'end', values=row)

update_invoice_list()

# Aba Início
def update_time():
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    time_label.config(text=current_time)
    frames['Início'].after(1000, update_time)

def set_background():
    file_path = filedialog.askopenfilename()
    if file_path:
        image = Image.open(file_path)
        screen_width, screen_height = get_screen_resolution()
        image = image.resize((screen_width, screen_height), Image.LANCZOS)
        background_image = ImageTk.PhotoImage(image)
        background_label.config(image=background_image)
        background_label.image = background_image

frames['Início'].pack(fill='both', expand=True)
time_label = tk.Label(frames['Início'], font=("Helvetica", 16))
time_label.pack(pady=20)
update_time()

background_button = tk.Button(frames['Início'], text="Selecionar Papel de Fundo", command=set_background)
background_button.pack(pady=10)

background_label = tk.Label(frames['Início'])
background_label.pack(fill='both', expand=True)

# Fechar a conexão ao banco de dados ao finalizar
root.protocol("WM_DELETE_WINDOW", lambda: (conn.close(), root.destroy()))

# Executa a aplicação
root.mainloop()
			
