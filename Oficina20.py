# -*- coding: utf-8 -*-
"""
Created on Thu Nov 14 10:40:34 2024

@author: tistahl
"""

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import filedialog
from datetime import datetime
import pyodbc


# Connect to the database
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=INF_01\\LOCAL;'
    'DATABASE=Oficina;'
    'UID=sa;'
    'PWD=Kkh@501350'
)
cursor = conn.cursor()


def validate_entries(entries):
    for label, entry in entries.items():
        if not entry.get():
            messagebox.showerror("Error", f"The field '{label}' cannot be empty.")
            return False
    # Add specific validations by field here
    if label == "CPF/CNPJ" and not validate_cpf_cnpj(entry.get()):
        messagebox.showerror("Error", f"The field '{label}' is invalid.")
        return False
    return True


def validate_cpf_cnpj(value):
    # Implement CPF/CNPJ validation here
    return True


def save_to_db(table, data):
    placeholders = ', '.join(['?'] * len(data))
    query = f'INSERT INTO {table} VALUES ({placeholders})'
    try:
        cursor.execute(query, data)
        conn.commit()
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Error", f"Error saving data: {e}")


def update_db(table, data, primary_key, key_value):
    set_clause = ', '.join([f'{col} = ?' for col in data.keys()])
    query = f'UPDATE {table} SET {set_clause} WHERE {primary_key} = ?'
    try:
        cursor.execute(query, list(data.values()) + [key_value])
        conn.commit()
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Error", f"Error updating data: {e}")


def delete_from_db(table, primary_key, key_value):
    query = f'DELETE FROM {table} WHERE {primary_key} = ?'
    try:
        cursor.execute(query, (key_value,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        messagebox.showerror("Error", f"Error deleting data: {e}")


# ------------------- Product Registration -------------------
def create_product_window():
    product_window = tk.Toplevel(root)
    product_window.title("Product Registration")

    labels_produtos = [
        "Código do Produto",
        "Tipo do Produto",
        "Nome da Peça",
        "Quantidade",
        "Data da Compra",
        "Fornecedor do Produto",
        "Número da Nota Fiscal",
        "Ponto de Pedido",
    ]
    entries_produtos = {}
    for i, label in enumerate(labels_produtos):
        tk.Label(product_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(product_window)
        entry.grid(row=i, column=1, sticky="w")
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
                entries_produtos["Ponto de Pedido"].get(),
            )
            save_to_db('Produtos', data)
            messagebox.showinfo("Success", "Product registered successfully!")
            update_product_list()
            product_window.destroy()

    def update_produto():
        if validate_entries(entries_produtos):
            data = {
                "TipoProduto": entries_produtos["Tipo do Produto"].get(),
                "NomePeca": entries_produtos["Nome da Peça"].get(),
                "Quantidade": entries_produtos["Quantidade"].get(),
                "DataCompra": entries_produtos["Data da Compra"].get(),
                "FornecedorProduto": entries_produtos["Fornecedor do Produto"].get(),
                "NumeroNotaFiscal": entries_produtos["Número da Nota Fiscal"].get(),
                "PontoPedido": entries_produtos["Ponto de Pedido"].get(),
            }
            update_db('Produtos', data, 'CodigoProduto', entries_produtos["Código do Produto"].get())
            messagebox.showinfo("Success", "Product updated successfully!")
            update_product_list()
            product_window.destroy()

    def delete_produto():
        codigo_produto = entries_produtos["Código do Produto"].get()
        if codigo_produto:
            delete_from_db('Produtos', 'CodigoProduto', codigo_produto)
            messagebox.showinfo("Success", "Product deleted successfully!")
            update_product_list()
            product_window.destroy()
        else:
            messagebox.showerror("Error", "Product code is required to delete.")

    tk.Button(product_window, text="Save", command=save_produto).grid(row=len(labels_produtos), column=0, sticky="w")
    tk.Button(product_window, text="Update", command=update_produto).grid(row=len(labels_produtos), column=1, sticky="w")
    tk.Button(product_window, text="Delete", command=delete_produto).grid(row=len(labels_produtos), column=2, sticky="w")

def update_product_list():
    for row in product_tree.get_children():
        product_tree.delete(row)
    cursor.execute("SELECT * FROM Produtos")
    for row in cursor.fetchall():
        product_tree.insert("", "end", values=row)

# Main window setup
root = tk.Tk()
root.title("Oficina Management System")

product_tree = ttk.Treeview(root, columns=("CodigoProduto", "TipoProduto", "NomePeca", "Quantidade", "DataCompra", "FornecedorProduto", "NumeroNotaFiscal", "PontoPedido"), show="headings")
product_tree.heading("CodigoProduto", text="Código do Produto")
product_tree.heading("TipoProduto", text="Tipo do Produto")
product_tree.heading("NomePeca", text="Nome da Peça")
product_tree.heading("Quantidade", text="Quantidade")
product_tree.heading("DataCompra", text="Data da Compra")
product_tree.heading("FornecedorProduto", text="Fornecedor do Produto")
product_tree.heading("NumeroNotaFiscal", text="Número da Nota Fiscal")
product_tree.heading("PontoPedido", text="Ponto de Pedido")
product_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Product", command=create_product_window).pack()

update_product_list()

# ------------------- Customer Registration -------------------
def create_customer_window():
    customer_window = tk.Toplevel(root)
    customer_window.title("Customer Registration")

    labels_clientes = [
        "Código do Cliente",
        "Nome do Cliente",
        "CPF/CNPJ",
        "Endereço",
        "Telefone",
        "Email",
    ]
    entries_clientes = {}
    for i, label in enumerate(labels_clientes):
        tk.Label(customer_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(customer_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_clientes[label] = entry

    def save_cliente():
        if validate_entries(entries_clientes):
            data = (
                entries_clientes["Código do Cliente"].get(),
                entries_clientes["Nome do Cliente"].get(),
                entries_clientes["CPF/CNPJ"].get(),
                entries_clientes["Endereço"].get(),
                entries_clientes["Telefone"].get(),
                entries_clientes["Email"].get(),
            )
            save_to_db('Clientes', data)
            messagebox.showinfo("Success", "Customer registered successfully!")
            update_customer_list()
            customer_window.destroy()

    def update_cliente():
        if validate_entries(entries_clientes):
            data = {
                "NomeCliente": entries_clientes["Nome do Cliente"].get(),
                "CPFCNPJ": entries_clientes["CPF/CNPJ"].get(),
                "Endereco": entries_clientes["Endereço"].get(),
                "Telefone": entries_clientes["Telefone"].get(),
                "Email": entries_clientes["Email"].get(),
            }
            update_db('Clientes', data, 'CodigoCliente', entries_clientes["Código do Cliente"].get())
            messagebox.showinfo("Success", "Customer updated successfully!")
            update_customer_list()
            customer_window.destroy()

    def delete_cliente():
        codigo_cliente = entries_clientes["Código do Cliente"].get()
        if codigo_cliente:
            delete_from_db('Clientes', 'CodigoCliente', codigo_cliente)
            messagebox.showinfo("Success", "Customer deleted successfully!")
            update_customer_list()
            customer_window.destroy()
        else:
            messagebox.showerror("Error", "Customer code is required to delete.")

    tk.Button(customer_window, text="Save", command=save_cliente).grid(row=len(labels_clientes), column=0, sticky="w")
    tk.Button(customer_window, text="Update", command=update_cliente).grid(row=len(labels_clientes), column=1, sticky="w")
    tk.Button(customer_window, text="Delete", command=delete_cliente).grid(row=len(labels_clientes), column=2, sticky="w")

def update_customer_list():
    for row in customer_tree.get_children():
        customer_tree.delete(row)
    cursor.execute("SELECT * FROM Clientes")
    for row in cursor.fetchall():
        customer_tree.insert("", "end", values=row)

# Customer list setup
customer_tree = ttk.Treeview(root, columns=("CodigoCliente", "NomeCliente", "CPFCNPJ", "Endereco", "Telefone", "Email"), show="headings")
customer_tree.heading("CodigoCliente", text="Código do Cliente")
customer_tree.heading("NomeCliente", text="Nome do Cliente")
customer_tree.heading("CPFCNPJ", text="CPF/CNPJ")
customer_tree.heading("Endereco", text="Endereço")
customer_tree.heading("Telefone", text="Telefone")
customer_tree.heading("Email", text="Email")
customer_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Customer", command=create_customer_window).pack()

update_customer_list()

# ------------------- Service Order Registration -------------------
def create_service_order_window():
    service_order_window = tk.Toplevel(root)
    service_order_window.title("Service Order Registration")

    labels_ordens_servico = [
        "Código da Ordem de Serviço",
        "Código do Cliente",
        "Código do Produto",
        "Data de Início",
        "Data de Término",
        "Descrição do Serviço",
        "Status",
    ]
    entries_ordens_servico = {}
    for i, label in enumerate(labels_ordens_servico):
        tk.Label(service_order_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(service_order_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_ordens_servico[label] = entry

    def save_ordem_servico():
        if validate_entries(entries_ordens_servico):
            data = (
                entries_ordens_servico["Código da Ordem de Serviço"].get(),
                entries_ordens_servico["Código do Cliente"].get(),
                entries_ordens_servico["Código do Produto"].get(),
                entries_ordens_servico["Data de Início"].get(),
                entries_ordens_servico["Data de Término"].get(),
                entries_ordens_servico["Descrição do Serviço"].get(),
                entries_ordens_servico["Status"].get(),
            )
            save_to_db('OrdensServico', data)
            messagebox.showinfo("Success", "Service order registered successfully!")
            update_service_order_list()
            service_order_window.destroy()

    def update_ordem_servico():
        if validate_entries(entries_ordens_servico):
            data = {
                "CodigoCliente": entries_ordens_servico["Código do Cliente"].get(),
                "CodigoProduto": entries_ordens_servico["Código do Produto"].get(),
                "DataInicio": entries_ordens_servico["Data de Início"].get(),
                "DataTermino": entries_ordens_servico["Data de Término"].get(),
                "DescricaoServico": entries_ordens_servico["Descrição do Serviço"].get(),
                "Status": entries_ordens_servico["Status"].get(),
            }
            update_db('OrdensServico', data, 'CodigoOrdemServico', entries_ordens_servico["Código da Ordem de Serviço"].get())
            messagebox.showinfo("Success", "Service order updated successfully!")
            update_service_order_list()
            service_order_window.destroy()

    def delete_ordem_servico():
        codigo_ordem_servico = entries_ordens_servico["Código da Ordem de Serviço"].get()
        if codigo_ordem_servico:
            delete_from_db('OrdensServico', 'CodigoOrdemServico', codigo_ordem_servico)
            messagebox.showinfo("Success", "Service order deleted successfully!")
            update_service_order_list()
            service_order_window.destroy()
        else:
            messagebox.showerror("Error", "Service order code is required to delete.")

    tk.Button(service_order_window, text="Save", command=save_ordem_servico).grid(row=len(labels_ordens_servico), column=0, sticky="w")
    tk.Button(service_order_window, text="Update", command=update_ordem_servico).grid(row=len(labels_ordens_servico), column=1, sticky="w")
    tk.Button(service_order_window, text="Delete", command=delete_ordem_servico).grid(row=len(labels_ordens_servico), column=2, sticky="w")

def update_service_order_list():
    for row in service_order_tree.get_children():
        service_order_tree.delete(row)
    cursor.execute("SELECT * FROM OrdensServico")
    for row in cursor.fetchall():
        service_order_tree.insert("", "end", values=row)

# Service order list setup
service_order_tree = ttk.Treeview(root, columns=("CodigoOrdemServico", "CodigoCliente", "CodigoProduto", "DataInicio", "DataTermino", "DescricaoServico", "Status"), show="headings")
service_order_tree.heading("CodigoOrdemServico", text="Código da Ordem de Serviço")
service_order_tree.heading("CodigoCliente", text="Código do Cliente")
service_order_tree.heading("CodigoProduto", text="Código do Produto")
service_order_tree.heading("DataInicio", text="Data de Início")
service_order_tree.heading("DataTermino", text="Data de Término")
service_order_tree.heading("DescricaoServico", text="Descrição do Serviço")
service_order_tree.heading("Status", text="Status")
service_order_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Service Order", command=create_service_order_window).pack()

update_service_order_list()

# ------------------- Employee Registration -------------------
def create_employee_window():
    employee_window = tk.Toplevel(root)
    employee_window.title("Employee Registration")

    labels_funcionarios = [
        "Código do Funcionário",
        "Nome do Funcionário",
        "CPF",
        "Endereço",
        "Telefone",
        "Email",
        "Cargo",
        "Salário",
    ]
    entries_funcionarios = {}
    for i, label in enumerate(labels_funcionarios):
        tk.Label(employee_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(employee_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_funcionarios[label] = entry

    def save_funcionario():
        if validate_entries(entries_funcionarios):
            data = (
                entries_funcionarios["Código do Funcionário"].get(),
                entries_funcionarios["Nome do Funcionário"].get(),
                entries_funcionarios["CPF"].get(),
                entries_funcionarios["Endereço"].get(),
                entries_funcionarios["Telefone"].get(),
                entries_funcionarios["Email"].get(),
                entries_funcionarios["Cargo"].get(),
                entries_funcionarios["Salário"].get(),
            )
            save_to_db('Funcionarios', data)
            messagebox.showinfo("Success", "Employee registered successfully!")
            update_employee_list()
            employee_window.destroy()

    def update_funcionario():
        if validate_entries(entries_funcionarios):
            data = {
                "NomeFuncionario": entries_funcionarios["Nome do Funcionário"].get(),
                "CPF": entries_funcionarios["CPF"].get(),
                "Endereco": entries_funcionarios["Endereço"].get(),
                "Telefone": entries_funcionarios["Telefone"].get(),
                "Email": entries_funcionarios["Email"].get(),
                "Cargo": entries_funcionarios["Cargo"].get(),
                "Salario": entries_funcionarios["Salário"].get(),
            }
            update_db('Funcionarios', data, 'CodigoFuncionario', entries_funcionarios["Código do Funcionário"].get())
            messagebox.showinfo("Success", "Employee updated successfully!")
            update_employee_list()
            employee_window.destroy()

    def delete_funcionario():
        codigo_funcionario = entries_funcionarios["Código do Funcionário"].get()
        if codigo_funcionario:
            delete_from_db('Funcionarios', 'CodigoFuncionario', codigo_funcionario)
            messagebox.showinfo("Success", "Employee deleted successfully!")
            update_employee_list()
            employee_window.destroy()
        else:
            messagebox.showerror("Error", "Employee code is required to delete.")

    tk.Button(employee_window, text="Save", command=save_funcionario).grid(row=len(labels_funcionarios), column=0, sticky="w")
    tk.Button(employee_window, text="Update", command=update_funcionario).grid(row=len(labels_funcionarios), column=1, sticky="w")
    tk.Button(employee_window, text="Delete", command=delete_funcionario).grid(row=len(labels_funcionarios), column=2, sticky="w")

def update_employee_list():
    for row in employee_tree.get_children():
        employee_tree.delete(row)
    cursor.execute("SELECT * FROM Funcionarios")
    for row in cursor.fetchall():
        employee_tree.insert("", "end", values=row)

# Employee list setup
employee_tree = ttk.Treeview(root, columns=("CodigoFuncionario", "NomeFuncionario", "CPF", "Endereco", "Telefone", "Email", "Cargo", "Salario"), show="headings")
employee_tree.heading("CodigoFuncionario", text="Código do Funcionário")
employee_tree.heading("NomeFuncionario", text="Nome do Funcionário")
employee_tree.heading("CPF", text="CPF")
employee_tree.heading("Endereco", text="Endereço")
employee_tree.heading("Telefone", text="Telefone")
employee_tree.heading("Email", text="Email")
employee_tree.heading("Cargo", text="Cargo")
employee_tree.heading("Salario", text="Salário")
employee_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Employee", command=create_employee_window).pack()

update_employee_list()

# ------------------- Supplier Registration -------------------
def create_supplier_window():
    supplier_window = tk.Toplevel(root)
    supplier_window.title("Supplier Registration")

    labels_fornecedores = [
        "Código do Fornecedor",
        "Nome do Fornecedor",
        "CNPJ",
        "Endereço",
        "Telefone",
        "Email",
    ]
    entries_fornecedores = {}
    for i, label in enumerate(labels_fornecedores):
        tk.Label(supplier_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(supplier_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_fornecedores[label] = entry

    def save_fornecedor():
        if validate_entries(entries_fornecedores):
            data = (
                entries_fornecedores["Código do Fornecedor"].get(),
                entries_fornecedores["Nome do Fornecedor"].get(),
                entries_fornecedores["CNPJ"].get(),
                entries_fornecedores["Endereço"].get(),
                entries_fornecedores["Telefone"].get(),
                entries_fornecedores["Email"].get(),
            )
            save_to_db('Fornecedores', data)
            messagebox.showinfo("Success", "Supplier registered successfully!")
            update_supplier_list()
            supplier_window.destroy()

    def update_fornecedor():
        if validate_entries(entries_fornecedores):
            data = {
                "NomeFornecedor": entries_fornecedores["Nome do Fornecedor"].get(),
                "CNPJ": entries_fornecedores["CNPJ"].get(),
                "Endereco": entries_fornecedores["Endereço"].get(),
                "Telefone": entries_fornecedores["Telefone"].get(),
                "Email": entries_fornecedores["Email"].get(),
            }
            update_db('Fornecedores', data, 'CodigoFornecedor', entries_fornecedores["Código do Fornecedor"].get())
            messagebox.showinfo("Success", "Supplier updated successfully!")
            update_supplier_list()
            supplier_window.destroy()

    def delete_fornecedor():
        codigo_fornecedor = entries_fornecedores["Código do Fornecedor"].get()
        if codigo_fornecedor:
            delete_from_db('Fornecedores', 'CodigoFornecedor', codigo_fornecedor)
            messagebox.showinfo("Success", "Supplier deleted successfully!")
            update_supplier_list()
            supplier_window.destroy()
        else:
            messagebox.showerror("Error", "Supplier code is required to delete.")

    tk.Button(supplier_window, text="Save", command=save_fornecedor).grid(row=len(labels_fornecedores), column=0, sticky="w")
    tk.Button(supplier_window, text="Update", command=update_fornecedor).grid(row=len(labels_fornecedores), column=1, sticky="w")
    tk.Button(supplier_window, text="Delete", command=delete_fornecedor).grid(row=len(labels_fornecedores), column=2, sticky="w")

def update_supplier_list():
    for row in supplier_tree.get_children():
        supplier_tree.delete(row)
    cursor.execute("SELECT * FROM Fornecedores")
    for row in cursor.fetchall():
        supplier_tree.insert("", "end", values=row)

# Supplier list setup
supplier_tree = ttk.Treeview(root, columns=("CodigoFornecedor", "NomeFornecedor", "CNPJ", "Endereco", "Telefone", "Email"), show="headings")
supplier_tree.heading("CodigoFornecedor", text="Código do Fornecedor")
supplier_tree.heading("NomeFornecedor", text="Nome do Fornecedor")
supplier_tree.heading("CNPJ", text="CNPJ")
supplier_tree.heading("Endereco", text="Endereço")
supplier_tree.heading("Telefone", text="Telefone")
supplier_tree.heading("Email", text="Email")
supplier_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Supplier", command=create_supplier_window).pack()

update_supplier_list()

# ------------------- Vehicle Registration -------------------
def create_vehicle_window():
    vehicle_window = tk.Toplevel(root)
    vehicle_window.title("Vehicle Registration")

    labels_veiculos = [
        "Código do Veículo",
        "Placa",
        "Modelo",
        "Ano",
        "Cor",
        "Código do Cliente",
    ]
    entries_veiculos = {}
    for i, label in enumerate(labels_veiculos):
        tk.Label(vehicle_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(vehicle_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_veiculos[label] = entry

    def save_veiculo():
        if validate_entries(entries_veiculos):
            data = (
                entries_veiculos["Código do Veículo"].get(),
                entries_veiculos["Placa"].get(),
                entries_veiculos["Modelo"].get(),
                entries_veiculos["Ano"].get(),
                entries_veiculos["Cor"].get(),
                entries_veiculos["Código do Cliente"].get(),
            )
            save_to_db('Veiculos', data)
            messagebox.showinfo("Success", "Vehicle registered successfully!")
            update_vehicle_list()
            vehicle_window.destroy()

    def update_veiculo():
        if validate_entries(entries_veiculos):
            data = {
                "Placa": entries_veiculos["Placa"].get(),
                "Modelo": entries_veiculos["Modelo"].get(),
                "Ano": entries_veiculos["Ano"].get(),
                "Cor": entries_veiculos["Cor"].get(),
                "CodigoCliente": entries_veiculos["Código do Cliente"].get(),
            }
            update_db('Veiculos', data, 'CodigoVeiculo', entries_veiculos["Código do Veículo"].get())
            messagebox.showinfo("Success", "Vehicle updated successfully!")
            update_vehicle_list()
            vehicle_window.destroy()

    def delete_veiculo():
        codigo_veiculo = entries_veiculos["Código do Veículo"].get()
        if codigo_veiculo:
            delete_from_db('Veiculos', 'CodigoVeiculo', codigo_veiculo)
            messagebox.showinfo("Success", "Vehicle deleted successfully!")
            update_vehicle_list()
            vehicle_window.destroy()
        else:
            messagebox.showerror("Error", "Vehicle code is required to delete.")

    tk.Button(vehicle_window, text="Save", command=save_veiculo).grid(row=len(labels_veiculos), column=0, sticky="w")
    tk.Button(vehicle_window, text="Update", command=update_veiculo).grid(row=len(labels_veiculos), column=1, sticky="w")
    tk.Button(vehicle_window, text="Delete", command=delete_veiculo).grid(row=len(labels_veiculos), column=2, sticky="w")

def update_vehicle_list():
    for row in vehicle_tree.get_children():
        vehicle_tree.delete(row)
    cursor.execute("SELECT * FROM Veiculos")
    for row in cursor.fetchall():
        vehicle_tree.insert("", "end", values=row)

# Vehicle list setup
vehicle_tree = ttk.Treeview(root, columns=("CodigoVeiculo", "Placa", "Modelo", "Ano", "Cor", "CodigoCliente"), show="headings")
vehicle_tree.heading("CodigoVeiculo", text="Código do Veículo")
vehicle_tree.heading("Placa", text="Placa")
vehicle_tree.heading("Modelo", text="Modelo")
vehicle_tree.heading("Ano", text="Ano")
vehicle_tree.heading("Cor", text="Cor")
vehicle_tree.heading("CodigoCliente", text="Código do Cliente")
vehicle_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Vehicle", command=create_vehicle_window).pack()

update_vehicle_list()

# ------------------- Service Registration -------------------
def create_service_window():
    service_window = tk.Toplevel(root)
    service_window.title("Service Registration")

    labels_servicos = [
        "Código do Serviço",
        "Nome do Serviço",
        "Descrição",
        "Preço",
    ]
    entries_servicos = {}
    for i, label in enumerate(labels_servicos):
        tk.Label(service_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(service_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_servicos[label] = entry

    def save_servico():
        if validate_entries(entries_servicos):
            data = (
                entries_servicos["Código do Serviço"].get(),
                entries_servicos["Nome do Serviço"].get(),
                entries_servicos["Descrição"].get(),
                entries_servicos["Preço"].get(),
            )
            save_to_db('Servicos', data)
            messagebox.showinfo("Success", "Service registered successfully!")
            update_service_list()
            service_window.destroy()

    def update_servico():
        if validate_entries(entries_servicos):
            data = {
                "NomeServico": entries_servicos["Nome do Serviço"].get(),
                "Descricao": entries_servicos["Descrição"].get(),
                "Preco": entries_servicos["Preço"].get(),
            }
            update_db('Servicos', data, 'CodigoServico', entries_servicos["Código do Serviço"].get())
            messagebox.showinfo("Success", "Service updated successfully!")
            update_service_list()
            service_window.destroy()

    def delete_servico():
        codigo_servico = entries_servicos["Código do Serviço"].get()
        if codigo_servico:
            delete_from_db('Servicos', 'CodigoServico', codigo_servico)
            messagebox.showinfo("Success", "Service deleted successfully!")
            update_service_list()
            service_window.destroy()
        else:
            messagebox.showerror("Error", "Service code is required to delete.")

    tk.Button(service_window, text="Save", command=save_servico).grid(row=len(labels_servicos), column=0, sticky="w")
    tk.Button(service_window, text="Update", command=update_servico).grid(row=len(labels_servicos), column=1, sticky="w")
    tk.Button(service_window, text="Delete", command=delete_servico).grid(row=len(labels_servicos), column=2, sticky="w")

def update_service_list():
    for row in service_tree.get_children():
        service_tree.delete(row)
    cursor.execute("SELECT * FROM Servicos")
    for row in cursor.fetchall():
        service_tree.insert("", "end", values=row)

# Service list setup
service_tree = ttk.Treeview(root, columns=("CodigoServico", "NomeServico", "Descricao", "Preco"), show="headings")
service_tree.heading("CodigoServico", text="Código do Serviço")
service_tree.heading("NomeServico", text="Nome do Serviço")
service_tree.heading("Descricao", text="Descrição")
service_tree.heading("Preco", text="Preço")
service_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Service", command=create_service_window).pack()

update_service_list()

# ------------------- Product Category Registration -------------------
def create_category_window():
    category_window = tk.Toplevel(root)
    category_window.title("Product Category Registration")

    labels_categorias = [
        "Código da Categoria",
        "Nome da Categoria",
        "Descrição",
    ]
    entries_categorias = {}
    for i, label in enumerate(labels_categorias):
        tk.Label(category_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(category_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_categorias[label] = entry

    def save_categoria():
        if validate_entries(entries_categorias):
            data = (
                entries_categorias["Código da Categoria"].get(),
                entries_categorias["Nome da Categoria"].get(),
                entries_categorias["Descrição"].get(),
            )
            save_to_db('Categorias', data)
            messagebox.showinfo("Success", "Category registered successfully!")
            update_category_list()
            category_window.destroy()

    def update_categoria():
        if validate_entries(entries_categorias):
            data = {
                "NomeCategoria": entries_categorias["Nome da Categoria"].get(),
                "Descricao": entries_categorias["Descrição"].get(),
            }
            update_db('Categorias', data, 'CodigoCategoria', entries_categorias["Código da Categoria"].get())
            messagebox.showinfo("Success", "Category updated successfully!")
            update_category_list()
            category_window.destroy()

    def delete_categoria():
        codigo_categoria = entries_categorias["Código da Categoria"].get()
        if codigo_categoria:
            delete_from_db('Categorias', 'CodigoCategoria', codigo_categoria)
            messagebox.showinfo("Success", "Category deleted successfully!")
            update_category_list()
            category_window.destroy()
        else:
            messagebox.showerror("Error", "Category code is required to delete.")

    tk.Button(category_window, text="Save", command=save_categoria).grid(row=len(labels_categorias), column=0, sticky="w")
    tk.Button(category_window, text="Update", command=update_categoria).grid(row=len(labels_categorias), column=1, sticky="w")
    tk.Button(category_window, text="Delete", command=delete_categoria).grid(row=len(labels_categorias), column=2, sticky="w")

def update_category_list():
    for row in category_tree.get_children():
        category_tree.delete(row)
    cursor.execute("SELECT * FROM Categorias")
    for row in cursor.fetchall():
        category_tree.insert("", "end", values=row)

# Category list setup
category_tree = ttk.Treeview(root, columns=("CodigoCategoria", "NomeCategoria", "Descricao"), show="headings")
category_tree.heading("CodigoCategoria", text="Código da Categoria")
category_tree.heading("NomeCategoria", text="Nome da Categoria")
category_tree.heading("Descricao", text="Descrição")
category_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Category", command=create_category_window).pack()

update_category_list()

# ------------------- Payment Type Registration -------------------
def create_payment_type_window():
    payment_type_window = tk.Toplevel(root)
    payment_type_window.title("Payment Type Registration")

    labels_tipos_pagamento = [
        "Código do Tipo de Pagamento",
        "Nome do Tipo de Pagamento",
        "Descrição",
    ]
    entries_tipos_pagamento = {}
    for i, label in enumerate(labels_tipos_pagamento):
        tk.Label(payment_type_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(payment_type_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_tipos_pagamento[label] = entry

    def save_tipo_pagamento():
        if validate_entries(entries_tipos_pagamento):
            data = (
                entries_tipos_pagamento["Código do Tipo de Pagamento"].get(),
                entries_tipos_pagamento["Nome do Tipo de Pagamento"].get(),
                entries_tipos_pagamento["Descrição"].get(),
            )
            save_to_db('TiposPagamento', data)
            messagebox.showinfo("Success", "Payment type registered successfully!")
            update_payment_type_list()
            payment_type_window.destroy()

    def update_tipo_pagamento():
        if validate_entries(entries_tipos_pagamento):
            data = {
                "NomeTipoPagamento": entries_tipos_pagamento["Nome do Tipo de Pagamento"].get(),
                "Descricao": entries_tipos_pagamento["Descrição"].get(),
            }
            update_db('TiposPagamento', data, 'CodigoTipoPagamento', entries_tipos_pagamento["Código do Tipo de Pagamento"].get())
            messagebox.showinfo("Success", "Payment type updated successfully!")
            update_payment_type_list()
            payment_type_window.destroy()

    def delete_tipo_pagamento():
        codigo_tipo_pagamento = entries_tipos_pagamento["Código do Tipo de Pagamento"].get()
        if codigo_tipo_pagamento:
            delete_from_db('TiposPagamento', 'CodigoTipoPagamento', codigo_tipo_pagamento)
            messagebox.showinfo("Success", "Payment type deleted successfully!")
            update_payment_type_list()
            payment_type_window.destroy()
        else:
            messagebox.showerror("Error", "Payment type code is required to delete.")

    tk.Button(payment_type_window, text="Save", command=save_tipo_pagamento).grid(row=len(labels_tipos_pagamento), column=0, sticky="w")
    tk.Button(payment_type_window, text="Update", command=update_tipo_pagamento).grid(row=len(labels_tipos_pagamento), column=1, sticky="w")
    tk.Button(payment_type_window, text="Delete", command=delete_tipo_pagamento).grid(row=len(labels_tipos_pagamento), column=2, sticky="w")

def update_payment_type_list():
    for row in payment_type_tree.get_children():
        payment_type_tree.delete(row)
    cursor.execute("SELECT * FROM TiposPagamento")
    for row in cursor.fetchall():
        payment_type_tree.insert("", "end", values=row)

# Payment type list setup
payment_type_tree = ttk.Treeview(root, columns=("CodigoTipoPagamento", "NomeTipoPagamento", "Descricao"), show="headings")
payment_type_tree.heading("CodigoTipoPagamento", text="Código do Tipo de Pagamento")
payment_type_tree.heading("NomeTipoPagamento", text="Nome do Tipo de Pagamento")
payment_type_tree.heading("Descricao", text="Descrição")
payment_type_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Payment Type", command=create_payment_type_window).pack()

update_payment_type_list()

# ------------------- Payment Order Registration -------------------
def create_payment_order_window():
    payment_order_window = tk.Toplevel(root)
    payment_order_window.title("Payment Order Registration")

    labels_ordens_pagamento = [
        "Código da Ordem de Pagamento",
        "Código do Cliente",
        "Código do Tipo de Pagamento",
        "Data de Pagamento",
        "Valor",
        "Status",
    ]
    entries_ordens_pagamento = {}
    for i, label in enumerate(labels_ordens_pagamento):
        tk.Label(payment_order_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(payment_order_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_ordens_pagamento[label] = entry

    def save_ordem_pagamento():
        if validate_entries(entries_ordens_pagamento):
            data = (
                entries_ordens_pagamento["Código da Ordem de Pagamento"].get(),
                entries_ordens_pagamento["Código do Cliente"].get(),
                entries_ordens_pagamento["Código do Tipo de Pagamento"].get(),
                entries_ordens_pagamento["Data de Pagamento"].get(),
                entries_ordens_pagamento["Valor"].get(),
                entries_ordens_pagamento["Status"].get(),
            )
            save_to_db('OrdensPagamento', data)
            messagebox.showinfo("Success", "Payment order registered successfully!")
            update_payment_order_list()
            payment_order_window.destroy()

    def update_ordem_pagamento():
        if validate_entries(entries_ordens_pagamento):
            data = {
                "CodigoCliente": entries_ordens_pagamento["Código do Cliente"].get(),
                "CodigoTipoPagamento": entries_ordens_pagamento["Código do Tipo de Pagamento"].get(),
                "DataPagamento": entries_ordens_pagamento["Data de Pagamento"].get(),
                "Valor": entries_ordens_pagamento["Valor"].get(),
                "Status": entries_ordens_pagamento["Status"].get(),
            }
            update_db('OrdensPagamento', data, 'CodigoOrdemPagamento', entries_ordens_pagamento["Código da Ordem de Pagamento"].get())
            messagebox.showinfo("Success", "Payment order updated successfully!")
            update_payment_order_list()
            payment_order_window.destroy()

    def delete_ordem_pagamento():
        codigo_ordem_pagamento = entries_ordens_pagamento["Código da Ordem de Pagamento"].get()
        if codigo_ordem_pagamento:
            delete_from_db('OrdensPagamento', 'CodigoOrdemPagamento', codigo_ordem_pagamento)
            messagebox.showinfo("Success", "Payment order deleted successfully!")
            update_payment_order_list()
            payment_order_window.destroy()
        else:
            messagebox.showerror("Error", "Payment order code is required to delete.")

    tk.Button(payment_order_window, text="Save", command=save_ordem_pagamento).grid(row=len(labels_ordens_pagamento), column=0, sticky="w")
    tk.Button(payment_order_window, text="Update", command=update_ordem_pagamento).grid(row=len(labels_ordens_pagamento), column=1, sticky="w")
    tk.Button(payment_order_window, text="Delete", command=delete_ordem_pagamento).grid(row=len(labels_ordens_pagamento), column=2, sticky="w")

def update_payment_order_list():
    for row in payment_order_tree.get_children():
        payment_order_tree.delete(row)
    cursor.execute("SELECT * FROM OrdensPagamento")
    for row in cursor.fetchall():
        payment_order_tree.insert("", "end", values=row)

# Payment order list setup
payment_order_tree = ttk.Treeview(root, columns=("CodigoOrdemPagamento", "CodigoCliente", "CodigoTipoPagamento", "DataPagamento", "Valor", "Status"), show="headings")
payment_order_tree.heading("CodigoOrdemPagamento", text="Código da Ordem de Pagamento")
payment_order_tree.heading("CodigoCliente", text="Código do Cliente")
payment_order_tree.heading("CodigoTipoPagamento", text="Código do Tipo de Pagamento")
payment_order_tree.heading("DataPagamento", text="Data de Pagamento")
payment_order_tree.heading("Valor", text="Valor")
payment_order_tree.heading("Status", text="Status")
payment_order_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Payment Order", command=create_payment_order_window).pack()

update_payment_order_list()

# ------------------- Appointment Registration -------------------
def create_appointment_window():
    appointment_window = tk.Toplevel(root)
    appointment_window.title("Appointment Registration")

    labels_agendamentos = [
        "Código do Agendamento",
        "Código do Cliente",
        "Código do Veículo",
        "Data do Agendamento",
        "Hora do Agendamento",
        "Descrição",
    ]
    entries_agendamentos = {}
    for i, label in enumerate(labels_agendamentos):
        tk.Label(appointment_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(appointment_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_agendamentos[label] = entry

    def save_agendamento():
        if validate_entries(entries_agendamentos):
            data = (
                entries_agendamentos["Código do Agendamento"].get(),
                entries_agendamentos["Código do Cliente"].get(),
                entries_agendamentos["Código do Veículo"].get(),
                entries_agendamentos["Data do Agendamento"].get(),
                entries_agendamentos["Hora do Agendamento"].get(),
                entries_agendamentos["Descrição"].get(),
            )
            save_to_db('Agendamentos', data)
            messagebox.showinfo("Success", "Appointment registered successfully!")
            update_appointment_list()
            appointment_window.destroy()

    def update_agendamento():
        if validate_entries(entries_agendamentos):
            data = {
                "CodigoCliente": entries_agendamentos["Código do Cliente"].get(),
                "CodigoVeiculo": entries_agendamentos["Código do Veículo"].get(),
                "DataAgendamento": entries_agendamentos["Data do Agendamento"].get(),
                "HoraAgendamento": entries_agendamentos["Hora do Agendamento"].get(),
                "Descricao": entries_agendamentos["Descrição"].get(),
            }
            update_db('Agendamentos', data, 'CodigoAgendamento', entries_agendamentos["Código do Agendamento"].get())
            messagebox.showinfo("Success", "Appointment updated successfully!")
            update_appointment_list()
            appointment_window.destroy()

    def delete_agendamento():
        codigo_agendamento = entries_agendamentos["Código do Agendamento"].get()
        if codigo_agendamento:
            delete_from_db('Agendamentos', 'CodigoAgendamento', codigo_agendamento)
            messagebox.showinfo("Success", "Appointment deleted successfully!")
            update_appointment_list()
            appointment_window.destroy()
        else:
            messagebox.showerror("Error", "Appointment code is required to delete.")

    tk.Button(appointment_window, text="Save", command=save_agendamento).grid(row=len(labels_agendamentos), column=0, sticky="w")
    tk.Button(appointment_window, text="Update", command=update_agendamento).grid(row=len(labels_agendamentos), column=1, sticky="w")
    tk.Button(appointment_window, text="Delete", command=delete_agendamento).grid(row=len(labels_agendamentos), column=2, sticky="w")

def update_appointment_list():
    for row in appointment_tree.get_children():
        appointment_tree.delete(row)
    cursor.execute("SELECT * FROM Agendamentos")
    for row in cursor.fetchall():
        appointment_tree.insert("", "end", values=row)

# Appointment list setup
appointment_tree = ttk.Treeview(root, columns=("CodigoAgendamento", "CodigoCliente", "CodigoVeiculo", "DataAgendamento", "HoraAgendamento", "Descricao"), show="headings")
appointment_tree.heading("CodigoAgendamento", text="Código do Agendamento")
appointment_tree.heading("CodigoCliente", text="Código do Cliente")
appointment_tree.heading("CodigoVeiculo", text="Código do Veículo")
appointment_tree.heading("DataAgendamento", text="Data do Agendamento")
appointment_tree.heading("HoraAgendamento", text="Hora do Agendamento")
appointment_tree.heading("Descricao", text="Descrição")
appointment_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Appointment", command=create_appointment_window).pack()

update_appointment_list()

# ------------------- Inventory Registration -------------------
def create_inventory_window():
    inventory_window = tk.Toplevel(root)
    inventory_window.title("Inventory Registration")

    labels_inventario = [
        "Código do Item",
        "Nome do Item",
        "Quantidade",
        "Localização",
        "Data de Entrada",
        "Data de Saída",
    ]
    entries_inventario = {}
    for i, label in enumerate(labels_inventario):
        tk.Label(inventory_window, text=label, anchor="w").grid(row=i, column=0, sticky="w")
        entry = tk.Entry(inventory_window)
        entry.grid(row=i, column=1, sticky="w")
        entries_inventario[label] = entry

    def save_inventario():
        if validate_entries(entries_inventario):
            data = (
                entries_inventario["Código do Item"].get(),
                entries_inventario["Nome do Item"].get(),
                entries_inventario["Quantidade"].get(),
                entries_inventario["Localização"].get(),
                entries_inventario["Data de Entrada"].get(),
                entries_inventario["Data de Saída"].get(),
            )
            save_to_db('Inventario', data)
            messagebox.showinfo("Success", "Inventory item registered successfully!")
            update_inventory_list()
            inventory_window.destroy()

    def update_inventario():
        if validate_entries(entries_inventario):
            data = {
                "NomeItem": entries_inventario["Nome do Item"].get(),
                "Quantidade": entries_inventario["Quantidade"].get(),
                "Localizacao": entries_inventario["Localização"].get(),
                "DataEntrada": entries_inventario["Data de Entrada"].get(),
                "DataSaida": entries_inventario["Data de Saída"].get(),
            }
            update_db('Inventario', data, 'CodigoItem', entries_inventario["Código do Item"].get())
            messagebox.showinfo("Success", "Inventory item updated successfully!")
            update_inventory_list()
            inventory_window.destroy()

    def delete_inventario():
        codigo_item = entries_inventario["Código do Item"].get()
        if codigo_item:
            delete_from_db('Inventario', 'CodigoItem', codigo_item)
            messagebox.showinfo("Success", "Inventory item deleted successfully!")
            update_inventory_list()
            inventory_window.destroy()
        else:
            messagebox.showerror("Error", "Inventory item code is required to delete.")

    tk.Button(inventory_window, text="Save", command=save_inventario).grid(row=len(labels_inventario), column=0, sticky="w")
    tk.Button(inventory_window, text="Update", command=update_inventario).grid(row=len(labels_inventario), column=1, sticky="w")
    tk.Button(inventory_window, text="Delete", command=delete_inventario).grid(row=len(labels_inventario), column=2, sticky="w")

def update_inventory_list():
    for row in inventory_tree.get_children():
        inventory_tree.delete(row)
    cursor.execute("SELECT * FROM Inventario")
    for row in cursor.fetchall():
        inventory_tree.insert("", "end", values=row)

# Inventory list setup
inventory_tree = ttk.Treeview(root, columns=("CodigoItem", "NomeItem", "Quantidade", "Localizacao", "DataEntrada", "DataSaida"), show="headings")
inventory_tree.heading("CodigoItem", text="Código do Item")
inventory_tree.heading("NomeItem", text="Nome do Item")
inventory_tree.heading("Quantidade", text="Quantidade")
inventory_tree.heading("Localizacao", text="Localização")
inventory_tree.heading("DataEntrada", text="Data de Entrada")
inventory_tree.heading("DataSaida", text="Data de Saída")
inventory_tree.pack(fill=tk.BOTH, expand=True)

tk.Button(root, text="Register Inventory Item", command=create_inventory_window).pack()

update_inventory_list()

import webbrowser

# Function to open the Sebrae NFe portal
def open_sebrae_nfe_portal():
    url = "https://amei.sebrae.com.br/auth/realms/externo/protocol/openid-connect/auth?client_id=emissor-nfe-frontend&redirect_uri=https%3A%2F%2Femissornfe.sebrae.com.br%2F&state=deda205a-1c67-4013-8e32-aaa0e1580118&response_mode=fragment&response_type=code&scope=openid&nonce=72cbb723-5708-46e5-8d00-df5ae4ad90d7"
    webbrowser.open(url)

# Add a button to the main window to open the Sebrae NFe portal
tk.Button(root, text="Emitir Nota Fiscal", command=open_sebrae_nfe_portal).pack()

root.mainloop()
