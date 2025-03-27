import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import openpyxl
import os
import sys
from datetime import datetime

class Produto:
    def __init__(self, id, nome, quantidade, preco):
        self.id = id
        self.nome = nome
        self.quantidade = quantidade
        self.preco = preco

    def adicionar_estoque(self, quantidade):
        self.quantidade += quantidade

    def remover_estoque(self, quantidade):
        if quantidade > self.quantidade:
            messagebox.showinfo("Erro", "Estoque insuficiente para a venda.")
            return False
        self.quantidade -= quantidade
        return True

    def __str__(self):
        return f"ID: {self.id} | Produto: {self.nome} | Quantidade: {self.quantidade} | Preço: R${self.preco:.2f}"


class GerenciadorEstoque:
    def __init__(self):
        self.estoque_file = 'estoque.xlsx'
        self.produtos = self.carregar_estoque()
        self.proximo_id = self.calcular_proximo_id()

    def calcular_proximo_id(self):
        if not self.produtos:
            return 1
        return max(int(p.id) for p in self.produtos.values()) + 1

    def carregar_estoque(self):
        try:
            if os.path.exists(self.estoque_file):
                df = pd.read_excel(self.estoque_file, engine='openpyxl')
                produtos = {}
                for _, row in df.iterrows():
                    produtos[row['Id']] = Produto(row['Id'], row['Nome'], row['Quantidade'], row['Preço'])
                return produtos
            return {}
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar estoque: {str(e)}")
            return {}

    def salvar_estoque(self):
        if not self.produtos:
            return
            
        data = {
            'Id': [produto.id for produto in self.produtos.values()],
            'Nome': [produto.nome for produto in self.produtos.values()],
            'Quantidade': [produto.quantidade for produto in self.produtos.values()],
            'Preço': [produto.preco for produto in self.produtos.values()]
        }
        df = pd.DataFrame(data)
        
        try:
            df.to_excel(self.estoque_file, index=False, engine='openpyxl')
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar estoque: {str(e)}")

    def gerar_novo_id(self):
        novo_id = self.proximo_id
        self.proximo_id += 1
        return novo_id

    def cadastrar_produto(self, nome, quantidade, preco):
        id = str(self.gerar_novo_id())
        if any(p.nome.lower() == nome.lower() for p in self.produtos.values()):
            messagebox.showinfo("Erro", "Produto já cadastrado.")
            return False
            
        try:
            self.produtos[id] = Produto(id, nome, int(quantidade), float(preco))
            self.salvar_estoque()
            messagebox.showinfo("Sucesso", f"Produto cadastrado com sucesso. ID: {id}")
            return id
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser inteiro e preço deve ser número decimal.")
            return None

    def buscar_por_id(self, id):
        return self.produtos.get(id)

    def buscar_por_nome(self, nome):
        for produto in self.produtos.values():
            if produto.nome.lower() == nome.lower():
                return produto
        return None

    def atualizar_estoque_venda(self, id, quantidade):
        produto = self.buscar_por_id(id)
        if produto:
            if produto.remover_estoque(int(quantidade)):
                self.salvar_estoque()
                messagebox.showinfo("Sucesso", f"Venda realizada. Estoque atual: {produto.quantidade}")
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def adicionar_estoque(self, id, quantidade):
        produto = self.buscar_por_id(id)
        if produto:
            try:
                produto.adicionar_estoque(int(quantidade))
                self.salvar_estoque()
                messagebox.showinfo("Sucesso", f"Estoque atualizado: {produto.quantidade}")
            except ValueError:
                messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def listar_estoque(self):
        if not self.produtos:
            messagebox.showinfo("Estoque", "Nenhum produto cadastrado.")
            return
            
        estoque = "Estoque atual:\n\n" + "\n".join(str(produto) for produto in sorted(
            self.produtos.values(), key=lambda x: int(x.id)))
        messagebox.showinfo("Estoque", estoque)

    def exportar_estoque_excel(self):
        if not self.produtos:
            messagebox.showinfo("Aviso", "Nenhum produto para exportar.")
            return
            
        data = {
            'Id': [produto.id for produto in self.produtos.values()],
            'Nome': [produto.nome for produto in self.produtos.values()],
            'Quantidade': [produto.quantidade for produto in self.produtos.values()],
            'Preço': [produto.preco for produto in self.produtos.values()]
        }
        df = pd.DataFrame(data)

        try:
            df.to_excel('estoque_exportado.xlsx', index=False, engine='openpyxl')
            messagebox.showinfo("Sucesso", "Estoque exportado para 'estoque_exportado.xlsx'")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar: {str(e)}")

    def abrir_excel(self):
        file_path = 'estoque_exportado.xlsx'
        if not os.path.exists(file_path):
            messagebox.showerror("Erro", "Arquivo não encontrado. Exporte o estoque primeiro.")
            return
            
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            else:
                os.system(f'open "{file_path}"')  # macOS
                # ou os.system(f'xdg-open "{file_path}"') para Linux
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")


class InterfaceEstoque:
    def __init__(self, root):
        self.gerenciador = GerenciadorEstoque()
        self.root = root
        self.root.title("Gerenciador de Estoque")
        self.root.geometry("500x600")
        
        # Configuração de grid
        for i in range(12):
            self.root.grid_rowconfigure(i, weight=1)
        for i in range(2):
            self.root.grid_columnconfigure(i, weight=1)

        # Frame de busca
        tk.Label(root, text="Buscar Produto:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.busca_var = tk.StringVar()
        self.radio_busca_id = tk.Radiobutton(root, text="Por ID", variable=self.busca_var, value="id")
        self.radio_busca_id.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.radio_busca_nome = tk.Radiobutton(root, text="Por Nome", variable=self.busca_var, value="nome")
        self.radio_busca_nome.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.busca_var.set("id")
        
        self.entry_busca = tk.Entry(root)
        self.entry_busca.grid(row=2, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        self.btn_buscar = tk.Button(root, text="Buscar", command=self.buscar_produto)
        self.btn_buscar.grid(row=3, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        # Separador
        ttk.Separator(root, orient='horizontal').grid(row=4, column=0, columnspan=2, sticky="we", pady=10)

        # Frame de cadastro
        tk.Label(root, text="ID do Produto:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.entry_id = tk.Entry(root, state='readonly')
        self.entry_id.grid(row=5, column=1, sticky="we", padx=5, pady=5)

        tk.Label(root, text="Nome do Produto:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.entry_nome = tk.Entry(root)
        self.entry_nome.grid(row=6, column=1, sticky="we", padx=5, pady=5)

        tk.Label(root, text="Quantidade:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        self.entry_quantidade = tk.Entry(root)
        self.entry_quantidade.grid(row=7, column=1, sticky="we", padx=5, pady=5)

        tk.Label(root, text="Preço:").grid(row=8, column=0, sticky="w", padx=5, pady=5)
        self.entry_preco = tk.Entry(root)
        self.entry_preco.grid(row=8, column=1, sticky="we", padx=5, pady=5)

        # Botões
        buttons = [
            ("Cadastrar Produto", self.cadastrar_produto),
            ("Atualizar Estoque (Venda)", self.atualizar_estoque_venda),
            ("Adicionar ao Estoque", self.adicionar_estoque),
            ("Listar Estoque", self.listar_estoque),
            ("Exportar para Excel", self.exportar_excel),
            ("Abrir Excel", self.abrir_excel)
        ]

        for i, (text, command) in enumerate(buttons, start=9):
            tk.Button(root, text=text, command=command).grid(
                row=i, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        # Atualiza o ID automaticamente
        self.atualizar_id()

    def atualizar_id(self):
        novo_id = self.gerenciador.gerar_novo_id()
        self.entry_id.config(state='normal')
        self.entry_id.delete(0, tk.END)
        self.entry_id.insert(0, str(novo_id))
        self.entry_id.config(state='readonly')

    def buscar_produto(self):
        termo = self.entry_busca.get().strip()
        if not termo:
            messagebox.showerror("Erro", "Digite um termo para busca.")
            return

        if self.busca_var.get() == "id":
            produto = self.gerenciador.buscar_por_id(termo)
        else:
            produto = self.gerenciador.buscar_por_nome(termo)

        if produto:
            self.entry_id.config(state='normal')
            self.entry_id.delete(0, tk.END)
            self.entry_id.insert(0, produto.id)
            self.entry_id.config(state='readonly')
            
            self.entry_nome.delete(0, tk.END)
            self.entry_nome.insert(0, produto.nome)
            
            self.entry_quantidade.delete(0, tk.END)
            self.entry_quantidade.insert(0, str(produto.quantidade))
            
            self.entry_preco.delete(0, tk.END)
            self.entry_preco.insert(0, f"{produto.preco:.2f}")
            
            messagebox.showinfo("Sucesso", f"Produto encontrado: {produto.nome}")
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def cadastrar_produto(self):
        nome = self.entry_nome.get().strip()
        quantidade = self.entry_quantidade.get().strip()
        preco = self.entry_preco.get().strip()

        if not all([nome, quantidade, preco]):
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        id_cadastrado = self.gerenciador.cadastrar_produto(nome, quantidade, preco)
        if id_cadastrado:
            # Limpa os campos e prepara para novo cadastro
            self.entry_nome.delete(0, tk.END)
            self.entry_quantidade.delete(0, tk.END)
            self.entry_preco.delete(0, tk.END)
            self.atualizar_id()

    def atualizar_estoque_venda(self):
        id = self.entry_id.get()
        quantidade = self.entry_quantidade.get().strip()
        
        if not quantidade:
            messagebox.showerror("Erro", "Quantidade é obrigatória.")
            return
            
        self.gerenciador.atualizar_estoque_venda(id, quantidade)

    def adicionar_estoque(self):
        id = self.entry_id.get()
        quantidade = self.entry_quantidade.get().strip()
        
        if not quantidade:
            messagebox.showerror("Erro", "Quantidade é obrigatória.")
            return
            
        self.gerenciador.adicionar_estoque(id, quantidade)

    def listar_estoque(self):
        self.gerenciador.listar_estoque()

    def exportar_excel(self):
        self.gerenciador.exportar_estoque_excel()

    def abrir_excel(self):
        self.gerenciador.abrir_excel()


if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceEstoque(root)
    root.mainloop()