import tkinter as tk
from tkinter import messagebox

class Produto:
    def __init__(self, nome, quantidade, preco):
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
        return f"Produto: {self.nome}, Quantidade: {self.quantidade}, Preço: R${self.preco:.2f}"


class GerenciadorEstoque:
    def __init__(self):
        self.produtos = {}

    def cadastrar_produto(self, nome, quantidade, preco):
        if nome in self.produtos:
            messagebox.showinfo("Erro", "Produto já cadastrado.")
            return
        self.produtos[nome] = Produto(nome, quantidade, preco)
        messagebox.showinfo("Sucesso", "Produto cadastrado com sucesso.")

    def checar_estoque(self, nome):
        produto = self.produtos.get(nome)
        if produto:
            messagebox.showinfo("Estoque", str(produto))
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def atualizar_estoque_venda(self, nome, quantidade):
        produto = self.produtos.get(nome)
        if produto:
            if produto.remover_estoque(quantidade):
                messagebox.showinfo("Sucesso", f"Venda realizada. Estoque atualizado: {produto.quantidade} unidades restantes.")
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def adicionar_estoque(self, nome, quantidade):
        produto = self.produtos.get(nome)
        if produto:
            produto.adicionar_estoque(quantidade)
            messagebox.showinfo("Sucesso", f"Estoque atualizado. Novo estoque: {produto.quantidade} unidades.")
        else:
            messagebox.showinfo("Erro", "Produto não encontrado.")

    def listar_estoque(self):
        if not self.produtos:
            messagebox.showinfo("Estoque", "Nenhum produto cadastrado.")
            return
        estoque = "Estoque atual:\n" + "\n".join(str(produto) for produto in self.produtos.values())
        messagebox.showinfo("Estoque", estoque)


class InterfaceEstoque:
    def __init__(self, root):
        self.gerenciador = GerenciadorEstoque()
        self.root = root
        self.root.title("Gerenciador de Estoque")

        self.label_nome = tk.Label(root, text="Nome do Produto:")
        self.label_nome.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nome = tk.Entry(root)
        self.entry_nome.grid(row=0, column=1, sticky="we", padx=5, pady=5)

        self.label_quantidade = tk.Label(root, text="Quantidade:")
        self.label_quantidade.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_quantidade = tk.Entry(root)
        self.entry_quantidade.grid(row=1, column=1, sticky="we", padx=5, pady=5)

        self.label_preco = tk.Label(root, text="Preço:")
        self.label_preco.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_preco = tk.Entry(root)
        self.entry_preco.grid(row=2, column=1, sticky="we", padx=5, pady=5)

        self.btn_cadastrar = tk.Button(root, text="Cadastrar Produto", command=self.cadastrar_produto)
        self.btn_cadastrar.grid(row=3, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        self.btn_checar = tk.Button(root, text="Checar Estoque", command=self.checar_estoque)
        self.btn_checar.grid(row=4, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        self.btn_vender = tk.Button(root, text="Atualizar Estoque para Venda", command=self.atualizar_estoque_venda)
        self.btn_vender.grid(row=5, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        self.btn_adicionar = tk.Button(root, text="Adicionar ao Estoque", command=self.adicionar_estoque)
        self.btn_adicionar.grid(row=6, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        self.btn_listar = tk.Button(root, text="Listar Estoque", command=self.listar_estoque)
        self.btn_listar.grid(row=7, column=0, columnspan=2, sticky="we", padx=5, pady=5)

        # Configuração de redimensionamento
        self.root.grid_columnconfigure(0, weight=1)  # Coluna 0 redimensionável
        self.root.grid_columnconfigure(1, weight=2)  # Coluna 1 mais redimensionável

        for i in range(8):  # Configura todas as linhas para redimensionar
            self.root.grid_rowconfigure(i, weight=1)

    def cadastrar_produto(self):
        nome = self.entry_nome.get()
        quantidade = self.entry_quantidade.get()
        preco = self.entry_preco.get()

        if not nome or not quantidade or not preco:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        try:
            quantidade = int(quantidade)
            preco = float(preco)
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro e preço deve ser um número decimal.")
            return

        self.gerenciador.cadastrar_produto(nome, quantidade, preco)

    def checar_estoque(self):
        nome = self.entry_nome.get()
        self.gerenciador.checar_estoque(nome)

    def atualizar_estoque_venda(self):
        nome = self.entry_nome.get()
        quantidade = self.entry_quantidade.get()

        if not nome or not quantidade:
            messagebox.showerror("Erro", "Nome e quantidade devem ser preenchidos.")
            return

        try:
            quantidade = int(quantidade)
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
            return

        self.gerenciador.atualizar_estoque_venda(nome, quantidade)

    def adicionar_estoque(self):
        nome = self.entry_nome.get()
        quantidade = self.entry_quantidade.get()

        if not nome or not quantidade:
            messagebox.showerror("Erro", "Nome e quantidade devem ser preenchidos.")
            return

        try:
            quantidade = int(quantidade)
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
            return

        self.gerenciador.adicionar_estoque(nome, quantidade)

    def listar_estoque(self):
        self.gerenciador.listar_estoque()


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x300")  # Tamanho inicial
    interface = InterfaceEstoque(root)
    root.mainloop()
