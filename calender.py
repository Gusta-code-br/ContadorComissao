import tkinter as tk
from tkinter import ttk
import pandas as pd
import customtkinter as ctk

class TabelaTkinter:
    def __init__(self, root):

        # Criar uma Treeview (tabela)
        self.tree = ttk.Treeview(root)
        self.tree["columns"] = ("ID", "Nome", "Função", "Valor do Serviço", "Comissão", "Data")

        # Formatar as colunas
        for col in self.tree["columns"]:
            self.tree.column(col, anchor=tk.W, width=100)
            self.tree.heading(col, text=col, anchor=tk.W)

        # Adicionar a tabela à janela
        self.tree.pack(expand=True, fill=tk.BOTH)

        # Carregar dados do Excel
        self.carregar_dados_excel()

    def carregar_dados_excel(self):
        # Ler dados do Excel usando pandas
        try:
            df = pd.read_excel("relatorio.xlsx")
        except FileNotFoundError:
            # Se o arquivo não for encontrado, exiba uma mensagem
            print("Arquivo Excel não encontrado.")
            return

        # Adicionar dados à tabela
        for _, row in df.iterrows():
            values = tuple(row)
            self.tree.insert("", tk.END, values=values)

def main():
    root = ctk.CTk()
    app = TabelaTkinter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
