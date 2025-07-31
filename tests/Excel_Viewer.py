import customtkinter as ctk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from datetime import datetime
import pandas as pd
import os

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Visualizador de Arquivo Excel")
        self.geometry("1200x600")

        # Título
        self.label_title = ctk.CTkLabel(self, 
        text="Visualizador de Arquivo Excel", 
        font=("Arial", 20))
        self.label_title.pack(pady=10)

        # Frame para Treeview
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Scrollbars
        self.tree_scroll_y = ctk.CTkScrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side="right", fill="y")

        self.tree_scroll_x = ctk.CTkScrollbar(self.tree_frame, orientation="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")

        # Treeview
        self.tree = ttk.Treeview(self.tree_frame,
                                 yscrollcommand=self.tree_scroll_y.set,
                                 xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(fill="both", expand=True)  # <- ESSENCIAL

        self.tree_scroll_y.configure(command=self.tree.yview)
        self.tree_scroll_x.configure(command=self.tree.xview)

        # Nomes personalizados das colunas
        self.colunas = ["Item", "Descrição", "Atividade"] + [str(i) for i in range(1, 32)] + ["Total"]
        self.tree["columns"] = self.colunas
        self.tree["show"] = "headings"

        for col in self.colunas:
            self.tree.heading(col, text=col)
            if col in ["Descrição", "Atividade"]:
                self.tree.column(col, width=300, anchor="w")  # Aumente conforme necessário
            else:
                self.tree.column(col, width=80, anchor="center")


        # Frame principal de edição (acima dos botões)
        self.frame_edicao = ctk.CTkFrame(self)
        self.frame_edicao.pack(fill="x", padx=10, pady=(0, 10))

        # Frame1 - Descrição e Atividade (lado a lado com botão Ok à direita)
        self.frame1 = ctk.CTkFrame(self.frame_edicao)
        self.frame1.pack(fill="x", pady=5)

        self.entry_descricao = ctk.CTkEntry(self.frame1, placeholder_text="Descrição")
        self.entry_descricao.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

        self.entry_atividade = ctk.CTkEntry(self.frame1, placeholder_text="Atividade")
        self.entry_atividade.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        self.btn_ok = ctk.CTkButton(
            self.frame1,
            text="Ok",
            fg_color=('#ff5722', '#ff5722'),
            hover_color=("#8E8E8E", "#505050"),
            width=60,
            command=self.atualizar_linha_treeview
        )
        self.btn_ok.grid(row=0, column=2, padx=5)

        # Expandir as colunas 0 e 1
        self.frame1.grid_columnconfigure(0, weight=1)
        self.frame1.grid_columnconfigure(1, weight=1)

        # Frame2 - Entradas de 1 a 31 (abaixo do frame1)
        self.frame2 = ctk.CTkFrame(self.frame_edicao)
        self.frame2.pack(fill="x", expand=True, pady=1)

        self.entries_dias = []
        for i in range(31):
            entry = ctk.CTkEntry(self.frame2, placeholder_text=str(i + 1))
            entry.grid(row=0, column=i, padx=1, pady=1, sticky="ew")
            self.frame2.grid_columnconfigure(i, weight=1)
            self.entries_dias.append(entry)

        self.tree.bind("<<TreeviewSelect>>", self.preencher_campos_edicao)

        # Botões
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=10)

        self.btn_open = ctk.CTkButton(
            self.button_frame, 
            text="Abrir Excel", 
            fg_color=('#ff5722', '#ff5722'),
            hover_color=("#8E8E8E", "#505050"),
            command=self.abrir_excel
        )
        self.btn_open.pack(side="left", padx=10)

        self.combo_abas = ctk.CTkOptionMenu(
            self.button_frame,
            values=["Meses"],
            command=self.carregar_aba_selecionada,
            fg_color="#505050",            # Cor de fundo do botão
            text_color="white",            # Cor do texto
            dropdown_fg_color="#505050",   # Cor de fundo do menu suspenso
            dropdown_text_color="white",   # Cor do texto da lista
            button_color="#ff5722",        # Cor do botão de seta
        )
        self.combo_abas.pack(side="left", padx=10)

        self.excel_abas = {}  # Guarda o ExcelFile e DataFrame para acesso posterior

        self.btn_salvar = ctk.CTkButton(
            self.button_frame, 
            text="Salvar",
            fg_color=('#ff5722', '#ff5722'),
            hover_color=("#8E8E8E", "#505050"),
            command=self.salvar_excel
        )
        self.btn_salvar.pack(side="left", padx=10)

    def carregar_aba_selecionada(self, nome_aba):
        if not nome_aba or "excel" not in self.excel_abas:
            return

        try:
            xls = self.excel_abas["excel"]
            df = pd.read_excel(xls, sheet_name=nome_aba, usecols="A:AI", skiprows=5, nrows=101)
            df.columns = df.columns.map(str)
            self.label_title.configure(text=f"Visualizador - {self.excel_abas['arquivo']} ({nome_aba})")
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Erro", f"Erro ao carregar aba '{nome_aba}':\n{e}")
            return

        self.tree.delete(*self.tree.get_children())
        df = df.fillna("")

        for _, row in df.iterrows():
            valores = []
            for col in self.colunas:
                valor = row.get(col, "")
                if col == "Item":
                    try:
                        valor = int(float(valor)) if str(valor).strip() else ""
                    except:
                        pass
                elif col.isdigit() or col == "Total":
                    try:
                        if isinstance(valor, float):
                            texto = str(valor).rstrip("0").rstrip(".")
                            valor = texto.replace(".", ",")
                        elif isinstance(valor, str) and valor.replace(".", "", 1).isdigit():
                            valor = valor.replace(".", ",")
                    except:
                        pass
                valores.append(valor)
            self.tree.insert("", "end", values=valores)

    def atualizar_linha_treeview(self):
        selected = self.tree.focus()
        if not selected:
            return

        # Recuperar os valores editados
        nova_descricao = self.entry_descricao.get()
        nova_atividade = self.entry_atividade.get()
        novos_dias = [entry.get() for entry in self.entries_dias]

        # Obter valores atuais da linha
        valores = self.tree.item(selected)["values"]

        if not valores:
            return

        # Atualizar valores
        valores[1] = nova_descricao
        valores[2] = nova_atividade
        for i in range(31):
            if 3 + i < len(valores):
                valores[3 + i] = novos_dias[i]

        self.tree.item(selected, values=valores)

    def preencher_campos_edicao(self, event):
        selected = self.tree.focus()
        if not selected:
            return
        valores = self.tree.item(selected)["values"]
        if not valores:
            return

        # Descrição
        if len(valores) > 1 and valores[1] != "":
            self.entry_descricao.delete(0, "end")
            self.entry_descricao.insert(0, valores[1])
        else:
            self.entry_descricao = self.resetar_entry(self.entry_descricao, self.frame1, "Descrição", 500)

        # Atividade
        if len(valores) > 2 and valores[2] != "":
            self.entry_atividade.delete(0, "end")
            self.entry_atividade.insert(0, valores[2])
        else:
            self.entry_atividade = self.resetar_entry(self.entry_atividade, self.frame1, "Atividade", 500)

        # Colunas 1 a 31
        for i in range(31):
            idx = 3 + i
            if len(valores) > idx and valores[idx] != "":
                self.entries_dias[i].delete(0, "end")
                self.entries_dias[i].insert(0, valores[idx])
            else:
                self.entries_dias[i] = self.resetar_entry(self.entries_dias[i], self.frame2, str(i + 1), 35)

    def resetar_entry(self, entry, parent, placeholder, width):
        index = None
        if parent == self.frame2:
            index = self.entries_dias.index(entry)

        entry.destroy()

        # JUSTIFICAR AO CENTRO se for entrada dos dias
        if parent == self.frame2:
            novo_entry = ctk.CTkEntry(parent, placeholder_text=placeholder, width=width, justify="center")
        else:
            novo_entry = ctk.CTkEntry(parent, placeholder_text=placeholder, width=width)

        if parent == self.frame1:
            col = 0 if placeholder == "Descrição" else 1
            novo_entry.grid(row=0, column=col, padx=5, pady=2, sticky="ew")
        else:
            novo_entry.grid(row=0, column=index, padx=2, pady=2, sticky="ew")
            parent.grid_columnconfigure(index, weight=1)

        return novo_entry


    def abrir_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            # Lista dos meses em português
            meses_pt = [
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            ]
            mes_atual = meses_pt[datetime.now().month - 1]

            # Abrir arquivo Excel e pegar as abas
            xls = pd.ExcelFile(file_path)
            abas_disponiveis = xls.sheet_names

            # Armazena referência
            self.excel_abas = {
                "excel": xls,
                "arquivo": os.path.splitext(os.path.basename(file_path))[0]
            }

            # Preencher o menu suspenso com CTkOptionMenu
            self.combo_abas.configure(values=abas_disponiveis)

            # Selecionar e carregar aba do mês atual se existir
            if mes_atual in abas_disponiveis:
                self.combo_abas.set(mes_atual)
                self.carregar_aba_selecionada(mes_atual)
            else:
                self.combo_abas.set(abas_disponiveis[0])
                self.carregar_aba_selecionada(abas_disponiveis[0])

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo:\n{e}")
            return

    def salvar_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            data.append(values)
        df = pd.DataFrame(data, columns=self.tree["columns"])
        df.to_excel(file_path, index=False)
        ctk.CTkMessagebox(title="Sucesso", message="Arquivo salvo com sucesso!")

if __name__ == "__main__":
    app = ExcelViewerApp()
    app.mainloop()
