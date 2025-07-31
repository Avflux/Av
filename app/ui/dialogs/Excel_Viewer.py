import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime
import os
import win32com.client
from app.utils.excel_processor import ExcelProcessor
from app.utils.window_manager import WindowManager

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelViewer(ctk.CTk):
    def __init__(self, parent=None):
        super().__init__()
        # Centralizar no mesmo monitor do parent, com tamanho 1200x600
        WindowManager().position_window(self, parent=parent)
        self.protocol("WM_DELETE_WINDOW", self.fechar_excel)
        self.title("Visualizador de Arquivo Excel")

        # Título
        self.label_title = ctk.CTkLabel(
            self,
            text="Visualizador de Arquivo Excel",
            font=("Arial", 20)
        )
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
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(fill="both", expand=True)
        self.tree_scroll_y.configure(command=self.tree.yview)
        self.tree_scroll_x.configure(command=self.tree.xview)

        self.colunas = ["Item", "Descrição", "Atividade"] + [str(i) for i in range(1, 32)] + ["Total"]
        self.tree["columns"] = self.colunas
        self.tree["show"] = "headings"
        for col in self.colunas:
            self.tree.heading(col, text=col)
            if col in ["Descrição", "Atividade"]:
                self.tree.column(col, width=500, anchor="w")
            else:
                self.tree.column(col, width=60, anchor="center")

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

        self.frame1.grid_columnconfigure(0, weight=1)
        self.frame1.grid_columnconfigure(1, weight=1)

        # Frame2 - Entradas de dias (lado a lado)
        self.frame2 = ctk.CTkFrame(self.frame_edicao)
        self.frame2.pack(fill="x", expand=True, pady=1)
        self.entries_dias = []
        for i in range(31):
            entry = ctk.CTkEntry(self.frame2, placeholder_text=str(i + 1))
            entry.grid(row=0, column=i, padx=1, pady=1, sticky="ew")
            self.frame2.grid_columnconfigure(i, weight=1)
            self.entries_dias.append(entry)

        self.tree.bind("<<TreeviewSelect>>", self.preencher_campos_edicao)

        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=10)

        self.combo_abas = ctk.CTkOptionMenu(
            self.button_frame,
            values=["Meses"],
            fg_color="#505050",            # Cor de fundo do botão
            text_color="white",            # Cor do texto
            dropdown_fg_color="#505050",   # Cor de fundo do menu suspenso
            dropdown_text_color="white",   # Cor do texto da lista
            button_color="#ff5722",        # Cor do botão de seta
            command=self.carregar_aba_selecionada)
        self.combo_abas.pack(side="left", padx=10)

        self.btn_salvar = ctk.CTkButton(
            self.button_frame, text="Salvar", 
            fg_color=('#ff5722', '#ff5722'), 
            hover_color=('#8E8E8E', '#505050'),
            command=self.salvar_excel)
        self.btn_salvar.pack(side="left", padx=10)

        self.excel_path = ""
        self.workbook = None
        self.excel = None
        self.protocol("WM_DELETE_WINDOW", self.fechar_excel)

    def abrir_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            self.excel_path = file_path
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.workbook = self.excel.Workbooks.Open(os.path.normpath(file_path))
            abas_disponiveis = [sheet.Name for sheet in self.workbook.Sheets]
            self.combo_abas.configure(values=abas_disponiveis)

            # Lista dos meses em português
            meses_pt = [
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            ]
            mes_atual = meses_pt[datetime.now().month - 1]

            if mes_atual in abas_disponiveis:
                self.combo_abas.set(mes_atual)
            else:
                self.combo_abas.set(abas_disponiveis[0])

            # Garantir que a aba selecionada seja carregada após setar
            self.carregar_aba_selecionada(self.combo_abas.get())

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo:\n{e}")

    def carregar_aba_selecionada(self, nome_aba):
        try:
            # Verificar se o Excel ainda está ativo e o arquivo ainda está aberto
            if not self.excel or not self.workbook:
                if not self.excel_path or not os.path.exists(self.excel_path):
                    messagebox.showerror("Erro", "Arquivo Excel não encontrado ou não carregado.")
                    return
                self.excel = win32com.client.Dispatch("Excel.Application")
                self.excel.Visible = False
                self.excel.DisplayAlerts = False
                self.workbook = self.excel.Workbooks.Open(os.path.normpath(self.excel_path))

            sheet = self.workbook.Worksheets(nome_aba)
            self.label_title.configure(text=f"Visualizador - {os.path.basename(self.excel_path)} ({nome_aba})")
            self.tree.delete(*self.tree.get_children())

            # Carregar todos os dados de uma vez (A7:AI107)
            dados = sheet.Range("A7:AI107").Value

            # Garantir que há dados válidos
            if not dados:
                return

            for linha in dados:
                valores = []
                for idx, valor in enumerate(linha):
                    col = self.colunas[idx]
                    if valor is None:
                        valor = ""
                    elif col == "Item":
                        try:
                            valor = str(int(valor))
                        except:
                            valor = str(valor)
                    elif isinstance(valor, float) and col.isdigit():
                        valor = str(valor).replace(".", ",")
                    valores.append(valor)
                self.tree.insert("", "end", values=valores)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar aba '{nome_aba}':\n{e}")

    def salvar_excel(self):
        aba_atual = self.combo_abas.get()
        if not aba_atual:
            messagebox.showerror("Erro", "Nenhuma aba selecionada.")
            return

        dados_treeview = [self.tree.item(item)['values'] for item in self.tree.get_children() if self.tree.item(item)['values']]
        if not dados_treeview:
            messagebox.showwarning("Aviso", "Nenhum dado encontrado na tabela para salvar.")
            return

        try:
            # Verificar se o Excel ainda está ativo e o arquivo ainda está aberto
            if not self.excel or not self.workbook:
                if not self.excel_path or not os.path.exists(self.excel_path):
                    messagebox.showerror("Erro", "Arquivo Excel não encontrado ou não carregado.")
                    return
                self.excel = win32com.client.Dispatch("Excel.Application")
                self.excel.Visible = False
                self.excel.DisplayAlerts = False
                self.workbook = self.excel.Workbooks.Open(os.path.normpath(self.excel_path))

            sheet = self.workbook.Worksheets(aba_atual)
            excel_processor = ExcelProcessor()
            if not excel_processor.desbloquear_planilha(self.workbook, sheet_name=aba_atual):
                messagebox.showerror("Erro", "Não foi possível desbloquear a planilha para edição.")
                return

            linha_inicial = 7
            for row_idx, row_data in enumerate(dados_treeview):
                linha_atual = linha_inicial + row_idx
                for col_idx, value in enumerate(row_data):
                    if col_idx >= len(self.colunas):
                        continue
                    col_name = self.colunas[col_idx]
                    if col_name == "Descrição":
                        celula = f"B{linha_atual}"
                    elif col_name == "Atividade":
                        celula = f"C{linha_atual}"
                    elif col_name.isdigit():
                        dia = int(col_name)
                        if 1 <= dia <= 31:
                            coluna_letra = excel_processor.get_column_letter(dia)
                            celula = f"{coluna_letra}{linha_atual}"
                        else:
                            continue
                    else:
                        continue
                    try:
                        if value is None or value == "":
                            sheet.Range(celula).Value = ""
                        elif col_name == "Item":
                            sheet.Range(celula).Value = int(float(str(value))) if str(value).strip() else ""
                        elif col_name in ["Descrição", "Atividade"]:
                            sheet.Range(celula).Value = str(value)
                        elif col_name.isdigit():
                            numeric_value = str(value).replace(",", ".")
                            sheet.Range(celula).Value = float(numeric_value) if numeric_value and numeric_value != "0" else ""
                    except:
                        sheet.Range(celula).Value = value

            sheet.Calculate()
            self.workbook.Application.CalculateFull()
            self.workbook.Save()
            excel_processor.bloquear_planilha(self.workbook, sheet_name=aba_atual)
            messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados:\n{str(e)}")

    def atualizar_linha_treeview(self):
        selected = self.tree.focus()
        if not selected:
            return
        nova_descricao = self.entry_descricao.get()
        nova_atividade = self.entry_atividade.get()
        novos_dias = [entry.get() for entry in self.entries_dias]
        valores = self.tree.item(selected)["values"]
        if not valores:
            return
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

        # Descrição
        if len(valores) > 1 and valores[1]:
            self.entry_descricao.delete(0, "end")
            self.entry_descricao.insert(0, valores[1])
            self.entry_descricao.configure(fg_color="#505050")
        else:
            self.entry_descricao.delete(0, "end")
            self.entry_descricao.configure(placeholder_text="Descrição", fg_color=("#343638", "#2b2b2b")
)

        # Atividade
        if len(valores) > 2 and valores[2]:
            self.entry_atividade.delete(0, "end")
            self.entry_atividade.insert(0, valores[2])
            self.entry_atividade.configure(fg_color="#505050")
        else:
            self.entry_atividade.delete(0, "end")
            self.entry_atividade.configure(placeholder_text="Atividade", fg_color=('#343638', '#2b2b2b'))

        # Colunas 1 a 31
        for i in range(31):
            idx = 3 + i
            self.entries_dias[i].delete(0, "end")
            if len(valores) > idx and valores[idx]:
                self.entries_dias[i].insert(0, valores[idx])
                self.entries_dias[i].configure(fg_color="#505050")
            else:
                self.entries_dias[i].configure(placeholder_text=str(i + 1), fg_color=('#343638', '#2b2b2b'))

    def fechar_excel(self):
        if self.workbook:
            self.workbook.Close(SaveChanges=True)
            self.workbook = None
        if self.excel:
            self.excel.Quit()
            self.excel = None
        self.destroy()  