#1
def salvar_excel(self):
        """Salva os dados da Treeview no arquivo Excel já aberto seguindo o padrão do ExcelProcessor"""
        try:
            # Verificar se há um arquivo Excel aberto
            if not hasattr(self, 'excel_abas') or not self.excel_abas or 'excel' not in self.excel_abas:
                messagebox.showerror("Erro", "Nenhum arquivo Excel aberto. Por favor, abra um arquivo primeiro.")
                return
            
            # Obter o nome da aba atual
            aba_atual = self.combo_abas.get()
            if not aba_atual:
                messagebox.showerror("Erro", "Nenhuma aba selecionada.")
                return
            
            # Obter dados da Treeview
            dados_treeview = []
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                if values:  # Verificar se há valores na linha
                    dados_treeview.append(values)
            
            if not dados_treeview:
                messagebox.showwarning("Aviso", "Nenhum dado encontrado na tabela para salvar.")
                return
            
            # Obter o caminho do arquivo original
            try:
                caminho_arquivo = self.excel_abas['excel'].io
                if not os.path.exists(caminho_arquivo):
                    messagebox.showerror("Erro", "Arquivo original não encontrado.")
                    return
            except:
                messagebox.showerror("Erro", "Não foi possível obter o caminho do arquivo.")
                return
            
            # Inicializar variáveis do Excel
            excel = None
            workbook = None
            
            try:
                # Inicializar Excel Application seguindo o padrão do ExcelProcessor
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                # Normalizar o caminho do arquivo
                caminho_arquivo = os.path.normpath(caminho_arquivo)
                
                # Abrir o arquivo existente
                workbook = excel.Workbooks.Open(caminho_arquivo)
                
                # Selecionar a planilha atual
                try:
                    sheet = workbook.Worksheets(aba_atual)
                    sheet.Select()
                except:
                    messagebox.showerror("Erro", f"Aba '{aba_atual}' não encontrada no arquivo.")
                    return
                
                # Inicializar o processador Excel para usar as funções de bloqueio/desbloqueio
                excel_processor = ExcelProcessor()
                
                # Desbloquear a planilha antes de fazer alterações
                if not excel_processor.desbloquear_planilha(workbook, sheet_name=aba_atual):
                    messagebox.showerror("Erro", "Não foi possível desbloquear a planilha para edição.")
                    return
                
                # Salvar dados na range B1:AH100 (coluna A mantida do arquivo original)
                linha_inicial = 7  # Começa na linha 7 como no ExcelProcessor
                
                for row_idx, row_data in enumerate(dados_treeview):
                    linha_atual = linha_inicial + row_idx
                    
                    # Processar cada coluna da linha
                    for col_idx, value in enumerate(row_data):
                        if col_idx >= len(self.colunas):
                            continue
                        
                        col_name = self.colunas[col_idx]
                        
                        # Mapear colunas para range B1:AH100
                        if col_name == "Descrição":
                            celula = f"B{linha_atual}"
                        elif col_name == "Atividade":
                            celula = f"C{linha_atual}"
                        elif col_name.isdigit():
                            # Colunas de dias (1-31) mapeadas para D-AH
                            dia = int(col_name)
                            if 1 <= dia <= 31:
                                # Usar a função do ExcelProcessor para obter a letra da coluna
                                coluna_letra = excel_processor.get_column_letter(dia)
                                celula = f"{coluna_letra}{linha_atual}"
                            else:
                                continue
                        else:
                            continue
                        
                        # Tratar e inserir o valor
                        if value is None or value == "":
                            sheet.Range(celula).Value = ""
                            continue
                        
                        # Tratamento específico por tipo de coluna
                        if col_name == "Item":
                            # Para coluna Item, garantir que seja número inteiro
                            try:
                                sheet.Range(celula).Value = int(float(str(value))) if str(value).strip() else ""
                            except:
                                sheet.Range(celula).Value = value
                        elif col_name in ["Descrição", "Atividade"]:
                            # Para colunas de texto
                            sheet.Range(celula).Value = str(value)
                        elif col_name.isdigit():
                            # Para colunas numéricas (dias)
                            try:
                                # Converter vírgula para ponto se necessário
                                numeric_value = str(value).replace(",", ".")
                                if numeric_value and numeric_value != "0":
                                    sheet.Range(celula).Value = float(numeric_value)
                                else:
                                    sheet.Range(celula).Value = ""
                            except:
                                # Se não conseguir converter, inserir como texto
                                sheet.Range(celula).Value = value
                
                # Recalcular fórmulas seguindo o padrão do ExcelProcessor
                sheet.Calculate()
                workbook.Application.CalculateFull()
                
                # Salvar o arquivo
                workbook.Save()
                
                # Bloquear a planilha novamente seguindo o padrão do ExcelProcessor
                excel_processor.bloquear_planilha(workbook, sheet_name=aba_atual)
                messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar dados: {str(e)}")
                # Tenta bloquear a planilha mesmo em caso de erro
                if workbook:
                    excel_processor.bloquear_planilha(workbook, sheet_name=aba_atual)
                return
            finally:
                try:
                    if workbook:
                        workbook.Close(SaveChanges=True)
                    if excel:
                        del excel
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao fechar recursos: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao manipular Excel: {str(e)}")
            if workbook:
                workbook.Close(SaveChanges=True)
            if excel: 
                del excel
            return

#2
def process_activities_to_excel(self, user_id: int, caminho_destino: str) -> bool:
        """Processa atividades do banco para o Excel"""
        self.log("\n[INFO] Iniciando processamento de atividades para Excel...")
        
        if not os.path.exists(caminho_destino):
            self.log("[ERRO] Arquivo de destino não encontrado!")
            return False

        activities = self.get_daily_activities(user_id)
        if not activities:
            self.log("[INFO] Nenhuma atividade encontrada para hoje.")
            return True

        self.log(f"[INFO] Total de atividades encontradas: {len(activities)}")
        
        excel = None
        workbook = None

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            caminho_destino = os.path.normpath(caminho_destino)
            workbook = excel.Workbooks.Open(caminho_destino)

            mes_atual = datetime.now().month
            sheet_name = self.MONTH_NAMES.get(mes_atual)
            if sheet_name:
                try:
                    sheet = workbook.Worksheets(sheet_name)
                    sheet.Select()
                except:
                    sheet = workbook.ActiveSheet
            else:
                sheet = workbook.ActiveSheet
            
            # Desbloqueia todas as planilhas antes de fazer alterações
            # Isso é necessário porque precisamos acessar meses anteriores
            if not self.desbloquear_planilha(workbook):
                self.log("[ERRO] Não foi possível desbloquear as planilhas!")
                return False

            dia_atual = datetime.now().day
            coluna_destino = self.get_column_letter(dia_atual)
            
            total_activities = len(activities)

            for idx, activity in enumerate(activities, 1):
                try:
                    self.log(f"\n[DEBUG] Processando atividade {idx}")
                    self.log(f"[DEBUG] Descrição: {activity['description']}")
                    self.log(f"[DEBUG] Atividade: {activity['atividade']}")
                    
                    linha_atual = 7
                    linha_encontrada = None
                    
                    # Procurar linha apropriada
                    while linha_atual <= 101:
                        desc_excel = sheet.Range(f"B{linha_atual}").Value
                        ativ_excel = sheet.Range(f"C{linha_atual}").Value
                        
                        self.log(f"[DEBUG] Linha {linha_atual}:")
                        self.log(f"[DEBUG] Excel Descrição: {desc_excel}")
                        self.log(f"[DEBUG] Excel Atividade: {ativ_excel}")
                        
                        if desc_excel is None:
                            # Encontrou linha vazia
                            self.log("[DEBUG] Encontrou linha vazia - usando para nova entrada")
                            linha_encontrada = linha_atual
                            break
                            
                        if desc_excel == activity['description']:
                            if ativ_excel == activity['atividade']:
                                # Descrição e atividade iguais
                                self.log("[DEBUG] Encontrou descrição e atividade iguais - usando linha existente")
                                linha_encontrada = linha_atual
                                break
                            else:
                                # Descrição igual mas atividade diferente
                                self.log("[DEBUG] Encontrou descrição igual mas atividade diferente - continuando busca")
                                
                        linha_atual += 1
                    
                    if linha_encontrada is None:
                        self.log("[DEBUG] Nenhuma linha adequada encontrada")
                        continue
                        
                    # Processamento da linha encontrada
                    celula_destino = f"{coluna_destino}{linha_encontrada}"
                    
                    # Se é uma nova linha, escreve descrição e atividade
                    if desc_excel is None:
                        self.log(f"[DEBUG] Escrevendo nova entrada na linha {linha_encontrada}")
                        sheet.Range(f"B{linha_encontrada}").Value = activity['description']
                        sheet.Range(f"C{linha_encontrada}").Value = activity['atividade']
                    
                    # Calcula valor acumulado dos dias anteriores no mês atual
                    valor_acumulado_mes_atual = 0
                    try:
                        for dia in range(1, dia_atual):
                            coluna_anterior = self.get_column_letter(dia)
                            celula_anterior = f"{coluna_anterior}{linha_encontrada}"
                            valor_anterior = sheet.Range(celula_anterior).Value
                            
                            if valor_anterior is not None:
                                self.log(f"[DEBUG] Dia {dia} - Valor anterior encontrado: {valor_anterior}")
                                if isinstance(valor_anterior, (int, float)):
                                    valor_acumulado_mes_atual += float(valor_anterior)
                        
                        self.log(f"[DEBUG] Valor total acumulado dos dias anteriores no mês atual: {valor_acumulado_mes_atual}")
                    except Exception as e:
                        self.log(f"[ERRO] Falha ao calcular valores anteriores no mês atual: {str(e)}")
                        valor_acumulado_mes_atual = 0
                    
                    # Verificar valores acumulados em meses anteriores
                    valor_acumulado_meses_anteriores = self.check_previous_months_activities(
                        workbook, 
                        activity['description'], 
                        activity['atividade'], 
                        mes_atual
                    )
                    
                    # Valor acumulado total (mês atual + meses anteriores)
                    valor_acumulado_total = valor_acumulado_mes_atual + valor_acumulado_meses_anteriores
                    self.log(f"[INFO] Valor acumulado total: {valor_acumulado_total} (mês atual: {valor_acumulado_mes_atual}, meses anteriores: {valor_acumulado_meses_anteriores})")
                    
                    # Processa o valor do tempo
                    try:
                        valor = str(activity['total_time']).replace(',', '.')
                        numero = float(valor)
                        
                        # Subtrai o valor acumulado se necessário
                        if valor_acumulado_total > 0:
                            numero_ajustado = numero - valor_acumulado_total
                            self.log(f"[DEBUG] Valor original: {numero}")
                            self.log(f"[DEBUG] Valor após subtração do acumulado total: {numero_ajustado}")
                            
                            # Se o valor ajustado for negativo, significa que já temos mais horas registradas do que o total
                            # Neste caso, não registramos nada
                            if numero_ajustado <= 0:
                                self.log(f"[INFO] Valor após subtração é zero ou negativo. Nada a registrar.")
                                continue
                                
                            numero = numero_ajustado
                        
                        # Verifica se a célula está vazia antes de escrever
                        if not sheet.Range(celula_destino).Value:
                            self.log(f"[DEBUG] Inserindo valor {numero} em {celula_destino}")
                            sheet.Range(celula_destino).Value = numero
                            sheet.Calculate()
                            workbook.Application.CalculateFull()
                    except ValueError as e:
                        self.log(f"[ERRO] Valor inválido para conversão: {valor}")
                        continue
                    
                    self.update_progress(idx, total_activities)
                    
                except Exception as e:
                    self.log(f"[ERRO] Falha ao processar atividade {idx}: {str(e)}")
                    continue

            workbook.Save()
            
            # Bloqueia todas as planilhas após as alterações
            if not self.bloquear_planilha(workbook):
                self.log("[AVISO] Não foi possível bloquear as planilhas após as alterações!")
            
            self.log("[INFO] Arquivo Excel atualizado com sucesso!")
            return True

        except Exception as e:
            self.log(f"[ERRO] Falha ao processar arquivo: {str(e)}")
            # Tenta bloquear as planilhas mesmo em caso de erro
            if workbook:
                self.bloquear_planilha(workbook)
            return False
            
        finally:
            try:
                if workbook:
                    workbook.Close(SaveChanges=True)
                if excel:
                    del excel
            except Exception as e:
                self.log(f"[ERRO] Falha ao fechar recursos: {str(e)}")