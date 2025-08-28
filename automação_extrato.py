import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
from datetime import datetime

def verificar_dependencias():
    """Verifica se as bibliotecas necessárias estão instaladas"""
    bibliotecas_faltando = []
    
    try:
        import openpyxl
    except ImportError:
        bibliotecas_faltando.append("openpyxl")
    
    try:
        import chardet
    except ImportError:
        print("⚠️ Biblioteca 'chardet' não encontrada (opcional para detecção automática de encoding)")
    
    if bibliotecas_faltando:
        print("⚠️ Bibliotecas em falta:")
        for lib in bibliotecas_faltando:
            print(f"   - {lib}")
        print("\n💡 Para instalar as bibliotecas faltando, execute:")
        print(f"pip install {' '.join(bibliotecas_faltando)}")
        return False
    return True

def processar_formato_valor_sicoob(valor_str):
    """
    Processa o formato específico do Sicoob: '- 125,69 D' ou '2.794,76 C'
    Retorna valor numérico (positivo para débitos, None para créditos)
    """
    if not valor_str or pd.isna(valor_str):
        return None
    
    valor_str = str(valor_str).strip()
    
    # Ignora valores vazios
    if not valor_str or valor_str == "nan":
        return None
    
    # Remove espaços extras
    valor_str = re.sub(r'\s+', ' ', valor_str)
    
    # Padrão para formato Sicoob: opcionalmente "- " seguido de número com vírgula e " D" ou " C"
    # Exemplos: "- 125,69 D", "2.794,76 C", "- 2.460,73 D"
    padrao = r'^-?\s*(\d{1,3}(?:\.\d{3})*,\d{2})\s*([DC])$'
    match = re.match(padrao, valor_str)
    
    if match:
        numero_str = match.group(1)  # Ex: "125,69" ou "2.460,73"
        tipo = match.group(2)        # "D" para débito ou "C" para crédito
        
        # Converte para float (troca vírgula por ponto, remove pontos dos milhares)
        numero_str = numero_str.replace('.', '').replace(',', '.')
        valor_numerico = float(numero_str)
        
        # Se for débito (D), retorna positivo (pois vamos remover o sinal de menos depois)
        # Se for crédito (C), retorna None (pois não queremos incluir)
        if tipo == 'D':
            return valor_numerico
        elif tipo == 'C':
            return None  # Créditos são explicitamente ignorados
    
    # Se não conseguiu processar no formato Sicoob, tenta formato genérico
    try:
        # Verifica se tem indicador de crédito
        if 'C' in valor_str.upper():
            return None  # Ignora créditos
        
        # Remove tudo que não é dígito, vírgula, ponto ou sinal de menos
        valor_limpo = re.sub(r'[^\d.,-]', '', valor_str)
        if valor_limpo:
            valor_limpo = valor_limpo.replace(',', '.')
            valor_numerico = float(valor_limpo)
            # Se tinha sinal de menos no original ou não tem indicador de crédito, considera como débito
            if '-' in valor_str or 'D' in valor_str.upper():
                return valor_numerico
    except:
        pass
    
    return None

def adicionar_dados_preservando_formatacao(caminho_planilha, novos_dados):
    """
    Adiciona novos dados à planilha preservando toda a formatação original
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import NamedStyle
        from openpyxl.utils import get_column_letter
        
        print("🎨 Carregando planilha preservando formatação...")
        
        # Carrega a planilha mantendo formatação
        wb = load_workbook(caminho_planilha)
        ws = wb['Banco']
        
        # Encontra a primeira linha vazia (após os dados existentes)
        linha_inicio = 1
        while linha_inicio <= ws.max_row:
            # Verifica se a linha está completamente vazia
            linha_vazia = True
            for col in range(1, 8):  # Colunas A até G
                cell_value = ws.cell(row=linha_inicio, column=col).value
                if cell_value is not None and str(cell_value).strip() != "":
                    linha_vazia = False
                    break
            
            if linha_vazia and linha_inicio > 1:  # Não conta a linha de cabeçalho
                break
            linha_inicio += 1
        
        print(f"📍 Iniciando inserção na linha {linha_inicio}")
        
        # Copia a formatação da linha de cabeçalho ou da última linha com dados
        linha_formato_referencia = 1 if linha_inicio <= 2 else linha_inicio - 1
        
        # Adiciona os novos dados
        for i, (index, row) in enumerate(novos_dados.iterrows()):
            linha_atual = linha_inicio + i
            
            # Data Vencimento (Coluna A)
            cell_data = ws.cell(row=linha_atual, column=1)
            try:
                # Converte string para data se necessário
                if isinstance(row['Data Vencimento'], str):
                    data_obj = datetime.strptime(row['Data Vencimento'], '%d/%m/%Y')
                    cell_data.value = data_obj
                else:
                    cell_data.value = row['Data Vencimento']
                # Aplica formatação de data
                cell_data.number_format = 'DD/MM/YYYY'
            except:
                cell_data.value = row['Data Vencimento']
            
            # Descrição (Coluna B)
            ws.cell(row=linha_atual, column=2, value=row['Descrição'])
            
            # Valor (Coluna C) - com formatação de moeda brasileira
            cell_valor = ws.cell(row=linha_atual, column=3)
            try:
                cell_valor.value = float(row['Valor'])
                cell_valor.number_format = 'R$ #,##0.00'
            except:
                cell_valor.value = row['Valor']
            
            # Fornecedor (Coluna D)
            ws.cell(row=linha_atual, column=4, value=row['Fornecedor'])
            
            # Numero Docto (Coluna E)
            ws.cell(row=linha_atual, column=5, value=row['Numero Docto'])
            
            # Conta Contábil (Coluna F)
            ws.cell(row=linha_atual, column=6, value=row['Conta Contábil'])
            
            # Observação (Coluna G)
            ws.cell(row=linha_atual, column=7, value=row['Observação (opcional)'])
            
            # Copia formatação da linha de referência (borda, alinhamento, etc.)
            if linha_formato_referencia > 0:
                for col in range(1, 8):
                    cell_origem = ws.cell(row=linha_formato_referencia, column=col)
                    cell_destino = ws.cell(row=linha_atual, column=col)
                    
                    # Copia formatação (exceto número que já definimos)
                    if cell_origem.font:
                        cell_destino.font = cell_origem.font
                    if cell_origem.border:
                        cell_destino.border = cell_origem.border
                    if cell_origem.fill:
                        cell_destino.fill = cell_origem.fill
                    if cell_origem.alignment:
                        cell_destino.alignment = cell_origem.alignment
        
        # Ajusta largura das colunas se necessário
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)  # Máximo de 50
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Salva a planilha
        wb.save(caminho_planilha)
        print(f"✅ Dados adicionados preservando formatação original!")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao preservar formatação: {e}")
        return False

def processar_extrato_e_transferir():
    """
    Função principal que automatiza o processo de transferência de dados do extrato
    para a planilha fixa 'Automação_Gransoft.xlsx'.
    """
    
    root = tk.Tk()
    root.withdraw()

    caminho_extrato = filedialog.askopenfilename(
        title="Selecione o arquivo de extrato do Sicoob",
        filetypes=[("Arquivos Excel/CSV", "*.xlsx;*.xls;*.csv")]
    )
    if not caminho_extrato:
        messagebox.showinfo("Aviso", "Operação cancelada. Nenhum arquivo de extrato foi selecionado.")
        return False
        
    caminho_script = os.path.dirname(os.path.abspath(__file__))
    caminho_planilha_fixa = os.path.join(caminho_script, 'Automação_Gransoft.xlsx')
    
    if not os.path.exists(caminho_planilha_fixa):
        messagebox.showerror("Erro", 
                            f"A planilha 'Automação_Gransoft.xlsx' não foi encontrada na pasta do script.\n"
                            f"Certifique-se de que o arquivo esteja na mesma pasta que o programa.")
        return False

    print(f"Processando extrato: {os.path.basename(caminho_extrato)}...")
    print(f"Atualizando a planilha fixa: Automação_Gransoft.xlsx")

    try:
        # --- 1. LEITURA ROBUSTA DOS ARQUIVOS ---
        df_extrato = None
        
        # Primeiro tenta ler como Excel (se openpyxl estiver disponível)
        try:
            df_extrato = pd.read_excel(caminho_extrato, skiprows=1, header=0)
            print("✅ Arquivo lido como Excel")
        except ImportError:
            print("⚠️ Biblioteca 'openpyxl' não está instalada. Tentando como CSV...")
            print("💡 Para usar arquivos Excel, instale: pip install openpyxl")
        except Exception as excel_error:
            print(f"⚠️ Erro ao ler como Excel: {excel_error}")
            
        # Se não conseguiu ler como Excel, tenta CSV com diferentes encodings
        if df_extrato is None:
            csv_lido = False
            estrategias_csv = [
                # Estratégia 1: Windows-1252 (comum em arquivos do Windows/bancos)
                {'sep': ',', 'quotechar': '"', 'encoding': 'windows-1252'},
                # Estratégia 2: CP1252 (alternativa para Windows)
                {'sep': ',', 'quotechar': '"', 'encoding': 'cp1252'},
                # Estratégia 3: ISO-8859-1 (Latin-1)
                {'sep': ',', 'quotechar': '"', 'encoding': 'iso-8859-1'},
                # Estratégia 4: Windows-1252 com ponto e vírgula
                {'sep': ';', 'quotechar': '"', 'encoding': 'windows-1252'},
                # Estratégia 5: UTF-8 com BOM
                {'sep': ',', 'quotechar': '"', 'encoding': 'utf-8-sig'},
                # Estratégia 6: Detecção automática com chardet
                {'sep': ',', 'quotechar': '"', 'encoding': None},  # Será detectado automaticamente
                # Estratégia 7: UTF-8 padrão
                {'sep': ',', 'quotechar': '"', 'encoding': 'utf-8'},
            ]
            
            for i, estrategia in enumerate(estrategias_csv):
                try:
                    print(f"🔄 Tentando estratégia CSV {i+1} (encoding: {estrategia.get('encoding', 'auto')})...")
                    
                    # Se encoding for None, tenta detectar automaticamente
                    if estrategia['encoding'] is None:
                        try:
                            import chardet
                            with open(caminho_extrato, 'rb') as f:
                                raw_data = f.read()
                                result = chardet.detect(raw_data)
                                detected_encoding = result['encoding']
                                estrategia['encoding'] = detected_encoding
                                print(f"🔍 Encoding detectado: {detected_encoding}")
                        except ImportError:
                            print("⚠️ Biblioteca 'chardet' não disponível, usando fallback")
                            estrategia['encoding'] = 'windows-1252'
                        except Exception:
                            estrategia['encoding'] = 'windows-1252'
                    
                    df_extrato = pd.read_csv(
                        caminho_extrato, 
                        skiprows=1, 
                        header=None,
                        on_bad_lines='skip',  # Pula linhas problemáticas
                        engine='python',      # Mais flexível
                        **estrategia
                    )
                    
                    # Verifica se conseguiu ler dados válidos
                    if not df_extrato.empty and df_extrato.shape[1] >= 3:
                        print(f"✅ Arquivo lido como CSV com estratégia {i+1} (encoding: {estrategia['encoding']})")
                        csv_lido = True
                        break
                        
                except Exception as csv_error:
                    print(f"❌ Estratégia CSV {i+1} falhou: {csv_error}")
                    continue
            
            if not csv_lido:
                raise Exception(
                    "Não foi possível ler o arquivo em nenhum formato suportado.\n\n"
                    "Soluções possíveis:\n"
                    "1. Instale o openpyxl para arquivos Excel: pip install openpyxl\n"
                    "2. Salve o extrato como arquivo .xlsx (Excel) em vez de .csv\n"
                    "3. Verifique se o arquivo não está corrompido"
                )
            
        # Lê apenas os dados existentes da planilha fixa para comparação (sem formatação)
        df_banco = pd.read_excel(caminho_planilha_fixa, sheet_name='Banco', engine='openpyxl')
        
        # --- 2. PREPARO DOS DADOS DO EXTRATO ---
        # Reorganiza para ter sempre as 4 colunas principais
        colunas_necessarias = ['DATA', 'DOCUMENTO', 'HISTORICO', 'VALOR']
        
        # Define as colunas baseado no que foi lido
        if df_extrato.shape[1] >= 4:
            df_extrato.columns = ['DATA', 'DOCUMENTO', 'HISTORICO', 'VALOR'] + [f'EXTRA_{i}' for i in range(df_extrato.shape[1] - 4)]
        elif df_extrato.shape[1] == 3:
            df_extrato.columns = ['DATA', 'HISTORICO', 'VALOR']
            df_extrato['DOCUMENTO'] = ''  # Adiciona coluna vazia para documento
        else:
            raise Exception(f"Arquivo tem estrutura inesperada com {df_extrato.shape[1]} colunas")
        
        for coluna in colunas_necessarias:
            if coluna not in df_extrato.columns:
                df_extrato[coluna] = ''
        
        # Seleciona apenas as colunas necessárias
        df_extrato = df_extrato[colunas_necessarias].copy()
        
        # Validação básica dos dados
        if df_extrato.empty:
            raise Exception("O arquivo está vazio ou não contém dados válidos")
        
        # Remove linhas completamente vazias
        df_extrato = df_extrato.dropna(how='all')
        
        if len(df_extrato) == 0:
            raise Exception("Não foram encontrados dados válidos no arquivo após limpeza")
        
        print(f"📊 Arquivo carregado: {len(df_extrato)} linhas, {df_extrato.shape[1]} colunas")
        
        # --- 3. CONSOLIDANDO A DESCRIÇÃO LINHA POR LINHA (VERSÃO CORRIGIDA PARA CRÉDITOS) ---
        registros_consolidados = []
        historico_atual = ""
        linha_principal = None
        transacoes_processadas = 0
        linhas_credito_ignoradas = 0
        
        print("🔄 Processando e consolidando transações...")
        
        for index, row in df_extrato.iterrows():
            # Converte valores para string para processamento
            data_str = str(row['DATA']).strip() if pd.notna(row['DATA']) else ""
            historico_str = str(row['HISTORICO']).strip() if pd.notna(row['HISTORICO']) else ""
            valor_str = str(row['VALOR']).strip() if pd.notna(row['VALOR']) else ""
            
            # Se tem DATA válida, pode ser uma linha principal de transação
            if data_str and data_str != "" and data_str != "nan":
                
                # PRIMEIRO: Verifica se esta linha tem crédito ou deve ser ignorada
                eh_credito = False
                eh_saldo = False
                
                # Verifica se é uma linha de crédito (termina com "C")
                if valor_str and ("C" in valor_str or "c" in valor_str.lower()):
                    eh_credito = True
                
                # Verifica se é uma linha de saldo (mesmo sem valor)
                frases_saldo = ['SALDO DO DIA', 'SALDO ANTERIOR', 'SALDO ATUAL', 'SALDO FINAL']
                for frase in frases_saldo:
                    if frase.lower() in historico_str.lower():
                        eh_saldo = True
                        break
                
                # Se for crédito ou saldo, ignora esta linha completamente
                if eh_credito or eh_saldo:
                    linhas_credito_ignoradas += 1
                    print(f"🚫 Ignorando linha de crédito/saldo: {historico_str[:50]}{'...' if len(historico_str) > 50 else ''}")
                    continue  # Pula esta linha sem fazer nada
                
                # Se chegou aqui, é uma transação de débito válida
                # Se já tínhamos uma linha principal anterior, salva ela
                if linha_principal is not None:
                    linha_principal['HISTORICO'] = historico_atual.strip()
                    registros_consolidados.append(linha_principal.copy())
                    transacoes_processadas += 1
                
                # Inicia uma nova linha principal (apenas se não for crédito/saldo)
                linha_principal = row.copy()
                historico_atual = historico_str
                
            # Se não tem DATA, é uma linha de continuação da descrição
            elif linha_principal is not None and historico_str and historico_str != "":
                # Verifica se esta continuação também não é uma linha de saldo
                eh_continuacao_saldo = False
                frases_saldo = ['SALDO DO DIA', 'SALDO ANTERIOR', 'SALDO ATUAL', 'SALDO FINAL']
                for frase in frases_saldo:
                    if frase.lower() in historico_str.lower():
                        eh_continuacao_saldo = True
                        break
                
                if not eh_continuacao_saldo:
                    # Adiciona o conteúdo ao histórico atual (com espaço)
                    if historico_atual:
                        historico_atual += " " + historico_str
                    else:
                        historico_atual = historico_str
        
        # Não esqueça de adicionar a última linha (se não for crédito)
        if linha_principal is not None:
            linha_principal['HISTORICO'] = historico_atual.strip()
            registros_consolidados.append(linha_principal.copy())
            transacoes_processadas += 1
        
        print(f"🚫 Linhas de crédito/saldo ignoradas durante consolidação: {linhas_credito_ignoradas}")
        print(f"🔄 Transações de débito consolidadas: {transacoes_processadas}")
        
        # Converte a lista de volta para DataFrame
        df_extrato_consolidado = pd.DataFrame(registros_consolidados)
        
        if df_extrato_consolidado.empty:
            print("⚠️ Nenhuma transação de débito foi encontrada após consolidação")
            messagebox.showinfo("Aviso", "Nenhuma transação de débito válida foi encontrada no extrato.")
            return True

        # --- 4. FILTRAR DADOS INDESEJADOS (SALDO, ETC.) ---
        frases_a_ignorar = [
            'SALDO DO DIA', 'SALDO ANTERIOR', 'SALDO ATUAL', 'SALDO FINAL',
            'saldo do dia', 'saldo anterior', 'saldo atual', 'saldo final',
            'Saldo bloqueado anterior', 'Saldo bloqueado', 'Saldo disponível', 'Saldo em conta'
        ]
        
        linhas_antes_filtro = len(df_extrato_consolidado)
        
        # Remove linhas que contêm frases de saldo
        for frase in frases_a_ignorar:
            df_extrato_consolidado = df_extrato_consolidado[
                ~df_extrato_consolidado['HISTORICO'].str.contains(frase, case=False, na=False)
            ]
        
        linhas_filtradas = linhas_antes_filtro - len(df_extrato_consolidado)
        print(f"🚫 Linhas de saldo ignoradas no filtro adicional: {linhas_filtradas}")

        # --- 5. PROCESSAMENTO ESPECÍFICO DOS VALORES DO SICOOB ---
        df_extrato_consolidado = df_extrato_consolidado.dropna(subset=['DATA'])
        
        # Limpar espaços extras no histórico
        df_extrato_consolidado['HISTORICO'] = df_extrato_consolidado['HISTORICO'].str.replace(r'\s+', ' ', regex=True).str.strip()
        
        # Processar valores no formato específico do Sicoob
        print("💰 Processando valores no formato Sicoob...")
        valores_processados = []
        valores_validos = 0
        valores_credito_ignorados = 0
        
        for index, row in df_extrato_consolidado.iterrows():
            valor_original = row['VALOR']
            valor_processado = processar_formato_valor_sicoob(valor_original)
            valores_processados.append(valor_processado)
            
            if valor_processado is not None:
                valores_validos += 1
                print(f"✅ Débito: {valor_original} → R$ {valor_processado:.2f}")
            elif str(valor_original).strip() and 'C' in str(valor_original).upper():
                valores_credito_ignorados += 1
                print(f"🚫 Crédito ignorado: {valor_original}")
        
        df_extrato_consolidado['VALOR_PROCESSADO'] = valores_processados
        
        print(f"💰 Valores de débito processados: {valores_validos}")
        print(f"🚫 Valores de crédito ignorados: {valores_credito_ignorados}")
        
        # --- 6. FILTRAR APENAS VALORES DE DÉBITO VÁLIDOS ---
        df_extrato_debitos = df_extrato_consolidado[df_extrato_consolidado['VALOR_PROCESSADO'].notna()].copy()
        
        if df_extrato_debitos.empty:
            messagebox.showinfo("Aviso", "⚠️ Nenhuma transação de débito válida foi encontrada no extrato.")
            return True
        
        print(f"📊 Total de débitos encontrados: {len(df_extrato_debitos)}")
        
        # --- 7. MAPEAMENTO PARA A ESTRUTURA DA PLANILHA FIXA ---
        novos_lancamentos = pd.DataFrame({
            'Data Vencimento': df_extrato_debitos['DATA'],
            'Descrição': df_extrato_debitos['HISTORICO'],
            'Valor': df_extrato_debitos['VALOR_PROCESSADO'],  # Usa o valor processado
            'Fornecedor': '',
            'Numero Docto': df_extrato_debitos['DOCUMENTO'],
            'Conta Contábil': '',
            'Observação (opcional)': df_extrato_debitos['HISTORICO']
        })

        # --- 8. PREVENÇÃO DE DUPLICIDADE ---
        colunas_para_comparar = ['Data Vencimento', 'Descrição', 'Valor']
        df_banco_temp = df_banco.dropna(subset=colunas_para_comparar).copy()
        novos_lancamentos_temp = novos_lancamentos.dropna(subset=colunas_para_comparar).copy()
        
        # Normalizar datas para comparação
        try:
            df_banco_temp['Data Vencimento'] = pd.to_datetime(df_banco_temp['Data Vencimento'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
            novos_lancamentos_temp['Data Vencimento'] = pd.to_datetime(novos_lancamentos_temp['Data Vencimento'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
        except:
            pass
        
        # Criar ID único para comparação
        df_banco_temp['ID'] = (
            df_banco_temp['Data Vencimento'].astype(str).str.strip() + '|' +
            df_banco_temp['Descrição'].astype(str).str.strip().str.lower() + '|' +
            df_banco_temp['Valor'].astype(str)
        )

        novos_lancamentos_temp['ID'] = (
            novos_lancamentos_temp['Data Vencimento'].astype(str).str.strip() + '|' +
            novos_lancamentos_temp['Descrição'].astype(str).str.strip().str.lower() + '|' +
            novos_lancamentos_temp['Valor'].astype(str)
        )

        novos_lancamentos_sem_duplicatas = novos_lancamentos[~novos_lancamentos_temp['ID'].isin(df_banco_temp['ID'])].copy()
        
        duplicatas_encontradas = len(novos_lancamentos) - len(novos_lancamentos_sem_duplicatas)
        print(f"🔍 Duplicatas ignoradas: {duplicatas_encontradas}")
        print(f"➕ Novos lançamentos a adicionar: {len(novos_lancamentos_sem_duplicatas)}")
        
        if novos_lancamentos_sem_duplicatas.empty:
            messagebox.showinfo("Aviso", "Não há novas transações para adicionar. A planilha está atualizada.")
            return True
            
        # --- 9. ADICIONANDO OS DADOS PRESERVANDO FORMATAÇÃO ---
        print("🎨 Adicionando dados preservando formatação original...")
        
        sucesso_formatacao = adicionar_dados_preservando_formatacao(
            caminho_planilha_fixa, 
            novos_lancamentos_sem_duplicatas
        )
        
        if not sucesso_formatacao:
            # Fallback: usar pandas se falhar a preservação de formatação
            print("⚠️ Fallback: usando método padrão sem preservar formatação completa")
            all_sheets = pd.read_excel(caminho_planilha_fixa, sheet_name=None, engine='openpyxl')
            df_banco_atualizado = pd.concat([df_banco, novos_lancamentos_sem_duplicatas], ignore_index=True)

            # Salvar na planilha preservando outras abas
            with pd.ExcelWriter(caminho_planilha_fixa, engine='openpyxl', mode='w') as writer:
                df_banco_atualizado.to_excel(writer, sheet_name='Banco', index=False)
                for sheet_name, df_sheet in all_sheets.items():
                    if sheet_name != 'Banco':
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        # Mostrar exemplos dos lançamentos adicionados
        print(f"\n📋 Exemplos dos lançamentos adicionados:")
        for i in range(min(5, len(novos_lancamentos_sem_duplicatas))):
            desc = novos_lancamentos_sem_duplicatas.iloc[i]['Descrição']
            valor = novos_lancamentos_sem_duplicatas.iloc[i]['Valor']
            data = novos_lancamentos_sem_duplicatas.iloc[i]['Data Vencimento']
            print(f"  {i+1}. {data} - {desc[:60]}{'...' if len(desc) > 60 else ''} - R$ {valor:.2f}")

        messagebox.showinfo("Sucesso", 
                            "✅ Automação concluída!\n\n"
                            f"📊 Transações processadas do extrato: {transacoes_processadas}\n"
                            f"🚫 Linhas de crédito/saldo ignoradas: {linhas_credito_ignoradas + valores_credito_ignorados}\n"
                            f"💰 Transações de débito encontradas: {len(df_extrato_debitos)}\n"
                            f"🔍 Duplicatas ignoradas: {duplicatas_encontradas}\n"
                            f"➕ Novos lançamentos adicionados: {len(novos_lancamentos_sem_duplicatas)}\n"
                            f"🎨 Formatação original preservada!\n\n"
                            f"O arquivo 'Automação_Gransoft.xlsx' foi atualizado com sucesso.")
        
        return True

    except Exception as e:
        error_msg = f"❌ Ocorreu um erro inesperado:\n\n{str(e)}"
        print(error_msg)
        
        # Informações adicionais de debugging
        import traceback
        error_details = traceback.format_exc()
        print(f"\n🔍 Detalhes técnicos do erro:\n{error_details}")
        
        # Se o erro for de leitura de arquivo, dá dicas específicas
        if "tokenizing" in str(e).lower() or "expected" in str(e).lower():
            error_msg += "\n\n💡 Dica: O arquivo parece ter problemas de formatação.\n"
            error_msg += "Tente salvar o extrato em formato .xlsx (Excel) em vez de .csv"
        elif "openpyxl" in str(e).lower():
            error_msg += "\n\n💡 Para corrigir este erro, execute no terminal:\n"
            error_msg += "pip install openpyxl\n\n"
            error_msg += "Ou salve o arquivo como .csv em vez de .xlsx"
        elif "encoding" in str(e).lower() or "decode" in str(e).lower():
            error_msg += "\n\n💡 Problema de codificação de caracteres.\n"
            error_msg += "Tente salvar o arquivo em formato .xlsx (Excel) que é mais confiável."
        
        messagebox.showerror("Erro", error_msg)
        return False

# Executa o processo principal
if __name__ == "__main__":
    print("🚀 Iniciando verificação do sistema...")
    if not verificar_dependencias():
        print("❌ Dependências necessárias não estão instaladas!")
        input("Pressione ENTER para sair...")
        exit()
    
    print("📊 Iniciando processamento de extrato bancário...")
    processar_extrato_e_transferir()