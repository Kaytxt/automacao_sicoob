import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

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
        print("\nO programa continuará, mas pode ter limitações na leitura de arquivos Excel.\n")

def processar_formato_valor_sicoob(valor_str):
    """
    Processa o formato específico do Sicoob: '- 125,69 D' ou '2.794,76 C'
    Retorna valor numérico (positivo para débitos, None para créditos)
    """
    if not valor_str or pd.isna(valor_str):
        return None
    
    valor_str = str(valor_str).strip()
    
    # Ignora valores vazios
    if not valor_str:
        return None
    
    # Remove espaços extras
    valor_str = re.sub(r'\s+', ' ', valor_str)
    
    # Padrão para formato Sicoob: opcionalmente "- " seguido de número com vírgula e " D" ou " C"
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
        else:
            return None  # Créditos são ignorados
    
    # Se não conseguiu processar no formato Sicoob, tenta formato genérico
    try:
        # Remove tudo que não é dígito, vírgula ou ponto
        valor_limpo = re.sub(r'[^\d.,-]', '', valor_str)
        if valor_limpo:
            valor_limpo = valor_limpo.replace(',', '.')
            valor_numerico = float(valor_limpo)
            # Se tinha sinal de menos no original, considera como débito
            if '-' in valor_str:
                return valor_numerico
    except:
        pass
    
    return None

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
            
        # Lê todas as abas da planilha fixa de uma vez
        all_sheets = pd.read_excel(caminho_planilha_fixa, sheet_name=None, engine='openpyxl')
        df_banco = all_sheets['Banco']
        
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
        
        # --- 3. CONSOLIDANDO A DESCRIÇÃO LINHA POR LINHA (VERSÃO MELHORADA PARA SICOOB) ---
        registros_consolidados = []
        historico_atual = ""
        linha_principal = None
        transacoes_processadas = 0
        
        print("🔄 Processando e consolidando transações...")
        
        for index, row in df_extrato.iterrows():
            # Converte valores para string para processamento
            data_str = str(row['DATA']).strip() if pd.notna(row['DATA']) else ""
            historico_str = str(row['HISTORICO']).strip() if pd.notna(row['HISTORICO']) else ""
            
            # Se tem DATA válida, é uma linha principal de transação
            if data_str and data_str != "" and data_str != "nan":
                # Se já tínhamos uma linha principal anterior, salva ela
                if linha_principal is not None:
                    linha_principal['HISTORICO'] = historico_atual.strip()
                    registros_consolidados.append(linha_principal.copy())
                    transacoes_processadas += 1
                
                # Inicia uma nova linha principal
                linha_principal = row.copy()
                historico_atual = historico_str
                
            # Se não tem DATA, é uma linha de continuação da descrição
            elif linha_principal is not None and historico_str and historico_str != "":
                # Adiciona o conteúdo ao histórico atual (com espaço)
                if historico_atual:
                    historico_atual += " " + historico_str
                else:
                    historico_atual = historico_str
        
        # Não esqueça de adicionar a última linha
        if linha_principal is not None:
            linha_principal['HISTORICO'] = historico_atual.strip()
            registros_consolidados.append(linha_principal.copy())
            transacoes_processadas += 1
        
        print(f"🔄 Transações consolidadas: {transacoes_processadas}")
        
        # Converte a lista de volta para DataFrame
        df_extrato_consolidado = pd.DataFrame(registros_consolidados)

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
        print(f"🚫 Linhas de saldo ignoradas: {linhas_filtradas}")

        # --- 5. PROCESSAMENTO ESPECÍFICO DOS VALORES DO SICOOB ---
        df_extrato_consolidado = df_extrato_consolidado.dropna(subset=['DATA'])
        
        # Limpar espaços extras no histórico
        df_extrato_consolidado['HISTORICO'] = df_extrato_consolidado['HISTORICO'].str.replace(r'\s+', ' ', regex=True).str.strip()
        
        # Processar valores no formato específico do Sicoob
        print("💰 Processando valores no formato Sicoob...")
        valores_processados = []
        valores_validos = 0
        
        for index, row in df_extrato_consolidado.iterrows():
            valor_processado = processar_formato_valor_sicoob(row['VALOR'])
            valores_processados.append(valor_processado)
            if valor_processado is not None:
                valores_validos += 1
        
        df_extrato_consolidado['VALOR_PROCESSADO'] = valores_processados
        
        print(f"💰 Valores de débito processados: {valores_validos}")
        
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
            
        # --- 9. ADICIONANDO OS DADOS E SALVANDO NA PLANILHA FIXA ---
        df_banco_atualizado = pd.concat([df_banco, novos_lancamentos_sem_duplicatas], ignore_index=True)

        # Ordenar por data
        try:
            df_banco_atualizado['Data Vencimento'] = pd.to_datetime(df_banco_atualizado['Data Vencimento'], dayfirst=True, errors='coerce')
            df_banco_atualizado.sort_values(by='Data Vencimento', inplace=True)
        except:
            pass
        
        # Salvar na planilha preservando outras abas
        with pd.ExcelWriter(caminho_planilha_fixa, engine='openpyxl', mode='w') as writer:
            # Primeiro, salva a aba 'Banco' atualizada
            df_banco_atualizado.to_excel(writer, sheet_name='Banco', index=False)
            
            # Em seguida, salva as outras abas
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
                            f"💰 Transações de débito encontradas: {len(df_extrato_debitos)}\n"
                            f"🚫 Duplicatas ignoradas: {duplicatas_encontradas}\n"
                            f"➕ Novos lançamentos adicionados: {len(novos_lancamentos_sem_duplicatas)}\n\n"
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
    verificar_dependencias()
    
    print("📊 Iniciando processamento de extrato bancário...")
    processar_extrato_e_transferir()