import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os

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
        try:
            df_extrato = pd.read_excel(caminho_extrato, skiprows=1, header=0)
        except:
            df_extrato = pd.read_csv(caminho_extrato, skiprows=1, header=None, encoding='utf-8', sep=',')
            df_extrato.columns = ['DATA', 'DOCUMENTO', 'HISTORICO', 'VALOR']
            
        # Lê todas as abas da planilha fixa de uma vez
        all_sheets = pd.read_excel(caminho_planilha_fixa, sheet_name=None, engine='openpyxl')
        df_banco = all_sheets['Banco']
        
        # --- 2. PREPARO DOS DADOS DO EXTRATO ---
        df_extrato.columns = ['DATA', 'DOCUMENTO', 'HISTORICO', 'VALOR']
        
        # --- 3. CONSOLIDANDO A DESCRIÇÃO EM UMA ÚNICA LINHA ---
        df_extrato['DATA'] = df_extrato['DATA'].ffill()
        df_extrato['DOCUMENTO'] = df_extrato['DOCUMENTO'].ffill()
        
        df_extrato_consolidado = df_extrato.groupby(['DATA', 'DOCUMENTO']).agg({
            'HISTORICO': lambda x: ' '.join(x.dropna().astype(str)),
            'VALOR': lambda x: x.dropna().iloc[0] if x.dropna().any() else np.nan
        }).reset_index()

        # --- 4. FILTRAR DADOS INDESEJADOS (SALDO, ETC.) ---
        frases_a_ignorar = [
            'Saldo bloqueado anterior', 'Saldo do dia', 'Saldo anterior',
            'Saldo bloqueado', 'Saldo atual', 'Saldo disponível', 'Saldo em conta',
            'SALDO FINAL', 'SALDO ANTERIOR'
        ]
        
        filtro_ignorados = '|'.join(frases_a_ignorar)
        df_extrato_consolidado = df_extrato_consolidado[
            ~df_extrato_consolidado['HISTORICO'].str.contains(filtro_ignorados, case=False, na=False)
        ].copy()

        # --- 5. LIMPEZA E CONVERSÃO DOS VALORES ---
        df_extrato_consolidado = df_extrato_consolidado.dropna(subset=['DATA'])
        
        df_extrato_consolidado['VALOR'] = pd.to_numeric(
            df_extrato_consolidado['VALOR'].astype(str).str.replace(r'[^\d.,-]', '', regex=True).str.replace(',', '.'),
            errors='coerce'
        )
        
        # --- 6. FILTRAR APENAS VALORES NEGATIVOS E CONVERTER PARA POSITIVOS ---
        df_extrato_negativos = df_extrato_consolidado[df_extrato_consolidado['VALOR'] < 0].copy()
        
        if df_extrato_negativos.empty:
            messagebox.showinfo("Aviso", "⚠️ Nenhuma transação de débito (valor negativo) foi encontrada no extrato.")
            return True
            
        df_extrato_negativos['VALOR'] = df_extrato_negativos['VALOR'].abs()
        
        # --- 7. MAPEAMENTO PARA A ESTRUTURA DA PLANILHA FIXA ---
        novos_lancamentos = pd.DataFrame({
            'Data Vencimento': df_extrato_negativos['DATA'],
            'Descrição': df_extrato_negativos['HISTORICO'],
            'Valor': df_extrato_negativos['VALOR'],
            'Fornecedor': '',
            'Numero Docto': df_extrato_negativos['DOCUMENTO'],
            'Conta Contábil': '',
            'Observação (opcional)': df_extrato_negativos['HISTORICO']
        })

        # --- 8. PREVENÇÃO DE DUPLICIDADE (FINAL) ---
        colunas_para_comparar = ['Data Vencimento', 'Descrição', 'Valor']
        df_banco_temp = df_banco.dropna(subset=colunas_para_comparar).copy()
        novos_lancamentos_temp = novos_lancamentos.dropna(subset=colunas_para_comparar).copy()
        
        df_banco_temp['Data Vencimento'] = pd.to_datetime(df_banco_temp['Data Vencimento'], dayfirst=True).dt.strftime('%d/%m/%Y')
        novos_lancamentos_temp['Data Vencimento'] = pd.to_datetime(novos_lancamentos_temp['Data Vencimento'], dayfirst=True).dt.strftime('%d/%m/%Y')
        
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
        
        if novos_lancamentos_sem_duplicatas.empty:
            messagebox.showinfo("Aviso", "Não há novas transações para adicionar. A planilha está atualizada.")
            return True
            
        # --- 9. ADICIONANDO OS DADOS E SALVANDO NA PLANILHA FIXA ---
        df_banco_atualizado = pd.concat([df_banco, novos_lancamentos_sem_duplicatas], ignore_index=True)

        df_banco_atualizado['Data Vencimento'] = pd.to_datetime(df_banco_atualizado['Data Vencimento'], dayfirst=True)
        df_banco_atualizado.sort_values(by='Data Vencimento', inplace=True)
        
        # O novo código de salvamento
        with pd.ExcelWriter(caminho_planilha_fixa, engine='openpyxl', mode='w') as writer:
            # Primeiro, salva a aba 'Banco' atualizada
            df_banco_atualizado.to_excel(writer, sheet_name='Banco', index=False)
            
            # Em seguida, salva as outras abas
            for sheet_name, df_sheet in all_sheets.items():
                if sheet_name != 'Banco':
                    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        messagebox.showinfo("Sucesso", 
                            "✅ Automação concluída!\n\n"
                            f"Total de novos lançamentos de débito adicionados: {len(novos_lancamentos_sem_duplicatas)}\n"
                            f"O arquivo 'Automação_Gransoft.xlsx' foi atualizado com sucesso.")
        
        return True

    except Exception as e:
        messagebox.showerror("Erro", f"❌ Ocorreu um erro inesperado:\n\n{e}")
        return False

# Executa o processo principal
if __name__ == "__main__":
    processar_extrato_e_transferir()