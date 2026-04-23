import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Central de Relatórios", layout="wide")
st.title("📊 Central de Relatórios (Lógica Original)")

tipo_relatorio = st.selectbox(
    "Selecione o relatório:",
    [
        "Mapa de Controle (1 Obra)",
        "Mapa de Controle (Múltiplas Obras)",
        "Apropriação de Obra",
        "Bens Sintético",
        "Diário de Equipamentos",
        "Equipamento Analítico",
        "Financeiro",
        "Histórico de Bens (Origem/Destino)"
    ]
)

uploaded_file = st.file_uploader("Anexe o arquivo Excel (.xlsx)", type=['xlsx'])

if st.button("🚀 Rodar Processamento"):
    if uploaded_file is not None:
        try:
            # Lemos o arquivo para memória uma vez
            # Para os códigos funcionarem, usaremos 'df' como o dataframe inicial
            
            output = io.BytesIO()

            # --- CÓDIGO: MAPA DE CONTROLE (1 OBRA) ---
            if tipo_relatorio == "Mapa de Controle (1 Obra)":
                df = pd.read_excel(uploaded_file, header=6)
                colunas_para_remover = [2, 4, 7, 9, 15, 17]
                df.drop(df.columns[colunas_para_remover], axis=1, inplace=True)
                df_final = df.dropna(subset=['Item'])

            # --- CÓDIGO: MAPA DE CONTROLE (MÚLTIPLAS OBRAS) ---
            elif tipo_relatorio == "Mapa de Controle (Múltiplas Obras)":
                df_temp = pd.read_excel(uploaded_file, header=None)
                header_row_index = None
                for i in range(len(df_temp)):
                    if pd.notna(df_temp.iloc[i, 0]) and 'item' in str(df_temp.iloc[i, 0]).lower():
                        header_row_index = i
                        break
                df = pd.read_excel(uploaded_file, header=header_row_index)
                colunas_para_remover_indices_originais = [2, 4, 7, 9, 15, 17]
                current_columns = df.columns.tolist()
                columns_to_drop_by_name = []
                for idx in colunas_para_remover_indices_originais:
                    if idx < len(current_columns):
                        columns_to_drop_by_name.append(current_columns[idx])
                df.drop(columns=list(set(columns_to_drop_by_name)), axis=1, inplace=True, errors='ignore')
                df_final = df.dropna(subset=[df.columns[0]])

            # --- CÓDIGO: APROPRIAÇÃO DE OBRA ---
            elif tipo_relatorio == "Apropriação de Obra":
                df = pd.read_excel(uploaded_file, header=None)
                periodo_atual = None
                selecao_atual = None
                obra_atual = None
                unidade_atual = None
                celula_atual = None
                etapa_atual = None
                subetapa_atual = None
                header_row_index = None
                dados_reestruturados = []
                date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')
                for index, row in df.iterrows():
                    if (isinstance(row[0], str) and row[0].strip() in ['Total da etapa', 'Total da subetapa', 'Total da célula construtiva', 'Total da unidade construtiva', 'Total da obra']) or pd.isna(row[0]) or (isinstance(row[0], str) and date_time_pattern.match(row[0].strip())):
                        continue
                    if row[0] == 'Período': periodo_atual = row[4]
                    if row[8] == 'Seleção por': selecao_atual = row[13]
                    if row[0] == 'Obra': obra_atual = row[4]
                    if row[0] == 'Unidade construtiva': unidade_atual = row[4]
                    if row[0] == 'Célula construtiva': celula_atual = row[4]
                    if row[0] == 'Etapa': etapa_atual = row[4]
                    if row[0] == 'Subetapa': subetapa_atual = row[4]
                    if row[0] == 'Data': header_row_index = index; continue
                    if header_row_index is not None and index > header_row_index:
                        dados_reestruturados.append({
                            'Período': periodo_atual, 'Seleção por': selecao_atual, 'Obra': obra_atual, 'Unidade construtiva': unidade_atual,
                            'Célula construtiva': celula_atual, 'Etapa': etapa_atual, 'Subetapa': subetapa_atual, 'Data': row[0],
                            'Documento': row[2], 'Título/Parcela': row[5], 'Or': row[8], 'Credor / Histórico': row[10],
                            'Valor do documento': row[14], 'Valor apropriado': row[17]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- CÓDIGO: BENS SINTÉTICO ---
            elif tipo_relatorio == "Bens Sintético":
                df = pd.read_excel(uploaded_file, header=None)
                centro_custo_atual, grupo_atual, header_row_index = None, None, None
                dados_reestruturados = []
                date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')
                for index, row in df.iterrows():
                    if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
                    if row[0] == 'Centro de custo': centro_custo_atual = row[3]
                    if row[0] == 'Grupo': grupo_atual = row[3]
                    if row[0] == 'Patrimônio': header_row_index = index; continue
                    if header_row_index is not None and index > header_row_index and pd.notna(row[0]):
                        dados_reestruturados.append({
                            'Centro de custo': centro_custo_atual, 'Grupo': grupo_atual, 'Patrimônio': row[0],
                            'Placa/Plaqueta': row[2], 'Cód barras': row[4], 'Descrição': row[5], 'Conservação': row[9],
                            'Dt. Incorporação': row[10], 'Situação': row[11], 'Localização atual': row[13]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- CÓDIGO: DIÁRIO DE EQUIPAMENTOS ---
            elif tipo_relatorio == "Diário de Equipamentos":
                df = pd.read_excel(uploaded_file, header=None)
                centro_custo_atual, n_registro_atual, equipamento_atual, placa_atual, responsavel_atual, observacao_atual = [None]*6
                header_row_index = None
                dados_reestruturados = []
                for index, row in df.iterrows():
                    if row[0] == 'Centro de custo': centro_custo_atual = row[2]
                    if row[0] == 'Nº registro': n_registro_atual = row[2]
                    if row[0] == 'Equipamento': equipamento_atual = row[2]
                    if row[4] == 'Placa/Plaqueta': placa_atual = row[5]
                    if row[0] == 'Responsável': responsavel_atual = row[2]
                    if row[0] == 'Observação': observacao_atual = row[2]
                    if row[0] == 'Número': header_row_index = index; continue
                    if header_row_index is not None and index > header_row_index and pd.notna(row[0]):
                        if isinstance(row[0], str) and 'Total' in row[0]: continue
                        dados_reestruturados.append({
                            'Centro de custo': centro_custo_atual, 'Nº registro': n_registro_atual, 'Equipamento': equipamento_atual, 'Placa/Plaqueta': placa_atual,
                            'Responsável': responsavel_atual, 'Observação': observacao_atual, 'Número': row[0], 'Obra': row[1], 'Utilização': row[4],
                            'Operador': row[7], 'Data saída': row[9], 'Data chegada': row[14]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- CÓDIGO: FINANCEIRO ---
            elif tipo_relatorio == "Financeiro":
                df = pd.read_excel(uploaded_file, header=None)
                header_row_index, dados_reestruturados = None, []
                for index, row in df.iterrows():
                    if row[0] == 'Emissão':
                        header_row_index = index
                        idx_em, idx_ve, idx_cl, idx_ti, idx_do, idx_pl, idx_cr, idx_de = 0, 1, 3, 5, 8, 10, 13, 17
                    if row[0] == 'Total do período': break
                    if header_row_index is not None and index > header_row_index and pd.notna(row[0]):
                        dados_reestruturados.append({
                            'Emissão': row[idx_em], 'Vencto': row[idx_ve], 'Cliente/Fornecedor/Complemento': row[idx_cl],
                            'Título/Parcela': row[idx_ti], 'Documento': row[idx_do], 'Plano financeiro': row[idx_pl],
                            'Crédito': row[idx_cr], 'Débito': row[idx_de]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- CÓDIGO: HISTÓRICO DE BENS (ORIGEM/DESTINO) ---
            # COPIADO EXATAMENTE DO SEU TXT "formula historico_bens orig_dest 0.1.txt"
            elif tipo_relatorio == "Histórico de Bens (Origem/Destino)":
                df = pd.read_excel(uploaded_file, header=None)
                patrimonio_atual, placa_atual, detalhamento_atual, header_row_index = None, None, None, None
                dados_reestruturados = []
                date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')
                for index, row in df.iterrows():
                    if row[0] == 'Patrimônio': patrimonio_atual = row[3]
                    if row[6] == 'Placa/Plaqueta': placa_atual = row[7]
                    if row[0] == 'Detalhamento': detalhamento_atual = row[3]
                    if row[0] == 'Data': header_row_index = index; continue
                    if header_row_index is not None and index > header_row_index:
                        if pd.isna(row[0]) or (isinstance(row[0], str) and date_time_pattern.match(row[0].strip())): continue
                        cc_raw = str(row[4])
                        setor_origem, setor_destino = "", ""
                        if "Origem:" in cc_raw and "Destino:" in cc_raw:
                            parts = cc_raw.split("Destino:")
                            setor_origem = parts[0].replace("Origem:", "").strip()
                            setor_destino = parts[1].strip()
                        elif "Destino:" in cc_raw:
                            setor_destino = cc_raw.replace("Destino:", "").strip()
                        dados_reestruturados.append({
                            'Patrimônio': patrimonio_atual, 'Placa/Plaqueta': placa_atual, 'Detalhamento': detalhamento_atual,
                            'Data': row[0], 'Tipo do movimento': row[1], 'Movimento': row[3], 'Centro(s) de Custo': row[4],
                            'Setor/obra origem': setor_origem, 'Setor/obra destino': setor_destino, 'Responsável': row[11]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- EXPORTAÇÃO ---
            if not df_final.empty:
                df_final.to_excel(output, index=False)
                st.success("✅ Processado com sua lógica original!")
                st.download_button("📥 Baixar Arquivo", output.getvalue(), f"{tipo_relatorio}.xlsx")
            
        except Exception as e:
            st.error(f"Erro: {e}")
