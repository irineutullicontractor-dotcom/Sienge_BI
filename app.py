import streamlit as st
import pandas as pd
import re
import io

# Configuração da página
st.set_page_config(page_title="Central de Relatórios", layout="wide")

st.title("📊 Central de Processamento de Relatórios")

# --- LISTA SUSPENSA ---
tipo_relatorio = st.selectbox(
    "Selecione o relatório para processar:",
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

if st.button("🚀 Gerar Relatório"):
    if uploaded_file is not None:
        try:
            # Lógica extraída exatamente dos arquivos TXT
            output = io.BytesIO()
            date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')
            
            # --- 1. MAPA DE CONTROLE (1 OBRA) ---
            if tipo_relatorio == "Mapa de Controle (1 Obra)":
                df = pd.read_excel(uploaded_file, header=6)
                colunas_para_remover = [2, 4, 7, 9, 15, 17]
                df.drop(df.columns[colunas_para_remover], axis=1, inplace=True)
                df_final = df.dropna(subset=['Item'])

            # --- 2. MAPA DE CONTROLE (MÚLTIPLAS OBRAS) ---
            elif tipo_relatorio == "Mapa de Controle (Múltiplas Obras)":
                df_temp = pd.read_excel(uploaded_file, header=None)
                header_row_index = None
                for i in range(len(df_temp)):
                    if pd.notna(df_temp.iloc[i, 0]) and 'item' in str(df_temp.iloc[i, 0]).lower():
                        header_row_index = i
                        break
                df = pd.read_excel(uploaded_file, header=header_row_index)
                colunas_indices = [2, 4, 7, 9, 15, 17]
                current_columns = df.columns.tolist()
                cols_to_drop = [current_columns[i] for i in colunas_indices if i < len(current_columns)]
                df.drop(columns=cols_to_drop, axis=1, inplace=True, errors='ignore')
                df_final = df.dropna(subset=[df.columns[0]])

            # --- 3. APROPRIAÇÃO DE OBRA ---
            elif tipo_relatorio == "Apropriação de Obra":
                df = pd.read_excel(uploaded_file, header=None)
                periodo_atual, selecao_atual, obra_atual, unidade_atual, celula_atual, etapa_atual, subetapa_atual = [None]*7
                header_row_index = None
                dados_reestruturados = []
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
                    if header_row_index and index > header_row_index:
                        dados_reestruturados.append({
                            'Período': periodo_atual, 'Seleção por': selecao_atual, 'Obra': obra_atual, 'Unidade construtiva': unidade_atual,
                            'Célula construtiva': celula_atual, 'Etapa': etapa_atual, 'Subetapa': subetapa_atual, 'Data': row[0],
                            'Documento': row[2], 'Título/Parcela': row[5], 'Or': row[8], 'Credor / Histórico': row[10],
                            'Valor do documento': row[14], 'Valor apropriado': row[17]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- 4. BENS SINTÉTICO ---
            elif tipo_relatorio == "Bens Sintético":
                df = pd.read_excel(uploaded_file, header=None)
                centro_custo_atual, grupo_atual, header_row_index = None, None, None
                dados_reestruturados = []
                for index, row in df.iterrows():
                    if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
                    if row[0] == 'Centro de custo': centro_custo_atual = row[3]
                    if row[0] == 'Grupo': grupo_atual = row[3]
                    if row[0] == 'Patrimônio': header_row_index = index; continue
                    if header_row_index and index > header_row_index and pd.notna(row[0]):
                        dados_reestruturados.append({
                            'Centro de custo': centro_custo_atual, 'Grupo': grupo_atual, 'Patrimônio': row[0],
                            'Placa/Plaqueta': row[2], 'Cód barras': row[4], 'Descrição': row[5], 'Conservação': row[9],
                            'Dt. Incorporação': row[10], 'Situação': row[11], 'Localização atual': row[13]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- 5. DIÁRIO DE EQUIPAMENTOS ---
            elif tipo_relatorio == "Diário de Equipamentos":
                df = pd.read_excel(uploaded_file, header=None)
                c_custo, n_reg, eqp, placa, resp, obs, h_idx, dados = [None]*7, []
                for index, row in df.iterrows():
                    if row[0] == 'Centro de custo': c_custo = row[2]
                    if row[0] == 'Nº registro': n_reg = row[2]
                    if row[0] == 'Equipamento': eqp = row[2]
                    if row[4] == 'Placa/Plaqueta': placa = row[5]
                    if row[0] == 'Responsável': resp = row[2]
                    if row[0] == 'Observação': obs = row[2]
                    if row[0] == 'Número': h_idx = index; continue
                    if h_idx and index > h_idx and pd.notna(row[0]):
                        if isinstance(row[0], str) and 'Total' in row[0]: continue
                        dados.append({
                            'Centro de custo': c_custo, 'Nº registro': n_reg, 'Equipamento': eqp, 'Placa/Plaqueta': placa,
                            'Responsável': resp, 'Observação': obs, 'Número': row[0], 'Obra': row[1], 'Utilização': row[4],
                            'Operador': row[7], 'Data saída': row[9], 'Data chegada': row[14]
                        })
                df_final = pd.DataFrame(dados)

            # --- 7. FINANCEIRO ---
            elif tipo_relatorio == "Financeiro":
                df = pd.read_excel(uploaded_file, header=None)
                h_idx, dados = None, []
                for index, row in df.iterrows():
                    if row[0] == 'Emissão': h_idx = index; continue
                    if row[0] == 'Total do período': break
                    if h_idx and index > h_idx and pd.notna(row[0]):
                        dados.append({
                            'Emissão': row[0], 'Vencto': row[1], 'Cliente/Fornecedor/Complemento': row[3], 'Título/Parcela': row[5],
                            'Documento': row[8], 'Plano financeiro': row[10], 'Crédito': row[13], 'Débito': row[17]
                        })
                df_final = pd.DataFrame(dados)

            # --- 8. HISTÓRICO DE BENS (ORIGEM/DESTINO) ---
            elif tipo_relatorio == "Histórico de Bens (Origem/Destino)":
                df = pd.read_excel(uploaded_file, header=None)
                pat_at, pla_at, cod_at, det_at, h_idx, dados = [None]*5, []
                for index, row in df.iterrows():
                    if row[0] == 'Patrimônio': pat_at = row[3]
                    if row[6] == 'Placa/Plaqueta': pla_at = row[7]
                    if row[0] == 'Detalhamento': det_at = row[3]
                    if row[0] == 'Data': h_idx = index; continue
                    if h_idx and index > h_idx and pd.notna(row[0]):
                        # Lógica exata de split de Origem/Destino do txt
                        cc_raw = str(row[4])
                        origem, destino = "", ""
                        if "Origem:" in cc_raw and "Destino:" in cc_raw:
                            parts = cc_raw.split("Destino:")
                            origem = parts[0].replace("Origem:", "").strip()
                            destino = parts[1].strip()
                        setor_raw = str(row[8])
                        dados.append({
                            'Patrimônio': pat_at, 'Placa/Plaqueta': pla_at, 'Detalhamento': det_at, 'Data': row[0],
                            'Tipo do movimento': row[1], 'Movimento': row[3], 'Centro(s) de Custo': row[4],
                            'Setor/obra origem': setor_raw, 'Responsável': row[11]
                        })
                df_final = pd.DataFrame(dados)

            # --- EXPORTAÇÃO ---
            if not df_final.empty:
                df_final.to_excel(output, index=False)
                st.success("✅ Processado com sucesso!")
                st.download_button("📥 Baixar Relatório", output.getvalue(), f"{tipo_relatorio}.xlsx")
            else:
                st.error("⚠️ O processamento não gerou dados. Verifique o arquivo de entrada.")

        except Exception as e:
            st.error(f"❌ Erro crítico: {e}")
