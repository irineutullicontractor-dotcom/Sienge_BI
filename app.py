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
            output = io.BytesIO()
            df_final = pd.DataFrame()
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
                periodo_atual = None
                selecao_atual = None
                obra_atual = None
                unidade_atual = None
                celula_atual = None
                etapa_atual = None
                subetapa_atual = None
                header_row_index = None
                dados_reestruturados = []
                for index, row in df.iterrows():
                    val0 = str(row[0]).strip() if pd.notna(row[0]) else ""
                    if val0 in ['Total da etapa', 'Total da subetapa', 'Total da célula construtiva', 'Total da unidade construtiva', 'Total da obra'] or date_time_pattern.match(val0):
                        continue
                    if val0 == 'Período': periodo_atual = row[4]
                    if str(row[8]) == 'Seleção por': selecao_atual = row[13]
                    if val0 == 'Obra': obra_atual = row[4]
                    if val0 == 'Unidade construtiva': unidade_atual = row[4]
                    if val0 == 'Célula construtiva': celula_atual = row[4]
                    if val0 == 'Etapa': etapa_atual = row[4]
                    if val0 == 'Subetapa': subetapa_atual = row[4]
                    if val0 == 'Data': header_row_index = index; continue
                    if header_row_index and index > header_row_index and pd.notna(row[0]):
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
                centro_custo_atual = None
                grupo_atual = None
                header_row_index = None
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
                centro_custo_atual = None
                n_registro_atual = None
                equipamento_atual = None
                placa_atual = None
                responsavel_atual = None
                observacao_atual = None
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
                    if header_row_index and index > header_row_index and pd.notna(row[0]):
                        if 'Total' in str(row[0]): continue
                        dados_reestruturados.append({
                            'Centro de custo': centro_custo_atual, 'Nº registro': n_registro_atual, 'Equipamento': equipamento_atual,
                            'Placa/Plaqueta': placa_atual, 'Responsável': responsavel_atual, 'Observação': observacao_atual,
                            'Número': row[0], 'Obra': row[1], 'Utilização': row[4], 'Operador': row[7], 'Data saída': row[9], 'Data chegada': row[14]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- 7. FINANCEIRO ---
            elif tipo_relatorio == "Financeiro":
                df = pd.read_excel(uploaded_file, header=None)
                header_row_index = None
                dados_reestruturados = []
                for index, row in df.iterrows():
                    if row[0] == 'Emissão': header_row_index = index; continue
                    if row[0] == 'Total do período': break
                    if header_row_index and index > header_row_index and pd.notna(row[0]):
                        dados_reestruturados.append({
                            'Emissão': row[0], 'Vencto': row[1], 'Cliente/Fornecedor/Complemento': row[3], 'Título/Parcela': row[5],
                            'Documento': row[8], 'Plano financeiro': row[10], 'Crédito': row[13], 'Débito': row[17]
                        })
                df_final = pd.DataFrame(dados_reestruturados)

            # --- 8. HISTÓRICO DE BENS (ORIGEM/DESTINO) ---
            elif tipo_relatorio == "Histórico de Bens (Origem/Destino)":
                df = pd.read_excel(uploaded_file, header=None)
                patrimonio_atual = None
                placa_atual = None
                detalhamento_atual = None
                header_row_index = None
                dados_reestruturados = []
                for index, row in df.iterrows():
                    if row[0] == 'Patrimônio': patrimonio_atual = row[3]
                    if row[6] == 'Placa/Plaqueta': placa_atual = row[7]
                    if row[0] == 'Detalhamento': detalhamento_atual = row[3]
                    if row[0] == 'Data': header_row_index = index; continue
                    if header_row_index and index > header_row_index and pd.notna(row[0]):
                        if date_time_pattern.match(str(row[0])): continue
                        
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

            # --- EXPORTAÇÃO FINAL ---
            if not df_final.empty:
                df_final.to_excel(output, index=False)
                st.success(f"✅ Relatório '{tipo_relatorio}' processado com sucesso!")
                st.download_button("📥 Baixar Relatório", output.getvalue(), f"{tipo_relatorio}.xlsx")
            else:
                st.error("⚠️ Nenhuma informação foi extraída. Verifique se o arquivo corresponde ao tipo selecionado.")

        except Exception as e:
            st.error(f"❌ Erro crítico: {e}")
