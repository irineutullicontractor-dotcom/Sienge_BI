import streamlit as st
import pandas as pd
import io
import re

# Configuração da página
st.set_page_config(page_title="Central de Relatórios", layout="wide")

st.title("📊 Processador de Relatórios Customizados")
st.markdown("Selecione o tipo de relatório, anexe o arquivo original e gere a versão tratada.")

# --- LISTA SUSPENSA PARA SELEÇÃO ---
tipo_relatorio = st.selectbox(
    "Qual relatório deseja processar?",
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

# --- ÁREA DE UPLOAD ---
uploaded_file = st.file_uploader("Arraste ou selecione o arquivo Excel (.xlsx)", type=['xlsx'])

# --- FUNÇÕES DE APOIO (Comuns a vários códigos) ---
def extrair_dados_basicos(df, padrao_data):
    # Função auxiliar para verificar se a linha é o fim do relatório (data/hora)
    for index, row in df.iterrows():
        if isinstance(row[0], str) and padrao_data.match(str(row[0])):
            return index
    return len(df)

# --- PROCESSAMENTO ---
if st.button("🚀 Gerar Relatório"):
    if uploaded_file is not None:
        try:
            output = io.BytesIO()
            df_final = pd.DataFrame()
            nome_arquivo_saida = f"{tipo_relatorio.lower().replace(' ', '_')}_tratado.xlsx"
            
            # Regex padrão para data/hora usado em vários arquivos
            date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

            # 1. MAPA DE CONTROLE (1 OBRA)
            if tipo_relatorio == "Mapa de Controle (1 Obra)":
                df = pd.read_excel(uploaded_file, header=6)
                colunas_para_remover = [2, 4, 7, 9, 15, 17]
                df.drop(df.columns[colunas_para_remover], axis=1, inplace=True, errors='ignore')
                df_final = df.dropna(subset=['Item'])

            # 2. MAPA DE CONTROLE (MÚLTIPLAS OBRAS)
            elif tipo_relatorio == "Mapa de Controle (Múltiplas Obras)":
                df_temp = pd.read_excel(uploaded_file, header=None)
                header_row = None
                for i in range(len(df_temp)):
                    if pd.notna(df_temp.iloc[i, 0]) and 'item' in str(df_temp.iloc[i, 0]).lower():
                        header_row = i
                        break
                df = pd.read_excel(uploaded_file, header=header_row)
                col_fantasmas = [2, 4, 7, 9, 15, 17]
                df.drop(df.columns[[c for c in col_fantasmas if c < len(df.columns)]], axis=1, inplace=True)
                df_final = df.dropna(subset=[df.columns[0]])

            # 3. APROPRIAÇÃO DE OBRA
            elif tipo_relatorio == "Apropriação de Obra":
                df = pd.read_excel(uploaded_file, header=None)
                periodo, obra, unidade, celula, etapa, subetapa = [None]*6
                dados = []
                for _, row in df.iterrows():
                    val0 = str(row[0])
                    if 'Período' in val0: periodo = row[4]
                    if row[0] == 'Obra': obra = row[4]
                    if row[0] == 'Unidade construtiva': unidade = row[4]
                    if row[0] == 'Célula construtiva': celula = row[4]
                    if row[0] == 'Etapa': etapa = row[4]
                    if row[0] == 'Subetapa': subetapa = row[4]
                    if isinstance(row[0], (int, float, str)) and re.match(r'\d{2}/\d{2}/\d{4}', str(row[0])):
                        dados.append({
                            'Período': periodo, 'Obra': obra, 'Unidade': unidade, 
                            'Célula': celula, 'Etapa': etapa, 'Subetapa': subetapa,
                            'Data': row[0], 'Documento': row[2], 'Credor': row[6], 'Valor': row[14]
                        })
                df_final = pd.DataFrame(dados)

            # 4. BENS SINTÉTICO
            elif tipo_relatorio == "Bens Sintético":
                df = pd.read_excel(uploaded_file, header=None)
                cc, grupo = None, None
                dados = []
                for _, row in df.iterrows():
                    if row[0] == 'Centro de custo': cc = row[3]
                    if row[0] == 'Grupo': grupo = row[3]
                    if str(row[0]).isdigit(): # Se for código do patrimônio
                        dados.append({
                            'Centro de Custo': cc, 'Grupo': grupo,
                            'Patrimônio': row[0], 'Descrição': row[5], 'Situação': row[11]
                        })
                df_final = pd.DataFrame(dados)

            # 5. DIÁRIO DE EQUIPAMENTOS
            elif tipo_relatorio == "Diário de Equipamentos":
                df = pd.read_excel(uploaded_file, header=None)
                dados = []
                # Exemplo simplificado da lógica do txt
                cc, eqp = None, None
                for _, row in df.iterrows():
                    if row[0] == 'Centro de custo': cc = row[2]
                    if row[0] == 'Equipamento': eqp = row[2]
                    if isinstance(row[0], (int, float)) and not pd.isna(row[0]):
                        dados.append({'Centro Custo': cc, 'Equipamento': eqp, 'Número': row[0], 'Obra': row[1]})
                df_final = pd.DataFrame(dados)

            # 6. EQUIPAMENTO ANALÍTICO
            elif tipo_relatorio == "Equipamento Analítico":
                df = pd.read_excel(uploaded_file, header=None)
                dados = []
                # Segue a lógica de "Centro de custo" e blocos de dados do txt
                df_final = df # Aplicar aqui a lógica de loop do txt correspondente

            # 7. FINANCEIRO
            elif tipo_relatorio == "Financeiro":
                df = pd.read_excel(uploaded_file, header=None)
                dados = []
                header_idx = None
                for idx, row in df.iterrows():
                    if row[0] == 'Emissão': header_idx = idx
                    if header_idx and idx > header_idx and pd.notna(row[0]):
                        if row[0] == 'Total do período': break
                        dados.append({
                            'Emissão': row[0], 'Vencto': row[1], 'Cliente/Fornecedor': row[3],
                            'Documento': row[8], 'Crédito': row[13], 'Débito': row[17]
                        })
                df_final = pd.DataFrame(dados)

            # 8. HISTÓRICO DE BENS
            elif tipo_relatorio == "Histórico de Bens (Origem/Destino)":
                df = pd.read_excel(uploaded_file, header=None)
                pat, placa = None, None
                dados = []
                for _, row in df.iterrows():
                    if row[0] == 'Patrimônio': pat = row[3]
                    if row[6] == 'Placa/Plaqueta': placa = row[7]
                    if isinstance(row[0], (int, float, str)) and re.match(r'\d{2}/\d{2}/\d{4}', str(row[0])):
                        dados.append({'Patrimônio': pat, 'Placa': placa, 'Data': row[0], 'Movimento': row[3]})
                df_final = pd.DataFrame(dados)

            # --- EXPORTAÇÃO FINAL ---
            if not df_final.empty:
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Processado')
                
                st.success(f"✅ Relatório '{tipo_relatorio}' processado com sucesso!")
                st.download_button(
                    label="📥 Baixar Arquivo Tratado",
                    data=output.getvalue(),
                    file_name=nome_arquivo_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nenhum dado foi extraído. Verifique se o formato do arquivo corresponde ao relatório selecionado.")

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
    else:
        st.info("Aguardando arquivo para iniciar...")
