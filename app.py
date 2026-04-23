import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Central de Relatórios", layout="wide")
st.title("📊 Central de Relatórios")

# =========================================================
# 🔹 FUNÇÃO DIÁRIO EQUIPAMENTO (ÍNTEGRA DO TXT)
# =========================================================

def relatorio_diario_eq(file):
    # Carregar exatamente como no TXT
    df = pd.read_excel(file, header=None)

    centro_custo_atual = None
    n_registro_atual = None
    equipamento_atual = None
    placa_atual = None
    responsavel_atual = None
    observacao_atual = None

    header_row_index = None
    dados_reestruturados = []

    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    # Índices base conforme o TXT
    col_idx_numero = 0
    col_idx_obra = 1
    col_idx_utilizacao = 4
    col_idx_operador = 7
    col_idx_data_saida = 9
    col_idx_data_chegada = 14
    
    col_idx_hodometro_saida = None
    col_idx_hodometro_chegada = None
    col_idx_horimetro_saida = None
    col_idx_horimetro_chegada = None

    for index, row in df.iterrows():
        # Stop processing if date/time pattern is matched
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])):
            break

        # Extração de Metadados do Bloco (Exatamente como no TXT)
        if isinstance(row[0], str):
            if 'Centro de custo' in row[0]:
                centro_custo_atual = row[3]
            elif 'Nº registro' in row[0]:
                n_registro_atual = row[3]
            elif 'Equipamento' in row[0]:
                equipamento_atual = row[3]
            elif 'Placa/Plaqueta' in row[0]:
                placa_atual = row[3]
            elif 'Responsável' in row[0]:
                responsavel_atual = row[3]
            elif 'Observação' in row[0]:
                observacao_atual = row[3]

        # Identificação do Cabeçalho e Colunas Dinâmicas
        # Busca por 'Hodômetro' ou 'Horímetro' na linha para definir os índices
        row_values_str = [str(v) for v in row.values]
        if any('Hodômetro' in v or 'Horímetro' in v for v in row_values_str):
            header_row_index = index
            
            # Resetar índices dinâmicos para o novo bloco
            col_idx_hodometro_saida = None
            col_idx_horimetro_saida = None
            col_idx_hodometro_chegada = None
            col_idx_horimetro_chegada = None

            for i, val in enumerate(row_values_str):
                if 'Hodômetro' in val and 'saída' in val: col_idx_hodometro_saida = i
                if 'Horímetro' in val and 'saída' in val: col_idx_horimetro_saida = i
                if 'Hodômetro' in val and 'chegada' in val: col_idx_hodometro_chegada = i
                if 'Horímetro' in val and 'chegada' in val: col_idx_horimetro_chegada = i
            continue

        # Processamento das linhas de dados
        if header_row_index is not None and index > header_row_index:
            # Se a linha for vazia ou for um novo cabeçalho/metadado, pula
            if pd.isna(row[0]) or row[0] == 'Número' or 'Centro de custo' in str(row[0]):
                continue
            
            # Se encontrar o rodapé do bloco (Total), pode resetar o header_row_index
            if 'Total' in str(row[0]):
                header_row_index = None
                continue

            # Extração dos valores conforme os índices mapeados
            dados_reestruturados.append({
                'Centro de custo': centro_custo_atual,
                'Nº registro': n_registro_atual,
                'Equipamento': equipamento_atual,
                'Placa/Plaqueta': placa_atual,
                'Responsável': responsavel_atual,
                'Observação': observacao_atual,
                'Número': row[col_idx_numero],
                'Obra': row[col_idx_obra],
                'Utilização': row[col_idx_utilizacao],
                'Operador': row[col_idx_operador],
                'Data saída': row[col_idx_data_saida],
                'Hodômetro saída': row[col_idx_hodometro_saida] if col_idx_hodometro_saida is not None else None,
                'Horímetro saída': row[col_idx_horimetro_saida] if col_idx_horimetro_saida is not None else None,
                'Data chegada': row[col_idx_data_chegada],
                'Hodômetro chegada': row[col_idx_hodometro_chegada] if col_idx_hodometro_chegada is not None else None,
                'Horímetro chegada': row[col_idx_horimetro_chegada] if col_idx_horimetro_chegada is not None else None,
            })

    return pd.DataFrame(dados_reestruturados)

# =========================================================
# 🔹 INTERFACE E EXECUÇÃO
# =========================================================

tipo = st.selectbox("Selecione o relatório", ["Diário Equipamento"]) # Adicione os outros aqui
arquivo = st.file_uploader("Envie o arquivo original", type=["xlsx"])

if st.button("🚀 Gerar"):
    if arquivo:
        try:
            if tipo == "Diário Equipamento":
                df_result = relatorio_diario_eq(arquivo)
            
            st.success("Processado!")
            st.dataframe(df_result)
            
            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False)
            st.download_button("📥 Baixar Excel", output.getvalue(), "resultado.xlsx")
            
        except Exception as e:
            st.error(f"Erro: {e}")
