import streamlit as st
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="Central de Relatórios", layout="wide")

st.title("📊 Central de Relatórios")

# =========================================================
# 🔹 FUNÇÃO DIÁRIO EQUIPAMENTO - ÍNTEGRA DO TXT
# =========================================================

def relatorio_diario_eq(file):
    # Load the data without skipping rows initially to find the header rows
    df = pd.read_excel(file, header=None)

    centro_custo_atual = None
    n_registro_atual = None
    equipamento_atual = None
    placa_atual = None
    responsavel_atual = None
    observacao_atual = None

    # Store the index of the data header row for the current block
    header_row_index = None

    dados_reestruturados = []

    # Regex pattern to match "DD/MM/YYYY - HH:MM:SS"
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    # Variables to store column indices for the current block
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

    # Loop through rows to find and process each block
    for index, row in df.iterrows():
        # Check if the first column matches the date and time pattern to stop processing
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])):
            break

        # Find rows containing the metadata for the current block
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

        # Find the data header row for the current block
        # We look for a row that contains "Número" and "Obra" in its expected positions
        # or more robustly, check for "Hodômetro" or "Horímetro" keywords in the row
        row_values_str = [str(v) for v in row.values]
        if any('Hodômetro' in v or 'Horímetro' in v for v in row_values_str):
            header_row_index = index
            
            # Reset column indices for Hodômetro and Horímetro for each block
            col_idx_hodometro_saida = None
            col_idx_horimetro_saida = None
            col_idx_hodometro_chegada = None
            col_idx_horimetro_chegada = None

            # Dynamically identify the columns for Hodômetro and Horímetro
            for i, val in enumerate(row_values_str):
                if 'Hodômetro' in val and 'saída' in val:
                    col_idx_hodometro_saida = i
                elif 'Horímetro' in val and 'saída' in val:
                    col_idx_horimetro_saida = i
                elif 'Hodômetro' in val and 'chegada' in val:
                    col_idx_hodometro_chegada = i
                elif 'Horímetro' in val and 'chegada' in val:
                    col_idx_horimetro_chegada = i
            continue

        # Process the data rows after the header row has been found
        if header_row_index is not None and index > header_row_index:
            # Check if we've reached a row that is not a data row (e.g., starts a new block or is empty)
            if pd.isna(row[0]) or row[0] == 'Número' or 'Centro de custo' in str(row[0]):
                # If we hit an empty row or a new block's start, stop processing for the current block
                # but don't break the main loop as there might be more blocks
                # Reset header_row_index only if it's truly the end of a block's data
                if pd.isna(row[0]) and index > header_row_index + 1:
                     header_row_index = None
                continue

            # Extract the data using the identified column indices
            numero = row[col_idx_numero]
            obra = row[col_idx_obra]
            utilizacao = row[col_idx_utilizacao]
            operador = row[col_idx_operador]
            data_saida = row[col_idx_data_saida]
            data_chegada = row[col_idx_data_chegada]

            # Access Hodômetro and Horímetro values using the dynamic indices
            hodometro_saida = row[col_idx_hodometro_saida] if col_idx_hodometro_saida is not None else None
            horimetro_saida = row[col_idx_horimetro_saida] if col_idx_horimetro_saida is not None else None
            hodometro_chegada = row[col_idx_hodometro_chegada] if col_idx_hodometro_chegada is not None else None
            horimetro_chegada = row[col_idx_horimetro_chegada] if col_idx_horimetro_chegada is not None else None

            # Store the extracted and organized data
            dados_reestruturados.append({
                'Centro de custo': centro_custo_atual,
                'Nº registro': n_registro_atual,
                'Equipamento': equipamento_atual,
                'Placa/Plaqueta': placa_atual,
                'Responsável': responsavel_atual,
                'Observação': observacao_atual,
                'Número': numero,
                'Obra': obra,
                'Utilização': utilizacao,
                'Operador': operador,
                'Data saída': data_saida,
                'Hodômetro saída': hodometro_saida,
                'Horímetro saída': horimetro_saida,
                'Data chegada': data_chegada,
                'Hodômetro chegada': hodometro_chegada,
                'Horímetro chegada': horimetro_chegada,
            })

    # Convert to DataFrame
    df_reestruturado = pd.DataFrame(dados_reestruturados)

    # Define the desired column order (Integral do TXT)
    column_order = [
         'Centro de custo', 'Nº registro', 'Equipamento', 'Placa/Plaqueta',
         'Responsável', 'Observação', 'Número', 'Obra', 'Utilização',
         'Operador', 'Data saída', 'Hodômetro saída', 'Horímetro saída',
         'Data chegada', 'Hodômetro chegada', 'Horímetro chegada'
    ]

    # Reorder the columns, dropping columns that don't exist
    df_reestruturado = df_reestruturado.reindex(columns=[c for c in column_order if c in df_reestruturado.columns])
    
    return df_reestruturado

# =========================
# 🔹 INTERFACE
# =========================

tipo = st.selectbox("Selecione o relatório", ["Diário Equipamento"])

arquivo = st.file_uploader("Anexe o arquivo Excel original", type=["xlsx"])

if st.button("🚀 Gerar Relatório"):
    if not arquivo:
        st.warning("Por favor, anexe um arquivo.")
    else:
        try:
            if tipo == "Diário Equipamento":
                df_resultado = relatorio_diario_eq(arquivo)
                
                st.success("Processamento concluído!")
                
                # Botão de Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Baixar Excel Processado",
                    data=output.getvalue(),
                    file_name="relatorio_diario_reestruturado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(df_resultado)
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
