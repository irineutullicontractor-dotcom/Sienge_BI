import streamlit as st
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="Central de Relatórios", layout="wide")

st.title("📊 Central de Relatórios")

st.markdown("Selecione o tipo de relatório, envie o arquivo e gere o resultado.")

# =========================
# 🔹 FUNÇÕES DE RELATÓRIOS
# =========================

def relatorio_financeiro(file):
    df = pd.read_excel(file, header=None)

    header_row_index = None
    dados = []

    for index, row in df.iterrows():

        if row[0] == 'Emissão':
            header_row_index = index

        if row[0] == 'Total do período':
            break

        if header_row_index is not None and index > header_row_index:
            dados.append({
                'Emissão': row[0],
                'Vencto': row[1],
                'Cliente': row[3],
                'Título': row[5],
                'Documento': row[8],
                'Plano': row[10],
                'Crédito': row[13],
                'Débito': row[17],
            })

    return pd.DataFrame(dados)


def relatorio_apropriacao(file):
    df = pd.read_excel(file, header=None)

    dados = []
    periodo = None
    obra = None

    for i, row in df.iterrows():

        if row[0] == 'Período':
            periodo = row[4]

        if row[0] == 'Obra':
            obra = row[4]

        if row[0] == 'Data':
            header = i
            continue

        if 'header' in locals() and i > header:
            dados.append({
                'Período': periodo,
                'Obra': obra,
                'Data': row[0],
                'Documento': row[1],
                'Valor': row[12],
            })

    return pd.DataFrame(dados)


def relatorio_bens(file):
    df = pd.read_excel(file, header=None)

    dados = []
    centro = None

    for i, row in df.iterrows():

        if row[0] == 'Centro de custo':
            centro = row[3]

        if row[0] == 'Patrimônio':
            header = i
            continue

        if 'header' in locals() and i > header:
            if pd.notna(row[0]):
                dados.append({
                    'Centro': centro,
                    'Patrimônio': row[0],
                    'Descrição': row[4],
                    'Situação': row[9],
                })

    return pd.DataFrame(dados)


def relatorio_diario(file):
    df = pd.read_excel(file, header=None)

    dados = []
    centro = None

    for i, row in df.iterrows():

        if isinstance(row[0], str) and "Centro de custo" in row[0]:
            centro = row[3]

        if isinstance(row[0], str) and "Número" in row.values:
            header = i
            continue

        if 'header' in locals() and i > header:
            if pd.notna(row[0]):
                dados.append({
                    'Centro': centro,
                    'Número': row[0],
                    'Obra': row[1],
                    'Operador': row[7],
                })

    return pd.DataFrame(dados)


def relatorio_eq_analitico(file):
    df = pd.read_excel(file, header=None)

    dados = []
    equipamento = None

    for i, row in df.iterrows():

        if 'Equipamento' in row.values:
            equipamento = row[2]

        if 'Centro de custo' in row.values and equipamento:
            dados.append({
                'Equipamento': equipamento,
                'Centro': row[2],
            })

    return pd.DataFrame(dados)


# 🔥 RELATÓRIO COMPLETO (SEU SCRIPT FIEL)
def relatorio_diario_eq_completo(file):

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

        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])):
            break

        if isinstance(row[0], str):

            if 'Centro de custo' in row[0]:
                centro_custo_atual = row[3]
                header_row_index = None
                col_idx_hodometro_saida = None
                col_idx_hodometro_chegada = None
                col_idx_horimetro_saida = None
                col_idx_horimetro_chegada = None

            elif 'Nº registro' in row[0]:
                n_registro_atual = row[3]

            elif 'Equipamento' in row[0]:
                equipamento_atual = row[3]

                for col_index, col_value in enumerate(row):
                    if isinstance(col_value, str) and 'Placa/Plaqueta' in col_value:
                        if col_index + 3 < len(row):
                            placa_atual = row[col_index + 3]
                        break

            elif 'Responsável' in row[0]:
                responsavel_atual = row[3]

            elif 'Observação' in row[0]:
                observacao_atual = row[3]

        if header_row_index is None and any(
            isinstance(col_value, str) and ('Hodômetro' in col_value or 'Horímetro' in col_value)
            for col_value in row
        ):

            header_row_index = index

            col_idx_hodometro_temp = []
            col_idx_horimetro_temp = []

            for col_index, col_value in enumerate(row):

                if isinstance(col_value, str):

                    if 'Número' in col_value:
                        col_idx_numero = col_index
                    elif 'Obra' in col_value:
                        col_idx_obra = col_index
                    elif 'Utilização' in col_value:
                        col_idx_utilizacao = col_index
                    elif 'Operador' in col_value:
                        col_idx_operador = col_index
                    elif 'Data saída' in col_value:
                        col_idx_data_saida = col_index
                    elif 'Data chegada' in col_value:
                        col_idx_data_chegada = col_index
                    elif 'Hodômetro' in col_value:
                        col_idx_hodometro_temp.append(col_index)
                    elif 'Horímetro' in col_value:
                        col_idx_horimetro_temp.append(col_index)

            col_idx_hodometro_saida = col_idx_hodometro_temp[0] if len(col_idx_hodometro_temp) >= 1 else None
            col_idx_hodometro_chegada = col_idx_hodometro_temp[1] if len(col_idx_hodometro_temp) >= 2 else None

            col_idx_horimetro_saida = col_idx_horimetro_temp[0] if len(col_idx_horimetro_temp) >= 1 else None
            col_idx_horimetro_chegada = col_idx_horimetro_temp[1] if len(col_idx_horimetro_temp) >= 2 else None

        if header_row_index is not None and index > header_row_index:

            if pd.notna(row[col_idx_numero]) and (
                isinstance(row[col_idx_numero], pd.Timestamp)
                or (isinstance(row[col_idx_numero], str) and '/' in str(row[col_idx_numero]) and 'Total' not in str(row[col_idx_numero]))
                or isinstance(row[col_idx_numero], (int, float))
            ):

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

    df_reestruturado = pd.DataFrame(dados_reestruturados)

    column_order = [
        'Centro de custo','Nº registro','Equipamento','Placa/Plaqueta',
        'Responsável','Observação','Número','Obra','Utilização','Operador',
        'Data saída','Hodômetro saída','Horímetro saída',
        'Data chegada','Hodômetro chegada','Horímetro chegada'
    ]

    return df_reestruturado.reindex(columns=column_order)


# =========================
# 🔹 DICIONÁRIO
# =========================

relatorios = {
    "Financeiro": relatorio_financeiro,
    "Apropriação de Obra": relatorio_apropriacao,
    "Bens Sintético": relatorio_bens,
    "Diário Equipamento (Simples)": relatorio_diario,
    "Equipamento Analítico": relatorio_eq_analitico,
    "🔥 Diário Equipamento COMPLETO": relatorio_diario_eq_completo,
}

# =========================
# 🔹 INTERFACE
# =========================

tipo = st.selectbox("Selecione o relatório", list(relatorios.keys()))

arquivo = st.file_uploader("Anexe o arquivo", type=["xlsx", "xls"])

if st.button("🚀 Gerar Relatório"):

    if not arquivo:
        st.warning("Envie um arquivo primeiro.")
    else:
        try:
            func = relatorios[tipo]
            df_resultado = func(arquivo)

            output = io.BytesIO()
            df_resultado.to_excel(output, index=False)

            st.success("Relatório gerado com sucesso!")

            st.download_button(
                "📥 Baixar",
                output.getvalue(),
                f"{tipo}.xlsx"
            )

            st.dataframe(df_resultado)

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
