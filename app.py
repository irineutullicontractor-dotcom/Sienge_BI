import streamlit as st
import pandas as pd
import io
import re

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


# =========================
# 🔹 DICIONÁRIO DE RELATÓRIOS
# =========================

relatorios = {
    "Financeiro": relatorio_financeiro,
    "Apropriação de Obra": relatorio_apropriacao,
    "Bens Sintético": relatorio_bens,
    "Diário Equipamento": relatorio_diario,
    "Equipamento Analítico": relatorio_eq_analitico,
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
