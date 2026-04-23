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
    dados_reestruturados = []

    for index, row in df.iterrows():
        if row[0] == 'Emissão':
            header_row_index = index
            col_idx_emissao, col_idx_vencto, col_idx_cliente = 0, 1, 3
            col_idx_titulo, col_idx_documento, col_idx_plano = 5, 8, 10
            col_idx_credito, col_idx_debito = 13, 17
            continue

        if row[0] == 'Total do período':
            break

        if header_row_index is not None and index > header_row_index:
            if pd.notna(row[col_idx_emissao]):
                dados_reestruturados.append({
                    'Emissão': row[col_idx_emissao],
                    'Vencto': row[col_idx_vencto],
                    'Cliente/Fornecedor/Complemento': row[col_idx_cliente],
                    'Título/Parcela': row[col_idx_titulo],
                    'Documento': row[col_idx_documento],
                    'Plano financeiro': row[col_idx_plano],
                    'Crédito': row[col_idx_credito],
                    'Débito': row[col_idx_debito]
                })
    return pd.DataFrame(dados_reestruturados)

def relatorio_apropriacao(file):
    df = pd.read_excel(file, header=None)
    periodo_atual, obra_atual, unidade_atual, celula_atual = None, None, None, None
    etapa_atual, subetapa_atual, header_row_index = None, None, None
    dados_reestruturados = []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if (isinstance(row[0], str) and row[0].strip() in ['Total da etapa', 'Total da subetapa', 'Total da célula construtiva', 'Total da unidade construtiva', 'Total da obra']) or \
           pd.isna(row[0]) or (isinstance(row[0], str) and date_time_pattern.match(row[0].strip())):
            continue

        if row[0] == 'Período': periodo_atual = row[4]
        if row[0] == 'Obra': obra_atual = row[4]
        if row[0] == 'Unidade construtiva': unidade_atual = row[4]
        if row[0] == 'Célula construtiva': celula_atual = row[4]
        if isinstance(row[0], str) and row[0].strip() == 'Etapa':
            etapa_atual = row[4]
            subetapa_atual = None
            continue
        if isinstance(row[0], str) and row[0].strip() == 'Subetapa':
            subetapa_atual = row[4]
            continue

        if row[0] == 'Data':
            header_row_index = index
            col_idxs = [0, 1, 4, 6, 7, 12, 14] # Data, Doc, Titulo, OR, Credor, Vlr Doc, Vlr Aprop
            continue

        if header_row_index is not None and index > header_row_index:
            if isinstance(row[0], str) and row[0].strip() in ['Período', 'Data']: continue
            dados_reestruturados.append({
                'Período': periodo_atual, 'Obra': obra_atual, 'Unidade construtiva': unidade_atual,
                'Célula construtiva': celula_atual, 'Etapa': etapa_atual, 'Subetapa': subetapa_atual,
                'Data': row[0], 'Documento': row[1], 'Título/Parcela': row[4], 'Credor / Histórico': row[7],
                'Valor do documento': row[12], 'Valor apropriado': row[14]
            })
    return pd.DataFrame(dados_reestruturados)

def relatorio_bens_sintetico(file):
    df = pd.read_excel(file, header=None)
    centro_custo_atual, grupo_atual, header_row_index = None, None, None
    dados_reestruturados = []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
        if row[0] == 'Centro de custo': centro_custo_atual = row[3]
        if row[0] == 'Grupo': grupo_atual = row[3]
        if row[0] == 'Patrimônio':
            header_row_index = index
            continue

        if header_row_index is not None and index > header_row_index:
            if pd.notna(row[0]) and row[0] != 'Patrimônio':
                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual, 'Grupo': grupo_atual,
                    'Patrimônio': row[0], 'Placa/Plaqueta': row[1], 'Descrição': row[4],
                    'Conservação': row[6], 'Dt. Incorporação': row[7], 'Situação': row[9]
                })
    return pd.DataFrame(dados_reestruturados)

def relatorio_diario_eq(file):
    df = pd.read_excel(file, header=None)
    centro_custo_atual, equipamento_atual, placa_atual = None, None, None
    header_row_index, dados_reestruturados = None, []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
        if isinstance(row[0], str):
            if 'Centro de custo' in row[0]: centro_custo_atual = row[3]
            elif 'Equipamento' in row[0]: equipamento_atual = row[3]
            elif 'Placa/Plaqueta' in row[0]: placa_atual = row[3]

        if header_row_index is None and any(isinstance(v, str) and ('Hodômetro' in v or 'Horímetro' in v) for v in row):
            header_row_index = index
            continue

        if header_row_index is not None and index > header_row_index:
            if pd.notna(row[0]):
                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual, 'Equipamento': equipamento_atual,
                    'Placa/Plaqueta': placa_atual, 'Número': row[0], 'Obra': row[1],
                    'Operador': row[7], 'Data saída': row[9], 'Data chegada': row[14]
                })
    return pd.DataFrame(dados_reestruturados)

def relatorio_eq_analitico(file):
    df = pd.read_excel(file, header=None)
    dados_reestruturados = []
    current_data = {}

    for index, row in df.iterrows():
        row_str = row.astype(str).fillna('')
        if 'Centro de custo' in row_str.values:
            if 'Equipamento' in current_data: dados_reestruturados.append(current_data.copy())
            current_data = {'Centro de custo': row_str[2]}
        if 'Equipamento' in row_str.values: current_data['Equipamento'] = row_str[2]
        if 'Placa/Plaqueta' in row_str.values: current_data['Placa/Plaqueta'] = row_str[4]
        if 'Situação' in row_str.values: current_data['Situação'] = row_str[2]

    if 'Equipamento' in current_data: dados_reestruturados.append(current_data)
    return pd.DataFrame(dados_reestruturados)

def relatorio_mapa_controle(file):
    df_temp = pd.read_excel(file, header=None)
    header_idx = None
    for i in range(len(df_temp)):
        if pd.notna(df_temp.iloc[i, 0]) and 'item' in str(df_temp.iloc[i, 0]).lower():
            header_idx = i
            break
    
    df = pd.read_excel(file, header=header_idx)
    colunas_fantasmas = [2, 4, 7, 9, 15, 17]
    df.drop(df.columns[colunas_fantasmas], axis=1, inplace=True, errors='ignore')
    df_final = df.dropna(subset=[df.columns[0]])
    return df_final

def relatorio_historico_bens(file):
    df = pd.read_excel(file, header=None)
    patrimonio_atual, placa_atual, header_row_index = None, None, None
    last_data, last_tipo, dados_reestruturados = None, None, []

    for index, row in df.iterrows():
        if row[0] == 'Patrimônio': patrimonio_atual = row[3]
        if row[6] == 'Placa/Plaqueta': placa_atual = row[7]
        if row[0] == 'Data':
            header_row_index = index
            continue

        if header_row_index is not None and index > header_row_index:
            if isinstance(row[0], str) and row[0].strip() in ['Patrimônio', 'Data']: continue
            data_to_use = row[0] if pd.notna(row[0]) else last_data
            if pd.notna(row[0]): last_data = row[0]
            
            dados_reestruturados.append({
                'Patrimônio': patrimonio_atual, 'Placa/Plaqueta': placa_atual,
                'Data': data_to_use, 'Movimento': row[3], 'Centro de Custo': row[4]
            })
    return pd.DataFrame(dados_reestruturados)

# =========================
# 🔹 DICIONÁRIO DE RELATÓRIOS
# =========================

relatorios = {
    "Financeiro": relatorio_financeiro,
    "Apropriação de Obra": relatorio_apropriacao,
    "Bens Sintético": relatorio_bens_sintetico,
    "Diário Equipamento": relatorio_diario_eq,
    "Equipamento Analítico": relatorio_eq_analitico,
    "Mapa de Controle": relatorio_mapa_controle,
    "Histórico de Bens (Origem/Destino)": relatorio_historico_bens
}

# =========================
# 🔹 INTERFACE
# =========================

tipo = st.selectbox("Selecione o relatório", list(relatorios.keys()))
arquivo = st.file_uploader("Anexe o arquivo Excel (.xlsx)", type=["xlsx"])

if st.button("🚀 Gerar Relatório"):
    if not arquivo:
        st.warning("Envie um arquivo primeiro.")
    else:
        try:
            df_resultado = relatorios[tipo](arquivo)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_resultado.to_excel(writer, index=False)
            
            st.success(f"Relatório '{tipo}' gerado com sucesso!")
            st.download_button("📥 Baixar Resultado", output.getvalue(), f"{tipo.replace(' ', '_')}.xlsx")
            st.dataframe(df_resultado)
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
