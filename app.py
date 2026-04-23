import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Central de Relatórios", layout="wide")

st.title("📊 Central de Relatórios")
st.markdown("Selecione o tipo de relatório, envie o arquivo original e gere o resultado processado.")

# =========================
# 🔹 FUNÇÕES INTEGRAIS (BASEADAS NOS ARQUIVOS TXT)
# =========================

def relatorio_financeiro(file):
    df = pd.read_excel(file, header=None)
    header_row_index = None
    dados_reestruturados = []
    
    for index, row in df.iterrows():
        if row[0] == 'Emissão':
            header_row_index = index
            col_idx_emissao = 0
            col_idx_vencto = 1
            col_idx_cliente = 3
            col_idx_titulo = 5
            col_idx_documento = 8
            col_idx_plano = 10
            col_idx_credito = 13
            col_idx_debito = 17
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
    periodo_atual = selecao_atual = obra_atual = unidade_atual = None
    celula_atual = etapa_atual = subetapa_atual = header_row_index = None
    dados_reestruturados = []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if (isinstance(row[0], str) and row[0].strip() in ['Total da etapa', 'Total da subetapa', 'Total da célula construtiva', 'Total da unidade construtiva', 'Total da obra']) or \
           pd.isna(row[0]) or (isinstance(row[0], str) and date_time_pattern.match(row[0].strip())):
            continue

        if row[0] == 'Período': periodo_atual = row[4]
        if row[8] == 'Seleção por': selecao_atual = row[13]
        if row[0] == 'Obra': obra_atual = row[4]
        if row[0] == 'Unidade construtiva': unidade_atual = row[4]
        if row[0] == 'Célula construtiva': celula_atual = row[4]
        if row[0] == 'Etapa':
            etapa_atual = row[4]
            subetapa_atual = None
            continue
        if row[0] == 'Subetapa':
            subetapa_atual = row[4]
            continue

        if row[0] == 'Data':
            header_row_index = index
            col_idx_data, col_idx_documento, col_idx_titulo = 0, 1, 4
            col_idx_or_, col_idx_credor, col_idx_valor_documento, col_idx_valor_apropriado = 6, 7, 12, 14
            continue

        if header_row_index is not None and index > header_row_index:
            if row[0] == 'Data': continue
            dados_reestruturados.append({
                'Período': periodo_atual, 'Seleção por': selecao_atual, 'Obra': obra_atual,
                'Unidade construtiva': unidade_atual, 'Célula construtiva': celula_atual,
                'Etapa': etapa_atual, 'Subetapa': subetapa_atual, 'Data': row[col_idx_data],
                'Documento': row[col_idx_documento], 'Título/Parcela': row[col_idx_titulo],
                'Or': row[col_idx_or_], 'Credor / Histórico': row[col_idx_credor],
                'Valor do documento': row[col_idx_valor_documento], 'Valor apropriado': row[col_idx_valor_apropriado]
            })
    return pd.DataFrame(dados_reestruturados)

def relatorio_bens_sintetico(file):
    df = pd.read_excel(file, header=None)
    centro_custo_atual = grupo_atual = header_row_index = None
    dados_reestruturados = []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
        if row[0] == 'Centro de custo': centro_custo_atual = row[3]
        if row[0] == 'Grupo': grupo_atual = row[3]
        if row[0] == 'Patrimônio':
            header_row_index = index
            col_idx_patrimonio, col_idx_placa, col_idx_barras, col_idx_desc = 0, 1, 2, 4
            col_idx_cons, col_idx_incorp, col_idx_situ, col_idx_loc = 6, 7, 9, 11
            continue
        if header_row_index is not None and index > header_row_index:
            if pd.notna(row[0]) and row[0] != 'Patrimônio':
                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual, 'Grupo': grupo_atual,
                    'Patrimônio': row[col_idx_patrimonio], 'Placa/Plaqueta': row[col_idx_placa],
                    'Cód barras': row[col_idx_barras], 'Descrição': row[col_idx_desc],
                    'Conservação': row[col_idx_cons], 'Dt. Incorporação': row[col_idx_incorp],
                    'Situação': row[col_idx_situ], 'Localização atual': row[col_idx_loc]
                })
    return pd.DataFrame(dados_reestruturados)

def relatorio_diario_eq(file):
    df = pd.read_excel(file, header=None)
    centro_custo_atual = n_registro_atual = equipamento_atual = placa_atual = responsavel_atual = observacao_atual = None
    header_row_index = None
    dados_reestruturados = []
    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():
        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])): break
        if isinstance(row[0], str):
            if 'Centro de custo' in row[0]: centro_custo_atual = row[3]
            if 'Nº registro' in row[0]: n_registro_atual = row[3]
            if 'Equipamento' in row[0]: equipamento_atual = row[3]
            if 'Placa/Plaqueta' in row[0]: placa_atual = row[3]
            if 'Responsável' in row[0]: responsavel_atual = row[3]
            if 'Observação' in row[0]: observacao_atual = row[3]

        if header_row_index is None and any(isinstance(v, str) and ('Hodômetro' in v or 'Horímetro' in v) for v in row):
            header_row_index = index
            col_idx_numero, col_idx_obra, col_idx_util, col_idx_oper = 0, 1, 4, 7
            col_idx_d_saida, col_idx_d_chegada = 9, 14
            # Busca dinâmica de hodômetro/horímetro
            col_idx_hodo_s = next((i for i, v in enumerate(row) if 'Hodômetro' in str(v) and 'saída' in str(v)), None)
            col_idx_hori_s = next((i for i, v in enumerate(row) if 'Horímetro' in str(v) and 'saída' in str(v)), None)
            col_idx_hodo_c = next((i for i, v in enumerate(row) if 'Hodômetro' in str(v) and 'chegada' in str(v)), None)
            col_idx_hori_c = next((i for i, v in enumerate(row) if 'Horímetro' in str(v) and 'chegada' in str(v)), None)
            continue

        if header_row_index is not None and index > header_row_index:
            if pd.notna(row[0]) and row[0] != 'Número':
                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual, 'Equipamento': equipamento_atual,
                    'Placa/Plaqueta': placa_atual, 'Número': row[col_idx_numero], 'Obra': row[col_idx_obra],
                    'Operador': row[col_idx_oper], 'Data saída': row[col_idx_d_saida],
                    'Hodômetro saída': row[col_idx_hodo_s] if col_idx_hodo_s else None,
                    'Horímetro saída': row[col_idx_hori_s] if col_idx_hori_s else None,
                    'Data chegada': row[col_idx_d_chegada],
                    'Hodômetro chegada': row[col_idx_hodo_c] if col_idx_hodo_c else None,
                    'Horímetro chegada': row[col_idx_hori_c] if col_idx_hori_c else None
                })
    return pd.DataFrame(dados_reestruturados)

def relatorio_eq_analitico(file):
    df = pd.read_excel(file, header=None)
    dados_reestruturados = []
    # Estado inicial do bloco
    c = {k: None for k in ['Centro', 'Eq', 'Barra', 'Placa', 'Grupo', 'Insumo', 'Det', 'Obs', 'Cons', 'Cor', 'Comb', 'Chassi', 'Pot', 'AnoF', 'AnoM', 'Obra', 'HoriA', 'HoriH', 'HodoA', 'HodoH']}

    for index, row in df.iterrows():
        row_str = row.astype(str).fillna('')
        if 'Centro de custo' in row_str.values:
            if c['Eq']: dados_reestruturados.append(c.copy())
            c['Centro'] = row[2]
        if 'Equipamento' in row_str.values: c['Eq'] = row[2]
        if 'Código barras' in row_str.values: c['Barra'] = row[2]
        if 'Placa/Plaqueta' in row_str.values: c['Placa'] = row[4]
        if 'Grupo' in row_str.values: c['Grupo'] = row[2]
        if 'Insumo' in row_str.values: c['Insumo'] = row[2]
        if 'Detalhamento' in row_str.values: c['Det'] = row[2]
        if 'Observação' in row_str.values: c['Obs'] = row[2]
        if 'Estado de conservação' in row_str.values: c['Cons'] = row[3]
        if 'Cor' in row_str.values: c['Cor'] = row[3]
        if 'Combustível' in row_str.values: c['Comb'] = row[3]
        if 'Nº de série/chassi' in row_str.values: c['Chassi'] = row[3]
        if 'Potência' in row_str.values: c['Pot'] = row[3]
        if 'Ano fabricação' in row_str.values: c['AnoF'] = row[3]
        if 'Ano modelo' in row_str.values: c['AnoM'] = row[3]
        if 'Setor/Obra atual' in row_str.values: c['Obra'] = row[3]
        if 'Horímetro' in row_str.values:
            c['HoriA'], c['HoriH'] = row[2], row[4]
        if 'Hodômetro' in row_str.values:
            c['HodoA'], c['HodoH'] = row[2], row[4]

    if c['Eq']: dados_reestruturados.append(c)
    return pd.DataFrame(dados_reestruturados)

def relatorio_mapa_controle(file):
    # Lógica unificada para 1 obra ou múltiplas obras (detecção dinâmica de header)
    df_temp = pd.read_excel(file, header=None)
    header_row_index = None
    for i in range(len(df_temp)):
        if pd.notna(df_temp.iloc[i, 0]) and 'item' in str(df_temp.iloc[i, 0]).lower():
            header_row_index = i
            break
    if header_row_index is None: return pd.DataFrame()
    
    df = pd.read_excel(file, header=header_row_index)
    colunas_fantasmas = [2, 4, 7, 9, 15, 17] # C, E, H, J, P, R
    df.drop(df.columns[colunas_fantasmas], axis=1, inplace=True, errors='ignore')
    return df.dropna(subset=[df.columns[0]])

def relatorio_historico_bens(file):
    df = pd.read_excel(file, header=None)
    pat_atual = placa_atual = cod_atual = det_atual = header_idx = last_data = None
    dados = []

    for index, row in df.iterrows():
        if row[0] == 'Patrimônio': pat_atual = row[3]
        if row[6] == 'Placa/Plaqueta': placa_atual = row[7]
        if row[0] == 'Detalhamento': det_atual = row[3]
        if row[0] == 'Data':
            header_idx = index
            continue

        if header_idx is not None and index > header_idx:
            if pd.isna(row[3]): continue # Pula se movimento for nulo
            data_to_use = row[0] if pd.notna(row[0]) else last_data
            if pd.notna(row[0]): last_data = row[0]
            
            # Lógica Origem/Destino do TXT
            cc_raw = str(row[4])
            origem = destino = ""
            if "Origem:" in cc_raw and "Destino:" in cc_raw:
                origem = cc_raw.split("Origem:")[1].split("Destino:")[0].strip()
                destino = cc_raw.split("Destino:")[1].strip()
            
            dados.append({
                'Patrimônio': pat_atual, 'Placa/Plaqueta': placa_atual, 'Detalhamento': det_atual,
                'Data': data_to_use, 'Movimento': row[3], 'Centro de Custo': row[4],
                'Origem': origem, 'Destino': destino, 'Responsável': row[11]
            })
    return pd.DataFrame(dados)

# =========================
# 🔹 DICIONÁRIO E INTERFACE
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

tipo = st.selectbox("Selecione o tipo de relatório para converter", list(relatorios.keys()))
arquivo = st.file_uploader("Arraste o arquivo .xlsx original aqui", type=["xlsx"])

if st.button("🚀 Processar e Gerar Download"):
    if arquivo:
        try:
            with st.spinner('Processando dados...'):
                df_final = relatorios[tipo](arquivo)
            
            if not df_final.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.success(f"Sucesso! {len(df_final)} registros processados.")
                st.download_button(
                    label="📥 Baixar Arquivo Convertido",
                    data=output.getvalue(),
                    file_name=f"RELATORIO_{tipo.upper().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df_final.head(100)) # Mostra preview
            else:
                st.error("O processamento resultou em um arquivo vazio. Verifique se o formato do Excel enviado é o correto para este relatório.")
        except Exception as e:
            st.error(f"Erro Crítico: {e}")
    else:
        st.warning("Por favor, anexe um arquivo antes de clicar em gerar.")
