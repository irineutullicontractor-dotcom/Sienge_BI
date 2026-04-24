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
                'Cliente/Fornecedor/Complemento': row[3],
                'Título/Parcela': row[5],
                'Documento': row[8],
                'Plano financeiro': row[10],
                'Crédito': row[13],
                'Débito': row[17]
            })

    return pd.DataFrame(dados)


def relatorio_apropriacao(file):

    df = pd.read_excel(file, header=None)

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

        if (isinstance(row[0], str) and row[0].strip() in [
            'Total da etapa','Total da subetapa','Total da célula construtiva',
            'Total da unidade construtiva','Total da obra'
        ]) or pd.isna(row[0]) or \
           (isinstance(row[0], str) and date_time_pattern.match(row[0].strip())):
            continue

        if row[0] == 'Período':
            periodo_atual = row[4]

        if row[8] == 'Seleção por':
            selecao_atual = row[13]

        if row[0] == 'Obra':
            obra_atual = row[4]

        if row[0] == 'Unidade construtiva':
            unidade_atual = row[4]

        if row[0] == 'Célula construtiva':
            celula_atual = row[4]

        if isinstance(row[0], str) and row[0].strip() == 'Etapa':
            etapa_atual = row[4]
            subetapa_atual = None
            continue

        if isinstance(row[0], str) and row[0].strip() == 'Subetapa':
            subetapa_atual = row[4]
            continue

        if row[0] == 'Data':
            header_row_index = index
            continue

        if header_row_index is not None and index > header_row_index:

            val_data = row[0]
            if isinstance(val_data, str) and val_data.strip() in ['Período','Seleção por','Data']:
                continue

            dados_reestruturados.append({
                'Período': periodo_atual,
                'Seleção por': selecao_atual,
                'Obra': obra_atual,
                'Unidade construtiva': unidade_atual,
                'Célula construtiva': celula_atual,
                'Etapa': etapa_atual,
                'Subetapa': subetapa_atual,
                'Data': row[0],
                'Documento': row[1],
                'Título/Parcela': row[4],
                'Or': row[6],
                'Credor / Histórico': row[7],
                'Valor do documento': row[12],
                'Valor apropriado': row[14],
            })

    df_reestruturado = pd.DataFrame(dados_reestruturados)

    column_order = [
        'Período','Seleção por','Obra','Unidade construtiva',
        'Célula construtiva','Etapa','Subetapa','Data','Documento',
        'Título/Parcela','Or','Credor / Histórico',
        'Valor do documento','Valor apropriado'
    ]

    return df_reestruturado[column_order]


def relatorio_bens(file):

    df = pd.read_excel(file, header=None)

    centro_custo_atual = None
    grupo_atual = None
    header_row_index = None
    dados_reestruturados = []

    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    for index, row in df.iterrows():

        if isinstance(row[0], str) and date_time_pattern.match(str(row[0])):
            break

        if row[0] == 'Centro de custo':
            centro_custo_atual = row[3]

        if row[0] == 'Grupo':
            grupo_atual = row[3]

        if row[0] == 'Patrimônio':
            header_row_index = index

        if header_row_index is not None and index > header_row_index:

            if pd.notna(row[0]) and row[0] not in ['Centro de custo','Grupo']:

                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual,
                    'Grupo': grupo_atual,
                    'Patrimônio': row[0],
                    'Placa/Plaqueta': row[1],
                    'Cód barras': row[2],
                    'Descrição': row[4],
                    'Conservação': row[6],
                    'Dt. Incorporação': row[7],
                    'Situação': row[9],
                    'Localização atual': row[10]
                })

    df = pd.DataFrame(dados_reestruturados)

    return df[
        ['Centro de custo','Grupo','Patrimônio','Placa/Plaqueta','Cód barras',
         'Descrição','Conservação','Dt. Incorporação','Situação','Localização atual']
    ]


def relatorio_historico_bens(file):

    df = pd.read_excel(file, header=None)

    patrimonio_atual = None
    placa_atual = None
    detalhamento_atual = None
    header_row_index = None
    dados = []

    last_data = None
    last_tipo = None

    for index, row in df.iterrows():

        if row[0] == 'Patrimônio':
            patrimonio_atual = row[3]

        if row[6] == 'Placa/Plaqueta':
            placa_atual = row[7]

        if row[0] == 'Detalhamento':
            detalhamento_atual = row[3]

        if row[0] == 'Data':
            header_row_index = index
            continue

        if header_row_index is not None and index > header_row_index:

            val_data = row[0]
            if isinstance(val_data, str) and val_data.strip() in ['Patrimônio','Detalhamento','Data']:
                continue

            data = row[0] if pd.notna(row[0]) else last_data
            if pd.notna(row[0]): last_data = row[0]

            tipo = row[1] if pd.notna(row[1]) else last_tipo
            if pd.notna(row[1]): last_tipo = row[1]

            dados.append({
                'Patrimônio': patrimonio_atual,
                'Placa/Plaqueta': placa_atual,
                'Detalhamento': detalhamento_atual,
                'Data': data,
                'Tipo do movimento': tipo,
                'Movimento': row[3],
                'Centro(s) de Custo': row[4],
                'Responsável': row[11],
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
    
def relatorio_eq_analitico(file):

    import pandas as pd
    import re

    df = pd.read_excel(file, header=None)

    dados_reestruturados = []

    date_time_pattern = re.compile(r'\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2}')

    centro_custo_atual = None
    equipamento_atual = None
    codigo_barras_atual = None
    placa_atual = None
    grupo_atual = None
    insumo_atual = None
    detalhamento_atual = None
    observacao_atual = None
    estado_conservacao_atual = None
    cor_atual = None
    combustivel_atual = None
    n_chassi_atual = None
    potencia_atual = None
    ano_fab_atual = None
    ano_mod_atual = None
    obra_atual = None
    atual_atual = None
    historico_atual = None
    horimetro_atual = None
    horimetro_historico = None
    hodometro_atual = None
    hodometro_historico = None

    for index, row in df.iterrows():
        row_str = row.astype(str).fillna('')

        if 'Centro de custo' in row_str.values:

            if equipamento_atual is not None:
                dados_reestruturados.append({
                    'Centro de custo': centro_custo_atual,
                    'Equipamento': equipamento_atual,
                    'Código barras': codigo_barras_atual,
                    'Placa/Plaqueta': placa_atual,
                    'Grupo': grupo_atual,
                    'Insumo': insumo_atual,
                    'Detalhamento': detalhamento_atual,
                    'Observação': observacao_atual,
                    'Estado de conservação': estado_conservacao_atual,
                    'Cor': cor_atual,
                    'Combustível': combustivel_atual,
                    'Nº de série/chassi': n_chassi_atual,
                    'Potência': potencia_atual,
                    'Ano fabricação': ano_fab_atual,
                    'Ano modelo': ano_mod_atual,
                    'Setor/Obra atual': obra_atual,
                    'Horímetro Atual': horimetro_atual,
                    'Horímetro Histórico': horimetro_historico,
                    'Hodômetro Atual': hodometro_atual,
                    'Hodômetro Histórico': hodometro_historico
                })

            # RESET
            centro_custo_atual = None
            equipamento_atual = None
            codigo_barras_atual = None
            placa_atual = None
            grupo_atual = None
            insumo_atual = None
            detalhamento_atual = None
            observacao_atual = None
            estado_conservacao_atual = None
            cor_atual = None
            combustivel_atual = None
            n_chassi_atual = None
            potencia_atual = None
            ano_fab_atual = None
            ano_mod_atual = None
            obra_atual = None
            horimetro_atual = None
            horimetro_historico = None
            hodometro_atual = None
            hodometro_historico = None

            centro_custo_atual = row_str[2]

        if 'Equipamento' in row_str.values:
            equipamento_atual = row_str[2]

        if 'Código barras' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Código barras', na=False)].index[0]
                codigo_barras_atual = row_str[col_index + 1]
            except:
                pass

        if 'Placa/Plaqueta' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Placa/Plaqueta', na=False)].index[0]
                placa_atual = row_str[col_index + 2]
            except:
                pass

        if 'Grupo' in row_str.values:
            grupo_atual = row_str[2]

        if 'Insumo' in row_str.values:
            insumo_atual = row_str[2]

        if 'Detalhamento' in row_str.values:
            detalhamento_atual = row_str[2]

        if 'Observação' in row_str.values:
            observacao_atual = row_str[2]

        if 'Estado de conservação' in row_str.values:
            estado_conservacao_atual = row_str[2]

        if 'Cor' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Cor', na=False)].index[0]
                cor_atual = row_str[col_index + 1]
            except:
                pass

        if 'Combustível' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Combustível', na=False)].index[0]
                combustivel_atual = row_str[col_index + 2]
            except:
                pass

        if 'Nº de série/chassi' in row_str.values:
            n_chassi_atual = row_str[2]

        if 'Potência' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Potência', na=False)].index[0]
                potencia_atual = row_str[col_index + 1]
            except:
                pass

        if 'Ano fabricação' in row_str.values:
            ano_fab_atual = row_str[2]

        if 'Ano modelo' in row_str.values:
            try:
                col_index = row_str[row_str.str.contains('Ano modelo', na=False)].index[0]
                ano_mod_atual = row_str[col_index + 1]
            except:
                pass

        if 'Setor/Obra atual' in row_str.values:
            obra_atual = row_str[2]

        if row_str[0] == 'Atual' and row_str[4] == 'Histórico':
            atual_value = row_str[2]
            historico_value = row_str[5]

            if 'h' in atual_value or 'h' in historico_value:
                horimetro_atual = atual_value
                horimetro_historico = historico_value
                hodometro_atual = None
                hodometro_historico = None
            else:
                hodometro_atual = atual_value
                hodometro_historico = historico_value
                horimetro_atual = None
                horimetro_historico = None

    if equipamento_atual is not None:
        dados_reestruturados.append({
            'Centro de custo': centro_custo_atual,
            'Equipamento': equipamento_atual,
            'Código barras': codigo_barras_atual,
            'Placa/Plaqueta': placa_atual,
            'Grupo': grupo_atual,
            'Insumo': insumo_atual,
            'Detalhamento': detalhamento_atual,
            'Observação': observacao_atual,
            'Estado de conservação': estado_conservacao_atual,
            'Cor': cor_atual,
            'Combustível': combustivel_atual,
            'Nº de série/chassi': n_chassi_atual,
            'Potência': potencia_atual,
            'Ano fabricação': ano_fab_atual,
            'Ano modelo': ano_mod_atual,
            'Setor/Obra atual': obra_atual,
            'Horímetro Atual': horimetro_atual,
            'Horímetro Histórico': horimetro_historico,
            'Hodômetro Atual': hodometro_atual,
            'Hodômetro Histórico': hodometro_historico
        })

    df_reestruturado = pd.DataFrame(dados_reestruturados)

    column_order = [
        'Centro de custo','Equipamento','Código barras','Placa/Plaqueta','Grupo',
        'Insumo','Detalhamento','Observação','Estado de conservação','Cor',
        'Combustível','Nº de série/chassi','Potência','Ano fabricação',
        'Ano modelo','Setor/Obra atual','Horímetro Atual','Horímetro Histórico',
        'Hodômetro Atual','Hodômetro Histórico'
    ]

    if not df_reestruturado.empty:
        df_reestruturado = df_reestruturado[column_order]

    return df_reestruturado

def relatorio_mapa_controle_1_obra(file):

    import pandas as pd

    df = pd.read_excel(file, header=6)

    colunas_para_remover = [2, 4, 7, 9, 15, 17]
    df.drop(df.columns[colunas_para_remover], axis=1, inplace=True)

    df_final = df.dropna(subset=['Item'])

    return df_final
    
# =========================
# 🔹 DICIONÁRIO
# =========================

relatorios = {
    "Financeiro": relatorio_financeiro,
    "Apropriação de Obra": relatorio_apropriacao,
    "Bens Sintético": relatorio_bens,
    "Histórico de Bens": relatorio_historico_bens,
    "Diário Equipamento COMPLETO": relatorio_diario_eq_completo,
    "Equipamento Analítico (Completo)": relatorio_eq_analitico,
    "Mapa Controle (1 Obra)": relatorio_mapa_controle_1_obra,
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
