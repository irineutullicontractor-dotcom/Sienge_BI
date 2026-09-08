import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Resumo de Telemetria e Faturamento", layout="wide")

st.title("🚗 Resumo de Telemetria e Faturamento de Veículos")
st.markdown("Faça upload da planilha de telemetria para geração automática do resumo por equipamento, cálculo de blocos de 7 dias e alertas de inatividade.")

# Parâmetros Editáveis na barra lateral
st.sidebar.header("⚙️ Parâmetros das Regras")
min_dist = st.sidebar.number_input("Distância mínima para considerar trabalho (km)", value=4.0, step=0.5)
min_tempo = st.sidebar.number_input("Tempo mínimo dirigido para considerar trabalho (min)", value=30.0, step=5.0)
bloco_dias = st.sidebar.number_input("Tamanho do Bloco de Cobrança (dias)", value=7, step=1)
dias_alerta_inativo = st.sidebar.number_input("Dias seguidos sem uso para Sinalizar Alerta", value=15, step=1)

uploaded_file = st.file_uploader("Selecione a planilha de telemetria (.xlsx ou .csv)", type=["xlsx", "xls", "csv"])

def clean_number(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).replace('.', '').replace(',', '.').strip()
    try:
        return float(val_str)
    except Exception:
        return 0.0

if uploaded_file is not None:
    try:
        # Leitura e detecção automática do cabeçalho com conversão explícita para string
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(uploaded_file, header=None)
            header_idx = 0
            for idx, row in df_raw.iterrows():
                row_str = [str(item).lower() for item in row if pd.notna(item)]
                if any('data' in s for s in row_str) and any(('dist' in s or 'placa' in s or 'tag' in s) for s in row_str):
                    header_idx = idx
                    break
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, skiprows=header_idx)
        else:
            df = pd.read_csv(uploaded_file)
            
        st.success("Arquivo carregado com sucesso!")
        
        # Mapeamento dinâmico das colunas
        cols = [str(c) for c in df.columns.tolist()]
        df.columns = cols
        
        st.subheader("📌 Mapeamento de Colunas")
        col1, col2, col3, col4 = st.columns(4)
        
        default_data = next((c for c in cols if 'data' in c.lower()), cols[0])
        default_tag = next((c for c in cols if any(k in c.lower() for k in ['tag', 'placa', 'veiculo', 'veículo', 'equipamento'])), cols[1 if len(cols)>1 else 0])
        default_dist = next((c for c in cols if 'dist' in c.lower()), cols[9 if len(cols)>9 else 0])
        default_tempo = next((c for c in cols if 'tempo' in c.lower() and 'dirigido' in c.lower()), cols[10 if len(cols)>10 else 0])
        
        col_data = col1.selectbox("Coluna Data (Coluna A)", cols, index=cols.index(default_data))
        col_tag = col2.selectbox("Coluna Tag / Equipamento (Coluna B)", cols, index=cols.index(default_tag))
        col_dist = col3.selectbox("Coluna Distância (Coluna J)", cols, index=cols.index(default_dist))
        col_tempo = col4.selectbox("Coluna Tempo Dirigido (Coluna K)", cols, index=cols.index(default_tempo))
        
        # Tratamento dos dados
        df['Data_Parsed'] = pd.to_datetime(df[col_data], format='%d/%m/%Y', errors='coerce')
        if df['Data_Parsed'].isna().sum() > len(df) * 0.5:
            df['Data_Parsed'] = pd.to_datetime(df[col_data], errors='coerce')
            
        df['Dist_Parsed'] = df[col_dist].apply(clean_number)
        df['Tempo_Parsed'] = df[col_tempo].apply(clean_number)
        
        # Condição de dia trabalhado
        df['Dia_Trabalhado'] = (df['Dist_Parsed'] > min_dist) & (df['Tempo_Parsed'] > min_tempo)
        df_valid = df.dropna(subset=['Data_Parsed']).copy()
        
        if len(df_valid) == 0:
            st.error("Não foi possível identificar datas válidas na coluna selecionada.")
        else:
            # Intervalo completo do mês
            min_date = df_valid['Data_Parsed'].min().replace(day=1)
            max_date = (min_date + pd.offsets.MonthEnd(1))
            total_days_in_month = max_date.day
            
            vehicles = sorted([str(v) for v in df_valid[col_tag].dropna().unique()])
            summary = []
            
            for v in vehicles:
                df_v = df_valid[df_valid[col_tag].astype(str) == v]
                worked_dates = sorted(df_v[df_v['Dia_Trabalhado']]['Data_Parsed'].dt.day.tolist())
                
                periodos = ["N/A", "N/A", "N/A", "N/A"]
                periodo_desconto = "N/A"
                dias_cobrar = total_days_in_month
                
                if not worked_dates:
                    # Sem nenhum uso validado
                    dias_cobrar = 0
                    periodo_desconto = f"01 a {total_days_in_month}"
                else:
                    # Cálculo dos blocos de 7 dias e verificação de inatividade >= 15 dias
                    first_day = worked_dates[0]
                    current_day = first_day
                    period_idx = 0
                    inactivity_periods = []
                    
                    # Verificação de inatividade inicial (do dia 1 até o primeiro uso)
                    if (first_day - 1) >= dias_alerta_inativo:
                        inactivity_periods.append(f"01 a {first_day - 1:02d}")
                    
                    # Processamento dos blocos a partir do primeiro uso
                    active_ranges = []
                    while current_day <= total_days_in_month and period_idx < 4:
                        start_block = current_day
                        end_block = min(current_day + bloco_dias - 1, total_days_in_month)
                        
                        periodos[period_idx] = f"{start_block:02d} a {end_block:02d}"
                        active_ranges.append((start_block, end_block))
                        period_idx += 1
                        
                        # Próximo dia trabalhado após o término do bloco de 7 dias
                        next_worked = [d for d in worked_dates if d > end_block]
                        if next_worked:
                            next_day = next_worked[0]
                            gap = next_day - (end_block + 1)
                            if gap >= dias_alerta_inativo:
                                inactivity_periods.append(f"{end_block + 1:02d} a {next_day - 1:02d}")
                            current_day = next_day
                        else:
                            # Verificação de inatividade no final do mês após o último bloco
                            gap_end = total_days_in_month - end_block
                            if gap_end >= dias_alerta_inativo:
                                inactivity_periods.append(f"{end_block + 1:02d} a {total_days_in_month:02d}")
                            break
                    
                    if inactivity_periods:
                        periodo_desconto = ", ".join(inactivity_periods)
                        # Se houver período(s) de desconto >= 15 dias, subtrai esses dias do total do mês
                        dias_desconto = 0
                        for p in inactivity_periods:
                            s_d, e_d = map(int, p.split(" a "))
                            dias_desconto += (e_d - s_d + 1)
                        dias_cobrar = total_days_in_month - dias_desconto
                    else:
                        dias_cobrar = total_days_in_month

                summary.append({
                    'Equipamento / Tag': v,
                    'Dias Trabalhados Validados': len(worked_dates),
                    'Período 1': periodos[0],
                    'Período 2': periodos[1],
                    'Período 3': periodos[2],
                    'Período 4': periodos[3],
                    'Período de Desconto (≥15d sem uso)': periodo_desconto,
                    'Total Dias a Cobrar': dias_cobrar
                })
                
            res_df = pd.DataFrame(summary)
            
            # Indicadores Principais
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total de Equipamentos", len(res_df))
            m2.metric("Equipamentos Ativos", len(res_df[res_df['Dias Trabalhados Validados'] > 0]))
            m3.metric("Equipamentos com Desconto (≥15d)", len(res_df[res_df['Período de Desconto (≥15d sem uso)'] != 'N/A']))
            m4.metric("Total de Dias a Cobrar (Geral)", f"{res_df['Total Dias a Cobrar'].sum()} dias")
            
            st.markdown("---")
            st.subheader("📊 Tabela de Resumo por Equipamento")
            st.dataframe(res_df, use_container_width=True)
            
            # Download do Resumo em Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False, sheet_name='Resumo_Faturamento')
            
            st.download_button(
                label="📥 Baixar Tabela em Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="resumo_telemetria_faturamento.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
