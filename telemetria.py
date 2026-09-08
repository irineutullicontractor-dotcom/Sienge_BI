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
            full_month_dates = pd.date_range(min_date, max_date)
            
            # Tratamento para evitar erro com tipos mistos nos nomes dos veículos/tags
            vehicles = sorted([str(v) for v in df_valid[col_tag].dropna().unique()])
            summary = []
            
            for v in vehicles:
                df_v = df_valid[df_valid[col_tag].astype(str) == v]
                
                # Calendário completo do equipamento
                calendar = pd.DataFrame({'Data_Parsed': full_month_dates})
                merged = pd.merge(calendar, df_v[['Data_Parsed', 'Dia_Trabalhado']], on='Data_Parsed', how='left')
                merged['Dia_Trabalhado'] = merged['Dia_Trabalhado'].fillna(False)
                
                merged['Dia_Mes'] = merged['Data_Parsed'].dt.day
                merged['Bloco'] = ((merged['Dia_Mes'] - 1) // bloco_dias) + 1
                
                # Cálculo dos blocos ativos
                active_blocks = merged.groupby('Bloco')['Dia_Trabalhado'].any().sum()
                dias_cobrar = active_blocks * bloco_dias
                dias_efetivos = merged['Dia_Trabalhado'].sum()
                
                # Sequência máxima de dias inativos seguidos
                streak = 0
                max_streak = 0
                for inactive in (~merged['Dia_Trabalhado']):
                    if inactive:
                        streak += 1
                        if streak > max_streak:
                            max_streak = streak
                    else:
                        streak = 0
                        
                alerta = "🚨 INATIVO (≥15d)" if max_streak >= dias_alerta_inativo else "✅ OK"
                
                summary.append({
                    'Equipamento / Tag': v,
                    'Dias Trabalhados Validados': dias_efetivos,
                    'Blocos de 7d Ativos': active_blocks,
                    'Mínimo Dias a Cobrar': dias_cobrar,
                    'Maior Seq. Inativa (Dias)': max_streak,
                    'Status Inatividade (≥15d)': alerta
                })
                
            res_df = pd.DataFrame(summary)
            
            # Indicadores Principais
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total de Equipamentos", len(res_df))
            m2.metric("Equipamentos Ativos", len(res_df[res_df['Dias Trabalhados Validados'] > 0]))
            m3.metric("Equipamentos Alertados (Inativos)", len(res_df[res_df['Status Inatividade (≥15d)'].astype(str).str.contains('INATIVO')]))
            m4.metric("Total Mínimo de Dias a Cobrar", f"{res_df['Mínimo Dias a Cobrar'].sum()} dias")
            
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
