import pandas as pd

# Carregar os dados da planilha fornecida
df = pd.read_excel('trajetos_e_paradas_01-08-2026_31-08-2026_1059442041_2.xlsx')

# Identificar nomes/índices das colunas (Colunas A, B, J, K correspondem aos índices 0, 1, 9, 10)
col_data = df.columns[0]
col_tag = df.columns[1]
col_dist = df.columns[9]
col_tempo = df.columns[10]

# Garantir tipos de dados corretos
df[col_data] = pd.to_datetime(df[col_data])
df[col_dist] = pd.to_numeric(df[col_dist], errors='coerce').fillna(0)

# Converter tempo (Coluna K) para minutos se estiver em string HH:MM ou HH:MM:SS
def converter_para_minutos(val):
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return val
    val_str = str(val).strip()
    partes = val_str.split(':')
    if len(partes) == 3:
        return int(partes[0]) * 60 + int(partes[1]) + int(partes[2]) / 60
    elif len(partes) == 2:
        return int(partes[0]) * 60 + int(partes[1])
    return 0

df['tempo_minutos'] = df[col_tempo].apply(converter_para_minutos)

# Aplicar o critério de dia trabalhado (Coluna J > 4 e Coluna K > 30 min)
df['trabalhou'] = (df[col_dist] > 4) & (df['tempo_minutos'] > 30)

# Processar regras por equipamento
resultados = []

for tag, grupo in df.groupby(col_tag):
    grupo = grupo.sort_values(by=col_data)
    
    dias_trabalhados = grupo[grupo['trabalhou']][col_data].dt.date.tolist()
    
    # 1. Cálculo dos blocos de cobrança (faturamento em períodos de 7 dias)
    dias_cobrados = 0
    periodos_cobrados = 0
    
    if dias_trabalhados:
        data_inicio_bloco = dias_trabalhados[0]
        periodos_cobrados = 1
        
        for d in dias_trabalhados[1:]:
            dias_decorridos = (d - data_inicio_bloco).days
            if dias_decorridos >= 7:
                blocos_adicionais = dias_decorridos // 7
                periodos_cobrados += blocos_adicionais
                data_inicio_bloco = data_inicio_bloco + pd.Timedelta(days=blocos_adicionais * 7)
                
        dias_cobrados = periodos_cobrados * 7

    # 2. Verificar se ficou sem uso por pelo menos 15 dias seguidos no período
    menor_data = grupo[col_data].min()
    maior_data = grupo[col_data].max()
    intervalo_completo = pd.date_range(start=menor_data, end=maior_data)
    
    dias_ativos_set = set(dias_trabalhados)
    
    max_dias_sem_uso = 0
    contador_sem_uso = 0
    
    for dt in intervalo_completo:
        if dt.date() not in dias_ativos_set:
            contador_sem_uso += 1
            if contador_sem_uso > max_dias_sem_uso:
                max_dias_sem_uso = contador_sem_uso
        else:
            contador_sem_uso = 0
            
    ficou_15_dias_parado = max_dias_sem_uso >= 15

    resultados.append({
        'Equipamento (Tag)': tag,
        'Dias Trabalhados (Efetivos)': len(dias_trabalhados),
        'Períodos de 7d': periodos_cobrados,
        'Dias a Cobrar': dias_cobrados,
        'Max Dias Consecutivos Parado': max_dias_sem_uso,
        'Sinalização (>=15d Parado)': 'SIM - Parado >= 15 dias' if ficou_15_dias_parado else 'Não'
    })

df_resumo = pd.DataFrame(resultados)
print(df_resumo.to_string(index=False))
