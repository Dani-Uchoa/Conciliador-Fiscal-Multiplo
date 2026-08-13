import streamlit as st
import pandas as pd
import io

# Configuração da Página
st.set_page_config(page_title="Conciliador Fiscal - Auditoria", layout="wide")

st.title("📊 Conciliador Fiscal Universal - Módulo de Auditoria")
st.markdown("---")

# Sidebar - Configurações Gerais
st.sidebar.header("⚙️ Configurações do Processamento")
tipo_operacao = st.sidebar.selectbox("Tipo de Operação", ["Saídas", "Entradas"])

st.sidebar.subheader("5. Fontes de Dados para Auditoria")
# ITEM 5: Nomenclatura ajustada para "Fonte Oficial" mantendo a "Fonte Secundária/Complementar"
fonte_oficial_nome = st.sidebar.text_input("Nome da Fonte 1:", value="Fonte Oficial")
fonte_secundaria_nome = st.sidebar.text_input("Nome da Fonte 2:", value="Fonte Domínio / SIEG")

# Upload dos Arquivos
col1, col2 = st.columns(2)

with col1:
    st.subheader(f"1. Arquivo: {fonte_oficial_nome}")
    file_oficial = st.file_uploader(f"Selecione o arquivo da {fonte_oficial_nome} (Excel/CSV)", type=["xlsx", "xls", "csv"], key="oficial")

with col2:
    st.subheader(f"2. Arquivo: {fonte_secundaria_nome}")
    file_secundario = st.file_uploader(f"Selecione o arquivo da {fonte_secundaria_nome} (Excel/CSV)", type=["xlsx", "xls", "csv"], key="secundario")

# Funções Auxiliares de Tratamento de Dados
def carregar_dados(file):
    if file is not None:
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, dtype=str)
            else:
                df = pd.read_excel(file, dtype=str)
            return df
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo: {e}")
            return None
    return None

def sanitizar_df(df):
    if df is not None:
        # Limpeza de nomes de colunas
        df.columns = df.columns.str.strip().str.upper()
        # Remoção de espaços nas células
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# Processamento da Auditoria
if file_oficial and file_secundario:
    df_oficial = sanitizar_df(carregar_dados(file_oficial))
    df_secundario = sanitizar_df(carregar_dados(file_secundario))

    st.markdown("---")
    st.header("🔍 Mapeamento das Colunas de Chave")

    col_m1, col_m2 = st.columns(2)
    
    with col_m1:
        st.markdown(f"**Mapeamento - {fonte_oficial_nome}**")
        col_num_oficial = st.selectbox(f"Coluna Número da Nota ({fonte_oficial_nome})", df_oficial.columns)
        col_valor_oficial = st.selectbox(f"Coluna Valor Total ({fonte_oficial_nome})", df_oficial.columns)
        col_status_oficial = st.selectbox(f"Coluna Status/Situação ({fonte_oficial_nome}) - Opcional", ["Nenhum"] + list(df_oficial.columns))

    with col_m2:
        st.markdown(f"**Mapeamento - {fonte_secundaria_nome}**")
        col_num_sec = st.selectbox(f"Coluna Número da Nota ({fonte_secundaria_nome})", df_secundario.columns)
        col_valor_sec = st.selectbox(f"Coluna Valor Total ({fonte_secundaria_nome})", df_secundario.columns)
        col_status_sec = st.selectbox(f"Coluna Status/Situação ({fonte_secundaria_nome}) - Opcional", ["Nenhum"] + list(df_secundario.columns))

    if st.button("🚀 Executar Auditoria / Conciliação"):
        # Expurgo de notas canceladas/denegadas se coluna selecionada
        if col_status_oficial != "Nenhum":
            df_oficial = df_oficial[~df_oficial[col_status_oficial].str.contains("CANCELAD|DENEGAD|INUTILIZAD", case=False, na=False)]
        
        if col_status_sec != "Nenhum":
            df_secundario = df_secundario[~df_secundario[col_status_sec].str.contains("CANCELAD|DENEGAD|INUTILIZAD", case=False, na=False)]

        # Tratamento numérico dos valores
        df_oficial['VALOR_CLEAN'] = pd.to_numeric(df_oficial[col_valor_oficial].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)
        df_secundario['VALOR_CLEAN'] = pd.to_numeric(df_secundario[col_valor_sec].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

        # Padronização do número da nota
        df_oficial['NUMERO_CLEAN'] = df_oficial[col_num_oficial].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
        df_secundario['NUMERO_CLEAN'] = df_secundario[col_num_sec].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')

        # Cruzamento dos Dados (Merge)
        df_cruzado = pd.merge(
            df_oficial,
            df_secundario,
            left_on='NUMERO_CLEAN',
            right_on='NUMERO_CLEAN',
            how='outer',
            suffixes=(f'_{fonte_oficial_nome}', f'_{fonte_secundaria_nome}')
        )

        # Identificação de Inconsistências
        df_cruzado['DIFERENCA_VALOR'] = df_cruzado['VALOR_CLEAN_' + fonte_oficial_nome].fillna(0) - df_cruzado['VALOR_CLEAN_' + fonte_secundaria_nome].fillna(0)
        
        def classificar_status(row):
            num_oficial = row.get('VALOR_CLEAN_' + fonte_oficial_nome)
            num_sec = row.get('VALOR_CLEAN_' + fonte_secundaria_nome)
            
            if pd.isna(num_oficial) or num_oficial == 0:
                return f"Falta na {fonte_oficial_nome}"
            elif pd.isna(num_sec) or num_sec == 0:
                return f"Falta na {fonte_secundaria_nome}"
            elif abs(row['DIFERENCA_VALOR']) > 0.01:
                return "Divergência de Valor"
            else:
                return "Conciliado OK"

        df_cruzado['STATUS_AUDITORIA'] = df_cruzado.apply(classificar_status, axis=1)

        # Exibição de Resumo em Métricas
        st.subheader("📈 Painel do Resultado da Auditoria")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Registros Conciliados", len(df_cruzado[df_cruzado['STATUS_AUDITORIA'] == "Conciliado OK"]))
        m2.metric("Divergências de Valor", len(df_cruzado[df_cruzado['STATUS_AUDITORIA'] == "Divergência de Valor"]))
        m3.metric(f"Ausentes na {fonte_oficial_nome}", len(df_cruzado[df_cruzado['STATUS_AUDITORIA'] == f"Falta na {fonte_oficial_nome}"]))
        m4.metric(f"Ausentes na {fonte_secundaria_nome}", len(df_cruzado[df_cruzado['STATUS_AUDITORIA'] == f"Falta na {fonte_secundaria_nome}"]))

        # Exibição da Tabela Filtrada de Divergências
        st.subheader("⚠️ Relatório Detalhado de Exceções e Divergências")
        df_excecoes = df_cruzado[df_cruzado['STATUS_AUDITORIA'] != "Conciliado OK"]
        st.dataframe(df_excecoes, use_container_width=True)

        # Exportação para Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_cruzado.to_excel(writer, sheet_name="Auditoria Completa", index=False)
            df_excecoes.to_excel(writer, sheet_name="Apenas Divergencias", index=False)
        
        st.download_button(
            label="📥 Baixar Relatório de Auditoria em Excel",
            data=buffer.getvalue(),
            file_name="Relatorio_Auditoria_Fiscal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
