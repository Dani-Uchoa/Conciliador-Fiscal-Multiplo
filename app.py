import streamlit as st
import pandas as pd
import re
import unicodedata
import io
from functools import reduce

st.set_page_config(page_title="Conciliador Fiscal Universal", layout="wide")
st.title("⚖️ Conciliador Fiscal Universal — Multi-Fonte")

# =========================================================
# UTILITÁRIOS & TRATAMENTO DE TEXTO/MOEDA
# =========================================================
def formatar_moeda_br(valor):
    try:
        valor = float(valor)
        s = f"{valor:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {s}"
    except:
        return "R$ 0,00"

def normalizar(txt):
    if pd.isna(txt): return ""
    return unicodedata.normalize('NFD', str(txt)).encode('ascii', 'ignore').decode('utf-8').upper().strip()

def limpar_valor(v):
    if pd.isna(v): return 0.0
    s = str(v).replace('R$', '').replace('"', '').replace('\xa0', '').replace(' ', '').strip()
    if not s: return 0.0
    if ',' in s: s = s.replace('.', '').replace(',', '.')
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

def converter_data(d):
    if pd.isna(d): return None
    s = str(d).strip()
    if s.replace('.', '').isdigit() and len(s) >= 5:
        try: return pd.to_datetime(float(s), unit='D', origin='1899-12-30').date()
        except: pass
    if '.' in s and not s.replace('.', '').isdigit(): s = s.replace('.', '/')
    if re.match(r'^\d{4}-\d{2}-\d{2}', s):
        try: return pd.to_datetime(s, errors='coerce').date()
        except: pass
    try: return pd.to_datetime(s, dayfirst=True, errors='raise').date()
    except: return pd.to_datetime(s, errors='coerce').date()

def extrair_identificador_nota(n, tipo_doc):
    """
    Retorna a Chave de 44 dígitos completa para NFe/NFCe/CTe.
    Usa apenas a numeração sequencial para NFSe ou quando a chave não estiver completa.
    """
    if pd.isna(n): return ""
    s = str(n).strip()
    digitos = re.sub(r'\D', '', s)
    
    if len(digitos) == 44 and tipo_doc != "NFSe":
        return digitos
    
    if len(digitos) == 44:
        digitos = digitos[25:34]
    return digitos.lstrip('0') if digitos else ""

def verificar_status_invalido(txt):
    """
    Retorna True se o registro indicar Cancelamento, Denegação, Inutilização ou Desconhecimento.
    """
    if pd.isna(txt): return False
    s = normalizar(txt)
    termos_invalidos = ["CANCEL", "DENEG", "INUTILIZ", "DESC. OP", "DESCONHEC"]
    return any(termo in s for termo in termos_invalidos)

# =========================================================
# DETECTOR DE CABEÇALHO E LEITURA
# =========================================================
def encontrar_cabecalho(df):
    termos_fortes = ["CHAVE", "NOTA", "DATA", "VALOR", "EMISSAO", "NUMERO", "NUM NFSE", "CNPJ", "EVENTO", "SITUACAO"]
    for i in range(min(len(df), 50)):
        linha = [normalizar(str(c)) for c in df.iloc[i]]
        matches = sum(1 for c in linha if any(t in c for t in termos_fortes))
        if matches >= 3: return i

    termos_simples = ["CHAVE", "NOTA", "DATA", "VALOR"]
    for i in range(min(len(df), 50)):
        linha = [normalizar(str(c)) for c in df.iloc[i]]
        matches = sum(1 for c in linha if any(t in c for t in termos_simples))
        if matches >= 2: return i
    return 0

def carregar_planilha(f):
    f.seek(0)
    conteudo = f.read()
    df = None

    if f.name.lower().endswith('.csv'):
        try: texto = conteudo.decode('utf-8')
        except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
        primeira_linha = texto.split('\n')[0] if '\n' in texto else texto
        separador = ';' if ';' in primeira_linha else ','
        df = pd.read_csv(io.StringIO(texto), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
    else:
        is_real_xls = conteudo.startswith(b'\xD0\xCF\x11\xE0')
        is_real_xlsx = conteudo.startswith(b'PK')

        if is_real_xls or is_real_xlsx:
            motor = 'openpyxl' if is_real_xlsx else 'xlrd'
            try:
                df = pd.read_excel(io.BytesIO(conteudo), header=None, dtype=str, engine=motor)
                if len(df.columns) == 1 or df.iloc[:, 1:].isna().all().all():
                    texto_esmagado = df.iloc[:, 0].dropna().astype(str).str.cat(sep='\n')
                    if texto_esmagado.strip():
                        separador = ';' if ';' in texto_esmagado.split('\n')[0] else ','
                        df = pd.read_csv(io.StringIO(texto_esmagado), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
            except Exception:
                st.error(f"🛑 **ARQUIVO CORROMPIDO:** O arquivo '{f.name}' possui falhas na matriz binária. Abra-o no Excel e salve-o como '.xlsx'.")
                return pd.DataFrame()
        else:
            try:
                dfs = pd.read_html(io.BytesIO(conteudo))
                df = dfs[0].astype(str)
            except Exception:
                try: texto = conteudo.decode('utf-8')
                except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
                sep = '\t' if '\t' in texto else (';' if ';' in texto else ',')
                df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, header=None, engine='python', on_bad_lines='skip')

    if df is None or df.empty: return pd.DataFrame()
    idx_cabecalho = encontrar_cabecalho(df)
    if idx_cabecalho >= len(df): return df

    nomes_colunas = [str(c).strip() if pd.notna(c) else f"Coluna_{i}" for i, c in enumerate(df.iloc[idx_cabecalho])]
    vistos = {}
    nomes_unicos = []
    for nome in nomes_colunas:
        if nome in vistos:
            vistos[nome] += 1
            nomes_unicos.append(f"{nome} ({vistos[nome]})")
        else:
            vistos[nome] = 0
            nomes_unicos.append(nome)

    df.columns = nomes_unicos
    df = df.iloc[idx_cabecalho + 1:].reset_index(drop=True)
    return df

def sugerir_colunas(cols):
    idx_nota = next((i for i, c in enumerate(cols) if any(t in normalizar(c) for t in ["CHAVE", "NOTA", "NUMERO", "NUM NFSE", "NUM", "DOC"])), 0)
    idx_data = next((i for i, c in enumerate(cols) if any(t in normalizar(c) for t in ["DATA", "EMISSAO", "ENTRADA"])), 0)
    idx_valor = next((i for i, c in enumerate(cols) if any(t in normalizar(c) for t in ["VALOR", "CONTABIL"])), 0)
    idx_evento = next((i for i, c in enumerate(cols) if any(t in normalizar(c) for t in ["EVENTO", "STATUS", "SITUACAO", "TIPO"])), None)
    return idx_nota, idx_data, idx_valor, idx_evento

# =========================================================
# PROCESSAMENTO DAS FONTES (COM PURGA DE INVALIDADAS)
# =========================================================
def processar_fonte(df, col_nota, col_data, col_valor, col_evento, usar_data, tipo_doc):
    res = pd.DataFrame()
    res['nota'] = df[col_nota].apply(lambda x: extrair_identificador_nota(x, tipo_doc))
    
    if col_data and col_data != "Nenhuma":
        res['data'] = df[col_data].apply(converter_data)
    else:
        res['data'] = None
        
    res['valor'] = df[col_valor].apply(limpar_valor)
    
    if col_evento and col_evento != "Nenhuma":
        res['invalida'] = df[col_evento].apply(verificar_status_invalido)
    else:
        res['invalida'] = False

    res = res[(res['nota'] != "") & (~res['invalida'])].copy()

    if usar_data:
        res = res.dropna(subset=['data'])
        chave = ['nota', 'data']
    else:
        chave = ['nota']

    agrupado = res.groupby(chave, as_index=False).agg({'valor': 'sum'})
    return agrupado

# =========================================================
# INTERFACE PRINCIPAL
# =========================================================
TIPOS_DOC = {
    "NFe Entrada": False,
    "NFe Saída": True,
    "NFCe": True,
    "CTe": True,
    "NFSe": True,
}

FONTES_DISPONIVEIS = {
    "Fonte Originária (SEFAZ/Prefeitura)": "Origem",
    "SIEG": "SIEG",
    "Domínio": "Dominio",
}

col_tipo, col_fontes = st.columns(2)
with col_tipo:
    tipo_doc = st.radio(
        "📄 **Tipo de documento fiscal:**", 
        list(TIPOS_DOC.keys()), 
        horizontal=False,
        key="radio_tipo_documento_main"
    )
    
    usar_data = st.checkbox(
        "Cruzar também pela DATA (além da Chave/Número)",
        value=TIPOS_DOC[tipo_doc],
        help="Habilitado por padrão para Saídas/Serviços. Para entradas, o ideal é desabilitar caso a data de emissão difira da data de entrada na Domínio.",
        key="chk_usar_data_main"
    )

with col_fontes:
    fontes_selecionadas = st.multiselect(
        "🔗 **Selecione as fontes para o cruzamento:** (mínimo 2)",
        list(FONTES_DISPONIVEIS.keys()),
        default=list(FONTES_DISPONIVEIS.keys()),
        key="multi_fontes_selecionadas_main"
    )

st.write("---")

if len(fontes_selecionadas) < 2:
    st.warning("Selecione ao menos 2 fontes para realizar o confronto.")
    st.stop()

dados_fontes = {}
n_fontes = len(fontes_selecionadas)
cols_layout = st.columns(n_fontes)

for i, fonte in enumerate(fontes_selecionadas):
    codigo = FONTES_DISPONIVEIS[fonte]
    with cols_layout[i]:
        st.markdown(f"**📊 {fonte}**")
        f_upload = st.file_uploader(f"Relatório — {fonte}", type=["xlsx", "xls", "csv"], key=f"upload_{codigo}")
        if f_upload:
            df_bruto = carregar_planilha(f_upload)
            if not df_bruto.empty:
                cols = list(df_bruto.columns)
                idx_nota, idx_data, idx_valor, idx_evento = sugerir_colunas(cols)
                opcoes_com_nenhuma = ["Nenhuma"] + cols

                col_nota = st.selectbox(
                    "Chave / Número NF", 
                    cols, 
                    index=idx_nota, 
                    key=f"sel_nota_{codigo}_{i}"
                )
                col_data = st.selectbox(
                    "Data" + (" (obrigatória)" if usar_data else " (opcional)"),
                    opcoes_com_nenhuma, 
                    index=idx_data + 1, 
                    key=f"sel_data_{codigo}_{i}"
                )
                col_valor = st.selectbox(
                    "Valor Total", 
                    cols, 
                    index=idx_valor, 
                    key=f"sel_valor_{codigo}_{i}"
                )
                col_evento = st.selectbox(
                    "Status / Situação (Filtra Canceladas)", 
                    opcoes_com_nenhuma,
                    index=(idx_evento + 1) if idx_evento is not None else 0, 
                    key=f"sel_evento_{codigo}_{i}"
                )

                dados_fontes[codigo] = {
                    "fonte": fonte, "df": df_bruto,
                    "col_nota": col_nota, "col_data": col_data,
                    "col_valor": col_valor, "col_evento": col_evento,
                }

st.write("---")

fontes_prontas = list(dados_fontes.keys())
if len(fontes_prontas) < 2:
    st.info("Carregue os arquivos e valide o mapeamento das colunas para liberar a conciliação.")
    st.stop()

if usar_data:
    faltando_data = [dados_fontes[c]["fonte"] for c in fontes_prontas if dados_fontes[c]["col_data"] == "Nenhuma"]
    if faltando_data:
        st.error(f"Selecione a coluna de data para as seguintes fontes: {', '.join(faltando_data)}.")
        st.stop()

if st.button("🚀 Processar Conciliação e Retornar Divergências", type="primary", use_container_width=True, key="btn_processar_conciliacao"):
    with st.spinner("Lendo relatórios, eliminando canceladas e cruzando bases..."):
        try:
            processados = {}
            for codigo in fontes_prontas:
                cfg = dados_fontes[codigo]
                agrupado = processar_fonte(
                    cfg["df"], cfg["col_nota"], cfg["col_data"], cfg["col_valor"], 
                    cfg["col_evento"], usar_data, tipo_doc
                )
                rename_map = {"valor": f"valor_{codigo}"}
                if not usar_data:
                    rename_map["data"] = f"data_{codigo}"
                agrupado = agrupado.rename(columns=rename_map)
                processados[codigo] = agrupado

            chave = ['nota', 'data'] if usar_data else ['nota']
            m = reduce(lambda l, r: pd.merge(l, r, on=chave, how='outer'), processados.values())

            st.subheader("📊 Totais Válidos Consolidados (Sem Canceladas)")
            metricas = st.columns(len(fontes_prontas))
            for i, codigo in enumerate(fontes_prontas):
                total = m[f"valor_{codigo}"].fillna(0).sum()
                with metricas[i]:
                    st.metric(f"Total Válido — {dados_fontes[codigo]['fonte']}", formatar_moeda_br(total))

            def analisar_divergencia(row):
                ausentes = [c for c in fontes_prontas if pd.isna(row[f"valor_{c}"])]
                presentes = {c: row[f"valor_{c}"] for c in fontes_prontas if c not in ausentes}
                situacoes = []
                
                if ausentes:
                    nomes_ausentes = [dados_fontes[c]["fonte"] for c in ausentes]
                    situacoes.append("Falta em: " + ", ".join(nomes_ausentes))
                
                if len(presentes) >= 2:
                    vals = list(presentes.values())
                    if max(vals) - min(vals) > 0.01:
                        situacoes.append("Divergência de Valor")

                return pd.Series({
                    "Situação": " | ".join(situacoes),
                    "_divergente": bool(situacoes)
                })

            extras = m.apply(analisar_divergencia, axis=1)
            m = pd.concat([m, extras], axis=1)
            divergencias = m[m["_divergente"]].copy()

            if not usar_data:
                data_ref_cols = [f"data_{c}" for c in fontes_prontas if f"data_{c}" in divergencias.columns]
                if data_ref_cols:
                    divergencias["Data (Ref)"] = divergencias[data_ref_cols].bfill(axis=1).iloc[:, 0]

            divergencias = divergencias.sort_values(by=chave)

            st.write("---")
            st.subheader("🔍 Relatório Exclusivo de Divergências")

            if divergencias.empty:
                st.success("✅ **Conciliação sem divergências:** Todas as notas válidas presentes nas bases possuem os mesmos valores.")
            else:
                st.warning(f"Foram encontradas **{len(divergencias)}** inconsistências entre os relatórios.")

                exib = pd.DataFrame()
                exib["Chave / Número NF"] = divergencias["nota"]
                if usar_data:
                    exib["Data"] = pd.to_datetime(divergencias["data"]).dt.strftime('%d/%m/%Y')
                elif "Data (Ref)" in divergencias.columns:
                    exib["Data (Ref)"] = pd.to_datetime(divergencias["Data (Ref)"], errors='coerce').dt.strftime('%d/%m/%Y')

                for c in fontes_prontas:
                    exib[f"Valor - {dados_fontes[c]['fonte']}"] = divergencias[f"valor_{c}"].apply(
                        lambda v: "— (Ausente)" if pd.isna(v) else formatar_moeda_br(v)
                    )

                exib["Motivo da Inconsistência"] = divergencias["Situação"]
                exib = exib.reset_index(drop=True)

                st.dataframe(exib, use_container_width=True)

                export_df = pd.DataFrame()
                export_df["Chave_Numero_NF"] = divergencias["nota"]
                if usar_data:
                    export_df["Data"] = pd.to_datetime(divergencias["data"]).dt.strftime('%d/%m/%Y')
                elif "Data (Ref)" in divergencias.columns:
                    export_df["Data"] = pd.to_datetime(divergencias["Data (Ref)"], errors='coerce').dt.strftime('%d/%m/%Y')

                for c in fontes_prontas:
                    export_df[f"Valor_{dados_fontes[c]['fonte']}"] = divergencias[f"valor_{c}"]

                export_df["Motivo_Inconsistencia"] = divergencias["Situação"]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Divergências')

                st.download_button(
                    label="📥 Baixar Planilha de Divergências (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"divergencias_conciliacao_{tipo_doc.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_excel"
                )
        except Exception as e:
            st.error(f"Erro no processamento da conciliação: {e}")
