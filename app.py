import streamlit as st
import pandas as pd
import re
import unicodedata
import io
from functools import reduce

st.set_page_config(page_title="Conciliador Fiscal Universal", layout="wide")
st.title("⚖️ Conciliador Fiscal Universal — Multi-Fonte")

# =========================================================
# UTILITÁRIOS
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

def extrair_nota_limpa(n):
    if pd.isna(n): return ""
    s = str(n).strip()
    s = re.sub(r'\D', '', s)
    if len(s) == 44: s = s[25:34]
    return s.lstrip('0') if s else ""

def extrair_ultimo_evento(txt):
    """Retorna o status quando a nota está Cancelada / Desc. Op. Denegada. Não filtra, só sinaliza."""
    if pd.isna(txt): return ""
    s = str(txt).strip()
    if not s: return ""
    ultimo = s.split(',')[-1].strip()
    teste = normalizar(ultimo)
    if "DESC" in teste or "CANCEL" in teste or "DENEG" in teste:
        return ultimo
    return ""

# =========================================================
# DETECTOR DE CABEÇALHO
# =========================================================
def encontrar_cabecalho(df):
    termos_fortes = ["CHAVE", "NOTA", "DATA", "VALOR", "EMISSAO", "NUMERO", "NUM NFSE", "CNPJ", "EVENTO"]
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

# =========================================================
# MOTOR DE LEITURA (mantido do original — cobre CSV, XLS/XLSX, HTML disfarçado)
# =========================================================
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
                st.error(f"🛑 **ARQUIVO CORROMPIDO NA ORIGEM:** O arquivo '{f.name}' exportado pelo sistema possui defeitos na matriz binária. \n\nPara garantir a integridade da conciliação, abra este arquivo no Excel do seu computador e selecione **'Salvar Como -> Pasta de Trabalho do Excel (.xlsx)'** antes de enviar.")
                return pd.DataFrame()
        else:
            try:
                dfs = pd.read_html(io.BytesIO(conteudo))
                df = dfs[0].astype(str)
            except Exception:
                try: texto = conteudo.decode('utf-8')
                except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
                if '\t' in texto: sep = '\t'
                elif ';' in texto: sep = ';'
                else: sep = ','
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
    idx_evento = next((i for i, c in enumerate(cols) if ("EVENTO" in normalizar(c) and "CODIGO" not in normalizar(c)) or ("TIPO" in normalizar(c) and "EVENTO" in normalizar(c)) or "STATUS" in normalizar(c) or "SITUACAO" in normalizar(c)), None)
    return idx_nota, idx_data, idx_valor, idx_evento

# =========================================================
# PROCESSAMENTO POR FONTE
# =========================================================
def processar_fonte(df, col_nota, col_data, col_valor, col_evento, usar_data):
    res = pd.DataFrame()
    res['nota'] = df[col_nota].apply(extrair_nota_limpa)
    if col_data and col_data != "Nenhuma":
        res['data'] = df[col_data].apply(converter_data)
    else:
        res['data'] = None
    res['valor'] = df[col_valor].apply(limpar_valor)
    if col_evento and col_evento != "Nenhuma":
        res['evento'] = df[col_evento].apply(extrair_ultimo_evento)
    else:
        res['evento'] = ""

    res = res[res['nota'] != ""]

    if usar_data:
        res = res.dropna(subset=['data'])
        chave = ['nota', 'data']
    else:
        chave = ['nota']

    agrupado = res.groupby(chave, as_index=False).agg({'valor': 'sum', 'evento': 'last'})
    return agrupado

# =========================================================
# INTERFACE
# =========================================================
st.info("💡 **Atenção:** Arquivos defeituosos de alguns sistemas exigem reparo no Excel (Salvar Como .xlsx).")

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
    tipo_doc = st.radio("📄 **Tipo de documento fiscal:**", list(TIPOS_DOC.keys()), horizontal=False)
    usar_data = st.checkbox(
        "Cruzar também pela DATA (além do número da nota)",
        value=TIPOS_DOC[tipo_doc],
        help="Ligado por padrão para Saída/NFCe/CTe/NFSe, desligado para Entrada — mas você pode ajustar."
    )

with col_fontes:
    fontes_selecionadas = st.multiselect(
        "🔗 **Quais fontes você vai subir nesta rodada?** (mín. 2)",
        list(FONTES_DISPONIVEIS.keys()),
        default=list(FONTES_DISPONIVEIS.keys())
    )

st.write("---")

if len(fontes_selecionadas) < 2:
    st.warning("Selecione ao menos 2 fontes para cruzar.")
    st.stop()

# --- Upload e mapeamento de colunas por fonte ---
dados_fontes = {}  # codigo -> dict com df bruto e colunas escolhidas

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

                col_nota = st.selectbox("Nota/Chave", cols, index=idx_nota, key=f"nota_{codigo}")
                col_data = st.selectbox(
                    "Data" + (" (obrigatória)" if usar_data else " (opcional, só referência)"),
                    opcoes_com_nenhuma, index=idx_data + 1, key=f"data_{codigo}"
                )
                col_valor = st.selectbox("Valor", cols, index=idx_valor, key=f"valor_{codigo}")
                col_evento = st.selectbox(
                    "Status/Evento (opcional)", opcoes_com_nenhuma,
                    index=(idx_evento + 1) if idx_evento is not None else 0, key=f"evento_{codigo}"
                )

                dados_fontes[codigo] = {
                    "fonte": fonte, "df": df_bruto,
                    "col_nota": col_nota, "col_data": col_data,
                    "col_valor": col_valor, "col_evento": col_evento,
                }

st.write("---")

fontes_prontas = list(dados_fontes.keys())
if len(fontes_prontas) < 2:
    st.info("Suba pelo menos 2 relatórios com as colunas mapeadas para liberar o cruzamento.")
    st.stop()

if usar_data:
    faltando_data = [dados_fontes[c]["fonte"] for c in fontes_prontas if dados_fontes[c]["col_data"] == "Nenhuma"]
    if faltando_data:
        st.error(f"Você marcou 'cruzar por data', mas não selecionou coluna de data para: {', '.join(faltando_data)}.")
        st.stop()

if st.button("🚀 Cruzar Dados e Buscar Divergências", type="primary", use_container_width=True):
    with st.spinner("Cruzando informações fiscais..."):
        try:
            processados = {}
            for codigo in fontes_prontas:
                cfg = dados_fontes[codigo]
                agrupado = processar_fonte(cfg["df"], cfg["col_nota"], cfg["col_data"], cfg["col_valor"], cfg["col_evento"], usar_data)
                rename_map = {"valor": f"valor_{codigo}", "evento": f"evento_{codigo}"}
                if not usar_data:
                    rename_map["data"] = f"data_{codigo}"
                agrupado = agrupado.rename(columns=rename_map)
                processados[codigo] = agrupado

            chave = ['nota', 'data'] if usar_data else ['nota']
            m = reduce(lambda l, r: pd.merge(l, r, on=chave, how='outer'), processados.values())

            # --- Totais consolidados (antes de filtrar por divergência) ---
            st.subheader("📊 Totais Consolidados")
            metricas = st.columns(len(fontes_prontas))
            for i, codigo in enumerate(fontes_prontas):
                total = m[f"valor_{codigo}"].fillna(0).sum()
                with metricas[i]:
                    st.metric(f"Soma {dados_fontes[codigo]['fonte']}", formatar_moeda_br(total))

            # --- Análise linha a linha: ausências e divergência de valor ---
            def analisar_linha(row):
                ausentes = [c for c in fontes_prontas if pd.isna(row[f"valor_{c}"])]
                presentes = {c: row[f"valor_{c}"] for c in fontes_prontas if c not in ausentes}
                situacoes = []
                if ausentes:
                    nomes_ausentes = [dados_fontes[c]["fonte"] for c in ausentes]
                    situacoes.append("Ausente em: " + ", ".join(nomes_ausentes))
                if len(presentes) >= 2:
                    vals = list(presentes.values())
                    if max(vals) - min(vals) > 0.01:
                        situacoes.append("Valor divergente")

                status_partes = []
                for c in fontes_prontas:
                    ev = row.get(f"evento_{c}", "")
                    if isinstance(ev, str) and ev.strip():
                        status_partes.append(f"{dados_fontes[c]['fonte']}: {ev.strip()}")

                return pd.Series({
                    "Situação": " | ".join(situacoes),
                    "Status (Cancelamento/Denegação)": " | ".join(status_partes),
                    "_divergente": bool(situacoes)
                })

            extras = m.apply(analisar_linha, axis=1)
            m = pd.concat([m, extras], axis=1)
            divergencias = m[m["_divergente"]].copy()

            # --- Data de referência quando a chave não usa data ---
            if not usar_data:
                data_ref_cols = [f"data_{c}" for c in fontes_prontas if f"data_{c}" in divergencias.columns]
                if data_ref_cols:
                    divergencias["Data (referência)"] = divergencias[data_ref_cols].bfill(axis=1).iloc[:, 0]

            divergencias = divergencias.sort_values(by=chave)

            st.write("---")
            st.subheader("🔍 Divergências Encontradas")

            if divergencias.empty:
                st.success("🎉 Excelente! As fontes bateram — nenhuma divergência encontrada.")
            else:
                st.warning(f"Foram identificadas {len(divergencias)} notas com inconsistências.")

                # Montagem da tabela de exibição (moeda formatada)
                exib = pd.DataFrame()
                exib["Número da Nota"] = divergencias["nota"]
                if usar_data:
                    exib["Data"] = pd.to_datetime(divergencias["data"]).dt.strftime('%d/%m/%Y')
                elif "Data (referência)" in divergencias.columns:
                    exib["Data (referência)"] = pd.to_datetime(divergencias["Data (referência)"], errors='coerce').dt.strftime('%d/%m/%Y')

                for c in fontes_prontas:
                    exib[f"Valor - {dados_fontes[c]['fonte']}"] = divergencias[f"valor_{c}"].apply(
                        lambda v: "— (ausente)" if pd.isna(v) else formatar_moeda_br(v)
                    )

                exib["Situação"] = divergencias["Situação"]
                exib["Status (Cancelamento/Denegação)"] = divergencias["Status (Cancelamento/Denegação)"]
                exib = exib.reset_index(drop=True)

                st.dataframe(exib, use_container_width=True)

                # Exportação em Excel (valores numéricos, não formatados, para o usuário trabalhar em cima)
                export_df = pd.DataFrame()
                export_df["Número da Nota"] = divergencias["nota"]
                if usar_data:
                    export_df["Data"] = pd.to_datetime(divergencias["data"]).dt.strftime('%d/%m/%Y')
                elif "Data (referência)" in divergencias.columns:
                    export_df["Data (referência)"] = pd.to_datetime(divergencias["Data (referência)"], errors='coerce').dt.strftime('%d/%m/%Y')
                for c in fontes_prontas:
                    export_df[f"Valor - {dados_fontes[c]['fonte']}"] = divergencias[f"valor_{c}"]
                export_df["Situação"] = divergencias["Situação"]
                export_df["Status (Cancelamento/Denegação)"] = divergencias["Status (Cancelamento/Denegação)"]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Divergências')

                nome_arquivo = f"divergencias_{tipo_doc.replace(' ', '_')}.xlsx"
                st.download_button(
                    label="📥 Baixar Planilha de Divergências",
                    data=output.getvalue(),
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro processual: {e}")
