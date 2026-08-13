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
    if pd.isna(txt): return ""
    s = str(txt).strip()
    if not s: return ""
    ultimo = s.split(',')[-1].strip()
    teste = normalizar(ultimo)
    if "DESC" in teste or "CANCEL" in teste or "DENEG" in teste:
        return ultimo
    return ""

# =========================================================
# DETECTOR DE CABEÇALHO E METADADOS
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

def valor_esta_vazio(v):
    if pd.isna(v): return True
    s = str(v).replace('R$', '').replace('"', '').replace('\xa0', '').replace(' ', '').strip()
    return s == ""

def extrair_metadados(df_raw, idx_cabecalho):
    metadados = {"empresa": None, "competencia": None}
    limite = min(idx_cabecalho, len(df_raw))
    for i in range(limite):
        linha = df_raw.iloc[i]
        for j, cel in enumerate(linha):
            if pd.isna(cel): continue
            texto = str(cel).strip()
            if not texto: continue
            texto_norm = normalizar(texto)

            if metadados["empresa"] is None and "EMPRESA" in texto_norm:
                if ':' in texto:
                    valor = texto.split(':', 1)[1].strip()
                    if valor:
                        metadados["empresa"] = valor
                        continue
                if j + 1 < len(linha) and pd.notna(linha.iloc[j + 1]):
                    valor = str(linha.iloc[j + 1]).strip()
                    if valor and "EMPRESA" not in normalizar(valor):
                        metadados["empresa"] = valor

            if metadados["competencia"] is None and any(t in texto_norm for t in ["COMPETENCIA", "PERIODO"]):
                if ':' in texto:
                    valor = texto.split(':', 1)[1].strip()
                    if valor:
                        metadados["competencia"] = valor
                        continue
                if j + 1 < len(linha) and pd.notna(linha.iloc[j + 1]):
                    valor = str(linha.iloc[j + 1]).strip()
                    if valor:
                        metadados["competencia"] = valor
    return metadados

# =========================================================
# MOTOR DE LEITURA
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
                st.error(f"🛑 **ARQUIVO CORROMPIDO NA ORIGEM:** O arquivo '{f.name}' possui defeitos na estrutura binária. Abra no Excel e salve como '.xlsx'.")
                return pd.DataFrame(), {}
        else:
            try:
                dfs = pd.read_html(io.BytesIO(conteudo))
                df = dfs[0].astype(str)
            except Exception:
                try: texto = conteudo.decode('utf-8')
                except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
                sep = '\t' if '\t' in texto else (';' if ';' in texto else ',')
                df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, header=None, engine='python', on_bad_lines='skip')

    if df is None or df.empty: return pd.DataFrame(), {}
    idx_cabecalho = encontrar_cabecalho(df)
    metadados = extrair_metadados(df, idx_cabecalho)
    if idx_cabecalho >= len(df): return df, metadados

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
    return df, metadados

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

    res['_valor_vazio'] = df[col_valor].apply(valor_esta_vazio)
    res['valor'] = df[col_valor].apply(limpar_valor)

    if col_evento and col_evento != "Nenhuma":
        res['evento'] = df[col_evento].apply(extrair_ultimo_evento)
    else:
        res['evento'] = ""

    # Filtra linhas sem nota, com valor totalmente vazio ou zerado (R$ 0,00)
    res = res[res['nota'] != ""]
    res = res[~res['_valor_vazio']]
    res = res[res['valor'] > 0.0]

    # Identificação de notas duplicadas antes de agrupar
    chave = ['nota', 'data'] if usar_data else ['nota']
    duplicados = res.duplicated(subset=chave, keep=False)
    notas_duplicadas = set(res[duplicados]['nota'].tolist())

    agrupado = res.groupby(chave, as_index=False).agg({'valor': 'sum', 'evento': 'last'})
    return agrupado, notas_duplicadas

def detectar_empresa(df):
    candidatos = [c for c in df.columns if any(t in normalizar(c) for t in
                  ["RAZAO SOCIAL", "NOME EMPRESA", "NOME DA EMPRESA", "EMITENTE", "EMPRESA"])
                  and "CODIGO" not in normalizar(c) and "CNPJ" not in normalizar(c)]
    for c in candidatos:
        valores = df[c].dropna().astype(str).str.strip()
        valores = valores[valores != ""]
        if not valores.empty:
            moda = valores.mode()
            if not moda.empty: return moda.iloc[0]
    return None

def detectar_competencia(df, col_data):
    if col_data and col_data != "Nenhuma":
        datas = df[col_data].apply(converter_data).dropna()
        if not datas.empty:
            meses = datas.apply(lambda d: f"{d.month:02d}/{d.year}")
            moda = meses.mode()
            if not moda.empty: return moda.iloc[0]
    candidatos = [c for c in df.columns if any(t in normalizar(c) for t in ["COMPETENCIA", "PERIODO"])]
    for c in candidatos:
        valores = df[c].dropna().astype(str).str.strip()
        valores = valores[valores != ""]
        if not valores.empty:
            moda = valores.mode()
            if not moda.empty: return moda.iloc[0]
    return None

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
    "Oficial (SEFAZ/Prefeitura)": "Oficial",
    "SIEG": "SIEG",
    "Domínio": "Dominio",
}

NOME_COLUNA_VALOR = {
    "Oficial": "Valor Oficial",
    "SIEG": "Valor SIEG",
    "Dominio": "Valor Domínio",
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
        "Cruzar também pela DATA (além do número da nota)",
        value=TIPOS_DOC[tipo_doc],
        key="chk_usar_data_main"
    )
    mostrar_data = tipo_doc != "NFe Entrada"

with col_fontes:
    fontes_selecionadas = st.multiselect(
        "🔗 **Quais fontes você vai subir nesta rodada?** (mín. 2)",
        list(FONTES_DISPONIVEIS.keys()),
        default=list(FONTES_DISPONIVEIS.keys()),
        key="multi_fontes_selecionadas_main"
    )

st.write("---")

if len(fontes_selecionadas) < 2:
    st.warning("Selecione ao menos 2 fontes para cruzar.")
    st.stop()

# --- Upload e mapeamento ---
dados_fontes = {}
cols_layout = st.columns(len(fontes_selecionadas))

for i, fonte in enumerate(fontes_selecionadas):
    codigo = FONTES_DISPONIVEIS[fonte]
    with cols_layout[i]:
        st.markdown(f"**📊 {fonte}**")
        f_upload = st.file_uploader(f"Relatório — {fonte}", type=["xlsx", "xls", "csv"], key=f"upload_{codigo}")
        if f_upload:
            df_bruto, metadados = carregar_planilha(f_upload)
            if not df_bruto.empty:
                cols = list(df_bruto.columns)
                idx_nota, idx_data, idx_valor, idx_evento = sugerir_colunas(cols)
                opcoes_com_nenhuma = ["Nenhuma"] + cols

                col_nota = st.selectbox("Nota/Chave", cols, index=idx_nota, key=f"nota_{codigo}_{i}")
                col_data = st.selectbox(
                    "Data" + (" (obrigatória)" if usar_data else " (opcional)"),
                    opcoes_com_nenhuma, index=idx_data + 1, key=f"data_{codigo}_{i}"
                )
                col_valor = st.selectbox("Valor", cols, index=idx_valor, key=f"valor_{codigo}_{i}")
                col_evento = st.selectbox(
                    "Status/Evento (opcional)", opcoes_com_nenhuma,
                    index=(idx_evento + 1) if idx_evento is not None else 0, key=f"evento_{codigo}_{i}"
                )

                dados_fontes[codigo] = {
                    "fonte": fonte, "df": df_bruto,
                    "col_nota": col_nota, "col_data": col_data,
                    "col_valor": col_valor, "col_evento": col_evento,
                    "empresa_detectada": metadados.get("empresa") or detectar_empresa(df_bruto),
                    "competencia_detectada": metadados.get("competencia") or detectar_competencia(df_bruto, col_data),
                }

st.write("---")

# --- Empresa e Competência (Prioridade Domínio) ---
empresa_sugerida = None
competencia_sugerida = None

if "Dominio" in dados_fontes and dados_fontes["Dominio"].get("empresa_detectada"):
    empresa_sugerida = dados_fontes["Dominio"]["empresa_detectada"]
else:
    empresa_sugerida = next((d["empresa_detectada"] for d in dados_fontes.values() if d.get("empresa_detectada")), None)

if "Dominio" in dados_fontes and dados_fontes["Dominio"].get("competencia_detectada"):
    competencia_sugerida = dados_fontes["Dominio"]["competencia_detectada"]
else:
    competencia_sugerida = next((d["competencia_detectada"] for d in dados_fontes.values() if d.get("competencia_detectada")), None)

if empresa_sugerida and not st.session_state.get("empresa_input"):
    st.session_state["empresa_input"] = empresa_sugerida
if competencia_sugerida and not st.session_state.get("competencia_input"):
    st.session_state["competencia_input"] = competencia_sugerida

col_emp, col_comp = st.columns(2)
with col_emp:
    empresa = st.text_input("🏢 Empresa", key="empresa_input")
with col_comp:
    competencia = st.text_input("🗓️ Competência (ex: 08/2026)", key="competencia_input")

st.write("---")

fontes_prontas = list(dados_fontes.keys())
if len(fontes_prontas) < 2:
    st.info("Suba pelo menos 2 relatórios para liberar o cruzamento.")
    st.stop()

if st.button("🚀 Cruzar Dados e Buscar Divergências", type="primary", use_container_width=True, key="btn_cruzar_dados_main"):
    with st.spinner("Cruzando informações fiscais..."):
        try:
            processados = {}
            duplicadas_por_fonte = {}

            for codigo in fontes_prontas:
                cfg = dados_fontes[codigo]
                agrupado, dupes = processar_fonte(cfg["df"], cfg["col_nota"], cfg["col_data"], cfg["col_valor"], cfg["col_evento"], usar_data)
                duplicadas_por_fonte[codigo] = dupes

                rename_map = {"valor": f"valor_{codigo}", "evento": f"evento_{codigo}"}
                if not usar_data: rename_map["data"] = f"data_{codigo}"
                agrupado = agrupado.rename(columns=rename_map)
                processados[codigo] = agrupado

            chave = ['nota', 'data'] if usar_data else ['nota']
            m = reduce(lambda l, r: pd.merge(l, r, on=chave, how='outer'), processados.values())

            # --- Totais Consolidados na Mesma Linha ---
            st.subheader("📊 Totais Consolidados")
            if empresa or competencia:
                st.markdown(f"### **{empresa or 'Empresa não informada'}** | Competência: {competencia or '—'} | {tipo_doc}")

            totais = {codigo: m[f"valor_{codigo}"].fillna(0).sum() for codigo in fontes_prontas}

            num_cols = len(fontes_prontas) + (1 if len(fontes_prontas) == 2 else 0)
            metricas = st.columns(num_cols)

            for i, codigo in enumerate(fontes_prontas):
                with metricas[i]:
                    st.metric(f"Soma {dados_fontes[codigo]['fonte']}", formatar_moeda_br(totais[codigo]))

            if len(fontes_prontas) == 2:
                codigo_a, codigo_b = fontes_prontas
                diferenca_total = totais[codigo_a] - totais[codigo_b]
                with metricas[-1]:
                    st.metric(
                        "Diferença Total",
                        formatar_moeda_br(diferenca_total),
                        delta=f"{diferenca_total:,.2f} R$" if abs(diferenca_total) > 0.01 else None,
                        delta_color="inverse" if abs(diferenca_total) > 0.01 else "normal"
                    )

            # --- Resumo em Tabela (Apenas para 3 fontes) ---
            resumo_df = None
            if len(fontes_prontas) == 3:
                resumo_linhas = [
                    {
                        "Comparação": f"{dados_fontes['Oficial']['fonte']} vs {dados_fontes['Dominio']['fonte']}",
                        "Diferença": totais['Oficial'] - totais['Dominio']
                    },
                    {
                        "Comparação": f"{dados_fontes['SIEG']['fonte']} vs {dados_fontes['Dominio']['fonte']}",
                        "Diferença": totais['SIEG'] - totais['Dominio']
                    },
                    {
                        "Comparação": f"{dados_fontes['Oficial']['fonte']} vs {dados_fontes['SIEG']['fonte']}",
                        "Diferença": totais['Oficial'] - totais['SIEG']
                    }
                ]
                resumo_df = pd.DataFrame(resumo_linhas)
                resumo_exib = resumo_df.copy()
                resumo_exib["Diferença"] = resumo_exib["Diferença"].apply(formatar_moeda_br)

                st.markdown("**Resumo Comparativo Par a Par**")
                st.dataframe(resumo_exib, use_container_width=True, hide_index=True)

            # --- Análise de Inconsistências ---
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
                        situacoes.append("Valor divergente entre fontes")

                nota_atual = row['nota']
                for c in fontes_prontas:
                    if nota_atual in duplicadas_por_fonte.get(c, set()):
                        situacoes.append(f"Nota duplicada no {dados_fontes[c]['fonte']}")

                for c in fontes_prontas:
                    ev = row.get(f"evento_{c}", "")
                    if isinstance(ev, str) and ev.strip():
                        situacoes.append(f"Evento ({dados_fontes[c]['fonte']}): {ev.strip()}")

                return pd.Series({
                    "Motivo da Inconsistência": " | ".join(situacoes),
                    "_divergente": bool(situacoes)
                })

            extras = m.apply(analisar_linha, axis=1)
            m = pd.concat([m, extras], axis=1)
            divergencias = m[m["_divergente"]].copy()

            if mostrar_data:
                if usar_data:
                    divergencias["Data"] = divergencias["data"]
                else:
                    data_ref_cols = [f"data_{c}" for c in fontes_prontas if f"data_{c}" in divergencias.columns]
                    divergencias["Data"] = divergencias[data_ref_cols].bfill(axis=1).iloc[:, 0] if data_ref_cols else None

            divergencias = divergencias.sort_values(by=chave)

            st.write("---")
            st.subheader("🔍 Divergências Encontradas")

            if divergencias.empty:
                st.success("🎉 Nenhuma divergência encontrada! As fontes conferem perfeitamente.")
            else:
                st.warning(f"Identificadas {len(divergencias)} notas com inconsistências.")

                exib = pd.DataFrame()
                exib["Número da Nota"] = divergencias["nota"]
                if mostrar_data:
                    exib["Data"] = pd.to_datetime(divergencias["Data"], errors='coerce').dt.strftime('%d/%m/%Y')

                for c in fontes_prontas:
                    exib[NOME_COLUNA_VALOR[c]] = divergencias[f"valor_{c}"].apply(
                        lambda v: "— (ausente)" if pd.isna(v) else formatar_moeda_br(v)
                    )

                exib["Motivo da Inconsistência"] = divergencias["Motivo da Inconsistência"]
                exib = exib.reset_index(drop=True)

                st.dataframe(exib, use_container_width=True)

                # Exportação limpa
                export_df = pd.DataFrame()
                export_df["Número da Nota"] = divergencias["nota"].values
                if mostrar_data:
                    export_df["Data"] = pd.to_datetime(divergencias["Data"], errors='coerce').dt.strftime('%d/%m/%Y').values

                for c in fontes_prontas:
                    export_df[NOME_COLUNA_VALOR[c]] = divergencias[f"valor_{c}"].values

                export_df["Motivo da Inconsistência"] = divergencias["Motivo da Inconsistência"].values

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Divergências')
                    if resumo_df is not None:
                        resumo_df.to_excel(writer, index=False, sheet_name='Resumo')

                nome_arquivo = f"divergencias_{(empresa or 'relatorio').replace(' ', '_')}_{competencia.replace('/', '-') if competencia else 'geral'}.xlsx"
                st.download_button(
                    label="📥 Baixar Planilha de Divergências (Excel)",
                    data=output.getvalue(),
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_excel_main"
                )
        except Exception as e:
            st.error(f"Erro processual: {e}")
