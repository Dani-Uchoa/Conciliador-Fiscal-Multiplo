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
    if '/' in s:
        # Formato "8221/8221" (mesmo número repetido) — usa a primeira parte
        partes = [p.strip() for p in s.split('/') if p.strip()]
        if partes: s = partes[0]
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

def valor_esta_vazio(v):
    """Distingue célula vazia/em branco de um valor real 0,00."""
    if pd.isna(v): return True
    s = str(v).replace('R$', '').replace('"', '').replace('\xa0', '').replace(' ', '').strip()
    return s == ""

def extrair_metadados(df_raw, idx_cabecalho):
    """Varre as linhas ANTES do cabeçalho da tabela em busca de 'Empresa:' e 'Competência:'/'Período:'.
    O valor nem sempre está na célula vizinha (colunas mescladas deixam NaN no meio),
    então pega a próxima célula NÃO VAZIA da mesma linha."""
    metadados = {"empresa": None, "competencia": None}
    limite = min(idx_cabecalho, len(df_raw))

    for i in range(limite):
        linha = df_raw.iloc[i]
        celulas = [(j, str(v).strip()) for j, v in enumerate(linha) if pd.notna(v) and str(v).strip() != ""]

        for pos, (j, texto) in enumerate(celulas):
            texto_norm = normalizar(texto)
            rotulo = texto_norm.replace(':', '').strip()

            if metadados["empresa"] is None and rotulo == "EMPRESA":
                if pos + 1 < len(celulas):
                    valor = celulas[pos + 1][1]
                    if valor:
                        metadados["empresa"] = valor

            if metadados["competencia"] is None and rotulo in ["COMPETENCIA", "PERIODO"]:
                if pos + 1 < len(celulas):
                    valor = celulas[pos + 1][1]
                    if valor:
                        metadados["competencia"] = valor

    # Heurística: quando não há rótulo "Empresa:" explícito, alguns relatórios (ex: Domínio Saídas)
    # trazem a razão social sozinha na primeira célula da primeira linha.
    if metadados["empresa"] is None and len(df_raw) > 0:
        primeira_celula = df_raw.iloc[0, 0]
        if pd.notna(primeira_celula):
            texto = str(primeira_celula).strip()
            texto_norm = normalizar(texto)
            termos_societarios = {"LTDA", "S/A", "S.A", "EIRELI", "ME", "EPP", "MEI"}
            if texto and texto == texto.upper() and termos_societarios & set(texto_norm.replace('.', ' ').split()):
                metadados["empresa"] = texto

    return metadados

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
                return pd.DataFrame(), {}
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
def processar_fonte(df, col_nota, col_data, col_valor, col_evento, usar_data, desconsiderar_zero=False):
    """Processa uma fonte e retorna (agrupado, notas_duplicadas).
    notas_duplicadas: notas que apareceram mais de uma vez com o MESMO valor antes do agrupamento
    (indício de linha duplicada no arquivo original, não de múltiplos lançamentos legítimos).
    """
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

    res = res[res['nota'] != ""]

    # Célula vazia sempre é descartada. Para a Domínio, valor 0,00 também é descartado
    # (lançamento cancelado no sistema não deve virar divergência falsa).
    descartar = res['_valor_vazio']
    if desconsiderar_zero:
        descartar = descartar | (res['valor'].abs() < 0.005)
    res = res[~descartar].drop(columns=['_valor_vazio'])

    if usar_data:
        res = res.dropna(subset=['data'])
        chave = ['nota', 'data']
    else:
        chave = ['nota']

    # Duplicidade: mesma NOTA e mesmo VALOR repetidos no arquivo original — independe da data,
    # já que notas com o mesmo número mas CFOP/valor diferentes são splits legítimos, não duplicidade.
    dup_mask = res.duplicated(subset=['nota', 'valor'], keep=False)
    notas_duplicadas = set(res.loc[dup_mask, 'nota'])

    agrupado = res.groupby(chave, as_index=False).agg({'valor': 'sum', 'evento': 'last'})
    return agrupado, notas_duplicadas

def detectar_empresa(df):
    """Procura coluna de razão social/nome da empresa e retorna o valor mais frequente."""
    candidatos = [c for c in df.columns if any(t in normalizar(c) for t in
                  ["RAZAO SOCIAL", "NOME EMPRESA", "NOME DA EMPRESA", "EMITENTE", "EMPRESA"])
                  and "CODIGO" not in normalizar(c) and "CNPJ" not in normalizar(c)]
    for c in candidatos:
        valores = df[c].dropna().astype(str).str.strip()
        valores = valores[valores != ""]
        if not valores.empty:
            moda = valores.mode()
            if not moda.empty:
                return moda.iloc[0]
    return None

def detectar_competencia(df, col_data):
    """Deriva mês/ano predominante a partir da coluna de data, ou de uma coluna explícita de competência."""
    if col_data and col_data != "Nenhuma":
        datas = df[col_data].apply(converter_data).dropna()
        if not datas.empty:
            meses = datas.apply(lambda d: f"{d.month:02d}/{d.year}")
            moda = meses.mode()
            if not moda.empty:
                return moda.iloc[0]
    candidatos = [c for c in df.columns if any(t in normalizar(c) for t in ["COMPETENCIA", "PERIODO"])]
    for c in candidatos:
        valores = df[c].dropna().astype(str).str.strip()
        valores = valores[valores != ""]
        if not valores.empty:
            moda = valores.mode()
            if not moda.empty:
                return moda.iloc[0]
    return None

def obter_empresa_sugerida(dados_fontes):
    """Prioriza o cabeçalho/metadados da Domínio, depois metadados de outras fontes,
    depois detecção por coluna (Domínio primeiro)."""
    ordem = (["Dominio"] if "Dominio" in dados_fontes else []) + [c for c in dados_fontes if c != "Dominio"]
    for codigo in ordem:
        meta = dados_fontes[codigo].get("metadados") or {}
        if meta.get("empresa"):
            return meta["empresa"]
    for codigo in ordem:
        if dados_fontes[codigo].get("empresa_detectada"):
            return dados_fontes[codigo]["empresa_detectada"]
    return None

def obter_competencia_sugerida(dados_fontes):
    ordem = (["Dominio"] if "Dominio" in dados_fontes else []) + [c for c in dados_fontes if c != "Dominio"]
    for codigo in ordem:
        meta = dados_fontes[codigo].get("metadados") or {}
        if meta.get("competencia"):
            return meta["competencia"]
    for codigo in ordem:
        if dados_fontes[codigo].get("competencia_detectada"):
            return dados_fontes[codigo]["competencia_detectada"]
    return None

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
    "Fonte Oficial (SEFAZ/Prefeitura)": "Origem",
    "SIEG": "SIEG",
    "Domínio": "Dominio",
}

# Nome fixo de coluna de valor por fonte, para o relatório final
NOME_COLUNA_VALOR = {
    "Origem": "Valor Fonte Oficial",
    "SIEG": "Valor SIEG",
    "Dominio": "Valor Domínio",
}

col_tipo, col_fontes = st.columns(2)
with col_tipo:
    tipo_doc = st.radio("📄 **Tipo de documento fiscal:**", list(TIPOS_DOC.keys()), horizontal=False)
    usar_data = TIPOS_DOC[tipo_doc]
    mostrar_data = tipo_doc != "NFe Entrada"
    st.caption(
        "Cruzamento automático: **só por número da nota** para Entrada; "
        "**nota + data** para os demais tipos."
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
            df_bruto, metadados = carregar_planilha(f_upload)
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

                # Status/Evento é detectado automaticamente pela varredura das colunas —
                # só entra no motivo da inconsistência se realmente existir.
                col_evento = cols[idx_evento] if idx_evento is not None else "Nenhuma"

                dados_fontes[codigo] = {
                    "fonte": fonte, "df": df_bruto,
                    "col_nota": col_nota, "col_data": col_data,
                    "col_valor": col_valor, "col_evento": col_evento,
                    "metadados": metadados,
                    "empresa_detectada": detectar_empresa(df_bruto),
                    "competencia_detectada": detectar_competencia(df_bruto, col_data),
                }

st.write("---")

# --- Empresa e Competência: auto-preenchidas a partir das planilhas (prioridade: Domínio), mas editáveis ---
empresa_sugerida = obter_empresa_sugerida(dados_fontes)
competencia_sugerida = obter_competencia_sugerida(dados_fontes)

if empresa_sugerida and not st.session_state.get("empresa_input"):
    st.session_state["empresa_input"] = empresa_sugerida
if competencia_sugerida and not st.session_state.get("competencia_input"):
    st.session_state["competencia_input"] = competencia_sugerida

col_emp, col_comp = st.columns(2)
with col_emp:
    empresa = st.text_input("🏢 Empresa", key="empresa_input")
    if empresa_sugerida:
        st.caption("🔎 Detectado automaticamente a partir da planilha — pode editar.")
with col_comp:
    competencia = st.text_input("🗓️ Competência (ex: 08/2026)", key="competencia_input")
    if competencia_sugerida:
        st.caption("🔎 Detectado automaticamente a partir da planilha — pode editar.")

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
            duplicadas_por_fonte = {}
            for codigo in fontes_prontas:
                cfg = dados_fontes[codigo]
                agrupado, notas_dup = processar_fonte(
                    cfg["df"], cfg["col_nota"], cfg["col_data"], cfg["col_valor"], cfg["col_evento"], usar_data,
                    desconsiderar_zero=(codigo == "Dominio")
                )
                duplicadas_por_fonte[codigo] = notas_dup
                rename_map = {"valor": f"valor_{codigo}", "evento": f"evento_{codigo}"}
                if not usar_data:
                    rename_map["data"] = f"data_{codigo}"
                agrupado = agrupado.rename(columns=rename_map)
                processados[codigo] = agrupado

            chave = ['nota', 'data'] if usar_data else ['nota']
            m = reduce(lambda l, r: pd.merge(l, r, on=chave, how='outer'), processados.values())

            # --- Totais consolidados (antes de filtrar por divergência) ---
            st.subheader("📊 Totais Consolidados")
            if empresa or competencia:
                st.caption(f"{empresa or '—'} · Competência: {competencia or '—'} · {tipo_doc}")

            totais = {codigo: m[f"valor_{codigo}"].fillna(0).sum() for codigo in fontes_prontas}

            # Cartões de totais — e, com só 2 fontes, a diferença já entra na MESMA linha
            if len(fontes_prontas) == 2:
                codigo_a, codigo_b = fontes_prontas
                diferenca_total = totais[codigo_a] - totais[codigo_b]
                metricas = st.columns(3)
                with metricas[0]:
                    st.metric(f"Soma {dados_fontes[codigo_a]['fonte']}", formatar_moeda_br(totais[codigo_a]))
                with metricas[1]:
                    st.metric(f"Soma {dados_fontes[codigo_b]['fonte']}", formatar_moeda_br(totais[codigo_b]))
                with metricas[2]:
                    st.metric(
                        "Diferença Total",
                        formatar_moeda_br(diferenca_total),
                        delta=f"{diferenca_total:,.2f} R$" if abs(diferenca_total) > 0.01 else None,
                        delta_color="inverse" if abs(diferenca_total) > 0.01 else "normal"
                    )
            else:
                metricas = st.columns(len(fontes_prontas))
                for i, codigo in enumerate(fontes_prontas):
                    with metricas[i]:
                        st.metric(f"Soma {dados_fontes[codigo]['fonte']}", formatar_moeda_br(totais[codigo]))

            # Resumo por relatório — só com 3 fontes (com 2, o cartão acima já fecha a informação).
            # Comparações em pares: primeiro cada fonte vs a referência (Domínio, se houver),
            # depois o par restante entre as duas não-referência — fechando a triangulação
            # (ex: Fonte Oficial vs SIEG), em vez de repetir Domínio x Domínio.
            resumo_df = None
            if len(fontes_prontas) == 3:
                referencia = "Dominio" if "Dominio" in fontes_prontas else fontes_prontas[0]
                nome_referencia = dados_fontes[referencia]["fonte"]
                outras = [c for c in fontes_prontas if c != referencia]
                pares = [(referencia, c) for c in outras] + [(outras[0], outras[1])]

                resumo_linhas = []
                for a, b in pares:
                    nome_a, nome_b = dados_fontes[a]["fonte"], dados_fontes[b]["fonte"]
                    resumo_linhas.append({
                        "Comparação": f"{nome_a} vs {nome_b}",
                        "Diferença": totais[a] - totais[b]
                    })
                resumo_df = pd.DataFrame(resumo_linhas)

                st.markdown(f"**Resumo por relatório** (referência: {nome_referencia})")
                resumo_exib = resumo_df.copy()
                resumo_exib["Diferença"] = resumo_exib["Diferença"].apply(formatar_moeda_br)
                st.dataframe(resumo_exib, use_container_width=True, hide_index=True)

            # --- Análise linha a linha: ausências (só disparam a linha, sem texto), ---
            # --- divergência de valor, duplicidade e status de cancelamento/denegação ---
            def analisar_linha(row):
                ausentes = [c for c in fontes_prontas if pd.isna(row[f"valor_{c}"])]
                presentes = {c: row[f"valor_{c}"] for c in fontes_prontas if c not in ausentes}

                # "Ausente" já fica visualmente claro pela célula em branco — não vira texto no Motivo.
                gatilhos = bool(ausentes)

                motivo_partes = []
                if len(presentes) >= 2:
                    vals = list(presentes.values())
                    if max(vals) - min(vals) > 0.01:
                        motivo_partes.append("Valor divergente")
                        gatilhos = True

                for c in fontes_prontas:
                    if row["nota"] in duplicadas_por_fonte.get(c, set()):
                        motivo_partes.append(f"Nota duplicada no {dados_fontes[c]['fonte']}")
                        gatilhos = True

                # Status de cancelamento/denegação (SIEG: "Tipo de Evento", ou qualquer coluna de
                # status com Cancelamento/Desc. de Operação) — quando detectado, entra no motivo
                # e por si só já justifica mostrar a nota (vale a pena revisar mesmo se os valores baterem).
                for c in fontes_prontas:
                    ev = row.get(f"evento_{c}", "")
                    if isinstance(ev, str) and ev.strip():
                        motivo_partes.append(f"{dados_fontes[c]['fonte']}: {ev.strip()}")
                        gatilhos = True

                return pd.Series({
                    "Motivo da Inconsistência": " | ".join(motivo_partes),
                    "_divergente": gatilhos
                })

            extras = m.apply(analisar_linha, axis=1)
            m = pd.concat([m, extras], axis=1)
            divergencias = m[m["_divergente"]].copy()

            # --- Coluna Data: exibida sempre, exceto para NFe Entrada ---
            # Se a data faz parte da chave, já está em 'data'. Senão, usa a primeira
            # data disponível entre as fontes como referência.
            if mostrar_data:
                if usar_data:
                    divergencias["Data"] = divergencias["data"]
                else:
                    data_ref_cols = [f"data_{c}" for c in fontes_prontas if f"data_{c}" in divergencias.columns]
                    if data_ref_cols:
                        divergencias["Data"] = divergencias[data_ref_cols].bfill(axis=1).iloc[:, 0]
                    else:
                        divergencias["Data"] = None

            divergencias = divergencias.sort_values(by=chave)

            st.write("---")
            st.subheader("🔍 Divergências Encontradas")

            if divergencias.empty:
                st.success("🎉 Excelente! As fontes bateram — nenhuma divergência encontrada.")
            else:
                st.warning(f"Foram identificadas {len(divergencias)} notas com inconsistências.")

                # A coluna Motivo só aparece se houver alguma informação real nela
                # (status de cancelamento/denegação, valor divergente ou duplicidade).
                tem_motivo = divergencias["Motivo da Inconsistência"].str.strip().ne("").any()

                # Montagem da tabela de exibição (moeda formatada) — Empresa/Competência ficam
                # só no título/cabeçalho, não repetidas linha a linha.
                exib = pd.DataFrame()
                exib["Número da Nota"] = divergencias["nota"]
                if mostrar_data:
                    exib["Data"] = pd.to_datetime(divergencias["Data"], errors='coerce').dt.strftime('%d/%m/%Y')

                for c in fontes_prontas:
                    exib[NOME_COLUNA_VALOR[c]] = divergencias[f"valor_{c}"].apply(
                        lambda v: "" if pd.isna(v) else formatar_moeda_br(v)
                    )

                if tem_motivo:
                    exib["Motivo da Inconsistência"] = divergencias["Motivo da Inconsistência"]
                exib = exib.reset_index(drop=True)

                st.dataframe(exib, use_container_width=True)

                # Exportação em Excel (valores numéricos, não formatados, para o usuário trabalhar em cima)
                export_df = pd.DataFrame()
                export_df["Número da Nota"] = divergencias["nota"].values
                if mostrar_data:
                    export_df["Data"] = pd.to_datetime(divergencias["Data"], errors='coerce').dt.strftime('%d/%m/%Y').values
                for c in fontes_prontas:
                    export_df[NOME_COLUNA_VALOR[c]] = divergencias[f"valor_{c}"].values
                if tem_motivo:
                    export_df["Motivo da Inconsistência"] = divergencias["Motivo da Inconsistência"].values

                output = io.BytesIO()
                linha_inicio = 3  # espaço para o cabeçalho (Empresa/Competência/Tipo)
                totais_df = pd.DataFrame([
                    {"Relatório": dados_fontes[c]["fonte"], "Valor Total": totais[c]} for c in fontes_prontas
                ])

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Divergências', startrow=linha_inicio)
                    ws = writer.sheets['Divergências']
                    ws.write(0, 0, f"Empresa: {empresa or '-'}")
                    ws.write(1, 0, f"Competência: {competencia or '-'}")
                    ws.write(2, 0, f"Tipo de Documento: {tipo_doc}")

                    # Aba Resumo: valor total de todas as fontes sempre presente,
                    # + comparação em pares quando há 3 fontes.
                    totais_df.to_excel(writer, index=False, sheet_name='Resumo', startrow=linha_inicio)
                    ws2 = writer.sheets['Resumo']
                    ws2.write(0, 0, f"Empresa: {empresa or '-'}")
                    ws2.write(1, 0, f"Competência: {competencia or '-'}")
                    ws2.write(2, 0, f"Tipo de Documento: {tipo_doc}")

                    if resumo_df is not None and not resumo_df.empty:
                        inicio_comparacao = linha_inicio + len(totais_df) + 3
                        ws2.write(inicio_comparacao - 1, 0, "Comparação entre fontes")
                        resumo_df.to_excel(writer, index=False, sheet_name='Resumo', startrow=inicio_comparacao)

                partes_nome = [p for p in [empresa, competencia, tipo_doc.replace(' ', '_')] if p]
                nome_arquivo = "divergencias_" + "_".join(partes_nome).replace('/', '-').replace(' ', '_') + ".xlsx"
                st.download_button(
                    label="📥 Baixar Planilha de Divergências",
                    data=output.getvalue(),
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro processual: {e}")
