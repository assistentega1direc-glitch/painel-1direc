"""
Painel de Acompanhamento Acadêmico — RAPP 2026 · 1ª DIREC (Natal/RN)

Arquitetura (resumo):
- A "Base componentes" (1 linha por aluno x componente x plataforma) é a FONTE ÚNICA da verdade.
  A situação de cada aluno (por plataforma e geral) é sempre DERIVADA dela, então todas as abas
  (Visão Geral, Prioritárias, Todas as Escolas, Caderneta, Detalhamento, Relatório) ficam coerentes.
- Importações do Plurall (Redash) e do Avalia atualizam a base de componentes, registram quem mudou
  de situação e gravam um "marco" (snapshot) datado para o gráfico de evolução.
- Visitas, frequência, base atualizada, marcos e log de mudanças são persistidos no Google Sheets
  (via conta de serviço). Sem credenciais, o app funciona em "modo local" (CSV na pasta dados_app/).
"""

import os
import re
import unicodedata
import uuid
from datetime import datetime, date
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURAÇÃO DA PÁGINA (precisa ser o primeiro comando st.*)
# ============================================================
st.set_page_config(
    page_title="Painel de Acompanhamento Acadêmico - 1ª DIREC",
    page_icon="🎓",
    layout="wide",
)

st.markdown("""
    <style>
        .main { background-color: #f8fafc; }
        .stMetric { background-color: #ffffff; padding: 12px; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
        @media (max-width: 768px) {
            .main .block-container { padding-left: 0.5rem !important; padding-right: 0.5rem !important; padding-top: 1rem !important; }
            .stMetric { padding: 10px !important; margin-bottom: 8px !important; }
            div[data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; min-width: 100% !important; }
        }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CONSTANTES
# ============================================================
EXCEL_DEFAULT = "RAPP_1DIREC_Situacao_Alunos_Tratada.xlsx"
DATA_BASE_INICIAL = date(2026, 9, 28)   # data do relatório que originou a base tratada
SENHA_ADMIN = "direc2026"               # altere quando desejar (ou defina admin_password em secrets)
TZ = ZoneInfo("America/Sao_Paulo")
PASTA_LOCAL = "dados_app"

AP, NAP, NAV = "Aprovado", "Não aprovado", "Não avaliado"
ORDEM_SIT = [AP, NAP, NAV]
CORES_SIT = {AP: "#10b981", NAP: "#f43f5e", NAV: "#f59e0b"}

COMP_COLS = ["Professor tutor", "Escola", "Matrícula", "Nome", "Série", "Componente",
             "Rendimento", "Situação", "Status plataforma", "Plataforma", "Prova", "Atualizado em"]

# Abas (worksheets) usadas na planilha do Google Drive
WS_VISITAS = "Visitas"
WS_FREQ = "Frequencia"
WS_HIST = "Historico_Situacao"
WS_LOG = "Log_Mudancas"
WS_IMPORT = "Importacoes"
WS_BASE = "Base_Atual"

COLS_VISITAS = ["ID", "Data_Visita", "Data_Registro", "Escola", "Turno", "Articulador",
                "Registrado_por", "Atendidos", "Aprovados_Liberados", "Presentes", "Observacoes"]
COLS_FREQ = ["ID_Visita", "Data_Visita", "Escola", "Matrícula", "Nome", "Plataforma",
             "Situação", "Frequencia", "Data_Registro"]
COLS_HIST = ["Import_ID", "Data_Referencia", "Data_Registro", "Fonte", "Escola", "Professor_tutor",
             "Assessor", "Total", "Aprovados", "Nao_Aprovados", "Nao_Avaliados"]
COLS_LOG = ["Import_ID", "Data_Referencia", "Matrícula", "Nome", "Escola", "Plataforma",
            "Situacao_anterior", "Situacao_nova"]
COLS_IMPORT = ["Import_ID", "Data_Importacao", "Data_Referencia", "Fonte", "Arquivos",
               "Linhas_lidas", "Componentes_atualizados", "Componentes_novos",
               "Alunos_mudaram", "Novos_aprovados", "Fora_da_base"]

# Variações de nome de escola entre plataformas / abas (chave já normalizada -> chave canônica)
ALIASES_ESCOLA = {
    "INST PE MIGUELINHO": "INSTITUTO PADRE MIGUELINHO",
    "E E GOV WALFREDO GURGEL":
        "E E GOV WALFREDO GURGEL ENSINO MEDIO EM TEMPO INTEGRAL E EDUC PROFISSIONAL TECNICA DE NIVEL MEDIO",
    "INSTITUTO ESTADUAL DE EDUCACAO PROFISSIONAL TECNOLOGIA E INOVACAO DO RN NATAL":
        "INSTITUTO ESTADUAL DE EDUCACAO PROFISSIONAL TECNOLOGIA E INOVACAO DO RIO GRANDE DO NORTE NATAL",
}
ALIASES_COMPONENTE = {"LINGUA INGLES": "LINGUA INGLESA"}


def agora():
    return datetime.now(TZ)


# ============================================================
# NORMALIZAÇÕES
# ============================================================
def norm_txt(v):
    """Sem acento, maiúsculo, sem pontuação, espaços simples."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", " ", s).upper()
    return " ".join(s.split())


def key_escola(v):
    k = norm_txt(v)
    if k.startswith("IERN "):
        k = k[5:]
    return ALIASES_ESCOLA.get(k, k)


def key_comp(v):
    k = norm_txt(v)
    return ALIASES_COMPONENTE.get(k, k)


def key_mat(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s).lstrip("0")
    return s


def norm_situacao(v):
    """Converte qualquer status de plataforma nas 3 situações oficiais."""
    k = norm_txt(v)
    if k == "APROVADO":
        return AP
    if k in ("REPROVADO", "NAO APROVADO"):
        return NAP
    # Não Avaliado, AGUARDANDO_INICIO, AGENDADA, NÃO AGENDADA, CANCELADA, Não obrigado...
    return NAV


def situacao_agregada(sits):
    s = set(sits)
    if NAP in s:
        return NAP
    if NAV in s:
        return NAV
    return AP


def fmt_data_br(iso):
    try:
        return pd.to_datetime(iso).strftime("%d/%m/%Y")
    except Exception:
        return str(iso)


# ============================================================
# CAMADA DE ARMAZENAMENTO (Google Sheets com fallback local)
# ============================================================
def gsheets_configurado():
    try:
        return "gcp_service_account" in st.secrets and "gsheets" in st.secrets
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _abrir_planilha():
    import gspread
    cliente = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    cfg = st.secrets["gsheets"]
    if "spreadsheet_id" in cfg:
        return cliente.open_by_key(cfg["spreadsheet_id"])
    return cliente.open_by_url(cfg["spreadsheet_url"])


def _nome_ws(nome):
    """Permite renomear abas via secrets: [gsheets] aba_visitas = 'Página1'."""
    try:
        mapa = {WS_VISITAS: "aba_visitas", WS_FREQ: "aba_frequencia", WS_HIST: "aba_historico",
                WS_LOG: "aba_log", WS_IMPORT: "aba_importacoes", WS_BASE: "aba_base"}
        return st.secrets["gsheets"].get(mapa.get(nome, ""), nome)
    except Exception:
        return nome


def _ws(nome, colunas):
    import gspread
    sh = _abrir_planilha()
    titulo = _nome_ws(nome)
    try:
        ws = sh.worksheet(titulo)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=titulo, rows=1000, cols=max(len(colunas), 12))
        ws.update(values=[colunas], range_name="A1")
        return ws, list(colunas)
    cab = ws.row_values(1)
    if not cab:
        ws.update(values=[colunas], range_name="A1")
        return ws, list(colunas)
    faltando = [c for c in colunas if c not in cab]
    if faltando:  # aba antiga com cabeçalho diferente: acrescenta as colunas que faltam
        cab = cab + faltando
        if ws.col_count < len(cab):
            ws.add_cols(len(cab) - ws.col_count)
        ws.update(values=[cab], range_name="A1")
    return ws, cab


def _csv_local(nome):
    os.makedirs(PASTA_LOCAL, exist_ok=True)
    return os.path.join(PASTA_LOCAL, f"{nome}.csv")


def _ler(nome, colunas):
    if gsheets_configurado():
        ws, _ = _ws(nome, colunas)
        valores = ws.get_all_values()
        if len(valores) <= 1:
            return pd.DataFrame(columns=colunas)
        df = pd.DataFrame(valores[1:], columns=valores[0])
    else:
        caminho = _csv_local(nome)
        if not os.path.exists(caminho):
            return pd.DataFrame(columns=colunas)
        df = pd.read_csv(caminho, dtype=str, keep_default_na=False)
    for c in colunas:
        if c not in df.columns:
            df[c] = ""
    df = df.replace("", pd.NA)
    return df.dropna(how="all")


@st.cache_data(ttl=60, show_spinner=False)
def ler_tabela(nome, colunas):
    return _ler(nome, list(colunas))


@st.cache_data(ttl=300, show_spinner=False)
def ler_base_persistida():
    return _ler(WS_BASE, COMP_COLS)


def acrescentar_linhas(nome, colunas, linhas):
    """linhas: lista de dicts."""
    if not linhas:
        return
    if gsheets_configurado():
        ws, cab = _ws(nome, colunas)
        valores = [["" if pd.isna(l.get(c, "")) else str(l.get(c, "")) for c in cab] for l in linhas]
        ws.append_rows(valores, value_input_option="RAW")
    else:
        caminho = _csv_local(nome)
        novo = pd.DataFrame(linhas, columns=colunas)
        existe = os.path.exists(caminho)
        novo.to_csv(caminho, mode="a", header=not existe, index=False)
    ler_tabela.clear()


def sobrescrever_base(df):
    df = df[COMP_COLS].fillna("").astype(str)
    if gsheets_configurado():
        ws, _ = _ws(WS_BASE, COMP_COLS)
        ws.clear()
        n_lin, n_col = len(df) + 1, len(COMP_COLS)
        if ws.row_count < n_lin or ws.col_count < n_col:
            ws.resize(rows=max(n_lin, ws.row_count), cols=max(n_col, ws.col_count))
        ws.update(values=[COMP_COLS] + df.values.tolist(), range_name="A1", value_input_option="RAW")
    else:
        df.to_csv(_csv_local(WS_BASE), index=False)
    ler_base_persistida.clear()


def apagar_base_persistida():
    if gsheets_configurado():
        ws, _ = _ws(WS_BASE, COMP_COLS)
        ws.clear()
        ws.update(values=[COMP_COLS], range_name="A1")
    else:
        caminho = _csv_local(WS_BASE)
        if os.path.exists(caminho):
            os.remove(caminho)
    ler_base_persistida.clear()


# ============================================================
# CARGA E PADRONIZAÇÃO DA BASE
# ============================================================
def mapa_escolas_canonicas(nomes):
    """chave normalizada -> nome de exibição (prefere a grafia em MAIÚSCULAS, a mais frequente)."""
    s = pd.Series(list(nomes)).dropna().astype(str).str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    df = pd.DataFrame({"nome": s, "k": s.map(key_escola)})
    df["maiusc"] = df["nome"].str.upper() == df["nome"]
    df = (df.groupby(["k", "nome", "maiusc"]).size().reset_index(name="n")
            .sort_values(["k", "maiusc", "n"], ascending=[True, False, False]))
    return df.drop_duplicates("k").set_index("k")["nome"].to_dict()


def padronizar_comp(comp):
    comp = comp.copy()
    comp.columns = [str(c).strip() for c in comp.columns]
    if "Rendimento (%)" in comp.columns and "Rendimento" not in comp.columns:
        comp = comp.rename(columns={"Rendimento (%)": "Rendimento"})
    if "Status plataforma" not in comp.columns or comp["Status plataforma"].isna().all():
        comp["Status plataforma"] = comp["Situação"]
    for c in COMP_COLS:
        if c not in comp.columns:
            comp[c] = ""
    comp = comp[COMP_COLS]
    comp["Matrícula"] = comp["Matrícula"].map(lambda v: key_mat(v) or str(v))
    for c in ["Professor tutor", "Escola", "Nome", "Série", "Componente", "Plataforma",
              "Status plataforma", "Prova", "Atualizado em", "Rendimento"]:
        comp[c] = comp[c].fillna("").astype(str).str.strip()
    comp["Componente"] = comp["Componente"].str.replace(r"\s+", " ", regex=True)
    comp["Situação"] = comp["Situação"].map(norm_situacao)
    comp["Plataforma"] = comp["Plataforma"].map(lambda p: "Avalia" if "AVALIA" in norm_txt(p) else "Plurall")
    # Escola com grafia única (corrige escolas que apareciam duplicadas: MAIÚSCULA x Título, siglas etc.)
    canon = mapa_escolas_canonicas(comp["Escola"])
    comp["Escola"] = comp["Escola"].map(lambda e: canon.get(key_escola(e), e))
    # Remove linhas 100% duplicadas (mesmo aluno, componente, prova, nota e situação)
    comp = comp.drop_duplicates(subset=["Matrícula", "Plataforma", "Componente", "Prova",
                                        "Rendimento", "Status plataforma"])
    return comp.reset_index(drop=True)


@st.cache_data(show_spinner="Carregando base oficial...")
def carregar_base_inicial(caminho, _mtime):
    xls = pd.ExcelFile(caminho)
    comp = pd.read_excel(xls, sheet_name="Base componentes", dtype={"Matrícula": str})
    aba_art = next((s for s in xls.sheet_names if "articulador" in s.lower()), None)
    art = pd.read_excel(xls, sheet_name=aba_art) if aba_art else pd.DataFrame(columns=["Escola", "Assessor(a)"])
    return padronizar_comp(comp), art


def mapear_articuladores(escolas, art):
    """Liga cada escola da base à linha da aba Articuladores (tolera acentos, caixa e nomes truncados)."""
    art = art.copy()
    art["_k"] = art["Escola"].map(key_escola)
    art = art[art["_k"] != ""]
    resultado = {}
    for esc in escolas:
        k = key_escola(esc)
        linha = art[art["_k"] == k]
        if linha.empty:  # nomes cortados em ~80 caracteres na aba Articuladores
            linha = art[art["_k"].map(lambda a: len(a) >= 25 and (k.startswith(a) or a.startswith(k)))]
        if not linha.empty:
            r = linha.iloc[0]
            resultado[esc] = {
                "Assessor(a)": str(r.get("Assessor(a)", "") or "Não Informado").strip(),
                "Articulador(a)": str(r.get("Articulador(a)", "") or "").strip(),
            }
    return resultado


def consolidar_plataforma(comp):
    """1 linha por aluno x plataforma (mesma regra da planilha tratada)."""
    c = comp.copy()
    c["_nap"] = c["Situação"] == NAP
    c["_nav"] = c["Situação"] == NAV
    c["_pend"] = c["Componente"].where(c["Situação"] != AP)
    g = c.groupby(["Matrícula", "Plataforma"], sort=False)
    df = g.agg(**{
        "Professor tutor": ("Professor tutor", "first"),
        "Escola": ("Escola", "first"),
        "Nome": ("Nome", "first"),
        "Série": ("Série", lambda x: " / ".join(sorted(set(v for v in x if v)))),
        "Componentes": ("Componente", "size"),
        "_nap": ("_nap", "any"),
        "_nav": ("_nav", "any"),
        "Pendências": ("_pend", lambda x: ", ".join(sorted(set(v for v in x.dropna())))),
    }).reset_index()
    df["Situação do aluno"] = AP
    df.loc[df["_nav"], "Situação do aluno"] = NAV
    df.loc[df["_nap"], "Situação do aluno"] = NAP
    return df.drop(columns=["_nap", "_nav"])


def consolidar_alunos(cons):
    """1 linha por aluno (une Plurall + Avalia). Liberado = sem nenhuma pendência em nenhuma plataforma."""
    if cons.empty:
        return pd.DataFrame(columns=["Matrícula", "Nome", "Escola", "Professor tutor", "Assessor(a)",
                                     "Plataforma", "Série", "Situação do aluno", "Pendências"])
    g = cons.groupby("Matrícula", sort=False)
    df = g.agg(**{
        "Nome": ("Nome", "first"),
        "Escola": ("Escola", "first"),
        "Professor tutor": ("Professor tutor", "first"),
        "Assessor(a)": ("Assessor(a)", "first"),
        "Plataforma": ("Plataforma", lambda x: " + ".join(sorted(set(x)))),
        "Série": ("Série", "first"),
        "Situação do aluno": ("Situação do aluno", situacao_agregada),
        "Pendências": ("Pendências", lambda x: "; ".join(v for v in x if v)),
    }).reset_index()
    return df


@st.cache_data(ttl=300, show_spinner="Atualizando painel...")
def obter_dados():
    caminho = EXCEL_DEFAULT
    comp_ini, art = carregar_base_inicial(caminho, os.path.getmtime(caminho))
    origem = "Base inicial (planilha tratada)"
    try:
        persistida = ler_base_persistida()
        if persistida is not None and len(persistida) > 0:
            comp = padronizar_comp(persistida)
            origem = "Base atualizada (salva no armazenamento)"
        else:
            comp = comp_ini
    except Exception as e:
        comp = comp_ini
        origem = f"Base inicial (não foi possível ler a base salva: {e})"
    mapa = mapear_articuladores(comp["Escola"].unique(), art)
    comp["Assessor(a)"] = comp["Escola"].map(lambda e: mapa.get(e, {}).get("Assessor(a)", "Não Informado"))
    cons = consolidar_plataforma(comp)
    cons["Assessor(a)"] = cons["Escola"].map(lambda e: mapa.get(e, {}).get("Assessor(a)", "Não Informado"))
    cons["Articulador(a)"] = cons["Escola"].map(lambda e: mapa.get(e, {}).get("Articulador(a)", ""))
    return comp, cons, mapa, origem


def recarregar_tudo():
    obter_dados.clear()
    ler_base_persistida.clear()
    ler_tabela.clear()


# ============================================================
# IMPORTAÇÃO: PLURALL (Redash) e AVALIA
# ============================================================
def identificar_planilha(df, nomes_abas=None):
    cols = {norm_txt(c) for c in df.columns}
    if "EMAIL ALUNO" in cols and "SITUACAO" in cols:
        return "Plurall"
    if "MATRICULA SISTEMA" in cols and "SITUACAO DA AVALIACAO" in cols:
        return "Avalia"
    if nomes_abas and "Base componentes" in nomes_abas:
        return "Base consolidada"
    return None


def _col(df, *alvos):
    mapa = {norm_txt(c): c for c in df.columns}
    for a in alvos:
        if norm_txt(a) in mapa:
            return mapa[norm_txt(a)]
    return None


def preparar_plurall(raw, comp):
    df = raw.copy()
    c_grupo = _col(df, "Grupo")
    if c_grupo:
        df = df[df[c_grupo].astype(str).str.contains(r"01ª DIREC|01 DIREC|1ª DIREC|NATAL", case=False, na=False)]
    c_email, c_aluno, c_escola = _col(df, "email_aluno"), _col(df, "aluno"), _col(df, "escola")
    c_prova, c_nome_prova = _col(df, "prova"), _col(df, "name")
    c_rend, c_sit = _col(df, "rendimento (%)", "rendimento"), _col(df, "Situação")

    df["_mk"] = df[c_email].astype(str).str.extract(r"^\s*(\d+)")[0].map(key_mat)
    # Linhas sem e-mail: recupera a matrícula pelo nome (+ escola) do aluno na base
    sem = df["_mk"] == ""
    if sem.any() and c_aluno:
        b = comp.drop_duplicates(["Matrícula"]).copy()
        b["_kn"] = b["Nome"].map(norm_txt)
        b["_ke"] = b["Escola"].map(key_escola)
        por_nome_escola = b.drop_duplicates(["_kn", "_ke"], keep=False).set_index(["_kn", "_ke"])["Matrícula"]
        por_nome = b.drop_duplicates("_kn", keep=False).set_index("_kn")["Matrícula"]
        for i in df.index[sem]:
            kn = norm_txt(df.at[i, c_aluno])
            ke = key_escola(df.at[i, c_escola]) if c_escola else ""
            m = por_nome_escola.get((kn, ke)) or por_nome.get(kn)
            if m:
                df.at[i, "_mk"] = key_mat(m)

    out = pd.DataFrame({
        "_mk": df["_mk"],
        "_ck": df[c_prova].map(key_comp),
        "Componente_raw": df[c_prova].astype(str).str.strip(),
        "Nome_raw": df[c_aluno].astype(str) if c_aluno else "",
        "Escola_raw": df[c_escola].astype(str) if c_escola else "",
        "Prova": df[c_nome_prova].astype(str).str.strip() if c_nome_prova else df[c_prova].astype(str),
        "Status plataforma": df[c_sit].astype(str).str.strip(),
        "_rend": pd.to_numeric(df[c_rend], errors="coerce") if c_rend else pd.NA,
    })
    out["Situação"] = out["Status plataforma"].map(norm_situacao)
    out["Rendimento"] = out["_rend"].map(lambda v: "—" if pd.isna(v) else f"{v:g}")
    out["_rank"] = out["Situação"].map({AP: 2, NAP: 1, NAV: 0})
    # Mesma prova repetida (refação): fica o melhor resultado
    out = (out.sort_values(["_rank", "_rend"], ascending=False)
              .drop_duplicates(["_mk", "_ck", "Prova"]))
    return out, len(df)


def preparar_avalia(raw):
    df = raw.copy()
    c_direc = _col(df, "DIREC")
    if c_direc:
        df = df[df[c_direc].astype(str).str.contains(r"01ª DIREC|01 DIREC|1ª DIREC|NATAL", case=False, na=False)]
    c_mat, c_comp = _col(df, "MATRÍCULA SISTEMA"), _col(df, "COMPONENTE CURRICULAR")
    c_sit, c_nota = _col(df, "SITUAÇÃO DA AVALIAÇÃO"), _col(df, "NOTA FINAL")
    c_nome, c_esc = _col(df, "NOME DO ESTUDANTE"), _col(df, "ESCOLA")
    out = pd.DataFrame({
        "_mk": df[c_mat].map(key_mat),
        "_ck": df[c_comp].map(key_comp),
        "Componente_raw": df[c_comp].astype(str).str.strip(),
        "Nome_raw": df[c_nome].astype(str) if c_nome else "",
        "Escola_raw": df[c_esc].astype(str) if c_esc else "",
        "Prova": "",
        "Status plataforma": df[c_sit].astype(str).str.strip(),
        "Rendimento": df[c_nota].astype(str).str.strip().replace({"-": "—", "nan": "—"}) if c_nota else "—",
    })
    out["Situação"] = out["Status plataforma"].map(norm_situacao)
    prioridade = {"APROVADO": 5, "REPROVADO": 4, "AGENDADA": 3, "AGUARDANDO INICIO": 2, "NAO AGENDADA": 1}
    out["_rank"] = out["Status plataforma"].map(lambda s: prioridade.get(norm_txt(s), 0))
    out = out.sort_values("_rank", ascending=False).drop_duplicates(["_mk", "_ck"])
    return out, len(df)


def aplicar_atualizacao(comp, novos, plataforma, data_ref_iso):
    """Substitui, para cada (aluno, componente) presente no arquivo novo, as linhas da plataforma.
    O que não aparece no arquivo novo é mantido como estava (nada é apagado por ausência)."""
    comp = comp.copy()
    comp["_mk"] = comp["Matrícula"].map(key_mat)
    comp["_ck"] = comp["Componente"].map(key_comp)

    novos = novos[novos["_mk"] != ""]
    alunos_base = set(comp["_mk"])
    fora = novos[~novos["_mk"].isin(alunos_base)]
    novos = novos[novos["_mk"].isin(alunos_base)]

    eh_plat = comp["Plataforma"] == plataforma
    chaves_novas = pd.MultiIndex.from_arrays([novos["_mk"], novos["_ck"]]).unique()
    idx_comp = pd.MultiIndex.from_arrays([comp["_mk"], comp["_ck"]])
    substituir = eh_plat & idx_comp.isin(chaves_novas)

    chaves_base_plat = set(zip(comp.loc[eh_plat, "_mk"], comp.loc[eh_plat, "_ck"]))
    n_novos_comp = sum(1 for k in zip(novos["_mk"], novos["_ck"]) if k not in chaves_base_plat)

    # metadados do aluno: prioriza a linha da mesma plataforma
    meta = pd.concat([comp[eh_plat], comp[~eh_plat]]).drop_duplicates("_mk").set_index("_mk")
    nome_comp = (comp.drop_duplicates(["_mk", "_ck"]).set_index(["_mk", "_ck"])["Componente"])
    nome_comp_geral = comp.drop_duplicates("_ck").set_index("_ck")["Componente"]

    linhas = []
    for r in novos.to_dict("records"):
        m = meta.loc[r["_mk"]]
        comp_nome = nome_comp.get((r["_mk"], r["_ck"])) or nome_comp_geral.get(r["_ck"]) or r["Componente_raw"]
        linhas.append({
            "Professor tutor": m["Professor tutor"], "Escola": m["Escola"], "Matrícula": m["Matrícula"],
            "Nome": m["Nome"], "Série": m["Série"], "Componente": comp_nome,
            "Rendimento": r["Rendimento"], "Situação": r["Situação"], "Status plataforma": r["Status plataforma"],
            "Plataforma": plataforma, "Prova": r["Prova"], "Atualizado em": data_ref_iso,
        })
    novo_comp = pd.concat([comp.loc[~substituir, COMP_COLS], pd.DataFrame(linhas, columns=COMP_COLS)],
                          ignore_index=True)
    stats = {
        "linhas_validas": len(novos),
        "componentes_atualizados": int(substituir.sum()),
        "componentes_novos": n_novos_comp,
        "fora_da_base": fora,
    }
    return novo_comp, stats


def comparar_situacoes(cons_antes, cons_depois):
    a = cons_antes[["Matrícula", "Plataforma", "Situação do aluno"]].rename(columns={"Situação do aluno": "Antes"})
    d = cons_depois[["Matrícula", "Plataforma", "Nome", "Escola", "Situação do aluno"]].rename(
        columns={"Situação do aluno": "Depois"})
    m = d.merge(a, on=["Matrícula", "Plataforma"], how="left")
    m["Antes"] = m["Antes"].fillna("(sem registro)")
    return m[m["Antes"] != m["Depois"]].reset_index(drop=True)


def calcular_snapshot(cons, import_id, data_ref_iso, fonte):
    uni = consolidar_alunos(cons)
    registro = agora().strftime("%Y-%m-%d %H:%M:%S")
    linhas = []
    for esc, g in uni.groupby("Escola"):
        vc = g["Situação do aluno"].value_counts()
        linhas.append({
            "Import_ID": import_id, "Data_Referencia": data_ref_iso, "Data_Registro": registro,
            "Fonte": fonte, "Escola": esc,
            "Professor_tutor": g["Professor tutor"].mode().iat[0] if len(g) else "",
            "Assessor": g["Assessor(a)"].mode().iat[0] if len(g) else "",
            "Total": len(g), "Aprovados": int(vc.get(AP, 0)),
            "Nao_Aprovados": int(vc.get(NAP, 0)), "Nao_Avaliados": int(vc.get(NAV, 0)),
        })
    return linhas


def registrar_marco(cons, data_ref_iso, fonte):
    import_id = agora().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:4]
    acrescentar_linhas(WS_HIST, COLS_HIST, calcular_snapshot(cons, import_id, data_ref_iso, fonte))
    return import_id


def processar_importacao(arquivos_lidos, data_ref, comp_atual, cons_atual, mapa_art):
    """arquivos_lidos: lista de (nome_arquivo, tipo, DataFrame)."""
    data_ref_iso = data_ref.isoformat()
    import_id = agora().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:4]
    comp = comp_atual[COMP_COLS].copy()
    total_lidas, atualizados, novos_comp = 0, 0, 0
    fora_lista, fontes = [], []
    for nome_arq, tipo, df in arquivos_lidos:
        if tipo == "Plurall":
            novos, lidas = preparar_plurall(df, comp)
        else:
            novos, lidas = preparar_avalia(df)
        comp, stt = aplicar_atualizacao(comp, novos, tipo, data_ref_iso)
        total_lidas += lidas
        atualizados += stt["componentes_atualizados"]
        novos_comp += stt["componentes_novos"]
        fontes.append(tipo)
        if len(stt["fora_da_base"]):
            f = stt["fora_da_base"][["_mk", "Nome_raw", "Escola_raw", "Componente_raw", "Situação"]].copy()
            f.columns = ["Matrícula", "Nome", "Escola", "Componente", "Situação"]
            f["Plataforma"] = tipo
            fora_lista.append(f)

    comp = padronizar_comp(comp)
    cons_novo = consolidar_plataforma(comp)
    cons_novo["Assessor(a)"] = cons_novo["Escola"].map(lambda e: mapa_art.get(e, {}).get("Assessor(a)", "Não Informado"))
    mudancas = comparar_situacoes(cons_atual, cons_novo)

    # Persistência: base atual, log de mudanças, marco para o gráfico de evolução e registro da importação
    sobrescrever_base(comp)
    acrescentar_linhas(WS_LOG, COLS_LOG, [{
        "Import_ID": import_id, "Data_Referencia": data_ref_iso, "Matrícula": r["Matrícula"],
        "Nome": r["Nome"], "Escola": r["Escola"], "Plataforma": r["Plataforma"],
        "Situacao_anterior": r["Antes"], "Situacao_nova": r["Depois"],
    } for _, r in mudancas.iterrows()])
    fonte_txt = " + ".join(sorted(set(fontes)))
    acrescentar_linhas(WS_HIST, COLS_HIST, calcular_snapshot(cons_novo, import_id, data_ref_iso, fonte_txt))
    fora = pd.concat(fora_lista, ignore_index=True) if fora_lista else pd.DataFrame()
    acrescentar_linhas(WS_IMPORT, COLS_IMPORT, [{
        "Import_ID": import_id, "Data_Importacao": agora().strftime("%Y-%m-%d %H:%M:%S"),
        "Data_Referencia": data_ref_iso, "Fonte": fonte_txt,
        "Arquivos": " | ".join(a[0] for a in arquivos_lidos), "Linhas_lidas": total_lidas,
        "Componentes_atualizados": atualizados, "Componentes_novos": novos_comp,
        "Alunos_mudaram": len(mudancas), "Novos_aprovados": int((mudancas["Depois"] == AP).sum()),
        "Fora_da_base": len(fora),
    }])
    recarregar_tudo()
    return {
        "import_id": import_id, "data_ref": data_ref_iso, "fonte": fonte_txt,
        "linhas_lidas": total_lidas, "atualizados": atualizados, "novos_comp": novos_comp,
        "mudancas": mudancas, "fora": fora,
    }


# ============================================================
# CARREGAMENTO PRINCIPAL
# ============================================================
if not os.path.exists(EXCEL_DEFAULT):
    st.error(f"Arquivo {EXCEL_DEFAULT} não encontrado no repositório.")
    st.stop()

try:
    df_comp, df_cons, MAPA_ART, ORIGEM_BASE = obter_dados()
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()

ARMAZENAMENTO_OK = True
try:
    df_visitas = ler_tabela(WS_VISITAS, tuple(COLS_VISITAS)).copy()
    df_hist = ler_tabela(WS_HIST, tuple(COLS_HIST)).copy()
    df_log = ler_tabela(WS_LOG, tuple(COLS_LOG)).copy()
    df_imports = ler_tabela(WS_IMPORT, tuple(COLS_IMPORT)).copy()
except Exception as e:
    ARMAZENAMENTO_OK = False
    ERRO_ARMAZENAMENTO = str(e)
    df_visitas = pd.DataFrame(columns=COLS_VISITAS)
    df_hist = pd.DataFrame(columns=COLS_HIST)
    df_log = pd.DataFrame(columns=COLS_LOG)
    df_imports = pd.DataFrame(columns=COLS_IMPORT)

# Visitas: compatibilidade com registros antigos (só com Data_Registro)
for c in ["Atendidos", "Aprovados_Liberados", "Presentes"]:
    df_visitas[c] = pd.to_numeric(df_visitas[c], errors="coerce").fillna(0).astype(int)
df_visitas["Data_Visita"] = df_visitas["Data_Visita"].fillna(df_visitas["Data_Registro"])
df_visitas["_data"] = pd.to_datetime(df_visitas["Data_Visita"], errors="coerce", dayfirst=False)
if len(df_visitas):
    canon_v = {key_escola(e): e for e in df_cons["Escola"].unique()}
    df_visitas["Escola"] = df_visitas["Escola"].map(lambda e: canon_v.get(key_escola(e), e))

# Marco inicial automático (1x): situação da base tratada, para o gráfico ter ponto de partida
if ARMAZENAMENTO_OK and df_hist.empty and not st.session_state.get("_marco_inicial_ok"):
    st.session_state["_marco_inicial_ok"] = True
    try:
        registrar_marco(df_cons, DATA_BASE_INICIAL.isoformat(), "Marco inicial (base tratada)")
        df_hist = ler_tabela(WS_HIST, tuple(COLS_HIST)).copy()
    except Exception:
        pass

for c in ["Total", "Aprovados", "Nao_Aprovados", "Nao_Avaliados"]:
    df_hist[c] = pd.to_numeric(df_hist[c], errors="coerce").fillna(0).astype(int)

ESCOLAS = sorted(df_cons["Escola"].unique())

# ============================================================
# NAVEGAÇÃO
# ============================================================
ABA_VISITAS = "Registro de Visitas & Campo"
opcoes_menu = [
    "Visão Geral",
    "Escolas Prioritárias (25)",
    "Todas as Escolas",
    ABA_VISITAS,
    "Detalhamento por Componentes",
    "📁 Atualizar Base de Dados",
    "Relatório Executivo / WhatsApp",
]
KEY_ESC_FORM = "sel_escola_form"
KEY_ESC_CAD = "sel_escola_caderneta"

st.session_state.setdefault("aba_ativa", "Visão Geral")
st.session_state.setdefault("escola_selecionada_visita", "Selecione...")


def ir_para_visita(nome_escola):
    """Callback do ✈️: roda ANTES do próximo rerun, então os widgets já nascem com a escola."""
    st.session_state["escola_selecionada_visita"] = nome_escola
    st.session_state[KEY_ESC_FORM] = nome_escola
    st.session_state[KEY_ESC_CAD] = nome_escola
    st.session_state["aba_ativa"] = ABA_VISITAS


def sincronizar_caderneta():
    esc = st.session_state.get(KEY_ESC_CAD, "Selecione...")
    st.session_state["escola_selecionada_visita"] = esc
    st.session_state[KEY_ESC_FORM] = esc


st.sidebar.markdown("### 🎓 1ª DIREC - Natal/RN")
st.sidebar.markdown("<p style='font-size: 12px; color: #64748b;'>Monitoramento de Ciclos, Visitas e Plataformas</p>",
                    unsafe_allow_html=True)
st.sidebar.markdown("---")
aba = st.sidebar.radio("Navegação", opcoes_menu, key="aba_ativa")

st.sidebar.markdown("---")
st.sidebar.markdown("🔒 **Área de Gestão / Trava**")
senha_admin = st.sidebar.text_input("Senha de Administrador", type="password", key="key_admin_pass")
try:
    senha_ok = st.secrets.get("admin_password", SENHA_ADMIN)
except Exception:
    senha_ok = SENHA_ADMIN
st.session_state.is_admin = senha_admin == senha_ok
if st.session_state.is_admin:
    st.sidebar.success("🔑 Modo Admin Ativo")
else:
    st.sidebar.caption("🔒 *Exclusões e resets bloqueados.*")

# --- FILTROS GLOBAIS (aplicados em todas as abas) ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 Filtros Globais")
filtro_assessor = st.sidebar.selectbox("Assessor(a)", ["Todos(as)"] + sorted(df_cons["Assessor(a)"].dropna().unique()))
filtro_tutor = st.sidebar.selectbox("Professor-Tutor", ["Todos(as)"] + sorted(df_cons["Professor tutor"].dropna().unique()))
filtro_escola = st.sidebar.selectbox("Escola", ["Todos(as)"] + ESCOLAS)
filtro_plataforma = st.sidebar.selectbox("Plataforma", ["Todas (Plural & Avalia)", "Plural (Redash)", "Avalia"])

st.sidebar.markdown("---")
if gsheets_configurado() and ARMAZENAMENTO_OK:
    st.sidebar.caption("🟢 Google Sheets conectado")
elif gsheets_configurado():
    st.sidebar.caption(f"🔴 Erro no Google Sheets: {ERRO_ARMAZENAMENTO[:120]}")
else:
    st.sidebar.caption("🟡 Modo local (configure o Google Sheets em secrets para salvar de forma permanente)")
st.sidebar.caption(f"📦 {ORIGEM_BASE}")


def filtrar(df, col_plat="Plataforma"):
    out = df
    if filtro_assessor != "Todos(as)":
        out = out[out["Assessor(a)"] == filtro_assessor]
    if filtro_tutor != "Todos(as)":
        out = out[out["Professor tutor"] == filtro_tutor]
    if filtro_escola != "Todos(as)":
        out = out[out["Escola"] == filtro_escola]
    if filtro_plataforma == "Plural (Redash)":
        out = out[out[col_plat] == "Plurall"]
    elif filtro_plataforma == "Avalia":
        out = out[out[col_plat] == "Avalia"]
    return out


cons_f = filtrar(df_cons)
comp_f = filtrar(df_comp)
alunos_f = consolidar_alunos(cons_f)                 # estudantes únicos (filtrados)
escolas_f = set(cons_f["Escola"].unique())


def contar(alunos):
    vc = alunos["Situação do aluno"].value_counts()
    return len(alunos), int(vc.get(AP, 0)), int(vc.get(NAP, 0)), int(vc.get(NAV, 0))


def evolucao_filtrada():
    """Série temporal (último marco de cada data) somando as escolas do filtro atual."""
    if df_hist.empty:
        return pd.DataFrame()
    h = df_hist.copy()
    if filtro_assessor != "Todos(as)":
        h = h[h["Assessor"] == filtro_assessor]
    if filtro_tutor != "Todos(as)":
        h = h[h["Professor_tutor"] == filtro_tutor]
    if filtro_escola != "Todos(as)":
        h = h[h["Escola"] == filtro_escola]
    if h.empty:
        return pd.DataFrame()
    ultimo = h.sort_values("Data_Registro").groupby("Data_Referencia")["Import_ID"].last()
    h = h[h["Import_ID"].isin(ultimo.values)]
    ev = (h.groupby("Data_Referencia")[["Total", "Aprovados", "Nao_Aprovados", "Nao_Avaliados"]]
            .sum().reset_index().sort_values("Data_Referencia"))
    ev["Data"] = pd.to_datetime(ev["Data_Referencia"])
    return ev


def visitas_filtradas():
    if df_visitas.empty:
        return df_visitas
    return df_visitas[df_visitas["Escola"].isin(escolas_f)]


def resumo_escolas(cons_rows, alunos):
    if alunos.empty:
        return pd.DataFrame()
    base = alunos.groupby("Escola").agg(
        Tutor=("Professor tutor", "first"),
        Assessor=("Assessor(a)", "first"),
        Total_Alunos=("Matrícula", "nunique"),
        Aprovados=("Situação do aluno", lambda x: (x == AP).sum()),
        Nao_Aprovados=("Situação do aluno", lambda x: (x == NAP).sum()),
        Nao_Avaliados=("Situação do aluno", lambda x: (x == NAV).sum()),
    )
    plat = cons_rows[cons_rows["Situação do aluno"] == NAV].groupby(["Escola", "Plataforma"]).size().unstack(fill_value=0)
    base["Nao_Avaliados_Plural"] = plat.get("Plurall", pd.Series(dtype=int)).reindex(base.index).fillna(0).astype(int)
    base["Nao_Avaliados_Avalia"] = plat.get("Avalia", pd.Series(dtype=int)).reindex(base.index).fillna(0).astype(int)
    base["Total_Pendentes"] = base["Nao_Aprovados"] + base["Nao_Avaliados"]
    v = visitas_filtradas()
    if len(v):
        vg = v.groupby("Escola").agg(Visitas=("Escola", "size"), Ultima_Visita=("_data", "max"))
        base = base.join(vg)
    else:
        base["Visitas"], base["Ultima_Visita"] = 0, pd.NaT
    base["Visitas"] = base["Visitas"].fillna(0).astype(int)
    return base.reset_index()


def lista_escolas_com_botao(tabela, prefixo):
    for i, row in enumerate(tabela.itertuples(index=False)):
        col_info, col_btn = st.columns([5, 1])
        ult = "" if pd.isna(row.Ultima_Visita) else f" | 🚐 {row.Visitas} visita(s), última {row.Ultima_Visita:%d/%m}"
        col_info.markdown(
            f"**{row.Escola}** | Tutor: {row.Tutor} | Assessor(a): {row.Assessor} | "
            f"Total Alunos: **{row.Total_Alunos}** | Aprovados: **{row.Aprovados}** | "
            f"Não Avaliados Plural: **{row.Nao_Avaliados_Plural}** | Não Avaliados Avalia: **{row.Nao_Avaliados_Avalia}** | "
            f"Não Aprovados: **{row.Nao_Aprovados}** | 🔴 **Pendentes: {row.Total_Pendentes}**{ult}"
        )
        col_btn.button("✈️ Visitar", key=f"{prefixo}_{i}_{key_escola(row.Escola)[:40]}",
                       on_click=ir_para_visita, args=(row.Escola,))


# ==========================================
# ABA 1: VISÃO GERAL
# ==========================================
if aba == "Visão Geral":
    st.title("📊 Painel de Acompanhamento Acadêmico")
    st.markdown("Monitoramento de Ciclos, Visitas e Plataformas (Plural & Avalia)")
    if len(df_imports):
        ult = df_imports.sort_values("Data_Importacao").iloc[-1]
        st.caption(f"🕒 Última atualização da base: {fmt_data_br(ult['Data_Referencia'])} "
                   f"({ult['Fonte']}, importada em {ult['Data_Importacao']})")
    st.markdown("---")

    total, ap, nap, nav = contar(alunos_f)
    ev = evolucao_filtrada()
    deltas = [None] * 4
    if len(ev) >= 1 and filtro_plataforma.startswith("Todas"):
        ini = ev.iloc[0]
        deltas = [None, ap - ini["Aprovados"], nap - ini["Nao_Aprovados"], nav - ini["Nao_Avaliados"]]
        st.caption(f"Variações calculadas em relação ao primeiro marco ({ini['Data']:%d/%m/%Y}).")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Estudantes", f"{total:,}".replace(",", "."))
    c2.metric("Aprovados / Liberados", f"{ap:,}".replace(",", "."),
              delta=None if deltas[1] is None else int(deltas[1]))
    c3.metric("Não Aprovados", f"{nap:,}".replace(",", "."),
              delta=None if deltas[2] is None else int(deltas[2]), delta_color="inverse")
    c4.metric("Não Avaliados", f"{nav:,}".replace(",", "."),
              delta=None if deltas[3] is None else int(deltas[3]), delta_color="inverse")
    st.caption("Liberado = estudante sem nenhuma pendência em nenhuma plataforma. "
               "Estudantes presentes no Plurall e no Avalia são contados uma única vez.")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Status por Situação (Estudantes Únicos)")
        df_status = alunos_f["Situação do aluno"].value_counts().reindex(ORDEM_SIT, fill_value=0).reset_index()
        df_status.columns = ["Situação", "Quantidade"]
        fig = px.pie(df_status, names="Situação", values="Quantidade", hole=0.5,
                     color="Situação", color_discrete_map=CORES_SIT)
        fig.update_layout(margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig, use_container_width=True)

    with col_g2:
        st.subheader("Fluxo das Visitas de Campo")
        v = visitas_filtradas()
        if v.empty:
            st.info("Nenhuma visita registrada até o momento.")
        else:
            vd = (v.dropna(subset=["_data"]).groupby(v["_data"].dt.date)
                    .agg(Visitas=("Escola", "size"), Atendidos=("Atendidos", "sum")).reset_index())
            vd.columns = ["Data", "Visitas", "Atendidos"]
            vd["Visitas acumuladas"] = vd["Visitas"].cumsum()
            fig = go.Figure()
            fig.add_bar(x=vd["Data"], y=vd["Atendidos"], name="Alunos atendidos", marker_color="#d97706")
            fig.add_scatter(x=vd["Data"], y=vd["Visitas acumuladas"], name="Visitas acumuladas",
                            mode="lines+markers", yaxis="y2", line=dict(color="#334155"))
            fig.update_layout(margin=dict(t=20, b=20, l=20, r=20), legend=dict(orientation="h", y=-0.2),
                              yaxis=dict(title="Atendidos"),
                              yaxis2=dict(title="Visitas", overlaying="y", side="right", rangemode="tozero"),
                              xaxis=dict(tickformat="%d/%m"))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("📈 Evolução da Situação dos Estudantes por Data")
    if not filtro_plataforma.startswith("Todas"):
        st.caption("A evolução é registrada com as duas plataformas somadas (filtro de plataforma não se aplica aqui).")
    if ev.empty:
        st.info("Ainda não há marcos registrados. Eles são criados a cada importação de planilha "
                "na aba 📁 Atualizar Base de Dados.")
    else:
        ev_long = ev.melt(id_vars=["Data"], value_vars=["Aprovados", "Nao_Aprovados", "Nao_Avaliados"],
                          var_name="Situação", value_name="Estudantes")
        ev_long["Situação"] = ev_long["Situação"].map({"Aprovados": AP, "Nao_Aprovados": NAP, "Nao_Avaliados": NAV})
        fig = px.line(ev_long, x="Data", y="Estudantes", color="Situação", markers=True, text="Estudantes",
                      color_discrete_map=CORES_SIT)
        fig.update_traces(textposition="top center")
        v = visitas_filtradas()
        for d in sorted(v["_data"].dropna().dt.normalize().unique()) if len(v) else []:
            fig.add_vline(x=pd.Timestamp(d).timestamp() * 1000, line_dash="dot", line_color="#94a3b8")
        fig.update_layout(margin=dict(t=20, b=20, l=20, r=20), xaxis=dict(tickformat="%d/%m"),
                          legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Linhas pontilhadas = dias com visita registrada.")

        tab = ev[["Data", "Total", "Aprovados", "Nao_Aprovados", "Nao_Avaliados"]].copy()
        tab["Δ Aprovados"] = tab["Aprovados"].diff().fillna(0).astype(int)
        v = visitas_filtradas()
        if len(v):
            datas = tab["Data"].tolist()
            vis = []
            for i, d in enumerate(datas):
                ini = datas[i - 1] if i else pd.Timestamp.min
                vis.append(int(((v["_data"] > ini) & (v["_data"] <= d)).sum()))
            tab["Visitas no período"] = vis
        tab["Data"] = tab["Data"].dt.strftime("%d/%m/%Y")
        tab = tab.rename(columns={"Nao_Aprovados": "Não aprovados", "Nao_Avaliados": "Não avaliados"})
        st.dataframe(tab, hide_index=True, use_container_width=True)

    log_f = df_log[df_log["Escola"].isin(escolas_f)] if len(df_log) else df_log
    if len(log_f):
        novos_ap = log_f[log_f["Situacao_nova"] == AP].copy()
        with st.expander(f"🎉 Estudantes que passaram a constar como Aprovados após atualizações ({len(novos_ap)})"):
            novos_ap["Data_Referencia"] = novos_ap["Data_Referencia"].map(fmt_data_br)
            st.dataframe(novos_ap[["Data_Referencia", "Escola", "Matrícula", "Nome", "Plataforma",
                                   "Situacao_anterior", "Situacao_nova"]],
                         hide_index=True, use_container_width=True)

# ==========================================
# ABA 2: ESCOLAS PRIORITÁRIAS (25)
# ==========================================
elif aba == "Escolas Prioritárias (25)":
    st.markdown("""
        <div style="background: linear-gradient(to right, #f59e0b, #d97706); padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h2 style="margin: 0; font-size: 22px;">Top 25 Escolas Prioritárias</h2>
            <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Escolas com maior volume de estudantes pendentes que exigem foco imediato de visitação de campo.</p>
        </div>
    """, unsafe_allow_html=True)
    resumo = resumo_escolas(cons_f, alunos_f)
    if resumo.empty:
        st.info("Nenhuma escola para os filtros selecionados.")
    else:
        top = resumo.sort_values(["Total_Pendentes", "Total_Alunos"], ascending=False).head(25)
        lista_escolas_com_botao(top, "prio")

# ==========================================
# ABA 3: TODAS AS ESCOLAS
# ==========================================
elif aba == "Todas as Escolas":
    st.title("🏫 Todas as Escolas da Rede (1ª DIREC)")
    st.markdown("Panorama completo de todas as unidades monitoradas na rede com conferência de quantitativos.")
    busca = st.text_input("🔍 Buscar escola pelo nome...", "")
    resumo = resumo_escolas(cons_f, alunos_f)
    if not resumo.empty:
        if busca:
            resumo = resumo[resumo["Escola"].map(norm_txt).str.contains(norm_txt(busca), na=False)]
        st.caption(f"{len(resumo)} escola(s)")
        lista_escolas_com_botao(resumo.sort_values("Escola"), "geral")

# ==========================================
# ABA 4: REGISTRO DE VISITAS & CAMPO
# ==========================================
elif aba == ABA_VISITAS:
    st.title("🚐 Caderneta de Campo & Atendimento do Articulador")
    st.markdown("Preencha o formulário de visita técnica para registrar os atendimentos nas escolas.")
    st.markdown("---")

    opcoes_escola = ["Selecione..."] + ESCOLAS
    esc_sel = st.session_state.get("escola_selecionada_visita", "Selecione...")
    if esc_sel not in opcoes_escola:
        esc_sel = "Selecione..."
    if esc_sel == "Selecione..." and filtro_escola in opcoes_escola:
        esc_sel = filtro_escola  # usa a escola do filtro lateral, se houver
    # Se o widget foi "limpo" (ao trocar de aba), ele renasce com a última escola escolhida
    for k in (KEY_ESC_FORM, KEY_ESC_CAD):
        if st.session_state.get(k) not in opcoes_escola:
            st.session_state[k] = esc_sel

    articulador_padrao = MAPA_ART.get(st.session_state[KEY_ESC_FORM], {}).get("Articulador(a)", "")

    with st.form("form_visita_manual", clear_on_submit=False):
        st.subheader("Registrar Visita Escolar & Atendimento de Campo")
        col_m1, col_m2 = st.columns(2)
        escola_manual = col_m1.selectbox("Escola Visitada", opcoes_escola, key=KEY_ESC_FORM)
        data_visita = col_m2.date_input("Data da Visita", value=agora().date(), format="DD/MM/YYYY")
        col_m3, col_m4 = st.columns(2)
        turno_manual = col_m3.selectbox("Turno da Visita", ["Matutino", "Vespertino", "Noturno"])
        articulador = col_m4.text_input("Articulador(a)", value=articulador_padrao)
        col_m5, col_m6 = st.columns(2)
        atendidos_manual = col_m5.number_input("Alunos Atendidos na Visita", min_value=0, value=0, step=1)
        aprovados_manual = col_m6.number_input("Alunos Aprovados / Liberados", min_value=0, value=0, step=1)
        registrado_por = st.text_input("Registrado por")
        obs = st.text_area("Observações", height=80)
        submitted = st.form_submit_button("💾 Salvar Visita", type="primary")
        if submitted:
            if escola_manual == "Selecione...":
                st.error("Selecione a escola visitada.")
            else:
                try:
                    acrescentar_linhas(WS_VISITAS, COLS_VISITAS, [{
                        "ID": uuid.uuid4().hex[:10], "Data_Visita": data_visita.isoformat(),
                        "Data_Registro": agora().strftime("%Y-%m-%d %H:%M:%S"),
                        "Escola": escola_manual, "Turno": turno_manual, "Articulador": articulador,
                        "Registrado_por": registrado_por, "Atendidos": int(atendidos_manual),
                        "Aprovados_Liberados": int(aprovados_manual), "Presentes": "", "Observacoes": obs,
                    }])
                    st.success(f"✅ Visita à escola '{escola_manual}' em {data_visita:%d/%m/%Y} salva na planilha.")
                    st.session_state["escola_selecionada_visita"] = escola_manual
                except Exception as e:
                    st.error(f"Não foi possível salvar a visita: {e}")

    st.markdown("---")
    st.subheader("Caderneta Nominal de Estudantes da Unidade")
    escola_visitada = st.selectbox("Selecione a Escola para carregar a lista nominal:", opcoes_escola,
                                   key=KEY_ESC_CAD, on_change=sincronizar_caderneta)

    if escola_visitada != "Selecione...":
        cons_esc = df_cons[df_cons["Escola"] == escola_visitada]
        alunos_esc = consolidar_alunos(cons_esc)
        t, a_, na_, nv_ = contar(alunos_esc)
        st.info(f"📊 **Resumo da Unidade ({escola_visitada}):** Total Estudantes Únicos: **{t}** | "
                f"Aprovados: **{a_}** | Não Aprovados: **{na_}** | Não Avaliados: **{nv_}**")

        st.markdown("##### Filtrar exibição nominal por situação:")
        c_ap, c_nap, c_nav = st.columns(3)
        status_permitidos = []
        if c_ap.checkbox(AP, value=False):
            status_permitidos.append(AP)
        if c_nap.checkbox(NAP, value=True):
            status_permitidos.append(NAP)
        if c_nav.checkbox(NAV, value=True):
            status_permitidos.append(NAV)

        df_alunos_escola = alunos_esc[alunos_esc["Situação do aluno"].isin(status_permitidos)].sort_values("Nome")
        if len(df_alunos_escola) > 0:
            st.markdown("##### 📋 Chamada / Presença de Campo:")
            df_chamada = df_alunos_escola[["Matrícula", "Nome", "Série", "Plataforma", "Situação do aluno", "Pendências"]].copy()
            df_chamada["Presente"] = False
            df_chamada["Ausente"] = False
            df_editado = st.data_editor(
                df_chamada,
                column_config={
                    "Presente": st.column_config.CheckboxColumn("Presente", default=False),
                    "Ausente": st.column_config.CheckboxColumn("Ausente", default=False),
                    "Nome": st.column_config.TextColumn("Nome Completo"),
                    "Situação do aluno": st.column_config.TextColumn("Situação Oficial"),
                    "Pendências": st.column_config.TextColumn("Componentes pendentes", width="large"),
                },
                disabled=["Matrícula", "Nome", "Série", "Plataforma", "Situação do aluno", "Pendências"],
                hide_index=True, use_container_width=True,
                key=f"editor_{key_escola(escola_visitada)[:40]}_{'-'.join(status_permitidos)}",
            )
            cf1, cf2 = st.columns(2)
            data_freq = cf1.date_input("Data da frequência", value=agora().date(), format="DD/MM/YYYY", key="data_freq")
            criar_visita = cf2.checkbox("Registrar também como visita (Atendidos = presentes)", value=True)
            if st.button("💾 Salvar Frequência da Visita"):
                conflito = df_editado[df_editado["Presente"] & df_editado["Ausente"]]
                if len(conflito):
                    st.error(f"{len(conflito)} estudante(s) marcados como Presente e Ausente ao mesmo tempo. Corrija antes de salvar.")
                else:
                    marcados = df_editado[df_editado["Presente"] | df_editado["Ausente"]]
                    if marcados.empty:
                        st.warning("Nenhum estudante marcado.")
                    else:
                        id_visita = uuid.uuid4().hex[:10]
                        reg = agora().strftime("%Y-%m-%d %H:%M:%S")
                        try:
                            acrescentar_linhas(WS_FREQ, COLS_FREQ, [{
                                "ID_Visita": id_visita, "Data_Visita": data_freq.isoformat(),
                                "Escola": escola_visitada, "Matrícula": r["Matrícula"], "Nome": r["Nome"],
                                "Plataforma": r["Plataforma"], "Situação": r["Situação do aluno"],
                                "Frequencia": "Presente" if r["Presente"] else "Ausente", "Data_Registro": reg,
                            } for _, r in marcados.iterrows()])
                            presentes = int(df_editado["Presente"].sum())
                            if criar_visita:
                                acrescentar_linhas(WS_VISITAS, COLS_VISITAS, [{
                                    "ID": id_visita, "Data_Visita": data_freq.isoformat(), "Data_Registro": reg,
                                    "Escola": escola_visitada, "Turno": "",
                                    "Articulador": MAPA_ART.get(escola_visitada, {}).get("Articulador(a)", ""),
                                    "Registrado_por": "Caderneta", "Atendidos": presentes,
                                    "Aprovados_Liberados": 0, "Presentes": presentes,
                                    "Observacoes": "Registrada pela caderneta nominal",
                                }])
                            st.success(f"Frequência registrada! **{presentes}** estudantes marcados como Presentes.")
                        except Exception as e:
                            st.error(f"Não foi possível salvar a frequência: {e}")
        else:
            st.caption("Nenhum estudante com as situações selecionadas.")

    st.markdown("---")
    st.subheader("Histórico de Visitas Registradas (Sincronizado)")
    vh = df_visitas.copy()
    if escola_visitada != "Selecione..." and st.checkbox("Mostrar só a escola selecionada", value=False):
        vh = vh[vh["Escola"] == escola_visitada]
    if len(vh):
        vh = vh.sort_values(["_data", "Data_Registro"], ascending=False)
        vh["Data da visita"] = vh["_data"].dt.strftime("%d/%m/%Y")
        st.dataframe(vh[["Data da visita", "Escola", "Turno", "Articulador", "Atendidos",
                         "Aprovados_Liberados", "Presentes", "Registrado_por", "Observacoes", "Data_Registro"]],
                     hide_index=True, use_container_width=True)
    else:
        st.markdown("<p style='color: #94a3b8; font-size: 13px;'>Nenhuma visita registrada na planilha até o momento.</p>",
                    unsafe_allow_html=True)

# ==========================================
# ABA 5: DETALHAMENTO POR COMPONENTES
# ==========================================
elif aba == "Detalhamento por Componentes":
    st.title("📚 Detalhamento de Componentes Curriculares por Escola")
    st.markdown("Visualização por disciplina respeitando a regra de multiplicidade (múltiplos componentes por aluno).")

    escolas_comp = sorted(comp_f["Escola"].unique())
    if not escolas_comp:
        st.info("Nenhuma escola para os filtros selecionados.")
        st.stop()
    idx = escolas_comp.index(filtro_escola) if filtro_escola in escolas_comp else 0
    escola_comp = st.selectbox("Selecione a Escola:", escolas_comp, index=idx)
    dce = comp_f[comp_f["Escola"] == escola_comp].copy()

    q_total = len(dce)
    vc = dce["Situação"].value_counts()
    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Componentes totais", q_total)
    mc2.metric("Aprovados", int(vc.get(AP, 0)))
    mc3.metric("Não aprovados", int(vc.get(NAP, 0)))
    mc4.metric("Não avaliados", int(vc.get(NAV, 0)))

    graf = dce.groupby(["Componente", "Situação"]).size().reset_index(name="Qtd")
    fig = px.bar(graf, y="Componente", x="Qtd", color="Situação", orientation="h",
                 color_discrete_map=CORES_SIT, category_orders={"Situação": ORDEM_SIT})
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=max(260, 28 * graf["Componente"].nunique()),
                      legend=dict(orientation="h", y=-0.15), yaxis_title=None)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### Filtrar componentes:")
    col_f1, col_f2, col_f3 = st.columns(3)
    sits = []
    if col_f1.checkbox("Exibir Aprovados", value=True, key="f_c_ap"):
        sits.append(AP)
    if col_f2.checkbox("Exibir Não Aprovados", value=True, key="f_c_nap"):
        sits.append(NAP)
    if col_f3.checkbox("Exibir Não avaliados", value=True, key="f_c_nav"):
        sits.append(NAV)
    cf1, cf2 = st.columns(2)
    comps_sel = cf1.multiselect("Componente", sorted(dce["Componente"].unique()))
    busca_aluno = cf2.text_input("Buscar estudante (nome ou matrícula)")

    filtro = dce["Situação"].isin(sits)
    if comps_sel:
        filtro &= dce["Componente"].isin(comps_sel)
    if busca_aluno:
        b = norm_txt(busca_aluno)
        filtro &= dce["Nome"].map(norm_txt).str.contains(b, na=False) | dce["Matrícula"].str.contains(b, na=False)
    exib = dce[filtro].sort_values(["Nome", "Componente"])

    st.info(f"Exibindo {len(exib)} de {q_total} componentes da unidade selecionada.")
    st.dataframe(
        exib[["Matrícula", "Nome", "Série", "Componente", "Plataforma", "Rendimento", "Situação",
              "Status plataforma", "Atualizado em"]],
        hide_index=True, use_container_width=True,
        column_config={"Rendimento": st.column_config.TextColumn("Rendimento (Plurall %) / Nota (Avalia)")},
    )
    st.download_button("⬇️ Baixar lista (CSV)", exib.to_csv(index=False, sep=";").encode("utf-8-sig"),
                       file_name=f"componentes_{key_escola(escola_comp)[:30].replace(' ', '_')}.csv", mime="text/csv")

# ==========================================
# ABA 6: ATUALIZAR BASE DE DADOS
# ==========================================
elif aba == "📁 Atualizar Base de Dados":
    st.title("📁 Importar / Atualizar Planilha de Dados")
    st.info(f"🕒 **Acesso à base:** {agora():%d/%m/%Y às %Hh%Mmin} · **Origem atual:** {ORIGEM_BASE}")
    st.markdown(
        "Envie a planilha do **Plurall (Redash)** — *SEEC-RN_-_Rendimento_e_participação...* — e/ou do "
        "**Avalia** — *relatorio_acompanhamento_rapp_ept_div* (pode enviar as duas juntas). "
        "Cada aluno x componente presente no arquivo é atualizado; o que não aparece no arquivo é mantido como estava."
    )

    if "ultimo_relatorio" in st.session_state:
        rel = st.session_state["ultimo_relatorio"]
        mud = rel["mudancas"]
        st.success(f"✅ Atualização de {fmt_data_br(rel['data_ref'])} ({rel['fonte']}) aplicada em todas as abas.")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Linhas lidas", rel["linhas_lidas"])
        r2.metric("Componentes atualizados", rel["atualizados"])
        r3.metric("Alunos que mudaram de situação", len(mud))
        r4.metric("Novos aprovados", int((mud["Depois"] == AP).sum()) if len(mud) else 0)
        if rel["novos_comp"]:
            st.caption(f"ℹ️ {rel['novos_comp']} componente(s) que não constavam na base para o aluno foram incluídos.")
        if len(mud):
            with st.expander("Ver alunos que mudaram de situação", expanded=True):
                st.dataframe(mud[["Escola", "Matrícula", "Nome", "Plataforma", "Antes", "Depois"]].sort_values(["Escola", "Nome"]),
                             hide_index=True, use_container_width=True)
            regrediu = mud[(mud["Antes"] == AP) & (mud["Depois"] != AP)]
            if len(regrediu):
                st.warning(f"⚠️ {len(regrediu)} aluno(s) deixaram de constar como Aprovados — vale conferir na plataforma.")
        if len(rel["fora"]):
            with st.expander(f"⚠️ {len(rel['fora'])} registro(s) de estudantes que não estão na base oficial (não importados)"):
                st.dataframe(rel["fora"], hide_index=True, use_container_width=True)
        if st.button("Fechar relatório"):
            del st.session_state["ultimo_relatorio"]
            st.rerun()
        st.markdown("---")

    arquivos = st.file_uploader("Escolha a(s) planilha(s) (.xlsx / .csv)", type=["xlsx", "csv"],
                                accept_multiple_files=True, key=f"up_{st.session_state.get('upload_n', 0)}")
    lidos, base_cheia = [], None
    datas_arquivo = []
    for arq in arquivos or []:
        try:
            if arq.name.lower().endswith(".csv"):
                df_up, abas = pd.read_csv(arq, sep=None, engine="python"), None
            else:
                xls = pd.ExcelFile(arq)
                abas = xls.sheet_names
                df_up = pd.read_excel(xls, sheet_name=0)
            tipo = identificar_planilha(df_up, abas)
            if tipo in ("Plurall", "Avalia"):
                lidos.append((arq.name, tipo, df_up))
                st.write(f"✔️ **{arq.name}** → planilha do **{tipo}** ({len(df_up)} linhas)")
                m = re.search(r"(20\d{2})[_\-](\d{2})[_\-](\d{2})", arq.name)
                if m:
                    datas_arquivo.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
            elif tipo == "Base consolidada":
                base_cheia = (arq.name, pd.read_excel(xls, sheet_name="Base componentes", dtype={"Matrícula": str}))
                st.write(f"✔️ **{arq.name}** → **Base consolidada** (aba Base componentes)")
            else:
                st.error(f"❌ {arq.name}: formato não reconhecido. Colunas: {', '.join(map(str, df_up.columns[:10]))}")
        except Exception as ex:
            st.error(f"Erro ao ler {arq.name}: {ex}")

    if lidos:
        if datas_arquivo:
            st.caption(f"Data encontrada no nome do arquivo: {max(datas_arquivo):%d/%m/%Y}")
        data_ref = st.date_input("Data de referência desta atualização (vai para o gráfico de evolução)",
                                 value=agora().date(), format="DD/MM/YYYY")
        if st.button("🔄 Processar e atualizar painel", type="primary"):
            with st.spinner("Cruzando dados e salvando..."):
                try:
                    rel = processar_importacao(lidos, data_ref, df_comp, df_cons, MAPA_ART)
                    st.session_state["ultimo_relatorio"] = rel
                    st.session_state["upload_n"] = st.session_state.get("upload_n", 0) + 1  # limpa o uploader
                    st.rerun()
                except Exception as ex:
                    st.error(f"Erro ao processar: {ex}")

    if base_cheia:
        if not st.session_state.is_admin:
            st.warning("🔒 Substituir a base consolidada inteira exige o modo administrador.")
        elif st.button("♻️ Substituir a base inteira por esta planilha consolidada"):
            nova = padronizar_comp(base_cheia[1])
            sobrescrever_base(nova)
            recarregar_tudo()
            novo_comp, novo_cons, _, _ = obter_dados()
            registrar_marco(novo_cons, agora().date().isoformat(), "Base consolidada substituída")
            st.session_state["upload_n"] = st.session_state.get("upload_n", 0) + 1
            st.rerun()

    st.markdown("---")
    st.subheader("Marcos e histórico")
    cm1, cm2 = st.columns([2, 1])
    data_marco = cm1.date_input("Registrar a situação atual como marco na data:", value=agora().date(),
                                format="DD/MM/YYYY", key="data_marco")
    if cm2.button("📌 Registrar marco"):
        registrar_marco(df_cons, data_marco.isoformat(), "Marco manual")
        st.success("Marco registrado.")
        st.rerun()
    if len(df_imports):
        st.markdown("##### Importações realizadas")
        st.dataframe(df_imports.sort_values("Data_Importacao", ascending=False), hide_index=True, use_container_width=True)

    st.markdown("##### Baixar base atual")
    st.download_button("⬇️ Base de componentes atual (CSV)",
                       df_comp[COMP_COLS].to_csv(index=False, sep=";").encode("utf-8-sig"),
                       file_name=f"base_componentes_{agora():%Y%m%d}.csv", mime="text/csv")

    if st.session_state.is_admin:
        st.markdown("---")
        st.markdown("##### 🔑 Administração")
        if st.button("⚠️ Restaurar base inicial (descarta atualizações salvas)"):
            apagar_base_persistida()
            recarregar_tudo()
            st.success("Base inicial restaurada.")
            st.rerun()

# ==========================================
# ABA 7: RELATÓRIO EXECUTIVO / WHATSAPP
# ==========================================
elif aba == "Relatório Executivo / WhatsApp":
    st.title("📄 Gerador de Resumo Executivo para WhatsApp")
    agora_txt = f"{agora():%d/%m/%Y às %Hh%M}"
    st.info(f"🕒 **Última verificação / Acesso à base:** {agora_txt}")

    total, ap, nap, nav = contar(alunos_f)
    pct = lambda x: f"{(100 * x / total):.1f}%".replace(".", ",") if total else "0%"
    v = visitas_filtradas()
    ev = evolucao_filtrada()
    linha_evol = ""
    if len(ev) >= 2:
        ini, fim = ev.iloc[0], ev.iloc[-1]
        linha_evol = (f"\n📈 *Evolução {ini['Data']:%d/%m} → {fim['Data']:%d/%m}:* "
                      f"Aprovados {int(ini['Aprovados'])} → {int(fim['Aprovados'])} "
                      f"({int(fim['Aprovados'] - ini['Aprovados']):+d})")
    resumo = resumo_escolas(cons_f, alunos_f)
    top5 = ""
    if not resumo.empty:
        t5 = resumo.sort_values("Total_Pendentes", ascending=False).head(5)
        top5 = "\n".join(f"  • {r.Escola.title()}: {r.Total_Pendentes} pendentes" for r in t5.itertuples())

    texto_whatsapp = f"""📊 *BOLETIM EXECUTIVO DE REDE - RAPP 2026 (1ª DIREC)*
🕒 *Atualizado em:* {agora_txt}

🏫 *Escolas Monitoradas:* {cons_f['Escola'].nunique()} unidades
👨‍🎓 *Total de Estudantes Únicos:* {total}
✅ *Aprovados/Liberados:* {ap} ({pct(ap)})
❌ *Não aprovados:* {nap} ({pct(nap)})
⏳ *Não avaliados:* {nav} ({pct(nav)}){linha_evol}

🚐 *Visitas Registradas:* {len(v)} visitas · {int(v['Atendidos'].sum()) if len(v) else 0} alunos atendidos

🔴 *Escolas com mais pendências:*
{top5}

📌 *Status Operacional:* acompanhamento de campo em andamento com foco nas unidades prioritárias e conclusão das avaliações Plural e Avalia."""

    st.text_area("Copie o texto abaixo para enviar no grupo de gestão do WhatsApp:", texto_whatsapp, height=380)
    st.success("Pronto! Basta selecionar o texto acima, copiar e colar no WhatsApp.")
