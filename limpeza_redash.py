"""
Limpeza diária do Redash (Plurall) — RAPP 2026 · 1ª DIREC

Regra: componente APROVADO no 1º ciclo continua aprovado, mesmo que o Redash do
2º ciclo mostre "Não aprovado" (prova não realizada ou nota menor).

Uso no painel (Streamlit): aba "🧹 Limpeza do Redash".
Uso pelo terminal:
    python limpeza_redash.py Redash_do_dia.xlsx --ciclo1 Redash_extra_do_15.10.xlsx --master Planilha_master_15.10.xlsx
    (gera Redash_DD.MM_limpo.xlsx na mesma pasta)
"""

import io
import re
import sys
import unicodedata
from datetime import date

import pandas as pd

APROVADO, NAO_APROVADO, NAO_AVALIADO = "Aprovado", "Não aprovado", "Não avaliado"
NOTA_MINIMA = 60  # mesma regra do Plurall no 1º ciclo (60% ou mais = aprovado)


# ------------------------------------------------------------------ normalização
def _norm(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^A-Za-z0-9]+", " ", s).upper().split())


def _comp(v):
    k = _norm(v)
    return {"LINGUA INGLES": "LINGUA INGLESA"}.get(k, k)


def _mat(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    m = re.match(r"\s*0*(\d+)", str(v))
    return m.group(1) if m else ""


def _col(df, *nomes):
    mapa = {_norm(c): c for c in df.columns}
    for n in nomes:
        if _norm(n) in mapa:
            return mapa[_norm(n)]
    return None


def _ler(arquivo):
    """Aceita caminho, bytes ou arquivo enviado pelo Streamlit (.xlsx ou .csv)."""
    nome = getattr(arquivo, "name", str(arquivo)).lower()
    if nome.endswith(".csv"):
        return pd.read_csv(arquivo, sep=None, engine="python")
    return pd.read_excel(arquivo, sheet_name=0)


def data_do_arquivo(nome, padrao=None):
    m = re.search(r"(20\d{2})[_\-.](\d{2})[_\-.](\d{2})", str(nome))
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{2})[.\-_](\d{2})(?![\d])", str(nome))
    if m and padrao:
        try:
            return date(padrao.year, int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return padrao or date.today()


# ------------------------------------------------------------------ limpeza
def limpar_redash(df_novo, df_ciclo1, df_master=None, filtro_grupo=r"01ª DIREC|1ª DIREC|01 DIREC|NATAL"):
    """Retorna (result, corrigidos, resumo)."""
    novo = df_novo.copy()
    c1 = df_ciclo1.copy()

    # mesma DIREC
    for d in (novo, c1):
        g = _col(d, "Grupo")
        if g and filtro_grupo:
            d.drop(d.index[~d[g].astype(str).str.contains(filtro_grupo, case=False, na=False)], inplace=True)

    # chaves (matrícula a partir do e-mail; nome+escola quando não há e-mail)
    def chaves(d):
        ce, ca, cs, cp = _col(d, "email_aluno"), _col(d, "aluno"), _col(d, "escola"), _col(d, "prova")
        d["_mk"] = d[ce].map(_mat) if ce else ""
        d["_kn"] = d[ca].map(_norm) if ca else ""
        d["_ks"] = d[cs].map(_norm) if cs else ""
        d["_ck"] = d[cp].map(_comp)
    chaves(novo)
    chaves(c1)

    ref = [c1.loc[c1["_mk"] != "", ["_mk", "_kn", "_ks"]]]
    if df_master is not None:
        m = df_master.copy()
        cm, cn, ce = _col(m, "Matrícula"), _col(m, "Nome"), _col(m, "Escola")
        if cm and cn:
            ref.append(pd.DataFrame({"_mk": m[cm].map(_mat), "_kn": m[cn].map(_norm),
                                     "_ks": m[ce].map(_norm) if ce else ""}))
    ref = pd.concat(ref).drop_duplicates()
    por_nome_escola = ref.drop_duplicates(["_kn", "_ks"], keep=False).set_index(["_kn", "_ks"])["_mk"]
    por_nome = ref.drop_duplicates("_kn", keep=False).set_index("_kn")["_mk"]
    for d in (novo, c1):
        for i in d.index[d["_mk"] == ""]:
            v = por_nome_escola.get((d.at[i, "_kn"], d.at[i, "_ks"])) or por_nome.get(d.at[i, "_kn"])
            if v:
                d.at[i, "_mk"] = v

    # situação do 1º ciclo por componente
    cs1, cr1 = _col(c1, "Situação"), _col(c1, "rendimento (%)", "rendimento")
    c1["_rend"] = pd.to_numeric(c1[cr1], errors="coerce") if cr1 else pd.NA
    if cs1:
        c1["_ap"] = c1[cs1].map(_norm) == "APROVADO"
    else:
        c1["_ap"] = c1["_rend"] >= NOTA_MINIMA
    c1v = c1[c1["_mk"] != ""]
    aprov1 = (c1v[c1v["_ap"]].sort_values("_rend", ascending=False)
              .drop_duplicates(["_mk", "_ck"]).set_index(["_mk", "_ck"])["_rend"])
    reprov1 = set(zip(c1v.loc[~c1v["_ap"], "_mk"], c1v.loc[~c1v["_ap"], "_ck"]))

    # alunos aprovados em tudo no 1º ciclo (Planilha master, Plurall)
    alunos_ap = set()
    if df_master is not None:
        m = df_master
        cm, csit, cplat = _col(m, "Matrícula"), _col(m, "Situação do aluno"), _col(m, "Plataforma")
        if cm and csit:
            ok = m[csit].map(_norm) == "APROVADO"
            if cplat:
                ok &= m[cplat].astype(str).str.contains("Plurall", case=False, na=False)
            alunos_ap = set(m.loc[ok, cm].map(_mat)) - {""}

    # situação do 2º ciclo informada pela plataforma
    cr, cpart, csit = _col(novo, "rendimento (%)", "rendimento"), _col(novo, "participação"), _col(novo, "Situação")
    rend = pd.to_numeric(novo[cr], errors="coerce") if cr else pd.Series(pd.NA, index=novo.index)
    if cpart:
        part = novo[cpart].astype(str).str.upper().isin(["TRUE", "1", "SIM", "VERDADEIRO"])
    else:
        part = pd.Series(True, index=novo.index)

    def sit_plataforma(i):
        if not part.at[i]:
            return "Não realizou"
        if csit and pd.notna(novo.at[i, csit]) and not cr:
            return APROVADO if _norm(novo.at[i, csit]) == "APROVADO" else NAO_APROVADO
        return APROVADO if pd.notna(rend.at[i]) and rend.at[i] >= NOTA_MINIMA else NAO_APROVADO

    sit_plat = pd.Series({i: sit_plataforma(i) for i in novo.index})

    situacao, motivo = [], []
    for i in novo.index:
        s, k = sit_plat.at[i], (novo.at[i, "_mk"], novo.at[i, "_ck"])
        if s == APROVADO:
            situacao.append(APROVADO); motivo.append("Aprovado no 2º ciclo")
        elif k[0] and k in aprov1.index:
            situacao.append(APROVADO); motivo.append(f"Aprovado no 1º ciclo ({aprov1[k]:g}%) — mantido")
        elif k[0] and k[0] in alunos_ap:
            situacao.append(APROVADO); motivo.append("Aluno já aprovado em todos os componentes no 1º ciclo — mantido")
        elif s == "Não realizou" and k[0] and k in reprov1:
            situacao.append(NAO_APROVADO); motivo.append("Não aprovado no 1º ciclo; prova do 2º ciclo não realizada")
        elif s == "Não realizou":
            situacao.append(NAO_AVALIADO); motivo.append("Prova do 2º ciclo não realizada")
        else:
            situacao.append(NAO_APROVADO); motivo.append("Não aprovado no 2º ciclo")

    # monta a saída: colunas originais + 4 novas (a coluna "Situação" original, se houver, é substituída)
    result = df_novo.loc[novo.index].copy()
    if csit:
        result = result.drop(columns=[csit])
    result["Situação 2º ciclo (plataforma)"] = sit_plat.values
    result["Situação"] = situacao
    result["Motivo da situação"] = motivo
    result["Corrigido?"] = [("Sim" if (s == APROVADO and p != APROVADO) else "")
                            for s, p in zip(situacao, sit_plat.values)]

    cols_corr = [c for c in [_col(result, "escola"), _col(result, "aluno"), _col(result, "email_aluno"),
                             _col(result, "prova"), _col(result, "name"), _col(result, "rendimento (%)"),
                             _col(result, "participação"), "Situação 2º ciclo (plataforma)", "Situação",
                             "Motivo da situação"] if c]
    corrigidos = result[result["Corrigido?"] == "Sim"][cols_corr]
    ord_cols = [c for c in (_col(result, "escola"), _col(result, "aluno")) if c]
    if ord_cols:
        corrigidos = corrigidos.sort_values(ord_cols)

    mk = novo["_mk"]
    resumo = pd.DataFrame({
        "Indicador": ["Linhas da planilha", "Aprovadas no 2º ciclo",
                      "Corrigidas para Aprovado (aprovação do 1º ciclo mantida)", "Não aprovado",
                      "Não avaliado (prova do 2º ciclo não realizada)", "Estudantes com algum componente corrigido",
                      "Linhas sem matrícula identificada"],
        "Quantidade": [len(result), motivo.count("Aprovado no 2º ciclo"), int((result["Corrigido?"] == "Sim").sum()),
                       situacao.count(NAO_APROVADO), situacao.count(NAO_AVALIADO),
                       int(mk[(result["Corrigido?"] == "Sim").values].nunique()), int((mk == "").sum())],
    })
    return result, corrigidos, resumo


# ------------------------------------------------------------------ Excel formatado
def gerar_excel(result, corrigidos, resumo, data_ref=None):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        result.to_excel(w, sheet_name="result", index=False)      # 1ª aba: o painel importa esta
        corrigidos.to_excel(w, sheet_name="Corrigidos", index=False)
        resumo.to_excel(w, sheet_name="Resumo", index=False)
        if data_ref:
            pd.DataFrame({"Data de referência": [data_ref.strftime("%d/%m/%Y")]}).to_excel(
                w, sheet_name="Resumo", index=False, startrow=len(resumo) + 2)
    buf.seek(0)
    wb = load_workbook(buf)
    verde = PatternFill("solid", fgColor="D1FAE5")
    cores = {APROVADO: "D1FAE5", NAO_APROVADO: "FFE4E6", NAO_AVALIADO: "FEF3C7"}
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for j, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 300)), 1):
            ws.column_dimensions[get_column_letter(j)].width = min(55, max(10, max(len(str(c.value or "")) for c in col) + 2))
    ws = wb["result"]
    cab = [c.value for c in ws[1]]
    i_corr, i_sit = cab.index("Corrigido?"), cab.index("Situação")
    for row in ws.iter_rows(min_row=2):
        if row[i_corr].value == "Sim":
            for c in row:
                c.fill = verde
        else:
            cor = cores.get(row[i_sit].value)
            if cor:
                row[i_sit].fill = PatternFill("solid", fgColor=cor)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def limpar_arquivos(arquivo_novo, arquivo_ciclo1, arquivo_master=None, data_ref=None):
    """Atalho: lê os arquivos, limpa e devolve (bytes do Excel, resumo, nome sugerido)."""
    df_novo, df_c1 = _ler(arquivo_novo), _ler(arquivo_ciclo1)
    df_m = _ler(arquivo_master) if arquivo_master is not None else None
    data_ref = data_ref or data_do_arquivo(getattr(arquivo_novo, "name", arquivo_novo), date.today())
    result, corrigidos, resumo = limpar_redash(df_novo, df_c1, df_m)
    return gerar_excel(result, corrigidos, resumo, data_ref), resumo, f"Redash_{data_ref:%d.%m}_limpo.xlsx"


# ------------------------------------------------------------------ terminal
if __name__ == "__main__":
    import argparse
    import os
    ap = argparse.ArgumentParser(description="Limpeza diária do Redash (RAPP 1ª DIREC)")
    ap.add_argument("novo", help="Redash do dia (.xlsx)")
    ap.add_argument("--ciclo1", required=True, help="Redash do 1º ciclo (ex.: Redash_extra_do_15.10.xlsx)")
    ap.add_argument("--master", help="Planilha master do 1º ciclo (opcional)")
    ap.add_argument("--saida", help="Pasta de saída (padrão: a mesma do arquivo do dia)")
    a = ap.parse_args()
    conteudo, resumo, nome = limpar_arquivos(a.novo, a.ciclo1, a.master)
    destino = os.path.join(a.saida or os.path.dirname(os.path.abspath(a.novo)), nome)
    with open(destino, "wb") as f:
        f.write(conteudo)
    print(resumo.to_string(index=False))
    print(f"\nArquivo gerado: {destino}")
    sys.exit(0)