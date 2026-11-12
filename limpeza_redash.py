"""
Limpeza diária do Redash (Plurall) — RAPP 2026 · 1ª DIREC

Regras (atualizadas em 12/11/2026):
  - nota 0, sem tempo de prova e sem data de entrega -> Não avaliado (prova não realizada),
    exceto se o componente foi NÃO APROVADO no 1º ciclo -> continua Não aprovado
  - nota de 0 a 59, com tempo de prova e data de entrega -> Não aprovado
  - nota de 60 a 100, com tempo de prova e data de entrega -> Aprovado
  - componente aprovado no 1º ciclo -> o 1º ciclo prevalece (Aprovado)

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


def _mat_email(v):
    """Matrícula a partir do e-mail: só vale se a parte antes do @ for toda numérica (8+ dígitos).
    E-mails fora do padrão (ex.: 1d.nome1234x@educar.rn.gov.br) ficam sem matrícula e são
    identificados pelo nome + escola."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    local = str(v).strip().split("@")[0]
    return local.lstrip("0") if local.isdigit() and len(local) >= 8 else ""


def _vazio(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "NaT", "-", "—")


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
        d["_mk"] = d[ce].map(_mat_email) if ce else ""
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

    # ---------------- situação do 2º ciclo (regras de 12/11) ----------------
    # 1) nota 0 SEM tempo de prova e SEM data de entrega  -> Não avaliado (prova não realizada)
    #    ... mas se foi NÃO APROVADO no 1º ciclo, continua Não aprovado
    # 2) nota 0 a 59 COM tempo de prova e data de entrega -> Não aprovado
    # 3) nota 60 a 100 COM tempo de prova e data de entrega -> Aprovado
    # 4) aprovado no 1º ciclo -> prevalece o 1º ciclo (Aprovado)
    cr, csit = _col(novo, "rendimento (%)", "rendimento"), _col(novo, "Situação")
    ctempo, cdata = _col(novo, "tempo de prova"), _col(novo, "data_entrega", "data de entrega")
    rend = pd.to_numeric(novo[cr], errors="coerce") if cr else pd.Series(pd.NA, index=novo.index)
    if ctempo and cdata:
        tem_tempo = ~novo[ctempo].map(_vazio)
        tem_data = ~novo[cdata].map(_vazio)
    else:  # planilhas antigas, sem essas colunas: usa a participação (ou considera realizada)
        cpart = _col(novo, "participação")
        part = (novo[cpart].astype(str).str.upper().isin(["TRUE", "1", "SIM", "VERDADEIRO"])
                if cpart else pd.Series(True, index=novo.index))
        tem_tempo = tem_data = part

    def sit_plataforma(i):
        r = rend.at[i]
        realizou = bool(tem_tempo.at[i]) and bool(tem_data.at[i])
        if pd.isna(r) and csit and not cr:  # formato antigo, só com a coluna Situação
            return (APROVADO if _norm(novo.at[i, csit]) == "APROVADO" else NAO_APROVADO), "ok"
        r = 0 if pd.isna(r) else r
        if realizou:
            return (APROVADO if r >= NOTA_MINIMA else NAO_APROVADO), "ok"
        if r == 0:  # só uma das duas colunas preenchida -> mantém Não avaliado, mas sinaliza
            return "Não realizou", ("conferir" if (bool(tem_tempo.at[i]) or bool(tem_data.at[i])) else "ok")
        # nota maior que 0 sem tempo/data: dado incompleto -> classifica pela nota e sinaliza
        return (APROVADO if r >= NOTA_MINIMA else NAO_APROVADO), "conferir"

    calc = {i: sit_plataforma(i) for i in novo.index}
    sit_plat = pd.Series({i: v[0] for i, v in calc.items()})
    alerta = pd.Series({i: v[1] for i, v in calc.items()})

    situacao, motivo = [], []
    for i in novo.index:
        s, k = sit_plat.at[i], (novo.at[i, "_mk"], novo.at[i, "_ck"])
        aviso = " — nota sem tempo/data de entrega, conferir" if alerta.at[i] == "conferir" else ""
        if k[0] and k in aprov1.index:
            situacao.append(APROVADO)
            motivo.append(f"Aprovado no 1º ciclo ({aprov1[k]:g}%) — prevalece" if s != APROVADO
                          else f"Aprovado no 1º ciclo ({aprov1[k]:g}%) e no 2º ciclo")
        elif k[0] and k[0] in alunos_ap:
            situacao.append(APROVADO); motivo.append("Aluno já aprovado em todos os componentes no 1º ciclo — prevalece")
        elif s == APROVADO:
            situacao.append(APROVADO); motivo.append("Aprovado no 2º ciclo (60% ou mais)" + aviso)
        elif s == "Não realizou" and k[0] and k in reprov1:
            situacao.append(NAO_APROVADO)
            motivo.append("Não aprovado no 1º ciclo; prova do 2º ciclo não realizada — continua não aprovado")
        elif s == "Não realizou":
            situacao.append(NAO_AVALIADO)
            motivo.append("Prova do 2º ciclo não realizada (0%, sem tempo e sem data de entrega)" if not aviso
                          else "Prova do 2º ciclo não realizada (0%; tempo ou data de entrega incompletos, conferir)")
        else:
            situacao.append(NAO_APROVADO); motivo.append("Fez a prova do 2º ciclo e não atingiu 60%" + aviso)

    # monta a saída: colunas originais + 4 novas (a coluna "Situação" original, se houver, é substituída)
    result = df_novo.loc[novo.index].copy()
    if csit:
        result = result.drop(columns=[csit])
    result["Situação 2º ciclo (plataforma)"] = sit_plat.values
    result["Situação"] = situacao
    result["Motivo da situação"] = motivo
    result["Corrigido?"] = [("Sim" if (s == APROVADO and p != APROVADO) else "")
                            for s, p in zip(situacao, sit_plat.values)]
    result["Conferir?"] = ["Sim" if a == "conferir" else "" for a in alerta.values]

    cols_corr = [c for c in [_col(result, "escola"), _col(result, "aluno"), _col(result, "email_aluno"),
                             _col(result, "prova"), _col(result, "name"), _col(result, "rendimento (%)"),
                             _col(result, "tempo de prova"), _col(result, "data_entrega"),
                             _col(result, "participação"), "Situação 2º ciclo (plataforma)", "Situação",
                             "Motivo da situação"] if c]
    corrigidos = result[result["Corrigido?"] == "Sim"][cols_corr]
    ord_cols = [c for c in (_col(result, "escola"), _col(result, "aluno")) if c]
    if ord_cols:
        corrigidos = corrigidos.sort_values(ord_cols)

    mk = novo["_mk"]
    resumo = pd.DataFrame({
        "Indicador": ["Linhas da planilha", "Aprovado (total)", "  aprovadas no 2º ciclo (60% ou mais)",
                      "  corrigidas para Aprovado (aprovação do 1º ciclo prevalece)",
                      "Não aprovado (fez a prova e não atingiu 60%, ou reprovado no 1º ciclo sem prova no 2º)",
                      "Não avaliado (0%, sem tempo e sem data de entrega)",
                      "Estudantes com algum componente corrigido", "Linhas para conferir (tempo/data incompletos)",
                      "Linhas sem matrícula identificada"],
        "Quantidade": [len(result), situacao.count(APROVADO), int((sit_plat == APROVADO).sum()),
                       int((result["Corrigido?"] == "Sim").sum()), situacao.count(NAO_APROVADO),
                       situacao.count(NAO_AVALIADO), int(mk[(result["Corrigido?"] == "Sim").values].nunique()),
                       int((alerta == "conferir").sum()), int((mk == "").sum())],
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