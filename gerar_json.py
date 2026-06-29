#!/usr/bin/env python3
"""
gerar_json.py — Converte dash_teste_tabelas.xlsx em dados.json
Estrutura conforme Apêndice da spec_anonimizada.docx.

Uso: python gerar_json.py
Saída: dados.json no mesmo diretório.
"""
import json
import math
from datetime import datetime
from pathlib import Path

import pandas as pd

XLSX_PATH = Path("dash_teste_tabelas (1).xlsx")
JSON_PATH = Path("dados.json")

MONTH_NAMES = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}

# Mapeamento etapa -> chave JSON no funil
ETAPAS_LM = ["FormFill", "ContatoT", "ContatoQ", "MQL", "MQLcPerfil", "SQL", "Oportunidade"]
ETAPAS_PASSIVOS = ["FormFill", "ContatoMKT", "MQLPassivo", "SQL", "Oportunidade"]

CARTEIRA_MAP = {
    "Medias": "medias",
    "Enterprise": "enterprise",
    "GrandesContas": "grandes_contas",
}


# ── utilidades ────────────────────────────────────────────────────────────────

def fmt_mes(year, month):
    return f"{MONTH_NAMES[month]}/{str(year)[-2:]}"


def prev_month(year, month):
    return (year, month - 1) if month > 1 else (year - 1, 12)


def safe_int(v):
    return 0 if (v is None or (isinstance(v, float) and math.isnan(v))) else int(v)


def safe_float(v, decimals=4):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    return round(float(v), decimals)


def pct_change(new_val, old_val):
    if old_val and old_val != 0:
        return round((new_val - old_val) / old_val, 4)
    return None


def filter_month(df, col, year, month):
    dt = pd.to_datetime(df[col])
    return df[(dt.dt.year == year) & (dt.dt.month == month)]


def vol(df, etapa_col, etapa_val, vol_col="volume"):
    return safe_int(df[df[etapa_col] == etapa_val][vol_col].sum())


# ── carregamento ──────────────────────────────────────────────────────────────

def load_sheets():
    xl = pd.ExcelFile(XLSX_PATH)
    return {
        name: xl.parse(name)
        for name in xl.sheet_names
        if not name.startswith("📋")
    }


# ── período de referência ─────────────────────────────────────────────────────

def get_ref_period(funil):
    funil["data"] = pd.to_datetime(funil["data"])
    max_date = funil["data"].max()
    return max_date.year, max_date.month


# ── hero ──────────────────────────────────────────────────────────────────────

def build_hero(funil, paginas_lm, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)

    # Pageviews (fato_paginas_lm)
    pv_cur = safe_int(filter_month(paginas_lm, "data", ref_y, ref_m)["pageviews"].sum())
    pv_prv = safe_int(filter_month(paginas_lm, "data", py, pm)["pageviews"].sum())

    # Filtros do funil por motion e mês
    lm_cur = filter_month(funil[funil["motion"] == "LM"], "data", ref_y, ref_m)
    lm_prv = filter_month(funil[funil["motion"] == "LM"], "data", py, pm)
    ps_cur = filter_month(funil[funil["motion"] == "Passivos"], "data", ref_y, ref_m)
    ps_prv = filter_month(funil[funil["motion"] == "Passivos"], "data", py, pm)
    all_cur = filter_month(funil, "data", ref_y, ref_m)
    all_prv = filter_month(funil, "data", py, pm)

    ct_lm_cur = vol(lm_cur, "etapa", "ContatoT")
    ct_lm_prv = vol(lm_prv, "etapa", "ContatoT")
    ct_ps_cur = vol(ps_cur, "etapa", "ContatoMKT")
    ct_ps_prv = vol(ps_prv, "etapa", "ContatoMKT")

    mql_lm = vol(lm_cur, "etapa", "MQL")
    mql_ps = vol(ps_cur, "etapa", "MQLPassivo")
    opp_lm = vol(lm_cur, "etapa", "Oportunidade")
    opp_ps = vol(ps_cur, "etapa", "Oportunidade")

    pipeline = safe_int(all_cur["pipeline_gerado"].sum())

    total_ct_cur = ct_lm_cur + ct_ps_cur
    total_ct_prv = ct_lm_prv + ct_ps_prv

    return {
        "pageviews": {
            "lm": pv_cur,
            "delta_mom": pct_change(pv_cur, pv_prv),
        },
        "contatos_totais": {
            "lm": ct_lm_cur,
            "passivos": ct_ps_cur,
            "delta_mom": pct_change(total_ct_cur, total_ct_prv),
        },
        "mqls": {"lm": mql_lm, "passivos": mql_ps},
        "oportunidades": {"lm": opp_lm, "passivos": opp_ps},
        "pipeline_gerado": pipeline,
    }


# ── alertas ───────────────────────────────────────────────────────────────────

def _taxa(df, etapa_from, etapa_to):
    vol_from = safe_int(df[df["etapa"] == etapa_from]["volume"].sum())
    vol_to = safe_int(df[df["etapa"] == etapa_to]["volume"].sum())
    return vol_to / vol_from if vol_from > 0 else 0, vol_from, vol_to


def _volume_kpi(df, etapa):
    return safe_int(df[df["etapa"] == etapa]["volume"].sum())


def build_alertas(funil, alertas_dim, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)
    alertas = []

    for _, row in alertas_dim.iterrows():
        motion = row["motion"]
        etapa_key = str(row["etapa"])
        limite_inf = float(row["limite_inferior"])
        limite_sup = float(row["limite_superior"])
        template = str(row["mensagem_template"])

        # Filtra pelo motion (Total = todos)
        if motion == "Total":
            df_cur = filter_month(funil, "data", ref_y, ref_m)
            df_prv = filter_month(funil, "data", py, pm)
        else:
            df_cur = filter_month(funil[funil["motion"] == motion], "data", ref_y, ref_m)
            df_prv = filter_month(funil[funil["motion"] == motion], "data", py, pm)

        # Taxa de conversão entre dois estágios (formato "EtapaA_EtapaB")
        if "_" in etapa_key and etapa_key not in ("Volume_MQL", "Pipeline_Gerado", "MQL_OPP", "MQL_SQL", "SQL_OPP"):
            partes = etapa_key.split("_", 1)
            etapa_from, etapa_to = partes[0], partes[1]
            valor_cur, _, _ = _taxa(df_cur, etapa_from, etapa_to)
            valor_prv, _, _ = _taxa(df_prv, etapa_from, etapa_to)
        elif etapa_key in ("MQL_SQL",):
            etapa_mql = "MQL" if motion == "LM" else "MQLPassivo"
            valor_cur, _, _ = _taxa(df_cur, etapa_mql, "SQL")
            valor_prv, _, _ = _taxa(df_prv, etapa_mql, "SQL")
        elif etapa_key in ("SQL_OPP",):
            valor_cur, _, _ = _taxa(df_cur, "SQL", "Oportunidade")
            valor_prv, _, _ = _taxa(df_prv, "SQL", "Oportunidade")
        elif etapa_key == "MQL_OPP":
            etapa_mql = "MQL" if motion == "LM" else "MQLPassivo"
            vol_mql = safe_int(df_cur[df_cur["etapa"] == etapa_mql]["volume"].sum())
            vol_opp = safe_int(df_cur[df_cur["etapa"] == "Oportunidade"]["volume"].sum())
            valor_cur = vol_opp / vol_mql if vol_mql > 0 else 0
            vol_mql_p = safe_int(df_prv[df_prv["etapa"] == etapa_mql]["volume"].sum())
            vol_opp_p = safe_int(df_prv[df_prv["etapa"] == "Oportunidade"]["volume"].sum())
            valor_prv = vol_opp_p / vol_mql_p if vol_mql_p > 0 else 0
        elif etapa_key == "Volume_MQL":
            etapa_mql = "MQL" if motion == "LM" else "MQLPassivo"
            valor_cur = _volume_kpi(df_cur, etapa_mql)
            valor_prv = _volume_kpi(df_prv, etapa_mql)
        elif etapa_key == "Pipeline_Gerado":
            valor_cur = safe_int(df_cur["pipeline_gerado"].sum())
            valor_prv = safe_int(df_prv["pipeline_gerado"].sum())
        else:
            continue

        variacao = round((valor_cur - valor_prv), 4) if valor_prv else None

        if valor_cur < limite_inf:
            tipo = "critical"
        elif valor_cur < limite_inf * 1.08:
            tipo = "warn"
        elif valor_cur > limite_sup * 1.10:
            tipo = "info"
        else:
            continue  # dentro da faixa normal

        # Formata valores para o template
        if etapa_key in ("Volume_MQL", "Pipeline_Gerado"):
            v_fmt = f"{int(valor_cur):,}".replace(",", ".")
            var_fmt = f"{variacao:+,.0f}" if variacao is not None else "n/d"
        else:
            v_fmt = f"{valor_cur*100:.1f}"
            var_fmt = f"{variacao*100:+.1f}" if variacao is not None else "n/d"

        try:
            texto = template.format(valor=v_fmt, variacao=var_fmt)
        except (KeyError, IndexError):
            texto = template

        alertas.append({"tipo": tipo, "texto": texto, "motion": motion, "kpi": etapa_key})

    return alertas


# ── funil por estágios ────────────────────────────────────────────────────────

def build_funil_stages(df, etapas, ref_y, ref_m):
    cur = filter_month(df, "data", ref_y, ref_m)
    return {e: vol(cur, "etapa", e) for e in etapas}


# ── carteiras ─────────────────────────────────────────────────────────────────

def build_carteiras(funil, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)
    result = {"lm": {}, "passivos": {}}

    lm_cur = filter_month(funil[funil["motion"] == "LM"], "data", ref_y, ref_m)
    lm_prv = filter_month(funil[funil["motion"] == "LM"], "data", py, pm)

    for raw_name, key in CARTEIRA_MAP.items():
        c_cur = lm_cur[lm_cur["carteira"] == raw_name]
        c_prv = lm_prv[lm_prv["carteira"] == raw_name]

        mqls = vol(c_cur, "etapa", "MQL")
        sqls = vol(c_cur, "etapa", "SQL")
        opps = vol(c_cur, "etapa", "Oportunidade")
        mqls_p = vol(c_prv, "etapa", "MQL")

        result["lm"][key] = {
            "mqls": mqls,
            "sqls": sqls,
            "opps": opps,
            "conv_mql_sql": safe_float(sqls / mqls) if mqls > 0 else 0,
            "conv_mql_opp": safe_float(opps / mqls) if mqls > 0 else 0,
            "delta_mqls_mom": pct_change(mqls, mqls_p),
        }

    # Passivos: sem split por carteira nos dados
    ps_cur = filter_month(funil[funil["motion"] == "Passivos"], "data", ref_y, ref_m)
    ps_prv = filter_month(funil[funil["motion"] == "Passivos"], "data", py, pm)
    mql_ps = vol(ps_cur, "etapa", "MQLPassivo")
    mql_ps_p = vol(ps_prv, "etapa", "MQLPassivo")

    result["passivos"]["total"] = {
        "mqls": mql_ps,
        "sqls": vol(ps_cur, "etapa", "SQL"),
        "opps": vol(ps_cur, "etapa", "Oportunidade"),
        "delta_mqls_mom": pct_change(mql_ps, mql_ps_p),
    }

    return result


# ── tendência 12 meses ────────────────────────────────────────────────────────

def build_tendencia_12m(funil, ref_y, ref_m):
    df = funil.copy()
    df["data"] = pd.to_datetime(df["data"])
    df["ym"] = df["data"].dt.to_period("M")

    ref_period = pd.Period(f"{ref_y}-{ref_m:02d}", "M")
    periods = [ref_period - i for i in range(11, -1, -1)]
    labels = [fmt_mes(p.year, p.month) for p in periods]

    lm_mql = df[(df["motion"] == "LM") & (df["etapa"] == "MQL")]
    ps_mql = df[(df["motion"] == "Passivos") & (df["etapa"] == "MQLPassivo")]

    series = {}

    # Por carteira (LM)
    for raw_name, key in CARTEIRA_MAP.items():
        monthly = lm_mql[lm_mql["carteira"] == raw_name].groupby("ym")["volume"].sum()
        series[key] = [safe_int(monthly.get(p, 0)) for p in periods]

    # Por fonte (LM)
    for fonte in sorted(lm_mql["fonte_mql"].dropna().unique()):
        monthly = lm_mql[lm_mql["fonte_mql"] == fonte].groupby("ym")["volume"].sum()
        series[fonte.lower()] = [safe_int(monthly.get(p, 0)) for p in periods]

    # Total passivos
    monthly_ps = ps_mql.groupby("ym")["volume"].sum()
    series["passivos"] = [safe_int(monthly_ps.get(p, 0)) for p in periods]

    return {"labels": labels, "series": series}


# ── MOFU / BOFU (páginas LM) ──────────────────────────────────────────────────

def build_mofu_bofu(paginas_lm, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)
    cur = filter_month(paginas_lm, "data", ref_y, ref_m)
    prv = filter_month(paginas_lm, "data", py, pm)

    pv_cur = safe_int(cur["pageviews"].sum())
    ff_cur = safe_int(cur["form_fills"].sum())
    vu_cur = safe_int(cur["visitantes_unicos"].sum())
    pv_prv = safe_int(prv["pageviews"].sum())
    ff_prv = safe_int(prv["form_fills"].sum())

    resumo = {
        "pageviews": pv_cur,
        "visitantes_unicos": vu_cur,
        "form_fills": ff_cur,
        "taxa_conversao": safe_float(ff_cur / vu_cur) if vu_cur > 0 else 0,
        "delta_pageviews_mom": pct_change(pv_cur, pv_prv),
        "delta_form_fills_mom": pct_change(ff_cur, ff_prv),
    }

    top = (
        cur.groupby(["url_pagina", "nome_pagina", "produto"])
        .agg(
            pageviews=("pageviews", "sum"),
            visitantes_unicos=("visitantes_unicos", "sum"),
            form_fills=("form_fills", "sum"),
        )
        .reset_index()
        .sort_values("form_fills", ascending=False)
        .head(10)
    )
    top["taxa_conversao"] = (
        top["form_fills"] / top["visitantes_unicos"].replace(0, float("nan"))
    ).round(4)
    top = top.fillna(0)
    for col in ["pageviews", "visitantes_unicos", "form_fills"]:
        top[col] = top[col].astype(int)

    return {"resumo": resumo, "top_paginas": top.to_dict("records")}


# ── TOFU (materiais passivos + SEO + AEO) ────────────────────────────────────

def build_tofu(paginas_passivos, seo_kws, aeo_prompts, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)

    # Materiais passivos
    cur = filter_month(paginas_passivos, "data", ref_y, ref_m)
    prv = filter_month(paginas_passivos, "data", py, pm)

    materiais = {
        "pageviews": safe_int(cur["pageviews"].sum()),
        "visitantes_unicos": safe_int(cur["visitantes_unicos"].sum()),
        "form_fills": safe_int(cur["form_fills"].sum()),
        "delta_pageviews_mom": pct_change(
            safe_int(cur["pageviews"].sum()), safe_int(prv["pageviews"].sum())
        ),
        "delta_form_fills_mom": pct_change(
            safe_int(cur["form_fills"].sum()), safe_int(prv["form_fills"].sum())
        ),
    }

    top_mat = (
        cur.groupby(["url_material", "nome_material", "tipo_material"])
        .agg(
            pageviews=("pageviews", "sum"),
            visitantes_unicos=("visitantes_unicos", "sum"),
            form_fills=("form_fills", "sum"),
        )
        .reset_index()
        .sort_values("form_fills", ascending=False)
        .head(10)
    )
    for col in ["pageviews", "visitantes_unicos", "form_fills"]:
        top_mat[col] = top_mat[col].astype(int)

    # SEO
    seo_kws["semana"] = pd.to_datetime(seo_kws["semana"])
    ultima_sem = seo_kws["semana"].max()
    sem_ant = ultima_sem - pd.Timedelta(days=7)
    kws_cur = seo_kws[seo_kws["semana"] == ultima_sem].copy()
    kws_prv = seo_kws[seo_kws["semana"] == sem_ant].copy()

    def dist_pos(df):
        pos = df["posicao"].dropna()
        return {
            "top3": int((pos <= 3).sum()),
            "p4_10": int(((pos > 3) & (pos <= 10)).sum()),
            "p11_20": int(((pos > 10) & (pos <= 20)).sum()),
            "p21_50": int(((pos > 20) & (pos <= 50)).sum()),
            "fora": int((pos > 50).sum()),
        }

    is_branded = kws_cur["is_branded"].astype(str).str.lower() == "true"

    top_kws = (
        kws_cur.nsmallest(20, "posicao")[
            ["keyword", "posicao", "posicao_anterior", "volume_busca", "url_ranqueada", "is_branded"]
        ]
        .copy()
    )
    top_kws["posicao"] = top_kws["posicao"].apply(lambda x: safe_int(x) if pd.notna(x) else None)
    top_kws["posicao_anterior"] = top_kws["posicao_anterior"].apply(lambda x: safe_int(x) if pd.notna(x) else None)
    top_kws["volume_busca"] = top_kws["volume_busca"].apply(lambda x: safe_int(x) if pd.notna(x) else None)
    top_kws["is_branded"] = top_kws["is_branded"].astype(str).str.lower() == "true"

    seo = {
        "semana_referencia": str(ultima_sem.date()),
        "total_kws": len(kws_cur),
        "distribuicao": dist_pos(kws_cur),
        "distribuicao_semana_anterior": dist_pos(kws_prv),
        "branded_vs_nao": {
            "branded": int(is_branded.sum()),
            "nao_branded": int((~is_branded).sum()),
        },
        "top_kws": top_kws.to_dict("records"),
    }

    # AEO
    aeo_prompts["semana"] = pd.to_datetime(aeo_prompts["semana"])
    ultima_sem_aeo = aeo_prompts["semana"].max()
    sem_ant_aeo = ultima_sem_aeo - pd.Timedelta(days=7)
    aeo_cur = aeo_prompts[aeo_prompts["semana"] == ultima_sem_aeo].copy()
    aeo_prv = aeo_prompts[aeo_prompts["semana"] == sem_ant_aeo].copy()

    mencoes_cur = int((aeo_cur["teste_mencionada"].astype(str).str.lower() == "true").sum())
    mencoes_prv = int((aeo_prv["teste_mencionada"].astype(str).str.lower() == "true").sum())
    total_cur_aeo = len(aeo_cur)

    def sov_por_llm(df):
        result = {}
        for llm, grp in df.groupby("llm"):
            mencionado = (grp["teste_mencionada"].astype(str).str.lower() == "true").sum()
            result[llm] = round(float(mencionado / len(grp)), 4)
        return result

    aeo = {
        "semana_referencia": str(ultima_sem_aeo.date()),
        "total_prompts_monitorados": total_cur_aeo,
        "mencoes": mencoes_cur,
        "share_of_voice": round(mencoes_cur / total_cur_aeo, 4) if total_cur_aeo > 0 else 0,
        "delta_sov_sem": pct_change(mencoes_cur, mencoes_prv) if mencoes_prv > 0 else None,
        "sov_por_llm": sov_por_llm(aeo_cur),
        "trafego_estimado": round(float(aeo_cur["trafego_estimado"].sum()), 1),
    }

    return {
        "materiais": materiais,
        "top_materiais": top_mat.to_dict("records"),
        "seo": seo,
        "aeo": aeo,
    }


# ── mídia paga ────────────────────────────────────────────────────────────────

def build_midia(midia_diaria, criativos, ref_y, ref_m):
    py, pm = prev_month(ref_y, ref_m)
    midia_diaria = midia_diaria.copy()
    criativos = criativos.copy()
    midia_diaria["data"] = pd.to_datetime(midia_diaria["data"])
    criativos["data"] = pd.to_datetime(criativos["data"])

    cur = filter_month(midia_diaria, "data", ref_y, ref_m)
    prv = filter_month(midia_diaria, "data", py, pm)

    inv_cur = float(cur["investimento"].sum())
    cliq_cur = safe_int(cur["cliques"].sum())
    leads_cur = safe_int(cur["leads"].sum())
    bfs_cur = round(float(cur["bfs"].sum()), 1)

    resumo = {
        "investimento": round(inv_cur, 2),
        "impressoes": safe_int(cur["impressoes"].sum()),
        "cliques": cliq_cur,
        "leads": leads_cur,
        "bfs": bfs_cur,
        "cpc_medio": round(inv_cur / cliq_cur, 2) if cliq_cur > 0 else None,
        "cpl_medio": round(inv_cur / leads_cur, 2) if leads_cur > 0 else None,
        "delta_investimento_mom": pct_change(inv_cur, float(prv["investimento"].sum())),
        "delta_leads_mom": pct_change(leads_cur, safe_int(prv["leads"].sum())),
    }

    por_plataforma = {}
    for plat, grp in cur.groupby("plataforma"):
        inv = float(grp["investimento"].sum())
        cl = safe_int(grp["cliques"].sum())
        ld = safe_int(grp["leads"].sum())
        por_plataforma[plat] = {
            "investimento": round(inv, 2),
            "impressoes": safe_int(grp["impressoes"].sum()),
            "cliques": cl,
            "leads": ld,
            "bfs": round(float(grp["bfs"].sum()), 1),
            "ctr_medio": safe_float(grp["ctr"].mean()),
            "cpc_medio": round(inv / cl, 2) if cl > 0 else None,
            "cpl_medio": round(inv / ld, 2) if ld > 0 else None,
        }

    top_camp = (
        cur.groupby(["id_campanha", "nome_campanha", "plataforma", "objetivo"])
        .agg(
            investimento=("investimento", "sum"),
            impressoes=("impressoes", "sum"),
            cliques=("cliques", "sum"),
            leads=("leads", "sum"),
            bfs=("bfs", "sum"),
        )
        .reset_index()
        .sort_values("leads", ascending=False)
        .head(10)
    )
    top_camp["investimento"] = top_camp["investimento"].round(2)
    top_camp["bfs"] = top_camp["bfs"].round(1)
    top_camp["impressoes"] = top_camp["impressoes"].astype(int)
    top_camp["cliques"] = top_camp["cliques"].astype(int)
    top_camp["leads"] = top_camp["leads"].astype(int)
    top_camp["cpl"] = (
        top_camp["investimento"] / top_camp["leads"].replace(0, float("nan"))
    ).round(2)

    cri_cur = filter_month(criativos, "data", ref_y, ref_m)
    cri_ativos = cri_cur[cri_cur["ativo"].astype(str).str.lower() == "true"]

    top_cri = (
        cri_ativos.groupby(["id_criativo", "nome_criativo", "plataforma", "formato", "audiencia"])
        .agg(
            impressoes=("impressoes", "sum"),
            cliques=("cliques", "sum"),
            leads=("leads", "sum"),
            investimento=("investimento", "sum"),
            bfs=("bfs", "sum"),
        )
        .reset_index()
        .sort_values("bfs", ascending=False)
        .head(10)
    )
    top_cri["investimento"] = top_cri["investimento"].round(2)
    top_cri["bfs"] = top_cri["bfs"].round(1)
    top_cri["impressoes"] = top_cri["impressoes"].astype(int)
    top_cri["cliques"] = top_cri["cliques"].astype(int)
    top_cri["leads"] = top_cri["leads"].astype(int)

    return {
        "resumo": resumo,
        "por_plataforma": por_plataforma,
        "top_campanhas": top_camp.to_dict("records"),
        "top_criativos": top_cri.to_dict("records"),
    }


# ── metas ─────────────────────────────────────────────────────────────────────

def build_metas(metas, ref_y, ref_m):
    metas = metas.copy()
    metas["mes"] = pd.to_datetime(metas["mes"])
    ref = metas[(metas["mes"].dt.year == ref_y) & (metas["mes"].dt.month == ref_m)]

    result = []
    for _, r in ref.iterrows():
        result.append(
            {
                "motion": r["motion"],
                "etapa": r["etapa"],
                "carteira": r["carteira"],
                "tipo_meta": r["tipo_meta"],
                "valor_meta": float(r["valor_meta"]),
                "comentario": str(r["comentario"]) if pd.notna(r["comentario"]) else None,
            }
        )
    return result


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"Lendo {XLSX_PATH}...")
    sheets = load_sheets()

    funil = sheets["fato_funil_diario"].copy()
    paginas_lm = sheets["fato_paginas_lm"].copy()
    paginas_passivos = sheets["fato_paginas_passivos"].copy()
    seo_kws = sheets["fato_seo_kws_semanal"].copy()
    aeo_prompts = sheets["fato_aeo_prompts_semanal"].copy()
    midia_diaria = sheets["fato_midia_paga_diaria"].copy()
    criativos = sheets["fato_midia_criativos"].copy()
    metas = sheets["dim_metas"].copy()
    alertas_dim = sheets["dim_alertas_funil"].copy()

    funil["data"] = pd.to_datetime(funil["data"])
    paginas_lm["data"] = pd.to_datetime(paginas_lm["data"])
    paginas_passivos["data"] = pd.to_datetime(paginas_passivos["data"])

    ref_y, ref_m = get_ref_period(funil)
    ref_label = fmt_mes(ref_y, ref_m)
    print(f"Período de referência: {ref_label}")

    updated_at = str(funil["updated_at"].max()) if "updated_at" in funil.columns else datetime.now().isoformat() + "Z"

    dados = {
        "metadata": {
            "atualizado_em": updated_at,
            "periodo_referencia": ref_label,
        },
        "funil_principal": {
            "hero": build_hero(funil, paginas_lm, ref_y, ref_m),
            "alertas": build_alertas(funil, alertas_dim, ref_y, ref_m),
            "funil_lm": build_funil_stages(
                funil[funil["motion"] == "LM"], ETAPAS_LM, ref_y, ref_m
            ),
            "funil_passivos": build_funil_stages(
                funil[funil["motion"] == "Passivos"], ETAPAS_PASSIVOS, ref_y, ref_m
            ),
            "carteiras": build_carteiras(funil, ref_y, ref_m),
            "tendencia_12m": build_tendencia_12m(funil, ref_y, ref_m),
        },
        "canais_aquisicao": {
            "mofu_bofu": build_mofu_bofu(paginas_lm, ref_y, ref_m),
            "tofu": build_tofu(paginas_passivos, seo_kws, aeo_prompts, ref_y, ref_m),
        },
        "midia_paga": build_midia(midia_diaria, criativos, ref_y, ref_m),
        "metas": build_metas(metas, ref_y, ref_m),
    }

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2, default=str)

    size_kb = JSON_PATH.stat().st_size / 1024
    print(f"Arquivo gerado: {JSON_PATH} ({size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
