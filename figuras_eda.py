"""Figuras e tabelas da analise exploratoria da BASE_SIPAER em tres classes (incidente, incidente grave, acidente).

Uso: python figuras_eda.py [BASE_SIPAER.xlsx]
Saidas: artigo/figuras/fig1_anual.png, fig2_regiao_local.png;
        artigo/tabelas/tab_etapas.tex, tab_posevento.tex, tab_operacao.tex, tab_tipos.tex; resultados/eda_numeros.json.
"""
import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

BASE = Path(sys.argv[1] if len(sys.argv) > 1 else "BASE_SIPAER.xlsx")
RAIZ = Path(__file__).resolve().parent
FIG, TAB, RES = RAIZ / "artigo" / "figuras", RAIZ / "artigo" / "tabelas", RAIZ / "resultados"
for d in (FIG, TAB, RES):
    d.mkdir(parents=True, exist_ok=True)

CLASSES = ["INCIDENTE", "INCIDENTE GRAVE", "ACIDENTE"]
b = pd.read_excel(BASE, sheet_name="base")
dic = pd.read_excel(BASE, sheet_name="dicionario")
fonte = pd.read_excel(BASE, sheet_name="fonte").set_index("item")["valor"]
b["data_ocorrencia"] = pd.to_datetime(b["data_ocorrencia"])
b["local"] = b.aerodromo.map(lambda a: "FORA DE AERÓDROMO" if a == "FAER"
                             else ("AERÓDROMOS SB" if str(a).startswith("SB") else "DEMAIS AERÓDROMOS"))
b["IG"] = (b.classificacao == "INCIDENTE GRAVE").astype(int)
b["AC"] = (b.classificacao == "ACIDENTE").astype(int)
n = len(b)
ini, fim = b.data_ocorrencia.min(), b.data_ocorrencia.max()


def mil(x):
    return f"{int(x):,}".replace(",", ".")


def pc(a, t, casas=1):
    return f"{100 * a / t:.{casas}f}".replace(".", ",")


def grupo(col, ordem):
    g = b.groupby(col).agg(n=("AC", "size"), ig=("IG", "sum"), ac=("AC", "sum")).reindex(ordem)
    g["tx_ig"], g["tx_ac"] = 100 * g.ig / g.n, 100 * g.ac / g.n
    return g


op = grupo("tipo_operacao", ["REGULAR", "NÃO REGULAR", "TÁXI AÉREO"])
mot = grupo("tipo_motor", ["JATO", "TURBOÉLICE", "TURBOEIXO", "PISTÃO"])
eqp = grupo("tipo_equipamento", ["AVIÃO", "HELICÓPTERO"])
qtd = grupo("qtd_motores", ["BIMOTOR", "MONOMOTOR"])
fase = grupo("fase_voo_grupo", ["SOLO", "DECOLAGEM E SUBIDA", "CRUZEIRO E DESCIDA", "APROXIMAÇÃO E POUSO"])
reg = grupo("regiao", ["NORTE", "CENTRO-OESTE", "NORDESTE", "SUL", "SUDESTE"])
loc = grupo("local", ["DEMAIS AERÓDROMOS", "FORA DE AERÓDROMO", "AERÓDROMOS SB"])
ORDEM_CAT = ["FALHA TÉCNICA", "COLISÃO COM FAUNA", "PISTA, POUSO E SOLO", "TRÁFEGO AÉREO", "METEOROLOGIA",
             "PERDA DE CONTROLE OU COLISÃO EM VOO", "OUTROS"]
cat = grupo("categoria_tipo", ORDEM_CAT)

# ------------------------------------------------------------------ tabelas
etapas = [v.rsplit(": ", 1) for k, v in fonte.items() if str(k).startswith("Etapa")]
linhas = [f"{nome} & {mil(q)} \\\\" for nome, q in etapas[:-1]] + [f"\\textbf{{{etapas[-1][0]}}} & \\textbf{{{mil(etapas[-1][1])}}} \\\\"]
(TAB / "tab_etapas.tex").write_text("\\begin{tabular}{lc}\n\\toprule\n\\textbf{Etapa} & \\textbf{Ocorrências} \\\\\n\\midrule\n"
                                    + "\n".join(linhas) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

pos = [("Dano substancial ou aeronave destruída", b.nivel_dano.isin(["SUBSTANCIAL", "DESTRUÍDA"])),
       ("Relatório final publicado", b.relatorio_publicado.eq("SIM")),
       ("Fatores contribuintes registrados", b.n_fatores > 0),
       ("Recomendações de segurança emitidas", b.n_recomendacoes > 0),
       ("Investigação ainda ativa", b.status_investigacao.eq("ATIVA"))]
nc = b.classificacao.value_counts()
cab = " & ".join(f"\\textbf{{{r} ({mil(nc[c])})}}" for r, c in
                 [("Acidentes", "ACIDENTE"), ("Inc. graves", "INCIDENTE GRAVE"), ("Incidentes", "INCIDENTE")])
linhas = []
for nome, m in pos:
    cel = [f"{int(m[b.classificacao == c].sum())} ({pc(m[b.classificacao == c].sum(), nc[c])}\\%)"
           for c in ["ACIDENTE", "INCIDENTE GRAVE", "INCIDENTE"]]
    linhas.append(f"{nome} & " + " & ".join(cel) + " \\\\")
(TAB / "tab_posevento.tex").write_text("\\begin{tabular}{lccc}\n\\toprule\n\\textbf{Atributo} & " + cab +
                                       " \\\\\n\\midrule\n" + "\n".join(linhas) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

ROT = {"REGULAR": "Regular", "NÃO REGULAR": "Não regular", "TÁXI AÉREO": "Táxi aéreo", "JATO": "Jato", "TURBOÉLICE": "Turboélice",
       "TURBOEIXO": "Turboeixo", "PISTÃO": "Pistão", "AVIÃO": "Avião", "HELICÓPTERO": "Helicóptero", "SOLO": "Solo",
       "DECOLAGEM E SUBIDA": "Decolagem e subida", "CRUZEIRO E DESCIDA": "Cruzeiro e descida", "APROXIMAÇÃO E POUSO": "Aproximação e pouso",
       "NORTE": "Norte", "CENTRO-OESTE": "Centro-\nOeste", "NORDESTE": "Nordeste", "SUL": "Sul", "SUDESTE": "Sudeste",
       "DEMAIS AERÓDROMOS": "Demais\naeródromos", "FORA DE AERÓDROMO": "Fora de\naeródromo", "AERÓDROMOS SB": "Aeródromos\nSB"}


def linha(nome, r):
    return f"\\hspace{{1em}}{nome} & {mil(r.n)} & {int(r.ig)} & {int(r.ac)} & {pc(r.ig, r.n, 2)} & {pc(r.ac, r.n, 2)} \\\\"


linhas = []
for titulo, g in [("Tipo de operação", op), ("Tipo de motor", mot), ("Equipamento", eqp), ("Fase do voo", fase)]:
    linhas.append(f"\\multicolumn{{6}}{{l}}{{\\textit{{{titulo}}}}} \\\\")
    linhas += [linha(ROT[k], r) for k, r in g.iterrows()]
(TAB / "tab_operacao.tex").write_text(
    "\\begin{tabular}{lccccc}\n\\toprule\n\\textbf{Atributo} & \\textbf{Ocorrências} & \\textbf{Inc. graves} & \\textbf{Acidentes} & "
    "\\textbf{\\% inc. graves} & \\textbf{\\% acidentes} \\\\\n\\midrule\n" + "\n".join(linhas) + "\n\\bottomrule\n\\end{tabular}\n",
    encoding="utf-8")

cap = lambda s: s[0] + s[1:].lower()
NOME_CAT = {"PERDA DE CONTROLE OU COLISÃO EM VOO": "Perda de controle / colisão em voo"}
linhas = [f"{NOME_CAT.get(c, cap(c))} & {mil(r.n)} & {pc(r.n, n)} & {int(r.ig)} & {pc(r.ig, r.n, 2)} & {int(r.ac)} & {pc(r.ac, r.n, 2)} \\\\"
          for c, r in cat.iterrows()]
linhas.append(f"\\midrule\n\\textbf{{Total}} & \\textbf{{{mil(n)}}} & \\textbf{{100,0}} & \\textbf{{{int(b.IG.sum())}}} & "
              f"\\textbf{{{pc(b.IG.sum(), n, 2)}}} & \\textbf{{{int(b.AC.sum())}}} & \\textbf{{{pc(b.AC.sum(), n, 2)}}} \\\\")
(TAB / "tab_tipos.tex").write_text(
    "\\begin{tabular}{lcccccc}\n\\toprule\n\\textbf{Categoria} & \\textbf{Ocorrências} & \\textbf{\\%} & \\textbf{Inc. graves} & "
    "\\textbf{\\% inc. gr.} & \\textbf{Acidentes} & \\textbf{\\% acid.} \\\\\n\\midrule\n" + "\n".join(linhas) +
    "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# ------------------------------------------------------------------ figuras
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
AZUL, AZUL_CLARO, LARANJA, VERMELHO = "#1f3b73", "#8faadc", "#e39a4c", "#b23a3a"

# Figura 1: ocorrencias por ano e tipo de operacao; incidentes graves e acidentes por ano
anual = pd.crosstab(b.ano, b.tipo_operacao).reindex(columns=["REGULAR", "NÃO REGULAR", "TÁXI AÉREO"], fill_value=0)
sev = b.groupby("ano")[["IG", "AC"]].sum().reindex(anual.index, fill_value=0)
cores = {"REGULAR": AZUL, "NÃO REGULAR": AZUL_CLARO, "TÁXI AÉREO": LARANJA}
fig, (ax, ax2) = plt.subplots(2, 1, figsize=(7.2, 4.3), sharex=True, gridspec_kw={"height_ratios": [2, 1.1]})
base = pd.Series(0, index=anual.index)
for c in anual.columns:
    ax.bar(anual.index, anual[c], bottom=base, color=cores[c], label=ROT[c], width=0.7)
    base += anual[c]
for a, tot in base.items():
    ax.text(a, tot + base.max() * 0.02, str(int(tot)), ha="center", fontsize=7.5)
ax.set_ylim(0, base.max() * 1.18)
ax.set_ylabel("Ocorrências")
ax.legend(fontsize=7.5, frameon=False, loc="upper left", title="Tipo de operação", title_fontsize=7.5)
w = 0.36
ax2.bar(sev.index - w / 2, sev.IG, width=w, color=AZUL_CLARO, label="Incidentes graves")
ax2.bar(sev.index + w / 2, sev.AC, width=w, color=VERMELHO, label="Acidentes")
for a, r in sev.iterrows():
    ax2.text(a - w / 2, r.IG + 0.4, str(int(r.IG)), ha="center", fontsize=6.5)
    ax2.text(a + w / 2, r.AC + 0.4, str(int(r.AC)), ha="center", fontsize=6.5)
ax2.set_ylim(0, sev.to_numpy().max() * 1.3)
ax2.set_ylabel("Ocorrências")
ax2.legend(fontsize=7.5, frameon=False, loc="upper right", ncol=2)
ax2.set_xticks(anual.index)
ax2.set_xticklabels([f"{a}*" if a in (ini.year, fim.year) else str(a) for a in anual.index])
for eixo in (ax, ax2):
    eixo.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(FIG / "fig1_anual.png", dpi=220)
plt.close(fig)

# Figura 2: incidentes graves e acidentes por 100 ocorrencias, por regiao e por local
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 2.9), gridspec_kw={"width_ratios": [1.3, 1]})
ymax = max(reg[["tx_ig", "tx_ac"]].to_numpy().max(), loc[["tx_ig", "tx_ac"]].to_numpy().max()) * 1.25
w = 0.38
for a, g, titulo in [(a1, reg, "(a) Região"), (a2, loc, "(b) Local da ocorrência")]:
    x = np.arange(len(g))
    a.bar(x - w / 2, g.tx_ig, width=w, color=AZUL_CLARO, label="Incidentes graves")
    a.bar(x + w / 2, g.tx_ac, width=w, color=VERMELHO, label="Acidentes")
    for i, (_, r) in enumerate(g.iterrows()):
        a.text(i - w / 2, r.tx_ig + ymax * 0.015, f"{int(r.ig)}", ha="center", fontsize=6.5)
        a.text(i + w / 2, r.tx_ac + ymax * 0.015, f"{int(r.ac)}", ha="center", fontsize=6.5)
    a.set_xticks(x)
    a.set_xticklabels([f"{ROT[k]}\n(n={mil(r.n)})" for k, r in g.iterrows()], fontsize=7)
    a.set_ylabel("Por 100 ocorrências")
    a.set_ylim(0, ymax)
    a.set_title(titulo, fontsize=9)
    a.spines[["top", "right"]].set_visible(False)
a1.legend(fontsize=7, frameon=False, loc="upper right")
fig.tight_layout()
fig.savefig(FIG / "fig2_regiao_local.png", dpi=220)
plt.close(fig)

# ------------------------------------------------------------------ numeros para o texto
acid, ig, inc = b[b.AC == 1], b[b.IG == 1], b[b.classificacao == "INCIDENTE"]
taxi = b[b.tipo_operacao == "TÁXI AÉREO"]
tl = b.tipos_ocorrencia.fillna("")
pre, pos_ = b[b.ano < 2023], b[b.ano >= 2023]
fat_acid = acid[acid.n_fatores > 0]
top_fat = fat_acid.fatores_lista.str.split(" | ", regex=False).explode().value_counts()
fat_ig = ig[ig.n_fatores > 0]
top_fat_ig = fat_ig.fatores_lista.str.split(" | ", regex=False).explode().value_counts()
nums = {
    "n": n, "ini": f"{ini:%d/%m/%Y}", "fim": f"{fim:%d/%m/%Y}", "colunas": int(b.shape[1] - 4), "multi": int((b.n_aeronaves_comerciais > 1).sum()),
    "n_ac": int(len(acid)), "n_ig": int(len(ig)), "n_inc": int(len(inc)), "pct_ac": pc(len(acid), n, 2), "pct_ig": pc(len(ig), n, 1), "pct_inc": pc(len(inc), n, 1),
    "fatais": int((acid.fatalidades > 0).sum()), "mortes": int(acid.fatalidades.sum()), "mortes_max": int(acid.fatalidades.max()),
    "fatalidades_ig": int(ig.fatalidades.sum()),
    "nulos_assentos": f"{dic.set_index('coluna')['nulos_%']['assentos']:.1f}".replace(".", ","),
    "nulos_ano_fab": f"{dic.set_index('coluna')['nulos_%']['ano_fabricacao']:.1f}".replace(".", ","),
    "fat_acid_n": int(len(fat_acid)), "fat_acid_oper": int(fat_acid.fator_operacional.sum()), "fat_acid_hum": int(fat_acid.fator_humano.sum()),
    "fat_acid_mat": int(fat_acid.fator_material.sum()),
    "top_fat": {k: int(top_fat.get(k, 0)) for k in ["PROCESSO DECISÓRIO", "JULGAMENTO DE PILOTAGEM", "SUPERVISÃO GERENCIAL"]},
    "fat_ig_n": int(len(fat_ig)), "top_fat_ig": {k: int(v) for k, v in top_fat_ig.head(3).items()},
    "op": op.round(2).to_dict("index"), "mot": mot.round(2).to_dict("index"), "eqp": eqp.round(2).to_dict("index"),
    "qtd": qtd.round(2).to_dict("index"), "fase": fase.round(2).to_dict("index"), "reg": reg.round(2).to_dict("index"),
    "loc": loc.round(2).to_dict("index"), "cat": cat.round(2).to_dict("index"),
    "fab_ab": {"n": int(b.fabricante_grupo.isin(["AIRBUS", "BOEING"]).sum()), "ig": int(b.IG[b.fabricante_grupo.isin(["AIRBUS", "BOEING"])].sum()),
               "ac": int(b.AC[b.fabricante_grupo.isin(["AIRBUS", "BOEING"])].sum())},
    "fab_cr": {"n": int(b.fabricante_grupo.isin(["CESSNA", "ROBINSON"]).sum()), "ig": int(b.IG[b.fabricante_grupo.isin(["CESSNA", "ROBINSON"])].sum()),
               "ac": int(b.AC[b.fabricante_grupo.isin(["CESSNA", "ROBINSON"])].sum())},
    "idade_mediana": {c: float(b.loc[b.classificacao == c, "idade_aeronave"].median()) for c in CLASSES},
    "pmd_mediano": {c: float(b.loc[b.classificacao == c, "pmd_kg"].median()) for c in CLASSES},
    "taxi_pistao": int((taxi.tipo_motor == "PISTÃO").sum()), "pistao": int((b.tipo_motor == "PISTÃO").sum()),
    "taxi_heli": int((taxi.tipo_equipamento == "HELICÓPTERO").sum()), "heli": int((b.tipo_equipamento == "HELICÓPTERO").sum()),
    "tipos_unicos": int(tl.str.split(" | ", regex=False).explode().replace("", pd.NA).dropna().nunique()), "multi_tipo": int((b.n_tipos > 1).sum()),
    "motor_voo": {"n": int(((b.categoria_tipo == "FALHA TÉCNICA") & tl.str.contains("FALHA DO MOTOR EM VOO", regex=False)).sum()),
                  "ac": int(b.AC[(b.categoria_tipo == "FALHA TÉCNICA") & tl.str.contains("FALHA DO MOTOR EM VOO", regex=False)].sum()),
                  "ig": int(b.IG[(b.categoria_tipo == "FALHA TÉCNICA") & tl.str.contains("FALHA DO MOTOR EM VOO", regex=False)].sum())},
    "excursao": {"n": int(tl.str.contains("EXCURSÃO DE PISTA", regex=False).sum()), "ac": int(b.AC[tl.str.contains("EXCURSÃO DE PISTA", regex=False)].sum()),
                 "ig": int(b.IG[tl.str.contains("EXCURSÃO DE PISTA", regex=False)].sum())},
    "pre": {"n": int(len(pre)), "ig": int(pre.IG.sum()), "ac": int(pre.AC.sum())}, "pos": {"n": int(len(pos_)), "ig": int(pos_.IG.sum()), "ac": int(pos_.AC.sum())},
    "pico_2024": int((b.ano == 2024).sum()), "fauna_23_24": int(((b.ano.isin([2023, 2024])) & (b.categoria_tipo == "COLISÃO COM FAUNA")).sum()),
    "taxi_2025": int(((b.ano == 2025) & (b.tipo_operacao == "TÁXI AÉREO")).sum()), "taxi_2026": int(((b.ano == 2026) & (b.tipo_operacao == "TÁXI AÉREO")).sum()),
    "norte_taxi_pct": pc(((b.regiao == "NORTE") & (b.tipo_operacao == "TÁXI AÉREO")).sum(), reg.loc["NORTE", "n"], 0), "taxi_pct": pc(len(taxi), n, 0),
    "n_aerodromos": int(b.aerodromo.nunique()), "n_cidades": int(b.cidade.nunique()),
}
(RES / "eda_numeros.json").write_text(json.dumps(nums, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
print(json.dumps({k: nums[k] for k in ["n_ac", "n_ig", "n_inc", "fatalidades_ig", "pre", "pos", "fat_ig_n", "top_fat_ig", "idade_mediana", "pmd_mediano", "motor_voo", "excursao", "fab_ab", "fab_cr"]}, ensure_ascii=False))
print("ok")
