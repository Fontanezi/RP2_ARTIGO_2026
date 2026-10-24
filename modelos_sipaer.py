"""Modelagem da BASE_SIPAER: classificacao em tres classes (INCIDENTE, INCIDENTE GRAVE, ACIDENTE).

Uso: python modelos_sipaer.py [BASE_SIPAER.xlsx]

Compara Regressao Logistica multinomial e Random Forest sob quatro estrategias de balanceamento
(nenhum, pesos de classe, SMOTENC, SMOTENC + pesos), com particao 80/20 estratificada e busca de
hiperparametros por validacao cruzada estratificada (k = 5) no treino, otimizando F1-macro.
Reporta as predicoes out-of-fold do treino (numero principal) e o teste (confirmacao).

Saidas: artigo/figuras/fig3_confusao.png, fig4_curvas.png, fig5_importancias.png;
        artigo/tabelas/tab_ablacao.tex, tab_metricas.tex; resultados/metricas.json, resultados/*.csv.
"""
import json
import sys
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from imblearn.over_sampling import SMOTENC
from imblearn.pipeline import Pipeline
from sklearn.base import clone
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.inspection import permutation_importance
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (average_precision_score, balanced_accuracy_score, confusion_matrix, f1_score,
                             precision_recall_curve, precision_recall_fscore_support, roc_auc_score, roc_curve)
from sklearn.model_selection import GridSearchCV, StratifiedKFold, cross_val_predict, train_test_split
from sklearn.preprocessing import OneHotEncoder, StandardScaler

# numpy 2.0 + BLAS Accelerate (macOS) emite avisos espurios em matmul; resultados conferidos contra einsum.
warnings.filterwarnings("ignore", message=".*encountered in matmul")

BASE = Path(sys.argv[1] if len(sys.argv) > 1 else "BASE_SIPAER.xlsx")
RAIZ = Path(__file__).resolve().parent
FIG = RAIZ / "artigo" / "figuras"
TAB = RAIZ / "artigo" / "tabelas"
RES = RAIZ / "resultados"
for d in (FIG, TAB, RES):
    d.mkdir(parents=True, exist_ok=True)

SEED = 42
CLASSES = ["INCIDENTE", "INCIDENTE GRAVE", "ACIDENTE"]
ROTULO = {"INCIDENTE": "Incidente", "INCIDENTE GRAVE": "Inc. grave", "ACIDENTE": "Acidente"}
COD = {c: i for i, c in enumerate(CLASSES)}
INC, IG, AC = 0, 1, 2

# ================================================================== atributos
NUM = ["log_pmd", "assentos", "idade_aeronave", "hora", "n_tipos", "total_aeronaves_envolvidas", "saida_pista"]
CAT = ["tipo_operacao", "tipo_equipamento", "tipo_motor", "qtd_motores", "fase_voo_grupo", "regiao", "local",
       "fabricante_grupo", "categoria_tipo", "mes", "dia_semana"]
NOME = {"log_pmd": "log10 PMD (kg)", "assentos": "Assentos", "idade_aeronave": "Idade da aeronave", "hora": "Hora",
        "n_tipos": "Nº de tipos de ocorrência", "total_aeronaves_envolvidas": "Aeronaves envolvidas",
        "saida_pista": "Saída de pista", "tipo_operacao": "Tipo de operação", "tipo_equipamento": "Equipamento",
        "tipo_motor": "Tipo de motor", "qtd_motores": "Qtd. de motores", "fase_voo_grupo": "Fase do voo",
        "regiao": "Região", "local": "Local", "fabricante_grupo": "Fabricante", "categoria_tipo": "Categoria do tipo",
        "mes": "Mês", "dia_semana": "Dia da semana"}


def preparar(b):
    """Deriva os preditores pre-evento a partir da base publicada (sem imputar: isso e feito no pipeline)."""
    X = pd.DataFrame(index=b.index)
    X["log_pmd"] = np.log10(b.pmd_kg)
    X["assentos"] = b.assentos
    X["idade_aeronave"] = b.idade_aeronave
    X["hora"] = b.hora
    X["n_tipos"] = b.n_tipos
    X["total_aeronaves_envolvidas"] = b.total_aeronaves_envolvidas
    X["saida_pista"] = b.saida_pista
    X["tipo_operacao"] = b.tipo_operacao
    X["tipo_equipamento"] = b.tipo_equipamento.replace({"ANFÍBIO": "AVIÃO"})
    X["tipo_motor"] = b.tipo_motor
    X["qtd_motores"] = b.qtd_motores.where(b.qtd_motores.isin(["MONOMOTOR", "BIMOTOR"]) | b.qtd_motores.isna(), "OUTROS")
    X["fase_voo_grupo"] = b.fase_voo_grupo
    X["regiao"] = b.regiao
    X["local"] = b.aerodromo.map(lambda a: "FORA DE AERÓDROMO" if a == "FAER"
                                 else ("AERÓDROMOS SB" if str(a).startswith("SB") else "DEMAIS AERÓDROMOS"))
    X["fabricante_grupo"] = b.fabricante_grupo
    X["categoria_tipo"] = b.categoria_tipo
    X["mes"] = b.mes.map(lambda m: f"M{int(m):02d}")
    X["dia_semana"] = b.dia_semana
    return X[NUM + CAT]


b = pd.read_excel(BASE, sheet_name="base")
X = preparar(b)
y = b.classificacao.map(COD).to_numpy()
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.2, stratify=y, random_state=SEED)
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
print(f"base {X.shape}, treino {len(X_tr)} {np.bincount(y_tr).tolist()}, teste {len(X_te)} {np.bincount(y_te).tolist()}")

# ================================================================== pipeline
IDX_NUM = list(range(len(NUM)))
IDX_CAT = list(range(len(NUM), len(NUM) + len(CAT)))


def pipeline(modelo, balanceamento):
    """imputar (treino) -> [SMOTENC] -> padronizar + one-hot -> modelo. Imputacao e codificacao sao ajustadas
    apenas nas dobras de treino; o SMOTENC so e aplicado em fit (nunca nos dados avaliados)."""
    imputar = ColumnTransformer([("num", SimpleImputer(strategy="median"), NUM),
                                 ("cat", SimpleImputer(strategy="most_frequent"), CAT)])
    codificar = ColumnTransformer([("num", StandardScaler(), IDX_NUM),
                                   ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), IDX_CAT)])
    pesos = "balanced" if balanceamento in ("pesos", "ambos") else None
    if modelo == "LR":
        clf = LogisticRegression(max_iter=5000, class_weight=pesos, random_state=SEED)
    else:
        clf = RandomForestClassifier(n_estimators=400, class_weight=pesos, random_state=SEED, n_jobs=1)
    smote = SMOTENC(categorical_features=IDX_CAT, random_state=SEED) if balanceamento in ("smote", "ambos") else "passthrough"
    return Pipeline([("imputar", imputar), ("smote", smote), ("codificar", codificar), ("clf", clf)])


GRADE = {"LR": {"clf__C": [0.01, 0.1, 1.0, 10.0]},
         "RF": {"clf__max_depth": [None, 10], "clf__min_samples_leaf": [1, 5, 20], "clf__max_features": ["sqrt", 0.3]}}
BALANCEAMENTOS = ["nenhum", "pesos", "smote", "ambos"]
ROT_BAL = {"nenhum": "Nenhum", "pesos": "Pesos de classe", "smote": "SMOTENC", "ambos": "SMOTENC + pesos"}


def metricas(y_true, y_pred, proba):
    p, r, f, _ = precision_recall_fscore_support(y_true, y_pred, labels=[INC, IG, AC], zero_division=0)
    m = {"f1_macro": f1_score(y_true, y_pred, average="macro"),
         "acuracia_balanceada": balanced_accuracy_score(y_true, y_pred),
         "acuracia": float((y_true == y_pred).mean())}
    for i, c in enumerate(["inc", "ig", "ac"]):
        m[f"prec_{c}"], m[f"rec_{c}"], m[f"f1_{c}"] = float(p[i]), float(r[i]), float(f[i])
    if proba is not None:
        m["auc_ovr_macro"] = roc_auc_score(y_true, proba, multi_class="ovr", average="macro")
        m["auc_ac"] = roc_auc_score((y_true == AC).astype(int), proba[:, AC])
        m["auc_ig"] = roc_auc_score((y_true == IG).astype(int), proba[:, IG])
        m["ap_ac"] = average_precision_score((y_true == AC).astype(int), proba[:, AC])
        m["ap_ig"] = average_precision_score((y_true == IG).astype(int), proba[:, IG])
    m["matriz"] = confusion_matrix(y_true, y_pred, labels=[INC, IG, AC]).tolist()
    return m


# ================================================================== baselines
def baseline_majoritaria(Xd):
    return np.full(len(Xd), INC)


def baseline_regra(Xd):
    return np.where(Xd.tipo_operacao.isin(["TÁXI AÉREO", "NÃO REGULAR"]), AC, INC)


resultados = {"n_treino": int(len(X_tr)), "n_teste": int(len(X_te)),
              "classes_treino": np.bincount(y_tr).tolist(), "classes_teste": np.bincount(y_te).tolist(),
              "baselines": {}, "modelos": {}}
for nome, fn in [("majoritaria", baseline_majoritaria), ("regra", baseline_regra)]:
    resultados["baselines"][nome] = {"treino": metricas(y_tr, fn(X_tr), None), "teste": metricas(y_te, fn(X_te), None)}

# ================================================================== busca + avaliacao
ajustados, oof_proba, te_proba = {}, {}, {}
for modelo in ["LR", "RF"]:
    for bal in BALANCEAMENTOS:
        chave = f"{modelo}|{bal}"
        gs = GridSearchCV(pipeline(modelo, bal), GRADE[modelo], scoring="f1_macro", cv=cv, n_jobs=-1, refit=True)
        gs.fit(X_tr, y_tr)
        melhor = gs.best_estimator_
        i = gs.best_index_
        proba_oof = cross_val_predict(clone(melhor), X_tr, y_tr, cv=cv, method="predict_proba", n_jobs=-1)
        proba_te = melhor.predict_proba(X_te)
        ajustados[chave], oof_proba[chave], te_proba[chave] = melhor, proba_oof, proba_te
        resultados["modelos"][chave] = {
            "modelo": modelo, "balanceamento": bal, "melhores_parametros": gs.best_params_,
            "cv_f1_macro_media": float(gs.cv_results_["mean_test_score"][i]),
            "cv_f1_macro_dp": float(gs.cv_results_["std_test_score"][i]),
            "oof": metricas(y_tr, proba_oof.argmax(1), proba_oof),
            "teste": metricas(y_te, proba_te.argmax(1), proba_te)}
        r = resultados["modelos"][chave]
        print(f"{chave:12s} F1m CV {r['cv_f1_macro_media']:.3f}±{r['cv_f1_macro_dp']:.3f} | OOF recAC {r['oof']['rec_ac']:.2f} "
              f"precAC {r['oof']['prec_ac']:.2f} recIG {r['oof']['rec_ig']:.2f} AUCac {r['oof']['auc_ac']:.3f} | "
              f"teste recAC {r['teste']['rec_ac']:.2f} precAC {r['teste']['prec_ac']:.2f} F1m {r['teste']['f1_macro']:.3f} "
              f"| {gs.best_params_}")

# configuracao principal de cada modelo: melhor F1-macro na validacao cruzada
principal = {}
for modelo in ["LR", "RF"]:
    cands = {k: v for k, v in resultados["modelos"].items() if v["modelo"] == modelo}
    principal[modelo] = max(cands, key=lambda k: cands[k]["cv_f1_macro_media"])
resultados["principal"] = principal
print("principais:", principal)


def precisao_na_sensibilidade(y_true, score, alvo):
    """Maior precisao alcancavel com sensibilidade >= alvo (curva precisao-sensibilidade, um-contra-resto)."""
    prec, rec, _ = precision_recall_curve(y_true, score)
    return float(prec[rec >= alvo].max())


# comparacao justa com a regra: precisao dos modelos no mesmo nivel de sensibilidade da regra (classe acidente)
rec_regra = resultados["baselines"]["regra"]["treino"]["rec_ac"]
resultados["precisao_na_sensibilidade_da_regra"] = {"sensibilidade_regra": rec_regra}
for modelo in ["LR", "RF"]:
    for conj, proba, yy in [("oof", oof_proba[principal[modelo]], y_tr), ("teste", te_proba[principal[modelo]], y_te)]:
        resultados["precisao_na_sensibilidade_da_regra"][f"{modelo}_{conj}"] = precisao_na_sensibilidade(
            (yy == AC).astype(int), proba[:, AC], rec_regra)
    resultados["precisao_na_sensibilidade_da_regra"][f"{modelo}_oof_rec80"] = precisao_na_sensibilidade(
        (y_tr == AC).astype(int), oof_proba[principal[modelo]][:, AC], 0.80)
print("precisao na sensibilidade da regra:", resultados["precisao_na_sensibilidade_da_regra"])

# ================================================================== tabelas
def pt(x, casas=2):
    return f"{x:.{casas}f}".replace(".", ",")


def pc(x):
    return f"{100 * x:.1f}".replace(".", ",")


linhas = []
for modelo in ["LR", "RF"]:
    for bal in BALANCEAMENTOS:
        r = resultados["modelos"][f"{modelo}|{bal}"]
        o = r["oof"]
        nome_modelo = "Regressão Logística" if modelo == "LR" else "Random Forest"
        marca = r"$\dagger$" if f"{modelo}|{bal}" == principal[modelo] else ""
        linhas.append(f"{nome_modelo if bal == 'nenhum' else ''} & {ROT_BAL[bal]}{marca} & {pt(r['cv_f1_macro_media'], 3)} "
                      f"$\\pm$ {pt(r['cv_f1_macro_dp'], 3)} & {pc(o['rec_ac'])} & {pc(o['prec_ac'])} & {pc(o['rec_ig'])} & "
                      f"{pc(o['prec_ig'])} & {pt(o['auc_ac'], 3)} \\\\")
    if modelo == "LR":
        linhas.append(r"\midrule")
(TAB / "tab_ablacao.tex").write_text(
    "\\begin{tabular}{llcccccc}\n\\toprule\n"
    "\\textbf{Modelo} & \\textbf{Balanceamento} & \\textbf{F1-macro CV} & \\textbf{Sens. AC} & \\textbf{Prec. AC} & "
    "\\textbf{Sens. IG} & \\textbf{Prec. IG} & \\textbf{AUC AC} \\\\\n\\midrule\n" + "\n".join(linhas) +
    "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")


def linha_metricas(nome, m, proba=True):
    auc = pt(m["auc_ac"], 3) if proba else "--"
    return (f"{nome} & {pc(m['rec_ac'])} & {pc(m['prec_ac'])} & {pc(m['rec_ig'])} & {pc(m['prec_ig'])} & "
            f"{pc(m['rec_inc'])} & {pt(m['f1_macro'], 3)} & {pt(m['acuracia_balanceada'], 3)} & {auc} \\\\")


blocos = []
for conj, chave_conj, n_ac, n_ig in [("Validação cruzada no treino", "treino", np.bincount(y_tr)[AC], np.bincount(y_tr)[IG]),
                                     ("Teste", "teste", np.bincount(y_te)[AC], np.bincount(y_te)[IG])]:
    blocos.append(f"\\multicolumn{{9}}{{l}}{{\\textit{{{conj} ({n_ac} acidentes, {n_ig} incidentes graves)}}}} \\\\")
    blocos.append(linha_metricas("Classe majoritária", resultados["baselines"]["majoritaria"][chave_conj], proba=False))
    blocos.append(linha_metricas("Regra pelo tipo de operação", resultados["baselines"]["regra"][chave_conj], proba=False))
    for modelo in ["LR", "RF"]:
        r = resultados["modelos"][principal[modelo]]
        nome_modelo = ("Regressão Logística" if modelo == "LR" else "Random Forest") + f" ({ROT_BAL[r['balanceamento']]})"
        blocos.append(linha_metricas(nome_modelo, r["oof" if chave_conj == "treino" else "teste"]))
    if chave_conj == "treino":
        blocos.append(r"\midrule")
(TAB / "tab_metricas.tex").write_text(
    "\\begin{tabular}{lcccccccc}\n\\toprule\n"
    "\\textbf{Modelo} & \\textbf{Sens. AC} & \\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & "
    "\\textbf{Sens. INC} & \\textbf{F1-macro} & \\textbf{Acur. bal.} & \\textbf{AUC AC} \\\\\n\\midrule\n" +
    "\n".join(blocos) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# ================================================================== figuras
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
AZUL, LARANJA = "#1f3b73", "#e39a4c"
rot_curto = [ROTULO[c] for c in CLASSES]

# Figura 3: matrizes de confusao (out-of-fold no treino) dos dois modelos principais
fig, eixos = plt.subplots(1, 2, figsize=(7.2, 3.0))
for ax, modelo in zip(eixos, ["LR", "RF"]):
    m = np.array(resultados["modelos"][principal[modelo]]["oof"]["matriz"])
    norm = m / m.sum(1, keepdims=True)
    ax.imshow(norm, cmap="Blues", vmin=0, vmax=1)
    for i in range(3):
        for j in range(3):
            ax.text(j, i, f"{m[i, j]}\n({100 * norm[i, j]:.0f}%)", ha="center", va="center", fontsize=8,
                    color="white" if norm[i, j] > 0.6 else "black")
    ax.set_xticks(range(3)); ax.set_yticks(range(3))
    ax.set_xticklabels(rot_curto); ax.set_yticklabels(rot_curto)
    ax.set_xlabel("Classe prevista"); ax.set_ylabel("Classe real")
    r = resultados["modelos"][principal[modelo]]
    ax.set_title(("(a) Regressão Logística" if modelo == "LR" else "(b) Random Forest") + f" ({ROT_BAL[r['balanceamento']]})", fontsize=9)
fig.tight_layout()
fig.savefig(FIG / "fig3_confusao.png", dpi=220)
plt.close(fig)

# Figura 4: curvas ROC e precisao-sensibilidade um-contra-resto (out-of-fold no treino)
def num_pt(x, casas=2):
    return f"{x:.{casas}f}".replace(".", ",")


fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 3.0))
cor = {"LR": AZUL, "RF": LARANJA}
for modelo in ["LR", "RF"]:
    proba = oof_proba[principal[modelo]]
    nome = "Regressão Logística" if modelo == "LR" else "Random Forest"
    for classe, ls, alpha in [(AC, "-", 1.0), (IG, "--", 0.8)]:
        yb = (y_tr == classe).astype(int)
        fpr, tpr, _ = roc_curve(yb, proba[:, classe])
        prec, rec, _ = precision_recall_curve(yb, proba[:, classe])
        auc = roc_auc_score(yb, proba[:, classe]); ap = average_precision_score(yb, proba[:, classe])
        a1.plot(fpr, tpr, ls, color=cor[modelo], alpha=alpha, lw=1.4, label=f"{nome}, {ROTULO[CLASSES[classe]].lower()} (AUC {num_pt(auc)})")
        a2.plot(rec, prec, ls, color=cor[modelo], alpha=alpha, lw=1.4, label=f"{nome}, {ROTULO[CLASSES[classe]].lower()} (AP {num_pt(ap)})")
# regra pelo tipo de operacao, classe acidente, como ponto de referencia
mr = np.array(resultados["baselines"]["regra"]["treino"]["matriz"])
fpr_regra = (mr[INC, AC] + mr[IG, AC]) / (mr[INC].sum() + mr[IG].sum())
a1.plot(fpr_regra, rec_regra, "D", color="black", ms=4, label="Regra pelo tipo de operação")
a2.plot(rec_regra, resultados["baselines"]["regra"]["treino"]["prec_ac"], "D", color="black", ms=4, label="Regra pelo tipo de operação")
a1.plot([0, 1], [0, 1], ":", color="gray", lw=1)
a1.set_xlabel("Taxa de falsos positivos"); a1.set_ylabel("Sensibilidade"); a1.set_title("(a) ROC", fontsize=9)
a2.set_xlabel("Sensibilidade"); a2.set_ylabel("Precisão"); a2.set_title("(b) Precisão-sensibilidade", fontsize=9)
a2.axhline((y_tr == AC).mean(), ls=":", color="gray", lw=1)
a1.legend(fontsize=6.3, frameon=False, loc="lower right"); a2.legend(fontsize=6.3, frameon=False, loc="upper right")
for a in (a1, a2):
    a.set_xlim(0, 1); a.set_ylim(0, 1.02); a.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(FIG / "fig4_curvas.png", dpi=220)
plt.close(fig)

# Figura 5: contrastes da LR (acidente vs incidente) e importancia por permutacao do RF (teste)
lr = ajustados[principal["LR"]]
enc = lr.named_steps["codificar"]
coef = lr.named_steps["clf"].coef_
log_or = coef[AC] - coef[INC]  # log P(acidente)/P(incidente) por unidade do atributo codificado
X_tr_imp = pd.DataFrame(lr.named_steps["imputar"].transform(X_tr), columns=NUM + CAT)
CURTO = {"tipo_operacao": "Operação", "tipo_equipamento": "Equipamento", "tipo_motor": "Motor", "qtd_motores": "Motores",
         "fase_voo_grupo": "Fase", "regiao": "Região", "local": "Local", "fabricante_grupo": "Fabricante",
         "categoria_tipo": "Tipo", "mes": "Mês", "dia_semana": "Dia"}
NIVEL = {"AERÓDROMOS SB": "aeródromos SB", "DEMAIS AERÓDROMOS": "demais aeródromos", "FORA DE AERÓDROMO": "fora de aeródromo",
         "AGUSTA/LEONARDO": "Agusta/Leonardo", "AIRBUS HELICOPTERS/HELIBRAS": "Airbus Helic./Helibras", "ATR": "ATR",
         "PERDA DE CONTROLE OU COLISÃO EM VOO": "perda de controle/colisão em voo", "PISTA, POUSO E SOLO": "pista, pouso e solo",
         "COLISÃO COM FAUNA": "colisão com fauna", "FALHA TÉCNICA": "falha técnica", "TRÁFEGO AÉREO": "tráfego aéreo",
         "APROXIMAÇÃO E POUSO": "aproximação e pouso", "DECOLAGEM E SUBIDA": "decolagem e subida",
         "CRUZEIRO E DESCIDA": "cruzeiro e descida", "TÁXI AÉREO": "táxi aéreo", "NÃO REGULAR": "não regular"}


def rot_nivel(col, c):
    c = str(c)
    if col == "mes":
        return c
    if c in NIVEL:
        return NIVEL[c]
    return c.capitalize() if col in ("fabricante_grupo", "dia_semana") else c.lower()


linhas_or = [{"atributo": f"{NOME[n]} (por desvio-padrão)", "grupo": NOME[n], "nivel": "", "referencia": "", "log_or": log_or[i]}
             for i, n in enumerate(NUM)]
pos = len(NUM)
for col, cats in zip(CAT, enc.named_transformers_["cat"].categories_):
    cats = list(cats)
    ref = X_tr_imp[col].mode()[0]  # nivel de referencia: o mais frequente no treino
    b_ref = log_or[pos + cats.index(ref)]
    for j, c in enumerate(cats):
        if c != ref:
            linhas_or.append({"atributo": f"{CURTO[col]}: {rot_nivel(col, c)} vs. {rot_nivel(col, ref)}", "grupo": NOME[col],
                              "nivel": c, "referencia": ref, "log_or": log_or[pos + j] - b_ref})
    pos += len(cats)
or_df = pd.DataFrame(linhas_or)
or_df["or"] = np.exp(or_df.log_or)
or_df = or_df.sort_values("log_or")
or_df.to_csv(RES / "lr_razoes_de_chance_acidente_vs_incidente.csv", index=False)
top = pd.concat([or_df.head(8), or_df.tail(8)])

rf = ajustados[principal["RF"]]
perm = permutation_importance(rf, X_te, y_te, scoring="roc_auc_ovr", n_repeats=30, random_state=SEED, n_jobs=-1)
imp_df = pd.DataFrame({"atributo": [NOME[c] for c in X_te.columns], "media": perm.importances_mean, "dp": perm.importances_std})
imp_df = imp_df.sort_values("media", ascending=True)
imp_df.to_csv(RES / "rf_importancia_permutacao.csv", index=False)

fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 3.6), gridspec_kw={"width_ratios": [1.15, 1]})
cores = [LARANJA if v > 0 else AZUL for v in top.log_or]
a1.barh(range(len(top)), top.log_or, color=cores)
a1.set_yticks(range(len(top))); a1.set_yticklabels(top.atributo, fontsize=6.6)
a1.axvline(0, color="black", lw=0.8)
a1.set_xlabel("log da razão de chances, acidente vs. incidente", fontsize=8)
a1.set_title("(a) Regressão Logística", fontsize=9)
a2.barh(range(len(imp_df)), imp_df.media, xerr=imp_df.dp, color=AZUL, error_kw={"lw": 0.7})
a2.set_yticks(range(len(imp_df))); a2.set_yticklabels(imp_df.atributo, fontsize=6.6)
a2.set_xlabel("Queda da AUC ao permutar", fontsize=8)
a2.set_title("(b) Random Forest", fontsize=9)
for a in (a1, a2):
    a.spines[["top", "right"]].set_visible(False)
    a.tick_params(axis="x", labelsize=7.5)
fig.tight_layout(w_pad=1.0)
fig.savefig(FIG / "fig5_importancias.png", dpi=220)
plt.close(fig)

# ================================================================== persistencia
resultados["lr_razoes_de_chance_top"] = top.assign(atributo=top.atributo.values).to_dict("records")
resultados["rf_importancia"] = imp_df.sort_values("media", ascending=False).to_dict("records")
pd.DataFrame({"codigo_ocorrencia": b.loc[X_te.index, "codigo_ocorrencia"].values, "y_real": y_te,
              **{f"p_{m}_{c}": te_proba[principal[m]][:, COD[c]] for m in ["LR", "RF"] for c in CLASSES}}
             ).to_csv(RES / "predicoes_teste.csv", index=False)
(RES / "metricas.json").write_text(json.dumps(resultados, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
print("ok", FIG, TAB, RES)
