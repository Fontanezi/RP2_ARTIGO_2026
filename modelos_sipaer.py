# -*- coding: utf-8 -*-
"""Modelagem da BASE_SIPAER: classificacao em tres classes (INCIDENTE, INCIDENTE GRAVE, ACIDENTE).

Uso: python modelos_sipaer.py [BASE_SIPAER.xlsx]

Etapas:
 1. Ablacao de calendario (mes, dia da semana, hora) com os balanceamentos do relatorio parcial; o conjunto
    reduzido e adotado se o F1-macro medio da validacao cruzada nao cair mais de TOLERANCIA em nenhum modelo.
 2. Comparacao de quatro estrategias de balanceamento (nenhum, pesos, SMOTENC, ambos) no conjunto final,
    para Regressao Logistica multinomial e Random Forest; busca em grade por validacao cruzada estratificada
    (k = 5) no treino (particao 80/20), maximizando F1-macro.
 3. Ablacao de atributos com o balanceamento selecionado: completo, sem calendario, sem descricao do evento
    (categoria do tipo, saida de pista, numero de tipos) e com codificacao pelo alvo de aerodromo e tipo detalhado.
 4. Validacao cruzada repetida (10 x 5) com os hiperparametros selecionados: media e desvio-padrao das metricas.
 5. Validacao temporal: treino em 2016-2022, teste em 2023-2026.
 6. Limiar de decisao para sensibilidade >= ALVO_SENS na classe acidente e avaliacao da calibracao (Brier,
    diagrama de confiabilidade).

Saidas: artigo/figuras/fig3..fig6; artigo/tabelas/tab_ablacao.tex, tab_atributos.tex, tab_metricas.tex,
        tab_temporal.tex; resultados/metricas.json e csvs.
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
from sklearn.metrics import (average_precision_score, balanced_accuracy_score, brier_score_loss, confusion_matrix,
                             f1_score, precision_recall_curve, precision_recall_fscore_support, roc_auc_score, roc_curve)
from sklearn.model_selection import GridSearchCV, StratifiedKFold, cross_val_predict, train_test_split
from sklearn.preprocessing import OneHotEncoder, StandardScaler, TargetEncoder

# numpy 2.0 + BLAS Accelerate (macOS) emite avisos espurios em matmul; resultados conferidos contra einsum.
warnings.filterwarnings("ignore", message=".*encountered in matmul")

BASE = Path(sys.argv[1] if len(sys.argv) > 1 else "BASE_SIPAER.xlsx")
RAIZ = Path(__file__).resolve().parent
FIG, TAB, RES = RAIZ / "artigo" / "figuras", RAIZ / "artigo" / "tabelas", RAIZ / "resultados"
for d in (FIG, TAB, RES):
    d.mkdir(parents=True, exist_ok=True)

SEED = 42
TOLERANCIA = 0.01      # queda maxima de F1-macro CV aceita para adotar o conjunto sem calendario
ALVO_SENS = 0.80       # sensibilidade minima para acidentes no ajuste do limiar
N_REPETICOES = 10      # repeticoes da validacao cruzada 5-fold
ANO_CORTE = 2023       # validacao temporal: treino < ANO_CORTE, teste >= ANO_CORTE
CLASSES = ["INCIDENTE", "INCIDENTE GRAVE", "ACIDENTE"]
ROTULO = {"INCIDENTE": "Incidente", "INCIDENTE GRAVE": "Inc. grave", "ACIDENTE": "Acidente"}
COD = {c: i for i, c in enumerate(CLASSES)}
INC, IG, AC = 0, 1, 2
NOME_MODELO = {"LR": "Regressão Logística", "RF": "Random Forest"}

# ================================================================== atributos
NUM_BASE = ["log_pmd", "assentos", "idade_aeronave", "n_tipos", "total_aeronaves_envolvidas", "saida_pista"]
CAT_BASE = ["tipo_operacao", "tipo_equipamento", "tipo_motor", "qtd_motores", "fase_voo_grupo", "regiao", "local",
            "fabricante_grupo", "categoria_tipo"]
CAL_NUM, CAL_CAT = ["hora"], ["mes", "dia_semana"]
EVENTO = ["categoria_tipo", "saida_pista", "n_tipos"]
TE_COLS = ["aerodromo", "tipo_principal"]
NOME = {"log_pmd": "log10 PMD (kg)", "assentos": "Assentos", "idade_aeronave": "Idade da aeronave", "hora": "Hora",
        "n_tipos": "Nº de tipos de ocorrência", "total_aeronaves_envolvidas": "Aeronaves envolvidas",
        "saida_pista": "Saída de pista", "tipo_operacao": "Tipo de operação", "tipo_equipamento": "Equipamento",
        "tipo_motor": "Tipo de motor", "qtd_motores": "Qtd. de motores", "fase_voo_grupo": "Fase do voo",
        "regiao": "Região", "local": "Local", "fabricante_grupo": "Fabricante", "categoria_tipo": "Categoria do tipo",
        "mes": "Mês", "dia_semana": "Dia da semana", "aerodromo": "Aeródromo (alvo)", "tipo_principal": "Tipo detalhado (alvo)"}
# ordem das colunas do conjunto completo identica a do relatorio parcial (afeta os sorteios do RF e do SMOTENC)
CONJUNTOS = {"completo": (["log_pmd", "assentos", "idade_aeronave", "hora", "n_tipos", "total_aeronaves_envolvidas", "saida_pista"], CAT_BASE + CAL_CAT, []),
             "sem_calendario": (NUM_BASE, CAT_BASE, [])}
ROT_CONJ = {"completo": "Completo", "sem_calendario": "Sem calendário", "sem_evento": "Sem descrição do evento",
            "com_te": "Com codificação pelo alvo"}


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
    X["aerodromo"] = b.aerodromo
    X["tipo_principal"] = b.tipo_principal
    return X


def colunas(conj):
    num, cat, te = CONJUNTOS[conj]
    return num + cat + te


b = pd.read_excel(BASE, sheet_name="base")
X_all = preparar(b)
y_all = b.classificacao.map(COD).to_numpy()
X_tr, X_te, y_tr, y_te = train_test_split(X_all, y_all, test_size=0.2, stratify=y_all, random_state=SEED)
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
print(f"base {X_all.shape}, treino {len(X_tr)} {np.bincount(y_tr).tolist()}, teste {len(X_te)} {np.bincount(y_te).tolist()}")


# ================================================================== pipeline
def pipeline(modelo, balanceamento, conj):
    """imputar (treino) -> [codificar pelo alvo] -> [SMOTENC] -> padronizar + one-hot -> modelo.
    Tudo e ajustado apenas nas dobras de treino; o SMOTENC so atua em fit; a codificacao pelo alvo usa
    cross-fitting interno (TargetEncoder) para nao vazar o rotulo da propria linha."""
    num, cat, te = CONJUNTOS[conj]
    n_num, n_cat, n_te = len(num), len(cat), len(te)
    idx_num = list(range(n_num))
    idx_cat = list(range(n_num, n_num + n_cat))
    idx_te = list(range(n_num + n_cat, n_num + n_cat + n_te))
    transf = [("num", SimpleImputer(strategy="median"), num), ("cat", SimpleImputer(strategy="most_frequent"), cat)]
    if te:
        transf.append(("te", SimpleImputer(strategy="most_frequent"), te))
    passos = [("imputar", ColumnTransformer(transf))]
    idx_te_saida = []
    if te:
        passos.append(("alvo", ColumnTransformer([("pass", "passthrough", idx_num + idx_cat),
                                                  ("te", TargetEncoder(target_type="multiclass", cv=5, random_state=SEED), idx_te)])))
        idx_te_saida = list(range(n_num + n_cat, n_num + n_cat + len(CLASSES) * n_te))
    pesos = "balanced" if balanceamento in ("pesos", "ambos") else None
    smote = SMOTENC(categorical_features=idx_cat, random_state=SEED) if balanceamento in ("smote", "ambos") else "passthrough"
    passos.append(("smote", smote))
    passos.append(("codificar", ColumnTransformer([("num", StandardScaler(), idx_num + idx_te_saida),
                                                   ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), idx_cat)])))
    if modelo == "LR":
        clf = LogisticRegression(max_iter=5000, class_weight=pesos, random_state=SEED)
    else:
        clf = RandomForestClassifier(n_estimators=400, class_weight=pesos, random_state=SEED, n_jobs=1)
    passos.append(("clf", clf))
    return Pipeline(passos)


GRADE = {"LR": {"clf__C": [0.01, 0.1, 1.0, 10.0]},
         "RF": {"clf__max_depth": [None, 10], "clf__min_samples_leaf": [1, 5, 20], "clf__max_features": ["sqrt", 0.3]}}
BALANCEAMENTOS = ["nenhum", "pesos", "smote", "ambos"]
ROT_BAL = {"nenhum": "Nenhum", "pesos": "Pesos de classe", "smote": "SMOTENC", "ambos": "SMOTENC + pesos"}


def metricas(y_true, y_pred, proba):
    p, r, f, _ = precision_recall_fscore_support(y_true, y_pred, labels=[INC, IG, AC], zero_division=0)
    m = {"f1_macro": f1_score(y_true, y_pred, average="macro"),
         "acuracia_balanceada": balanced_accuracy_score(y_true, y_pred),
         "acuracia": float((y_true == y_pred).mean()),
         "n_previstos_ac": int((y_pred == AC).sum()), "n_previstos_ig": int((y_pred == IG).sum())}
    for i, c in enumerate(["inc", "ig", "ac"]):
        m[f"prec_{c}"], m[f"rec_{c}"], m[f"f1_{c}"] = float(p[i]), float(r[i]), float(f[i])
    if proba is not None:
        m["auc_ovr_macro"] = roc_auc_score(y_true, proba, multi_class="ovr", average="macro")
        m["auc_ac"] = roc_auc_score((y_true == AC).astype(int), proba[:, AC])
        m["auc_ig"] = roc_auc_score((y_true == IG).astype(int), proba[:, IG])
        m["ap_ac"] = average_precision_score((y_true == AC).astype(int), proba[:, AC])
        m["ap_ig"] = average_precision_score((y_true == IG).astype(int), proba[:, IG])
        m["brier_ac"] = brier_score_loss((y_true == AC).astype(int), proba[:, AC])
    m["matriz"] = confusion_matrix(y_true, y_pred, labels=[INC, IG, AC]).tolist()
    return m


cache = {}


def ajustar(modelo, bal, conj, X_treino, y_treino, tag="aleatoria"):
    """Busca em grade (5-fold, F1-macro) + predicoes out-of-fold; resultados em cache por configuracao."""
    chave = f"{modelo}|{bal}|{conj}|{tag}"
    if chave not in cache:
        Xc = X_treino[colunas(conj)]
        gs = GridSearchCV(pipeline(modelo, bal, conj), GRADE[modelo], scoring="f1_macro", cv=cv, n_jobs=-1, refit=True)
        gs.fit(Xc, y_treino)
        proba_oof = cross_val_predict(clone(gs.best_estimator_), Xc, y_treino, cv=cv, method="predict_proba", n_jobs=-1)
        cache[chave] = {"modelo_ajustado": gs.best_estimator_, "params": gs.best_params_, "oof_proba": proba_oof,
                        "cv_media": float(gs.cv_results_["mean_test_score"][gs.best_index_]),
                        "cv_dp": float(gs.cv_results_["std_test_score"][gs.best_index_]),
                        "oof": metricas(y_treino, proba_oof.argmax(1), proba_oof)}
        r = cache[chave]
        print(f"{chave:40s} F1m CV {r['cv_media']:.3f}±{r['cv_dp']:.3f} | OOF recAC {r['oof']['rec_ac']:.2f} precAC {r['oof']['prec_ac']:.2f} "
              f"recIG {r['oof']['rec_ig']:.2f} AUCac {r['oof']['auc_ac']:.3f} | {r['params']}")
    return cache[chave]


# ================================================================== 1. ablacao de calendario -> conjunto final
BAL_PARCIAL = {"LR": "nenhum", "RF": "smote"}  # balanceamentos selecionados no relatorio parcial
cal = {(m, conj): ajustar(m, BAL_PARCIAL[m], conj, X_tr, y_tr)["cv_media"] for conj in ["completo", "sem_calendario"] for m in ["LR", "RF"]}
queda = {m: cal[(m, "completo")] - cal[(m, "sem_calendario")] for m in ["LR", "RF"]}
BASE_FINAL = "sem_calendario" if all(q <= TOLERANCIA for q in queda.values()) else "completo"
print(f"ablacao de calendario: quedas de F1-macro {queda} -> conjunto final: {BASE_FINAL}")
num_f, cat_f, _ = CONJUNTOS[BASE_FINAL]
CONJUNTOS["sem_evento"] = ([c for c in num_f if c not in EVENTO], [c for c in cat_f if c not in EVENTO], [])
CONJUNTOS["com_te"] = (num_f, cat_f, TE_COLS)
NUM, CAT = num_f, cat_f

# ================================================================== 2. balanceamento no conjunto final
for modelo in ["LR", "RF"]:
    for bal in BALANCEAMENTOS:
        ajustar(modelo, bal, BASE_FINAL, X_tr, y_tr)
principal = {}
for modelo in ["LR", "RF"]:
    cands = {bal: cache[f"{modelo}|{bal}|{BASE_FINAL}|aleatoria"]["cv_media"] for bal in BALANCEAMENTOS}
    principal[modelo] = max(cands, key=cands.get)
print("balanceamento principal:", principal)
final = {m: cache[f"{m}|{principal[m]}|{BASE_FINAL}|aleatoria"] for m in ["LR", "RF"]}
oof_proba = {m: final[m]["oof_proba"] for m in ["LR", "RF"]}
te_proba = {m: final[m]["modelo_ajustado"].predict_proba(X_te[colunas(BASE_FINAL)]) for m in ["LR", "RF"]}
teste_final = {m: metricas(y_te, te_proba[m].argmax(1), te_proba[m]) for m in ["LR", "RF"]}

# ================================================================== 3. ablacao de atributos
ORDEM_CONJ = ["completo", "sem_calendario", "sem_evento", "com_te"]
for modelo in ["LR", "RF"]:
    for conj in ORDEM_CONJ:
        ajustar(modelo, principal[modelo], conj, X_tr, y_tr)

# ================================================================== 4. validacao cruzada repetida
repetida = {m: [] for m in ["LR", "RF"]}
for r in range(N_REPETICOES):
    cv_r = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED + r)  # r = 0 reproduz a particao principal
    for m in ["LR", "RF"]:
        proba = oof_proba[m] if r == 0 else cross_val_predict(clone(final[m]["modelo_ajustado"]), X_tr[colunas(BASE_FINAL)], y_tr,
                                                              cv=cv_r, method="predict_proba", n_jobs=-1)
        repetida[m].append(metricas(y_tr, proba.argmax(1), proba))
CHAVES_REP = ["rec_ac", "prec_ac", "rec_ig", "prec_ig", "f1_macro", "acuracia_balanceada", "auc_ac", "auc_ig", "ap_ac", "brier_ac"]
resumo_rep = {m: {k: {"media": float(np.mean([x[k] for x in repetida[m]])), "dp": float(np.std([x[k] for x in repetida[m]], ddof=1)),
                      "min": float(np.min([x[k] for x in repetida[m]])), "max": float(np.max([x[k] for x in repetida[m]]))}
                  for k in CHAVES_REP} for m in ["LR", "RF"]}
# diferenca pareada RF - LR em cada repeticao (mesmas particoes para os dois modelos)
resumo_rep["diferenca_RF_menos_LR"] = {}
for k in CHAVES_REP:
    d = np.array([rf_[k] - lr_[k] for rf_, lr_ in zip(repetida["RF"], repetida["LR"])])
    resumo_rep["diferenca_RF_menos_LR"][k] = {"media": float(d.mean()), "dp": float(d.std(ddof=1)), "n_positivas": int((d > 0).sum()), "n": int(len(d))}
print("RF - LR por repeticao:", {k: f"{v['media']:+.3f}±{v['dp']:.3f} ({v['n_positivas']}/{v['n']} > 0)" for k, v in resumo_rep["diferenca_RF_menos_LR"].items()})
for m in ["LR", "RF"]:
    print(f"CV repetida {m}: F1m {resumo_rep[m]['f1_macro']['media']:.3f}±{resumo_rep[m]['f1_macro']['dp']:.3f} "
          f"recAC {resumo_rep[m]['rec_ac']['media']:.3f}±{resumo_rep[m]['rec_ac']['dp']:.3f} precAC {resumo_rep[m]['prec_ac']['media']:.3f}±{resumo_rep[m]['prec_ac']['dp']:.3f}")

# ================================================================== 5. validacao temporal
ano = b.loc[X_all.index, "ano"]
mask_tr_t = (ano < ANO_CORTE).to_numpy()
X_tr_t, y_tr_t, X_te_t, y_te_t = X_all[mask_tr_t], y_all[mask_tr_t], X_all[~mask_tr_t], y_all[~mask_tr_t]
temporal = {"n_treino": int(mask_tr_t.sum()), "classes_treino": np.bincount(y_tr_t).tolist(),
            "n_teste": int((~mask_tr_t).sum()), "classes_teste": np.bincount(y_te_t).tolist(), "modelos": {}}


def baseline_regra(Xd):
    return np.where(Xd.tipo_operacao.isin(["TÁXI AÉREO", "NÃO REGULAR"]), AC, INC)


temporal["regra"] = metricas(y_te_t, baseline_regra(X_te_t), None)
for m in ["LR", "RF"]:
    r = ajustar(m, principal[m], BASE_FINAL, X_tr_t, y_tr_t, tag="temporal")
    proba_t = r["modelo_ajustado"].predict_proba(X_te_t[colunas(BASE_FINAL)])
    temporal["modelos"][m] = {"params": r["params"], "cv_media": r["cv_media"], "cv_dp": r["cv_dp"],
                              "teste": metricas(y_te_t, proba_t.argmax(1), proba_t)}
    t = temporal["modelos"][m]["teste"]
    print(f"temporal {m}: recAC {t['rec_ac']:.2f} precAC {t['prec_ac']:.2f} recIG {t['rec_ig']:.2f} F1m {t['f1_macro']:.3f} AUCac {t['auc_ac']:.3f}")


# ================================================================== 6. limiar e calibracao
def limiar_para_sensibilidade(y_bin, p, alvo):
    """Maior limiar em p que garante sensibilidade >= alvo."""
    ordem = np.argsort(-p)
    recall = np.cumsum(y_bin[ordem]) / y_bin.sum()
    return float(p[ordem][int(np.argmax(recall >= alvo))])


def prever_com_limiar(proba, tau):
    """Acidente se P(acidente) >= tau; caso contrario, a mais provavel entre incidente e incidente grave."""
    resto = np.where(proba[:, IG] > proba[:, INC], IG, INC)
    return np.where(proba[:, AC] >= tau, AC, resto)


QUANTIS_CALIB = [0, 0.5, 0.7, 0.8, 0.9, 0.95, 1.0]  # metade inferior em uma faixa; faixas finas onde ficam os alertas


def calibracao(y_bin, p, quantis=QUANTIS_CALIB):
    bins = np.unique(np.quantile(p, quantis))
    idx = np.clip(np.digitize(p, bins[1:-1]), 0, len(bins) - 2)
    linhas = []
    for k in range(len(bins) - 1):
        sel = idx == k
        if sel.sum():
            linhas.append({"p_medio": float(p[sel].mean()), "observado": float(y_bin[sel].mean()), "n": int(sel.sum())})
    return linhas


limiar, calib = {}, {}
y_tr_ac, y_te_ac = (y_tr == AC).astype(int), (y_te == AC).astype(int)
prev_ac = float(y_tr_ac.mean())
for m in ["LR", "RF"]:
    tau = limiar_para_sensibilidade(y_tr_ac, oof_proba[m][:, AC], ALVO_SENS)
    limiar[m] = {"tau": tau, "oof": metricas(y_tr, prever_com_limiar(oof_proba[m], tau), oof_proba[m]),
                 "teste": metricas(y_te, prever_com_limiar(te_proba[m], tau), te_proba[m])}
    calib[m] = {"oof": calibracao(y_tr_ac, oof_proba[m][:, AC]), "brier_oof": float(brier_score_loss(y_tr_ac, oof_proba[m][:, AC])),
                "brier_teste": float(brier_score_loss(y_te_ac, te_proba[m][:, AC])),
                "p_medio_oof": float(oof_proba[m][:, AC].mean()), "p_medio_teste": float(te_proba[m][:, AC].mean())}
    print(f"limiar {m}: tau={tau:.3f} | OOF recAC {limiar[m]['oof']['rec_ac']:.2f} precAC {limiar[m]['oof']['prec_ac']:.2f} "
          f"alertas {limiar[m]['oof']['n_previstos_ac']} | teste recAC {limiar[m]['teste']['rec_ac']:.2f} precAC {limiar[m]['teste']['prec_ac']:.2f} "
          f"| Brier OOF {calib[m]['brier_oof']:.4f} (prevalência {prev_ac * (1 - prev_ac):.4f}) p_medio {calib[m]['p_medio_oof']:.3f} vs prev {prev_ac:.3f}")

# baselines e precisao na sensibilidade da regra
resultados = {"conjunto_final": BASE_FINAL, "ablacao_calendario_cv": {f"{m}|{c}": v for (m, c), v in cal.items()}, "quedas_calendario": queda,
              "n_treino": int(len(X_tr)), "n_teste": int(len(X_te)), "classes_treino": np.bincount(y_tr).tolist(), "classes_teste": np.bincount(y_te).tolist(),
              "baselines": {"majoritaria": {"treino": metricas(y_tr, np.full(len(y_tr), INC), None), "teste": metricas(y_te, np.full(len(y_te), INC), None)},
                            "regra": {"treino": metricas(y_tr, baseline_regra(X_tr), None), "teste": metricas(y_te, baseline_regra(X_te), None)}},
              "principal": principal, "balanceamento": {}, "atributos": {}, "final": {}, "cv_repetida": resumo_rep, "temporal": temporal,
              "limiar": limiar, "calibracao": calib, "prevalencia_ac_treino": prev_ac}
for m in ["LR", "RF"]:
    for bal in BALANCEAMENTOS:
        r = cache[f"{m}|{bal}|{BASE_FINAL}|aleatoria"]
        resultados["balanceamento"][f"{m}|{bal}"] = {"params": r["params"], "cv_media": r["cv_media"], "cv_dp": r["cv_dp"], "oof": r["oof"]}
    for conj in ORDEM_CONJ:
        r = cache[f"{m}|{principal[m]}|{conj}|aleatoria"]
        resultados["atributos"][f"{m}|{conj}"] = {"n_atributos": len(colunas(conj)), "params": r["params"], "cv_media": r["cv_media"], "cv_dp": r["cv_dp"], "oof": r["oof"]}
    resultados["final"][m] = {"balanceamento": principal[m], "params": final[m]["params"], "cv_media": final[m]["cv_media"], "cv_dp": final[m]["cv_dp"],
                              "oof": final[m]["oof"], "teste": teste_final[m]}
rec_regra = resultados["baselines"]["regra"]["treino"]["rec_ac"]


def precisao_na_sensibilidade(y_bin, score, alvo):
    prec, rec, _ = precision_recall_curve(y_bin, score)
    return float(prec[rec >= alvo].max())


resultados["precisao_na_sensibilidade_da_regra"] = {"sensibilidade_regra": rec_regra}
for m in ["LR", "RF"]:
    resultados["precisao_na_sensibilidade_da_regra"][f"{m}_oof"] = precisao_na_sensibilidade(y_tr_ac, oof_proba[m][:, AC], rec_regra)
    resultados["precisao_na_sensibilidade_da_regra"][f"{m}_teste"] = precisao_na_sensibilidade(y_te_ac, te_proba[m][:, AC], rec_regra)


# ================================================================== tabelas
def pt(x, casas=2):
    return f"{x:.{casas}f}".replace(".", ",")


def pc(x):
    return f"{100 * x:.1f}".replace(".", ",")


def pm(k, m, pct=True, casas=3):
    """media ± dp da validacao cruzada repetida."""
    s = resumo_rep[m][k]
    return f"{pc(s['media'])} $\\pm$ {pc(s['dp'])}" if pct else f"{pt(s['media'], casas)} $\\pm$ {pt(s['dp'], casas)}"


# Tabela: balanceamento (conjunto final); "ambos" e identico a "smote" (pesos unitarios apos o SMOTENC) e fica fora da tabela
for m in ["LR", "RF"]:
    a, s_ = resultados["balanceamento"][f"{m}|ambos"], resultados["balanceamento"][f"{m}|smote"]
    assert abs(a["cv_media"] - s_["cv_media"]) < 1e-9 and a["oof"]["matriz"] == s_["oof"]["matriz"], f"ambos != smote para {m}"
linhas = []
for m in ["LR", "RF"]:
    for bal in ["nenhum", "pesos", "smote"]:
        r = resultados["balanceamento"][f"{m}|{bal}"]
        o = r["oof"]
        marca = r"$\dagger$" if bal == principal[m] else ""
        linhas.append(f"{NOME_MODELO[m] if bal == 'nenhum' else ''} & {ROT_BAL[bal]}{marca} & {pt(r['cv_media'], 3)} $\\pm$ {pt(r['cv_dp'], 3)} & "
                      f"{pc(o['rec_ac'])} & {pc(o['prec_ac'])} & {pc(o['rec_ig'])} & {pc(o['prec_ig'])} & {pt(o['auc_ac'], 3)} \\\\")
    if m == "LR":
        linhas.append(r"\midrule")
(TAB / "tab_ablacao.tex").write_text(
    "\\begin{tabular}{llcccccc}\n\\toprule\n\\textbf{Modelo} & \\textbf{Balanceamento} & \\textbf{F1-macro CV} & \\textbf{Sens. AC} & "
    "\\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & \\textbf{AUC AC} \\\\\n\\midrule\n" + "\n".join(linhas) +
    "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# Tabela: ablacao de atributos (balanceamento selecionado)
linhas = []
for m in ["LR", "RF"]:
    for conj in ORDEM_CONJ:
        r = resultados["atributos"][f"{m}|{conj}"]
        o = r["oof"]
        marca = r"$\dagger$" if conj == BASE_FINAL else ""
        linhas.append(f"{NOME_MODELO[m] if conj == 'completo' else ''} & {ROT_CONJ[conj]} ({r['n_atributos']}){marca} & {pt(r['cv_media'], 3)} $\\pm$ {pt(r['cv_dp'], 3)} & "
                      f"{pc(o['rec_ac'])} & {pc(o['prec_ac'])} & {pc(o['rec_ig'])} & {pc(o['prec_ig'])} & {pt(o['auc_ac'], 3)} \\\\")
    if m == "LR":
        linhas.append(r"\midrule")
(TAB / "tab_atributos.tex").write_text(
    "\\begin{tabular}{llcccccc}\n\\toprule\n\\textbf{Modelo} & \\textbf{Conjunto de atributos ($p$)} & \\textbf{F1-macro CV} & \\textbf{Sens. AC} & "
    "\\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & \\textbf{AUC AC} \\\\\n\\midrule\n" + "\n".join(linhas) +
    "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")


# Tabela: selecao (balanceamento + atributos em um so ambiente)
bal_rows = [l for l in (TAB / "tab_ablacao.tex").read_text(encoding="utf-8").splitlines() if l.endswith("\\\\") and "textbf" not in l or l == r"\midrule"]
atr_rows = [l for l in (TAB / "tab_atributos.tex").read_text(encoding="utf-8").splitlines() if l.endswith("\\\\") and "textbf" not in l or l == r"\midrule"]
(TAB / "tab_selecao.tex").write_text(
    "\\begin{tabular}{llcccccc}\n\\toprule\n\\textbf{Modelo} & \\textbf{Configuração} & \\textbf{F1-macro CV} & \\textbf{Sens. AC} & "
    "\\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & \\textbf{AUC AC} \\\\\n\\midrule\n"
    f"\\multicolumn{{8}}{{l}}{{\\textit{{Balanceamento (conjunto completo, {len(colunas('completo'))} atributos)}}}} \\\\\n" + "\n".join(bal_rows) +
    "\n\\midrule\n\\multicolumn{8}{l}{\\textit{Conjunto de atributos (balanceamento selecionado); entre parênteses, o número de atributos}} \\\\\n" +
    "\n".join(atr_rows) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# Tabela: desempenho (CV repetida com ± e teste)
def linha_ponto(nome, m, proba=True):
    return (f"{nome} & {pc(m['rec_ac'])} & {pc(m['prec_ac'])} & {pc(m['rec_ig'])} & {pc(m['prec_ig'])} & {pt(m['f1_macro'], 3)} & "
            f"{pt(m['auc_ac'], 3) if proba else '--'} \\\\")


def linha_rep(nome, m):
    return (f"{nome} & {pm('rec_ac', m)} & {pm('prec_ac', m)} & {pm('rec_ig', m)} & {pm('prec_ig', m)} & {pm('f1_macro', m, False)} & "
            f"{pm('auc_ac', m, False)} \\\\")


n_ac_tr, n_ig_tr = np.bincount(y_tr)[AC], np.bincount(y_tr)[IG]
n_ac_te, n_ig_te = np.bincount(y_te)[AC], np.bincount(y_te)[IG]
blocos = [f"\\multicolumn{{7}}{{l}}{{\\textit{{Validação cruzada no treino ({n_ac_tr} acidentes, {n_ig_tr} incidentes graves); modelos: média $\\pm$ desvio-padrão em {N_REPETICOES} repetições}}}} \\\\",
          linha_ponto("Classe majoritária", resultados["baselines"]["majoritaria"]["treino"], False),
          linha_ponto("Regra pelo tipo de operação", resultados["baselines"]["regra"]["treino"], False)]
for m in ["LR", "RF"]:
    blocos.append(linha_rep(f"{NOME_MODELO[m]} ({ROT_BAL[principal[m]]})", m))
blocos += [r"\midrule", f"\\multicolumn{{7}}{{l}}{{\\textit{{Teste ({n_ac_te} acidentes, {n_ig_te} incidentes graves)}}}} \\\\",
           linha_ponto("Classe majoritária", resultados["baselines"]["majoritaria"]["teste"], False),
           linha_ponto("Regra pelo tipo de operação", resultados["baselines"]["regra"]["teste"], False)]
for m in ["LR", "RF"]:
    blocos.append(linha_ponto(f"{NOME_MODELO[m]} ({ROT_BAL[principal[m]]})", teste_final[m]))
(TAB / "tab_metricas.tex").write_text(
    "\\begin{tabular}{lcccccc}\n\\toprule\n\\textbf{Modelo} & \\textbf{Sens. AC} & \\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & "
    "\\textbf{F1-macro} & \\textbf{AUC AC} \\\\\n\\midrule\n" + "\n".join(blocos) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# Tabela: validacao temporal
def mil(x):
    return f"{int(x):,}".replace(",", ".")


ct, cte = temporal["classes_treino"], temporal["classes_teste"]
blocos = [f"\\multicolumn{{7}}{{l}}{{\\textit{{Treino 2016--2022 ({mil(temporal['n_treino'])} ocorrências: {ct[AC]} acidentes e {ct[IG]} incidentes graves); "
          f"teste 2023--2026 ({mil(temporal['n_teste'])} ocorrências: {cte[AC]} e {cte[IG]})}}}} \\\\",
          linha_ponto("Regra pelo tipo de operação", temporal["regra"], False)]
for m in ["LR", "RF"]:
    blocos.append(linha_ponto(f"{NOME_MODELO[m]} ({ROT_BAL[principal[m]]})", temporal["modelos"][m]["teste"]))
(TAB / "tab_temporal.tex").write_text(
    "\\begin{tabular}{lcccccc}\n\\toprule\n\\textbf{Modelo} & \\textbf{Sens. AC} & \\textbf{Prec. AC} & \\textbf{Sens. IG} & \\textbf{Prec. IG} & "
    "\\textbf{F1-macro} & \\textbf{AUC AC} \\\\\n\\midrule\n" + "\n".join(blocos) + "\n\\bottomrule\n\\end{tabular}\n", encoding="utf-8")

# ================================================================== figuras
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})
AZUL, LARANJA = "#1f3b73", "#e39a4c"
cor = {"LR": AZUL, "RF": LARANJA}
rot_curto = [ROTULO[c] for c in CLASSES]


def num_pt(x, casas=2):
    return f"{x:.{casas}f}".replace(".", ",")


# Figura 3: matrizes de confusao (OOF)
fig, eixos = plt.subplots(1, 2, figsize=(7.2, 3.0))
for ax, m in zip(eixos, ["LR", "RF"]):
    mat = np.array(final[m]["oof"]["matriz"])
    norm = mat / mat.sum(1, keepdims=True)
    ax.imshow(norm, cmap="Blues", vmin=0, vmax=1)
    for i in range(3):
        for j in range(3):
            ax.text(j, i, f"{mat[i, j]}\n({100 * norm[i, j]:.0f}%)", ha="center", va="center", fontsize=8, color="white" if norm[i, j] > 0.6 else "black")
    ax.set_xticks(range(3)); ax.set_yticks(range(3)); ax.set_xticklabels(rot_curto); ax.set_yticklabels(rot_curto)
    ax.set_xlabel("Classe prevista"); ax.set_ylabel("Classe real")
    ax.set_title(f"({'a' if m == 'LR' else 'b'}) {NOME_MODELO[m]} ({ROT_BAL[principal[m]]})", fontsize=9)
fig.tight_layout(); fig.savefig(FIG / "fig3_confusao.png", dpi=220); plt.close(fig)

# Figura 4: ROC e precisao-sensibilidade (OOF)
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 3.0))
for m in ["LR", "RF"]:
    for classe, ls, alpha in [(AC, "-", 1.0), (IG, "--", 0.8)]:
        yb = (y_tr == classe).astype(int)
        fpr, tpr, _ = roc_curve(yb, oof_proba[m][:, classe]); prec, rec, _ = precision_recall_curve(yb, oof_proba[m][:, classe])
        auc = roc_auc_score(yb, oof_proba[m][:, classe]); ap = average_precision_score(yb, oof_proba[m][:, classe])
        a1.plot(fpr, tpr, ls, color=cor[m], alpha=alpha, lw=1.4, label=f"{NOME_MODELO[m]}, {ROTULO[CLASSES[classe]].lower()} (AUC {num_pt(auc)})")
        a2.plot(rec, prec, ls, color=cor[m], alpha=alpha, lw=1.4, label=f"{NOME_MODELO[m]}, {ROTULO[CLASSES[classe]].lower()} (AP {num_pt(ap)})")
mr = np.array(resultados["baselines"]["regra"]["treino"]["matriz"])
fpr_regra = (mr[INC, AC] + mr[IG, AC]) / (mr[INC].sum() + mr[IG].sum())
a1.plot(fpr_regra, rec_regra, "D", color="black", ms=4, label="Regra pelo tipo de operação")
a2.plot(rec_regra, resultados["baselines"]["regra"]["treino"]["prec_ac"], "D", color="black", ms=4, label="Regra pelo tipo de operação")
a1.plot([0, 1], [0, 1], ":", color="gray", lw=1)
a1.set_xlabel("Taxa de falsos positivos"); a1.set_ylabel("Sensibilidade"); a1.set_title("(a) ROC", fontsize=9)
a2.set_xlabel("Sensibilidade"); a2.set_ylabel("Precisão"); a2.set_title("(b) Precisão-sensibilidade", fontsize=9)
a2.axhline(prev_ac, ls=":", color="gray", lw=1)
a1.legend(fontsize=6.3, frameon=False, loc="lower right"); a2.legend(fontsize=6.3, frameon=False, loc="upper right")
for a in (a1, a2):
    a.set_xlim(0, 1); a.set_ylim(0, 1.02); a.spines[["top", "right"]].set_visible(False)
fig.tight_layout(); fig.savefig(FIG / "fig4_curvas.png", dpi=220); plt.close(fig)

# Figura 5: contrastes da LR e importancia por permutacao do RF
assert not CONJUNTOS[BASE_FINAL][2], "contrastes da LR pressupoem conjunto final sem codificacao pelo alvo"
lr = final["LR"]["modelo_ajustado"]
enc = lr.named_steps["codificar"]
coef = lr.named_steps["clf"].coef_
log_or = coef[AC] - coef[INC]
X_tr_imp = pd.DataFrame(lr.named_steps["imputar"].transform(X_tr[colunas(BASE_FINAL)]), columns=NUM + CAT)
CURTO = {"tipo_operacao": "Operação", "tipo_equipamento": "Equipamento", "tipo_motor": "Motor", "qtd_motores": "Motores", "fase_voo_grupo": "Fase",
         "regiao": "Região", "local": "Local", "fabricante_grupo": "Fabricante", "categoria_tipo": "Tipo", "mes": "Mês", "dia_semana": "Dia"}
NIVEL = {"AERÓDROMOS SB": "aeródromos SB", "DEMAIS AERÓDROMOS": "demais aeródromos", "FORA DE AERÓDROMO": "fora de aeródromo",
         "AGUSTA/LEONARDO": "Agusta/Leonardo", "AIRBUS HELICOPTERS/HELIBRAS": "Airbus Helic./Helibras", "ATR": "ATR",
         "PERDA DE CONTROLE OU COLISÃO EM VOO": "perda de controle/colisão em voo", "PISTA, POUSO E SOLO": "pista, pouso e solo",
         "COLISÃO COM FAUNA": "colisão com fauna", "FALHA TÉCNICA": "falha técnica", "TRÁFEGO AÉREO": "tráfego aéreo",
         "APROXIMAÇÃO E POUSO": "aproximação e pouso", "DECOLAGEM E SUBIDA": "decolagem e subida", "CRUZEIRO E DESCIDA": "cruzeiro e descida",
         "TÁXI AÉREO": "táxi aéreo", "NÃO REGULAR": "não regular"}


def rot_nivel(col, c):
    c = str(c)
    if col == "mes":
        return c
    if c in NIVEL:
        return NIVEL[c]
    return c.capitalize() if col in ("fabricante_grupo", "dia_semana") else c.lower()


linhas_or = [{"atributo": f"{NOME[n]} (por desvio-padrão)", "grupo": NOME[n], "nivel": "", "referencia": "", "log_or": log_or[i]} for i, n in enumerate(NUM)]
pos = len(NUM)
for col, cats in zip(CAT, enc.named_transformers_["cat"].categories_):
    cats = list(cats)
    ref = X_tr_imp[col].mode()[0]
    b_ref = log_or[pos + cats.index(ref)]
    for j, c in enumerate(cats):
        if c != ref:
            linhas_or.append({"atributo": f"{CURTO[col]}: {rot_nivel(col, c)} vs. {rot_nivel(col, ref)}", "grupo": NOME[col], "nivel": c, "referencia": ref,
                              "log_or": log_or[pos + j] - b_ref})
    pos += len(cats)
or_df = pd.DataFrame(linhas_or)
or_df["or"] = np.exp(or_df.log_or)
or_df = or_df.sort_values("log_or")
or_df.to_csv(RES / "lr_razoes_de_chance_acidente_vs_incidente.csv", index=False)
top = pd.concat([or_df.head(8), or_df.tail(8)])

rf = final["RF"]["modelo_ajustado"]
perm = permutation_importance(rf, X_te[colunas(BASE_FINAL)], y_te, scoring="roc_auc_ovr", n_repeats=30, random_state=SEED, n_jobs=-1)
imp_df = pd.DataFrame({"atributo": [NOME[c] for c in colunas(BASE_FINAL)], "media": perm.importances_mean, "dp": perm.importances_std}).sort_values("media")
imp_df.to_csv(RES / "rf_importancia_permutacao.csv", index=False)

fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 3.6), gridspec_kw={"width_ratios": [1.15, 1]})
a1.barh(range(len(top)), top.log_or, color=[LARANJA if v > 0 else AZUL for v in top.log_or])
a1.set_yticks(range(len(top))); a1.set_yticklabels(top.atributo, fontsize=6.6); a1.axvline(0, color="black", lw=0.8)
a1.set_xlabel("log da razão de chances, acidente vs. incidente", fontsize=8); a1.set_title("(a) Regressão Logística", fontsize=9)
a2.barh(range(len(imp_df)), imp_df.media, xerr=imp_df.dp, color=AZUL, error_kw={"lw": 0.7})
a2.set_yticks(range(len(imp_df))); a2.set_yticklabels(imp_df.atributo, fontsize=6.6)
a2.set_xlabel("Queda da AUC ao permutar", fontsize=8); a2.set_title("(b) Random Forest", fontsize=9)
for a in (a1, a2):
    a.spines[["top", "right"]].set_visible(False); a.tick_params(axis="x", labelsize=7.5)
fig.tight_layout(w_pad=1.0); fig.savefig(FIG / "fig5_importancias.png", dpi=220); plt.close(fig)

# Figura 6: calibracao e limiar
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 3.6))
a1.plot([1e-4, 1], [1e-4, 1], ":", color="gray", lw=1, label="Calibração perfeita")
for m in ["LR", "RF"]:
    c = calib[m]["oof"]
    a1.plot([x["p_medio"] for x in c], [max(x["observado"], 2e-4) for x in c], "o-", color=cor[m], ms=3.5, lw=1.2,
            label=f"{NOME_MODELO[m]} (Brier {num_pt(calib[m]['brier_oof'], 3)})")
a1.axvline(prev_ac, color="gray", lw=0.8, ls="--"); a1.axhline(prev_ac, color="gray", lw=0.8, ls="--")
a1.text(prev_ac * 1.15, 2.6e-4, "prevalência", fontsize=6.5, color="gray")
a1.set_xscale("log"); a1.set_yscale("log")
a1.set_xlim(2e-4, 1); a1.set_ylim(2e-4, 1)
a1.set_xlabel("Probabilidade prevista de acidente (média da faixa)"); a1.set_ylabel("Fração observada de acidentes")
a1.set_title("(a) Confiabilidade de P(acidente), out-of-fold", fontsize=9)
a1.legend(fontsize=6.5, frameon=False, loc="upper left")
taus = np.linspace(0.005, 0.95, 400)
for m in ["LR", "RF"]:
    p = oof_proba[m][:, AC]
    sens = [((p >= t) & (y_tr_ac == 1)).sum() / y_tr_ac.sum() for t in taus]
    prec = [((p >= t) & (y_tr_ac == 1)).sum() / max((p >= t).sum(), 1) for t in taus]
    a2.plot(taus, sens, "-", color=cor[m], lw=1.4, label=f"{NOME_MODELO[m]}: sensibilidade")
    a2.plot(taus, prec, "--", color=cor[m], lw=1.4, label=f"{NOME_MODELO[m]}: precisão")
    a2.axvline(limiar[m]["tau"], color=cor[m], ls=":", lw=1)
a2.axhline(ALVO_SENS, color="gray", ls=":", lw=1)
a2.text(0.93, ALVO_SENS + 0.02, f"sensibilidade {int(ALVO_SENS * 100)}%", fontsize=6.5, color="gray", ha="right")
a2.set_xlabel("Limiar de decisão para a classe acidente"); a2.set_ylabel("Sensibilidade / precisão (acidente)")
a2.set_title("(b) Sensibilidade e precisão em função do limiar", fontsize=9)
a2.set_xlim(0, 0.95); a2.set_ylim(0, 1.02)
a2.legend(fontsize=6.5, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=2)
for a in (a1, a2):
    a.spines[["top", "right"]].set_visible(False)
fig.tight_layout(); fig.savefig(FIG / "fig6_limiar_calibracao.png", dpi=220); plt.close(fig)

# ================================================================== persistencia
resultados["lr_razoes_de_chance_top"] = top.to_dict("records")
resultados["rf_importancia"] = imp_df.sort_values("media", ascending=False).to_dict("records")
pd.DataFrame({"codigo_ocorrencia": b.loc[X_te.index, "codigo_ocorrencia"].values, "y_real": y_te,
              **{f"p_{m}_{c}": te_proba[m][:, COD[c]] for m in ["LR", "RF"] for c in CLASSES}}).to_csv(RES / "predicoes_teste.csv", index=False)
(RES / "metricas.json").write_text(json.dumps(resultados, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
print("ok", FIG, TAB, RES)
