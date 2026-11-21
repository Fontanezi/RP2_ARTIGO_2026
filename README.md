# Previsão da severidade de ocorrências aeronáuticas na aviação comercial brasileira (RBAC 121 e 135)

Trabalho de RP2 (EACH-USP, 2026). Classificação da severidade de ocorrências do CENIPA/SIPAER em três
classes (incidente, incidente grave, acidente) com Regressão Logística multinomial e Random Forest.
Relatório final em `artigo/artigo.pdf`; o relatório parcial (out/2026) está no histórico do git.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `BASE_SIPAER.xlsx` | Base final (abas `base`, `dicionario`, `fonte`): 5.038 ocorrências, 44 colunas |
| `figuras_eda.py` | Análise exploratória em três classes: figuras anuais/geográficas e tabelas descritivas (`.tex`) |
| `modelos_sipaer.py` | Modelagem: LR e RF × 4 balanceamentos; ablação de atributos (calendário, descrição do evento, codificação pelo alvo); CV repetida 10×5; validação temporal (≤2022 → ≥2023); limiar para sensibilidade ≥ 80%; calibração (Brier, confiabilidade); figuras e tabelas do artigo |
| `eda_sipaer.py` | Versão anterior da EDA (alvo binário, saída em `.docx`); mantida como histórico |
| `gerar_slides.py` | Slides da apresentação em vídeo (17 slides, 16:9), lidos das tabelas `.tex` e figuras do artigo; notas com roteiro e tempo-alvo |
| `apresentacao/slides_relatorio_parcial.pptx` | Deck do relatório parcial (números da versão de outubro) |
| `artigo/artigo.tex` | Artigo (relatório parcial), com tabelas via `\input{tabelas/...}` |
| `artigo/artigo_overleaf.tex` | Mesmo artigo com as tabelas embutidas, para colar no Overleaf |
| `artigo/figuras/` | PNGs das figuras |
| `artigo/tabelas/` | Fragmentos `.tex` gerados pelos scripts |
| `artigo/artigo.bib` | Referências |
| `resultados/` | `metricas.json` (todas as métricas), razões de chance da LR, importâncias do RF, predições do teste |

## Reproduzir

```bash
python3 -m venv venv && source venv/bin/activate
pip install pandas==2.2.3 openpyxl==3.1.5 matplotlib==3.9.2 scikit-learn==1.5.2 imbalanced-learn==0.12.4
python figuras_eda.py BASE_SIPAER.xlsx      # ~5 s
python modelos_sipaer.py BASE_SIPAER.xlsx   # ~3 min (grades, ablações, CV repetida, validação temporal)
python3 gerar_slides.py                     # requer python-pptx
cd artigo && tectonic -X compile artigo.tex # ou compile no Overleaf (pdfLaTeX)
```

Sementes fixas (`random_state=42`): os números do artigo são reproduzidos exatamente.

## Decisões de modelagem

- Alvo com três classes (NSCA 3-13); escopo RBAC 121 + 135 (só a operação regular tem 12 acidentes).
- Preditores apenas pré-investigação; nível de dano, fatalidades, fatores contribuintes, recomendações,
  situação da investigação e ano ficam fora (vazamento / mudança de cobertura em 2023).
- Imputação, padronização, one-hot e SMOTENC ajustados apenas nas dobras de treino (pipeline imblearn).
- Partição 80/20 estratificada; grade com 5-fold estratificado otimizando F1-macro; predições
  out-of-fold do treino como resultado principal e teste como confirmação.
- Baselines: classe majoritária e regra "táxi aéreo ou não regular ⇒ acidente".
- Conjunto final: completo (18 atributos); a remoção do calendário reduziu o F1-macro CV em 0,02 (> tolerância 0,01).
- Resultado principal (CV repetida 10×5): RF+SMOTENC F1-macro 0,598 ± 0,010 (sens. acidente 56,8%, inc. grave 37,8%);
  LR 0,568 ± 0,016; AUC acidente 0,94–0,95 nos dois. Limiar para 80% de sensibilidade: LR τ = 0,029 (263 alertas, prec. 22,4%).
