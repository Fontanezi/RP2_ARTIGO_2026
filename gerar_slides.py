# -*- coding: utf-8 -*-
"""Slides da apresentacao do relatorio final (8-10 min), gerados a partir do artigo.

Uso: python3 gerar_slides.py            (python-pptx; sem dependencias de dados)
Le as tabelas de artigo/tabelas/*.tex e as figuras de artigo/figuras/, para que os numeros dos slides
sejam exatamente os do artigo. Saida: apresentacao/slides_relatorio_final.pptx.
Cada slide traz, nas notas do apresentador, o roteiro sugerido e o tempo-alvo.
(O deck do relatorio parcial, de outubro, permanece em apresentacao/slides_relatorio_parcial.pptx.)
"""
import re
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

RAIZ = Path(__file__).resolve().parent
FIG = RAIZ / "artigo" / "figuras"
TAB = RAIZ / "artigo" / "tabelas"
SAIDA = RAIZ / "apresentacao" / "slides_relatorio_final.pptx"
SAIDA.parent.mkdir(exist_ok=True)

AZUL = RGBColor(0x1F, 0x3B, 0x73)
LARANJA = RGBColor(0xE3, 0x9A, 0x4C)
VERMELHO = RGBColor(0xB2, 0x3A, 0x3A)
CINZA = RGBColor(0x59, 0x59, 0x59)
CINZA_CLARO = RGBColor(0xF2, 0xF2, 0xF2)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
PRETO = RGBColor(0x1A, 0x1A, 0x1A)
DESTAQUE = RGBColor(0xFD, 0xEB, 0xD3)
FONTE = "Calibri"

LARG, ALT = Inches(13.333), Inches(7.5)
MARGEM = Inches(0.6)
TOPO_CONTEUDO = Inches(1.35)
ALT_CONTEUDO = Inches(5.45)
LARG_CONTEUDO = LARG - 2 * MARGEM

prs = Presentation()
prs.slide_width, prs.slide_height = LARG, ALT
BRANCO_LAYOUT = prs.slide_layouts[6]
TOTAL_SLIDES = 19
RODAPE = "Previsão da severidade de ocorrências aeronáuticas na aviação comercial brasileira · Relatório final · RP2 EACH-USP"


# ------------------------------------------------------------------ utilitarios
def tex_limpo(s):
    """Converte um fragmento LaTeX de celula em texto simples; devolve (texto, negrito, italico, recuo)."""
    s = s.strip()
    negrito = "\\textbf{" in s
    italico = "\\textit{" in s
    recuo = "\\hspace" in s
    s = re.sub(r"\\hspace\{[^}]*\}", "", s)
    s = re.sub(r"\\multicolumn\{\d+\}\{[^}]*\}\{(.*)\}$", r"\1", s)
    for _ in range(3):
        s = re.sub(r"\\(textbf|textit)\{([^{}]*)\}", r"\2", s)
    s = s.replace("$\\pm$", "±").replace("$\\dagger$", "†").replace("\\%", "%").replace("--", "–")
    s = re.sub(r"\$n=([\d.]+)\$", r"n = \1", s)
    s = s.replace("$", "").replace("\\\\", "").strip()
    return s, negrito, italico, recuo


def ler_tabela_tex(nome):
    """Le um fragmento tabular e devolve lista de linhas; cada linha e lista de celulas (texto, negrito, italico, recuo)."""
    linhas = []
    for raw in (TAB / nome).read_text(encoding="utf-8").splitlines():
        raw = raw.strip()
        if not raw or raw.startswith("\\begin") or raw.startswith("\\end") or raw.endswith("rule"):
            continue
        raw = raw.rstrip("\\")
        linhas.append([tex_limpo(c) for c in re.split(r"\s*&\s*", raw)])
    return linhas


def caixa_texto(slide, x, y, w, h, texto, tam=18, cor=PRETO, negrito=False, alinh=PP_ALIGN.LEFT, ancora=MSO_ANCHOR.TOP, italico=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ancora
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = alinh
    r = p.add_run()
    r.text = texto
    r.font.name, r.font.size, r.font.bold, r.font.italic, r.font.color.rgb = FONTE, Pt(tam), negrito, italico, cor
    return tb


def topicos(slide, x, y, w, h, itens, tam=18, espaco=6):
    """Lista de topicos; cada item pode ser str ou (str, nivel). Trechos entre ** ficam em negrito."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    primeiro = True
    for item in itens:
        texto, nivel = (item, 0) if isinstance(item, str) else item
        p = tf.paragraphs[0] if primeiro else tf.add_paragraph()
        primeiro = False
        p.space_after = Pt(espaco)
        p.level = nivel
        marcador = "•  " if nivel == 0 else "–  "
        partes = texto.split("**")
        for k, parte in enumerate(partes):
            if k == 0:
                parte = (marcador if nivel == 0 else "      " + marcador) + parte
            if not parte:
                continue
            r = p.add_run()
            r.text = parte
            r.font.name, r.font.size, r.font.color.rgb = FONTE, Pt(tam - 2 * nivel), PRETO
            r.font.bold = k % 2 == 1
    return tb


def novo_slide(titulo, parte=None, numero=None):
    slide = prs.slides.add_slide(BRANCO_LAYOUT)
    barra = slide.shapes.add_shape(1, 0, 0, LARG, Inches(0.12))  # 1 = retangulo
    barra.fill.solid(); barra.fill.fore_color.rgb = AZUL; barra.line.fill.background()
    caixa_texto(slide, MARGEM, Inches(0.3), LARG_CONTEUDO, Inches(0.9), titulo, tam=30, cor=AZUL, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
    if parte:
        caixa_texto(slide, LARG - MARGEM - Inches(4.2), Inches(0.18), Inches(4.2), Inches(0.35), parte, tam=11, cor=CINZA, alinh=PP_ALIGN.RIGHT)
    caixa_texto(slide, MARGEM, ALT - Inches(0.45), LARG_CONTEUDO - Inches(1), Inches(0.3), RODAPE, tam=10, cor=CINZA)
    if numero:
        caixa_texto(slide, LARG - MARGEM - Inches(1), ALT - Inches(0.45), Inches(1), Inches(0.3), f"{numero}/{TOTAL_SLIDES}", tam=10, cor=CINZA, alinh=PP_ALIGN.RIGHT)
    return slide


def notas(slide, apresentador, tempo, roteiro):
    tf = slide.notes_slide.notes_text_frame
    tf.text = f"[{apresentador} | tempo-alvo: {tempo}]\n\n{roteiro}"


def figura(slide, nome, x, y, w_max, h_max):
    """Insere a figura preservando a proporcao dentro da caixa (x, y, w_max, h_max)."""
    pw, ph = Image.open(FIG / nome).size
    escala = min(w_max / pw, h_max / ph)
    w, h = int(pw * escala), int(ph * escala)
    return slide.shapes.add_picture(str(FIG / nome), x + (w_max - w) // 2, y + (h_max - h) // 2, w, h)


def tabela(slide, linhas, x, y, w, tam=12, larguras=None, alt_linha=Inches(0.3), destaque=()):
    """Tabela nativa a partir das linhas lidas do .tex. Primeira linha = cabecalho; linhas de uma celula = grupo."""
    ncol = max(len(l) for l in linhas)
    tb = slide.shapes.add_table(len(linhas), ncol, x, y, w, alt_linha * len(linhas)).table
    if larguras:
        total = sum(larguras)
        for j, frac in enumerate(larguras):
            tb.columns[j].width = int(w * frac / total)
    for i, linha in enumerate(linhas):
        tb.rows[i].height = alt_linha
        grupo = len(linha) == 1 and i > 0
        for j in range(ncol):
            cel = tb.cell(i, j)
            texto, negrito, italico, recuo = linha[j] if j < len(linha) else ("", False, False, False)
            cel.margin_left = cel.margin_right = Inches(0.06)
            cel.margin_top = cel.margin_bottom = Inches(0.02)
            cel.vertical_anchor = MSO_ANCHOR.MIDDLE
            cel.fill.solid()
            if i == 0:
                cel.fill.fore_color.rgb = AZUL
            elif i in destaque:
                cel.fill.fore_color.rgb = DESTAQUE
            else:
                cel.fill.fore_color.rgb = BRANCO if i % 2 else CINZA_CLARO
            p = cel.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 or grupo else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = ("    " if recuo else "") + texto
            r.font.name, r.font.size = FONTE, Pt(tam)
            r.font.bold = (i == 0) or negrito
            r.font.italic = italico and i > 0
            r.font.color.rgb = BRANCO if i == 0 else PRETO
        if grupo:
            tb.cell(i, 0).merge(tb.cell(i, ncol - 1))
    return tb


def legenda(slide, texto, y=None):
    caixa_texto(slide, MARGEM, y or (ALT - Inches(0.95)), LARG_CONTEUDO, Inches(0.4), texto, tam=12, cor=CINZA, italico=True, alinh=PP_ALIGN.CENTER)


def bloco(slide, x, y, w, h, titulo, cor, texto, tam=14):
    cab = slide.shapes.add_shape(1, x, y, w, Inches(0.5)); cab.fill.solid(); cab.fill.fore_color.rgb = cor; cab.line.fill.background()
    cab.text_frame.text = titulo
    pr = cab.text_frame.paragraphs[0]; pr.alignment = PP_ALIGN.CENTER
    pr.runs[0].font.name, pr.runs[0].font.size, pr.runs[0].font.bold, pr.runs[0].font.color.rgb = FONTE, Pt(16), True, BRANCO
    corpo = slide.shapes.add_shape(1, x, y + Inches(0.5), w, h - Inches(0.5)); corpo.fill.solid(); corpo.fill.fore_color.rgb = CINZA_CLARO; corpo.line.fill.background()
    caixa_texto(slide, x + Inches(0.1), y + Inches(0.6), w - Inches(0.2), h - Inches(0.7), texto, tam=tam)


# ================================================================== 1: titulo
s = prs.slides.add_slide(BRANCO_LAYOUT)
faixa = s.shapes.add_shape(1, 0, 0, LARG, Inches(4.2)); faixa.fill.solid(); faixa.fill.fore_color.rgb = AZUL; faixa.line.fill.background()
caixa_texto(s, MARGEM, Inches(1.1), LARG_CONTEUDO, Inches(2.0), "Previsão da severidade de ocorrências aeronáuticas na aviação comercial brasileira (RBAC 121 e 135)", tam=34, cor=BRANCO, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
caixa_texto(s, MARGEM, Inches(3.1), LARG_CONTEUDO, Inches(0.6), "Relatório final · RP2 · 2026", tam=18, cor=RGBColor(0xC9, 0xD6, 0xEE))
caixa_texto(s, MARGEM, Inches(4.6), LARG_CONTEUDO, Inches(0.6), "Eric I. W. Chou · João Fontanezi · Marcos R. da S. Santos", tam=22, cor=PRETO, negrito=True)
caixa_texto(s, MARGEM, Inches(5.2), LARG_CONTEUDO, Inches(0.5), "Escola de Artes, Ciências e Humanidades – Universidade de São Paulo (EACH-USP)", tam=16, cor=CINZA)
caixa_texto(s, MARGEM, Inches(6.2), LARG_CONTEUDO, Inches(0.5), "Dados abertos do CENIPA / Painel SIPAER · Regressão Logística e Random Forest · github.com/Fontanezi/RP2_ARTIGO_2026", tam=13, cor=CINZA)
notas(s, "Todos (abertura)", "0:15", "Título e integrantes. Uma frase: 'com o que se sabe no momento da notificação, dá para antecipar se uma ocorrência é incidente, incidente grave ou acidente?'")

# ================================================================== PARTE 1
s = novo_slide("Contexto e problema", "Parte 1 · Introdução", 2)
topicos(s, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, ALT_CONTEUDO, [
    "Transporte aéreo público no Brasil: **RBAC 121** (aviões com mais de 19 assentos) e **RBAC 135** (até 19 assentos ou helicópteros: táxi aéreo e operação não regular)",
    "O **CENIPA** investiga e publica todas as ocorrências no Painel SIPAER e em dados.gov.br",
    "A análise usual desses registros é **retrospectiva e caso a caso**: explica o que aconteceu, não antecipa",
    "A literatura nacional traz séries temporais agregadas e estudos descritivos, mas **não classifica a severidade de uma ocorrência individual**",
    "Pergunta: com os atributos conhecidos na notificação, é possível estimar a probabilidade de a ocorrência ser **incidente, incidente grave ou acidente**?",
], tam=19, espaco=10)
notas(s, "Integrante 1", "0:25", "Setor (RBAC 121 e 135), fonte (CENIPA) e lacuna. Terminar na pergunta do trabalho.")

s = novo_slide("Objetivo e as três classes do CENIPA (NSCA 3-13)", "Parte 1 · Introdução", 3)
topicos(s, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, Inches(1.1), [
    "**Objetivo:** desenvolver e comparar Regressão Logística multinomial e Random Forest para classificar a severidade de ocorrências da aviação comercial em três classes, usando apenas atributos disponíveis antes da investigação",
], tam=18)
y0 = TOPO_CONTEUDO + Inches(1.3)
for k, (titulo, cor, texto) in enumerate([
    ("Acidente", VERMELHO, "Lesão grave ou morte, dano ou falha estrutural que afete a resistência ou o desempenho da aeronave (ou exija grande reparo), aeronave desaparecida ou inacessível. Exceções: danos localizados (um motor, hélices, pneus, colisão com fauna)."),
    ("Incidente grave", LARANJA, "Incidente cujas circunstâncias indicam elevado potencial de risco de acidente. A diferença para o acidente está apenas nas consequências: é o quase-acidente."),
    ("Incidente", AZUL, "Toda ocorrência que não se caracteriza como acidente, mas que afeta ou pode afetar a segurança da operação."),
]):
    x = MARGEM + k * (LARG_CONTEUDO // 3)
    bloco(s, x, y0, LARG_CONTEUDO // 3 - Inches(0.2), Inches(3.1), titulo, cor, texto)
caixa_texto(s, MARGEM, y0 + Inches(3.3), LARG_CONTEUDO, Inches(0.5), "Três classes, e não duas: as definições separam o incidente grave pelo potencial de risco, não pela consequência, e os dados mostram perfis diferentes (slide 9).", tam=14, cor=CINZA, italico=True)
notas(s, "Integrante 1", "0:35", "Objetivo e as três definições, com ênfase em 'a diferença entre incidente grave e acidente está apenas nas consequências'. Justificar manter as três classes separadas.")

s = novo_slide("Trabalhos correlatos", "Parte 1 · Introdução", 4)
corr = [[("Referência", 1, 0, 0), ("Modelo", 1, 0, 0), ("Dados utilizados", 1, 0, 0), ("Contribuição deste trabalho", 1, 0, 0)],
        [("Reinert (2023)", 0, 0, 0), ("Ajuste de curvas e ARIMA sobre séries anuais agregadas", 0, 0, 0), ("SIPAER 2012–2022, aviões Ipanema e Sêneca", 0, 0, 0), ("Classifica cada ocorrência individual em três classes com atributos pré-investigação; cobre todo o transporte aéreo comercial", 0, 0, 0)],
        [("Soares (2012)", 0, 0, 0), ("Análise sinótica qualitativa de casos", 0, 0, 0), ("Acidentes dos anos 1990 associados à meteorologia", 0, 0, 0), ("Quantifica o risco com modelos discriminativos validados por métricas; meteorologia é uma das categorias de tipo", 0, 0, 0)],
        [("Barros e Balero (2012)", 0, 0, 0), ("Pesquisa bibliográfica e estatística descritiva", 0, 0, 0), ("Estações meteorológicas do Centro-Oeste e estatísticas do SIPAER", 0, 0, 0), ("Abrangência nacional com a região como preditor; validação cruzada repetida, teste e validação temporal, com baselines explícitos", 0, 0, 0)]]
tabela(s, corr, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, tam=14, larguras=[1.6, 2.6, 2.8, 4.2], alt_linha=Inches(0.95))
caixa_texto(s, MARGEM, TOPO_CONTEUDO + Inches(4.1), LARG_CONTEUDO, Inches(0.8), "Em comum: nenhum estima a severidade de uma ocorrência individual a partir de atributos operacionais, geográficos e técnicos. É essa a lacuna que este trabalho ataca.", tam=15, cor=AZUL, negrito=True)
notas(s, "Integrante 1", "0:30", "Um trabalho de séries temporais e dois descritivos sobre meteorologia. Destacar a última coluna. Não ler a tabela inteira.")

# ================================================================== PARTE 2
s = novo_slide("Dados: construção da base e escopo", "Parte 2 · Metodologia", 5)
etapas = [("14.626", "ocorrências nos arquivos do CENIPA"), ("9.471", "entre 10/05/2016 e 09/05/2026"), ("5.038", "com aeronave em operação regular, não regular ou táxi aéreo")]
for k, (num, desc) in enumerate(etapas):
    x = MARGEM + k * Inches(4.1)
    bl = s.shapes.add_shape(1, x, TOPO_CONTEUDO, Inches(3.7), Inches(1.5)); bl.fill.solid(); bl.fill.fore_color.rgb = AZUL if k == 2 else CINZA_CLARO; bl.line.fill.background()
    caixa_texto(s, x, TOPO_CONTEUDO + Inches(0.1), Inches(3.7), Inches(0.7), num, tam=32, cor=BRANCO if k == 2 else AZUL, negrito=True, alinh=PP_ALIGN.CENTER)
    caixa_texto(s, x + Inches(0.1), TOPO_CONTEUDO + Inches(0.8), Inches(3.5), Inches(0.65), desc, tam=13, cor=BRANCO if k == 2 else CINZA, alinh=PP_ALIGN.CENTER)
    if k < 2:
        caixa_texto(s, x + Inches(3.7), TOPO_CONTEUDO + Inches(0.4), Inches(0.4), Inches(0.7), "→", tam=28, cor=CINZA, alinh=PP_ALIGN.CENTER)
topicos(s, MARGEM, TOPO_CONTEUDO + Inches(1.8), LARG_CONTEUDO, Inches(3.6), [
    "Cruzamento das tabelas de Ocorrências, Tipos, Aeronaves, Fatores Contribuintes e Recomendações pelo código da ocorrência; **uma linha por ocorrência**",
    "Classes: **4.845 incidentes (96,2%) · 102 incidentes graves (2,0%) · 91 acidentes (1,8%)**; 17 acidentes com vítimas fatais, 122 mortes (62 em Vinhedo, 2024)",
    "Escopo = todo o transporte aéreo comercial (RBAC 121 + 135): mesma família regulatória; a operação regular sozinha tem só **12 acidentes**; o tipo de operação entra como preditor",
    "Base publicada com dicionário de dados: 5.038 linhas × 44 colunas (BASE_SIPAER.xlsx)",
], tam=17, espaco=8)
notas(s, "Integrante 2", "0:30", "Funil de construção (três números), desbalanceamento (menos de 2% de acidentes) e os três motivos do escopo RBAC 121 + 135.")

s = novo_slide("Vazamento de informação: o que não pode ser preditor", "Parte 2 · Metodologia", 6)
pos = ler_tabela_tex("tab_posevento.tex")
tabela(s, pos, MARGEM, TOPO_CONTEUDO, Inches(7.4), tam=12, larguras=[3.9, 1.4, 1.5, 1.6], alt_linha=Inches(0.45))
caixa_texto(s, MARGEM, TOPO_CONTEUDO + Inches(2.85), Inches(7.4), Inches(0.6), "Atributos produzidos pela investigação, por classe. A investigação completa é feita quase só em acidentes e incidentes graves: esses atributos entregariam a resposta ao modelo.", tam=12, cor=CINZA, italico=True)
topicos(s, MARGEM + Inches(7.7), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(7.7), ALT_CONTEUDO, [
    "**Fora dos preditores:** nível de dano, fatalidades, situação da investigação, relatório final, fatores contribuintes, recomendações; e o ano (mudança de cobertura em 2023)",
    "**18 preditores disponíveis na notificação:** tipo de operação, equipamento, tipo e quantidade de motores, fabricante, peso máximo de decolagem, assentos, idade, fase do voo, região, local, categoria do tipo, saída de pista, nº de tipos, aeronaves envolvidas, mês, dia da semana, hora",
    "Aeródromo (203 níveis) e tipo detalhado (65) entram agrupados; a codificação pelo alvo foi testada como experimento",
], tam=15, espaco=10)
notas(s, "Integrante 2", "0:30", "Achado metodológico central: relatório, fatores e recomendações existem em 61–68% dos acidentes, 44–45% dos incidentes graves e 0,1% dos incidentes. Listar rapidamente os 18 preditores.")

s = novo_slide("Pipeline, protocolo e experimentos complementares", "Parte 2 · Metodologia", 7)
caixas = ["Imputação\n(mediana / moda)", "SMOTENC\n(só no treino)", "Padronização\n+ one-hot", "Regressão Logística\nmultinomial  ·  Random Forest"]
for k, txt in enumerate(caixas):
    x = MARGEM + k * Inches(3.1)
    w = Inches(2.7) if k < 3 else Inches(3.1)
    cx = s.shapes.add_shape(1, x, TOPO_CONTEUDO, w, Inches(0.9)); cx.fill.solid(); cx.fill.fore_color.rgb = AZUL if k == 3 else CINZA_CLARO; cx.line.fill.background()
    caixa_texto(s, x, TOPO_CONTEUDO, w, Inches(0.9), txt, tam=14, cor=BRANCO if k == 3 else PRETO, negrito=k == 3, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)
    if k < 3:
        caixa_texto(s, x + w, TOPO_CONTEUDO + Inches(0.1), Inches(0.4), Inches(0.7), "→", tam=26, cor=CINZA, alinh=PP_ALIGN.CENTER)
caixa_texto(s, MARGEM, TOPO_CONTEUDO + Inches(0.95), LARG_CONTEUDO, Inches(0.35), "Tudo ajustado apenas nas dobras de treino, nunca nos dados avaliados.", tam=12, cor=CINZA, italico=True, alinh=PP_ALIGN.CENTER)
topicos(s, MARGEM, TOPO_CONTEUDO + Inches(1.4), Inches(6.0), Inches(4.0), [
    "**Partição 80/20 estratificada:** treino 4.030 (3.876 / 81 / 73), teste 1.008 (969 / 21 / 18)",
    "**Hiperparâmetros:** busca em grade com validação cruzada estratificada (k = 5) no treino, maximizando F1-macro",
    "**Balanceamento, 4 estratégias:** nenhum · pesos de classe · SMOTENC · SMOTENC + pesos",
    "**Métricas:** sensibilidade, precisão e F1 por classe; F1-macro; AUC e precisão média (um-contra-resto); Brier",
    "**Baselines:** classe majoritária; regra 'táxi aéreo ou não regular ⇒ acidente'",
], tam=14, espaco=7)
topicos(s, MARGEM + Inches(6.3), TOPO_CONTEUDO + Inches(1.4), LARG_CONTEUDO - Inches(6.3), Inches(4.0), [
    "**Experimentos complementares**",
    ("**Ablação de atributos:** sem calendário; sem descrição do evento; com codificação pelo alvo de aeródromo e tipo detalhado", 1),
    ("**Validação cruzada repetida** 10 × 5: média ± desvio-padrão e diferença pareada RF − RL", 1),
    ("**Validação temporal:** treino 2016–2022, teste 2023–2026", 1),
    ("**Limiar de decisão** para sensibilidade ≥ 80% em acidentes", 1),
    ("**Calibração:** score de Brier e diagrama de confiabilidade", 1),
], tam=14, espaco=6)
notas(s, "Integrante 2", "0:40", "Percorrer o fluxo. Explicar por que o número principal é a validação cruzada (18 acidentes no teste). Depois os cinco experimentos complementares: cada um responde a uma pergunta deixada pelo relatório parcial.")

# ================================================================== PARTE 3: RESULTADOS (>= 4 min)
s = novo_slide("Resultados · Operação e aeronave separam as classes", "Parte 3 · Resultados", 8)
op = ler_tabela_tex("tab_operacao.tex")
tabela(s, op, MARGEM, TOPO_CONTEUDO, Inches(7.6), tam=12, larguras=[2.6, 1.3, 1.2, 1.2, 1.4, 1.4], alt_linha=Inches(0.285))
topicos(s, MARGEM + Inches(7.9), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(7.9), ALT_CONTEUDO, [
    "Operação regular: 85% das ocorrências, **0,28%** de acidentes",
    "Táxi aéreo: 13% das ocorrências, **72 dos 91 acidentes** (11,1%)",
    "Pistão 22,6% vs. jato 0,32%; helicóptero 7,5% vs. avião 1,3%",
    "Idade mediana: 11 anos (incidente), 16 (inc. grave), **32 (acidente)**; peso: 70,5 t, 7,3 t, **2,2 t**",
    "Nenhum acidente em solo; incidentes graves concentram-se na aproximação e pouso",
    "Táxi aéreo, pistão, helicóptero e aeronave leve e antiga são **o mesmo subgrupo** (colinearidade)",
], tam=14, espaco=7)
notas(s, "Integrante 3", "0:25", "Tipo de operação é o que mais separa as classes. Avisar da colinearidade: vai aparecer na interpretação da Regressão Logística.")

s = novo_slide("Resultados · Tipos de ocorrência: perfis distintos", "Parte 3 · Resultados", 9)
tip = ler_tabela_tex("tab_tipos.tex")
tabela(s, tip, MARGEM, TOPO_CONTEUDO, Inches(8.4), tam=13, larguras=[3.2, 1.3, 0.9, 1.3, 1.3, 1.2, 1.1], alt_linha=Inches(0.4))
topicos(s, MARGEM + Inches(8.7), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(8.7), ALT_CONTEUDO, [
    "Falha técnica + fauna = 85% das ocorrências; **colisão com fauna quase nunca é grave** (0 e 1 em 1.767)",
    "**Acidente:** falha do motor em voo (14 acidentes, 0 inc. graves em 122) e perda de controle (60%)",
    "**Incidente grave:** tráfego aéreo (21 inc. graves, **0 acidentes**) e pista, pouso e solo",
    "Perfis diferentes ⇒ três classes separadas",
], tam=14, espaco=9)
notas(s, "Integrante 3", "0:25", "Tráfego aéreo gera 21 incidentes graves e nenhum acidente; falha do motor em voo, 14 acidentes e nenhum incidente grave. É o argumento empírico das três classes.")

s = novo_slide("Resultados · A cobertura das notificações mudou em 2023", "Parte 3 · Resultados", 10)
figura(s, "fig1_anual.png", MARGEM, TOPO_CONTEUDO, Inches(7.4), ALT_CONTEUDO - Inches(0.1))
topicos(s, MARGEM + Inches(7.7), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(7.7), ALT_CONTEUDO, [
    "2016–2022: 1.097 ocorrências; 2023–2026: **3.941** (pico de 1.582 em 2024), quase todas na operação regular e nas colisões com fauna",
    "Acidentes e incidentes graves por ano **não crescem**: a proporção de acidentes cai de 5,9% para 0,7%. Mais notificação, não menos risco",
    "Por isso o **ano não é preditor**, e a mudança motiva a **validação temporal** (slide 15)",
    "Geografia: Norte tem 10,5% das ocorrências e **44 dos 91 acidentes** (táxi aéreo); 49 acidentes fora de aeródromo; 57 dos 102 incidentes graves em aeródromos SB",
], tam=14, espaco=9)
legenda(s, "Figura 1. Ocorrências por ano e tipo de operação (acima) e incidentes graves e acidentes por ano (abaixo); 2016 e 2026 parciais", y=ALT - Inches(0.85))
notas(s, "Integrante 3", "0:25", "Descontinuidade em 2023 no painel de cima; acidentes estáveis no de baixo. Conclusão: mudança de cobertura, não de risco; isso justifica a validação temporal.")

s = novo_slide("Resultados · Seleção: balanceamento e atributos", "Parte 3 · Modelos", 11)
sel = ler_tabela_tex("tab_selecao.tex")
tabela(s, sel, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, tam=11, larguras=[2.0, 2.9, 1.7, 1.0, 1.0, 1.0, 1.0, 1.0], alt_linha=Inches(0.255), destaque=(2, 7, 9, 13))
caixa_texto(s, MARGEM, TOPO_CONTEUDO + Inches(4.45), LARG_CONTEUDO, Inches(0.9),
            "Balanceamento: na RL, pesos de classe trocam precisão por sensibilidade (F1-macro cai); no RF, pesos e SMOTENC empatam em 0,62 e o SMOTENC varia menos entre dobras.  "
            "Atributos: sem calendário o F1-macro cai 0,02 (conjunto completo mantido); sem a descrição do evento cai para 0,50–0,55, mas ainda bate os baselines; a codificação pelo alvo não traz ganho.",
            tam=13, cor=PRETO)
legenda(s, "Tabela 5. F1-macro na validação cruzada (média ± dp entre dobras) e métricas out-of-fold; † configuração selecionada; AC = acidente, IG = incidente grave", y=ALT - Inches(0.78))
notas(s, "Integrante 1 ou 2", "0:40", "Ler só as linhas destacadas. Duas mensagens: (1) o SMOTENC ajuda o Random Forest sem destruir a precisão; (2) a descrição do evento é o grupo de atributos mais informativo, e o calendário e a codificação pelo alvo não mudam nada de forma mensurável.")

s = novo_slide("Resultados · Desempenho com variabilidade (CV repetida 10 × 5)", "Parte 3 · Modelos", 12)
met = ler_tabela_tex("tab_metricas.tex")
tabela(s, met, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, tam=12, larguras=[3.2, 1.3, 1.3, 1.3, 1.3, 1.4, 1.4], alt_linha=Inches(0.3), destaque=(5, 10))
topicos(s, MARGEM, TOPO_CONTEUDO + Inches(3.5), LARG_CONTEUDO, Inches(1.85), [
    "**Random Forest + SMOTENC:** F1-macro **0,598 ± 0,010**; detecta 56,8 ± 2,9% dos acidentes e 37,8 ± 3,5% dos inc. graves. **Regressão Logística:** 0,568 ± 0,016; 39,3% e 21,5%, com precisão maior (50,1% vs. 43,7%)",
    "**Diferença pareada:** RF vence em F1-macro em 9 de 10 repetições (+0,030 ± 0,019) e em sensibilidade em todas; **AUC indistinguível** (0,947 vs. 0,943): os dois ordenam igual, o SMOTENC muda o ponto de operação",
    "**Leitura honesta:** o 0,624 do relatório parcial era o máximo das 10 repetições (seleção e avaliação na mesma partição); a média é 0,598. Teste confirma a ordem: RF 11/18 acidentes, RL 10/18",
], tam=13, espaco=6)
legenda(s, "Tabela 6. Validação cruzada repetida no treino (73 acidentes, 81 inc. graves) e teste (18 e 21); sensibilidade e precisão em %", y=ALT - Inches(0.78))
notas(s, "Integrante 1 ou 2", "0:45", "Slide central. Números com ± da CV repetida. Explicar: RF e RL têm a mesma AUC; o RF ganha F1 porque o SMOTENC troca precisão por sensibilidade. Declarar o otimismo do parcial (0,624 → 0,598) antes que perguntem.")

s = novo_slide("Resultados · Onde os modelos erram", "Parte 3 · Modelos", 13)
figura(s, "fig3_confusao.png", MARGEM, TOPO_CONTEUDO, Inches(8.6), Inches(3.6))
topicos(s, MARGEM + Inches(8.9), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(8.9), ALT_CONTEUDO, [
    "**Random Forest:** acidentes perdidos vão para 'incidente' (23), raramente para 'inc. grave' (6); custo: 107 incidentes (2,8%) previstos como graves ou acidentes",
    "**Regressão Logística** quase não prevê 'incidente grave' (35 predições em 4.030): é a classe de sinal mais fraco",
    "Incidente grave é difícil por definição: difere do acidente só nas consequências, e os atributos descrevem as condições do evento, não o desfecho",
], tam=14, espaco=9)
legenda(s, "Matrizes de confusão das predições out-of-fold na partição principal (contagem e percentual da linha)", y=TOPO_CONTEUDO + Inches(3.7))
notas(s, "Integrante 1 ou 2", "0:25", "Diagonal do Random Forest (97%, 43%, 60%) e o custo em falsos alarmes. Coluna 'inc. grave' quase vazia na Regressão Logística.")

s = novo_slide("Resultados · Independente do limiar: ROC e precisão-sensibilidade", "Parte 3 · Modelos", 14)
figura(s, "fig4_curvas.png", MARGEM, TOPO_CONTEUDO, Inches(8.6), Inches(3.6))
topicos(s, MARGEM + Inches(8.9), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(8.9), ALT_CONTEUDO, [
    "**Acidente:** AUC 0,94 nos dois modelos; precisão média 0,41–0,44 (prevalência 1,8%)",
    "**Incidente grave:** AUC 0,85–0,88, precisão média 0,23–0,33",
    "Na sensibilidade da regra (86%), os modelos chegam a 14–17% de precisão contra 10,7%: ganho modesto nesse ponto",
    "A vantagem está na faixa intermediária: RF com 45–70% de precisão para 40–65% de sensibilidade",
], tam=14, espaco=9)
legenda(s, "Figura 2. Curvas ROC (a) e precisão-sensibilidade (b) um-contra-resto, predições out-of-fold; o losango marca a regra pelo tipo de operação", y=TOPO_CONTEUDO + Inches(3.7))
notas(s, "Integrante 1 ou 2", "0:25", "Desempenho sem depender do limiar. Ponto honesto: na sensibilidade da regra o ganho é pequeno; o modelo vale mais na faixa intermediária.")

s = novo_slide("Resultados · Validação temporal: treino até 2022, teste 2023–2026", "Parte 3 · Modelos", 15)
tmp = ler_tabela_tex("tab_temporal.tex")
tabela(s, tmp, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, tam=13, larguras=[3.2, 1.3, 1.3, 1.3, 1.3, 1.4, 1.4], alt_linha=Inches(0.4))
topicos(s, MARGEM, TOPO_CONTEUDO + Inches(2.4), LARG_CONTEUDO, Inches(2.9), [
    "**A ordenação se transfere:** AUC para acidente 0,974 (RL) e 0,954 (RF), iguais ou melhores que na partição aleatória",
    "**O ponto de operação não:** precisão para acidentes cai para 30,8% e 20,6%; o RF prevê 206 incidentes graves (5,2% do teste) com 11,7% de precisão; a regra cai para 5,8%",
    "**Causa:** a taxa-base de acidentes caiu de 5,9% para 0,7% com a ampliação das notificações; um modelo treinado com a prevalência antiga (e inflada pelo SMOTENC) emite alertas demais na nova",
    "**Implicação:** para uso prospectivo, limiar e probabilidades precisam ser recalibrados à prevalência corrente, e o modelo retreinado periodicamente",
], tam=14, espaco=8)
legenda(s, "Tabela 7. Modelos treinados em 2016–2022 (65 acidentes) e avaliados em 2023–2026 (26 acidentes, 38 inc. graves); sensibilidade e precisão em %", y=ALT - Inches(0.78))
notas(s, "Integrante 1 ou 2", "0:35", "O experimento mais importante para o uso real. Separar duas coisas: a ordenação (AUC) transferiu; o ponto de operação quebrou por mudança de taxa-base. Implicação: recalibrar e retreinar.")

s = novo_slide("Resultados · O que os modelos aprenderam", "Parte 3 · Modelos", 16)
figura(s, "fig5_importancias.png", MARGEM, TOPO_CONTEUDO, Inches(8.4), Inches(3.9))
topicos(s, MARGEM + Inches(8.7), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(8.7), ALT_CONTEUDO, [
    "**Random Forest:** tipo de operação, peso, fase do voo, local e categoria do tipo lideram; mês, dia, hora e assentos têm importância ~0, coerente com a ablação",
    "**Regressão Logística** (chance de acidente vs. incidente): perda de controle em voo (~10³), Robinson (58×), turboeixo (48×), fora de aeródromo (26×), táxi aéreo e não regular (15×), Norte (13×)",
    "Cessna, Sikorsky e Embraer com chance **menor** que Airbus apesar de taxas maiores: **colinearidade** com operação, motor e peso; ler em conjunto",
], tam=13, espaco=9)
legenda(s, "Figura 3. (a) Contrastes de log da razão de chances na Regressão Logística; (b) importância por permutação no Random Forest (queda da AUC no teste)", y=TOPO_CONTEUDO + Inches(4.0))
notas(s, "Integrante 1 ou 2", "0:30", "O Random Forest confirma a análise exploratória. Na Regressão Logística, explicar o sinal 'errado' de Cessna/Sikorsky/Embraer pela colinearidade.")

s = novo_slide("Resultados · Limiar de decisão e calibração", "Parte 3 · Modelos", 17)
figura(s, "fig6_limiar_calibracao.png", MARGEM, TOPO_CONTEUDO, Inches(8.2), Inches(3.4))
topicos(s, MARGEM + Inches(8.5), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(8.5), Inches(3.5), [
    "**Limiar para 80% de sensibilidade:** RL τ = 0,029, RF τ = 0,130; ambos detectam 59 dos 73 acidentes",
    "**Alertas:** RL 263 (6,5% das ocorrências, precisão 22,4%); RF 342 (17,3%); regra 588 (10,7%). A RL **reduz os alertas à metade**, e 36% deles são ocorrências graves (regra: 19%)",
    "Teste, mesmos limiares: RL 17/18 acidentes com 73 alertas; RF 15/18 com 78",
], tam=13, espaco=7)
topicos(s, MARGEM, TOPO_CONTEUDO + Inches(3.55), LARG_CONTEUDO, Inches(1.3), [
    "**Calibração:** RL tem Brier 0,0144 (19% abaixo do previsor constante, 0,0178) e probabilidade média = prevalência (1,8%). RF tem Brier 0,0162 e probabilidade média 4,4%, **2,4× a prevalência**: superestima o risco 2 a 4× em todas as faixas, efeito direto do SMOTENC (treino com 1/3 de acidentes). Como score de risco, a RL serve sem ajuste; o RF ordena igual, mas precisa de recalibração",
], tam=13, espaco=6)
legenda(s, "Figura 4. (a) Diagrama de confiabilidade de P(acidente), out-of-fold, escalas logarítmicas; (b) sensibilidade e precisão em função do limiar; RL = Regressão Logística, RF = Random Forest", y=ALT - Inches(0.78))
notas(s, "Integrante 1 ou 2", "0:40", "Dois resultados operacionais. Limiar: para o mesmo nível de detecção da regra, a Regressão Logística emite metade dos alertas. Calibração: as probabilidades da RL são utilizáveis; as do RF estão infladas 2–4× pelo SMOTENC e precisam de recalibração.")

# ================================================================== CONCLUSAO
s = novo_slide("Discussão", "Parte 4 · Discussão e conclusão", 18)
topicos(s, MARGEM, TOPO_CONTEUDO, Inches(6.1), ALT_CONTEUDO, [
    "**Resposta com ressalva:** os atributos da notificação distinguem acidentes com AUC 0,94–0,95 e concentram os acidentes em 6,5% das ocorrências; mas, nesse ponto, 4 em 5 alertas não são acidentes, e o incidente grave é a classe difícil (AUC 0,85–0,89)",
    "**RF vs. RL é mais sutil que o F1-macro:** mesma ordenação; o RF ganha pelo ponto de operação do SMOTENC e paga com probabilidades infladas. Decisão fixa: RF. Score de risco calibrado e interpretável: RL",
    "**Tensão na descrição do evento:** é o grupo mais informativo, mas descreve em parte o desfecho; sem ela o modelo ainda bate os baselines (F1-macro 0,545, AUC 0,92)",
], tam=15, espaco=10)
topicos(s, MARGEM + Inches(6.4), TOPO_CONTEUDO, LARG_CONTEUDO - Inches(6.4), ALT_CONTEUDO, [
    "**Validação temporal é a limitação mais importante:** a prevalência caiu 8× em 2023; limiares e probabilidades fixados antes emitem alertas em excesso. Um sistema real precisa ser recalibrado e retreinado",
    "**Limitações:** 91 acidentes (desvio-padrão de 3 pontos na sensibilidade; razões de chance de categorias raras instáveis); 30% dos acidentes e inc. graves com investigação ativa; modelos discriminativos, sem leitura causal (o Norte eleva a chance porque concentra táxi aéreo em aeronaves leves e antigas)",
    "**Seleção e avaliação na mesma partição inflam o resultado:** o 0,624 do parcial virou 0,598 ± 0,010 na CV repetida",
], tam=15, espaco=10)
notas(s, "Integrante 3", "0:35", "Responder à pergunta com a ressalva. Explicar a nuance RF vs. RL. Declarar as limitações antes que perguntem, inclusive o otimismo do parcial.")

s = novo_slide("Conclusão", "Parte 4 · Discussão e conclusão", 19)
topicos(s, MARGEM, TOPO_CONTEUDO, LARG_CONTEUDO, ALT_CONTEUDO, [
    "Base de **5.038 ocorrências** do transporte aéreo comercial (2016–2026), alvo em **três classes do CENIPA**, preditores só da notificação (atributos da investigação excluídos)",
    "**Random Forest + SMOTENC:** F1-macro 0,60 ± 0,01, 57% dos acidentes e 38% dos incidentes graves fora da amostra. **Regressão Logística:** 0,57 ± 0,02, probabilidades calibradas e mesma AUC (0,94–0,95)",
    "Com limiar para 80% de sensibilidade, a RL concentra os acidentes em **6,5% das ocorrências**, metade dos alertas da regra pelo tipo de operação",
    "A validação temporal mostra que a ordenação se transfere, mas **limiar e probabilidades devem ser recalibrados** à prevalência corrente",
    "**Continuidade:** recalibrar o RF, atualizar com novos dados abertos do CENIPA e incorporar atributos de exposição (horas voadas, movimentos por operador) para estimar taxas de risco, não só a severidade condicional à ocorrência",
    "Código, base, figuras e artigo: **github.com/Fontanezi/RP2_ARTIGO_2026**",
], tam=16, espaco=10)
notas(s, "Integrante 3 (fechamento)", "0:25", "Cinco frases: o que foi feito, o melhor resultado com o ±, o ganho operacional do limiar, a lição da validação temporal e a continuidade. Agradecer.")

assert len(prs.slides) == TOTAL_SLIDES, len(prs.slides)
prs.save(SAIDA)
print("ok", SAIDA, len(prs.slides), "slides")
