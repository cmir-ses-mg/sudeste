# -*- coding: utf-8 -*-
"""
build_boletim_sudeste.py — Boletim de Acompanhamento do Plano Sudeste (Fonte 95)
Lê a planilha de Monitoramento (abas 'Consolidado' e 'Plano Sudeste') e produz
um boletim.html autocontido, pronto para impressão em A4.
Uso: python build_boletim_sudeste.py [pasta_ou_arquivo.xlsx] [saida.html]
"""
import sys, os, glob, datetime, unicodedata
from collections import defaultdict, Counter
import openpyxl


def achar_planilha(alvo):
    if os.path.isfile(alvo):
        return alvo
    if os.path.isdir(alvo):
        c = [f for f in glob.glob(os.path.join(alvo, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
        if not c:
            sys.exit(f"ERRO: nenhum .xlsx em '{alvo}'.")
        novo = max(c, key=os.path.getmtime)
        if len(c) > 1:
            print(f"[aviso] {len(c)} planilhas — usando {os.path.basename(novo)}")
        return novo
    sys.exit(f"ERRO: caminho não encontrado: '{alvo}'")


XLSX = achar_planilha(sys.argv[1] if len(sys.argv) > 1 else "dados")
OUT = sys.argv[2] if len(sys.argv) > 2 else "boletim.html"
# 3º argumento (opcional): index.html do dashboard, onde o botão será injetado
DASH = sys.argv[3] if len(sys.argv) > 3 else None
print(f"Planilha: {os.path.basename(XLSX)}")


def injetar_botao(caminho_index, destino="boletim.html"):
    """Insere no dashboard um botão flutuante que abre o boletim.
    É idempotente: rodar de novo não duplica o botão."""
    if not caminho_index or not os.path.isfile(caminho_index):
        print(f"[aviso] dashboard não encontrado em '{caminho_index}' — botão não injetado")
        return
    with open(caminho_index, encoding="utf-8") as f:
        pagina = f.read()
    if "id='btn-boletim'" in pagina or 'id="btn-boletim"' in pagina:
        print("[info] botão do boletim já existe no dashboard")
        return
    botao = """
<style>
  #btn-boletim {{ position:fixed; right:22px; bottom:22px; z-index:999; display:flex; align-items:center;
                 gap:8px; text-decoration:none; font-family:'Segoe UI',sans-serif; font-size:13px;
                 font-weight:600; color:#fff; background:#1B4F72; border-radius:26px;
                 padding:11px 18px; box-shadow:0 4px 14px rgba(14,48,73,.32); transition:background .15s; }}
  #btn-boletim:hover {{ background:#0E3049; }}
  #btn-boletim svg {{ width:16px; height:16px; fill:none; stroke:#fff; stroke-width:2;
                     stroke-linecap:round; stroke-linejoin:round; }}
  @media print {{ #btn-boletim {{ display:none; }} }}
</style>
<a id='btn-boletim' href='{destino}' target='_blank' rel='noopener' title='Abrir o boletim em PDF'>
  <svg viewBox='0 0 24 24'><path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/>
    <path d='M14 2v6h6'/><path d='M8 13h8'/><path d='M8 17h5'/></svg>
  Boletim de Acompanhamento
</a>
""".format(destino=destino)
    if "</body>" in pagina:
        pagina = pagina.replace("</body>", botao + "</body>", 1)
    else:
        pagina += botao
    with open(caminho_index, "w", encoding="utf-8") as f:
        f.write(pagina)
    print(f"[ok] botão do boletim injetado em {caminho_index}")


def brl(v, dec=2):
    s = f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def mi(v):
    return ("R$ %.1f mi" % (v / 1_000_000)).replace(".", ",")


def pct(x, dec=1):
    return f"{x:.{dec}f}".replace(".", ",") + "%"


def w(x):
    return f"{max(0, min(100, x)):.1f}"


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())


MINUSC = {"e", "de", "da", "do", "das", "dos", "para", "em", "no", "na", "nos",
          "nas", "a", "o", "as", "os", "com", "por", "ao", "aos", "à", "às"}
SIGLAS = {"UBS", "HPS", "APS", "SEI", "UPA", "CAPS", "SAMU", "NA", "MPMG", "URS",
          "SUS", "HRJP", "CEO", "UTI"}

def titulo(s):
    palavras = " ".join(str(s or "").split()).split(" ")
    out = []
    for i, p in enumerate(palavras):
        limpo = p.strip(".,;:()-").upper()
        if limpo in SIGLAS:
            out.append(p.upper())
        elif i > 0 and p.lower() in MINUSC:
            out.append(p.lower())
        else:
            out.append(p[:1].upper() + p[1:].lower() if p else p)
    return " ".join(out)


# ── LEITURA ─────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, data_only=True)
for aba in ("Consolidado", "Plano Sudeste"):
    if aba not in wb.sheetnames:
        sys.exit(f"ERRO: falta a aba '{aba}'. Abas: {', '.join(wb.sheetnames)}")

ws = wb["Consolidado"]
HC = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
IX = {h: i + 1 for i, h in enumerate(HC) if h}
COL_STATUS = next((h for h in HC if h.startswith("STATUS")), None)
if not COL_STATUS:
    sys.exit("ERRO: não achei a coluna STATUS na aba Consolidado.")
DATA_STATUS = COL_STATUS.replace("STATUS", "").strip(" ()")


def g(r, n):
    return str(ws.cell(r, IX[n]).value or "").strip() if n in IX else ""


def gv(r, n):
    return ws.cell(r, IX[n]).value if n in IX else None


# indicações válidas (ACORDO SUDESTE = SIM)
linhas = [r for r in range(2, ws.max_row + 1)
          if ws.cell(r, 1).value and norm(g(r, "ACORDO SUDESTE")) == "SIM"]
fora_acordo = sum(1 for r in range(2, ws.max_row + 1)
                  if ws.cell(r, 1).value and norm(g(r, "ACORDO SUDESTE")) != "SIM")

PAGO = "PAGO"
mapeado = sum(gv(r, "VALOR DO PLEITO") or 0 for r in linhas)
n_ind = len(linhas)

nome_exib = {}          # chave normalizada -> grafia com acento
status_exib = {}
fonte_status = defaultdict(lambda: defaultdict(lambda: [0, 0.0]))  # status -> fonte -> [qtd, valor]
obj_status = defaultdict(lambda: defaultdict(float))
obj_qtd = defaultdict(Counter)
obj_polo = defaultdict(list)            # polo -> [(beneficiário, descrição, valor)]
por_status_v = defaultdict(float)
por_status_n = Counter()
por_instr = defaultdict(lambda: [0, 0.0, 0.0])       # n, valor, pago
por_mun = defaultdict(lambda: [0, 0.0, 0.0])
por_padrinho = defaultdict(lambda: [0, 0.0])
objetos_instr = defaultdict(lambda: defaultdict(float))
etapas = defaultdict(lambda: [0, 0.0, []])

for r in linhas:
    v = gv(r, "VALOR DO PLEITO") or 0
    st = norm(g(r, COL_STATUS))
    inst = norm(g(r, "INSTRUMENTO"))
    mun = norm(g(r, "MUNICÍPIO"))
    nome_exib.setdefault(mun, titulo(g(r, "MUNICÍPIO")))
    status_exib.setdefault(norm(g(r, COL_STATUS)), g(r, COL_STATUS))
    pg = v if st == PAGO else 0
    por_status_v[st] += v
    por_status_n[st] += 1
    por_instr[inst][0] += 1
    por_instr[inst][1] += v
    por_instr[inst][2] += pg
    por_mun[mun][0] += 1
    por_mun[mun][1] += v
    por_mun[mun][2] += pg
    pad = g(r, "PADRINHO") or "Sem padrinho declarado"
    por_padrinho[norm(pad) if pad != "Sem padrinho declarado" else pad][0] += 1
    por_padrinho[norm(pad) if pad != "Sem padrinho declarado" else pad][1] += v
    objetos_instr[inst][titulo(g(r, "OBJETO"))] += v
    _f = (g(r, "FONTE FINAL") or "Não informada").replace("Fonte ", "").strip()
    fonte_status[st][_f][0] += 1
    fonte_status[st][_f][1] += v
    obj_status[st][titulo(g(r, "OBJETO"))] += v
    obj_qtd[st][titulo(g(r, "OBJETO"))] += 1
    if st != PAGO:
        et = g(r, "ETAPA DE CELEBRAÇÃO") or "NA"
        etapas[et][0] += 1
        etapas[et][1] += v
        etapas[et][2].append((titulo(g(r, "MUNICÍPIO")), titulo(g(r, "BENEFICIÁRIO")), v,
                              g(r, "DETALHAMENTO ETAPA DE CELEBRAÇÃO")))

pago = por_status_v.get(PAGO, 0)
n_pago = por_status_n.get(PAGO, 0)
nao_pago = mapeado - pago
n_nao_pago = n_ind - n_pago

# ── PLANO SUDESTE (previsão) ────────────────────────────────────────────────
ps = wb["Plano Sudeste"]
HP = [str(ps.cell(1, c).value or "").strip() for c in range(1, ps.max_column + 1)]
IP = {h: i + 1 for i, h in enumerate(HP) if h}


def p(r, n):
    return str(ps.cell(r, IP[n]).value or "").strip() if n in IP else ""


def pvv(r, n):
    return ps.cell(r, IP[n]).value if n in IP else None


pl_linhas = [r for r in range(2, ps.max_row + 1) if ps.cell(r, 1).value]
plano_total = sum(pvv(r, "Valor total") or 0 for r in pl_linhas)
plano_polo = defaultdict(float)
plano_status_v = defaultdict(float)
plano_status_n = Counter()
# De/Para dos status do plano — mesmos rótulos do dashboard (planosudeste.netlify.app)
DE_PARA_STATUS = {
    "JA RODANDO": "Em execução",
    "RODAR": "Aprovado para execução",
    "DEFINIR": "Em definição",
    "AGUARDA PROJETO": "Aguardando elaboração de projeto",
    "ANALISE PRELIMINAR": "Em análise preliminar",
    "FORA - DESISTIRAM": "Desistência confirmada",
    "MELHOR EM BICAS": "Sugestão de realocação territorial",
    "NAO VIAVEL PARA O MOMENTO": "Inviável no momento",
    "REGIONAL VAI PEDIR OFICIO PARA COMECARMOS A RODAR.": "Aguardando ofício regional",
    "REGIONAL VAI PEDIR OFICIO PARA COMECARMOS A RODAR": "Aguardando ofício regional",
}


def status_plano(bruto):
    """Rótulo do dashboard. 'Fora' se desdobra: com recurso já publicado é
    remanejamento de fonte; sem observação, saiu do escopo do plano."""
    k = norm(bruto)
    if k == "FORA":
        return "Remapeado"  # detalhado por item na função status_plano_item
    return DE_PARA_STATUS.get(k, " ".join(str(bruto or "").split()))


# situações que retiram o item da carteira ativa
FORA = ("FORA", "FORA - DESISTIRAM", "NAO VIAVEL PARA O MOMENTO")
plano_fora_v = plano_fora_n = 0
definir_itens = []
for r in pl_linhas:
    v = pvv(r, "Valor total") or 0
    st = norm(p(r, "Status"))
    polo_key = norm(p(r, "Município Polo"))
    nome_exib[polo_key] = " ".join(p(r, "Município Polo").split())
    status_exib[st] = status_plano(p(r, "Status"))
    plano_polo[polo_key] += v
    plano_status_v[st] += v
    plano_status_n[st] += 1
    if st in FORA:
        plano_fora_v += v
        plano_fora_n += 1
    if st == "DEFINIR":
        definir_itens.append((" ".join(p(r, "Município Polo").split()), v))
    if st not in FORA:
        obj_polo[polo_key].append((" ".join(p(r, "Destino do Investimeto").split()),
                                   " ".join(p(r, "Descrição").split()), v,
                                   status_plano(p(r, "Status")),
                                   " ".join(p(r, "Observações").split())))

carteira_ativa = plano_total - plano_fora_v
polos_plano = {k for k in plano_polo if k not in ("TODA A MACRORREGIAO", "CONSORCIOS GENERALISTAS")}
polos_com_ind = {k for k in por_mun}
polos_sem_ind = sorted(polos_plano - polos_com_ind)

# funil
p_map = 100 * mapeado / plano_total
p_pg = 100 * pago / plano_total
p_pg_map = 100 * pago / mapeado

hoje = datetime.date.today()
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
data_ext = f"{hoje.day} de {MESES[hoje.month-1]} de {hoje.year}"

# ── FRAGMENTOS ──────────────────────────────────────────────────────────────
ST_LABEL = {"PAGO": "Pago", "EM TRAMITACAO": "Em tramitação",
            "EM ANALISE PREVIA": "Em análise prévia",
            "AGUARDA CONVENIO PROJETO": "Aguarda convênio/projeto"}
ST_COR = {"PAGO": "ok", "EM TRAMITACAO": "at", "EM ANALISE PREVIA": "at",
          "AGUARDA CONVENIO PROJETO": "cr"}

ORDEM_EXEC = ["AGUARDA CONVENIO PROJETO", "EM ANALISE PREVIA", "EM TRAMITACAO", "PAGO"]


def _grupo_fonte(f):
    """Fonte 10 e Fonte 10 - SEGOV entram na mesma coluna."""
    return "10" if f.startswith("10") else ("95" if f.startswith("95") else "outra")


def valor_fonte(st, grupo):
    tot = sum(val for f, (q, val) in fonte_status.get(st, {}).items() if _grupo_fonte(f) == grupo)
    qtd = sum(q for f, (q, val) in fonte_status.get(st, {}).items() if _grupo_fonte(f) == grupo)
    if not qtd:
        return "<span class='vazio'>&mdash;</span>"
    return f"{brl(tot,0)}<span class='qtd'>{qtd} ind.</span>"


tem_segov = any(f.startswith("10") and "SEGOV" in f.upper()
                for st in fonte_status for f in fonte_status[st])
segov_v = sum(val for st in fonte_status for f, (q, val) in fonte_status[st].items()
              if "SEGOV" in f.upper())
segov_n = sum(q for st in fonte_status for f, (q, val) in fonte_status[st].items()
              if "SEGOV" in f.upper())


def corta(t, lim):
    t = " ".join(str(t or "").split())
    if len(t) <= lim:
        return t
    return t[:lim].rsplit(" ", 1)[0].rstrip(" -,;") + "…"


# categorias de objeto — resume o texto bruto em rótulos legíveis
CATEGORIAS = [
    ("Obra do HPS",              ["CONSTRUCAO HPS", "HPS"]),
    ("Projetos e estudos",       ["DIAGNOSTICO", "ESTUDO DE VIABILIDADE", "PROJETO BASICO", "PROJETOS EXECUTIVOS"]),
    ("Obras de UBS",             ["UBS"]),
    ("Obras de UPA",             ["UPA"]),
    ("Obra hospitalar",          ["OBRA E REFORMA", "HOSPITAL REGIONAL", "REFORMA"]),
    ("Oncologia",                ["ACELERADOR LINEAR"]),
    ("SAMU e transporte",        ["SAMU", "AMBULANCIA", "VEICULO", "VACIMOVEL", "PICK UP", "PICK-UP",
                                  "UNIDADE ODONTOLOGICA MOVEL", "TRANSPORTE SANITARIO"]),
    ("Equipamentos de oftalmologia", ["VITREOFAGO", "FACOEMULSIFICADOR", "RETINOGRAFO", "CAMPIMETRO",
                                      "AUTOREFRATOR", "TOPOGRAFO DE CORNEA", "OFTALMOSCOPIO", "REFRATOR",
                                      "TONOMETRO", "BIOMETRO", "OFTALMOLOGIC", "COERENCIA OPTICA"]),
    ("Equipamentos de APS e saúde bucal", ["APS", "ATENCAO PRIMARIA", "SAUDE BUCAL", "CEO", "ODONTOLOG"]),
    ("Custeio e manutenção",     ["CUSTEIO", "MANUTENCAO"]),
    ("Equipamentos hospitalares", ["EQUIPAMENT", "APARELHO", "ULTRASSOM", "RAIOS X", "MESA CIRURGICA",
                                   "VENTILADOR", "ARCO CIRURGICO", "ENDOSCOPIA", "ESTERILIZADOR",
                                   "MONITOR", "CAMAS", "ASPIRADOR", "CAMARAS", "MICROSCOPIO", "ANESTESIA"]),
]


def categoria(obj):
    t = unicodedata.normalize("NFKD", str(obj or "")).encode("ascii", "ignore").decode().upper()
    for rotulo, chaves in CATEGORIAS:
        if any(k in t for k in chaves):
            return rotulo
    return "Outros objetos"


def objetos_status(st, n=4):
    """Categorias do status, da maior para a menor em valor, separadas por barra."""
    cats = defaultdict(float)
    for o, val in obj_status[st].items():
        cats[categoria(o)] += val
    ordenadas = [c for c, _ in sorted(cats.items(), key=lambda x: -x[1])]
    return " / ".join(ordenadas[:n])


_ordem = [k for k in ORDEM_EXEC if k in por_status_v] + \
         [k for k in por_status_v if k not in ORDEM_EXEC]
linhas_status = "\n".join(
    f"""<tr>
      <td class='nome'><span class='dot d-{ST_COR.get(k,'at')}'></span>{ST_LABEL.get(k, status_exib.get(k, titulo(k)))}</td>
      <td class='num'>{por_status_n[k]}</td>
      <td class='num tot'>{brl(por_status_v[k], 0)}</td>
      <td class='num f95'>{valor_fonte(k, "95")}</td>
      <td class='num f10'>{valor_fonte(k, "10")}</td>
      <td class='fin'>{objetos_status(k)}</td>
    </tr>"""
    for k in _ordem)
linhas_status += f"""<tr class='soma'>
      <td class='nome'>Total</td>
      <td class='num'>{n_ind}</td>
      <td class='num tot'>{brl(mapeado, 0)}</td>
      <td class='num f95'>{brl(sum(val for st in fonte_status for f,(q,val) in fonte_status[st].items() if _grupo_fonte(f)=="95"), 0)}</td>
      <td class='num f10'>{brl(sum(val for st in fonte_status for f,(q,val) in fonte_status[st].items() if _grupo_fonte(f)=="10"), 0)}</td>
      <td></td>
    </tr>"""

INSTR_LABEL = {"RESOLUCAO": "Resolução", "RESOLUCAO UBS": "Resolução — UBS",
               "CONVENIO": "Convênio", "EXECUCAO DIRETA": "Execução direta"}


def objetos(inst, n=2):
    vistos, saida = set(), []
    for o, _ in sorted(objetos_instr[inst].items(), key=lambda x: -x[1]):
        curto = o[:46].rstrip(" -,;")
        if curto.upper() in vistos:
            continue
        vistos.add(curto.upper())
        saida.append(curto)
        if len(saida) == n:
            break
    return "; ".join(saida)


linhas_instr = "\n".join(
    f"""<tr>
      <td class='nome'>{INSTR_LABEL.get(k, titulo(k))}</td>
      <td class='num'>{n}</td>
      <td class='num'>{brl(v)}</td>
      <td class='barcel'><span class='pct'>{pct(100*pg/v if v else 0)}</span>
        <div class='mini'><div class='mfill f-ok' style='width:{w(100*pg/v if v else 0)}%'></div></div></td>
      <td class='fin'>{objetos(k)}</td>
    </tr>"""
    for k, (n, v, pg) in sorted(por_instr.items(), key=lambda x: -x[1][1]))

# municípios: plano × mapeado × pago
mun_todos = sorted(set(por_mun) | polos_plano,
                   key=lambda m: -(plano_polo.get(m, 0) + por_mun.get(m, [0, 0, 0])[1]))
linhas_mun = ""
for m in mun_todos:
    prev = plano_polo.get(m, 0)
    n, vmap, vpg = por_mun.get(m, [0, 0, 0])
    if prev == 0 and vmap == 0:
        continue
    linhas_mun += f"""<tr class='{"tem" if vmap else "sem"}'>
      <td class='nome'>{nome_exib.get(m, titulo(m))}</td>
      <td class='num'>{brl(prev,0) if prev else '&mdash;'}</td>
      <td class='num'>{brl(vmap,0) if vmap else '&mdash;'}<span class='sub'>{f' · {n} ind.' if n else ''}</span></td>
      <td class='num'>{brl(vpg,0) if vpg else '&mdash;'}</td>
      <td class='barcel'>{f"<span class='pct'>{pct(100*vpg/vmap)}</span><div class='mini'><div class='mfill f-ok' style='width:{w(100*vpg/vmap)}%'></div></div>" if vmap else "<span class='pct sub'>sem indicação</span>"}</td>
    </tr>"""

# etapas pendentes
ordem_etapas = sorted(etapas.items(), key=lambda x: -x[1][1])
linhas_etapas = ""
for et, (n, v, itens) in ordem_etapas:
    quem = "; ".join(f"{mu} — {be[:34]}" for mu, be, _, _ in sorted(itens, key=lambda x: -x[2]))
    obs = next((d for _, _, _, d in sorted(itens, key=lambda x: -x[2]) if d and norm(d) != "NA"), "")
    rot = et if norm(et) != "NA" else "Sem etapa registrada"
    linhas_etapas += f"""<tr>
      <td class='nome'>{titulo(rot) if norm(rot)!='NA' else rot}</td>
      <td class='num'>{n}</td>
      <td class='num'>{brl(v,0)}</td>
      <td class='fin'>{quem}{f"<br><i>{titulo(obs)}</i>" if obs else ""}</td>
    </tr>"""

# blocos de atenção
maior_parado = max(((r, gv(r, "VALOR DO PLEITO") or 0) for r in linhas
                    if norm(g(r, COL_STATUS)) != PAGO), key=lambda x: x[1], default=(None, 0))
blocos = ""
if maior_parado[0]:
    r = maior_parado[0]
    blocos += f"""<div class='bl'>
      <div class='bl-h'><div class='bl-n'>{mi(maior_parado[1])}</div><div>
        <div class='bl-t'>{titulo(g(r,'BENEFICIÁRIO'))} — {titulo(g(r,'MUNICÍPIO'))}</div>
        <div class='bl-s'>{titulo(g(r,'OBJETO'))} · {ST_LABEL.get(norm(g(r,COL_STATUS)), titulo(g(r,COL_STATUS)))}</div></div></div>
      <div class='bl-m'>Maior valor individual parado — sozinho representa {pct(100*maior_parado[1]/nao_pago,0)} de tudo o que ainda não foi pago.</div></div>"""
if polos_sem_ind:
    # o total de cada polo é a SOMA DOS ITENS LISTADOS (itens descartados ficam de fora)
    ativos = {m: sorted(obj_polo.get(m, []), key=lambda x: -x[2]) for m in polos_sem_ind}
    ativos = {m: itens for m, itens in ativos.items() if itens}
    v_sem = sum(sum(i[2] for i in itens) for itens in ativos.values())
    n_itens = sum(len(itens) for itens in ativos.values())
    itens_polo = ""
    for m, itens in sorted(ativos.items(), key=lambda x: -sum(i[2] for i in x[1])):
        tot_m = sum(i[2] for i in itens)
        linhas_i = ""
        for benef, desc, v, st_i, obs in itens:
            chip = ("" if st_i == "Em definição"
                    else f"<span class='st-chip'>{st_i}</span>")
            OBS_GENERICA = ("REQUER ANALISE TECNICA APROFUNDADA",)
            nota = (f"<span class='it-obs'>{corta(obs, 78)}</span>"
                    if obs and norm(obs) not in ("NONE", "") + OBS_GENERICA else "")
            linhas_i += (f"<div class='it'>"
                         f"<div class='it-esq'><div class='it-d'>{corta(desc, 64)}</div>"
                         f"<div class='it-b'>{corta(benef, 46)}{chip}</div>{nota}</div>"
                         f"<div class='it-v'>{brl(v,0)}</div></div>")
        itens_polo += (f"<div class='pl'>"
                       f"<div class='pl-h'><span class='pl-m'>{nome_exib.get(m, titulo(m))}</span>"
                       f"<span class='pl-c'>{len(itens)} {'item' if len(itens)==1 else 'itens'}</span>"
                       f"<span class='pl-v'>{brl(tot_m,0)}</span></div>{linhas_i}</div>")
    blocos += f"""<div class='bl'>
      <div class='bl-h'><div class='bl-n'>{len(ativos)}</div><div>
        <div class='bl-t'>Municípios-polo sem nenhuma indicação mapeada</div>
        <div class='bl-s'>{n_itens} itens previstos, somando {brl(v_sem,0)}, ainda sem contrapartida na Fonte 95</div></div></div>
      <div class='bl-lista'>{itens_polo}</div></div>"""


# ── FRENTES (cartões com barras empilhadas, no padrão Rio Doce) ──────────────
pl_rodando = plano_status_n.get("JA RODANDO", 0)
pl_fora_n = plano_fora_n
pl_definindo = len(pl_linhas) - pl_rodando - pl_fora_n

ind_pagas = n_pago
ind_analise = por_status_n.get("EM TRAMITACAO", 0) + por_status_n.get("EM ANALISE PREVIA", 0)
ind_travadas = n_ind - ind_pagas - ind_analise

cel_pend = sum(n for _, (n, _, _) in etapas.items())
cel_avancadas = sum(n for et, (n, _, _) in etapas.items() if norm(et).startswith("ETAPA 7"))
cel_inicio = sum(n for et, (n, _, _) in etapas.items() if norm(et).startswith(("ETAPA 2", "ETAPA 3")))
cel_sem = cel_pend - cel_avancadas - cel_inicio

polos_tot = len(polos_plano)
polos_ok = len(polos_plano & polos_com_ind)
polos_nao = len(polos_sem_ind)


def card(titulo_, sufixo, legenda, segs, obs):
    """segs = lista de (classe, quantidade); a última fatia usa flex:1."""
    tot = max(1, sum(q for _, q in segs))
    barras = ""
    for i, (cls, q) in enumerate(segs):
        estilo = "flex:1" if i == len(segs) - 1 else f"width:{w(100*q/tot)}%"
        barras += f"<div class='s-{cls}' style='{estilo}'></div>"
    return f"""<div class='fr'>
      <div class='fr-t'>{titulo_} <span>· {sufixo}</span></div>
      <div class='fr-s'>{legenda}</div>
      <div class='stk'>{barras}</div>
      <div class='fr-l'>{obs}</div>
    </div>"""


cards = (
    card("Indicações Fonte 95", f"{n_ind} indicações", "status do pagamento",
         [("ok", ind_pagas), ("at", ind_analise), ("cr", ind_travadas)],
         f"<b>{ind_pagas}</b> pagas · <b>{ind_analise}</b> em tramitação ou análise · <b style='color:var(--barro)'>{ind_travadas}</b> aguardando convênio") +
    card("Celebração pendente", f"{cel_pend} processos", "etapa de tramitação alcançada",
         [("ok", cel_avancadas), ("at", cel_inicio), ("cr", cel_sem)],
         f"<b>{cel_avancadas}</b> assinado/publicado · <b>{cel_inicio}</b> em etapas iniciais · <b style='color:var(--barro)'>{cel_sem}</b> sem etapa registrada") +
    card("Cobertura dos polos", f"{polos_tot} municípios-polo", "polos com indicação mapeada",
         [("ok", polos_ok), ("cr", polos_nao)],
         f"<b>{polos_ok}</b> com indicação · <b style='color:var(--barro)'>{polos_nao}</b> sem nenhuma indicação")
)

# ── HTML ────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Boletim de Acompanhamento — Plano Sudeste · Fonte 95</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,400..900&family=Source+Serif+4:opsz,wght@8..60,400..700&display=swap" rel="stylesheet">
<style>
  :root {{
    --tinta:#161E28; --serra:#1B4F72; --serra-esc:#0E3049; --serra-cl:#E9F0F6;
    --barro:#B04A21; --barro-cl:#FBEDE6; --ambar:#8A5E00; --ambar-cl:#FBF3DF;
    --mata:#1B7A55; --mata-cl:#E6F3EC; --painel:#F1F4F7; --fio:#D9E1E8;
    --cinza:#5A6672; --cinza-cl:#8C97A2;
  }}
  * {{ box-sizing:border-box; margin:0; }}
  html {{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }}
  body {{ font-family:'Source Serif 4','Cambria',Georgia,serif; color:var(--tinta); background:#E4E9ED;
         font-size:10.6pt; line-height:1.55; }}
  .arch {{ font-family:'Archivo','Segoe UI','Helvetica Neue',Arial,sans-serif; }}
  .pagina {{ background:#fff; width:210mm; min-height:297mm; margin:10mm auto; padding:16mm 17mm 14mm; display:flex; flex-direction:column;
             box-shadow:0 2px 18px rgba(22,30,40,.14); position:relative; }}
  .fio-top {{ height:2.5px; border:0; background:linear-gradient(90deg,var(--serra) 0%,var(--serra) 60%,var(--barro) 100%); margin:0 0 5mm; }}
  .cabo {{ display:flex; justify-content:space-between; align-items:flex-end; padding-bottom:4mm; }}
  .cabo-org {{ font-family:'Archivo'; font-size:8pt; font-weight:600; letter-spacing:.14em; text-transform:uppercase; color:var(--cinza); line-height:1.5; }}
  .cabo-num {{ font-family:'Archivo'; font-size:8pt; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--serra); text-align:right; line-height:1.5; }}
  .titulo {{ font-family:'Archivo'; font-weight:800; font-stretch:87%; font-size:23pt; line-height:1.05;
             letter-spacing:-.01em; color:var(--serra-esc); margin:5mm 0 1.5mm; }}
  .titulo em {{ font-style:normal; color:var(--barro); }}
  .subtitulo {{ font-family:'Archivo'; font-size:9.2pt; font-weight:500; color:var(--cinza); margin-bottom:4mm; }}
  .sec {{ margin-top:5mm; }}
  .eyebrow {{ font-family:'Archivo'; font-size:8.5pt; font-weight:700; letter-spacing:.16em;
              text-transform:uppercase; color:var(--serra); display:flex; align-items:center; gap:8px; margin-bottom:2.8mm; }}
  .eyebrow::after {{ content:''; flex:1; height:1px; background:var(--fio); }}
  .eyebrow.barro {{ color:var(--barro); }}

  /* barra empilhada — mesmo desenho do dashboard */
  .tot-lbl {{ font-family:'Archivo'; font-size:8.5pt; color:var(--cinza-cl); margin-bottom:.6mm; }}
  .tot-val {{ font-family:'Archivo'; font-size:17pt; font-weight:600; color:var(--tinta); letter-spacing:-.01em; }}
  .exec-hdr {{ display:flex; justify-content:space-between; align-items:baseline; margin:3.5mm 0 1.6mm; }}
  .exec-titulo {{ font-family:'Archivo'; font-size:8.5pt; color:var(--cinza-cl); }}
  .exec-pct {{ font-family:'Archivo'; font-size:12pt; font-weight:700; color:#1D9E75; }}
  .exec-bar-wrap {{ height:6mm; background:#EEF2F5; border-radius:2px; overflow:hidden;
                    position:relative; margin-bottom:2mm; }}
  .exec-bar-wrap > div {{ position:absolute; top:0; left:0; height:100%; }}
  .seg-tram {{ background:#EF9F27; z-index:1; }}
  .seg-pago {{ background:#1D9E75; z-index:2; }}
  .seg-map {{ border-top:.8mm dashed #378ADD; border-bottom:.8mm dashed #378ADD; z-index:3; }}
  .exec-legenda {{ display:flex; flex-wrap:nowrap; gap:4mm; font-family:'Archivo'; font-size:7.4pt;
                   color:var(--cinza); white-space:nowrap; }}
  .leg-item {{ display:flex; align-items:center; gap:1.2mm; min-width:0; }}
  .leg-item b {{ color:var(--tinta); font-weight:600; }}
  .leg-dot {{ width:2mm; height:2mm; border-radius:50%; flex-shrink:0; }}
  .d-prev {{ background:#EEF2F5; border:.3mm solid #D9E1E8; }}
  .d-map {{ background:#378ADD; }} .d-tram {{ background:#EF9F27; }} .d-pago {{ background:#1D9E75; }}
  .marco {{ font-family:'Archivo'; font-size:8pt; color:var(--cinza); margin-top:2.5mm; line-height:1.6; }}
  .marco b {{ color:var(--serra-esc); }}
  .tot {{ font-weight:700; }}
  .f95, .f10 {{ color:var(--serra); }}
  .qtd {{ display:block; font-size:7.6pt; font-weight:400; color:var(--cinza-cl); margin-top:.3mm; }}
  .vazio {{ color:#C7D0D8; }}
  tr.soma td {{ border-top:1.2px solid var(--tinta); border-bottom:none;
                font-weight:700; color:var(--tinta); padding-top:2.4mm; }}
  .pl {{ background:#fff; border:.5px solid #F0D9CB; border-radius:3px;
         padding:2.6mm 3.4mm; margin-bottom:2mm; }}
  .pl-h {{ display:flex; align-items:baseline; gap:3mm; padding-bottom:1.6mm;
           margin-bottom:1.6mm; border-bottom:.5px solid #F4E6DC; }}
  .pl-m {{ font-family:'Archivo'; font-weight:700; font-size:9pt; color:var(--tinta); }}
  .pl-c {{ font-family:'Archivo'; font-size:7.6pt; color:var(--cinza-cl); }}
  .pl-v {{ font-family:'Archivo'; font-weight:700; font-size:8.8pt; color:#8A3A15; margin-left:auto; }}
  .it {{ display:flex; justify-content:space-between; align-items:flex-start; gap:4mm; padding:1.1mm 0; }}
  .it + .it {{ border-top:.5px solid #FAF2EC; }}
  .it-esq {{ min-width:0; }}
  .it-d {{ font-family:'Source Serif 4'; font-size:8.6pt; color:var(--tinta); line-height:1.4; }}
  .it-b {{ font-family:'Archivo'; font-size:7.8pt; color:var(--cinza); margin-top:.5mm;
           display:flex; align-items:center; gap:2mm; flex-wrap:wrap; }}
  .it-obs {{ display:block; font-family:'Source Serif 4'; font-size:8.2pt; font-style:italic;
             color:var(--cinza-cl); margin-top:.5mm; line-height:1.4; }}
  .it-v {{ font-family:'Archivo'; font-size:8.4pt; font-weight:600; color:var(--tinta); white-space:nowrap; }}
  .st-chip {{ font-family:'Archivo'; font-size:7.2pt; font-weight:700; color:#8A5E00;
              background:var(--ambar-cl); border-radius:8px; padding:.3mm 1.8mm; }}
  .bl-lista {{ margin-top:2.5mm; padding-left:24mm; }}
  /* tabela de municípios */
  .tmun td {{ padding-top:2.2mm; padding-bottom:2.2mm; }}
  .tmun tbody tr:nth-child(odd) {{ background:#F7F9FB; }}
  .tmun td:first-child {{ padding-left:2.5mm; border-left:2.5px solid transparent; }}
  .tmun tr.tem td:first-child {{ border-left-color:var(--mata); }}
  .tmun tr.sem td:first-child {{ border-left-color:#DDE3E8; }}
  .tmun tr.sem td {{ color:var(--cinza-cl); }}
  .tmun .sub {{ display:block; font-size:7.8pt; }}
  table {{ width:100%; border-collapse:collapse; font-family:'Archivo'; font-size:8.6pt; }}
  th {{ text-align:left; font-size:7.1pt; font-weight:700; letter-spacing:.09em; text-transform:uppercase;
        color:var(--cinza); padding:0 3mm 1.8mm 0; border-bottom:1.5px solid var(--tinta); }}
  th.num {{ text-align:right; }}
  td {{ padding:1.55mm 3mm 1.55mm 0; border-bottom:.5px solid var(--fio); vertical-align:top; }}
  tr:last-child td {{ border-bottom:none; }}
  .nome {{ font-weight:600; }}
  .num {{ text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }}
  .sub {{ color:var(--cinza-cl); font-weight:400; }}
  .fin {{ font-family:'Source Serif 4'; font-size:8.4pt; color:var(--cinza); line-height:1.45; }}
  .fin i {{ color:var(--cinza-cl); }}
  .barcel {{ width:30mm; }}
  .pct {{ font-weight:700; display:block; margin-bottom:.8mm; font-variant-numeric:tabular-nums; }}
  .mini {{ height:2mm; background:#E3E9EE; border-radius:2px; overflow:hidden; }}
  .mfill {{ height:100%; border-radius:2px; }}
  .f-ok {{ background:var(--mata); }} .f-at {{ background:#D3922B; }} .f-cr {{ background:var(--barro); }}
  .dot {{ display:inline-block; width:6px; height:6px; border-radius:50%; margin-right:6px; vertical-align:middle; }}
  .d-ok {{ background:var(--mata); }} .d-at {{ background:#D3922B; }} .d-cr {{ background:var(--barro); }}

  .frentes {{ display:grid; grid-template-columns:1fr 1fr; gap:3mm; }}
  .fr {{ background:var(--painel); border-radius:3px; padding:3.2mm 4mm; }}
  .fr-t {{ font-family:'Archivo'; font-weight:700; font-size:10pt; }}
  .fr-t span {{ color:var(--cinza); font-weight:500; }}
  .fr-s {{ font-family:'Archivo'; font-size:8pt; color:var(--cinza); margin:.3mm 0 2.5mm; }}
  .stk {{ display:flex; height:4.5mm; border-radius:2px; overflow:hidden; margin-bottom:2mm; }}
  .s-ok {{ background:var(--mata); }} .s-at {{ background:#D3922B; }} .s-cr {{ background:var(--barro); }}
  .fr-l {{ font-family:'Archivo'; font-size:8.4pt; color:var(--cinza); line-height:1.55; }}
  .fr-l b {{ color:var(--tinta); }}
  .bl {{ background:var(--barro-cl); border-left:3px solid var(--barro); border-radius:3px; padding:3.2mm 4.5mm; margin-bottom:2.5mm; }}
  .bl-h {{ display:flex; gap:4mm; align-items:baseline; }}
  .bl-n {{ font-family:'Archivo'; font-weight:800; font-size:14pt; color:var(--barro); min-width:20mm; }}
  .bl-t {{ font-family:'Archivo'; font-weight:700; font-size:10.3pt; }}
  .bl-s {{ font-family:'Archivo'; font-size:8.6pt; color:#8A3A15; }}
  .bl-m {{ font-family:'Archivo'; font-size:8.6pt; color:var(--cinza); margin-top:1.8mm; padding-left:24mm; line-height:1.6; }}

  .rodape {{ margin-top:10mm; padding-top:2mm; display:flex; justify-content:space-between;
             font-family:'Archivo'; font-size:7.5pt; color:var(--cinza-cl); border-top:.5px solid var(--fio); }}
  .btn-print {{ position:fixed; top:14px; right:14px; z-index:9; font-family:'Archivo'; font-weight:700; font-size:13px;
                background:var(--serra); color:#fff; border:0; border-radius:22px; padding:10px 18px; cursor:pointer;
                box-shadow:0 3px 10px rgba(14,48,73,.35); }}
  .btn-print:hover {{ background:var(--serra-esc); }}
  @page {{ size:A4; margin:15mm 16mm 13mm; }}
  @media print {{
    body {{ background:#fff; font-size:10.2pt; }}
    .pagina {{ margin:0; padding:0; box-shadow:none; width:auto; }}
    .btn-print {{ display:none; }}
    .eyebrow, .fr, .exec-bar-wrap, .exec-legenda, tr, thead, .pl, .it {{ page-break-inside:avoid; break-inside:avoid; }}
    .bl {{ break-inside:auto; }}
    .eyebrow {{ break-after:avoid; page-break-after:avoid; }}
    thead {{ display:table-header-group; }}
    p, .marco {{ orphans:3; widows:3; }}
    .rodape {{ margin-top:8mm; }}
    .fin {{ font-size:8.6pt; line-height:1.4; }}
    .subtitulo {{ font-size:9.4pt; }}
    .fr-l, .bl-m {{ line-height:1.45; }}
    td {{ padding-top:1.4mm; padding-bottom:1.4mm; }}
  }}
  @media screen and (max-width:820px) {{ .pagina {{ width:auto; min-height:auto; padding:8mm 6mm; }} }}
</style>
</head>
<body>
<button class="btn-print" onclick="window.print()">🖨&nbsp; Imprimir / salvar PDF</button>

<!-- ══════════ PÁGINA 1 ══════════ -->
<div class="pagina">
  <div class="cabo">
    <div class="cabo-org">Secretaria de Estado de Saúde de Minas Gerais<br>Subsecretaria de Regionalização · CMIR</div>
    <div class="cabo-num">Boletim de acompanhamento<br>Edição {hoje.strftime("%m/%Y")} · {hoje.strftime("%d/%m/%Y")}</div>
  </div>
  <hr class="fio-top">
  <div class="titulo">Plano de Investimentos<br>da Macrorregião <em>Sudeste</em></div>
  

  <div class="sec">
    <div class="eyebrow arch">O curso do recurso</div>
    <div class="tot-lbl">Previsão inicial (Plano Sudeste)</div>
    <div class="tot-val">{brl(plano_total, 0)}</div>
    <div class="exec-hdr">
      <span class="exec-titulo">Distribuição por status (% da previsão)</span>
      <span class="exec-pct">{pct(p_pg)} pago</span>
    </div>
    <div class="exec-bar-wrap">
      <div class="seg-tram" style="width:{w(100*mapeado/plano_total)}%"></div>
      <div class="seg-pago" style="width:{w(p_pg)}%"></div>
      <div class="seg-map" style="width:{w(100*mapeado/plano_total)}%"></div>
    </div>
    <div class="exec-legenda">
      <div class="leg-item"><span class="leg-dot d-prev"></span>Previsão inicial <b>{brl(plano_total,0)}</b></div>
      <div class="leg-item"><span class="leg-dot d-map"></span>Pleitos mapeados <b>{brl(mapeado,0)}</b></div>
      <div class="leg-item"><span class="leg-dot d-tram"></span>Em tramitação <b>{brl(nao_pago,0)}</b></div>
      <div class="leg-item"><span class="leg-dot d-pago"></span>Pago <b>{brl(pago,0)}</b></div>
    </div>
    

  <div class="sec">
    
    <table>
      <thead><tr><th>Status em {DATA_STATUS}</th><th class="num">Ind.</th><th class="num">Valor total</th><th class="num">Fonte 95</th><th class="num">Fonte 10</th><th>Principais Objetos</th></tr></thead>
      <tbody>{linhas_status}</tbody>
    </table>
    

  <div class="sec">
    <div class="eyebrow arch barro">Pontos de atenção</div>
    {blocos}
  </div>

  <div class="sec">
    <div class="eyebrow arch">Previsão × execução por município</div>
    <table class="tmun">
      <thead><tr><th>Município</th><th class="num">Previsto no plano</th><th class="num">Mapeado</th><th class="num">Pago</th><th>Execução</th></tr></thead>
      <tbody>{linhas_mun}</tbody>
    </table>
    <div class="marco">Descontados os {plano_fora_n} itens fora do plano ({brl(plano_fora_v,0)}), a carteira ativa é de
    <b>{brl(carteira_ativa,0)}</b>. {f"Foram desconsideradas {fora_acordo} indicações marcadas como fora do Acordo Sudeste." if fora_acordo else ""}</div>
  </div>

  <div class="rodape">
    <span>Fonte: Monitoramento Fonte 95.xlsx (ASPAR)</span>
    <span>Elaboração: CMIR / Subsecretaria de Regionalização — SES-MG</span>
  </div>
</div>
</body>
</html>"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print(f"gerado: {OUT} ({len(html):,} chars)")

if DASH:
    injetar_botao(DASH, os.path.basename(OUT))
