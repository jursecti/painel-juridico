#!/usr/bin/env python3
"""Lê a planilha JURIDICO (somente leitura), separa as colunas públicas,
calcula os prazos e gera site/data.json. Com --email, envia o resumo semanal.

Uso no GitHub Actions:  python scripts/build.py --out site/data.json [--email]
Teste local:            python scripts/build.py --fixture tests/fixture --out site/data.json --previa-email /tmp/email.html
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import smtplib
import sys
import unicodedata
from collections import Counter, defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

RAIZ = Path(__file__).resolve().parent.parent
FUSO = ZoneInfo("America/Recife")

# ---------------------------------------------------------------- texto

def norm(s):
    """Maiúsculas, sem acento, só letras e números separados por um espaço."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def limpo(s):
    return re.sub(r"[ \t]+", " ", str(s or "")).strip()


RE_CPF = re.compile(r"(CPF\s*:?\s*(n[ºo°]\.?\s*)?)?\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", re.I)
RE_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(\.[\w-]+)+")
RE_FONE = re.compile(
    r"\(\d{2}\)\s*\d{4,5}[-.\s]?\d{4}"
    r"|\b9\d{4}-\d{4}\b"
    r"|\b(fone|fones|tel|telefone|cel|celular|whatsapp|ramal)\b\.?\s*:?\s*[\d()][\d\s()./-]{6,}",
    re.I,
)
RE_MATRICULA = re.compile(
    r",?\s*\(?\bmat(r[ií]cula|r)?\b\.?\s*(SGP)?\s*(n[ºo°]\.?)?\s*:?\s*\d[\d./-]*\)?", re.I
)


RE_CIENCIA = re.compile(r"\(\s*Termo de Ci[êe]ncia[^)]*\)", re.I)


def publico(s):
    """Rede de segurança: tira CPF, telefone, e-mail e matrícula de qualquer texto publicado."""
    s = str(s or "")
    for rx in (RE_CPF, RE_EMAIL, RE_FONE, RE_MATRICULA, RE_CIENCIA):
        s = rx.sub("", s)
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r" ?- ?(?=\n|$)", "", s)
    s = re.sub(r"\n\s*\n+", "\n", s)
    return s.strip(" \n,;-")


RE_DATA = re.compile(r"\b(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})\b")


def datas(s):
    achadas = []
    for d, m, a in RE_DATA.findall(str(s or "")):
        a = int(a)
        a += 2000 if a < 100 else 0
        try:
            achadas.append(dt.date(a, int(m), int(d)))
        except ValueError:
            pass
    return achadas


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", norm(s).lower()).strip("-")


# ---------------------------------------------------------------- leitura

def ler_fixture(pasta):
    abas = {}
    for arq in sorted(Path(pasta).glob("*.json")):
        d = json.loads(arq.read_text(encoding="utf-8"))
        abas[d["aba"]] = d["valores"]
    return abas, list(abas)


def ler_planilha(planilha_id, nomes):
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account

    if not os.environ.get("GOOGLE_CREDENTIALS"):
        sys.exit("ERRO: falta o segredo GOOGLE_CREDENTIALS no repositório (veja o README, passo 4).")
    info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    cred = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    sessao = AuthorizedSession(cred)
    base = f"https://sheets.googleapis.com/v4/spreadsheets/{planilha_id}"
    r = sessao.get(base, params={"fields": "sheets.properties(title,hidden)"}, timeout=60)
    r.raise_for_status()
    existentes = [s["properties"]["title"] for s in r.json()["sheets"] if not s["properties"].get("hidden")]
    pedir = [n for n in nomes if n in existentes]
    abas = {}
    for i in range(0, len(pedir), 4):
        lote = pedir[i:i + 4]
        r = sessao.get(
            base + "/values:batchGet",
            params={"ranges": ["'" + n.replace("'", "''") + "'" for n in lote], "valueRenderOption": "FORMATTED_VALUE"},
            timeout=180,
        )
        r.raise_for_status()
        for nome, vr in zip(lote, r.json()["valueRanges"]):
            abas[nome] = vr.get("values", [])
    return abas, existentes


class Tabela:
    """Uma aba, com o cabeçalho localizado pelo nome das colunas."""

    def __init__(self, valores, esperados):
        alvo = {norm(e.split("|")[0]) for e in esperados if e}
        melhor, pontos = 0, -1
        for i, linha in enumerate(valores[:8]):
            ns = {norm(c) for c in linha if limpo(c)}
            p = sum(1 for a in alvo if a in ns or any(n.startswith(a + " ") for n in ns))
            if p > pontos:
                melhor, pontos = i, p
        self.linha_cab = melhor
        self.cab = [limpo(str(c).replace("\n", " ")) for c in (valores[melhor] if valores else [])]
        self.ncab = [norm(c) for c in self.cab]
        self.linhas = [(melhor + 2 + j, l) for j, l in enumerate(valores[melhor + 1:]) if any(limpo(c) for c in l)]

    def col(self, nome):
        if not nome:
            return None
        n = norm(nome.split("|")[0])
        if n in self.ncab:
            return self.ncab.index(n)
        for i, c in enumerate(self.ncab):
            if c.startswith(n + " "):
                return i
        return None

    @staticmethod
    def val(linha, i):
        return limpo(linha[i]) if i is not None and i < len(linha) else ""

    @staticmethod
    def bruto(linha, i):
        return str(linha[i]).strip() if i is not None and i < len(linha) else ""


# ---------------------------------------------------------------- regras

INATIVO = ("ENCERR", "RESCIND", "CANCEL", "CONCLU", "EXTINT", "REVOGAD", "FINALIZ", "SUSPENS", "ANULAD", "SEM EFEITO", "ARQUIV", "VENCID")
ATIVO = ("VIGENTE", "EXECUC", "A VENCER", "ATIVO", "ANDAMENTO", "ABERTO")
NAO_PESSOA = {"SECRETARIO", "SECRETARIA", "DIRETOR", "DIRETORA", "DIRETORIA", "GERENTE", "GERENCIA", "GESTOR", "GESTORA", "PRESIDENTE", "COORDENADOR", "COORDENADORA", "COORDENACAO", "MEMBRO", "TITULAR", "SUPLENTE", "NAO", "SEM", "UNIDADE", "ASSESSORIA", "ASSESSOR", "ASSESSORA", "GABINETE", "SUPERINTENDENCIA", "SUPERINTENDENTE", "COMISSAO", "PORTARIA", "TERMO", "ANTES", "ATE", "A", "VAGO", "CHEFE", "FISCAL", "SERVIDOR", "REPRESENTANTE", "INDICACAO", "AGUARDANDO"}


def situacao(textos):
    """True = vigente, False = encerrado, None = a planilha não diz."""
    t = norm(" ".join(textos))
    if not t:
        return None
    if any(p in t for p in INATIVO):
        return False
    if any(p in t for p in ATIVO):
        return True
    return None


def nomes_de_pessoas(texto):
    saida = []
    for parte in re.split(r"[\n;]+", publico(texto)):
        parte = re.split(r",|\(| - | – | — |:|/", parte)[0].strip(" .-")
        palavras = parte.split()
        if len(palavras) < 2 or len(parte) > 60 or re.search(r"\d", parte):
            continue
        if norm(palavras[0]) in NAO_PESSOA or not all(re.match(r"^[^\W\d_][\w'’.-]*$", p) for p in palavras):
            continue
        saida.append(" ".join(palavras))
    return saida


SIGLAS = ("SEI", "DOE", "DOU", "PGE", "TCE", "TCEPE", "CNPJ", "PNCP", "PE", "TA", "OS", "ISI", "SECTI", "MROSC", "CTI", "SAD")


def rotular(cab):
    """Cabeçalho em caixa alta vira rótulo legível, preservando siglas."""
    cab = cab.strip()
    if not cab.replace("º", "").replace("ª", "").isupper():
        return cab
    palavras = [p if p.strip("()/-,.").upper() in SIGLAS or p.upper() == "Nº" else p.lower() for p in cab.split()]
    texto = " ".join(palavras)
    return texto[:1].upper() + texto[1:]


def campos_publicos(tab, linha, lista):
    campos = []
    for item in lista:
        nome, _, rot = item.partition("|")
        i = tab.col(nome)
        if i is None:
            continue
        v = publico(Tabela.bruto(linha, i))
        if v and v not in ("-", "–") and not norm(v).startswith("FALTA"):
            campos.append([rot or rotular(tab.cab[i]), v])
    return campos


def processar_aba(cfg, tab, hoje, avisos):
    c = {k: tab.col(cfg.get(k)) for k in ("numero", "tipo", "demais", "sei", "objeto", "unidade", "projeto", "titulo", "chave")}
    c_status = [tab.col(s) for s in cfg.get("status", [])]
    c_partes = [tab.col(p) for p in cfg.get("partes", [])]
    if cfg.get("partes_regex"):
        rx = re.compile(cfg["partes_regex"])
        c_partes += [i for i, n in enumerate(tab.ncab) if rx.match(n)]
    c_partes = [i for i in c_partes if i is not None]
    c_cnpj = [i for i in (tab.col(x) for x in cfg.get("cnpj", [])) if i is not None]
    c_pessoas = {papel: tab.col(col) for papel, col in cfg.get("pessoas", {}).items()}
    modo = cfg["modo"]

    for chave, nome in list(cfg.items()):
        if chave in ("numero", "tipo", "demais", "sei", "objeto", "titulo", "chave") and nome and tab.col(nome) is None:
            avisos["colunas_sumidas"].append(f"{cfg['aba']}: coluna “{nome}” não encontrada")

    registros, atual = [], None
    for nlinha, linha in tab.linhas:
        v = lambda k: Tabela.val(linha, c[k])
        numero, demais = v("numero"), v("demais")
        filho = False
        if modo == "acessorio":
            if atual and cfg.get("itens") and not numero and not v("sei"):
                item = campos_publicos(tab, linha, cfg.get("publicas_item", []))
                if item:
                    atual["itens"].append(item)
                continue
            filho = bool(demais) and atual is not None
            if filho and numero and norm(numero) not in ("S N", "SN") and norm(numero) != norm(atual["numero"]):
                outro = next((r for r in reversed(registros) if norm(r["numero"]) == norm(numero)), None)
                if outro:
                    atual = outro
        elif modo == "consecutivo":
            tem_status = any(Tabela.val(linha, i) for i in c_status)
            filho = atual is not None and norm(v("chave")) == norm(atual["_chave"]) and not tem_status
        if filho:
            if modo == "consecutivo":
                item = campos_publicos(tab, linha, cfg.get("publicas_item", []))
                if item:
                    atual["itens"].append(item)
            else:
                atual["filhos"].append({
                    "titulo": " ".join(x for x in (v("tipo"), demais) if x),
                    "campos": campos_publicos(tab, linha, cfg["publicas"]),
                })
            for papel, i in c_pessoas.items():
                for nome in nomes_de_pessoas(Tabela.bruto(linha, i)):
                    atual["_pessoas"].append((nome, papel))
            continue

        status_txt = [Tabela.val(linha, i) for i in c_status]
        partes = [publico(Tabela.val(linha, i)) for i in c_partes]
        partes = [p for p in partes if p and p not in ("-", "–")]
        tipo = v("tipo")
        sn = norm(numero) in ("", "S N", "SN")
        if cfg.get("titulo_composto"):
            pedacos = [Tabela.val(linha, tab.col(x)) for x in cfg["titulo_composto"]]
            titulo = f"{pedacos[0]} nº {pedacos[1]}/{pedacos[2]}" if all(pedacos[:3]) else " ".join(p for p in pedacos if p)
        elif c["titulo"] is not None:
            titulo = v("titulo")
        elif sn:
            titulo = " – ".join(x for x in (tipo or cfg["singular"], v("projeto") or (partes[1] if len(partes) > 1 else "")) if x)
        else:
            titulo = f"{tipo or cfg['singular']} nº {numero}"
        atual = {
            "id": "", "aba": cfg["id"], "linha": nlinha,
            "titulo": publico(titulo) or cfg["singular"], "numero": "" if sn else numero,
            "sei": v("sei"), "objeto": publico(v("objeto")), "unidade": publico(v("unidade")).split("\n")[0][:60],
            "partes": partes, "cnpj": [Tabela.val(linha, i) for i in c_cnpj if Tabela.val(linha, i)],
            "status": publico(" / ".join(dict.fromkeys(s for s in status_txt if s))),
            "ativo": True if cfg.get("sempre_ativo") else situacao(status_txt),
            "campos": campos_publicos(tab, linha, cfg["publicas"]),
            "filhos": [], "itens": [], "prazos": [], "pessoas": [],
            "_chave": v("chave"), "_pessoas": [], "_linha": linha,
        }
        for papel, i in c_pessoas.items():
            for nome in nomes_de_pessoas(Tabela.bruto(linha, i)):
                atual["_pessoas"].append((nome, papel))
        registros.append(atual)

    # prazos, pendências e itens internos
    internos, pendencias = [], []
    for r in registros:
        linha = r.pop("_linha")
        fins = []
        for regra in cfg.get("prazos", []):
            data = None
            for col in regra["colunas"]:
                i = tab.col(col)
                bruto = Tabela.val(linha, i)
                ds = datas(bruto)
                if ds:
                    data = max(ds)
                    break
                if bruto and re.search(r"\d", bruto) and r["ativo"] is not False:
                    avisos["datas_ilegiveis"].append(f"{cfg['aba']}, linha {r['linha']} ({r['titulo']}): “{col}” = “{bruto[:40]}”")
            if not data or data.year >= 2099:
                continue
            fins.append(data)
            dias = (data - hoje).days
            ativo = r["ativo"] if r["ativo"] is not None else dias >= 0
            if (ativo and dias <= regra["antecedencia"]) or (r["ativo"] is False and 0 <= dias <= regra["antecedencia"]):
                r["prazos"].append({"tipo": regra["tipo"], "data": data.isoformat(), "dias": dias})
        if r["ativo"] is None:
            r["ativo"] = bool(fins) and max(fins) >= hoje if fins else not cfg.get("prazos")
        if r["ativo"]:
            for col in cfg.get("pendencias", []):
                i = tab.col(col)
                if "FALTA" in norm(Tabela.val(linha, i)):
                    pendencias.append(f"{r['titulo']} (linha {r['linha']}): {tab.cab[i]} consta como “{Tabela.val(linha, i)[:40]}”")
            for rot, col in cfg.get("internas", {}).items():
                valor = Tabela.val(linha, tab.col(col))
                if valor and norm(valor) not in ("NAO", "N A", "X"):
                    internos.append(f"{rot}: {r['titulo']} (linha {r['linha']}) – {valor[:120]}")
        vistos = set()
        for nome, papel in r.pop("_pessoas"):
            if (norm(nome), papel) not in vistos:
                vistos.add((norm(nome), papel))
                r["pessoas"].append([nome, papel])
        r.pop("_chave")

    # colunas com cara de prazo que nenhuma regra usa
    usadas = {tab.col(col) for regra in cfg.get("prazos", []) for col in regra["colunas"]}
    for i, n in enumerate(tab.ncab):
        if i in usadas or not n:
            continue
        if re.search(r"\b(TERMINO|FINAL|FIM|PRAZO|VENCIMENTO|GARANTIA|TITULACAO|RESOLUTIVA|PRESTACAO DE CONTAS)\b", n) and not re.search(r"\b(INICI\w*|DIAS|NATUREZA|HISTORICO|VALOR|ULTIMA TITULACAO)\b", n):
            avisos["sem_regra"].append(f"{cfg['aba']}: “{tab.cab[i]}”")
    return registros, pendencias, internos


def processar_internas(cfgs, abas, hoje, avisos):
    saida = []
    for cfg in cfgs:
        if cfg["aba"] not in abas:
            avisos["abas_sumidas"].append(cfg["aba"])
            continue
        tab = Tabela(abas[cfg["aba"]], [cfg["coluna"]] + cfg["descricao"])
        i = tab.col(cfg["coluna"])
        if i is None:
            avisos["colunas_sumidas"].append(f"{cfg['aba']}: coluna “{cfg['coluna']}” não encontrada")
            continue
        for nlinha, linha in tab.linhas:
            ds = datas(Tabela.val(linha, i))
            if not ds:
                continue
            desc = " | ".join(x for x in (Tabela.val(linha, tab.col(d)).replace("\n", " ")[:90] for d in cfg["descricao"]) if x)
            if cfg["tipo"] == "prazo":
                if situacao([Tabela.val(linha, tab.col(cfg.get("status")))]) is False:
                    continue
                dias = (max(ds) - hoje).days
                if dias <= cfg["antecedencia"]:
                    saida.append({"rotulo": cfg["rotulo"], "desc": desc, "data": max(ds), "dias": dias})
            else:
                parado = (hoje - max(ds)).days
                if parado > cfg["dias"]:
                    saida.append({"rotulo": cfg["rotulo"], "desc": desc, "data": max(ds), "dias": -parado})
    return saida


# ---------------------------------------------------------------- e-mail

def fmt(d):
    return d.strftime("%d/%m/%Y") if isinstance(d, dt.date) else dt.date.fromisoformat(d).strftime("%d/%m/%Y")


def quando(dias):
    if dias < 0:
        return f"vencido há {-dias} dia{'s' if dias != -1 else ''}"
    if dias == 0:
        return "vence hoje"
    return f"em {dias} dia{'s' if dias != 1 else ''}"


def montar_email(cfg, registros, rotulos, pendencias, internos, internas, avisos, hoje, url):
    e = html.escape
    linhas = []
    for r in registros:
        for p in r["prazos"]:
            linhas.append((p["dias"], rotulos[r["aba"]], r, p))
    linhas.sort(key=lambda x: x[0])
    faixas = [("Vencidos", lambda d: d < 0), ("Até 7 dias", lambda d: 0 <= d <= 7), ("De 8 a 30 dias", lambda d: 8 <= d <= 30),
              ("De 31 a 60 dias", lambda d: 31 <= d <= 60), ("Acima de 60 dias", lambda d: d > 60)]
    css_td = "padding:6px 8px;border-bottom:1px solid #e3e6ea;font-size:13px;vertical-align:top"
    partes = [f"<div style='font-family:Arial,Helvetica,sans-serif;color:#1c2733;max-width:860px'>",
              f"<h2 style='margin:0 0 4px'>Prazos do Jurídico – semana de {fmt(hoje)}</h2>",
              f"<p style='margin:0 0 16px;color:#5b6876;font-size:13px'>Gerado automaticamente a partir da planilha JURIDICO."
              + (f" Painel: <a href='{e(url)}'>{e(url)}</a>" if url else "") + "</p>"]
    if not linhas:
        partes.append("<p><b>Nenhum prazo dentro das antecedências configuradas.</b> Este e-mail é enviado mesmo vazio, para confirmar que a rotina está funcionando.</p>")
    for nome, teste in faixas:
        sel = [l for l in linhas if teste(l[0])]
        if not sel:
            continue
        cor = "#b42318" if nome == "Vencidos" else "#1c2733"
        partes.append(f"<h3 style='margin:18px 0 6px;color:{cor}'>{nome} ({len(sel)})</h3><table style='border-collapse:collapse;width:100%'>")
        for dias, aba, r, p in sel:
            resp = "; ".join(f"{n} ({papel})" for n, papel in r["pessoas"][:3])
            extra = " · ".join(x for x in ([r["partes"][0]] if r["partes"] and r["aba"] in ("contratos", "atas", "cgs") else []) + ([f"SEI {r['sei'].splitlines()[0]}"] if r["sei"] else []) + ([resp] if resp else []))
            partes.append(f"<tr><td style='{css_td};white-space:nowrap'><b>{fmt(p['data'])}</b><br><span style='color:#5b6876'>{quando(dias)}</span></td>"
                          f"<td style='{css_td}'><b>{e(p['tipo'])}</b> – {e(r['titulo'])} <span style='color:#5b6876'>({e(aba)}, linha {r['linha']})</span><br><span style='color:#5b6876'>{e(extra)}</span></td></tr>")
        partes.append("</table>")

    def lista(titulo, itens, nota=""):
        if not itens:
            return
        partes.append(f"<h3 style='margin:22px 0 6px'>{titulo} ({len(itens)})</h3>" + (f"<p style='margin:0 0 6px;color:#5b6876;font-size:13px'>{nota}</p>" if nota else "")
                      + "<ul style='margin:0;padding-left:18px;font-size:13px'>" + "".join(f"<li style='margin-bottom:3px'>{e(i)}</li>" for i in itens[:80])
                      + (f"<li>… e mais {len(itens) - 80}</li>" if len(itens) > 80 else "") + "</ul>")

    lista("Acompanhamento interno", [f"{i['rotulo']}: {i['desc']} – {fmt(i['data'])} ({quando(i['dias']) if i['dias'] >= 0 or 'fim' in i['rotulo'] else f'há {-i['dias']} dias'})" for i in sorted(internas, key=lambda x: x["dias"])], "Não aparece no site.")
    lista("Notas técnicas e outros itens internos", internos, "Não aparece no site.")
    lista("Cadastro ou publicação pendente", pendencias, "Instrumentos vigentes com “FALTA” em coluna de cadastro ou publicação.")
    lista("Datas que a rotina não conseguiu ler", sorted(set(avisos["datas_ilegiveis"])), "Estes prazos não estão sendo monitorados até a célula ter uma data.")
    lista("Colunas com cara de prazo sem regra de alerta", sorted(set(avisos["sem_regra"])), "Se alguma delas deve gerar alerta, peça a inclusão da regra.")
    lista("Abas novas na planilha, ainda fora do painel", avisos["abas_novas"])
    lista("Abas ou colunas que o painel esperava e não achou", avisos["abas_sumidas"] + avisos["colunas_sumidas"], "Provável renomeação na planilha.")
    partes.append("</div>")
    venc = sum(1 for l in linhas if l[0] < 0)
    em30 = sum(1 for l in linhas if 0 <= l[0] <= 30)
    assunto = f"[Painel Jurídico] Prazos da semana – {fmt(hoje)} ({venc} vencido{'s' if venc != 1 else ''}, {em30} em até 30 dias)"
    return assunto, "\n".join(partes)


def enviar(assunto, corpo, para):
    if not os.environ.get("GMAIL_USER") or not os.environ.get("GMAIL_APP_PASSWORD"):
        sys.exit("ERRO: faltam os segredos GMAIL_USER e GMAIL_APP_PASSWORD (veja o README, passo 4).")
    usuario, senha = os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"].replace(" ", "")
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = assunto, f"Painel do Jurídico <{usuario}>", ", ".join(para)
    msg.attach(MIMEText(corpo, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(usuario, senha)
        s.sendmail(usuario, para, msg.as_string())


# ---------------------------------------------------------------- principal

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fixture")
    ap.add_argument("--out", default=str(RAIZ / "site" / "data.json"))
    ap.add_argument("--email", action="store_true")
    ap.add_argument("--previa-email")
    ap.add_argument("--hoje")
    a = ap.parse_args()

    cfg = json.loads((RAIZ / "config.json").read_text(encoding="utf-8"))
    agora = dt.datetime.now(FUSO)
    hoje = dt.date.fromisoformat(a.hoje) if a.hoje else agora.date()
    nomes = [x["aba"] for x in cfg["abas"]] + [x["aba"] for x in cfg["internas"]]
    abas, existentes = ler_fixture(a.fixture) if a.fixture else ler_planilha(cfg["planilha_id"], nomes)

    avisos = defaultdict(list)
    conhecidas = set(nomes) | set(cfg["abas_ignoradas"])
    avisos["abas_novas"] = [n for n in existentes if n not in conhecidas]

    registros, pendencias, internos, lista_abas, rotulos = [], [], [], [], {}
    for c in cfg["abas"]:
        rotulos[c["id"]] = c["rotulo"]
        if c["aba"] not in abas:
            avisos["abas_sumidas"].append(c["aba"])
            continue
        esperados = c["publicas"] + [c.get(k) for k in ("numero", "tipo", "demais", "sei", "objeto")]
        tab = Tabela(abas[c["aba"]], esperados)
        regs, pend, inte = processar_aba(c, tab, hoje, avisos)
        usados = Counter()
        for r in regs:
            base = f"{c['id']}-{slug(r['numero'] or r['titulo'])[:48] or r['linha']}"
            usados[base] += 1
            r["id"] = base if usados[base] == 1 else f"{base}-{usados[base]}"
        registros += regs
        pendencias += pend
        internos += inte
        lista_abas.append({"id": c["id"], "rotulo": c["rotulo"], "singular": c["singular"], "total": len(regs), "vigentes": sum(1 for r in regs if r["ativo"])})

    internas = processar_internas(cfg["internas"], abas, hoje, avisos)

    # índice de pessoas
    grupos = defaultdict(lambda: {"formas": Counter(), "itens": []})
    for r in registros:
        for nome, papel in r["pessoas"]:
            g = grupos[norm(nome)]
            g["formas"][nome] += 1
            g["itens"].append([r["id"], papel])
    pessoas = sorted(({"id": slug(k), "nome": g["formas"].most_common(1)[0][0], "itens": g["itens"]} for k, g in grupos.items()), key=lambda p: norm(p["nome"]))

    saida = {"titulo": cfg["titulo"], "gerado_em": agora.strftime("%d/%m/%Y às %H:%M"), "hoje": hoje.isoformat(),
             "abas": lista_abas, "registros": registros, "pessoas": pessoas}
    texto = json.dumps(saida, ensure_ascii=False, separators=(",", ":"))
    for rx, nome in ((RE_CPF, "CPF"), (RE_EMAIL, "e-mail")):
        if rx.search(texto):
            sys.exit(f"ERRO: {nome} encontrado nos dados públicos; publicação cancelada.")
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(texto, encoding="utf-8")
    nprazos = sum(len(r["prazos"]) for r in registros)
    print(f"{len(registros)} registros, {len(pessoas)} pessoas, {nprazos} prazos -> {a.out} ({len(texto) // 1024} KB)")
    for chave, itens in avisos.items():
        if itens:
            print(f"aviso [{chave}]: {len(itens)}")

    if a.email or a.previa_email:
        assunto, corpo = montar_email(cfg, registros, rotulos, pendencias, internos, internas, avisos, hoje, os.environ.get("PAINEL_URL", ""))
        if a.previa_email:
            Path(a.previa_email).write_text(f"<!-- {assunto} -->\n{corpo}", encoding="utf-8")
            print("prévia do e-mail:", assunto)
        if a.email:
            enviar(assunto, corpo, cfg["email_para"])
            print("e-mail enviado para", ", ".join(cfg["email_para"]))


if __name__ == "__main__":
    main()
