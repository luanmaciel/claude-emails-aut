# -*- coding: utf-8 -*-
"""Mesa de Emails no Notion: espelha triagem/painel.json num database e da o kit de escrita que o poller usa.

Uso:
  python notion_mesa.py sync     # painel.json -> paginas/etapas no Notion (mapa em triagem/notion-mesa.json)
  python notion_mesa.py criar    # recria o database (caso de perda) e imprime os ids para colar nas constantes
  python notion_mesa.py --demo   # self-check sem rede
  python notion_mesa.py --teste  # cria uma pagina "TESTE · apagar" no database real, exercita tudo e joga na lixeira
  python notion_mesa.py temas [--simular | --criar]   # tags por palavra-chave nos cards com Tema vazio
Token: %USERPROFILE%\\.notion-token (lido so quando a API e chamada; o poller importa este modulo).
Desenho: docs/superpowers/specs/2026-09-18-mesa-emails-notion-design.md
"""
import json
import mimetypes
import os
import re
import sys
import time
import urllib.error
import urllib.request
import uuid
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from notion_resumo import rich

DB_ID = "<id-do-database-mesa>"
DS_ID = "<id-do-database-2>"
OWNER_ID = "COLOQUE-AQUI-SEU-USER-ID-DO-NOTION"
BOT_ID = "<id-do-bot-da-integracao>"
PAI = "<id-da-pagina-resumo>"  # pagina "Resumos do Dia": parent do database em criar()
VER = "2025-09-03"
TOKEN_FILE = Path.home() / ".notion-token"
REPO = Path(__file__).resolve().parent
PAINEL = REPO / "triagem" / "painel.json"
MAPA = REPO / "triagem" / "notion-mesa.json"
ETAPAS = {"mapeado": "Mapeado", "acao": "Aguardando minha ação",
          "aguardando": "Aguardando resposta deles", "encerrado": "Encerrado"}
# Tema (multi-select) por palavra-chave em remetente + assunto + resumo. So preenche quando o card nasce ou
# esta vazio; o que o Luan muda a mao fica. Lista aprovada em 19/09/2026; ajuste as regras aqui.
TEMAS = {"ALFA": {
    "Novas bibliotecas": r"librar|bibliotec|integra|implant list|exoplan|exocad|zirkonz|carestream|realguide|codiagnost|implastation|nemotec|blue ?sky|3shape|dentiq|\bstl\b|scanbody|kit ips",
    "Regulatório": r"\bfda\b|510\s?\(?k\)?|anvisa|regulat|\bmdr\b",
    "Contratos / NDA": r"\bnda\b|contrat|confidencial|agreement",
    "Centro de usinagem": r"milling|abutment|open implants|skdla|alfa ?360",
    "P&D / Protótipos": r"p&d|r&d|prot[oó]tip|relat[oó]rio de (p&d|r&d)|iolanda",
    "INPI / PI": r"\binpi\b|mentoria|patente|propriedade intelectual|pmpi",
}, "BETA": {
    "Almoxarifado": r"leonardo miranda|almoxarif|estoque|solicita[cç][aã]o de compra|\bsc\b|ruptura|painel",
    "Ferramentas": r"ferramenta|broca|inserto|turbilhona|afia[cç]|fresa|pastilha|guhring|sanposs|fungaro|madaula",
    "Compras / Cotações": r"cota[cç]|fornecedor|proposta comercial|or[cç]amento\b(?! 2027)",
    "Produção / PCP": r"programa[cç][aã]o|\bkits?\b|gabarito|produ[cç][aã]o|\bpcp\b|ordem de produ|acompanhamento|caroline rodrigues|eros napole",
    "Board / Diretoria": r"leandro rocca|\bceo\b|conselho|diretoria|\bboard\b|rateio|plano de a[cç][aã]o",
    "Controladoria / Orçamento": r"controladoria|centro de custo|custeio|or[cç]ament[aá]rio|rafael ribeiro|valida[cç][aã]o de horas|taxa.hora",
    "Qualidade / Regulatório": r"\beto\b|certificad|\bfda\b|anvisa|abimaq|cat[aá]logo|qualidade|regulat|amostras",
    "Consultoria Externa": r"consultor|mapeamento|p-excellence|\bas is\b|eumara|bpmn|sim[oõ]es pires|lei do bem",
    "T.I. / Software": r"\bt\.?i\.?\b|datasul|chamado|telas|solidshop|solidcam|software|licen[cç]a|suporte|zeiss",
}}
CORES_TEMA = {"Novas bibliotecas": "blue", "Regulatório": "purple", "Contratos / NDA": "gray",
              "Centro de usinagem": "pink", "P&D / Protótipos": "default", "INPI / PI": "brown",
              "Almoxarifado": "orange", "Ferramentas": "yellow", "Compras / Cotações": "red",
              "Produção / PCP": "green", "Board / Diretoria": "purple", "Controladoria / Orçamento": "brown",
              "Qualidade / Regulatório": "blue", "Consultoria Externa": "pink", "T.I. / Software": "gray"}
_token = None


SO_ASSUNTO = {"Board / Diretoria", "Produção / PCP", "Ferramentas"}  # palavras genericas demais para o resumo


def temas_do(card):
    """Tags do card pelas regras de TEMAS (empresa do card); [] quando nada casa (fica para o Luan)."""
    cab = " ".join(card.get(k, "") or "" for k in ("remetente", "assunto")).lower()
    tudo = cab + " " + (card.get("pedido") or "").lower()
    return [t for t, rx in TEMAS.get(card.get("empresa", ""), {}).items()
            if re.search(rx, cab if t in SO_ASSUNTO else tudo)]


class NotionErro(Exception):
    def __init__(self, code, body):
        super().__init__(f"Notion HTTP {code}: {body[:300]}")
        self.code, self.body = code, body


def api(metodo, caminho, corpo=None, raw=None, ctype="application/json"):
    """Chamada a API do Notion. Levanta NotionErro(code, body); em 429/5xx tenta ate 4x (respeita Retry-After)."""
    global _token
    _token = _token or TOKEN_FILE.read_text().strip()
    data = raw if raw is not None else (json.dumps(corpo).encode() if corpo is not None else None)
    for tentativa in range(4):
        req = urllib.request.Request("https://api.notion.com/v1" + caminho, data=data, method=metodo, headers={
            "Authorization": "Bearer " + _token, "Notion-Version": VER, "Content-Type": ctype})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            if tentativa < 3 and (e.code == 429 or e.code >= 500):
                try:
                    espera = float(e.headers.get("Retry-After", 0))
                except (TypeError, ValueError):
                    espera = 0
                time.sleep(min(max(espera, 1 + tentativa), 30))
                continue
            raise NotionErro(e.code, body) from None


# ---------- texto -> blocos ----------

def _rich_longo(texto, n=2000):
    """rich() por pedacos de <= n chars (rich() trunca em 2000). Corta em quebra de linha ou espaco quando da.

    ponytail: marcacao **negrito** cortada no meio fica literal; so em linhas > n chars.
    """
    saida = []
    while len(texto) > n:
        corte = max(texto.rfind("\n", 0, n), texto.rfind(" ", 0, n))
        if corte < 1:
            corte = n
        saida += rich(texto[:corte])
        texto = texto[corte:]
    return saida + rich(texto)


def _bloco(tipo, texto):
    # ponytail: rich_text tem teto de 100 elementos por bloco; uma linha com >100 marcacoes inline passaria disso
    if tipo == "code":
        return {"object": "block", "type": "code", "code": {"language": "plain text", "rich_text": [
            {"type": "text", "text": {"content": texto[i:i + 2000]}} for i in range(0, len(texto), 2000)]}}
    return {"object": "block", "type": tipo, tipo: {"rich_text": _rich_longo(texto)}}


def blocos_md(md):
    """Markdown -> blocos do Notion. Ignora H1 e '---'; linhas '> ' seguidas viram um quote so."""
    itens = []
    for linha in md.splitlines():
        s = linha.strip()
        if not s or s.startswith("# ") or (len(s) > 2 and set(s) <= {"-", "|", ":", " "}):
            continue
        if s.startswith("## "):
            tipo, texto = "heading_2", s[3:]
        elif s.startswith("### "):
            tipo, texto = "heading_3", s[4:]
        elif s[:2] in ("- ", "* "):
            tipo, texto = "bulleted_list_item", s[2:]
        elif re.match(r"\d+\. ", s):
            tipo, texto = "numbered_list_item", s.split(". ", 1)[1]
        elif s == ">" or s.startswith("> "):
            tipo, texto = "quote", s[2:]
        elif s.startswith("|"):
            tipo, texto = "code", s  # tabela markdown: linhas viram um bloco code (monoespacado, legivel)
        else:
            tipo, texto = "paragraph", s
        if tipo in ("quote", "code") and itens and itens[-1][0] == tipo:
            itens[-1] = (tipo, itens[-1][1] + "\n" + texto)
        else:
            itens.append((tipo, texto))
    return [_bloco(t, x) for t, x in itens]


# ---------- escrita na pagina ----------

def upload(caminho):
    """Sobe um arquivo (<= 5 MiB) e devolve o id do file_upload; None se for grande ou der erro."""
    p = Path(caminho)
    try:
        if p.stat().st_size > 20 * 1024 * 1024:  # teto single-part da API; workspace Free recusa acima de 5 MiB (vira None)
            print(f"upload pulado ({p.name}): > 20 MiB")
            return None
        ctype = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
        fu = api("POST", "/file_uploads", {"filename": p.name, "content_type": ctype})
        b = "----" + uuid.uuid4().hex
        # nome do part so ASCII (acentos no header multipart sao terreno instavel); o nome real vai no bloco
        parte = re.sub(r"[^A-Za-z0-9._-]", "_", p.name) or "upload"
        corpo = (f'--{b}\r\nContent-Disposition: form-data; name="file"; filename="{parte}"\r\n'
                 f"Content-Type: {ctype}\r\n\r\n").encode() + p.read_bytes() + f"\r\n--{b}--\r\n".encode()
        api("POST", f"/file_uploads/{fu['id']}/send", raw=corpo, ctype=f"multipart/form-data; boundary={b}")
        return fu["id"]
    except (OSError, NotionErro, KeyError) as ex:
        print(f"upload falhou ({p.name}): {ex}")
        return None


def _quando(q):
    if q is None:
        return datetime.now()
    return q if isinstance(q, datetime) else datetime.fromisoformat(q)


def etapa(page_id, titulo, md="", arquivos=(), quando=None):
    """Timeline: heading_3 '<titulo> · DD/MM HH:MM' + blocos_md(md) + arquivos; e atualiza 'Última execução'."""
    cab = f"{titulo} · {_quando(quando):%d/%m %H:%M}"
    filhos = [_bloco("heading_3", cab)] + blocos_md(md)
    fora = []  # > 5 MiB ou extensao que o Notion recusa (.stl, .step...): uma linha so, nao uma por arquivo
    for a in arquivos:
        fid = upload(a)
        if fid:
            filhos.append({"object": "block", "type": "file",
                           "file": {"type": "file_upload", "file_upload": {"id": fid}, "name": Path(a).name}})
        else:
            fora.append(Path(a))
    if fora:
        pastas = sorted({str(p.parent) for p in fora})
        filhos.append(_bloco("paragraph", f"Não anexados ({len(fora)}: extensão não aceita pelo Notion ou > 5 MiB), "
                                          f"ficam em {'; '.join(pastas)}: " + ", ".join(p.name for p in fora)))
    for i in range(0, len(filhos), 100):  # limite da API por request
        api("PATCH", f"/blocks/{page_id}/children", {"children": filhos[i:i + 100]})
    set_props(page_id, ultima=cab)


def _rt(texto):
    return {"rich_text": [{"text": {"content": texto[:2000]}}] if texto else []}


def _data(iso):
    """ISO local sem fuso (painel.json) -> ISO com fuso, como o Notion pede."""
    return datetime.fromisoformat(iso).astimezone().isoformat(timespec="seconds")


def set_props(page_id, etapa=None, rascunho=None, respondido=None, ultima=None):
    """PATCH so das propriedades passadas. Rascunho > 2000 chars e truncado com ' … (ver timeline)'."""
    props = {}
    if etapa is not None:
        props["Etapa"] = {"select": {"name": etapa}}
    if rascunho is not None:
        props["Rascunho"] = _rt(rascunho if len(rascunho) <= 2000 else rascunho[:1983] + " … (ver timeline)")
    if respondido is not None:
        props["Respondido"] = {"date": {"start": _data(respondido)}}
    if ultima is not None:
        props["Última execução"] = _rt(ultima)
    if props:
        api("PATCH", f"/pages/{page_id}", {"properties": props})


# ---------- comentarios e consultas ----------

def _corpo_comentario(page_id, texto, discussion_id=None):
    corpo = {"rich_text": _rich_longo(texto or "(vazio)", 1900)}
    if discussion_id:
        corpo["discussion_id"] = discussion_id
    else:
        corpo["parent"] = {"page_id": page_id}
    return corpo


def comentar(page_id, texto, discussion_id=None):
    """Comenta na pagina ou responde no fio (discussion_id). 403 = integracao sem capacidade de comentarios."""
    return api("POST", "/comments", _corpo_comentario(page_id, texto, discussion_id))


def _paginado(metodo, caminho, corpo=None):
    cursor = None
    while True:
        if metodo == "GET":
            r = api("GET", caminho + (f"&start_cursor={cursor}" if cursor else ""))
        else:
            r = api("POST", caminho, {**corpo, "start_cursor": cursor} if cursor else corpo)
        yield from r.get("results", [])
        cursor = r.get("next_cursor") if r.get("has_more") else None
        if not cursor:
            return


def listar_comentarios(page_id):
    """Comentarios da pagina em ordem cronologica: [{id, discussion_id, created_time, created_by, texto}]."""
    lista = [{"id": c["id"], "discussion_id": c.get("discussion_id"), "created_time": c.get("created_time", ""),
              "created_by": (c.get("created_by") or {}).get("id"),
              "texto": "".join(t.get("plain_text", "") for t in c.get("rich_text", []))}
             for c in _paginado("GET", f"/comments?block_id={page_id}&page_size=100")]
    return sorted(lista, key=lambda c: c["created_time"])


def _prop(pagina, nome):
    v = pagina.get("properties", {}).get(nome) or {}
    t = v.get("type")
    if t == "select":
        return (v["select"] or {}).get("name", "")
    if t in ("rich_text", "title"):
        return "".join(x.get("plain_text", "") for x in v[t])
    return ""


def paginas_ativas():
    """Paginas com Etapa != Encerrado: [{page_id, etapa, ref}]."""
    filtro = {"filter": {"property": "Etapa", "select": {"does_not_equal": "Encerrado"}}, "page_size": 100}
    return [{"page_id": p["id"], "etapa": _prop(p, "Etapa"), "ref": _prop(p, "ref")}
            for p in _paginado("POST", f"/data_sources/{DS_ID}/query", filtro)]


def achar_por_ref(ref):
    """Pagina cujo ref contem o id (thread continuada acumula os ids dos emails na mesma pagina)."""
    r = api("POST", f"/data_sources/{DS_ID}/query",
            {"filter": {"property": "ref", "rich_text": {"contains": ref}}, "page_size": 1})
    return r["results"][0]["id"] if r.get("results") else None


# ---------- mapa local ----------

def mapa_ler():
    try:
        m = json.loads(MAPA.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        m = {}
    m.setdefault("owner_id", OWNER_ID)
    m.setdefault("ultimo_sync", "")
    m.setdefault("cards", {})
    return m


def mapa_gravar(m):
    MAPA.parent.mkdir(exist_ok=True)
    tmp = MAPA.with_suffix(".tmp")
    tmp.write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, MAPA)


# ---------- sync ----------

def limpar_assunto(assunto):
    return re.sub(r"^((re|res|fw|fwd|enc)\s*:\s*)+", "", (assunto or "").strip(), flags=re.I).strip()


def nome_pagina(card):
    return f"{card.get('remetente', '')} · {card.get('termo') or limpar_assunto(card.get('assunto', ''))}"


def _linha(s):
    return " ".join((s or "").split())


def triagem_md(card):
    return (f"- Pedido: {_linha(card.get('pedido')) or '(sem resumo)'}\n"
            f"- Ação sugerida: {_linha(card.get('acao')) or '(sem ação sugerida)'}")


def _mtime(caminho):
    try:
        return os.path.getmtime(REPO / caminho) if caminho else 0
    except OSError:
        return 0


def _ler(caminho):
    try:
        return (REPO / caminho).read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return ""


def _agora_utc():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")  # mesmo formato do created_time


def nova_entrada(card):
    return {"page_id": "", "ref": card["id"], "conversa": card.get("conversa", ""), "slug": card.get("slug", ""),
            "sessao": card.get("sessao", ""), "termo": card.get("termo", ""), "dossie": card.get("dossie", ""),
            "empresa": card.get("empresa", ""), "coluna_vista": "", "dossie_mtime": 0, "rascunho_mtime": 0,
            "rascunho_v": 0, "respondido_visto": "", "comentarios_ate": _agora_utc(), "preview": None}


def decidir(card, entrada, dossie_mtime, rascunho_mtime, pagina_thread=None):
    """Decisao pura do sync (regras 1 a 4 da spec): card + entrada do mapa + mtimes -> lista de acoes, sem rede.

    entrada None = card novo: ('continuar', page_id) se outra entrada tem a mesma conversa, senao ('criar',)
    (quem executa tenta adotar a pagina por ref antes de criar). Depois: ('etapa', coluna) so se coluna_auto
    mudou desde a ultima vista e o Luan nao encerrou o card pela Mesa ('encerrado' na entrada),
    ('dossie', mtime), ('rascunho', v, mtime), ('respondido', iso).
    """
    e = entrada or {}
    acoes = []
    if entrada is None:
        acoes.append(("continuar", pagina_thread) if pagina_thread else ("criar",))
    elif card.get("coluna_auto") != e.get("coluna_vista") and not e.get("encerrado"):
        acoes.append(("etapa", card.get("coluna_auto", "mapeado")))
    if dossie_mtime > e.get("dossie_mtime", 0):
        acoes.append(("dossie", dossie_mtime))
    if rascunho_mtime > e.get("rascunho_mtime", 0):
        acoes.append(("rascunho", e.get("rascunho_v", 0) + 1, rascunho_mtime))
    resp = card.get("respondido", "")
    if resp and resp != e.get("respondido_visto", ""):
        acoes.append(("respondido", resp))
    return acoes


def criar_pagina(card):
    props = {"Nome": {"title": [{"text": {"content": nome_pagina(card)[:2000]}}]},
             "Empresa": {"select": {"name": card.get("empresa") or "ALFA"}},
             "Etapa": {"select": {"name": ETAPAS.get(card.get("coluna_auto"), "Mapeado")}},
             "Remetente": _rt(card.get("remetente", "")), "Resumo": _rt(card.get("pedido", "")),
             "Ação proposta": _rt(card.get("acao", "")), "ref": _rt(card["id"])}
    if card.get("recebido"):
        props["Recebido"] = {"date": {"start": _data(card["recebido"])}}
    if temas_do(card):
        props["Tema"] = {"multi_select": [{"name": t} for t in temas_do(card)]}
    return api("POST", "/pages", {"parent": {"type": "data_source_id", "data_source_id": DS_ID},
                                  "properties": props})["id"]


def sync(mapa=None):
    """Espelha triagem/painel.json no Notion (regras da spec) e devolve o mapa atualizado."""
    mapa = mapa if mapa is not None else mapa_ler()
    cards = mapa["cards"]
    painel = json.loads(PAINEL.read_text(encoding="utf-8-sig"))
    n = Counter()
    # do mais antigo ao mais novo: numa thread, o primeiro cria a pagina e os seguintes continuam nela
    for card in sorted(painel.get("cards", []), key=lambda c: c.get("recebido", "")):
        eid = card["entry_id"]
        entrada = cards.get(eid)
        dossie = card.get("dossie", "")
        rasc = dossie.replace(".md", "-rascunho.txt") if dossie else ""
        dm, rm = _mtime(dossie), _mtime(rasc)
        thread = None
        if entrada is None and card.get("conversa"):
            thread = next((v for k, v in cards.items()
                           if k != eid and v.get("page_id") and v.get("conversa") == card["conversa"]), None)
        e = entrada if entrada is not None else nova_entrada(card)
        try:
            for a in decidir(card, entrada, dm, rm, thread and thread["page_id"]):
                if a[0] == "criar":
                    pid = achar_por_ref(card["id"])
                    if pid:  # ponytail: pagina adotada apos perda do mapa: nao repete a timeline nem mexe na Etapa
                        e.update(page_id=pid, dossie_mtime=dm, rascunho_mtime=rm,
                                 respondido_visto=card.get("respondido", ""), coluna_vista=card.get("coluna_auto", ""))
                        n["adotados"] += 1
                        break
                    e["page_id"] = criar_pagina(card)
                    e["coluna_vista"] = card.get("coluna_auto", "")
                    etapa(e["page_id"], "🔎 Triagem", triagem_md(card), quando=card.get("recebido"))
                    n["criados"] += 1
                elif a[0] == "continuar":
                    # ponytail: rascunho_v recomeca em 1 no email novo; ref acumula os ids da thread (adocao
                    # apos perda do mapa acha a pagina por qualquer um deles)
                    e.update(page_id=a[1], comentarios_ate=thread.get("comentarios_ate", ""),
                             coluna_vista=card.get("coluna_auto", ""))
                    if thread.get("encerrado"):  # card encerrado pela Mesa: registra a resposta, nao reabre
                        e["encerrado"] = True
                    etapa(a[1], "📥 Resposta deles", triagem_md(card), quando=card.get("recebido"))
                    refs = [v["ref"] for v in cards.values() if v.get("page_id") == a[1]] + [card["id"]]
                    props = {"ref": _rt(" ".join(dict.fromkeys(refs)))}
                    if not e.get("encerrado"):
                        props["Etapa"] = {"select": {"name": "Aguardando minha ação"}}
                    api("PATCH", f"/pages/{a[1]}", {"properties": props})
                    n["threads"] += 1
                elif a[0] == "etapa":
                    set_props(e["page_id"], etapa=ETAPAS.get(a[1], "Mapeado"))
                    e["coluna_vista"] = a[1]
                    n["etapas"] += 1
                elif a[0] == "dossie":
                    etapa(e["page_id"], "📄 Dossiê", _ler(dossie), arquivos=[REPO / dossie])
                    e["dossie_mtime"] = a[1]
                    n["dossies"] += 1
                elif a[0] == "rascunho":
                    txt = _ler(rasc).strip()
                    etapa(e["page_id"], f"✍️ Rascunho v{a[1]}",
                          "\n".join("> " + l for l in txt.splitlines()), arquivos=[REPO / rasc])
                    set_props(e["page_id"], rascunho=txt)
                    e["rascunho_v"], e["rascunho_mtime"] = a[1], a[2]
                    n["rascunhos"] += 1
                elif a[0] == "respondido":
                    set_props(e["page_id"], respondido=a[1])
                    if e.get("respondido_visto") != "poller":  # enviado pela Mesa: a etapa "Enviado" ja esta la
                        etapa(e["page_id"], "✅ Respondido pelo Outlook", quando=a[1])
                    e["respondido_visto"] = a[1]
        except NotionErro as ex:
            print(f"erro no card {card.get('slug') or card['id']}: {ex}")
            n["erros"] += 1
            if entrada is not None and (ex.code == 404 or "archived" in ex.body or "trash" in ex.body):
                cards.pop(eid, None)  # pagina apagada no Notion: o proximo sync recria
                mapa_gravar(mapa)
                continue
        if not e.get("page_id"):
            continue  # criar falhou: tenta de novo no proximo sync
        for k in ("conversa", "slug", "sessao", "termo", "dossie", "empresa"):
            e[k] = card.get(k) or e.get(k, "")  # sessao/termo chegam depois (rotina das 13h): sempre refrescar
        cards[eid] = e
        mapa_gravar(mapa)  # por card: se o run cair, o que ja subiu nao sobe de novo
    mapa["ultimo_sync"] = datetime.now().isoformat(timespec="minutes")
    mapa_gravar(mapa)
    print(f"sync: {n['criados']} criados, {n['adotados']} adotados, {n['threads']} threads continuadas, "
          f"{n['dossies']} dossies, {n['rascunhos']} rascunhos, {n['etapas']} etapas movidas, {n['erros']} erros")
    return mapa


# ---------- temas ----------

def criar_tema():
    """Cria (ou recria a lista de opcoes d)a propriedade Tema no data source. Nao toca nas outras propriedades."""
    opcoes = [{"name": t, "color": CORES_TEMA.get(t, "default")} for emp in ("ALFA", "BETA") for t in TEMAS[emp]]
    api("PATCH", f"/data_sources/{DS_ID}", {"properties": {"Tema": {"multi_select": {"options": opcoes}}}})
    print(f"propriedade Tema com {len(opcoes)} opções")


def preencher_temas(simular=False):
    """Preenche Tema nos cards do mapa cujo Tema esta vazio (nunca sobrescreve). --simular so imprime."""
    mapa = mapa_ler()
    por_eid = {c["entry_id"]: c for c in json.loads(PAINEL.read_text(encoding="utf-8-sig")).get("cards", [])}
    n = Counter()
    for eid, e in mapa["cards"].items():
        card = por_eid.get(eid)
        if not card or not e.get("page_id"):
            continue
        temas = temas_do(card)
        print(f"[{card['empresa']}] {card['remetente'][:22]:22} | {card['assunto'][:45]:45} -> {', '.join(temas) or '(sem tag)'}")
        if simular or not temas:
            n["sem tag" if not temas else "simulados"] += 1
            continue
        pg = api("GET", f"/pages/{e['page_id']}")
        if (pg.get("properties", {}).get("Tema") or {}).get("multi_select"):
            n["ja tinham"] += 1
            continue
        api("PATCH", f"/pages/{e['page_id']}", {"properties": {"Tema": {"multi_select": [{"name": t} for t in temas]}}})
        n["preenchidos"] += 1
    print("temas:", dict(n))


# ---------- criar database ----------

def criar():
    """Recria o database Mesa de Emails (schema da sonda de 18/09) e imprime os ids para as constantes."""
    def sel(*nomes):
        return {"select": {"options": [{"name": x} for x in nomes]}}
    props = {
        "Nome": {"title": {}},
        "Empresa": sel("ALFA", "BETA"),
        "Etapa": sel("Mapeado", "Aguardando minha ação", "Executando", "Confirmar envio",
                     "Aguardando resposta deles", "Encerrado", "Erro"),
        "Remetente": {"rich_text": {}}, "Resumo": {"rich_text": {}}, "Ação proposta": {"rich_text": {}},
        "Rascunho": {"rich_text": {}}, "Recebido": {"date": {}}, "Respondido": {"date": {}},
        "Dias": {"formula": {"expression": 'if(empty(prop("Respondido")), dateBetween(now(), prop("Recebido"), "days"), '
                                           'dateBetween(now(), prop("Respondido"), "days"))'}},
        "Última execução": {"rich_text": {}}, "ref": {"rich_text": {}},
    }
    db = api("POST", "/databases", {
        "parent": {"type": "page_id", "page_id": PAI},
        "title": [{"type": "text", "text": {"content": "Mesa de Emails"}}],
        "icon": {"type": "emoji", "emoji": "📬"},
        "initial_data_source": {"properties": props},
    })
    print(f'DB_ID = "{db["id"]}"\nDS_ID = "{db["data_sources"][0]["id"]}"\nurl: {db.get("url")}')
    print("Cole os ids nas constantes de notion_mesa.py e apague triagem/notion-mesa.json.")


# ---------- self-checks ----------

def demo():
    md = ("# H1 fora\n\n## Sec\n### Sub\n- um **b**\n* dois\n1. tres\n> q1\n>\n> q2\n---\ntexto `c`\n"
          + "x" * 4500 + "\n| a | b |\n|---|---|\n| 1 | 2 |\n")
    b = blocos_md(md)
    assert [x["type"] for x in b] == ["heading_2", "heading_3", "bulleted_list_item", "bulleted_list_item",
                                      "numbered_list_item", "quote", "paragraph", "paragraph", "code"], [x["type"] for x in b]
    assert b[8]["code"]["rich_text"][0]["text"]["content"] == "| a | b |\n| 1 | 2 |"
    assert b[0]["heading_2"]["rich_text"][0]["text"]["content"] == "Sec"
    assert b[2]["bulleted_list_item"]["rich_text"][1] == {"type": "text", "text": {"content": "b"},
                                                          "annotations": {"bold": True}}
    assert b[4]["numbered_list_item"]["rich_text"][0]["text"]["content"] == "tres"
    assert b[5]["quote"]["rich_text"][0]["text"]["content"] == "q1\n\nq2"
    partes = b[7]["paragraph"]["rich_text"]
    assert len(partes) == 3 and all(len(p["text"]["content"]) <= 2000 for p in partes)
    assert "".join(p["text"]["content"] for p in partes) == "x" * 4500
    # comentar: fatias de <= 1900 sem perder texto; fio via discussion_id
    c = _corpo_comentario("pg", "a " * 2000)
    assert c["parent"] == {"page_id": "pg"} and all(len(p["text"]["content"]) <= 1900 for p in c["rich_text"])
    assert "".join(p["text"]["content"] for p in c["rich_text"]) == "a " * 2000
    fio = _corpo_comentario("pg", "oi", "disc-1")
    assert fio["discussion_id"] == "disc-1" and "parent" not in fio
    # mapeamento e nome
    assert ETAPAS == {"mapeado": "Mapeado", "acao": "Aguardando minha ação",
                      "aguardando": "Aguardando resposta deles", "encerrado": "Encerrado"}
    assert limpar_assunto("RES: Re: FW: Pedido 123") == "Pedido 123"
    assert nome_pagina({"remetente": "Fulano", "assunto": "Re: Orçamento"}) == "Fulano · Orçamento"
    assert nome_pagina({"remetente": "Fulano", "assunto": "Re: Orçamento", "termo": "Orc"}) == "Fulano · Orc"
    # decisao pura do sync
    card = {"id": "abc", "entry_id": "x", "coluna_auto": "acao", "respondido": "", "conversa": "c1"}
    assert decidir(card, None, 0, 0) == [("criar",)]
    assert decidir(card, None, 0, 0, "pg-1") == [("continuar", "pg-1")]
    assert decidir(card, None, 10, 20) == [("criar",), ("dossie", 10), ("rascunho", 1, 20)]
    e = {"coluna_vista": "acao", "dossie_mtime": 10, "rascunho_mtime": 20, "rascunho_v": 1, "respondido_visto": ""}
    assert decidir(card, e, 10, 20) == []
    assert decidir(card, e, 11, 20) == [("dossie", 11)]
    assert decidir(card, e, 10, 25) == [("rascunho", 2, 25)]
    assert decidir({**card, "coluna_auto": "aguardando", "respondido": "2026-09-17T10:40"}, e, 10, 20) == [
        ("etapa", "aguardando"), ("respondido", "2026-09-17T10:40")]
    # encerrado pela Mesa: a coluna automatica mudou, mas a Etapa fica; o resto continua subindo
    assert decidir({**card, "coluna_auto": "aguardando", "respondido": "2026-09-17T10:40"},
                   {**e, "encerrado": True}, 11, 20) == [("dossie", 11), ("respondido", "2026-09-17T10:40")]
    assert set_props.__defaults__ == (None, None, None, None)
    # temas por palavra-chave (nomes ficticios)
    assert temas_do({"empresa": "ALFA", "remetente": "X", "assunto": "Alfa Implant Library Integration", "pedido": ""}) == ["Novas bibliotecas"]
    assert temas_do({"empresa": "ALFA", "remetente": "X", "assunto": "MT 16 library for FDA", "pedido": ""}) == ["Novas bibliotecas", "Regulatório"]
    assert temas_do({"empresa": "BETA", "remetente": "Leonardo Miranda", "assunto": "Inserto 3PKT", "pedido": ""}) == ["Almoxarifado", "Ferramentas"]
    assert temas_do({"empresa": "BETA", "remetente": "X", "assunto": "Planejamento Orçamentário 2027", "pedido": "Controladoria pede"}) == ["Controladoria / Orçamento"]
    assert temas_do({"empresa": "BETA", "remetente": "X", "assunto": "Orçamento referente ao porta ferramenta", "pedido": ""}) == ["Ferramentas", "Compras / Cotações"]
    assert temas_do({"empresa": "BETA", "remetente": "X", "assunto": "Assunto neutro", "pedido": "nada"}) == []
    assert all(t in CORES_TEMA for emp in TEMAS for t in TEMAS[emp])
    print("demo ok")


def teste():
    """Rede: pagina de teste no database real, todas as chamadas, e lixeira no final."""
    import tempfile
    tmp = Path(tempfile.gettempdir()) / "mesa-teste.txt"
    tmp.write_text("arquivo de teste\n", encoding="utf-8")
    tmp_md = Path(tempfile.gettempdir()) / "mesa-teste-validação.md"  # nome com acento: cobre o multipart
    tmp_md.write_text("# teste\n\ntexto\n", encoding="utf-8")
    card = {"id": "teste-ref-mesa", "entry_id": "teste", "empresa": "ALFA", "remetente": "Fulano de Tal",
            "assunto": "RES: TESTE · apagar", "pedido": "pedido de teste", "acao": "acao de teste",
            "coluna_auto": "acao", "recebido": "2026-09-16T07:32"}
    pid = criar_pagina(card)
    print("pagina:", pid)
    try:
        etapa(pid, "🔎 Triagem", triagem_md(card), quando=card["recebido"])
        etapa(pid, "📄 Dossiê", "## Pedido\n- item **forte**\n1. passo\n> citacao\n>\n> segunda\n\nparagrafo `codigo`",
              arquivos=[tmp, tmp_md])
        print("upload .txt:", upload(tmp), "| upload .md:", upload(tmp_md))
        set_props(pid, etapa="Executando", rascunho="r" * 2500, respondido="2026-09-17T10:40", ultima="⏳ teste (10:02)")
        print("set_props ok")
        for nome, fn in (("comentar", lambda: comentar(pid, "teste **negrito** " + "z" * 2000)),
                         ("listar_comentarios", lambda: listar_comentarios(pid))):
            try:
                print(nome + ":", str(fn())[:160])
            except NotionErro as ex:
                print(f"{nome}: HTTP {ex.code} (esperado enquanto a integracao nao tem capacidade de comentarios)")
        ativas = paginas_ativas()
        print(f"paginas_ativas: {len(ativas)} | teste presente: {any(p['page_id'] == pid for p in ativas)}")
        print("achar_por_ref:", achar_por_ref(card["id"]) == pid)
    finally:
        r = api("PATCH", f"/pages/{pid}", {"in_trash": True})
        print("lixeira:", r.get("in_trash"))
    print("teste ok")


def main(argv):
    sys.stdout.reconfigure(encoding="utf-8")
    cmd = argv[1] if len(argv) > 1 else ""
    if cmd == "sync":
        sync()
    elif cmd == "criar":
        criar()
    elif cmd == "temas":  # temas --criar (propriedade) | temas --simular (so imprime) | temas (preenche vazios)
        criar_tema() if "--criar" in argv else preencher_temas(simular="--simular" in argv)
    elif cmd == "--demo":
        demo()
    elif cmd == "--teste":
        teste()
    else:
        print(__doc__)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
