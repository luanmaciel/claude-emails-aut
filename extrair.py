# -*- coding: utf-8 -*-
"""Extrai emails das caixas de entrada do Outlook (MAPI/COM) para JSON.

Uso: python extrair.py [dias]   (default: 7 dias; saida: emails.json)
Sem LLM. Pre-filtro deterministico por regras (ver filtrar()).
"""
import json
import re
import sys

import win32com.client

DIAS_PADRAO = 7
SAIDA = "emails.json"

RE_NOREPLY = re.compile(r"no.?reply|donotreply|newsletter|mailer-daemon", re.I)
# disparos automaticos proprios a ignorar: um endereco por linha em ignorados.txt (fora do git)
import os
REMETENTES_IGNORADOS = set()
if os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ignorados.txt")):
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ignorados.txt"), encoding="utf-8") as _f:
        REMETENTES_IGNORADOS = {l.strip().lower() for l in _f if l.strip()}
RE_SISTEMAS = re.compile(r"datalus|totvs|fluig", re.I)
RE_APROVACAO = re.compile(r"aprovaç|aprovacao|pendente de aprova|approval", re.I)
RE_UNSUBSCRIBE = re.compile(r"unsubscribe|cancelar (a )?inscrição|descadastr", re.I)

# Tag automática "Bibliotecas" (só conta SIN). Única escrita no Outlook desta
# rotina: campo Categories, nada mais. Newsletters ficam de fora via RE_NOREPLY/UNSUBSCRIBE.
TAG_BIBLIO = "Bibliotecas"
RE_BIBLIO = re.compile(
    r"library|libraries|bibliotec|3shape|exocad|codiagnostix|realguide|blue ?sky|"
    r"implastation|zimvie|medit|tibase|dentiq|atomica|yomi|sicat|nemotec|dtx", re.I)
RE_BIBLIO_REMETENTE = re.compile(
    r"exocad\.com|dental-wings\.com|blueskyplan|blueskybio|prodigident|zimvie|"
    r"3shape|meditcompany|sin360|sindentalusa", re.I)


def marcar_biblioteca(item, email):
    """Aplica a categoria Bibliotecas em emails da SIN sobre bibliotecas digitais."""
    if "sinimplantsystem" not in email["conta"]:
        return
    if not (RE_BIBLIO.search(email["assunto"]) or RE_BIBLIO_REMETENTE.search(email["remetente_smtp"])):
        return
    if RE_NOREPLY.search(email["remetente_smtp"]) or RE_UNSUBSCRIBE.search(email["corpo"].lower()):
        return  # marketing/newsletter não ganha tag
    cats = [c.strip() for c in (email["categoria"] or "").split(",") if c.strip()]
    if TAG_BIBLIO in cats:
        return
    cats.append(TAG_BIBLIO)
    try:
        item.Categories = ", ".join(cats)
        item.Save()
        email["categoria"] = item.Categories
    except Exception:
        pass


def smtp_do_item(item):
    """SMTP do remetente (resolve endereco EX do Exchange)."""
    try:
        if item.SenderEmailType == "EX":
            exch = item.Sender.GetExchangeUser()
            if exch:
                return exch.PrimarySmtpAddress or item.SenderEmailAddress
        return item.SenderEmailAddress or ""
    except Exception:
        return ""


def smtp_do_recipient(r):
    try:
        if r.AddressEntry.Type == "EX":
            exch = r.AddressEntry.GetExchangeUser()
            if exch:
                return exch.PrimarySmtpAddress or r.Address
        return r.Address or ""
    except Exception:
        return ""


def iterar_recentes(pasta, cutoff):
    """Itera itens da pasta do mais novo ao mais velho, parando no cutoff.

    Evita Items.Restrict porque o formato de data do Restrict depende do
    locale do Windows — ordenar e cortar é à prova de locale.
    """
    items = pasta.Items
    items.Sort("[ReceivedTime]", True)
    for item in items:
        try:
            recebido = item.ReceivedTime
        except Exception:
            continue
        if recebido is None:
            continue
        if recebido.replace(tzinfo=None) < cutoff:
            break
        if item.Class == 43:  # olMail
            yield item


def mapa_ultimas_respostas(store, cutoff):
    """ConversationID -> data do meu envio mais recente (Itens Enviados)."""
    mapa = {}
    try:
        enviados = store.GetDefaultFolder(5)  # olFolderSentMail
    except Exception:
        return mapa
    for item in iterar_recentes(enviados, cutoff):
        try:
            conv = item.ConversationID
        except Exception:
            continue
        dt = item.ReceivedTime.replace(tzinfo=None)
        if conv and (conv not in mapa or dt > mapa[conv]):
            mapa[conv] = dt
    return mapa


def filtrar(email, meus_enderecos, meu_nome, ultima_minha):
    """Retorna motivo do descarte ou None (mantém). Regras do CLAUDE.md."""
    corpo = email["corpo"].lower()

    if email["remetente_smtp"].lower() in REMETENTES_IGNORADOS:
        return "remetente_ignorado"

    if RE_NOREPLY.search(email["remetente_smtp"]) or RE_UNSUBSCRIBE.search(corpo):
        return "newsletter_noreply"

    if ultima_minha and ultima_minha >= email["_recebido_dt"]:
        return "ultima_mensagem_minha"

    if RE_SISTEMAS.search(email["remetente_smtp"]) or RE_SISTEMAS.search(email["assunto"]):
        if not (RE_APROVACAO.search(email["assunto"]) or RE_APROVACAO.search(corpo)):
            return "notificacao_sistema"

    to_lower = [t.lower() for t in email["para"]]
    estou_no_para = any(e in t for e in meus_enderecos for t in to_lower)
    if not estou_no_para:
        partes_nome = [p for p in meu_nome.lower().split() if len(p) > 2]
        if not any(p in corpo for p in partes_nome):
            return "cc_sem_mencao"

    return None


def extrair_conta(store, conta_smtp, meu_nome, cutoff):
    inbox = store.GetDefaultFolder(6)  # olFolderInbox
    ultimas = mapa_ultimas_respostas(store, cutoff)
    meus_enderecos = [conta_smtp.lower()]

    mantidos, descartados = [], {}
    for item in iterar_recentes(inbox, cutoff):
        para, cc = [], []
        try:
            for r in item.Recipients:
                (para if r.Type == 1 else cc).append(smtp_do_recipient(r))
        except Exception:
            pass
        try:
            conv = item.ConversationID
        except Exception:
            conv = ""
        recebido_dt = item.ReceivedTime.replace(tzinfo=None)
        ultima_minha = ultimas.get(conv)
        email = {
            "remetente": item.SenderName or "",
            "remetente_smtp": smtp_do_item(item),
            "para": para,
            "cc": cc,
            "assunto": item.Subject or "",
            "recebido": recebido_dt.isoformat(),
            "ultima_resposta_minha": ultima_minha.isoformat() if ultima_minha else None,
            "flag": item.FlagRequest or "",
            "categoria": item.Categories or "",
            "tem_anexo": item.Attachments.Count > 0,
            "corpo": (item.Body or "")[:500],
            "entry_id": item.EntryID,
            "conta": conta_smtp,
            "_recebido_dt": recebido_dt,
        }
        marcar_biblioteca(item, email)
        motivo = filtrar(email, meus_enderecos, meu_nome, ultima_minha)
        if motivo:
            descartados[motivo] = descartados.get(motivo, 0) + 1
        else:
            del email["_recebido_dt"]
            mantidos.append(email)
    return mantidos, descartados


def main():
    import datetime

    dias = int(sys.argv[1]) if len(sys.argv) > 1 else DIAS_PADRAO
    global SAIDA
    if len(sys.argv) > 2:
        SAIDA = sys.argv[2]
    cutoff = datetime.datetime.now() - datetime.timedelta(days=dias)

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # mapa store -> smtp da conta dona
    contas = {}
    for acc in ns.Accounts:
        try:
            contas[acc.DeliveryStore.StoreID] = acc.SmtpAddress
        except Exception:
            pass

    meu_nome = ns.CurrentUser.Name or ""
    resultado = {"gerado_em": datetime.datetime.now().isoformat(), "contas": []}

    for store in ns.Stores:
        smtp = contas.get(store.StoreID)
        if not smtp:
            continue  # arquivos de dados locais, caixas compartilhadas etc.
        print(f"Lendo {smtp} ({store.DisplayName})...", file=sys.stderr)
        mantidos, descartados = extrair_conta(store, smtp, meu_nome, cutoff)
        resultado["contas"].append({
            "conta": smtp,
            "emails": mantidos,
            "descartados": descartados,
        })
        print(f"  mantidos: {len(mantidos)}, descartados: {descartados}", file=sys.stderr)

    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
    print(f"OK: {SAIDA}", file=sys.stderr)


if __name__ == "__main__":
    main()
