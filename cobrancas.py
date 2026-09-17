# -*- coding: utf-8 -*-
"""Levanta threads paradas: eu mandei, ninguem respondeu.

Uso: python cobrancas.py [dias_janela] [dias_minimos]   (default: 60 e 4)
Saida: cobrancas.json. Sem LLM, sem escrita no Outlook.

Complementa extrair.py, que por desenho descarta justamente estas threads
(filtrar() -> "ultima_mensagem_minha"): a triagem diaria olha o que chegou,
esta rotina olha o que saiu e nao voltou.
"""
import datetime
import json
import re
import sys

import win32com.client

DIAS_PADRAO = 60      # janela de varredura dos Itens Enviados
MIN_DIAS_PADRAO = 4   # so reporta se o envio ja tem esta idade
SAIDA = "cobrancas.json"

RE_NOREPLY = re.compile(r"no.?reply|donotreply|newsletter|mailer-daemon", re.I)
RE_AUTO = re.compile(r"automatic reply|resposta automática|out of office|ausência do escritório", re.I)


def data_do_item(item):
    """SentOn nos enviados, ReceivedTime nos recebidos; None se nenhum servir."""
    for campo in ("SentOn", "ReceivedTime"):
        try:
            dt = getattr(item, campo)
        except Exception:
            continue
        if dt is not None:
            try:
                return dt.replace(tzinfo=None)
            except Exception:
                continue
    return None


def iterar_recentes(pasta, cutoff):
    """Itens de mail mais novos que cutoff, do mais novo para o mais antigo.

    Nao usa Restrict: o filtro por data depende do locale do Windows (ver README).
    """
    items = pasta.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    for item in items:
        dt = data_do_item(item)
        if dt is None:
            continue
        if dt < cutoff:
            break
        try:
            if item.Class != 43:  # olMail
                continue
        except Exception:
            continue
        yield item, dt


def pastas_de_entrada(inbox):
    """Inbox e suas subpastas (BioHorizon e afins ficam fora da Inbox raiz)."""
    pilha, saida = [inbox], []
    while pilha:
        pasta = pilha.pop()
        saida.append(pasta)
        try:
            pilha.extend(list(pasta.Folders))
        except Exception:
            pass
    return saida


def mapa_recebidos(store, cutoff):
    """ConversationID -> data da mensagem recebida mais recente."""
    mapa = {}
    try:
        inbox = store.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        return mapa
    for pasta in pastas_de_entrada(inbox):
        for item, dt in iterar_recentes(pasta, cutoff):
            try:
                conv = item.ConversationID
            except Exception:
                continue
            if conv and dt > mapa.get(conv, cutoff - datetime.timedelta(days=1)):
                mapa[conv] = dt
    return mapa


def enviados_por_thread(store, cutoff):
    """ConversationID -> dados do meu envio mais recente na thread."""
    threads = {}
    try:
        enviados = store.GetDefaultFolder(5)  # olFolderSentMail
    except Exception:
        return threads
    for item, dt in iterar_recentes(enviados, cutoff):
        try:
            conv = item.ConversationID
        except Exception:
            continue
        if not conv:
            continue
        anterior = threads.get(conv)
        if anterior and anterior["_dt"] >= dt:
            continue
        para, cc = [], []
        try:
            for r in item.Recipients:
                endereco = smtp_do_recipient(r)
                (para if r.Type == 1 else cc).append(endereco)
        except Exception:
            pass
        try:
            anexos = item.Attachments.Count > 0
        except Exception:
            anexos = False
        threads[conv] = {
            "assunto": item.Subject or "",
            "para": para,
            "cc": cc,
            "enviado_em": dt.isoformat(),
            "tem_anexo": anexos,
            "corpo": (item.Body or "")[:500],
            "entry_id": item.EntryID,
            "_dt": dt,
        }
    return threads


def smtp_do_recipient(r):
    try:
        return r.AddressEntry.GetExchangeUser().PrimarySmtpAddress or r.Address
    except Exception:
        try:
            return r.Address or ""
        except Exception:
            return ""


def descartar(envio):
    """Motivo para nao cobrar esta thread, ou None."""
    destinos = " ".join(envio["para"] + envio["cc"]).lower()
    if not destinos.strip():
        return "sem_destinatario"
    if RE_NOREPLY.search(destinos):
        return "destino_noreply"
    if RE_AUTO.search(envio["assunto"]):
        return "resposta_automatica"
    return None


def paradas_da_conta(store, conta_smtp, cutoff, min_dias, agora):
    recebidos = mapa_recebidos(store, cutoff)
    enviados = enviados_por_thread(store, cutoff)

    paradas, descartados = [], {}
    for conv, envio in enviados.items():
        resposta = recebidos.get(conv)
        if resposta and resposta > envio["_dt"]:
            continue  # alguem respondeu depois do meu envio
        dias = (agora - envio["_dt"]).days
        if dias < min_dias:
            continue  # ainda e cedo para cobrar
        motivo = descartar(envio)
        if motivo:
            descartados[motivo] = descartados.get(motivo, 0) + 1
            continue
        del envio["_dt"]
        envio["dias_parado"] = dias
        envio["conta"] = conta_smtp
        paradas.append(envio)

    paradas.sort(key=lambda e: e["dias_parado"], reverse=True)
    return paradas, descartados


def main():
    dias = int(sys.argv[1]) if len(sys.argv) > 1 else DIAS_PADRAO
    min_dias = int(sys.argv[2]) if len(sys.argv) > 2 else MIN_DIAS_PADRAO
    agora = datetime.datetime.now()
    cutoff = agora - datetime.timedelta(days=dias)

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    contas = {}
    for acc in ns.Accounts:
        try:
            contas[acc.DeliveryStore.StoreID] = acc.SmtpAddress
        except Exception:
            pass

    resultado = {
        "gerado_em": agora.isoformat(),
        "janela_dias": dias,
        "min_dias_parado": min_dias,
        "contas": [],
    }

    for store in ns.Stores:
        smtp = contas.get(store.StoreID)
        if not smtp:
            continue  # arquivos de dados locais, caixas compartilhadas etc.
        print(f"Lendo enviados de {smtp} ({store.DisplayName})...", file=sys.stderr)
        paradas, descartados = paradas_da_conta(store, smtp, cutoff, min_dias, agora)
        resultado["contas"].append({
            "conta": smtp,
            "paradas": paradas,
            "descartados": descartados,
        })
        print(f"  paradas: {len(paradas)}, descartados: {descartados}", file=sys.stderr)

    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
    print(f"OK: {SAIDA}", file=sys.stderr)


if __name__ == "__main__":
    main()
