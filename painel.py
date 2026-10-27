# -*- coding: utf-8 -*-
"""Monta triagem/painel.json (cards do painel kanban) a partir dos pendentes + Outlook.

Uso: python painel.py [dias]   (default: 14 dias de pendentes-*.json)
Coluna automatica por card:
  mapeado    - esta em pendentes-*.json mas ainda sem sessao/dossie
  acao       - dossie + rascunho prontos, sem resposta minha depois do email
  aguardando - respondi depois do email (Itens Enviados na mesma ConversationID)
  encerrado  - respondi ha mais de DIAS_ENCERRA dias e nada novo chegou na thread
A sessao do dia empurra o JSON para o db do artifact (write_db batch, op "set" em cards/<id> e meta/painel) e le comandos/. Coluna movida a mao fica em overrides/<id> e nao e tocada aqui.
"""
import glob
import json
import os
import re
import sys
from datetime import datetime, timedelta

import win32com.client

DIAS_ENCERRA = 7  # ponytail: fixo; vira config se o Luan pedir
TRIAGEM = os.path.join(os.path.dirname(os.path.abspath(__file__)), "triagem")


def carregar_json(caminho, padrao):
    try:
        with open(caminho, encoding="utf-8-sig") as f:
            return json.load(f)
    except (OSError, ValueError):
        return padrao


def ler(caminho):
    try:
        with open(caminho, encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


def tokens(s):
    return set(re.findall(r"[a-z0-9]{3,}", s.lower()))


def limpa(assunto):
    """Assunto sem prefixos Re:/RES:/FW:/FWD:/ENC: (repetidos, qualquer caixa) = termo para enviar.py."""
    return re.sub(r"^(?:(?:re|res|fw|fwd|enc)\s*:\s*)+", "", (assunto or "").strip(), flags=re.I).strip()


def slug_do(dossie):
    """triagem/pendentes/AAAA-MM-DD-<slug>.md -> <slug>; '' sem dossie."""
    return re.sub(r"^\d{4}-\d{2}-\d{2}-", "", os.path.basename(dossie))[:-3] if dossie else ""


def achar_dossie(dia, remetente, assunto):
    """Dossie do dia com mais tokens em comum com remetente+assunto (a sessao nem sempre usa o slug exato)."""
    alvo = tokens(remetente + " " + assunto)
    melhor, pontos = "", 0
    for arq in glob.glob(os.path.join(TRIAGEM, "pendentes", f"{dia}-*.md")):
        n = len(alvo & tokens(os.path.basename(arq)))
        if n > pontos:
            melhor, pontos = arq, n
    return os.path.relpath(melhor) if melhor else ""


def main():
    dias = int(sys.argv[1]) if len(sys.argv) > 1 else 14
    cutoff = datetime.now() - timedelta(days=dias)

    pendentes, sessoes = {}, {}
    for arq in sorted(glob.glob(os.path.join(TRIAGEM, "pendentes-*.json"))):
        dia = re.search(r"pendentes-(\d{4}-\d{2}-\d{2})", arq).group(1)
        if datetime.fromisoformat(dia) < cutoff:
            continue
        for it in carregar_json(arq, []):
            it["dia"] = dia
            pendentes[it["entry_id"]] = it  # o mais recente vence
        for eid, s in carregar_json(arq.replace("pendentes-", "sessoes-"), {}).items():
            s["dia"] = dia
            sessoes[eid] = s

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # ConversationID -> (ultimo envio meu, ultimo recebido) nos ultimos `dias`+30.
    # ass_env: assunto limpo -> ultimo envio meu; fallback porque o enviar.py, quando reenvia como item novo,
    # gera ConversationID diferente do email recebido (7 de 9 replies dos ultimos 30 dias, conferido em 18/09).
    conv_env, conv_rec, ass_env = {}, {}, {}
    for store in ns.Stores:
        for pasta_id, mapa in ((5, conv_env), (6, conv_rec)):
            try:
                items = store.GetDefaultFolder(pasta_id).Items
            except Exception:
                continue
            items.Sort("[ReceivedTime]", True)
            for item in items:
                try:
                    dt = item.ReceivedTime.replace(tzinfo=None)
                except Exception:
                    continue
                if dt < cutoff - timedelta(days=30):
                    break
                try:
                    conv = item.ConversationID
                except Exception:
                    continue
                if conv and (conv not in mapa or dt > mapa[conv]):
                    mapa[conv] = dt
                if pasta_id == 5:
                    try:
                        chave = limpa(item.Subject).lower()
                    except Exception:
                        continue
                    if chave and (chave not in ass_env or dt > ass_env[chave]):
                        ass_env[chave] = dt

    cards = []
    for eid, it in pendentes.items():
        try:
            item = ns.GetItemFromID(eid)
            conv = item.ConversationID
            recebido = item.ReceivedTime.replace(tzinfo=None)
        except Exception:
            conv, recebido = "", datetime.fromisoformat(it.get("recebido", it["dia"])[:19])
        env = conv_env.get(conv) or ass_env.get(limpa(it.get("assunto", "")).lower())
        rec = conv_rec.get(conv)
        s = sessoes.get(eid)
        if env and env > recebido:
            if datetime.now() - env > timedelta(days=DIAS_ENCERRA) and not (rec and rec > env):
                coluna = "encerrado"
            else:
                coluna = "aguardando"
        elif s:
            coluna = "acao"
        else:
            coluna = "mapeado"
        # dossie existe independente de sessao (cards montados a mao tambem tem)
        dossie = achar_dossie(it["dia"], it.get("remetente", ""), it.get("assunto", ""))
        if dossie and coluna == "mapeado":
            coluna = "acao"
        cards.append({
            "id": eid[-24:],
            "entry_id": eid,
            "empresa": "BETA" if "beta" in it.get("conta", "") else "ALFA",
            "remetente": it.get("remetente", ""),
            "assunto": it.get("assunto", ""),
            "conversa": conv,
            "termo": limpa(it.get("assunto", ""))[:45].strip(),
            "slug": slug_do(dossie),
            "pedido": it.get("pedido", ""),
            "acao": it.get("acao_sugerida", ""),
            "rascunho": ler(dossie.replace(".md", "-rascunho.txt")) if dossie else "",
            "dossie": dossie if dossie and os.path.exists(dossie) else "",
            "sessao": s["sessao"] if s else "",
            "recebido": recebido.isoformat(timespec="minutes"),
            "respondido": env.isoformat(timespec="minutes") if env and env > recebido else "",
            "coluna_auto": coluna,
            "atualizado": datetime.now().isoformat(timespec="minutes"),
        })

    saida = {"gerado_em": datetime.now().isoformat(timespec="minutes"),
             "resumo": ler(os.path.join(TRIAGEM, "RESUMO-DIA.md")), "cards": cards}
    with open(os.path.join(TRIAGEM, "painel.json"), "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=1)
    from collections import Counter
    print(f"{len(cards)} cards:", dict(Counter(c["coluna_auto"] for c in cards)))
    return 0


def demo():
    assert limpa("RE: Res: FW: Assunto: x") == "Assunto: x"
    assert limpa("  enc:  Proposta ") == "Proposta"
    assert limpa("Resumo sem prefixo") == "Resumo sem prefixo"
    assert slug_do("triagem\\pendentes\\2026-09-17-leonardo-miranda-inserto-3pkt.md") == "leonardo-miranda-inserto-3pkt"
    assert slug_do("") == ""
    print("demo ok")


if __name__ == "__main__":
    sys.exit(demo() if "--demo" in sys.argv else main())
