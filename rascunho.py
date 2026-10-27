# -*- coding: utf-8 -*-
"""Cria um rascunho de resposta no Outlook a partir de um arquivo de texto.

Uso: python rascunho.py <EntryID> <arquivo.txt> [--all] [--anexo <caminho>]...
  --all = ReplyAll (default: Reply so ao remetente)
  --anexo = anexa o arquivo ao rascunho (repetivel)
O arquivo e texto puro; linhas em branco separam paragrafos. Nunca envia: so Save().
Assinatura: ALFA (conta padrao) via GetInspector; Beta montada a partir do .htm
(GetInspector inseriria a assinatura da ALFA, incidente 2026-08-21).
"""
import html
import os
import sys

import win32com.client

ASSINATURA_BIO = os.path.join(os.environ["APPDATA"], "Microsoft", "Signatures",
                              "Beta (eu@beta.example.com).htm")


def para_html(texto):
    pars = [p.strip() for p in texto.replace("\r\n", "\n").split("\n\n") if p.strip()]
    return "".join("<p>" + html.escape(p).replace("\n", "<br>") + "</p>" for p in pars)


def main():
    if len(sys.argv) < 3:
        print("Uso: python rascunho.py <EntryID> <arquivo.txt> [--all]")
        return 1
    entry_id, arquivo = sys.argv[1], sys.argv[2]
    todos = "--all" in sys.argv
    anexos = [sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == "--anexo"]
    with open(arquivo, encoding="utf-8") as f:
        corpo = para_html(f.read())

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    original = ns.GetItemFromID(entry_id)
    resp = original.ReplyAll() if todos else original.Reply()
    conta = original.Parent.Store.DisplayName.lower()

    if "beta" in conta:
        # ponytail: assinatura .htm sem embutir imagens (podem sair como link); se incomodar, anexar com cid
        try:
            with open(ASSINATURA_BIO, encoding="utf-8", errors="ignore") as f:
                assinatura = f.read()
        except OSError:
            assinatura = ""
        resp.HTMLBody = corpo + assinatura + resp.HTMLBody
    else:
        resp.GetInspector  # insere a assinatura da conta padrao (ALFA) no HTMLBody
        resp.HTMLBody = corpo + resp.HTMLBody
    for a in anexos:
        resp.Attachments.Add(os.path.abspath(a))
    resp.Save()

    dest = "; ".join(f"{r.Address}({'Cc' if r.Type == 2 else 'Para'})" for r in resp.Recipients)
    print(f"RASCUNHO SALVO [{original.Parent.Store.DisplayName}] '{resp.Subject}' -> {dest}"
          + (f" | anexos: {[os.path.basename(a) for a in anexos]}" if anexos else ""))
    print('Nao enviado. Para enviar: python enviar.py "<termo do assunto>"')
    return 0


if __name__ == "__main__":
    sys.exit(main())
