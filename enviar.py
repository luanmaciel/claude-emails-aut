# -*- coding: utf-8 -*-
"""Envia UM rascunho existente do Outlook, identificado por termo do assunto.

Uso: python enviar.py "termo do assunto"
Guardrails: só envia itens já salvos em Rascunhos; recusa se 0 ou 2+ correspondem.
"""
import sys

import win32com.client


def main():
    if len(sys.argv) < 2:
        print("Uso: python enviar.py \"termo do assunto\"")
        return 1
    termo = sys.argv[1].lower()

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    achados = []
    for store in ns.Stores:
        try:
            rascunhos = store.GetDefaultFolder(16)  # olFolderDrafts
        except Exception:
            continue
        for it in rascunhos.Items:
            try:
                if it.Class == 43 and termo in (it.Subject or "").lower():
                    achados.append((store.DisplayName, it))
            except Exception:
                pass

    if len(achados) != 1:
        for conta, it in achados:
            print(f"- [{conta}] {it.Subject}")
        print(f"ERRO: {len(achados)} rascunhos correspondem a '{termo}' — precisa ser exatamente 1.")
        return 1

    conta, it = achados[0]
    dest = "; ".join(f"{r.Address}({'Cc' if r.Type == 2 else 'Para'})" for r in it.Recipients)
    print(f"Enviando [{conta}] '{it.Subject}' -> {dest}")
    it.Send()
    print("ENVIADO")
    return 0


if __name__ == "__main__":
    sys.exit(main())
