# -*- coding: utf-8 -*-
"""Apaga rascunhos do Outlook por termo do assunto (passo do "ajusta" e limpeza de duplicados).

Uso: python apagar_rascunho.py "termo do assunto" [--manter-ultimo]
  --manter-ultimo = preserva o rascunho mais recente e apaga os demais (enviar.py recusa com 2+).
Só toca a pasta Rascunhos; nunca envia.
"""
import sys

import win32com.client


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    termo, manter = sys.argv[1].strip().lower(), "--manter-ultimo" in sys.argv
    if not termo:
        print("ERRO: termo vazio; informe um trecho do assunto.")
        return 1
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    achados = []
    for store in ns.Stores:
        try:
            items = store.GetDefaultFolder(16).Items  # olFolderDrafts
        except Exception:
            continue
        for it in items:
            try:
                if termo in (it.Subject or "").lower():
                    achados.append(it)
            except Exception:
                pass
    achados.sort(key=lambda it: it.LastModificationTime)
    if manter and achados:
        achados = achados[:-1]
    elif len(achados) > 1:  # sem --manter-ultimo, apagar em lote so de proposito (termo curto casaria com tudo)
        for it in achados:
            print("-", it.Subject[:70], "|", it.LastModificationTime.strftime("%d/%m %H:%M"))
        print(f"ERRO: {len(achados)} rascunhos correspondem a '{termo}'; use um termo mais especifico ou --manter-ultimo.")
        return 1
    for it in achados:
        print("apagado:", it.Subject[:70], "|", it.LastModificationTime.strftime("%d/%m %H:%M"))
        it.Delete()
    print(f"{len(achados)} rascunho(s) apagado(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
