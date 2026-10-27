# -*- coding: utf-8 -*-
"""Envia UM rascunho existente do Outlook, identificado por termo do assunto.

Uso: python enviar.py "termo do assunto" [--dry-run] [--modificado <iso>]
Guardrails: só envia itens já salvos em Rascunhos; recusa se 0 ou 2+ correspondem, ou termo vazio.
--dry-run: acha o rascunho e imprime CONTA:/PARA:/CC:/ANEXOS:/ASSUNTO:/MODIFICADO:/CORPO: + corpo, sem enviar.
--modificado <iso>: recusa se o LastModificationTime do rascunho não for esse (a Mesa no Notion amarra
o envio à prévia que o Luan leu).
"""
import os
import sys
import tempfile

import win32com.client


def smtp(r):
    """Endereço SMTP de um Recipient (Exchange vem como /O=EXCHANGELABS/...)."""
    try:
        if r.AddressEntry.Type == "EX":
            return r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    except Exception:
        pass
    return r.Address


def conta_do_rascunho(app, it):
    """Account cuja store contém o rascunho (para não enviar pela conta padrão)."""
    try:
        store_id = it.Parent.Store.StoreID
        for acct in app.Session.Accounts:
            if acct.DeliveryStore.StoreID == store_id:
                return acct
    except Exception:
        pass
    return None


def reenviar_como_item_novo(it):
    """Rascunho de Reply/ReplyAll recusa Send() com 'Parâmetro incorreto'.

    Copiar assunto/corpo para um item novo com destinatários em SMTP puro envia normal.
    """
    para = [smtp(r) for r in it.Recipients if r.Type != 2]
    cc = [smtp(r) for r in it.Recipients if r.Type == 2]
    app = win32com.client.Dispatch("Outlook.Application")
    novo = app.CreateItem(0)
    acct = conta_do_rascunho(app, it)
    if acct is not None:
        # atribuição direta de SendUsingAccount não pega no win32com; Invoke no dispid resolve
        novo._oleobj_.Invoke(64209, 0, 8, 0, acct)
        print(f"Enviando pela conta: {acct.SmtpAddress}")
    else:
        print("AVISO: conta do rascunho nao identificada; vai sair pela conta padrao.")
    novo.Subject = it.Subject
    novo.HTMLBody = it.HTMLBody
    novo.To = "; ".join(para)
    novo.CC = "; ".join(cc)
    # anexos não passam entre itens: salvar em disco (pasta que se apaga sozinha) e reanexar
    with tempfile.TemporaryDirectory() as tmp:
        for a in it.Attachments:
            caminho = os.path.join(tmp, os.path.basename(a.FileName or "anexo"))
            a.SaveAsFile(caminho)
            novo.Attachments.Add(caminho)
        novo.Recipients.ResolveAll()
        novo.Send()
    it.Delete()


def main():
    if len(sys.argv) < 2:
        print("Uso: python enviar.py \"termo do assunto\"")
        return 1
    if os.environ.get("MESA_HEADLESS") and "--dry-run" not in sys.argv:
        print("ERRO: envio bloqueado dentro de sessão automática da Mesa; o envio só sai pelo card (manda + confirma) ou pelo chat.")
        return 1
    termo = sys.argv[1].strip().lower()
    if not termo:
        print("ERRO: termo vazio; informe um trecho do assunto ('' casaria com todo rascunho).")
        return 1
    # --modificado <iso>: só envia se o rascunho ainda for o da prévia (a Mesa no Notion passa o MODIFICADO do --dry-run)
    esperado = sys.argv[sys.argv.index("--modificado") + 1] if "--modificado" in sys.argv else None

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
        print(f"ERRO: {len(achados)} rascunhos correspondem a '{termo}'; precisa ser exatamente 1.")
        return 1

    conta, it = achados[0]
    modificado = it.LastModificationTime.replace(tzinfo=None).isoformat(timespec="seconds")
    if "--dry-run" in sys.argv:
        print("CONTA:", conta)
        print("PARA:", "; ".join(smtp(r) for r in it.Recipients if r.Type != 2))
        print("CC:", "; ".join(smtp(r) for r in it.Recipients if r.Type == 2))
        print("ANEXOS:", "; ".join(f"{a.FileName} ({a.Size} B)" for a in it.Attachments))
        print("ASSUNTO:", it.Subject or "")
        print("MODIFICADO:", modificado)
        print("CORPO:")
        print(it.Body or "")
        return 0
    if esperado and modificado != esperado:
        print(f"ERRO: rascunho mudou depois da prévia ({modificado} != {esperado}); peça a prévia de novo.")
        return 1
    dest = "; ".join(f"{r.Address}({'Cc' if r.Type == 2 else 'Para'})" for r in it.Recipients)
    print(f"Enviando [{conta}] '{it.Subject}' -> {dest}")
    try:
        it.Send()
    except Exception as ex:
        print(f"Send() direto falhou ({ex}); reenviando como item novo...")
        reenviar_como_item_novo(it)
    print("ENVIADO")
    return 0


if __name__ == "__main__":
    sys.exit(main())
