# -*- coding: utf-8 -*-
"""Checagem mínima da regra cc_sem_mencao: python test_filtrar.py"""
from extrair import filtrar

def email(remetente, para, corpo=""):
    return {"remetente_smtp": remetente, "para": para, "cc": [], "assunto": "x",
            "corpo": corpo, "_corpo_completo": corpo}

EU = ["eu@beta.example.com"]
NOME = "Garcilasso, Luan"
# colega interno com eu só em Cc: mantém
assert filtrar(email("colega@beta.example.com", ["contato@fornecedor.example.com"]), EU, NOME, None) is None
# externo com eu só em Cc e sem me citar: descarta
assert filtrar(email("fornecedor@fornecedor.example.com", ["colega@beta.example.com"]), EU, NOME, None) == "cc_sem_mencao"
# externo em Cc que cita só o sobrenome (com a vírgula do Outlook no nome): mantém
assert filtrar(email("fornecedor@fornecedor.example.com", ["x@y.com"], "conforme o Garcilasso pediu"), EU, NOME, None) is None
print("ok")
