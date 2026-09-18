# -*- coding: utf-8 -*-
"""Testes da logica do cobrancas.py com Outlook falso (rodam em qualquer SO).

Uso: python teste_cobrancas.py

Cobre o pareamento enviado/recebido, a varredura de subpastas e os cortes.
A camada COM em si (Dispatch, MAPI) nao e testada aqui: so o miolo.
"""
import datetime
import sys
import types
import unittest

# cobrancas.py importa win32com no topo; num SO sem Outlook, substituimos por um talo.
if "win32com" not in sys.modules:
    win32com = types.ModuleType("win32com")
    win32com.client = types.ModuleType("win32com.client")
    win32com.client.Dispatch = lambda *a, **k: None
    sys.modules["win32com"] = win32com
    sys.modules["win32com.client"] = win32com.client

import cobrancas

AGORA = datetime.datetime(2026, 9, 18, 9, 0, 0)


def dias_atras(n):
    return AGORA - datetime.timedelta(days=n)


class AddressEntryFalso:
    def GetExchangeUser(self):
        raise RuntimeError("sem Exchange: cai no Address")


class DestinatarioFalso:
    def __init__(self, endereco, tipo=1):
        self.Address = endereco
        self.Type = tipo  # 1 = Para, 2 = Cc
        self.AddressEntry = AddressEntryFalso()


class AnexosFalsos:
    Count = 0


class ItemFalso:
    Class = 43  # olMail

    def __init__(self, conv, data, assunto="Assunto", destinos=(), corpo="corpo", classe=43):
        self.ConversationID = conv
        self.ReceivedTime = data
        self.SentOn = data
        self.Subject = assunto
        self.Recipients = [DestinatarioFalso(d) for d in destinos]
        self.Attachments = AnexosFalsos()
        self.Body = corpo
        self.EntryID = f"EID-{conv}-{data:%Y%m%d}"
        self.Class = classe


class ItemsFalsos(list):
    def Sort(self, campo, desc):
        self.sort(key=lambda i: i.ReceivedTime, reverse=desc)


class PastaFalsa:
    def __init__(self, itens=(), subpastas=()):
        self.Items = ItemsFalsos(itens)
        self.Folders = list(subpastas)


class StoreFalso:
    def __init__(self, inbox, enviados):
        self._pastas = {6: inbox, 5: enviados}

    def GetDefaultFolder(self, n):
        return self._pastas[n]


def rodar(inbox, enviados, janela=60, min_dias=4):
    store = StoreFalso(inbox, enviados)
    cutoff = AGORA - datetime.timedelta(days=janela)
    return cobrancas.paradas_da_conta(store, "eu@empresa.com", cutoff, min_dias, AGORA)


class TesteParadas(unittest.TestCase):
    def test_envio_sem_resposta_vira_cobranca(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), "Dongle Exoplan", ["fornecedor@x.com"])])
        paradas, _ = rodar(PastaFalsa(), enviados)
        self.assertEqual(len(paradas), 1)
        self.assertEqual(paradas[0]["dias_parado"], 10)
        self.assertEqual(paradas[0]["assunto"], "Dongle Exoplan")
        self.assertEqual(paradas[0]["para"], ["fornecedor@x.com"])
        self.assertNotIn("_dt", paradas[0])

    def test_resposta_posterior_encerra_a_cobranca(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), destinos=["f@x.com"])])
        inbox = PastaFalsa([ItemFalso("C1", dias_atras(8))])
        paradas, _ = rodar(inbox, enviados)
        self.assertEqual(paradas, [])

    def test_resposta_anterior_ao_envio_nao_conta(self):
        # responderam, eu respondi de volta: a bola voltou para eles
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(6), destinos=["f@x.com"])])
        inbox = PastaFalsa([ItemFalso("C1", dias_atras(9))])
        paradas, _ = rodar(inbox, enviados)
        self.assertEqual(len(paradas), 1)
        self.assertEqual(paradas[0]["dias_parado"], 6)

    def test_resposta_em_subpasta_encerra_a_cobranca(self):
        # a pasta "BioHorizon" fica fora da Inbox raiz
        biohorizon = PastaFalsa([ItemFalso("C1", dias_atras(7))])
        inbox = PastaFalsa([], subpastas=[biohorizon])
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), destinos=["f@x.com"])])
        paradas, _ = rodar(inbox, enviados)
        self.assertEqual(paradas, [])

    def test_envio_recente_ainda_nao_e_cobranca(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(2), destinos=["f@x.com"])])
        paradas, _ = rodar(PastaFalsa(), enviados, min_dias=4)
        self.assertEqual(paradas, [])

    def test_destino_noreply_e_descartado(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), destinos=["no-reply@x.com"])])
        paradas, descartados = rodar(PastaFalsa(), enviados)
        self.assertEqual(paradas, [])
        self.assertEqual(descartados.get("destino_noreply"), 1)

    def test_envio_sem_destinatario_e_descartado(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), destinos=[])])
        paradas, descartados = rodar(PastaFalsa(), enviados)
        self.assertEqual(paradas, [])
        self.assertEqual(descartados.get("sem_destinatario"), 1)

    def test_so_o_envio_mais_recente_da_thread_conta(self):
        enviados = PastaFalsa([
            ItemFalso("C1", dias_atras(30), "Primeira cobranca", ["f@x.com"]),
            ItemFalso("C1", dias_atras(9), "Segunda cobranca", ["f@x.com"]),
        ])
        paradas, _ = rodar(PastaFalsa(), enviados)
        self.assertEqual(len(paradas), 1)
        self.assertEqual(paradas[0]["dias_parado"], 9)
        self.assertEqual(paradas[0]["assunto"], "Segunda cobranca")

    def test_ordena_da_mais_antiga_para_a_mais_nova(self):
        enviados = PastaFalsa([
            ItemFalso("C1", dias_atras(5), "Nova", ["a@x.com"]),
            ItemFalso("C2", dias_atras(20), "Velha", ["b@x.com"]),
            ItemFalso("C3", dias_atras(12), "Media", ["c@x.com"]),
        ])
        paradas, _ = rodar(PastaFalsa(), enviados)
        self.assertEqual([p["assunto"] for p in paradas], ["Velha", "Media", "Nova"])

    def test_fora_da_janela_nao_aparece(self):
        enviados = PastaFalsa([
            ItemFalso("C1", dias_atras(80), "Antiga demais", ["a@x.com"]),
            ItemFalso("C2", dias_atras(10), "Dentro", ["b@x.com"]),
        ])
        paradas, _ = rodar(PastaFalsa(), enviados, janela=60)
        self.assertEqual([p["assunto"] for p in paradas], ["Dentro"])

    def test_item_que_nao_e_email_e_ignorado(self):
        enviados = PastaFalsa([ItemFalso("C1", dias_atras(10), destinos=["f@x.com"], classe=53)])
        paradas, _ = rodar(PastaFalsa(), enviados)
        self.assertEqual(paradas, [])

    def test_corte_usa_o_campo_da_ordenacao(self):
        # SentOn bem mais antigo que ReceivedTime nao pode interromper a varredura
        armadilha = ItemFalso("C9", dias_atras(3), "Armadilha", ["a@x.com"])
        armadilha.SentOn = dias_atras(400)
        alvo = ItemFalso("C1", dias_atras(10), "Alvo", ["b@x.com"])
        enviados = PastaFalsa([armadilha, alvo])
        paradas, _ = rodar(PastaFalsa(), enviados)
        self.assertIn("Alvo", [p["assunto"] for p in paradas])


if __name__ == "__main__":
    unittest.main(verbosity=2)
