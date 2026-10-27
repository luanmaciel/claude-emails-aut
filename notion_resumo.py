"""Publica triagem/RESUMO-DIA.md como subpagina de "Resumos do Dia" no Notion.

Token: %USERPROFILE%\\.notion-token (fora do repo, que e publico).
Uso: python notion_resumo.py
"""
import json
import os
import re
import sys
import urllib.request
from datetime import date
from pathlib import Path

PAI = "<id-da-pagina-resumo>"  # pagina "📅 Resumos do Dia"
RESUMO = Path(__file__).parent / "triagem" / "RESUMO-DIA.md"
TOKEN_FILE = Path(os.environ["USERPROFILE"]) / ".notion-token"


def api(metodo, caminho, corpo=None):
    req = urllib.request.Request(
        "https://api.notion.com/v1" + caminho,
        data=json.dumps(corpo).encode() if corpo else None,
        method=metodo,
        headers={
            "Authorization": "Bearer " + TOKEN_FILE.read_text().strip(),
            "Notion-Version": "2022-06-28",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def rich(texto):
    """Converte **negrito**, *italico* e `codigo` inline em rich_text do Notion."""
    partes = []
    for p in re.split(r"(\*\*.+?\*\*|`.+?`|\*[^*]+?\*)", texto):
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            conteudo, anot = p[2:-2], {"bold": True}
        elif p.startswith("`") and p.endswith("`"):
            conteudo, anot = p[1:-1], {"code": True}
        elif p.startswith("*") and p.endswith("*"):
            conteudo, anot = p[1:-1], {"italic": True}
        else:
            conteudo, anot = p, {}
        partes.append({"type": "text", "text": {"content": conteudo[:2000]},
                       "annotations": anot})
    return partes


def blocos(markdown):
    """Markdown do resumo -> blocos do Notion. Ignora o H1 (vira titulo da pagina)."""
    saida = []
    for linha in markdown.splitlines():
        linha = linha.rstrip()
        if not linha or linha.startswith("# ") or set(linha) <= {"-"} and len(linha) > 2:
            continue
        if linha.startswith("## "):
            tipo, texto = "heading_2", linha[3:]
        elif linha.startswith("- "):
            tipo, texto = "bulleted_list_item", linha[2:]
        else:
            tipo, texto = "paragraph", linha
        saida.append({"object": "block", "type": tipo, tipo: {"rich_text": rich(texto)}})
    return saida[:100]  # ponytail: limite da API por request; o resumo tem <=40 linhas


def titulo_do_dia(markdown):
    """Data do H1 (DD/MM/AAAA) -> 'AAAA-MM-DD — Resumo do Dia'. Fallback: hoje."""
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", markdown.split("\n", 1)[0])
    d = f"{m[3]}-{m[2]}-{m[1]}" if m else date.today().isoformat()
    return f"{d} — Resumo do Dia"


def ja_existe(titulo):
    filhos = api("GET", f"/blocks/{PAI}/children?page_size=100")["results"]
    return any(f["type"] == "child_page" and f["child_page"]["title"] == titulo
               for f in filhos)


def main():
    if not TOKEN_FILE.exists():
        sys.exit(f"sem token: crie {TOKEN_FILE} com o secret da integracao Notion")
    md = RESUMO.read_text(encoding="utf-8")
    titulo = titulo_do_dia(md)
    if ja_existe(titulo):
        print(f"ja publicado: {titulo}")
        return
    r = api("POST", "/pages", {
        "parent": {"page_id": PAI},
        "icon": {"type": "emoji", "emoji": "☀️"},
        "properties": {"title": [{"text": {"content": titulo}}]},
        "children": blocos(md),
    })
    print("publicado:", r["url"])


def demo():
    md = "# Resumo do Dia — 05/08/2026\n\n## ALFA\n- **Nemotec** faltam `25 STLs`.\n\n---\n*nota final*\n"
    assert titulo_do_dia(md) == "2026-08-05 — Resumo do Dia"
    b = blocos(md)
    assert [x["type"] for x in b] == ["heading_2", "bulleted_list_item", "paragraph"]
    item = b[1]["bulleted_list_item"]["rich_text"]
    assert item[0] == {"type": "text", "text": {"content": "Nemotec"},
                       "annotations": {"bold": True}}
    assert item[2]["annotations"] == {"code": True}
    assert b[2]["paragraph"]["rich_text"][0]["annotations"] == {"italic": True}
    assert titulo_do_dia("# sem data") == f"{date.today().isoformat()} — Resumo do Dia"
    print("demo ok")


if __name__ == "__main__":
    demo() if "--demo" in sys.argv else main()
