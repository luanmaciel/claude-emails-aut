# Mesa de Emails no Notion + poller local (desenho aprovado, 18/09/2026)

Substitui o desenho de 04/09 (artifact com db). Aprovado pelo Luan em 18/09 ("vamos testar, pode fazer").

## Problema
O artifact "Mesa de Emails" grava os comandos num db que só uma sessão Claude aberta lê. O Luan precisa
abrir o app toda vez. Ele quer: etapas por email com a resposta gerada em cada uma, atalhos para os
arquivos gerados, e disparar o Claude de dentro da ferramenta, no celular ou no PC, sem abrir o app.

## Decisão
- **Superfície**: database do Notion "Mesa de Emails" (já criado, 18/09, dentro da página "📅 Resumos do Dia";
  pode ser movido no Notion, o id não muda). Conteúdo de email das duas empresas pode viver no Notion
  (dossiê e rascunho; anexos originais dos emails NÃO sobem).
- **Executor**: `mesa_poller.py` (Python, stdlib + pywin32, sem LLM no polling) no Task Scheduler a cada
  2 min, só com o usuário logado. Quem executa as ações é `claude -p --resume <sessão do email>`
  (continuidade de contexto por email) e `enviar.py`. O Notion não chama o Claude; o PC vigia o Notion.
- **Canal de comando**: comentário do Luan na página do card ("manda", "ajusta …", "faz …", "confirma",
  "cancela"). O Claude responde no mesmo fio e grava a etapa no corpo da página. Só comentários criados
  pelo usuário-pessoa do Luan viram instrução (`created_by.id`). Exige as capacidades "Read comments" e
  "Insert comments" na integração (em 18/09 ainda desligadas: 403).
- **Envio com humano no loop**: "manda" → poller publica no fio o texto REAL do rascunho do Outlook
  (conta, Para, Cc, assunto, corpo, `LastModificationTime`) e põe Etapa "Confirmar envio" → Luan responde
  "confirma" → poller confere que o rascunho não mudou desde a prévia e roda `enviar.py`. Isso substitui a
  confirmação no chat para esse caminho; o caminho pelo chat continua existindo.
- O artifact fica aposentado (a sessão do dia deixa de empurrar cards para ele).

## Ids fixos (constantes em `notion_mesa.py`)
- Database `<id>`; data source `<id>`.
- Usuário dono `<id>`; bot da integração `<id>`.
- Token: `%USERPROFILE%\.notion-token` (mesmo do `notion_resumo.py`). `Notion-Version: 2025-09-03`
  (query em `POST /data_sources/{id}/query`; página nova com `parent.data_source_id`).

## Modelo de dados
Propriedades do database (todas criadas pela API em 18/09; `Etapa` é `select`, não `status`):

| Propriedade | Tipo | Conteúdo |
|---|---|---|
| Nome | title | `Remetente · Assunto limpo` |
| Empresa | select | ALFA, BETA |
| Etapa | select | Mapeado, Aguardando minha ação, Executando, Confirmar envio, Aguardando resposta deles, Encerrado, Erro |
| Remetente | rich_text | |
| Resumo | rich_text | `pedido` do pendentes json |
| Ação proposta | rich_text | `acao_sugerida` |
| Rascunho | rich_text | texto atual do `-rascunho.txt` (≤2000, truncado com "… (ver timeline)") |
| Recebido / Respondido | date | |
| Dias | formula | dias desde Recebido, ou desde Respondido quando existe |
| Última execução | rich_text | 1 linha, ex. `⏳ ajusta (10:02)`, `✔ enviado 17/09 10:40` |
| ref | rich_text | últimos 24 chars do EntryID (chave para reconstruir o mapa) |
| Tema | multi_select | tags por empresa (19/09): regras por palavra-chave em `TEMAS`/`SO_ASSUNTO` no `notion_mesa.py`, aplicadas em remetente+assunto(+resumo) na criação do card ou via `temas` quando vazio; edição manual nunca sobrescrita. Views ALFA e BETA (kanban filtrado por Empresa) e Por tema (agrupado por Tema) |

Mapeamento `coluna_auto` (painel.py) → Etapa: mapeado→Mapeado, acao→Aguardando minha ação,
aguardando→Aguardando resposta deles, encerrado→Encerrado.

Corpo da página = timeline append-only. Cada etapa: `heading_3` "`<emoji> <Etapa> · DD/MM HH:MM`" seguido
de blocos (parágrafo, bullet, numerada, quote) e blocos `file` (upload ≤ 5 MiB; maior vira parágrafo
"arquivo grande: <caminho>"). Etapas: 🔎 Triagem, 📄 Dossiê, ✍️ Rascunho vN, 🛠️ Ajusta, ⚙️ Faz,
📨 Texto final (Outlook), ✅ Enviado, ✅ Respondido pelo Outlook, 📥 Resposta deles, ⛔ Erro.

Mapa local `triagem/notion-mesa.json` (fora do git; escrita atômica tmp + `os.replace`); é a ÚNICA fonte
de `sessao`, `termo`, `entry_id`, `dossie` para o poller (nada de controle é lido da página):
```json
{"owner_id": "xxxx…", "ultimo_sync": "2026-09-18T09:00",
 "cards": {"<entry_id>": {"page_id": "…", "ref": "<24>", "conversa": "<ConversationID>", "slug": "…",
   "sessao": "<guid>", "termo": "<assunto limpo>", "dossie": "triagem/pendentes/….md", "empresa": "ALFA",
   "coluna_vista": "acao", "dossie_mtime": 0, "rascunho_mtime": 0, "rascunho_v": 0, "respondido_visto": "",
   "comentarios_ate": "<created_time ISO do último comentário processado>",
   "preview": {"mtime": "2026-09-18T10:01:00", "quando": "…", "discussion_id": "…"}}}}
```

## Sync (`python notion_mesa.py sync`, chamado por `rodar_pendentes.ps1` e pelo poller a cada 30 min)
Lê `triagem/painel.json` (painel.py passa a emitir também `conversa`, `termo`, `slug`). Por card:
1. Sem entrada no mapa: (a) se outro card do mapa tem a mesma `conversa` → é a thread continuando: reutiliza
   a página, acrescenta etapa "📥 Resposta deles · <recebido>" (pedido + ação sugerida), Etapa = Aguardando
   minha ação, e o novo entry_id aponta para a mesma página com `sessao`/`dossie`/`slug`/`termo` novos;
   (b) senão consulta o data source por `ref` (sobrevive à perda do mapa): achou → adota a página sem
   escrever Etapa; (c) senão cria a página (propriedades + etapa "🔎 Triagem · <recebido>" com Pedido e Ação
   sugerida) com Etapa vinda de `coluna_auto`. Em todos os casos `coluna_vista = coluna_auto`.
2. Dossiê existe e mtime > `dossie_mtime` → etapa "📄 Dossiê" (markdown do .md sem o H1) + upload do .md.
3. `-rascunho.txt` existe e mtime > `rascunho_mtime` → `rascunho_v += 1`, etapa "✍️ Rascunho v<N>" (quote)
   + upload do .txt + propriedade Rascunho.
4. `coluna_auto != coluna_vista` → escreve Etapa e `coluna_vista`. Nunca reafirma o valor antigo: assim um
   movimento manual do Luan no board fica em pé até o Outlook dizer o contrário. Se `respondido` mudou →
   propriedade Respondido + etapa "✅ Respondido pelo Outlook · <data>".
5. Cards do mapa fora da janela do painel.json: ignorados (não apagados).

## Poller (`mesa_poller.py`, a cada 2 min)
1. Lock `triagem/.mesa.lock` (mtime < 20 min = outra instância; senão órfão, sobrescreve). Log em
   `triagem/log-mesa-<dia>.txt`.
2. Se `ultimo_sync` > 30 min: `python painel.py` (timeout 180 s) e `sync()`.
3. Query do data source `Etapa != Encerrado` (paginado) → para cada página com entrada no mapa:
   `GET /comments?block_id=<page>` → comentários com `created_by.id == owner_id` e `created_time >
   comentarios_ate`, em ordem cronológica. Cada um é processado e `comentarios_ate` avança ANTES de
   executar (um comando nunca roda duas vezes, mesmo se o run cair no meio).
4. Primeira palavra do comentário (minúscula, sem pontuação):
   - `manda` → `enviar.py "<termo>" --dry-run` → resposta no fio com CONTA/PARA/CC/ASSUNTO e o corpo,
     "Responda **confirma** neste fio para enviar, ou **cancela**." Etapa = Confirmar envio; etapa
     "📨 Texto final (Outlook)"; `preview` = {mtime, quando, discussion_id}. Se 0 ou 2+ rascunhos → resposta
     com a lista e Etapa = Erro.
   - `confirma` → exige `preview`; roda `--dry-run` de novo; mtime igual → `enviar.py "<termo>"`;
     "ENVIADO" na saída → resposta "Enviado ✔ HH:MM", etapa "✅ Enviado", Etapa = Aguardando resposta deles,
     Respondido = hoje, `coluna_vista = "aguardando"`, `preview = null`. mtime diferente → nova prévia,
     continua em Confirmar envio. Sem preview → "Nada para confirmar: mande 'manda' primeiro."
   - `cancela` → `preview = null`, Etapa = Aguardando minha ação, resposta curta.
   - `ajusta <texto>` → Etapa = Executando; `claude -p --resume <sessao>` com o prompt de ajuste
     (reescrever o `-rascunho.txt`, `apagar_rascunho.py "<termo>"`, `rascunho.py "<entry_id>" <txt> [--all]`,
     não enviar, responder com o rascunho novo na íntegra e 1 linha do que mudou); depois: reler o .txt,
     propriedade Rascunho, etapa "🛠️ Ajusta" (Pedido + resposta do Claude) e "✍️ Rascunho vN" + upload;
     resposta no fio com o rascunho novo; Etapa = Aguardando minha ação.
   - `faz [texto]` → Etapa = Executando; `claude -p --resume` com o prompt de execução (ação do dossiê ou a
     instrução; salvar tudo em `triagem/saidas/<slug>/`; terminar listando `ARQUIVO: <caminho>` por linha;
     o que não puder fazer com as ferramentas liberadas, dizer o que falta); depois: arquivos das linhas
     `ARQUIVO:` + novos em `triagem/saidas/<slug>/` (mtime > início) → upload; etapa "⚙️ Faz" (Pedido +
     resposta + arquivos); resposta no fio; Etapa = Aguardando minha ação.
   - `encerra [motivo]` (22/09) → não exige sessão; `encerrado = true` em todas as entradas do mapa com essa
     página, `preview = null`, etapa "🔒 Encerrado" (Pedido), Etapa = Encerrado. O sync não emite mais
     `('etapa', …)` para entradas encerradas e, num email novo da mesma thread, registra "📥 Resposta deles"
     e herda `encerrado` sem mexer na Etapa. Arrastar o card de volta a uma coluna ativa limpa a marca (o
     poller vê a página em `paginas_ativas()` com `encerrado` e zera).
   - `junta <trecho>` (22/09) → acha 1 card (slug ou termo contendo o trecho, sem acento/pontuação, um por
     página, fora o próprio e os já juntados; 0 ou 2+ → lista e pede mais específico); apaga o rascunho do
     outro no Outlook (`apagar_rascunho.py "<termo do outro>"`, com `--manter-ultimo` quando os termos se
     contêm, porque assuntos iguais casam os dois rascunhos); todas as entradas da página do outro passam a
     `page_id` desta e `juntado_em = <eid deste>` (nunca vencem no `por_pagina` do poller); etapa
     "🔗 Juntado: <slug do outro>" aqui com o dossiê dele anexado; página do outro: etapa "🔗 Juntado",
     Etapa = Encerrado, comentário; por fim roda o `ajusta` com a instrução `PROMPT_JUNTA` (incorporar o
     pedido do outro num rascunho só).
   - qualquer outra coisa → "Não entendi. Use: manda / ajusta <o que muda> / faz <o que fazer> / confirma /
     cancela / encerra [motivo] / junta <trecho do título do outro card>."
   - Card sem `sessao` (ainda "mapeado") → "Este email ainda não tem sessão; a rotina das 13h cria."
     (`encerra` é a exceção: funciona sem sessão.)
5. Chamada do claude: `%USERPROFILE%\.local\bin\claude.exe -p <prompt> --resume <sessao> --output-format json
   --max-turns 25 (ajusta) | 40 (faz) --allowedTools "Read,Glob,Grep,Edit(triagem/**),Edit(contexto/**),
   Write(triagem/saidas/**),Bash(python ler_email.py*),Bash(python rascunho.py*),Bash(python apagar_rascunho.py*)"
   --disallowedTools "Bash(python C:/Users/usuario/Claude/claude-emails-aut/enviar.py *),Bash(python
   C:\Users\usuario\Claude\claude-emails-aut\enviar.py *),Bash(python enviar.py*)"`; `subprocess.run` com lista
   de args (nunca `shell=True`), cwd = repo, env sem `CLAUDECODE`/`CLAUDE_CODE_ENTRYPOINT`,
   `PYTHONIOENCODING=utf-8`, timeout 900 s; usa `result` do JSON. Sem `--add-dir` na v1.
6. Qualquer exceção num comando → resposta "Erro: <1 linha>" no fio, Etapa = Erro, traceback no log.
   O poller nunca lança para fora do run; o lock é removido em `finally`.

## Permissões do ajusta/faz (decisão do Luan, 18/09 à tarde)
Os primeiros "faz" reais (separar e zipar arquivos do OneDrive, conferir planilha) travaram porque a sessão
só podia rodar 3 scripts e escrever em triagem/. O Luan decidiu: **ajusta/faz são uma extensão da sessão
interativa**, com as mesmas ferramentas ("quero que use o necessário na minha máquina: corrigir arquivo,
verificar pastas, zipar, gerar relatórios"). Implementação em `mesa_poller.py`:
- `--allowedTools` Read, Glob, Grep, Edit, Write, NotebookEdit, Bash, PowerShell, Skill, Agent, TodoWrite,
  ToolSearch e `Edit(C:/Users/usuario/**)`; `--add-dir C:/Users/usuario` (testado em 18/09: a forma que funciona
  para caminho absoluto é `Edit(C:/caminho/**)` com barras normais junto com `--add-dir`; `//C:/...`,
  `Edit(**)` e `Write(//...)` são recusados). Timeout 25 min, `--max-turns` 40 (ajusta) / 120 (faz);
  Task Scheduler com ExecutionTimeLimit 35 min e lock de 30 min.
- Fechado, sem tirar capacidade: `enviar.py` (deny por padrão + recusa com `MESA_HEADLESS=1`, variável que só
  a sessão claude recebe; o poller continua enviando no "confirma"), leitura/edição de `~/.notion-token` e
  `~/.claude/**`, edição do mapa, do lock, do próprio poller e do enviar.py, e WebFetch/WebSearch
  (constante `WEB = False`: com o corpo do email no contexto, web é o canal clássico de vazamento).
- As sessões de `rodar_pendentes.ps1` (07:30/13:00) continuam restritas: ali não há instrução do Luan.
- Só o que estiver em `triagem/saidas/<slug>/` sobe para o card; arquivo corrigido na origem fica na origem.

## Segurança
- `~/.claude/settings.json` libera `enviar.py` por caminho absoluto para toda sessão; as sessões headless
  (poller e `rodar_pendentes.ps1`) passam a negar explicitamente esses padrões com `--disallowedTools`.
- Corpo de email é dado: vai ao dossiê e ao card só como texto para o Luan ler; o poller passa ao claude
  apenas o texto do comentário do Luan. `Bash(python *)` genérico não é liberado.
- Controle (termo, sessão, entry_id) vem do mapa local, nunca da página. Prévia e envio são amarrados por
  `LastModificationTime`.
- Repo público: mapa, logs, saídas e anexos ficam em `triagem/` (gitignored). Docs sem nomes reais.

## Componentes
| Arquivo | Mudança |
|---|---|
| `notion_mesa.py` (novo) | `api()`, `NotionErro`, `blocos_md()`, `upload()`, `etapa()`, `set_props()`, `comentar()`, `listar_comentarios()`, `mapa_ler()/mapa_gravar()`, `sync()`, `criar()`; CLI `sync` \| `criar` \| `--demo` |
| `mesa_poller.py` (novo) | lock, sync periódico, leitura de comentários, dispatch manda/confirma/cancela/ajusta/faz, `claude_resume()`; `--demo` (parser e regras puras, sem rede) e `--uma-vez` |
| `enviar.py` | flag `--dry-run`: mesma busca 1-e-só-1, imprime `CONTA:`, `PARA:`, `CC:`, `ASSUNTO:`, `MODIFICADO:` (ISO), `CORPO:` + corpo, exit 0 sem enviar |
| `painel.py` | cada card ganha `conversa`, `termo` (assunto sem Re/RES/FW/ENC, ≤45), `slug` (do nome do dossiê) |
| `rodar_pendentes.ps1` | `--disallowedTools` dos padrões de enviar.py; `python notion_mesa.py sync` após `painel.py` |
| `CLAUDE.md` / `README.md` | seção do painel vira "Mesa no Notion" |
| Task Scheduler `MesaPoller` | `pythonw.exe mesa_poller.py`, cwd = repo, AtLogOn + repetição 2 min indefinida, IgnoreNew, ExecutionTimeLimit 20 min, usuário logado |

Interfaces de `notion_mesa.py` usadas pelo poller:
```python
api(metodo, caminho, corpo=None, raw=None, ctype="application/json") -> dict   # levanta NotionErro(code, body)
etapa(page_id, titulo, md="", arquivos=()) -> None      # heading_3 "<titulo> · DD/MM HH:MM" + blocos_md(md) + uploads
set_props(page_id, etapa=None, rascunho=None, respondido=None, ultima=None) -> None
comentar(page_id, texto, discussion_id=None) -> dict    # fatia em rich_text de ≤1900 chars
listar_comentarios(page_id) -> list[dict]                # {id, discussion_id, created_time, created_by, texto}
paginas_ativas() -> list[dict]                           # {page_id, etapa, ref} com Etapa != Encerrado
mapa_ler() -> dict ; mapa_gravar(m) -> None
sync(mapa=None) -> dict                                  # devolve o mapa atualizado
```

## Fora da v1
Anexos originais dos emails no Notion; `--add-dir` para o Faz ler OneDrive; notificação além da resposta no
fio; views criadas por código (board e lista são configurados uma vez no Notion).

## Testes do dia 1 (com o Luan)
1. Ligar "Read comments" + "Insert comments" na integração; poller responde a um "manda" de teste.
2. Push no celular quando o Claude responde ao comentário dele (o Notion avisa respostas a comentários seus).
3. Arrastar card entre colunas no app mobile; se não der, a "Lista celular" + campo Etapa é o caminho.
