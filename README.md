# Triagem diária de emails corporativos com Claude Code (Outlook local, sem OAuth)

Rotina local que transforma duas caixas corporativas do Outlook Desktop em um assistente diário: triagem priorizada, memória de contexto por empresa, briefings por projeto e uma sessão de chat acessível do celular — tudo sem Graph API, sem OAuth e sem add-in aprovado por TI (cenário comum: contas corporativas sem acesso de administrador).

> Versão pública: os nomes das duas empresas, domínios, endereços e IDs foram trocados por Alfa, Beta e placeholders. O código é o mesmo que roda na rotina real.

## Arquitetura

```
07:30  extrair.py (pywin32/COM) ──> emails.json ──> claude -p (triagem-prompt.md)
                                                     ├─> triagem/AAAA-MM-DD.md   (AÇÃO HOJE / SEMANA / MONITORAR)
                                                     └─> contexto/<empresa>/     (vaults Obsidian por empresa)
07:45  claude -p por projeto (BRIEFING-PROMPT.md) ──> BRIEFING-DIA.md + plano atualizado
07:35  rodar_pendentes.ps1: 1 sessão claude -p por email pendente (pendente-prompt.md)
       ├─> triagem/pendentes/AAAA-MM-DD-<slug>.md (dossiê: pedido, ação proposta, rascunho)
       ├─> rascunho de resposta salvo no Outlook (rascunho.py); sessão nomeada fica na aba Code
       └─> painel.py + notion_mesa.py sync ──> card no database "Mesa de Emails" do Notion
08:05  claude -p (resumo-prompt.md) ──> triagem/RESUMO-DIA.md (formato celular) ──> Notion
13:00  rodar_tudo.ps1 -Tarde: re-extrai, triagem incremental, sessões só dos emails novos
2 min  mesa_poller.py: comentário do usuário no card ("manda", "ajusta …", "faz …", "confirma")
       ──> claude -p --resume <sessão do email> | enviar.py ──> resposta no mesmo fio + etapa na timeline
```

No Notion, cada card tem resumo, ação proposta, rascunho e a timeline do que foi feito, com os arquivos
gerados anexados. O usuário comanda por comentário no card, do celular ou do PC, sem abrir o app Claude;
o poller no PC executa (o Outlook só existe lá) e responde no fio. No app, cada sessão
"DD/MM ALFA - Remetente - Assunto" continua abrindo com o mesmo contexto para quem preferir o chat.

Tudo agendado no Task Scheduler do Windows (`schtasks`), rodando com a sessão do usuário logada.

## Componentes

| Arquivo | Papel |
|---|---|
| `extrair.py` | Lê as caixas via COM/MAPI (sem LLM), aplica pré-filtro determinístico, gera JSON |
| `triagem-prompt.md` | Prompt do `claude -p` diário: triagem priorizada + atualização dos vaults de contexto |
| `resumo-prompt.md` | Consolida triagem + briefings num resumo de celular |
| `rodar_tudo.ps1` | Orquestra tudo no logon (incremental) e às 13h com `-Tarde` |
| `rodar_triagem.ps1` / `rodar_briefing.ps1` / `rodar_resumo.ps1` | Runners das etapas |
| `rodar_pendentes.ps1` + `pendente-prompt.md` | Uma sessão `claude -p --session-id --name` por email pendente; idempotente via `triagem/sessoes-<dia>.json` |
| `ler_email.py` | Lê um email por EntryID (corpo completo) e salva anexos em `triagem/anexos/` |
| `painel.py` | Monta `triagem/painel.json`: um card por email pendente com coluna automática (Itens Enviados por ConversationID) |
| `notion_mesa.py` | Cliente do database "Mesa de Emails" no Notion: sync dos cards (timeline por etapa, uploads ≤ 20 MiB, tags `Tema` por palavra-chave), comentários, mapa local `triagem/notion-mesa.json` |
| `mesa_poller.py` | Task Scheduler a cada 2 min, sem LLM no polling: lê comentários do usuário nos cards (manda / ajusta / faz / confirma / cancela / encerra / junta) e executa com `claude -p --resume <sessão do email>` ou `enviar.py`; envio em duas etapas (prévia do rascunho real → "confirma"). ajusta/faz rodam com ferramentas de sessão interativa (`--add-dir` na pasta do usuário); envio de email, token e web ficam bloqueados na sessão |
| `registrar_mesa_poller.ps1` | Registra a tarefa MesaPoller (pythonw, 2 min, logon, IgnoreNew, limite 35 min) |
| `rascunho.py` | Cria o rascunho de resposta (Reply/ReplyAll) a partir de um .txt, com a assinatura da conta certa |
| `enviar.py` | Envia UM rascunho existente do Outlook (guardrail: nunca compõe e envia direto) |
| `CLAUDE.md` | Regras da sessão: "bom dia" apresenta o resumo; fluxo de envio; segurança |

## Pré-filtro determinístico (sem modelo)

Descarta antes da triagem: newsletters/no-reply, threads onde a última mensagem é do usuário
(mapa `ConversationID → data` dos Itens Enviados), emails de remetente externo onde o usuário está só em Cc
sem ser citado no corpo (colega do mesmo domínio em Cc é mantido), notificações de sistemas (exceto aprovações pendentes) e remetentes listados
em `ignorados.txt` (fora do git). Redução típica observada: ~50% dos emails cortados antes do LLM.

## Decisões de segurança

- **Corpo de email é conteúdo não confiável.** Todos os prompts instruem a nunca executar
  instruções vindas de dentro de emails — só usar como dado.
- **Rotinas agendadas são somente-leitura no Outlook** (única exceção: aplicar uma categoria).
  `--allowedTools` restringe escrita às pastas de saída do próprio projeto.
- **Envio de email exige humano no loop**: o fluxo é rascunho no Outlook (com assinatura via
  `GetInspector`) → revisão no chat → confirmação explícita → `enviar.py "termo do assunto"`,
  que se recusa a enviar se o termo casar com mais de um rascunho.
- **Dados de email nunca entram no repositório** (`.gitignore` cobre JSONs, .msg, triagens e vaults).

## Truques de implementação que custaram caro descobrir

- `Items.Restrict` por data depende do locale do Windows → ordenar decrescente e cortar no cutoff.
- Rascunho criado via COM não recebe assinatura → `mail.GetInspector` força o Outlook a inserir.
- Rascunho aberto no painel de leitura não pode ser enviado via COM ("resposta embutida") →
  fechar no Outlook e reenviar, ou usar `GetItemFromID` para referência limpa.
- Task Scheduler não acha `pwsh` da Store → usar `powershell.exe` do sistema.
- PowerShell 5.1 lê `.ps1` UTF-8 sem BOM com acentos quebrados → scripts agendados em ASCII.
- Cache COM do pywin32 (`gen_py`) corrompe → apagar `%LOCALAPPDATA%\Temp\gen_py` resolve.
- `claude --remote-control` morre se a janela fechar → janela oculta (`-WindowStyle Hidden`) + PID
  guardado para encerrar a sessão do dia anterior.
- Notion API 2025-09-03: consulta em `POST /data_sources/{id}/query` (a de `/databases` está deprecada),
  página nova com `parent.data_source_id`, propriedade tipo `status` não é criável pela API (usar `select`),
  comentários exigem as capacidades "Read/Insert comments" na integração (senão 403), views não se criam
  pela API pública (board e lista configurados uma vez no Notion).
- `~/.claude/settings.json` libera `enviar.py` por caminho absoluto em toda sessão; as sessões headless
  negam explicitamente esses padrões com `--disallowedTools`.

## Evolução (changelog de estratégia)

- **v1** — extrator + pré-filtro + validação COM nas duas contas.
- **v2** — camada `claude -p`: triagem priorizada com rascunhos de resposta + EntryID.
- **v3** — vaults de contexto por empresa (Obsidian) + histórico de 1 ano resumido em janelas
  escaladas (3d/1sem/1m/3m/1ano) por agentes paralelos.
- **v4** — envio com humano no loop (`enviar.py` + permissão explícita no settings).
- **v5** — categoria automática por assunto (tag "Bibliotecas") no extrator.
- **v6** — briefings diários por projeto (cada projeto com CLAUDE.md + plano vivo).
- **v7** — resumo consolidado de celular + sessão do dia em remote-control aberta por agendamento.
- **v8** — resumo publicado no Notion; uma sessão por email pendente (dossiê + rascunho no Outlook),
  aberta no logon e às 13h, para o Luan só decidir em vez de tratar um a um.
- **v9** — Mesa de Emails: database no Notion (um card por email, timeline por etapa, tags por tema,
  kanban por empresa) + poller local que executa comentários do usuário com `claude -p --resume` e
  envia em duas etapas; "faz" com ferramentas de sessão interativa. Substitui o artifact.
