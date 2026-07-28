# Triagem diária de emails corporativos com Claude Code (Outlook local, sem OAuth)

Rotina local que transforma duas caixas corporativas do Outlook Desktop em um assistente diário: triagem priorizada, memória de contexto por empresa, briefings por projeto e uma sessão de chat acessível do celular — tudo sem Graph API, sem OAuth e sem add-in aprovado por TI (cenário comum: contas corporativas sem acesso de administrador).

## Arquitetura

```
07:30  extrair.py (pywin32/COM) ──> emails.json ──> claude -p (triagem-prompt.md)
                                                     ├─> triagem/AAAA-MM-DD.md   (AÇÃO HOJE / SEMANA / MONITORAR)
                                                     └─> contexto/<empresa>/     (vaults Obsidian por empresa)
07:45  claude -p por projeto (BRIEFING-PROMPT.md) ──> BRIEFING-DIA.md + plano atualizado
08:05  claude -p (resumo-prompt.md) ──> triagem/RESUMO-DIA.md (formato celular)
08:10  claude --remote-control "DD/MM - Resumo diário..." (janela oculta)
       └─> sessão do dia acessível pelo app do celular: "bom dia" → resumo → direcionamento
```

Tudo agendado no Task Scheduler do Windows (`schtasks`), rodando com a sessão do usuário logada.

## Componentes

| Arquivo | Papel |
|---|---|
| `extrair.py` | Lê as caixas via COM/MAPI (sem LLM), aplica pré-filtro determinístico, gera JSON |
| `triagem-prompt.md` | Prompt do `claude -p` diário: triagem priorizada + atualização dos vaults de contexto |
| `resumo-prompt.md` | Consolida triagem + briefings num resumo de celular |
| `rodar_triagem.ps1` / `rodar_briefing.ps1` / `rodar_resumo.ps1` | Runners agendados |
| `abrir_sessao_dia.ps1` | Abre a sessão do dia em remote-control (oculta), mata a de ontem |
| `enviar.py` | Envia UM rascunho existente do Outlook (guardrail: nunca compõe e envia direto) |
| `CLAUDE.md` | Regras da sessão: "bom dia" apresenta o resumo; fluxo de envio; segurança |

## Pré-filtro determinístico (sem modelo)

Descarta antes da triagem: newsletters/no-reply, threads onde a última mensagem é do usuário
(mapa `ConversationID → data` dos Itens Enviados), emails onde o usuário está só em Cc sem ser
citado no corpo, notificações de sistemas (exceto aprovações pendentes) e remetentes listados
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

## Evolução (changelog de estratégia)

- **v1** — extrator + pré-filtro + validação COM nas duas contas.
- **v2** — camada `claude -p`: triagem priorizada com rascunhos de resposta + EntryID.
- **v3** — vaults de contexto por empresa (Obsidian) + histórico de 1 ano resumido em janelas
  escaladas (3d/1sem/1m/3m/1ano) por agentes paralelos.
- **v4** — envio com humano no loop (`enviar.py` + permissão explícita no settings).
- **v5** — categoria automática por assunto (tag "Bibliotecas") no extrator.
- **v6** — briefings diários por projeto (cada projeto com CLAUDE.md + plano vivo).
- **v7** — resumo consolidado de celular + sessão do dia em remote-control aberta por agendamento.
