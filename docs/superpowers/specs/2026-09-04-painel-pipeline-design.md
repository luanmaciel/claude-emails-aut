# Painel de pipeline dos emails (rascunho de desenho, 04/09/2026)

## Problema
As sessões por email criadas pelo `rodar_pendentes.ps1` (`claude -p --session-id`) não aparecem na
barra lateral do app Claude (o app só lista sessões próprias, IDs `local_...`). Elas só abrem pelo
terminal (`claude --resume` na pasta do projeto). O Luan quer uma interface única em vez de abrir
sessão a sessão.

## Proposta
Um Artifact (página publicada, privada) em formato kanban, alimentado pelos arquivos do pipeline:

Colunas: **Mapeado** (na triagem, sem sessão ainda) → **Aguardando minha ação** (dossiê + rascunho
prontos) → **Aguardando resposta deles** (enviado, esperando o contato) → **Encerrado**.

Cada card: empresa (ALFA/BETA), remetente, assunto, resumo do pedido, ação proposta, rascunho,
dias parado, link para o dossiê. Ordem ALFA primeiro, Beta depois.

Interação com o Claude pela própria página (capacidades do Artifact, carregar a skill
`artifact-capabilities` antes de construir):
- `db`: estado dos cards (coluna, notas, data de envio) compartilhado e persistente.
- Comentário "manda" / "ajusta ..." / "faz ..." em um card, enviado ao Claude: a sessão do dia
  (que fica vigiando o artifact) executa na máquina: `enviar.py`, `rascunho.py`, ação do dossiê.
  O envio continua exigindo o comando explícito do Luan; nada sai sozinho.
- Botão "resumo do dia" no topo mostra o RESUMO-DIA.

## Alimentação
- `publicar_painel.py` (novo) roda ao final do `rodar_pendentes.ps1`: lê `pendentes-<dia>.json`,
  `sessoes-<dia>.json`, dossiês em `triagem/pendentes/` e o mapa de "última resposta minha" do
  `extrair.py` para mover cards para "Aguardando resposta deles" e "Encerrado" (thread sem resposta
  minha pendente e sem email novo há N dias).
- Escrita no artifact pela ação `write_db` (batch) a partir da sessão do dia, ou republicação da
  página com os dados embutidos se a capacidade de db não estiver disponível para a conta.

## Decisões pendentes do Luan
- Encerrar card automaticamente após quantos dias sem movimento (sugestão: 7)?
- O painel substitui a página diária do Notion ou convive com ela?

## Fora de escopo
Remote control / sessões no celular. O painel já resolve o acesso pelo celular.
