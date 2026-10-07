# Setup — Painel de Produção Revita

## 1. O que já está pronto
- Worker (Cloudflare) em `src/`, painel estático em `public/`.
- Banco D1: `apontamentos` (janela recente, ver `RETENTION_DAYS`), `producao_diaria` (um total por dia, recalculado da planilha — a produção do mês é a soma dos dias), `sync_state`. As tabelas `metas` e `producao_mensal` existem no schema mas não são mais usadas — metas viraram valores fixos em `src/lib/metasFixas.ts`, e o contador mensal foi substituído pelos totais diários em 2026-10-07.
- Sincronização automática com o Excel do SharePoint via Microsoft Graph, checando a cada 1 minuto (cron trigger) se o arquivo mudou, com botão de "Atualizar" manual e upload manual como caminho alternativo. O painel recarrega os dados do navegador a cada 30 segundos — na prática, uma alteração salva na planilha aparece no painel em até ~1-2 minutos.
- Testado localmente com `wrangler dev`: API, filtros por turma/desaguadora, metas, upload e exportação para Excel — todos validados com dados de exemplo.

## 2. Pendências que só você resolve
1. ~~Registro do app no Azure AD~~ — feito, sincronização automática ativa.
2. **Confirmar nomes de coluna reais da planilha** — o parser (`src/lib/parseExcel.ts`) já reconhece "Lote", "Cliente", "Número do Fardo", "TURMA", "Soma de Peso Seco 51%", "Data", "Hora do Apontamento", "Máquina"/"Desaguadora", "Produto" (com variações). Se algum nome real for diferente, é só me avisar ou ajustar a lista `HEADER_ALIASES`.
3. ~~**Plano do Cloudflare Workers — recomendo o plano Paid (US$ 5/mês).**~~ Resolvido em 2026-08-31 com a migração pra conta PRO da Forest. Pendência atual: o Cron Trigger nessa conta não dispara (ver `SUPORTE_CLOUDFLARE.txt`); o navegador da TV faz o papel do cron chamando `/api/sync` a cada 1 min. Cron Trigger no plano Free tem só 10ms de CPU por execução; a checagem "o arquivo mudou?" é barata e cabe tranquilamente nisso, mas o processamento completo (baixar + interpretar a planilha + gravar no banco), que só roda quando o Excel realmente mudou, pode passar de 10ms se a planilha crescer — no plano Paid o limite sobe pra 30 segundos, então isso deixa de ser risco. Sem o upgrade, o pior cenário é a sincronização de um ciclo específico falhar e tentar de novo no minuto seguinte (não trava o painel, só atrasa a atualização).
4. **Deploy automático via GitHub Actions** — configurado (veja seção 7), mas exige que você cadastre um token da Cloudflare como secret no repositório do GitHub antes do primeiro push.

Cores: desde 2026-10-05 o painel segue o visual do BI original (ver seção 8); as variáveis estão no topo de `public/styles.css`.

## 3. Registrar o app no Azure AD (sincronização automática)
Precisa de alguém com permissão de administrador no Microsoft 365 / Azure AD da Forest.

1. Portal do Azure → **Azure Active Directory** → **App registrations** → **New registration**.
   - Nome: `revita-painel-producao-sync`
   - Tipo de conta: só a organização (single tenant)
   - Não precisa de Redirect URI (é autenticação app-only).
2. Anote o **Application (client) ID** e o **Directory (tenant) ID** da página Overview.
3. **Certificates & secrets** → **New client secret** → copie o valor (só aparece uma vez).
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → adicione **`Files.Read.All`** (ou `Sites.Read.All`, se preferir escopo mais amplo).
5. Clique em **Grant admin consent** (precisa ser um admin do tenant).

Esse app só recebe permissão de **leitura**; ele não escreve nada na planilha do SharePoint.

## 4. Criar os recursos no Cloudflare
Banco D1 `revita-producao-db` já criado (desde 2026-08-31 na conta PRO da Forest `ecc2e58aa45d10954dccbf42fd4d45ee`, login bruno.carvalho@forest.ind.br; antes ficava na conta pessoal `7ce2347a5c26e50ba4ff191d3a173084`), `database_id` já preenchido em `wrangler.jsonc`. Worker em produção: https://revita-painel-producao.chamados-push.workers.dev

**Atenção:** se for rodar comandos `wrangler` nesta máquina, confira antes com `npx wrangler whoami` — durante o setup encontramos uma sessão OAuth já logada aqui numa conta Cloudflare diferente ("Gustavo.oliveira@forest.ind.br's Account", `b780fecba22b416b28b057fef5fa7797`), não a `7ce2347a5c26e50ba4ff191d3a173084` deste projeto. Rode `npx wrangler login` de novo escolhendo a conta certa antes de qualquer comando `--remote`, ou prefira os workflows do GitHub Actions abaixo (que já usam o token certo via secret).

```bash
npm install

# aplica o schema no banco remoto — prefira rodar via GitHub Actions:
# aba Actions → "DB Migrate" → Run workflow (usa o secret CLOUDFLARE_API_TOKEN,
# já na conta certa). Alternativa local, só se tiver certeza da conta logada:
npx wrangler d1 migrations apply revita-producao-db --remote

# segredos (não fica no código, fica só na Cloudflare — precisa estar logado
# na conta certa pra rodar isso local)
npx wrangler secret put MS_TENANT_ID
npx wrangler secret put MS_CLIENT_ID
npx wrangler secret put MS_CLIENT_SECRET
npx wrangler secret put MS_SHARE_URL   # cole o link do SharePoint (o mesmo compartilhado nesta conversa)
npx wrangler secret put ADMIN_TOKEN    # protege /api/upload e /api/sync (upload manual e forçar sincronização)
```

## 5. Deploy
```bash
npx wrangler deploy
```
O Worker sobe com o painel estático e a API no mesmo domínio (`*.workers.dev` ou um domínio próprio, se configurado depois). O cron de sincronização (`* * * * *`, a cada 1 minuto) começa a rodar automaticamente após o deploy.

## 6. Testar localmente antes de mandar pra produção
```bash
npx wrangler d1 execute revita-producao-db --local --file ./migrations/0001_init.sql
npx wrangler d1 execute revita-producao-db --local --file ./scripts/seed_dev.sql   # dados fictícios, só local
npx wrangler dev
```
Crie um `.dev.vars` (não versionado) com `ADMIN_TOKEN` e valores fictícios de `MS_*` para testar a UI sem depender do Azure AD.

**Importante:** `scripts/seed_dev.sql` fica fora da pasta `migrations/` de propósito — qualquer `.sql` dentro de `migrations/` é tratado como migração de schema e pode ser aplicado no banco remoto (de produção) pelo workflow "DB Migrate". Nunca coloque dados fictícios lá.

## 7. Deploy automático via GitHub Actions
A cada `git push` na branch `main`, o workflow `.github/workflows/deploy.yml` publica o Worker sozinho (`cloudflare/wrangler-action`). Ele **não** roda migração de banco — mudanças em `migrations/` continuam aplicadas manualmente (passo 4), pra nunca alterar dados de produção sem você revisar antes.

Antes do primeiro push, cadastre 2 secrets no repositório do GitHub (**Settings → Secrets and variables → Actions → New repository secret**):

1. `CLOUDFLARE_API_TOKEN` — crie em [dash.cloudflare.com/profile/api-tokens](https://dash.cloudflare.com/profile/api-tokens) → **Create Token** → use o template **"Edit Cloudflare Workers"** (já vem com a permissão certa: `Account.Workers Scripts:Edit`).
2. `CLOUDFLARE_ACCOUNT_ID` — `ecc2e58aa45d10954dccbf42fd4d45ee` (conta PRO da Forest que hospeda este painel, dash.cloudflare.com/ecc2e58aa45d10954dccbf42fd4d45ee).

Os outros segredos (`MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_SHARE_URL`, `ADMIN_TOKEN`) **não** entram no GitHub — eles ficam só na Cloudflare (`wrangler secret put`, passo 4), porque é lá que o Worker roda e os lê.

## 8. Decisões de design (para validar com você)
- **"Turno" = a turma selecionada.** A planilha não tem uma coluna separada de "turno" (janela de horário); o que existe é a coluna TURMA. Então "Produção Total do Turno" = produção da turma selecionada no dia; "Produção Total do Dia" = todas as turmas somadas. Se isso não bater com a operação real, me diga como o turno deveria ser calculado.
- **Metas são fixas, definidas em `src/lib/metasFixas.ts`** (2026-08-25): não existe mais edição de meta pelo gestor no painel — é um número visual/informativo só, quem decide os valores é a operação. Meta por desaguadora = 10.000 cada (fixo, não é fração da meta do turno). Meta do turno = 40.000 (fixo). **Meta do dia = 40.000 × `TURNOS_POR_DIA`** (var no `wrangler.jsonc`; em 2026-10-05 são 2 turnos, turmas A e B, = 80.000, igual ao BI original). Meta por hora = 40.000 ÷ 6. Não existe meta do mês. Se a capacidade mudar, atualiza `metasFixas.ts`; se o número de turnos mudar, atualiza só a var — nos dois casos é `git push` e o deploy é automático.
- **Turmas listadas = as que têm apontamento no dia** (2026-10-05, igual ao segmentador do BI). A var `TURMAS` (A..E) é só fallback pra dia sem nenhum dado.
- **Turma em operação é automática** (2026-10-05): o painel segue a turma do *último apontamento do dia* (`turmaAtual`; o front pede `?turma=auto`). Quando entra o primeiro apontamento da turma seguinte, a tela vira pra ela sozinha — a TV não precisa de clique. Clicar numa turma trava nela por 10 min e depois volta ao automático. "Produção x Turma" e "% da meta" continuam mostrando todas as turmas do dia (a em operação destacada); gauge do turno, meta por hora, desaguadoras e a tabela seguem a turma em operação.
- **Visual segue o BI original** (2026-10-05): fundo branco com molduras pretas, verde vivo pra produção, azul-marinho pra meta, vermelho pra "falta meta", barras em par (produção sobre meta, mesma escala), gauges grandes com o número dentro do arco e "% da meta" com uma barra por turma. A paleta Forest (verde-claro/roxo) foi substituída a pedido.
- **Sincronização:** o cron roda a cada 1 minuto e só baixa/reprocessa a planilha se o `lastModifiedDateTime` do arquivo mudou (evita trabalho à toa — a maioria das execuções só faz essa checagem barata). O botão "Atualizar" força a sincronização na hora. O painel também recarrega os dados do banco a cada 30 segundos sozinho, então qualquer atualização na planilha aparece no telão em no máximo ~1-2 minutos.
- **Sincronização = espelho da planilha (2026-10-07).** Cada linha vira uma "impressão digital" (hash da linha inteira: lote, fardo, turma, máquina, data/hora, peso, cliente, produto; linhas idênticas ganham #1, #2 e contam as duas, como no BI). A cada leitura, nos dias que ela cobre por inteiro, o banco fica IGUAL à planilha: linhas novas entram, linhas apagadas ou corrigidas na planilha saem. Antes o sync só acrescentava — cada correção de lote/hora virava um "fantasma" que inflava a turma e o mês (6 fardos a mais na turma A e +8 t no mês em 07/10). Os totais diários (`producao_diaria`) são recalculados do mesmo jeito; o mês é a soma deles. A leitura normal pega as últimas 1500 linhas (~2 semanas); na primeira vez que `producao_diaria` está vazia o sync lê 6000 linhas sozinho pra montar o histórico. Pra forçar uma recontagem maior: `POST /api/recontar?linhas=8000` com o `ADMIN_TOKEN`.
