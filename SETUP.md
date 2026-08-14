# Setup — Painel de Produção Revita

## 1. O que já está pronto
- Worker (Cloudflare) em `src/`, painel estático em `public/`.
- Banco D1 com 3 tabelas: `apontamentos`, `metas`, `sync_state` (`migrations/0001_init.sql`).
- Sincronização automática com o Excel do SharePoint via Microsoft Graph, checando a cada 1 minuto (cron trigger) se o arquivo mudou, com botão de "Atualizar" manual e upload manual como caminho alternativo. O painel recarrega os dados do navegador a cada 30 segundos — na prática, uma alteração salva na planilha aparece no painel em até ~1-2 minutos.
- Testado localmente com `wrangler dev`: API, filtros por turma/desaguadora, metas, upload e exportação para Excel — todos validados com dados de exemplo.

## 2. Pendências que só você resolve
1. **Registro do app no Azure AD** (necessário para a sincronização automática — veja passo a passo abaixo). Sem isso, o painel funciona só com upload manual do Excel pelo gestor.
2. **Confirmar nomes de coluna reais da planilha** — o parser (`src/lib/parseExcel.ts`) já reconhece "Lote", "Cliente", "Número do Fardo", "TURMA", "Soma de Peso Seco 51%", "Data", "Hora do Apontamento", "Máquina"/"Desaguadora", "Produto" (com variações). Se algum nome real for diferente, é só me avisar ou ajustar a lista `HEADER_ALIASES`.
3. **Plano do Cloudflare Workers — recomendo o plano Paid (US$ 5/mês).** Cron Trigger no plano Free tem só 10ms de CPU por execução; a checagem "o arquivo mudou?" é barata e cabe tranquilamente nisso, mas o processamento completo (baixar + interpretar a planilha + gravar no banco), que só roda quando o Excel realmente mudou, pode passar de 10ms se a planilha crescer — no plano Paid o limite sobe pra 30 segundos, então isso deixa de ser risco. Sem o upgrade, o pior cenário é a sincronização de um ciclo específico falhar e tentar de novo no minuto seguinte (não trava o painel, só atrasa a atualização).
4. **Deploy automático via GitHub Actions** — configurado (veja seção 7), mas exige que você cadastre um token da Cloudflare como secret no repositório do GitHub antes do primeiro push.

Cores já aplicadas: paleta oficial Forest (Verde Claro `#C5F249`, Verde Escuro `#0A252A`, Bege Kraft `#B19873`, Roxo Destaque `#C1B4EE`) em `public/styles.css`. Verde-claro e roxo são usados só como preenchimento (barras/arcos/badges com fundo escuro) porque têm baixo contraste como texto sobre fundo branco; o texto de leitura usa sempre verde-escuro.

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
```bash
npm install
npx wrangler login

# cria o banco D1 e copia o database_id retornado para wrangler.jsonc
npx wrangler d1 create revita-producao-db
# cole o database_id em wrangler.jsonc -> d1_databases[0].database_id

# aplica o schema no banco remoto
npx wrangler d1 migrations apply revita-producao-db --remote

# segredos (não fica no código, fica só na Cloudflare)
npx wrangler secret put MS_TENANT_ID
npx wrangler secret put MS_CLIENT_ID
npx wrangler secret put MS_CLIENT_SECRET
npx wrangler secret put MS_SHARE_URL   # cole o link do SharePoint (o mesmo compartilhado nesta conversa)
npx wrangler secret put ADMIN_TOKEN    # senha que o gestor vai usar para editar metas / upload / forçar sync
```

## 5. Deploy
```bash
npx wrangler deploy
```
O Worker sobe com o painel estático e a API no mesmo domínio (`*.workers.dev` ou um domínio próprio, se configurado depois). O cron de sincronização (`* * * * *`, a cada 1 minuto) começa a rodar automaticamente após o deploy.

## 6. Testar localmente antes de mandar pra produção
```bash
npx wrangler d1 execute revita-producao-db --local --file ./migrations/0001_init.sql
npx wrangler dev
```
Crie um `.dev.vars` (não versionado) com `ADMIN_TOKEN` e valores fictícios de `MS_*` para testar a UI sem depender do Azure AD.

## 7. Deploy automático via GitHub Actions
A cada `git push` na branch `main`, o workflow `.github/workflows/deploy.yml` publica o Worker sozinho (`cloudflare/wrangler-action`). Ele **não** roda migração de banco — mudanças em `migrations/` continuam aplicadas manualmente (passo 4), pra nunca alterar dados de produção sem você revisar antes.

Antes do primeiro push, cadastre 2 secrets no repositório do GitHub (**Settings → Secrets and variables → Actions → New repository secret**):

1. `CLOUDFLARE_API_TOKEN` — crie em [dash.cloudflare.com/profile/api-tokens](https://dash.cloudflare.com/profile/api-tokens) → **Create Token** → use o template **"Edit Cloudflare Workers"** (já vem com a permissão certa: `Account.Workers Scripts:Edit`).
2. `CLOUDFLARE_ACCOUNT_ID` — `7ce2347a5c26e50ba4ff191d3a173084` (conta usada para hospedar este painel, dash.cloudflare.com/7ce2347a5c26e50ba4ff191d3a173084).

Os outros segredos (`MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_SHARE_URL`, `ADMIN_TOKEN`) **não** entram no GitHub — eles ficam só na Cloudflare (`wrangler secret put`, passo 4), porque é lá que o Worker roda e os lê.

## 8. Decisões de design (para validar com você)
- **"Turno" = a turma selecionada.** A planilha não tem uma coluna separada de "turno" (janela de horário); o que existe é a coluna TURMA. Então "Produção Total do Turno" = produção da turma selecionada no dia; "Produção Total do Dia" = todas as turmas somadas. Se isso não bater com a operação real, me diga como o turno deveria ser calculado.
- **Meta do turno = Meta do dia ÷ turnos por dia** (padrão 4) e **meta por hora = meta do turno ÷ horas por turno** (padrão 6), exatamente como pedido ("metas de hora/dia/turno baseadas na meta do dia"). Esses dois divisores (turnos por dia, horas por turno) ficam editáveis junto com a meta do dia na área do gestor.
- **Meta por desaguadora = meta do turno ÷ número de desaguadoras ativas** (divisão igual entre as 4).
- **Sincronização:** o cron roda a cada 1 minuto e só baixa/reprocessa a planilha se o `lastModifiedDateTime` do arquivo mudou (evita trabalho à toa — a maioria das execuções só faz essa checagem barata). O botão "Atualizar" força a sincronização na hora. O painel também recarrega os dados do banco a cada 30 segundos sozinho, então qualquer atualização (planilha, upload manual, edição de meta) aparece no telão em no máximo ~1-2 minutos.
- **Deduplicação:** cada linha da planilha vira uma "impressão digital" (hash de lote + fardo + máquina + data/hora). Isso deixa a sincronização segura para rodar repetidamente sem duplicar dados, mesmo que o Excel seja reexportado do zero.
