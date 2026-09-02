import type { Env } from "./types";
import { buildDashboardPayload } from "./lib/dashboard";
import { runSync, upsertApontamentos, RETENTION_DAYS } from "./lib/graphSync";
import { parseWorkbook } from "./lib/parseExcel";
import { buildExportWorkbook } from "./lib/exportXlsx";
import { isAdminRequest } from "./lib/authAdmin";
import { dayBoundsLocal, daysAgoLocalISOStart } from "./lib/date";

function json(data: unknown, init?: ResponseInit): Response {
  return new Response(JSON.stringify(data), {
    ...init,
    headers: { "Content-Type": "application/json; charset=utf-8", ...init?.headers },
  });
}

function unauthorized(): Response {
  return json({ erro: "Não autorizado. Informe o token do gestor." }, { status: 401 });
}

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const { pathname } = url;

    try {
      if (pathname === "/api/dashboard" && request.method === "GET") {
        const payload = await buildDashboardPayload(env, {
          turma: url.searchParams.get("turma") ?? undefined,
          maquina: url.searchParams.get("maquina") ?? undefined,
          date: url.searchParams.get("data") ?? undefined,
        });
        return json(payload);
      }

      // Sem token: é só "checar agora", usado pelo botão Atualizar do painel
      // (que qualquer um no chão de fábrica pode clicar) e pelo auto-sync
      // periódico do navegador. Upload manual e o /api/debug continuam
      // protegidos, esse não precisa.
      //
      // ?force=1 (só o clique manual do botão) pula a checagem barata de
      // "o arquivo mudou desde a última vez?" e força reler mesmo sem
      // mudança. O auto-sync periódico NÃO usa force: sem ele, quando nada
      // mudou no SharePoint a checagem é barata (1 consulta ao Graph); com
      // force sempre reescreveria as linhas da janela inteira no D1 a cada
      // chamada, mesmo sem nada novo — foi isso que quase estourou a cota
      // do banco quando o auto-sync passou a rodar de 1 em 1 minuto.
      if (pathname === "/api/sync" && request.method === "POST") {
        const force = url.searchParams.get("force") === "1";
        const result = await runSync(env, { force });
        return json(result);
      }

      if (pathname === "/api/upload" && request.method === "POST") {
        if (!isAdminRequest(request, env)) return unauthorized();
        const form = await request.formData();
        const file = form.get("arquivo");
        if (!(file instanceof File)) {
          return json({ erro: "Envie o arquivo no campo 'arquivo' (multipart/form-data)." }, { status: 400 });
        }
        const buffer = await file.arrayBuffer();
        const windowStart = daysAgoLocalISOStart(RETENTION_DAYS);
        const { rows, warnings } = await parseWorkbook(buffer, env.MS_SHEET_NAME, windowStart);
        await upsertApontamentos(env, rows);
        await env.DB.prepare(
          `UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'sincronizado_manual',
           last_sync_rows = ?, last_error = NULL WHERE id = 1`
        )
          .bind(rows.length)
          .run();
        return json({ ok: true, linhas: rows.length, warnings });
      }

      if (pathname === "/api/debug" && request.method === "GET") {
        if (!isAdminRequest(request, env)) return unauthorized();
        const apontamentos = await env.DB.prepare(
          "SELECT COUNT(*) AS total, MIN(data_hora) AS mais_antigo, MAX(data_hora) AS mais_novo FROM apontamentos"
        ).first();
        const producaoMensal = await env.DB.prepare("SELECT * FROM producao_mensal ORDER BY ano_mes").all();
        const syncState = await env.DB.prepare("SELECT * FROM sync_state WHERE id = 1").first();
        return json({ apontamentos, producaoMensal: producaoMensal.results, syncState });
      }

      if (pathname === "/api/export" && request.method === "GET") {
        const filters = {
          turma: url.searchParams.get("turma") ?? undefined,
          maquina: url.searchParams.get("maquina") ?? undefined,
          date: url.searchParams.get("data") ?? undefined,
        };
        const payload = await buildDashboardPayload(env, filters);
        const { start, end } = dayBoundsLocal(payload.data);
        const conds = ["data_hora >= ?", "data_hora < ?"];
        const args: (string | number)[] = [start, end];
        if (payload.filtros.turma) {
          conds.push("turma = ?");
          args.push(payload.filtros.turma);
        }
        if (payload.filtros.maquina) {
          conds.push("maquina = ?");
          args.push(payload.filtros.maquina);
        }
        const rows = await env.DB.prepare(
          `SELECT lote, cliente, numero_fardo, turma, peso_seco, data_hora, maquina, produto
           FROM apontamentos WHERE ${conds.join(" AND ")} ORDER BY data_hora DESC`
        )
          .bind(...args)
          .all();
        const workbook = buildExportWorkbook(payload, rows.results as Record<string, unknown>[]);
        return new Response(workbook, {
          headers: {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": `attachment; filename="revita-producao-${payload.data}.xlsx"`,
          },
        });
      }

      return json({ erro: "Rota não encontrada." }, { status: 404 });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return json({ erro: message }, { status: 500 });
    }
  },

  async scheduled(_event: ScheduledController, env: Env, ctx: ExecutionContext): Promise<void> {
    ctx.waitUntil(runSync(env).then(() => undefined));
  },
} satisfies ExportedHandler<Env>;
