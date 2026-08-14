import type { Env } from "./types";
import { buildDashboardPayload } from "./lib/dashboard";
import { runSync, upsertApontamentos } from "./lib/graphSync";
import { parseWorkbook } from "./lib/parseExcel";
import { upsertMetaDia, upsertMetaMes } from "./lib/metas";
import { buildExportWorkbook } from "./lib/exportXlsx";
import { isAdminRequest } from "./lib/authAdmin";
import { currentBrazilYearMonth, dayBoundsLocal, todayBrazilISODate } from "./lib/date";

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

      if (pathname === "/api/metas" && request.method === "POST") {
        if (!isAdminRequest(request, env)) return unauthorized();
        const body = (await request.json()) as {
          escopo: "dia" | "mes";
          referencia?: string;
          valor: number;
          turnosPorDia?: number;
          horasPorTurno?: number;
          updatedBy?: string;
        };
        if (typeof body.valor !== "number" || body.valor < 0) {
          return json({ erro: "Valor de meta inválido." }, { status: 400 });
        }
        if (body.escopo === "dia") {
          const referencia = body.referencia ?? todayBrazilISODate();
          await upsertMetaDia(
            env,
            referencia,
            body.valor,
            body.turnosPorDia ?? (Number(env.TURNOS_POR_DIA_PADRAO) || 4),
            body.horasPorTurno ?? (Number(env.HORAS_POR_TURNO_PADRAO) || 6),
            body.updatedBy ?? "gestor"
          );
        } else if (body.escopo === "mes") {
          const referencia = body.referencia ?? currentBrazilYearMonth();
          await upsertMetaMes(env, referencia, body.valor, body.updatedBy ?? "gestor");
        } else {
          return json({ erro: "Escopo inválido, use 'dia' ou 'mes'." }, { status: 400 });
        }
        return json({ ok: true });
      }

      if (pathname === "/api/sync" && request.method === "POST") {
        if (!isAdminRequest(request, env)) return unauthorized();
        const result = await runSync(env, { force: true });
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
        const { rows, warnings } = await parseWorkbook(buffer, env.MS_SHEET_NAME);
        await upsertApontamentos(env, rows);
        await env.DB.prepare(
          `UPDATE sync_state SET last_sync_at = datetime('now'), last_sync_status = 'sincronizado_manual',
           last_sync_rows = ?, last_error = NULL WHERE id = 1`
        )
          .bind(rows.length)
          .run();
        return json({ ok: true, linhas: rows.length, warnings });
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
