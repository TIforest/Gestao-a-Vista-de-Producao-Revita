import type { Env } from "../types";
import { timingSafeEqual } from "./hash";

/** Gate simples para ações do gestor (editar metas, upload manual, forçar sync). */
export function isAdminRequest(request: Request, env: Env): boolean {
  const header = request.headers.get("Authorization") ?? "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (!token || !env.ADMIN_TOKEN) return false;
  return timingSafeEqual(token, env.ADMIN_TOKEN);
}
