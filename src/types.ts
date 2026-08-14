// Gerado manualmente porque o projeto ainda não tem `wrangler login` / D1 real.
// Assim que rodar `wrangler d1 create` e `wrangler login`, regenere com `wrangler types`
// e apague este arquivo (troque os imports para `./worker-configuration`).
export interface Env {
  DB: D1Database;
  ASSETS: Fetcher;

  // Secrets (wrangler secret put NOME)
  MS_TENANT_ID: string;
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  MS_SHARE_URL: string; // link de compartilhamento do SharePoint (o mesmo colado na conversa)
  MS_SHEET_NAME?: string; // opcional: nome da aba, senão usa a primeira
  ADMIN_TOKEN: string; // senha simples para liberar edição de metas/upload manual

  // Vars (wrangler.jsonc)
  TURNOS_POR_DIA_PADRAO: string;
  HORAS_POR_TURNO_PADRAO: string;
  DESAGUADORAS: string; // "01,02,03,04"
  TURMAS: string; // "A,B,C,D,E"
}

export interface Apontamento {
  lote: string;
  cliente: string;
  numero_fardo: number | null;
  turma: string;
  peso_seco: number;
  data_hora: string; // ISO local (America/Sao_Paulo), ex: 2026-08-14T13:38:00
  maquina: string; // "01".."04"
  produto: string;
  row_hash: string;
}
