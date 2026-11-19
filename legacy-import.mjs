/* =====================================================================================
   QUOTE FOLLOW UP — planilhas da versão anterior → comandos SQL do banco
   Lê "Quote Follow Up data.xlsx" (follow-ups, notas, histórico, metas, usuários) e/ou
   "Quote Follow Up reports.xlsx" (cotações e vendas) com o PRÓPRIO core.js do programa e devolve os
   INSERTs (sem USE / START TRANSACTION: quem chama decide). Usado por:
     • import-xlsx.mjs (importar-banco.bat) → grava o import.sql;
     • server.mjs (aba Data → "Importar da versão anterior") → roda direto no banco, numa transação.

   Regras (as mesmas nos dois caminhos):
     • histórico: só entra o que ainda não está no banco (mesma chave do programa); quem fez cada registro é mantido;
     • cotações: vale a versão mais recente (updated_at) — um arquivo antigo não volta o status de uma cotação;
     • usuários: vale a mudança mais recente (changed_at);
     • follow-ups e metas: mode "replace" = o arquivo vale (como o importar-banco.bat);
                           mode "keep"    = só acrescenta o que falta, sem mudar o que já está no banco.
   ===================================================================================== */
import crypto from "crypto";

const esc = (s) => "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "''").replace(/\0/g, "") + "'";
const txt = (v) => (v == null ? "''" : esc(v));
const nul = (v) => (v == null || v === "" ? "NULL" : esc(v));
const num = (v, d = 6) => (Number.isFinite(+v) ? (+v).toFixed(d) : "0");
const isoD = (v) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v || "").slice(0, 10)) ? esc(String(v).slice(0, 10)) : "NULL");
const isoT = (v) => (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(v || "")) ? esc(v) : "NULL");
const sha = (s) => crypto.createHash("sha256").update(s, "utf8").digest("hex");
const ROLE_OK = new Set(["admin", "manager", "outside", "inside"]);

/* dataWb / repWb: pastas de trabalho já lidas (XLSX.read); qualquer uma pode faltar.
   opts.mode: "replace" (padrão) | "keep" · opts.cur: settings atuais do banco ({name: value}), para não trocar
   o nome/data do último relatório por um mais antigo. Devolve { statements: [{ sql, what }], counts, warnings }. */
export function buildLegacyImport(core, dataWb, repWb, opts = {}) {
  const keep = opts.mode === "keep";
  const cur = opts.cur || {};
  const empty = { meta: {}, excl: [], acts: [], users: {}, dayGoals: {}, salesRows: [], quotesRep: null, salesInfo: null };
  const D = dataWb ? core.parseDataWorkbook(dataWb) : empty;
  const R = repWb ? core.parseDataWorkbook(repWb) : empty;
  const warnings = [];
  if (D.usersBad || R.usersBad) warnings.push("users-bad");

  const quotesRep = (R.quotesRep && R.quotesRep.quotes.length ? R.quotesRep : D.quotesRep) || { quotes: [] };
  const quotes = quotesRep.quotes || [];
  const salesRows = (R.salesRows || []).length ? R.salesRows : D.salesRows || [];
  const salesInfo = R.salesInfo && R.salesInfo.fileName ? R.salesInfo : D.salesInfo;
  const meta = { ...(R.meta || {}), ...(D.meta || {}) };
  const excl = new Set([...(R.excl || []), ...(D.excl || [])]);
  const actSeen = new Set();
  const acts = [...(D.acts || []), ...(R.acts || [])].filter((a) => {
    const k = core.actKey(a);
    if (actSeen.has(k)) return false;
    actSeen.add(k);
    return true;
  });
  const users = core.mergeUsers(R.users || {}, D.users || {});
  const dailyGoal = D.dailyGoal || R.dailyGoal || 0;
  const dayGoals = { ...(R.dayGoals || {}), ...(D.dayGoals || {}) };

  const statements = [];
  // INSERT em blocos de 500 linhas (rápido e dentro do max_allowed_packet padrão)
  function insert(what, table, cols, rows, tail = "", verb = "INSERT") {
    for (let i = 0; i < rows.length; i += 500)
      statements.push({ what, sql: `${verb} INTO ${table} (${cols.join(", ")}) VALUES\n` + rows.slice(i, i + 500).map((r) => "(" + r.join(", ") + ")").join(",\n") + (tail ? "\n" + tail : "") });
  }
  const upd = (cols) => "ON DUPLICATE KEY UPDATE " + cols.map((c) => `${c} = VALUES(${c})`).join(", ");

  /* cotações: só atualiza quando o arquivo é mais novo (updated_at é a última coluna: as outras comparam com o valor antigo) */
  const qCols = ["quote_id", "quote_date", "customer", "customer_key", "account", "classification", "status", "quoted_by", "sales_person", "amount", "sales_amt", "amount_pt", "amount_cp", "amount_mh", "line_read", "first_seen", "updated_at"];
  const qNewer = "(updated_at IS NULL OR VALUES(updated_at) IS NULL OR VALUES(updated_at) >= updated_at)";
  insert(
    "quotes",
    "quotes",
    qCols,
    quotes.map((q) => [esc(q.id), isoD(q.d), txt(q.c), txt(q.nn), txt(q.conta), txt(q.cls), txt(q.st), txt(q.qb), txt(q.sp), num(q.a), num(q.sa), num((q.al || {}).PT || 0), num((q.al || {}).CP || 0), num((q.al || {}).MH || 0), q.lr ? "1" : "0", isoT(q.fs), isoT(q.up)]),
    "ON DUPLICATE KEY UPDATE " + qCols.slice(1).map((c) => (c === "first_seen" ? `first_seen = COALESCE(first_seen, VALUES(first_seen))` : `${c} = IF(${qNewer}, VALUES(${c}), ${c})`)).join(", "),
  );
  const orders = quotes.flatMap((q) => (q.so || []).filter(Boolean).map((s) => [esc(q.id), esc(s)]));
  insert("orders", "quote_orders", ["quote_id", "sales_id"], orders, "", "INSERT IGNORE");
  const pgs = quotes.flatMap((q) => Object.entries(q.pg || {}).map(([k, v]) => [esc(q.id), esc(k), num(v, 2)]));
  insert("pgs", "quote_product_groups", ["quote_id", "product_group", "amount"], pgs, upd(["amount"]));

  /* vendas */
  insert(
    "sales",
    "sales",
    ["sales_id", "booked_date", "customer", "classification", "line", "booked", "salesperson"],
    salesRows.map((r) => [txt(r.so), isoD(r.d), txt(r.c), txt(r.cls), txt(["PT", "CP", "MH"].includes(r.ln) ? r.ln : ""), num(r.v), txt(r.sp)]),
    upd(["booked"]),
  );

  /* follow-ups (campos manuais) */
  let winFixed = 0;
  const fIds = [...new Set([...Object.keys(meta), ...excl])];
  const fCols = ["quote_id", "customer", "fup", "fup_since", "win", "left_out_of_plan", "cycle_step", "next_date", "requisition", "closed_as", "close_reason"];
  insert(
    "followups",
    "followups",
    fCols,
    fIds.map((id) => {
      const m = meta[id] || {};
      const win = [25, 50, 75, 99].includes(+m.win) ? String(+m.win) : (m.win && winFixed++, "NULL");
      return [esc(id), txt(m.name || (quotes.find((q) => q.id === id) || {}).c || ""), nul(core.FUP[m.fup] ? m.fup : ""), isoD(m.since), win, excl.has(id) ? "1" : "0", m.step ? String(+m.step) : "NULL", isoD(m.at), txt(m.req || ""), nul(m.close === "lost" || m.close === "cancelled" ? m.close : ""), txt(m.why || "")];
    }),
    keep ? "" : upd(fCols.slice(1)),
    keep ? "INSERT IGNORE" : "INSERT",
  );
  const doneDays = fIds.flatMap((id) => [...new Set(((meta[id] || {}).log || []).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d)))].map((d) => [esc(id), esc(d)]));
  insert("done", "followup_done_days", ["quote_id", "done_day"], doneDays, "", "INSERT IGNORE");

  /* histórico: só entra o que ainda não está no banco */
  const vtxt = (v) => (v == null || v === "" ? "NULL" : esc(String(v)));
  insert(
    "acts",
    "activities",
    ["act_key", "at", "quote_id", "customer", "kind", "from_value", "to_value", "note", "created_by"],
    acts.map((a) => [esc(sha(core.actKey(a))), isoT(a.t), esc(a.id), txt(a.c), esc(a.k), vtxt(a.o), vtxt(a.n), vtxt(a.txt), txt(a.by)]),
    "",
    "INSERT IGNORE",
  );

  /* metas */
  insert("goals", "day_goals", ["day", "goal"], Object.entries(dayGoals).map(([d, v]) => [esc(d), num(v, 2)]), keep ? "" : upd(["goal"]), keep ? "INSERT IGNORE" : "INSERT");
  const settings = [["program_version", D.ver || R.ver || ""]];
  if (dailyGoal && (!keep || !+cur.daily_goal)) settings.push(["daily_goal", String(dailyGoal)]);
  // nome e data do último relatório: só quando o do arquivo é mais novo que o do banco
  if (quotesRep.fileName && String(quotesRep.at || "") >= String(cur.quote_report_imported_at || ""))
    settings.push(["last_quote_report", quotesRep.fileName], ["quote_report_imported_at", quotesRep.at || ""]);
  if (salesInfo && salesInfo.fileName && String(salesInfo.at || "") >= String(cur.sales_report_imported_at || ""))
    settings.push(["last_sales_report", salesInfo.fileName], ["sales_report_imported_at", salesInfo.at || ""]);
  insert("settings", "settings", ["name", "value"], settings.map(([k, v]) => [esc(k), esc(v)]), upd(["value"]));

  /* usuários: o papel antigo "seller" vira sem acesso (o Admin escolhe a tela de cada um); vale a mudança mais recente */
  const uRows = Object.entries(users).map(([name, u]) => [esc(name), ROLE_OK.has(u.role) ? esc(u.role) : "NULL", txt(u.salt || ""), txt(u.hash || ""), isoT(u.at)]);
  const newer = "(VALUES(changed_at) IS NOT NULL AND (changed_at IS NULL OR VALUES(changed_at) >= changed_at))";
  insert(
    "users",
    "users",
    ["name", "role", "pass_salt", "pass_hash", "changed_at"],
    uRows,
    `ON DUPLICATE KEY UPDATE role = IF(${newer}, VALUES(role), role), pass_salt = IF(${newer}, VALUES(pass_salt), pass_salt), pass_hash = IF(${newer}, VALUES(pass_hash), pass_hash), changed_at = IF(${newer}, VALUES(changed_at), changed_at)`,
  );

  const roles = {};
  Object.values(users).forEach((u) => {
    const r = ROLE_OK.has(u.role) ? u.role : "none";
    roles[r] = (roles[r] || 0) + 1;
  });
  return {
    statements,
    warnings,
    counts: { quotes: quotes.length, orders: orders.length, pgs: pgs.length, sales: salesRows.length, followups: fIds.length, done: doneDays.length, acts: acts.length, goals: Object.keys(dayGoals).length, users: Object.keys(users).length, roles, winFixed },
  };
}
