/* =====================================================================================
   QUOTE FOLLOW UP — importação das planilhas para o MariaDB
   Lê "Quote Follow Up data.xlsx" e "Quote Follow Up reports.xlsx" com o PRÓPRIO core.js do programa
   (mesmas regras de leitura, inclusive a aba Config cifrada dos usuários) e gera o arquivo import.sql.

   Como usar (na pasta do programa, onde estão core.js e node_modules — rode o build.bat uma vez antes):
     node import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx"
     mariadb -u root -p < import.sql

   Pode rodar quantas vezes quiser: o import.sql atualiza o que já existe e não duplica nada.
   O histórico (activities) só ganha os registros que ainda não estão no banco.
   ===================================================================================== */
import fs from "fs";
import path from "path";
import crypto from "crypto";
import { fileURLToPath, pathToFileURL } from "url";
import * as XLSX from "xlsx";

const here = path.dirname(fileURLToPath(import.meta.url));
const [dataPath, repPath, outArg] = process.argv.slice(2);
if (!dataPath || !repPath) {
  console.error('Uso: node import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx" [import.sql]');
  process.exit(1);
}
const OUT = outArg || path.join(here, "import.sql");

/* core.js é ESM com extensão .js: uma cópia temporária .mjs garante que o Node o carregue como módulo
   em qualquer versão (a cópia fica na mesma pasta para achar o node_modules/xlsx). */
const tmpCore = path.join(here, ".core-import.tmp.mjs");
fs.copyFileSync(path.join(here, "core.js"), tmpCore);
let core;
try {
  core = await import(pathToFileURL(tmpCore).href);
} finally {
  fs.rmSync(tmpCore, { force: true });
}

const readWb = (p) => XLSX.read(fs.readFileSync(p), { type: "buffer" });
const D = core.parseDataWorkbook(readWb(dataPath)); // follow-up, histórico, metas, usuários
const R = core.parseDataWorkbook(readWb(repPath)); // cotações e vendas
if (D.usersBad) console.warn("ATENÇÃO: a aba Config (usuários) foi alterada por fora e não pôde ser lida: os usuários ficam de fora.");

const quotesRep = (R.quotesRep && R.quotesRep.quotes.length ? R.quotesRep : D.quotesRep) || { quotes: [] };
const quotes = quotesRep.quotes;
const salesRows = R.salesRows.length ? R.salesRows : D.salesRows;
const salesInfo = R.salesInfo && R.salesInfo.fileName ? R.salesInfo : D.salesInfo;
const meta = { ...(R.meta || {}), ...(D.meta || {}) };
const excl = new Set([...(R.excl || []), ...(D.excl || [])]);
// histórico dos dois arquivos, sem repetir (mesma chave do programa)
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

/* ---------- SQL ---------- */
const esc = (s) => "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "''").replace(/\0/g, "") + "'";
const txt = (v) => (v == null ? "''" : esc(v));
const nul = (v) => (v == null || v === "" ? "NULL" : esc(v));
const num = (v, d = 6) => (Number.isFinite(+v) ? (+v).toFixed(d) : "0");
const isoD = (v) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v || "").slice(0, 10)) ? esc(String(v).slice(0, 10)) : "NULL");
const isoT = (v) => (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(v || "")) ? esc(v) : "NULL");
const sha = (s) => crypto.createHash("sha256").update(s, "utf8").digest("hex");

const out = [];
const put = (s) => out.push(s);
// INSERT em blocos de 500 linhas (rápido e dentro do max_allowed_packet padrão)
function insert(table, cols, rows, tail = "", verb = "INSERT") {
  for (let i = 0; i < rows.length; i += 500) {
    put(`${verb} INTO ${table} (${cols.join(", ")}) VALUES\n` + rows.slice(i, i + 500).map((r) => "(" + r.join(", ") + ")").join(",\n") + (tail ? "\n" + tail : "") + ";");
  }
}
const upd = (cols) => "ON DUPLICATE KEY UPDATE " + cols.map((c) => `${c} = VALUES(${c})`).join(", ");

put("-- Gerado por import-xlsx.mjs em " + core.nowStamp());
put(`-- data: ${path.basename(dataPath)} · reports: ${path.basename(repPath)}`);
put("SET NAMES utf8mb4;");
put("USE quote_follow_up;");
put("START TRANSACTION;");

/* cotações */
const qCols = ["quote_id", "quote_date", "customer", "customer_key", "account", "classification", "status", "quoted_by", "sales_person", "amount", "sales_amt", "amount_pt", "amount_cp", "amount_mh", "line_read", "first_seen", "updated_at"];
insert(
  "quotes",
  qCols,
  quotes.map((q) => [esc(q.id), isoD(q.d), txt(q.c), txt(q.nn), txt(q.conta), txt(q.cls), txt(q.st), txt(q.qb), txt(q.sp), num(q.a), num(q.sa), num((q.al || {}).PT || 0), num((q.al || {}).CP || 0), num((q.al || {}).MH || 0), q.lr ? "1" : "0", isoT(q.fs), isoT(q.up)]),
  upd(qCols.slice(1)),
);
const orders = quotes.flatMap((q) => (q.so || []).filter(Boolean).map((s) => [esc(q.id), esc(s)]));
insert("quote_orders", ["quote_id", "sales_id"], orders, "", "INSERT IGNORE");
const pgs = quotes.flatMap((q) => Object.entries(q.pg || {}).map(([k, v]) => [esc(q.id), esc(k), num(v, 2)]));
insert("quote_product_groups", ["quote_id", "product_group", "amount"], pgs, upd(["amount"]));

/* vendas */
insert(
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
  fCols,
  fIds.map((id) => {
    const m = meta[id] || {};
    const win = [25, 50, 75, 99].includes(+m.win) ? String(+m.win) : (m.win && winFixed++, "NULL");
    return [esc(id), txt(m.name || (quotes.find((q) => q.id === id) || {}).c || ""), nul(core.FUP[m.fup] ? m.fup : ""), isoD(m.since), win, excl.has(id) ? "1" : "0", m.step ? String(+m.step) : "NULL", isoD(m.at), txt(m.req || ""), nul(m.close === "lost" || m.close === "cancelled" ? m.close : ""), txt(m.why || "")];
  }),
  upd(fCols.slice(1)),
);
const doneDays = fIds.flatMap((id) => [...new Set(((meta[id] || {}).log || []).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d)))].map((d) => [esc(id), esc(d)]));
insert("followup_done_days", ["quote_id", "done_day"], doneDays, "", "INSERT IGNORE");

/* histórico (só entra o que ainda não está no banco) */
const vtxt = (v) => (v == null || v === "" ? "NULL" : esc(String(v)));
insert(
  "activities",
  ["act_key", "at", "quote_id", "customer", "kind", "from_value", "to_value", "note", "created_by"],
  acts.map((a) => [esc(sha(core.actKey(a))), isoT(a.t), esc(a.id), txt(a.c), esc(a.k), vtxt(a.o), vtxt(a.n), vtxt(a.txt), txt(a.by)]),
  "",
  "INSERT IGNORE",
);

/* metas e informações gerais */
insert("day_goals", ["day", "goal"], Object.entries(dayGoals).map(([d, v]) => [esc(d), num(v, 2)]), upd(["goal"]));
const settings = [
  ["daily_goal", String(dailyGoal || 0)],
  ["last_quote_report", quotesRep.fileName || ""],
  ["quote_report_imported_at", quotesRep.at || ""],
  ["last_sales_report", (salesInfo && salesInfo.fileName) || ""],
  ["sales_report_imported_at", (salesInfo && salesInfo.at) || ""],
  ["program_version", D.ver || R.ver || ""],
];
insert("settings", ["name", "value"], settings.map(([k, v]) => [esc(k), esc(v)]), upd(["value"]));

/* usuários: o papel antigo "seller" vira sem acesso (o Admin escolhe a tela de cada um) */
const ROLE_OK = new Set(["admin", "manager", "outside", "inside"]);
const uRows = Object.entries(users).map(([name, u]) => [esc(name), ROLE_OK.has(u.role) ? esc(u.role) : "NULL", txt(u.salt || ""), txt(u.hash || ""), isoT(u.at)]);
// vale a mudança mais recente (changed_at por último: as atribuições do UPDATE são feitas da esquerda para a direita)
const newer = "(VALUES(changed_at) IS NOT NULL AND (changed_at IS NULL OR VALUES(changed_at) >= changed_at))";
insert(
  "users",
  ["name", "role", "pass_salt", "pass_hash", "changed_at"],
  uRows,
  `ON DUPLICATE KEY UPDATE role = IF(${newer}, VALUES(role), role), pass_salt = IF(${newer}, VALUES(pass_salt), pass_salt), pass_hash = IF(${newer}, VALUES(pass_hash), pass_hash), changed_at = IF(${newer}, VALUES(changed_at), changed_at)`,
);

put("COMMIT;");
fs.writeFileSync(OUT, out.join("\n\n") + "\n");

const roles = {};
Object.values(users).forEach((u) => {
  const r = ROLE_OK.has(u.role) ? u.role : "sem acesso";
  roles[r] = (roles[r] || 0) + 1;
});
console.log(`Pronto! ${OUT} (${(fs.statSync(OUT).size / 1024 / 1024).toFixed(1)} MB)`);
console.log(`  cotações ${quotes.length} · pedidos ${orders.length} · product groups ${pgs.length} · vendas ${salesRows.length}`);
console.log(`  follow-ups ${fIds.length} · dias com Done ${doneDays.length} · histórico ${acts.length} · metas por dia ${Object.keys(dayGoals).length}`);
console.log(`  usuários ${Object.keys(users).length} (${Object.entries(roles).map(([k, v]) => `${k}: ${v}`).join(", ")})`);
if (winFixed) console.log(`  ${winFixed} % Win fora de 25/50/75/99 ficaram vazios`);
