/* =====================================================================================
   QUOTE FOLLOW UP — servidor (API)
   Fica no meio do caminho entre a página e o MariaDB:
     navegador / celular  ──►  túnel Cloudflare  ──►  este servidor (http://localhost:8080)  ──►  MariaDB
   • entrega a página (QuoteFollowUp.html) e a API em /api/…;
   • faz o login (as senhas de hoje continuam valendo: o banco guarda o mesmo hash do programa);
   • lê e grava os dados no banco. O banco NUNCA fica exposto: só este servidor fala com ele.

   Como rodar (na pasta do programa, depois do build.bat):
     1. copie servidor\server.config.example.json para servidor\server.config.json e preencha a senha do banco;
     2. node servidor/server.mjs        (ou o iniciar-servidor.bat)
   Veja docs/SERVIDOR.md para deixar como serviço do Windows e ligar o túnel.
   Pastas: codigo/ (core.js) · pagina/ (QuoteFollowUp.html e ícones) · servidor/ (este arquivo e a configuração).
   ===================================================================================== */
import fs from "fs";
import path from "path";
import zlib from "zlib";
import crypto from "crypto";
import { fileURLToPath, pathToFileURL } from "url";
import express from "express";
import compression from "compression";
import * as mariadb from "mariadb";
import * as XLSX from "xlsx";
import { buildLegacyImport } from "./legacy-import.mjs";
import { createPush } from "./push.mjs";

const here = path.dirname(fileURLToPath(import.meta.url)); // pasta servidor/
const ROOT = path.join(here, ".."); // pasta do programa

/* ---------- configuração: servidor/server.config.json (ou, de versões antigas, na pasta do programa) ---------- */
const CFG_FILE = process.env.QFU_CONFIG || [path.join(here, "server.config.json"), path.join(ROOT, "server.config.json")].find((f) => fs.existsSync(f)) || path.join(here, "server.config.json");
if (!fs.existsSync(CFG_FILE)) {
  console.error(`Falta o arquivo servidor\\server.config.json.\nCopie servidor\\server.config.example.json para servidor\\server.config.json e preencha a senha do banco.`);
  process.exit(1);
}
if (path.dirname(CFG_FILE) === ROOT) console.warn("[aviso] O server.config.json está na pasta do programa: mova para a pasta servidor (funciona dos dois jeitos).");
const cfg = JSON.parse(fs.readFileSync(CFG_FILE, "utf8"));
const PORT = +cfg.port || 8080;
const HOST = cfg.host || "127.0.0.1"; // só a própria máquina (o túnel roda nela); "0.0.0.0" = toda a rede
const SESSION_DAYS = Math.max(1, +cfg.sessionDays || 30);
const PAGE = [cfg.page && path.join(ROOT, cfg.page), path.join(ROOT, "pagina", "QuoteFollowUp.html")].find((f) => f && fs.existsSync(f)) || path.join(ROOT, "pagina", "QuoteFollowUp.html");
const ICONS = path.join(ROOT, "pagina", "icones");

/* ---------- core.js: as MESMAS regras do programa (vendedores, chaves do histórico, senhas) ---------- */
const tmpCore = path.join(here, ".core-server.tmp.mjs");
fs.copyFileSync(path.join(ROOT, "codigo", "core.js"), tmpCore);
let core;
try {
  core = await import(pathToFileURL(tmpCore).href);
} finally {
  fs.rmSync(tmpCore, { force: true });
}
const { ADMIN_USER, PASS_ROUNDS, actKey, sellerOptions, sellerSame, quoteCancelled, normName, clean, FUP, ACT_TYPES, APP_VERSION } = core;

/* ---------- banco ---------- */
const pool = mariadb.createPool({
  host: (cfg.db && cfg.db.host) || "127.0.0.1",
  port: +(cfg.db && cfg.db.port) || 3306,
  user: (cfg.db && cfg.db.user) || "root",
  password: (cfg.db && cfg.db.password) || "",
  database: (cfg.db && cfg.db.database) || "quote_follow_up",
  connectionLimit: +(cfg.db && cfg.db.connections) || 6,
  charset: "utf8mb4",
  dateStrings: true, // DATE / DATETIME chegam como texto ("2026-11-17", "2026-11-17 09:30:00"), sem fuso
  decimalAsNumber: true,
  bigIntAsNumber: true,
  insertIdAsNumber: true,
});

/* ---------- utilidades ---------- */
const sha256 = (s) => crypto.createHash("sha256").update(String(s), "utf8").digest("hex");
// mesmo cálculo do hashPassword do core.js (SHA-256 repetido), só que com o crypto do Node (muito mais rápido)
const hashPwd = (pwd, salt) => {
  let h = sha256(salt + ":" + pwd);
  for (let i = 1; i < PASS_ROUNDS; i++) h = sha256(h + salt);
  return h;
};
const sameHex = (a, b) => typeof a === "string" && typeof b === "string" && a.length === b.length && a.length > 0 && crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
const str = (v) => (v == null ? "" : String(v).trim());
const arr = (v) => (Array.isArray(v) ? v : []);
const isObj = (v) => !!v && typeof v === "object" && !Array.isArray(v);
const cut = (v, n) => str(v).slice(0, n);
const RX_DAY = /^\d{4}-\d{2}-\d{2}$/;
const RX_TS = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
const RX_ID = /^[A-Z0-9][A-Z0-9._\-/]{0,19}$/;
const day = (v) => (RX_DAY.test(str(v).slice(0, 10)) ? str(v).slice(0, 10) : null);
const stamp = (v) => (RX_TS.test(str(v)) ? str(v) : null);
const qid = (v) => str(v).toUpperCase();
const num = (v) => (Number.isFinite(+v) ? +v : 0);
const ROLE_OK = new Set(["admin", "manager", "outside", "inside"]);
const MGR = new Set(["admin", "manager"]);
const WIN_OK = new Set([25, 50, 75, 99]);
const KINDS = new Set(Object.keys(ACT_TYPES));
const REPORT_KINDS = new Set(["status", "amount", "order"]); // vêm do Quote Report (importação da gerência)
const roleOf = (name, role) => (name === ADMIN_USER ? "admin" : ROLE_OK.has(role) ? role : "");
const httpErr = (status, error, message) => Object.assign(new Error(message || error), { status, code: error });
const wrap = (fn) => (req, res, next) => fn(req, res, next).catch(next);
const now = () => core.nowStamp();
const chunks = (list, n) => {
  const out = [];
  for (let i = 0; i < list.length; i += n) out.push(list.slice(i, i + n));
  return out;
};
const affected = (r) => (Array.isArray(r) ? r.reduce((s, x) => s + (x.affectedRows || 0), 0) : (r && r.affectedRows) || 0);

/* ---------- versões: cada gravação soma 1 (é assim que os outros sabem que precisam buscar) ---------- */
async function readVers(conn) {
  const rows = await (conn || pool).query("SELECT name, value FROM settings WHERE name IN ('data_version', 'rep_version')");
  const v = { data: 1, rep: 1 };
  rows.forEach((r) => (v[r.name === "rep_version" ? "rep" : "data"] = +r.value || 1));
  return v;
}

/* ---------- cotações e vendas (parte grande): montadas uma vez por versão e guardadas já compactadas ---------- */
let repCache = null;
let repBuilding = null;
async function buildRep() {
  const conn = await pool.getConnection();
  try {
    await conn.query("START TRANSACTION WITH CONSISTENT SNAPSHOT");
    const ver = (await readVers(conn)).rep;
    const qs = await conn.query(
      "SELECT quote_id, quote_date, customer, account, classification, status, quoted_by, sales_person, amount, sales_amt, amount_pt, amount_cp, amount_mh, line_read, first_seen, updated_at FROM quotes ORDER BY quote_date, quote_id",
    );
    const orders = await conn.query("SELECT quote_id, sales_id FROM quote_orders ORDER BY quote_id, sales_id");
    const pgs = await conn.query("SELECT quote_id, product_group, amount FROM quote_product_groups");
    const sales = await conn.query("SELECT sales_id, booked_date, customer, classification, line, booked, salesperson FROM sales ORDER BY booked_date, id");
    const st = Object.fromEntries((await conn.query("SELECT name, value FROM settings")).map((r) => [r.name, r.value]));
    await conn.query("COMMIT");
    const soBy = new Map();
    orders.forEach((o) => (soBy.get(o.quote_id) || soBy.set(o.quote_id, []).get(o.quote_id)).push(o.sales_id));
    const pgBy = new Map();
    pgs.forEach((p) => ((pgBy.get(p.quote_id) || pgBy.set(p.quote_id, {}).get(p.quote_id))[p.product_group] = p.amount));
    // formato curto: só os campos que têm valor (a página completa o resto)
    const quotes = qs.map((r) => {
      const q = { id: r.quote_id, d: r.quote_date, c: r.customer };
      if (r.account) q.conta = r.account;
      if (r.classification) q.cls = r.classification;
      if (r.status) q.st = r.status;
      if (r.quoted_by) q.qb = r.quoted_by;
      if (r.sales_person) q.sp = r.sales_person;
      if (r.amount) q.a = r.amount;
      if (r.sales_amt) q.sa = r.sales_amt;
      if (soBy.has(r.quote_id)) q.so = soBy.get(r.quote_id);
      const al = {};
      if (r.amount_pt) al.PT = r.amount_pt;
      if (r.amount_cp) al.CP = r.amount_cp;
      if (r.amount_mh) al.MH = r.amount_mh;
      if (Object.keys(al).length) q.al = al;
      if (r.line_read) q.lr = 1;
      if (pgBy.has(r.quote_id)) q.pg = pgBy.get(r.quote_id);
      if (r.first_seen) q.fs = r.first_seen;
      if (r.updated_at) q.up = r.updated_at;
      return q;
    });
    const salesRows = sales.map((r) => ({ so: r.sales_id, d: r.booked_date, c: r.customer, cls: r.classification, ln: r.line, v: r.booked, sp: r.salesperson }));
    const body = {
      ver,
      quotesRep: { fileName: st.last_quote_report || "", at: st.quote_report_imported_at || "", by: st.quote_report_by || "", fillAt: st.quote_fill_at || "", quotes },
      salesRows,
      salesInfo: { fileName: st.last_sales_report || "", at: st.sales_report_imported_at || "", by: st.sales_report_by || "" },
    };
    const json = JSON.stringify(body);
    const gz = zlib.gzipSync(json, { level: 6 });
    // para conferir o vendedor de cada cotação e montar a lista de nomes do login
    const byId = new Map(quotes.map((q) => [q.id, { sp: q.sp || "", qb: q.qb || "", c: q.c }]));
    const live = quotes.filter((q) => !quoteCancelled({ st: q.st || "" }));
    const spGroups = sellerOptions(live, "sp");
    const qbGroups = sellerOptions(live, "qb");
    console.info(`[rep] versão ${ver}: ${quotes.length} cotações, ${salesRows.length} vendas · ${(json.length / 1e6).toFixed(1)} MB → ${(gz.length / 1e6).toFixed(2)} MB compactado`);
    return { ver, json, gz, byId, spGroups, qbGroups };
  } catch (e) {
    try {
      await conn.query("ROLLBACK");
    } catch {}
    throw e;
  } finally {
    conn.release();
  }
}
async function getRep() {
  const v = (await readVers()).rep;
  if (repCache && repCache.ver === v) return repCache;
  if (!repBuilding)
    repBuilding = buildRep()
      .then((r) => (repCache = r))
      .finally(() => (repBuilding = null));
  return repBuilding;
}

/* nomes que podem entrar: como no programa — colunas Sales Person e Quoted By do Quote Report (cotações não
   canceladas) + as pessoas que o Admin acrescentou com um papel */
async function loginNames() {
  const rep = await getRep();
  const users = await pool.query("SELECT name, role, pass_hash FROM users");
  const sp = new Set(rep.spGroups.map((g) => g.name)),
    qb = new Set(rep.qbGroups.map((g) => g.name));
  const hand = users.filter((u) => u.name !== ADMIN_USER && !sp.has(u.name) && !qb.has(u.name) && ROLE_OK.has(u.role)).map((u) => u.name);
  const names = [...new Set([...sp, ...qb, ...hand])].sort((a, b) => a.localeCompare(b, "pt-BR", { sensitivity: "base" }));
  return { names, users, rep };
}

/* vendedor só mexe nas próprias cotações: externo = Sales Person, interno = Quoted By (mesmo agrupamento de
   grafias da tela; na dúvida, a mesma regra de nomes do programa) */
function scopeOf(me, rep) {
  if (MGR.has(me.role)) return () => true;
  // externo e interno: a carteira é sempre a do Sales Person (o Quoted By não define o que a pessoa vê)
  const field = "sp";
  const g = rep.spGroups.find((x) => x.name === me.name);
  const raw = new Set(g ? g.raw : []);
  const memo = new Map();
  return (id) => {
    const q = rep.byId.get(id);
    if (!q) return false;
    const v = clean(q[field]);
    if (!v) return false;
    if (!memo.has(v)) memo.set(v, raw.has(v) || sellerSame(me.name, v));
    return memo.get(v);
  };
}

/* ---------- avisos no celular / PC (servidor/push.mjs) ---------- */
const push = createPush({ pool, core, scopeOf, MGR, ADMIN_USER, proxy: str(cfg.proxy) }); // "proxy": só se a rede da empresa exigir (ex.: "http://proxy:3128")
push.setRepGetter(() => getRep());
// meta diária atual: [total, PT, CP, MH] (para saber se mudou de verdade numa gravação)
async function goalVec(conn) {
  const r = await conn.query("SELECT name, value FROM settings WHERE name IN ('daily_goal', 'daily_goal_pt', 'daily_goal_cp', 'daily_goal_mh')");
  const m = Object.fromEntries(r.map((x) => [x.name, Math.round(num(x.value))]));
  return [m.daily_goal || 0, m.daily_goal_pt || 0, m.daily_goal_cp || 0, m.daily_goal_mh || 0];
}

/* ---------- sessão (cookie) ---------- */
const COOKIE = "qfu_sid";
const cookies = (req) =>
  Object.fromEntries(
    String(req.headers.cookie || "")
      .split(";")
      .map((p) => p.trim().split("="))
      .filter((p) => p[0])
      .map(([k, ...v]) => [k, decodeURIComponent(v.join("="))]),
  );
const isHttps = (req) => req.secure || str(req.headers["x-forwarded-proto"]).split(",")[0] === "https";
function setCookie(req, res, token, maxAgeS) {
  res.append("Set-Cookie", `${COOKIE}=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAgeS}${isHttps(req) ? "; Secure" : ""}`);
}
async function newSession(req, res, name) {
  const token = crypto.randomBytes(32).toString("base64url");
  await pool.query("INSERT INTO sessions (token_hash, name, expires_at, user_agent) VALUES (?, ?, NOW() + INTERVAL ? DAY, ?)", [sha256(token), name, SESSION_DAYS, cut(req.headers["user-agent"], 200)]);
  setCookie(req, res, token, SESSION_DAYS * 86400);
}
async function sessionUser(req, res) {
  const token = cookies(req)[COOKIE];
  if (!token) return null;
  const th = sha256(token);
  const rows = await pool.query(
    "SELECT s.name, s.last_seen < NOW() - INTERVAL 10 MINUTE AS stale, u.role, u.pass_hash FROM sessions s JOIN users u ON u.name = s.name WHERE s.token_hash = ? AND s.expires_at > NOW()",
    [th],
  );
  if (!rows.length || !rows[0].pass_hash) return null;
  const r = rows[0];
  const user = { name: r.name, role: roleOf(r.name, r.role), th };
  if (r.stale) {
    // quem usa todo dia não precisa entrar de novo: a sessão vai sendo estendida
    await pool.query("UPDATE sessions SET last_seen = NOW(), expires_at = NOW() + INTERVAL ? DAY WHERE token_hash = ?", [SESSION_DAYS, th]);
    setCookie(req, res, token, SESSION_DAYS * 86400);
  }
  return user;
}
const auth = wrap(async (req, res, next) => {
  req.user = await sessionUser(req, res);
  if (!req.user) throw httpErr(401, "login", "Sign in again.");
  next();
});
const needRole = (req, res, next) => (req.user.role ? next() : next(httpErr(403, "no-access", "The admin has not chosen your screen yet.")));
const needMgr = (req, res, next) => (MGR.has(req.user.role) ? next() : next(httpErr(403, "forbidden", "Only the management can do this.")));
const needAdmin = (req, res, next) => (req.user.role === "admin" ? next() : next(httpErr(403, "forbidden", "Only the Admin can do this.")));

/* ---------- limite de tentativas de senha (por endereço e por nome) ---------- */
const tries = new Map();
const TRY_WINDOW = 15 * 60 * 1000;
const clientIp = (req) => str(req.headers["cf-connecting-ip"]) || req.ip || "";
function tooMany(req, name) {
  const t = Date.now();
  const keys = [`ip|${clientIp(req)}`, `nm|${clientIp(req)}|${name}`];
  const lim = [30, 8];
  return keys.some((k, i) => {
    const e = tries.get(k);
    return e && t - e.t0 < TRY_WINDOW && e.n >= lim[i];
  });
}
function failed(req, name) {
  const t = Date.now();
  [`ip|${clientIp(req)}`, `nm|${clientIp(req)}|${name}`].forEach((k) => {
    const e = tries.get(k);
    if (!e || t - e.t0 >= TRY_WINDOW) tries.set(k, { t0: t, n: 1 });
    else e.n++;
  });
}
setInterval(() => {
  const t = Date.now();
  for (const [k, e] of tries) if (t - e.t0 >= TRY_WINDOW) tries.delete(k);
}, 60000).unref();

/* ---------- parte "data": follow-ups, histórico, metas e usuários (pequena, vai inteira) ---------- */
const sanitizeUsers = (rows, me) => {
  const out = {};
  rows.forEach((u) => {
    if (me.role !== "admin" && u.name !== me.name) return; // cada um só vê o próprio cadastro; o Admin vê todos
    out[u.name] = { role: roleOf(u.name, u.role), pwd: !!u.pass_hash };
  });
  return out;
};
const NUM_KINDS = new Set(["win", "amount"]); // o programa grava esses valores como número
const actVal = (k, v) => (v == null ? undefined : NUM_KINDS.has(k) && /^-?\d+(\.\d+)?$/.test(v) ? +v : v);
async function readData(me) {
  const conn = await pool.getConnection();
  try {
    await conn.query("START TRANSACTION WITH CONSISTENT SNAPSHOT");
    const ver = (await readVers(conn)).data;
    const fups = await conn.query("SELECT quote_id, customer, fup, fup_since, win, left_out_of_plan, cycle_step, next_date, requisition, closed_as, close_reason FROM followups");
    const done = await conn.query("SELECT quote_id, done_day FROM followup_done_days ORDER BY quote_id, done_day");
    const acts = await conn.query("SELECT at, quote_id, customer, kind, from_value, to_value, note, created_by FROM activities ORDER BY at, id");
    const goals = await conn.query("SELECT day, goal FROM day_goals");
    const dg = await conn.query("SELECT value FROM settings WHERE name = 'daily_goal'");
    // metas por linha: settings daily_goal_pt / _cp / _mh → na página, dayGoals["line:PT"]…
    const lg = await conn.query("SELECT name, value FROM settings WHERE name IN ('daily_goal_pt', 'daily_goal_cp', 'daily_goal_mh')");
    const users = await conn.query("SELECT name, role, pass_hash FROM users");
    await conn.query("COMMIT");
    const logBy = new Map();
    done.forEach((d) => (logBy.get(d.quote_id) || logBy.set(d.quote_id, []).get(d.quote_id)).push(d.done_day));
    const meta = {},
      excl = [];
    fups.forEach((f) => {
      const m = {};
      if (f.customer) m.name = f.customer;
      if (f.fup) m.fup = f.fup;
      if (f.fup_since) m.since = f.fup_since;
      if (f.win) m.win = f.win;
      if (logBy.has(f.quote_id)) m.log = logBy.get(f.quote_id);
      if (f.cycle_step) m.step = f.cycle_step;
      if (f.next_date) m.at = f.next_date;
      if (f.requisition) m.req = f.requisition;
      if (f.closed_as) m.close = f.closed_as;
      if (f.close_reason) m.why = f.close_reason;
      if (Object.keys(m).some((k) => k !== "name")) meta[f.quote_id] = m;
      if (f.left_out_of_plan) excl.push(f.quote_id);
    });
    return {
      ver,
      meta,
      excl,
      acts: acts.map((r) => {
        const a = { t: r.at, id: r.quote_id, c: r.customer, k: r.kind };
        const o = actVal(r.kind, r.from_value),
          n = actVal(r.kind, r.to_value);
        if (o !== undefined) a.o = o;
        if (n !== undefined) a.n = n;
        if (r.note) a.txt = r.note;
        if (r.created_by) a.by = r.created_by;
        return a;
      }),
      dailyGoal: num(dg.length && dg[0].value),
      dayGoals: {
        ...Object.fromEntries(goals.map((g) => [g.day, g.goal])),
        ...Object.fromEntries(lg.filter((r) => num(r.value) > 0).map((r) => ["line:" + r.name.slice(-2).toUpperCase(), num(r.value)])),
      },
      users: sanitizeUsers(users, me),
    };
  } catch (e) {
    try {
      await conn.query("ROLLBACK");
    } catch {}
    throw e;
  } finally {
    conn.release();
  }
}

/* ---------- gravação ---------- */
async function lockVers(conn) {
  const rows = await conn.query("SELECT name, value FROM settings WHERE name IN ('data_version', 'rep_version') FOR UPDATE");
  const v = { data: 1, rep: 1 };
  rows.forEach((r) => (v[r.name === "rep_version" ? "rep" : "data"] = +r.value || 1));
  return v;
}
async function setSetting(conn, name, value) {
  await conn.query("INSERT INTO settings (name, value) VALUES (?, ?) ON DUPLICATE KEY UPDATE value = VALUES(value)", [name, String(value)]);
}
// campos de uma cotação no follow-up (o mesmo que o programa guardava na aba Follow-up)
function fupRow(m) {
  m = isObj(m) ? m : {};
  return {
    fup: FUP[m.fup] ? m.fup : null,
    fup_since: day(m.since),
    win: WIN_OK.has(+m.win) ? +m.win : null,
    cycle_step: Number.isInteger(+m.step) && +m.step > 0 && +m.step < 256 ? +m.step : null,
    next_date: day(m.at),
    requisition: cut(m.req, 255),
    closed_as: m.close === "lost" || m.close === "cancelled" ? m.close : null,
    close_reason: cut(m.why, 40),
    log: [...new Set(arr(m.log).map(day).filter(Boolean))],
  };
}
const fupEmpty = (r) => !r.fup && !r.fup_since && !r.win && !r.cycle_step && !r.next_date && !r.requisition && !r.closed_as && !r.close_reason && !r.log.length && !r.left_out_of_plan;

async function applyFollowups(conn, me, b, okId, rep, skipped) {
  const metaIn = isObj(b.meta) ? b.meta : {};
  const exAdd = new Set(arr(b.excl && b.excl.add).map(qid)),
    exDel = new Set(arr(b.excl && b.excl.del).map(qid));
  const metaBy = new Map(Object.entries(metaIn).map(([k, v]) => [qid(k), v]));
  const ids = [...new Set([...metaBy.keys(), ...exAdd, ...exDel])].filter((id) => RX_ID.test(id));
  const mine = ids.filter((id) => (okId(id) ? true : (skipped.add(id), false)));
  if (!mine.length) return 0;
  const cur = new Map();
  for (const part of chunks(mine, 500)) {
    const rows = await conn.query("SELECT * FROM followups WHERE quote_id IN (?) FOR UPDATE", [part]);
    rows.forEach((r) => cur.set(r.quote_id, r));
    const dd = await conn.query("SELECT quote_id, done_day FROM followup_done_days WHERE quote_id IN (?)", [part]);
    dd.forEach((d) => {
      const r = cur.get(d.quote_id);
      if (r) (r.log || (r.log = [])).push(d.done_day);
    });
  }
  let n = 0;
  for (const id of mine) {
    const c = cur.get(id);
    const hasMeta = metaBy.has(id);
    const m = hasMeta ? metaBy.get(id) : null;
    const base = hasMeta
      ? fupRow(m)
      : c
      ? { fup: c.fup, fup_since: c.fup_since, win: c.win, cycle_step: c.cycle_step, next_date: c.next_date, requisition: c.requisition, closed_as: c.closed_as, close_reason: c.close_reason, log: c.log || [] }
      : fupRow(null);
    base.left_out_of_plan = exAdd.has(id) ? 1 : exDel.has(id) ? 0 : c ? c.left_out_of_plan : 0;
    const customer = cut((isObj(m) && m.name) || (c && c.customer) || (rep.byId.get(id) || {}).c || "", 150);
    if (fupEmpty(base)) {
      if (c) await conn.query("DELETE FROM followups WHERE quote_id = ?", [id]);
    } else {
      await conn.query(
        `INSERT INTO followups (quote_id, customer, fup, fup_since, win, left_out_of_plan, cycle_step, next_date, requisition, closed_as, close_reason)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE customer = VALUES(customer), fup = VALUES(fup), fup_since = VALUES(fup_since), win = VALUES(win), left_out_of_plan = VALUES(left_out_of_plan),
           cycle_step = VALUES(cycle_step), next_date = VALUES(next_date), requisition = VALUES(requisition), closed_as = VALUES(closed_as), close_reason = VALUES(close_reason)`,
        [id, customer, base.fup, base.fup_since, base.win, base.left_out_of_plan, base.cycle_step, base.next_date, base.requisition, base.closed_as, base.close_reason],
      );
      if (hasMeta) {
        await conn.query("DELETE FROM followup_done_days WHERE quote_id = ?", [id]);
        if (base.log.length) await conn.batch("INSERT IGNORE INTO followup_done_days (quote_id, done_day) VALUES (?, ?)", base.log.map((d) => [id, d]));
      }
    }
    n++;
  }
  return n;
}

/* o banco aceita o tipo "dcs" (formulário DCS / Value Added)? Bancos criados antes precisam do schema.sql novo
   (a regra ck_activities_kind). Sem isso, o formulário NÃO é gravado — a página avisa e manda de novo depois. */
let dcsOk = null,
  dcsCheckedAt = 0;
async function dcsAllowed() {
  if (dcsOk || Date.now() - dcsCheckedAt < 60000) return dcsOk;
  dcsCheckedAt = Date.now();
  try {
    const r = await pool.query("SELECT CHECK_CLAUSE FROM information_schema.CHECK_CONSTRAINTS WHERE CONSTRAINT_SCHEMA = DATABASE() AND CONSTRAINT_NAME = 'ck_activities_kind'");
    dcsOk = !r.length || /'dcs'/.test(r[0].CHECK_CLAUSE);
  } catch {
    dcsOk = true;
  }
  return dcsOk;
}
async function applyActs(conn, me, b, okId, skipped, fresh) {
  const mgr = MGR.has(me.role);
  if (arr(b.acts).some((a) => isObj(a) && a.k === "dcs") && !(await dcsAllowed()))
    throw httpErr(503, "db-old", "The database needs an update to save the DCS / Value Added form: run the new banco/schema.sql (see docs/SERVIDOR.md). Nothing was lost: the page sends it again by itself.");
  const rows = [];
  let bad = 0;
  for (const a of arr(b.acts)) {
    if (!isObj(a)) continue;
    const id = qid(a.id),
      k = str(a.k),
      t = stamp(a.t);
    if (!t || !RX_ID.test(id) || !KINDS.has(k)) {
      bad++;
      continue;
    }
    const rk = REPORT_KINDS.has(k);
    if (rk ? !(mgr || b.rep) : !okId(id)) { // registros da importação: só junto com o relatório (ou da gerência)
      skipped.add(id);
      continue;
    }
    const v = (x) => (x == null || x === "" ? null : String(x).slice(0, 255));
    // a chave é a do programa (actKey): o mesmo registro mandado duas vezes não entra de novo
    rows.push([sha256(actKey(a)), t, id, cut(a.c, 150), k, v(a.o), v(a.n), a.txt ? String(a.txt) : null, rk ? cut(a.by, 100) : me.name]);
    if (fresh) fresh.push([rows[rows.length - 1][0], a]);
  }
  // avisos: só o que entra agora (o mesmo registro mandado de novo não avisa duas vezes)
  if (fresh && fresh.length) {
    const had = new Set();
    for (const part of chunks(fresh.map((x) => x[0]), 1000)) (await conn.query("SELECT act_key FROM activities WHERE act_key IN (?)", [part])).forEach((r) => had.add(r.act_key));
    const keep = fresh.filter((x) => !had.has(x[0])).map((x) => x[1]);
    fresh.length = 0;
    fresh.push(...keep);
  }
  let n = 0;
  for (const part of chunks(rows, 1000))
    n += affected(await conn.batch("INSERT IGNORE INTO activities (act_key, at, quote_id, customer, kind, from_value, to_value, note, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", part));
  if (bad) console.warn(`[sync] ${me.name}: ${bad} registro(s) de histórico inválido(s) ignorado(s)`);
  return n;
}

async function applyGoals(conn, b) {
  let n = 0;
  if (b.dailyGoal !== undefined && Number.isFinite(+b.dailyGoal) && +b.dailyGoal >= 0) {
    await setSetting(conn, "daily_goal", +b.dailyGoal);
    n++;
  }
  if (isObj(b.dayGoals))
    for (const [d0, v] of Object.entries(b.dayGoals)) {
      // meta diária de uma linha (PT, CP, MH)
      const ln = /^line:(PT|CP|MH)$/.exec(d0);
      if (ln) {
        const val = v == null ? 0 : +v;
        if (!Number.isFinite(val) || val < 0) continue;
        await setSetting(conn, "daily_goal_" + ln[1].toLowerCase(), val);
        n++;
        continue;
      }
      const d = day(d0);
      if (!d) continue;
      if (v == null) await conn.query("DELETE FROM day_goals WHERE day = ?", [d]);
      else if (Number.isFinite(+v) && (+v >= 0 || +v === -1)) await conn.query("INSERT INTO day_goals (day, goal) VALUES (?, ?) ON DUPLICATE KEY UPDATE goal = VALUES(goal)", [d, +v]);
      else continue;
      n++;
    }
  return n;
}

// importação dos relatórios (gerência): cotações novas/alteradas, cotações que saíram e as vendas inteiras
async function applyRep(conn, rep, byName = "") {
  const quotes = arr(rep.quotes).filter((q) => isObj(q) && RX_ID.test(qid(q.id)) && day(q.d));
  const qRows = quotes.map((q) => {
    const al = isObj(q.al) ? q.al : {};
    return [qid(q.id), day(q.d), cut(q.c, 150), cut(normName(q.c), 150), cut(q.conta, 20), cut(q.cls, 40), cut(q.st, 40), cut(q.qb, 100), cut(q.sp, 100), num(q.a), num(q.sa), num(al.PT), num(al.CP), num(al.MH), q.lr ? 1 : 0, stamp(q.fs), stamp(q.up)];
  });
  const cols = ["quote_id", "quote_date", "customer", "customer_key", "account", "classification", "status", "quoted_by", "sales_person", "amount", "sales_amt", "amount_pt", "amount_cp", "amount_mh", "line_read", "first_seen", "updated_at"];
  const ids = qRows.map((r) => r[0]);
  for (const part of chunks(qRows, 1000))
    await conn.batch(`INSERT INTO quotes (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")}) ON DUPLICATE KEY UPDATE ${cols.slice(1).map((c) => `${c} = VALUES(${c})`).join(", ")}`, part);
  for (const part of chunks(ids, 1000)) {
    await conn.query("DELETE FROM quote_orders WHERE quote_id IN (?)", [part]);
    await conn.query("DELETE FROM quote_product_groups WHERE quote_id IN (?)", [part]);
  }
  const so = quotes.flatMap((q) => [...new Set(arr(q.so).map((s) => cut(s, 20)).filter(Boolean))].map((s) => [qid(q.id), s]));
  for (const part of chunks(so, 1000)) await conn.batch("INSERT IGNORE INTO quote_orders (quote_id, sales_id) VALUES (?, ?)", part);
  const pg = quotes.flatMap((q) => Object.entries(isObj(q.pg) ? q.pg : {}).map(([k, v]) => [qid(q.id), cut(k, 80), Math.round(num(v) * 100) / 100]));
  for (const part of chunks(pg, 1000)) await conn.batch("INSERT INTO quote_product_groups (quote_id, product_group, amount) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE amount = VALUES(amount)", part);
  const removed = arr(rep.removed).map(qid).filter((id) => RX_ID.test(id));
  for (const part of chunks(removed, 1000)) await conn.query("DELETE FROM quotes WHERE quote_id IN (?)", [part]);
  let nSales = -1;
  if (Array.isArray(rep.sales)) {
    const rows = rep.sales
      .filter((r) => isObj(r) && day(r.d))
      .map((r) => [cut(r.so, 20), day(r.d), cut(r.c, 150), cut(r.cls, 40), ["PT", "CP", "MH"].includes(r.ln) ? r.ln : "", num(r.v), cut(r.sp, 100)]);
    await conn.query("DELETE FROM sales");
    for (const part of chunks(rows, 1000))
      await conn.batch("INSERT INTO sales (sales_id, booked_date, customer, classification, line, booked, salesperson) VALUES (?, ?, ?, ?, ?, ?, ?) ON DUPLICATE KEY UPDATE booked = VALUES(booked)", part);
    nSales = rows.length;
  }
  const info = isObj(rep.info) ? rep.info : {};
  if ("quoteFile" in info) await setSetting(conn, "last_quote_report", cut(info.quoteFile, 255));
  if ("quoteAt" in info) await setSetting(conn, "quote_report_imported_at", cut(info.quoteAt, 255));
  if ("quoteFile" in info && byName) await setSetting(conn, "quote_report_by", cut(byName, 255)); // quem importou (notificações)
  if ("fillAt" in info) await setSetting(conn, "quote_fill_at", cut(info.fillAt, 255));
  if ("salesFile" in info) await setSetting(conn, "last_sales_report", cut(info.salesFile, 255));
  if ("salesAt" in info) await setSetting(conn, "sales_report_imported_at", cut(info.salesAt, 255));
  if ("salesFile" in info && byName) await setSetting(conn, "sales_report_by", cut(byName, 255));
  await setSetting(conn, "program_version", APP_VERSION);
  return { quotes: qRows.length, removed: removed.length, sales: nSales };
}

/* ---------- aplicação web ---------- */
const app = express();
app.disable("x-powered-by");
app.set("trust proxy", "loopback"); // o túnel (cloudflared) conecta nesta mesma máquina
app.use((req, res, next) => {
  res.set({
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "same-origin",
    "Content-Security-Policy": "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; font-src 'self' data:; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
  });
  next();
});
app.use(compression());
// os POST precisam deste cabeçalho: outro site não consegue mandá-lo (proteção contra CSRF)
app.use("/api", (req, res, next) => {
  res.set("Cache-Control", "no-store");
  if (req.method !== "GET" && req.get("X-Requested-With") !== "qfu") return next(httpErr(403, "csrf", "Request refused."));
  next();
});
app.use("/api", express.json({ limit: "80mb" }));

const sendPage = (req, res) => {
  if (!fs.existsSync(PAGE)) return res.status(500).type("text").send(`Falta pagina\\QuoteFollowUp.html: rode o build.bat na pasta do programa.`);
  res.set("Cache-Control", "no-cache");
  res.sendFile(PAGE);
};
app.get(["/", "/index.html", "/QuoteFollowUp.html"], sendPage);

// avisos: o service worker precisa vir deste mesmo endereço, na raiz
app.get("/sw.js", (req, res) => {
  const file = path.join(ROOT, "pagina", "sw.js");
  if (!fs.existsSync(file)) return res.status(404).end();
  res.type("application/javascript").set({ "Cache-Control": "no-cache", "Service-Worker-Allowed": "/" });
  res.sendFile(file);
});

// celular: "Adicionar à tela inicial" abre o programa em tela cheia, com ícone próprio
app.get("/manifest.webmanifest", (req, res) => {
  res.type("application/manifest+json").set("Cache-Control", "no-cache");
  res.send(
    JSON.stringify({
      name: "Quote Follow Up",
      short_name: "Quotes",
      start_url: "/",
      scope: "/",
      display: "standalone",
      background_color: "#F3F6FA",
      theme_color: "#0D3D6E",
      icons: [
        { src: "/icon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
        { src: "/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
      ],
    }),
  );
});
for (const f of ["icon-192.png", "icon-512.png", "apple-touch-icon.png"])
  app.get("/" + f, (req, res) => {
    const file = path.join(ICONS, f);
    if (!fs.existsSync(file)) return res.status(404).end();
    res.set("Cache-Control", "public, max-age=86400");
    res.sendFile(file);
  });

app.get(
  "/api/ping",
  wrap(async (req, res) => {
    let db = "ok";
    try {
      await pool.query("SELECT 1");
    } catch {
      db = "down";
    }
    res.json({ app: "quote-follow-up", version: APP_VERSION, db });
  }),
);

// lista do login (não precisa estar logado): nomes e quem já tem senha
app.get(
  "/api/login-names",
  wrap(async (req, res) => {
    const { names, users } = await loginNames();
    res.json({ names, users: Object.fromEntries(users.map((u) => [u.name, { pwd: !!u.pass_hash }])) });
  }),
);

app.post(
  "/api/login",
  wrap(async (req, res) => {
    const name = str(req.body && req.body.name),
      pwd = String((req.body && req.body.password) || "");
    if (tooMany(req, name)) throw httpErr(429, "too-many", "Too many wrong passwords. Wait 15 minutes and try again.");
    const rows = name ? await pool.query("SELECT name, role, pass_salt, pass_hash FROM users WHERE name = ?", [name]) : [];
    const u = rows[0];
    const ok = !!(u && u.pass_hash && sameHex(hashPwd(pwd, u.pass_salt), u.pass_hash));
    if (!ok) {
      failed(req, name);
      if (u && !u.pass_hash) throw httpErr(409, "no-password", "This user has no password yet: create one.");
      throw httpErr(401, "wrong", "Wrong password.");
    }
    tries.delete(`nm|${clientIp(req)}|${name}`);
    await newSession(req, res, u.name);
    console.info(`[login] ${u.name}`);
    res.json({ name: u.name, role: roleOf(u.name, u.role) });
  }),
);

// primeiro acesso: cria a senha (só para quem ainda não tem)
app.post(
  "/api/password",
  wrap(async (req, res) => {
    const name = str(req.body && req.body.name),
      pwd = String((req.body && req.body.password) || "");
    if (tooMany(req, name)) throw httpErr(429, "too-many", "Too many attempts. Wait 15 minutes and try again.");
    if (pwd.length < 6) throw httpErr(400, "short", "The password needs at least 6 characters.");
    const { names } = await loginNames();
    if (name !== ADMIN_USER && !names.includes(name)) {
      failed(req, name);
      throw httpErr(400, "unknown", "Choose your user.");
    }
    const conn = await pool.getConnection();
    let role = "";
    try {
      await conn.beginTransaction();
      const rows = await conn.query("SELECT role, pass_hash FROM users WHERE name = ? FOR UPDATE", [name]);
      if (rows.length && rows[0].pass_hash) {
        await conn.rollback();
        failed(req, name);
        throw httpErr(409, "has-password", "This user already has a password. Forgot it? Ask the admin to reset it.");
      }
      const salt = crypto.randomBytes(16).toString("hex");
      role = name === ADMIN_USER ? "admin" : (rows[0] && rows[0].role) || null;
      await conn.query(
        "INSERT INTO users (name, role, pass_salt, pass_hash, changed_at) VALUES (?, ?, ?, ?, NOW()) ON DUPLICATE KEY UPDATE pass_salt = VALUES(pass_salt), pass_hash = VALUES(pass_hash), changed_at = NOW()",
        [name, role, salt, hashPwd(pwd, salt)],
      );
      await conn.query("UPDATE settings SET value = value + 1 WHERE name = 'data_version'");
      await conn.commit();
    } catch (e) {
      try {
        await conn.rollback();
      } catch {}
      throw e;
    } finally {
      conn.release();
    }
    await newSession(req, res, name);
    console.info(`[login] ${name} criou a senha`);
    res.json({ name, role: roleOf(name, role) });
  }),
);

app.post(
  "/api/logout",
  wrap(async (req, res) => {
    const token = cookies(req)[COOKIE];
    if (token) await pool.query("DELETE FROM sessions WHERE token_hash = ?", [sha256(token)]);
    setCookie(req, res, "", 0);
    res.json({ ok: true });
  }),
);

// quem está logado neste navegador (ninguém = name null)
app.get(
  "/api/me",
  wrap(async (req, res) => {
    const u = await sessionUser(req, res);
    res.json(u ? { name: u.name, role: u.role } : { name: null, role: "" });
  }),
);

// versões: a página olha a cada pouco e só busca os dados quando alguém mudou alguma coisa
app.get(
  "/api/version",
  auth,
  wrap(async (req, res) => {
    const v = await readVers();
    res.json({ ...v, me: { name: req.user.name, role: req.user.role } });
  }),
);

app.get(
  "/api/data",
  auth,
  needRole,
  wrap(async (req, res) => {
    if (req.query.part === "rep") {
      const rep = await getRep();
      res.set({ "Content-Type": "application/json; charset=utf-8", ETag: `"rep-${rep.ver}"` });
      if (/\bgzip\b/.test(req.get("Accept-Encoding") || "")) {
        res.set({ "Content-Encoding": "gzip", Vary: "Accept-Encoding" });
        return res.end(rep.gz);
      }
      return res.end(rep.json);
    }
    res.json(await readData(req.user));
  }),
);

/* Gravação: só as diferenças desde a última vez (o que esta página mudou), tudo numa transação.
   corpo: { meta: {QT: {...} | null}, excl: {add, del}, acts: [...], dailyGoal, dayGoals: {dia: valor | null},
            rep: {quotes, removed, sales, info} }  ← rep: importação dos relatórios (qualquer perfil com papel) */
app.post(
  "/api/sync",
  auth,
  needRole,
  wrap(async (req, res) => {
    const me = req.user,
      b = isObj(req.body) ? req.body : {};
    const mgr = MGR.has(me.role);
    // os três perfis (vendedor externo, interno e gerência) importam o Quote Report e o Sales Report; quem não tem papel nem chega aqui (needRole)
    let rep = await getRep();
    const skipped = new Set();
    let skippedGoals = false;
    const conn = await pool.getConnection();
    let out, goalBefore, goalAfter, newIds;
    const freshActs = [];
    const t0 = Date.now();
    try {
      await conn.beginTransaction();
      const prev = await lockVers(conn); // uma gravação por vez: a próxima espera esta terminar
      let repN = null;
      if (b.rep) {
        // avisos: QTs que entram pela primeira vez nesta importação
        const at = stamp(b.rep.info && b.rep.info.quoteAt);
        const cand = at ? arr(b.rep.quotes).filter((q) => isObj(q) && q.fs === at).map((q) => qid(q.id)) : [];
        const had = new Set();
        for (const part of chunks(cand, 1000)) (await conn.query("SELECT quote_id FROM quotes WHERE quote_id IN (?)", [part])).forEach((r) => had.add(r.quote_id));
        newIds = new Set(cand.filter((id) => !had.has(id)));
        repN = await applyRep(conn, b.rep, me.name);
      }
      if (repN && repN.quotes) {
        // cotações novas da importação: o vendedor delas também precisa ser conhecido na conferência abaixo
        rep = { ...rep, byId: new Map(rep.byId) };
        arr(b.rep.quotes).forEach((q) => isObj(q) && rep.byId.set(qid(q.id), { sp: str(q.sp), qb: str(q.qb), c: str(q.c) }));
      }
      // gerência só visualiza: follow-ups, notas e formulários de cotação vêm dos vendedores (o registro da importação continua)
      const okId = mgr ? () => false : scopeOf(me, rep);
      const nF = await applyFollowups(conn, me, b, okId, rep, skipped);
      const nA = await applyActs(conn, me, b, okId, skipped, freshActs);
      let nG = 0;
      if (b.dailyGoal !== undefined || isObj(b.dayGoals)) {
        if (mgr) {
          goalBefore = await goalVec(conn);
          nG = await applyGoals(conn, b);
          goalAfter = await goalVec(conn);
        } else skippedGoals = true;
      }
      const dataCh = nF + nA + nG > 0;
      const next = { data: prev.data + (dataCh ? 1 : 0), rep: prev.rep + (repN ? 1 : 0) };
      if (dataCh) await setSetting(conn, "data_version", next.data);
      if (repN) await setSetting(conn, "rep_version", next.rep);
      await conn.commit();
      out = { ok: true, data: next.data, prevData: prev.data, rep: next.rep, prevRep: prev.rep, skipped: [...skipped], skippedGoals };
      const parts = [nF && `${nF} follow-up(s)`, nA && `${nA} histórico`, nG && `${nG} meta(s)`, repN && `relatórios: ${repN.quotes} cotações, ${repN.removed} removidas${repN.sales >= 0 ? `, ${repN.sales} vendas` : ""}`, skipped.size && `${skipped.size} fora da carteira (ignorado)`].filter(Boolean);
      if (parts.length) console.info(`[sync] ${me.name}: ${parts.join(" · ")} (${Date.now() - t0} ms)`);
    } catch (e) {
      try {
        await conn.rollback();
      } catch {}
      throw e;
    } finally {
      conn.release();
    }
    if (out.rep !== out.prevRep) getRep().catch((e) => console.error("[rep] erro ao montar", e)); // já deixa pronto para os outros
    res.json(out);
    // avisos no celular / PC de quem foi afetado (depois da resposta: não atrasa a gravação)
    if (push.ready && ((newIds && newIds.size) || goalAfter || freshActs.length))
      push.afterSync({ me, rep, newIds, acts: freshActs, quotes: b.rep ? arr(b.rep.quotes) : [], goalBefore, goalAfter }).catch((e) => console.error("[avisos]", e.message || e));
  }),
);

/* Admin: papel (tela) de cada pessoa e reset de senha */
app.post(
  "/api/users/role",
  auth,
  needAdmin,
  wrap(async (req, res) => {
    const name = cut(req.body && req.body.name, 100),
      r = str(req.body && req.body.role);
    if (!name || name === ADMIN_USER) throw httpErr(400, "bad-user", "Choose a user.");
    const role = ["manager", "outside", "inside"].includes(r) ? r : null;
    await pool.query("INSERT INTO users (name, role, changed_at) VALUES (?, ?, NOW()) ON DUPLICATE KEY UPDATE role = VALUES(role), changed_at = NOW()", [name, role]);
    await pool.query("UPDATE settings SET value = value + 1 WHERE name = 'data_version'");
    console.info(`[admin] ${name} → ${role || "sem acesso"}`);
    res.json({ ok: true, users: sanitizeUsers(await pool.query("SELECT name, role, pass_hash FROM users"), req.user) });
  }),
);
app.post(
  "/api/users/reset",
  auth,
  needAdmin,
  wrap(async (req, res) => {
    const name = cut(req.body && req.body.name, 100);
    if (!name) throw httpErr(400, "bad-user", "Choose a user.");
    await pool.query("UPDATE users SET pass_salt = '', pass_hash = '', changed_at = NOW() WHERE name = ?", [name]);
    await pool.query("DELETE FROM sessions WHERE name = ?", [name]); // sai de todos os aparelhos
    await pool.query("UPDATE settings SET value = value + 1 WHERE name = 'data_version'");
    console.info(`[admin] senha de ${name} zerada`);
    res.json({ ok: true, users: sanitizeUsers(await pool.query("SELECT name, role, pass_hash FROM users"), req.user) });
  }),
);

/* avisos no celular / PC: estado, ativar / desativar um aparelho, teste e o horário do lembrete da manhã */
app.get(
  "/api/push",
  auth,
  needRole,
  wrap(async (req, res) => {
    res.json({ on: push.ready, key: push.publicKey, hour: push.hour, devices: push.ready ? await push.count(req.user.name) : 0 });
  }),
);
app.post(
  "/api/push/subscribe",
  auth,
  needRole,
  wrap(async (req, res) => {
    const origin = isHttps(req) ? "https://" + str(req.get("host")).toLowerCase() : "";
    await push.subscribe(req.user.name, req.body && req.body.sub, req.headers["user-agent"], origin);
    res.json({ ok: true, devices: await push.count(req.user.name) });
  }),
);
app.post(
  "/api/push/unsubscribe",
  auth,
  wrap(async (req, res) => {
    if (push.ready) await push.unsubscribe(str(req.body && req.body.endpoint), req.user.name);
    res.json({ ok: true, devices: push.ready ? await push.count(req.user.name) : 0 });
  }),
);
app.post(
  "/api/push/test",
  auth,
  needRole,
  wrap(async (req, res) => {
    if (!push.ready) throw httpErr(503, "push-off", "Push is off on the server.");
    const n = process.env.QFU_PUSH_TEST === "1" && req.body && req.body.morning ? await push.morning(true, req.user.name) : await push.sendTo([req.user.name], { title: "Avisos ligados", body: "Este aparelho vai receber os avisos do Quote Follow Up.", tag: "test" });
    res.json({ ok: true, sent: n });
  }),
);
app.post(
  "/api/push/hour",
  auth,
  needMgr,
  wrap(async (req, res) => {
    if (!push.ready) throw httpErr(503, "push-off", "Push is off on the server.");
    await push.setHour(str(req.body && req.body.hour));
    console.info(`[avisos] lembrete da manhã: ${push.hour} (${req.user.name})`);
    res.json({ ok: true, hour: push.hour });
  }),
);

/* notificações: "lido até" de cada pessoa (settings "notif:<nome>", até 255 caracteres: { at, g, r }) */
const notifKey = (name) => "notif:" + String(name).slice(0, 54);
const notifKindKey = (name) => "notifk:" + String(name).slice(0, 53); // o que já foi visto de cada tipo (clicado no sino)
const NOTIF_KINDS = new Set(["new", "status", "order", "amount", "note", "dcs", "sales"]);
app.get(
  "/api/notif",
  auth,
  needRole,
  wrap(async (req, res) => {
    const r = await pool.query("SELECT name, value FROM settings WHERE name IN (?, ?)", [notifKey(req.user.name), notifKindKey(req.user.name)]);
    const get = (n) => {
      try {
        const x = r.find((y) => y.name === n);
        return x ? JSON.parse(x.value) : null;
      } catch {
        return null;
      }
    };
    const seen = get(notifKey(req.user.name));
    if (seen) seen.k = get(notifKindKey(req.user.name)) || {};
    res.json({ seen });
  }),
);
app.post(
  "/api/notif",
  auth,
  needRole,
  wrap(async (req, res) => {
    const s = isObj(req.body && req.body.seen) ? req.body.seen : null;
    if (!s || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(s.at || ""))) throw httpErr(400, "bad-request", "Invalid data.");
    const TS = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
    const v = JSON.stringify({ at: s.at, g: Array.isArray(s.g) ? s.g.slice(0, 4).map((x) => Math.round(+x || 0)) : [], r: /^\d{4}-\d{2}-\d{2}$/.test(String(s.r || "")) ? s.r : "", mine: (Array.isArray(s.mine) ? s.mine : []).map(String).filter((t) => TS.test(t)).slice(-4), u: TS.test(String(s.u || "")) ? s.u : s.at });
    const k = JSON.stringify(Object.fromEntries(Object.entries(isObj(s.k) ? s.k : {}).filter(([n, t]) => NOTIF_KINDS.has(n) && TS.test(String(t)) && String(t) > s.at)));
    const up = "INSERT INTO settings (name, value) VALUES (?, ?) ON DUPLICATE KEY UPDATE value = VALUES(value)";
    await pool.query(up, [notifKey(req.user.name), v.slice(0, 255)]);
    await pool.query(up, [notifKindKey(req.user.name), k.slice(0, 255)]);
    res.json({ ok: true });
  }),
);

/* Admin: planilhas da versão anterior (Quote Follow Up data.xlsx e/ou reports.xlsx), direto pela página — o mesmo
   que o importar-banco.bat faz, numa transação só. Corpo = os bytes dos dois arquivos, um depois do outro;
   ?d=<bytes do data>&r=<bytes do reports>&mode=keep|replace. Histórico: só entra o que falta (quem fez é mantido). */
app.post(
  "/api/legacy-import",
  auth,
  needAdmin,
  express.raw({ type: "application/octet-stream", limit: "150mb" }),
  wrap(async (req, res) => {
    const buf = Buffer.isBuffer(req.body) ? req.body : Buffer.alloc(0);
    const dLen = Math.max(0, parseInt(req.query.d, 10) || 0),
      rLen = Math.max(0, parseInt(req.query.r, 10) || 0);
    const mode = req.query.mode === "replace" ? "replace" : "keep";
    if (!dLen && !rLen) throw httpErr(400, "no-file", "Choose the data file and/or the reports file.");
    if (dLen + rLen !== buf.length) throw httpErr(400, "bad-upload", "The upload was cut off: try again.");
    const read = (b, bad) => {
      if (!b.length) return null;
      let wb = null;
      try {
        wb = XLSX.read(b, { type: "buffer" });
      } catch {}
      if (!wb || !core.isDataWorkbook(wb)) throw httpErr(400, "bad-file", bad);
      return wb;
    };
    const dataWb = read(buf.subarray(0, dLen), "The data file is not a Quote Follow Up data file of the previous version."),
      repWb = read(buf.subarray(dLen), "The reports file is not a Quote Follow Up reports file of the previous version.");
    const t0 = Date.now();
    const conn = await pool.getConnection();
    let out;
    try {
      await conn.beginTransaction();
      const prev = await lockVers(conn); // uma gravação por vez
      const cur = Object.fromEntries((await conn.query("SELECT name, value FROM settings")).map((r) => [r.name, r.value]));
      const { statements, counts, warnings } = buildLegacyImport(core, dataWb, repWb, { mode, cur });
      const done = {};
      for (const st of statements) {
        const r = await conn.query(st.sql);
        done[st.what] = (done[st.what] || 0) + Number((r && r.affectedRows) || 0);
      }
      // histórico novo = linhas que entraram de fato (INSERT IGNORE não conta as repetidas)
      const next = { data: prev.data + 1, rep: prev.rep + (counts.quotes || counts.sales ? 1 : 0) };
      await setSetting(conn, "data_version", next.data);
      if (next.rep !== prev.rep) await setSetting(conn, "rep_version", next.rep);
      await conn.commit();
      out = { ok: true, mode, counts, newActs: done.acts || 0, warnings, data: next.data, rep: next.rep };
      console.info(`[importação] ${req.user.name} (${mode}): ${counts.quotes} cotações · ${counts.sales} vendas · ${counts.followups} follow-ups · histórico ${done.acts || 0} novo(s) de ${counts.acts} · ${counts.users} usuários (${Date.now() - t0} ms)`);
    } catch (e) {
      try {
        await conn.rollback();
      } catch {}
      throw e;
    } finally {
      conn.release();
    }
    if (out.rep) getRep().catch((e) => console.error("[rep] erro ao montar", e));
    res.json(out);
  }),
);

app.use("/api", (req, res) => res.status(404).json({ error: "not-found", message: "Not found." }));
app.use((req, res) => res.status(404).type("text").send("Not found"));
// erros: a página recebe uma mensagem curta; o detalhe fica só no console do servidor
app.use((err, req, res, next) => {
  const status = err.status || (err.type === "entity.too.large" ? 413 : err.type === "entity.parse.failed" ? 400 : 500);
  const known = !!err.code && !!err.status; // erro previsto (httpErr): a mensagem vai para a página
  if (status >= 500 && !known) console.error(`[erro] ${req.method} ${req.path}`, err);
  else if (status >= 500) console.warn(`[aviso] ${req.method} ${req.path}: ${err.message}`);
  res.status(status).json({ error: err.code || (status >= 500 ? "server" : "bad-request"), message: status >= 500 && !known ? "Server error: try again in a moment." : err.message });
});

/* ---------- partida ---------- */
async function prepareDb() {
  // tabela de sessões e contadores de versão (o schema.sql novo já cria; aqui é para bancos criados antes)
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS sessions (
      token_hash  CHAR(64)     NOT NULL,
      name        VARCHAR(100) NOT NULL,
      created_at  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      last_seen   DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      expires_at  DATETIME     NOT NULL,
      user_agent  VARCHAR(200) NOT NULL DEFAULT '',
      PRIMARY KEY (token_hash),
      KEY ix_sessions_name (name),
      KEY ix_sessions_expires (expires_at)
    ) ENGINE=InnoDB`);
  } catch (e) {
    const has = await pool.query("SHOW TABLES LIKE 'sessions'");
    if (!has.length) throw new Error("a tabela sessions não existe e o usuário do banco não pode criá-la: rode o schema.sql novo (mariadb -u root -p < banco\\schema.sql)");
  }
  await pool.query("INSERT IGNORE INTO settings (name, value) VALUES ('data_version', '1'), ('rep_version', '1'), ('quote_fill_at', '')");
  // formulário DCS / Value Added: tenta liberar o tipo "dcs" no histórico (precisa de permissão ALTER; senão, avisa)
  if (!(await dcsAllowed())) {
    try {
      await pool.query("ALTER TABLE activities DROP CONSTRAINT IF EXISTS ck_activities_kind");
      await pool.query("ALTER TABLE activities ADD CONSTRAINT ck_activities_kind CHECK (kind IN ('note', 'fup', 'win', 'done', 'resched', 'req', 'plan', 'status', 'amount', 'order', 'close', 'reopen', 'why', 'dcs'))");
      dcsOk = true;
      console.info("[banco] histórico liberado para o formulário DCS / Value Added");
    } catch {
      console.warn("\n[AVISO] O banco ainda não aceita o formulário DCS / Value Added. Rode o schema.sql novo como root:\n        mariadb -u root -p < banco\\schema.sql\n");
    }
  }
  await pool.query("DELETE FROM sessions WHERE expires_at < NOW()");
  try {
    await push.init();
  } catch (e) {
    console.warn("[avisos] desligados:", e.message || e);
  }
}

try {
  await prepareDb();
  await getRep();
} catch (e) {
  console.error("\nNão foi possível falar com o banco de dados:", e.message || e);
  console.error(`Confira o "db" em ${path.basename(CFG_FILE)} (usuário, senha, banco) e se o MariaDB está rodando.\n`);
  process.exit(1);
}
setInterval(() => pool.query("DELETE FROM sessions WHERE expires_at < NOW()").catch(() => {}), 3600000).unref();

const server = app.listen(PORT, HOST, () => {
  console.info(`\nQuote Follow Up · servidor no ar em http://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);
  console.info(`Banco: ${(cfg.db && cfg.db.database) || "quote_follow_up"} · página: ${path.basename(PAGE)} · versão do programa ${APP_VERSION}`);
  console.info("Para parar: Ctrl+C\n");
});
server.requestTimeout = 5 * 60 * 1000; // importação grande em conexão lenta
const stop = () => {
  console.info("Parando o servidor…");
  server.close(() => pool.end().finally(() => process.exit(0)));
  setTimeout(() => process.exit(0), 5000).unref();
};
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
