/* =====================================================================================
   QUOTE FOLLOW UP — regras de negócio (sem React)
   Leitura do Quote Report e do Sales Report, montagem da carteira por cliente,
   agenda de follow-up (FUP) e cálculo das cotações necessárias para bater a meta.
   Os parsers vieram do BRAZIL OEM PLAN (pulz) e foram enxugados para este programa.
   ===================================================================================== */
import * as XLSX from "xlsx";

/* ===== utilitários de texto, número e data ===== */
export const norm = (s) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
export const clean = (v) => (v == null ? "" : String(v).trim());
export const normName = (s) =>
  String(s || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\b(LTDA|EIRELI|EPP|ME|S\/?A|SA|DO BRASIL|INDUSTRIA E COMERCIO|IND E COM|COMERCIO|INDUSTRIA)\b/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
export const contaCode = (s) => {
  const m = String(s || "")
    .toUpperCase()
    .match(/C\d{3,}/);
  return m ? m[0] : "";
};
export const money = (n) =>
  "R$ " + (Number(n) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
// valor curto para os ladrilhos: R$ 1,2 mi / R$ 350,4 k
export const moneyShort = (v) => {
  const a = Math.abs(v || 0);
  if (a >= 1e6) return "R$ " + (v / 1e6).toFixed(2).replace(".", ",") + " mi";
  if (a >= 1e4) return "R$ " + (v / 1e3).toFixed(1).replace(".", ",") + " k";
  return money(v);
};
export const fmtDate = (iso) => {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
};
export const fmtDay = (iso) => {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}` : "";
};
export const pad2 = (n) => String(n).padStart(2, "0");
// data de HOJE no fuso do computador (toISOString daria o dia em UTC, que vira o dia seguinte à noite)
export const todayISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
};
/* Datas como texto ISO (AAAA-MM-DD). As contas são feitas em UTC para o horário de verão não
   pular nem repetir dia. */
const isoToUTC = (iso) => {
  const [y, m, d] = iso.split("-").map(Number);
  return Date.UTC(y, m - 1, d);
};
const utcToISO = (t) => new Date(t).toISOString().slice(0, 10);
export const addDays = (iso, n) => utcToISO(isoToUTC(iso) + n * 86400000);
export const dowOf = (iso) => new Date(isoToUTC(iso)).getUTCDay(); // 0 = domingo
export const daysBetween = (a, b) => Math.round((isoToUTC(b) - isoToUTC(a)) / 86400000);
export const addMonths = (iso, k) => {
  const [y, m, d] = iso.split("-").map(Number);
  const tot = y * 12 + (m - 1) + k;
  const ny = Math.floor(tot / 12),
    nm = (tot % 12) + 1;
  const last = new Date(Date.UTC(ny, nm, 0)).getUTCDate(); // 31/01 + 1 mês = 28 ou 29/02
  return `${ny}-${pad2(nm)}-${pad2(Math.min(d, last))}`;
};
export const lastDayOfMonth = (ym) => {
  const [y, m] = ym.split("-").map(Number);
  return `${ym}-${pad2(new Date(Date.UTC(y, m, 0)).getUTCDate())}`;
};
const isWeekend = (iso) => {
  const w = dowOf(iso);
  return w === 0 || w === 6;
};
/* ===== Feriados (nacionais, estadual de SP e municipais de Artur Nogueira) =====
   block: true = não é dia útil: nenhum follow-up cai nele (vai para o próximo dia útil) e não tem meta.
   Os pontos facultativos (Carnaval e Corpus Christi) aparecem no calendário, mas continuam dias úteis;
   para bloquear também, troque "block: false" por "block: true" neles. */
export const HOLIDAYS_FIXED = {
  "01-01": { name: "New Year (Ano Novo)", kind: "national", block: true },
  "04-10": { name: "Artur Nogueira anniversary", kind: "municipal", block: true },
  "04-21": { name: "Tiradentes", kind: "national", block: true },
  "05-01": { name: "Labour Day (Dia do Trabalhador)", kind: "national", block: true },
  "07-09": { name: "Constitutionalist Revolution (SP)", kind: "state", block: true },
  "09-07": { name: "Independence Day", kind: "national", block: true },
  "09-15": { name: "N. Sra. das Dores (Artur Nogueira)", kind: "municipal", block: true },
  "10-12": { name: "N. Sra. Aparecida", kind: "national", block: true },
  "11-02": { name: "All Souls' Day (Finados)", kind: "national", block: true },
  "11-15": { name: "Republic Day", kind: "national", block: true },
  "11-20": { name: "Black Consciousness Day", kind: "national", block: true },
  "12-25": { name: "Christmas (Natal)", kind: "national", block: true },
};
// datas que mudam todo ano: dias contados a partir da Páscoa
export const HOLIDAYS_EASTER = [
  { off: -48, name: "Carnival Monday", kind: "optional", block: false },
  { off: -47, name: "Carnival Tuesday", kind: "optional", block: false },
  { off: -2, name: "Good Friday (Sexta-feira Santa)", kind: "national", block: true },
  { off: 60, name: "Corpus Christi", kind: "optional", block: false },
];
// domingo de Páscoa (algoritmo de Meeus/Jones/Butcher)
const easterOf = (y) => {
  const a = y % 19,
    b = Math.floor(y / 100),
    c = y % 100,
    d = Math.floor(b / 4),
    e = b % 4,
    f = Math.floor((b + 8) / 25),
    g = Math.floor((b - f + 1) / 3),
    h = (19 * a + b - d - g + 15) % 30,
    i = Math.floor(c / 4),
    k = c % 4,
    l = (32 + 2 * e + 2 * i - h - k) % 7,
    m = Math.floor((a + 11 * h + 22 * l) / 451),
    mo = Math.floor((h + l - 7 * m + 114) / 31),
    da = ((h + l - 7 * m + 114) % 31) + 1;
  return `${y}-${pad2(mo)}-${pad2(da)}`;
};
const holYears = {};
const holidaysOfYear = (y) => {
  if (holYears[y]) return holYears[y];
  const m = {};
  Object.entries(HOLIDAYS_FIXED).forEach(([md, h]) => (m[`${y}-${md}`] = h));
  const e = easterOf(y);
  HOLIDAYS_EASTER.forEach((h) => (m[addDays(e, h.off)] = m[addDays(e, h.off)] || h));
  return (holYears[y] = m);
};
// feriado do dia ({ name, kind, block }) ou null
export const holidayOf = (iso) => (/^\d{4}-\d{2}-\d{2}$/.test(iso || "") ? holidaysOfYear(+iso.slice(0, 4))[iso] || null : null);
// dia sem expediente: sábado, domingo ou feriado
export const isOffDay = (iso) => isWeekend(iso) || !!(holidayOf(iso) || {}).block;
// sábado, domingo e feriado vão para o próximo dia útil
export const toBusinessDay = (iso) => {
  let d = iso;
  while (isOffDay(d)) d = addDays(d, 1);
  return d;
};
// dia útil anterior (pula fim de semana e feriado)
export const prevBusinessDay = (iso) => {
  let d = addDays(iso, -1);
  while (isOffDay(d)) d = addDays(d, -1);
  return d;
};
export const MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const monthLabel = (ym) => {
  const [y, m] = ym.split("-").map(Number);
  return `${MONTHS_EN[m - 1]} ${y}`;
};

function toISO(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) return `${br[3]}-${br[2]}-${br[1]}`;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
}
// datas em texto: "us" = mês/dia/ano (padrão dos relatórios do sistema); senão dia/mês/ano
function parseDate(v, us) {
  if (v == null || v === "") return "";
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === "number") return toISO(v);
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let a = +m[1],
      b = +m[2],
      y = m[3];
    if (y.length === 2) y = "20" + y;
    const mm = us ? a : b,
      dd = us ? b : a;
    if (mm < 1 || mm > 12) return "";
    return `${y}-${pad2(mm)}-${pad2(dd)}`;
  }
  return toISO(v);
}
// vota se as datas em texto estão em mês/dia (EUA) ou dia/mês (Brasil). Empate = EUA, como no pulz.
function guessUS(aoa, hr, col) {
  let us = 0,
    br = 0;
  for (let i = hr + 1; i < Math.min(aoa.length, hr + 800); i++) {
    const v = aoa[i] && aoa[i][col];
    if (typeof v !== "string") continue;
    const m = v.match(/^(\d{1,2})\/(\d{1,2})\//);
    if (!m) continue;
    if (+m[1] > 12) br++;
    else if (+m[2] > 12) us++;
  }
  return !(br > us);
}
export const toNum = (v) => {
  if (typeof v === "number") return v;
  if (v == null) return 0;
  const n = parseFloat(
    String(v)
      .replace(/[^0-9,.-]/g, "")
      .replace(/\.(?=\d{3}(\D|$))/g, "")
      .replace(",", "."),
  );
  return isNaN(n) ? 0 : n;
};
const r2 = (v) => Math.round(v * 100) / 100;
const custName = (raw) => clean(String(raw || "").split(/[\n(]/)[0]);

const sheetAoA = (wb, name) => XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });

/* ===== Linha de produto (PT / CP / MH) =====
   Mesma regra do BRAZIL OEM PLAN (pulz): o Product Group diz a linha; "Purchase Components" (90) é PT
   ou MH conforme o subgrupo que vem no Product Code, ex. "Transmission Chain (89154)". */
export const ALL_LINES = ["PT", "CP", "MH"];
const PT_PURCHASE_COMPONENT_SUBGROUPS = new Set([
  "057",
  "05709",
  "073",
  "197",
  "89154",
  "89156",
  "89192",
  "90082",
  "90412",
  "90418",
  "90421",
]);
const PG_LINE = {
  sprockets: "PT",
  gears: "PT",
  sheaves: "PT",
  bushings: "PT",
  "timing pulleys": "PT",
  couplings: "PT",
  "specialty items": "PT",
  "material handling": "MH",
  "conveyor pulleys": "CP",
  idlers: "CP",
  "service sales pt": "PT",
  "service sales cp": "CP",
  "service sales mh": "MH",
};


export function lineOf(pg, subItem) {
  /* serviços com a linha no código: "No Booking ,Services (99 CP)" / "(99 PT)" / "(99 H)" no Quote Report,
     "Service Sales MH (99 MH)" no Sales Report. "(99)" sozinho não tem linha e não conta. */
  const svc = /\(\s*99\s*(CP|PT|MH|H)\s*\)/i.exec(String(pg || ""));
  if (svc) return { CP: "CP", PT: "PT", MH: "MH", H: "MH" }[svc[1].toUpperCase()];
  let s = String(pg || "")
    .toLowerCase()
    .replace(/\s*-\s*mto/g, "")
    .replace(/\([^)]*\)/g, "")
    .trim();
  const sub = String(subItem || "").replace(/[^0-9]/g, "");
  if (s === "90" || s.startsWith("90 ") || s.startsWith("90-"))
    return PT_PURCHASE_COMPONENT_SUBGROUPS.has(sub) ? "PT" : "MH";
  if (s.startsWith("purchase components"))
    return PT_PURCHASE_COMPONENT_SUBGROUPS.has(sub) ? "PT" : "MH";
  if (s.includes("service sales cp")) return "CP";
  if (s.includes("service sales pt")) return "PT";
  if (s.includes("service sales mh")) return "MH";
  for (const key in PG_LINE) if (s.startsWith(key)) return PG_LINE[key];
  return "";
}
/* PRODUCT GROUP QUE NÃO CONTA: item com Product Group preenchido que não é de nenhuma linha
   (PT / CP / MH) — No Booking / Services (99), Customer Sample (91), Shop Supplies (95/959),
   Raw materials (89), Machinery maintenance (959)… Essas linhas são ignoradas na leitura do
   Quote Report e do Sales Report: não somam valor, e a cotação que só tem esses itens não entra. */
// nome do Product Group para mostrar: sem o "- MTO" e com o código entre parênteses (ex.: "Sheaves (21)")
export const pgLabel = (pg) =>
  clean(pg)
    .replace(/\s*-\s*MTO\b/i, "")
    .replace(/\(\s*(\d+)\s*M\s*\)/i, "($1)")
    .trim();
export const skipPG = (pg, sub) => clean(pg) !== "" && !lineOf(pg, sub);

/* ===== Quote Report =====
   Identificado pela coluna "Quote ID". Uma entrada por cotação, somando as linhas (itens) dela.
   Diferente do pulz, aqui NÃO filtramos OEM na leitura: a classificação fica guardada em cada cotação
   e o filtro "Classification" da tela decide o recorte. */
export function isQuoteWorkbook(wb) {
  for (const name of wb.SheetNames) {
    const aoa = sheetAoA(wb, name);
    for (let r = 0; r < Math.min(aoa.length, 30); r++)
      if ((aoa[r] || []).some((c) => /^(quote id|quote no|quote number)$/.test(norm(c)))) return true;
  }
  return false;
}
export function parseQuoteWorkbook(wb) {
  let best = null;
  for (const name of wb.SheetNames) {
    const aoa = sheetAoA(wb, name);
    for (let hr = 0; hr < Math.min(aoa.length, 30); hr++) {
      const row = (aoa[hr] || []).map((c) => norm(c));
      const eq = (k) => row.findIndex((c) => c === k);
      const has = (kws) => row.findIndex((c) => c && kws.some((k) => c.includes(k)));
      const pick = (exact, kws) => (eq(exact) >= 0 ? eq(exact) : has(kws));
      const qi = has(["quote id", "quote no", "quote number", "cotação", "cotacao"]);
      if (qi < 0) continue;
      best = {
        name,
        aoa,
        hr,
        qi,
        cli: has(["classification", "classifica"]),
        ci: has(["customer", "cliente"]),
        di: has(["create date", "quote date", "data"]),
        si: pick("status", ["status"]),
        sidi: has(["sales id", "sales order", "pedido"]),
        sai: has(["sales amt", "sales amount"]),
        qbi: pick("quoted by", ["quoted by", "cotado por"]),
        spi: pick("sales person", ["sales person", "salesperson"]),
        ai: pick("amount", ["amount", "valor"]),
        pgi: has(["product group", "item group"]),
        pci: has(["product code"]),
      };
      break;
    }
    if (best) break;
  }
  if (!best || best.di < 0 || best.ci < 0) return null;
  const us = guessUS(best.aoa, best.hr, best.di);
  const byId = {},
    seen = new Set();
  let rows = 0,
    skipped = 0,
    skippedAmt = 0,
    minD = "",
    maxD = "";
  for (let i = best.hr + 1; i < best.aoa.length; i++) {
    const row = best.aoa[i] || [];
    const id = clean(row[best.qi]).toUpperCase();
    if (!id) continue;
    const d = parseDate(row[best.di], us);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
    rows++;
    if (!minD || d < minD) minD = d;
    if (!maxD || d > maxD) maxD = d;
    seen.add(id);
    if (best.pgi >= 0 && skipPG(row[best.pgi], best.pci >= 0 ? row[best.pci] : "")) {
      skipped++;
      skippedAmt += best.ai >= 0 ? toNum(row[best.ai]) : 0;
      continue;
    }
    const custRaw = String(row[best.ci] || "");
    const q =
      byId[id] ||
      (byId[id] = {
        id,
        d, // data da cotação (Create Date)
        c: custName(custRaw),
        conta: contaCode(custRaw),
        cls: "", // Classification (OEM, Distributor...)
        st: "", // status (normalizado)
        qb: "", // Quoted By — quem montou
        sp: "", // Sales Person — vendedor da conta
        a: 0, // valor cotado (Amount)
        sa: 0, // Sales Amt (valor que a própria cotação diz ter virado pedido)
        so: [], // pedidos (Sales Id) ligados à cotação
        al: {}, // valor cotado por linha de produto { PT, CP, MH }
        lr: best.pgi >= 0, // a linha dos itens foi lida (o relatório tem Product Group)
        pg: {}, // valor cotado por Product Group (ex.: "Sprockets (11)": 1234,5)
      });
    if (d < q.d) q.d = d;
    const st = best.si >= 0 ? norm(row[best.si]) : "";
    if (st) q.st = st;
    if (!q.cls && best.cli >= 0) q.cls = clean(row[best.cli]);
    if (!q.qb && best.qbi >= 0) q.qb = clean(row[best.qbi]);
    if (!q.sp && best.spi >= 0) q.sp = clean(row[best.spi]);
    const amt = best.ai >= 0 ? toNum(row[best.ai]) : 0;
    q.a += amt;
    const ln = best.pgi >= 0 ? lineOf(row[best.pgi], best.pci >= 0 ? row[best.pci] : "") : "";
    if (ln) q.al[ln] = (q.al[ln] || 0) + amt;
    if (best.pgi >= 0) {
      const pgName = pgLabel(row[best.pgi]);
      if (pgName) q.pg[pgName] = (q.pg[pgName] || 0) + amt;
    }
    if (best.sai >= 0) q.sa += toNum(row[best.sai]);
    const so = best.sidi >= 0 ? clean(row[best.sidi]).toUpperCase() : "";
    if (so && !q.so.includes(so)) q.so.push(so);
  }
  // valores sem arredondar (como o Excel soma); o arredondamento fica só na hora de mostrar
  const quotes = Object.values(byId).map((q) => ({ ...q, nn: normName(q.c) }));
  // cotações que só tinham itens que não contam: saem também do que já estava guardado
  const dropped = [...seen].filter((id) => !byId[id]);
  return { quotes, rows, skipped, skippedAmt: r2(skippedAmt), dropped, minD, maxD, sheet: best.name, hasClass: best.cli >= 0, hasSO: best.sidi >= 0, hasLine: best.pgi >= 0 };
}

/* ===== Sales Report (Booked) =====
   Colunas reconhecidas por palavra-chave, como no pulz. Só entram as linhas BOOKED (quando existe a
   coluna de tipo/status). As linhas são somadas por pedido + dia + classificação + linha de produto
   + cliente: é isso que fica guardado (e vai para o arquivo de dados). Os índices usados nas telas
   saem dessas linhas em buildSalesIndex. */
export function parseSalesWorkbook(wb) {
  const kwDate = ["create date", "order date", "invoice date", "booking date", "date", "data", "emiss", "fatur"];
  const kwStatus = ["type", "status", "tipo", "situa"];
  const kwCust = ["customer", "cliente", "conta", "account", "razao", "razão", "sold to"];
  const kwAmount = ["amount", "valor", "value", "net", "líquido", "liquido", "total"];
  const kwClass = ["classification", "classifica", "cust class"];
  const kwSalesId = ["sales id", "sales order", "order id", "pedido"];
  const kwPG = ["product group", "item group", "grupo de produto"];
  const kwSubItem = ["sub item", "subitem", "sub-item", "subgrupo", "sub group", "product code"];
  let best = null;
  for (const name of wb.SheetNames) {
    const aoa = sheetAoA(wb, name);
    for (let hr = 0; hr < Math.min(aoa.length, 30); hr++) {
      const row = (aoa[hr] || []).map((c) => norm(c));
      const find = (kws) => row.findIndex((c) => c && kws.some((k) => c.includes(k)));
      const di = find(kwDate),
        ci = find(kwCust);
      if (di >= 0 && ci >= 0) {
        const exact = (ks) => row.findIndex((c) => ks.includes(c));
        best = {
          name, aoa, hr, di, ci, si: find(kwStatus), ai: find(kwAmount), cli: find(kwClass), sidi: find(kwSalesId), pgi: find(kwPG), sii: find(kwSubItem),
          // vendedor: código ("Sales Responsible" = SP0178) e nome ("Name" = Julia Pulz Campo)
          spi: exact(["sales responsible", "salesperson", "sales person", "sales rep", "vendedor"]),
          nmi: exact(["name", "salesperson name", "nome"]),
        };
        break;
      }
    }
    if (best) break;
  }
  if (!best) return null;
  const us = guessUS(best.aoa, best.hr, best.di);
  const agg = {};
  let rows = 0,
    booked = 0,
    minD = "",
    maxD = "";
  for (let i = best.hr + 1; i < best.aoa.length; i++) {
    const row = best.aoa[i] || [];
    if (!row.some((c) => c != null && c !== "")) continue;
    const d = parseDate(row[best.di], us);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
    rows++;
    if (best.si >= 0) {
      const st = norm(row[best.si]);
      if (st && !st.includes("book")) continue; // Invoiced e outros tipos ficam de fora
    }
    if (!minD || d < minD) minD = d;
    if (!maxD || d > maxD) maxD = d;
    if (best.pgi >= 0 && skipPG(row[best.pgi], best.sii >= 0 ? row[best.sii] : "")) continue; // Product Group que não conta
    booked++;
    const amt = best.ai >= 0 ? toNum(row[best.ai]) : 0;
    const cls = best.cli >= 0 ? clean(row[best.cli]) : "";
    const ln = best.pgi >= 0 ? lineOf(row[best.pgi], best.sii >= 0 ? row[best.sii] : "") : "";
    const so = best.sidi >= 0 ? clean(row[best.sidi]).toUpperCase() : "";
    const c = custName(row[best.ci]);
    const code = best.spi >= 0 ? clean(row[best.spi]) : "",
      nm = best.nmi >= 0 ? clean(row[best.nmi]) : "";
    // vendedor no mesmo formato do Quote Report: "SP0178 - Julia Pulz Campo"
    const sp = code && nm ? (/^sp\d+$/i.test(code) ? `${code.toUpperCase()} - ${nm}` : nm) : nm || code;
    const k = [so, d, cls, ln, c, sp].join("|");
    const o = agg[k] || (agg[k] = { so, d, c, cls, ln, sp, v: 0 });
    o.v += amt;
  }
  const out = Object.values(agg); // sem arredondar: o total bate com a soma do Excel
  return { rows: out, read: rows, booked, minD, maxD, sheet: best.name, hasSO: best.sidi >= 0, hasClass: best.cli >= 0, hasLine: best.pgi >= 0, hasSeller: best.spi >= 0 || best.nmi >= 0 };
}
/* Índices do Sales Report a partir das linhas guardadas:
   • so[SalesId] = { v, d, c, cls, PT, CP, MH } → liga a cotação ganha ao valor real do pedido;
   • days[AAAA-MM-DD]["classificação|linha"] = valor booked do dia (meta diária). */
export function buildSalesIndex(rows, info) {
  const so = {},
    days = {};
  let minD = "",
    maxD = "",
    hasSO = false,
    hasClass = false,
    hasLine = false,
    hasSeller = false;
  for (const r of rows || []) {
    if (r.sp) hasSeller = true;
    if (!minD || r.d < minD) minD = r.d;
    if (!maxD || r.d > maxD) maxD = r.d;
    if (r.cls) hasClass = true;
    if (r.ln) hasLine = true;
    const dy = days[r.d] || (days[r.d] = {});
    const k = r.cls + "|" + r.ln;
    dy[k] = (dy[k] || 0) + r.v;
    if (r.so) {
      hasSO = true;
      const o = so[r.so] || (so[r.so] = { v: 0, d: r.d, c: r.c, cls: r.cls });
      o.v += r.v;
      if (r.ln) o[r.ln] = (o[r.ln] || 0) + r.v;
      if (r.d < o.d) o.d = r.d;
    }
  }
  return { v: 2, rows: rows || [], so, days, minD, maxD, hasSO, hasClass, hasLine, hasSeller, booked: (rows || []).length, ...(info || {}) };
}
/* Junta um Sales Report novo ao que já estava guardado: o período do relatório novo (do primeiro ao
   último dia dele) é substituído pelas linhas novas; o resto do histórico continua. */
/* Quote Report MAIS VELHO que as cotações guardadas: serve só para completar o que falta (Product Groups e
   linhas PT / CP / MH) nas cotações que já estão guardadas. Situação, valores, pedidos e as cotações que
   não estão guardadas ficam como estão — nada volta para trás. */
export function fillQuotes(oldQuotes, incoming) {
  const inc = new Map((incoming || []).map((q) => [q.id, q]));
  const has = (o) => !!(o && Object.keys(o).length);
  let filled = 0,
    found = 0;
  const quotes = (oldQuotes || []).map((o) => {
    const q = inc.get(o.id);
    if (!q) return o;
    found++;
    const needPg = !has(o.pg) && has(q.pg);
    const needAl = !has(o.al) && has(q.al);
    if (!needPg && !needAl) return o;
    filled++;
    return { ...o, ...(needPg ? { pg: q.pg } : null), ...(needAl ? { al: q.al, lr: true } : null) };
  });
  return { quotes, filled, found, outside: inc.size - found };
}
export function mergeSalesRows(oldRows, rep) {
  const keep = (oldRows || []).filter((r) => r.d < rep.minD || r.d > rep.maxD);
  return [...keep, ...rep.rows].sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : 0));
}

/* ===== situação da cotação (mesmas regras do pulz) ===== */
export const quoteWon = (q) => /confirm/.test(q.st) || (q.so && q.so.length > 0);
export const quoteLost = (q) => /lost|perd|declin/.test(q.st);
export const quoteCancelled = (q) => /cancel/.test(q.st);
export const quoteState = (q) => (quoteWon(q) ? "won" : quoteCancelled(q) ? "cancelled" : quoteLost(q) ? "lost" : "open");
/* Fechamento marcado à mão (FUP → Cancelled / Lost, com o motivo): a cotação que o Quote Report ainda diz
   que está aberta passa a contar como cancelada ou perdida no programa. closeOf(id) = "cancelled" | "lost" | "". */
export const CLOSE_REASONS = { cancelled: ["Budget Only", "Project Discontinued", "Stalled project", "No Response"], lost: ["Lead Time", "Price"] };
// motivos gravados com nomes antigos (em português ou os primeiros em inglês) passam para os nomes atuais
const WHY_OLD = {
  "Levantamento de custo": "Budget Only",
  "Cost Estimate": "Budget Only",
  Desistência: "Project Discontinued",
  Desistence: "Project Discontinued",
  "Customer withdrew": "Project Discontinued",
  Withdrawal: "Project Discontinued",
  "Projeto paralisado": "Stalled project",
  Prazo: "Lead Time",
  Deadline: "Lead Time",
  Preço: "Price",
};
export const normWhy = (w) => WHY_OLD[w] || w || "";
export const CLOSE_LABEL = { cancelled: "Cancelled", lost: "Lost" };
export const quoteStateM = (q, close) => {
  const s = quoteState(q);
  return s === "open" && (close === "cancelled" || close === "lost") ? close : s;
};

// data do pedido: a mais antiga entre as SOs da cotação que já aparecem no Sales Report
const soDateOf = (q, soMap) =>
  (q.so || [])
    .map((s) => soMap && soMap[s] && soMap[s].d)
    .filter(Boolean)
    .sort()[0] || "";
/* Âncora de período (igual ao Customer Tendency do pulz): a cotação GANHA conta no mês do pedido;
   aberta, perdida, ou ganha cujo pedido ainda não chegou no Sales Report, conta na data da cotação. */
export const anchorDate = (q, soMap) => (quoteWon(q) ? soDateOf(q, soMap) || q.d : q.d);

/* Valor VENDIDO de uma cotação ganha: Booked do Sales Id no Sales Report; sem o pedido no relatório,
   cai para o Sales Amt da própria cotação (e, sem ele, para o valor cotado). "used" evita contar o
   mesmo pedido duas vezes quando mais de uma cotação aponta para ele. */
export function soldValue(q, soMap, used) {
  let v = 0,
    found = false;
  for (const s of q.so || []) {
    const o = soMap && soMap[s];
    if (!o) continue;
    found = true;
    if (used) {
      if (used.has(s)) continue;
      used.add(s);
    }
    v += o.v;
  }
  if (found) return { v, src: "sales" };
  return { v: q.sa || q.a || 0, src: "quote" };
}

// Sold só pelo Sales Report (sem cair no Sales Amt da cotação)
export function soldBooked(q, soMap, used) {
  let v = 0;
  for (const s of q.so || []) {
    const o = soMap && soMap[s];
    if (!o || used.has(s)) continue;
    used.add(s);
    v += o.v || 0;
  }
  return { v, src: "sales" };
}

/* ===== recorte por classificação ===== */
export const clsMatch = (cls, sel) => !sel || norm(cls) === norm(sel);
export function clsOptions(quotes) {
  const m = {};
  (quotes || []).forEach((q) => {
    const k = clean(q.cls);
    if (k) m[k] = (m[k] || 0) + 1;
  });
  return Object.entries(m)
    .sort((a, b) => b[1] - a[1])
    .map(([v, n]) => ({ v, n }));
}

/* ===== ano fiscal (julho a junho), como no pulz: FY 2026/27 = jul/2026 a jun/2027 ===== */
export const fyOf = (iso) => {
  const y = +iso.slice(0, 4),
    m = +iso.slice(5, 7);
  return m >= 7 ? y : y - 1;
};
export const fyLabel = (y) => `FY ${y}/${String(y + 1).slice(2)}`;
export const fyRange = (y) => ({ from: `${y}-07-01`, to: `${y + 1}-06-30` });
/* Seletor de período igual ao do pulz: cada ano fiscal é um grupo, com a opção do ano inteiro e os
   meses que têm cotação (a ganha conta no mês do pedido, como na tabela). O ano fiscal atual sempre
   aparece. Valores: "all" · "fy:2026" · "m:2026-10". */
export const MON_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
export const fyRangeLabel = (y) => `Jul/${y} – Jun/${y + 1}`;
export function periodGroups(quotes, sales) {
  const soMap = (sales && sales.so) || null;
  const by = {};
  (quotes || []).forEach((q) => {
    const ym = anchorDate(q, soMap).slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(ym)) return;
    (by[fyOf(ym + "-01")] || (by[fyOf(ym + "-01")] = new Set())).add(ym);
  });
  const cur = fyOf(todayISO());
  if (!by[cur]) by[cur] = new Set();
  return Object.keys(by)
    .map(Number)
    .sort((a, b) => b - a)
    .map((y) => ({ y, label: fyLabel(y), range: fyRangeLabel(y), current: y === cur, months: [...by[y]].sort().reverse() }));
}
export function periodValid(p, groups) {
  if (p === "all") return true;
  if (/^fy:\d{4}$/.test(p || "")) return groups.some((g) => g.y === +p.slice(3));
  if (/^m:\d{4}-\d{2}$/.test(p || "")) return groups.some((g) => g.months.includes(p.slice(2)));
  return false;
}
export function periodLabelOf(p) {
  if (!p || p === "all") return "All periods";
  if (p.startsWith("fy:")) return fyLabel(+p.slice(3));
  if (p.startsWith("m:")) return `${MON_SHORT[+p.slice(7) - 1]}/${p.slice(2, 6)}`;
  return "";
}
export function periodRange(p) {
  if (!p || p === "all") return { from: "", to: "9999-12-31" };
  if (p === "12m") {
    const t = todayISO();
    return { from: addDays(addMonths(t, -12), 1), to: t };
  }
  if (p.startsWith("fy:")) return fyRange(+p.slice(3));
  if (p.startsWith("m:")) return { from: p.slice(2) + "-01", to: lastDayOfMonth(p.slice(2)) };
  return { from: "", to: "9999-12-31" };
}

/* ===== carteira por COTAÇÃO =====
   Uma linha por Quote ID. Quoted = valor cotado; Open, Lost ou Sold recebem o valor conforme a
   situação da cotação. Cotação cancelada fica de fora (mesma regra do pulz); com cancelled = true a função
   devolve SÓ as canceladas (aba Cancelled da carteira, que não entra nos totais das outras). */
export function buildQuoteRows(quotes, sales, { from, to, cls, cancelled, closeOf }) {
  const stOf = (q) => quoteStateM(q, closeOf ? closeOf(q.id) : "");
  const soMap = (sales && sales.so) || null;
  const used = new Set(); // um pedido (SO) ligado a duas cotações conta uma vez só
  return (quotes || [])
    .filter((q) => clsMatch(q.cls, cls) && (cancelled ? stOf(q) === "cancelled" : stOf(q) !== "cancelled"))
    .map((q) => ({ q, ad: anchorDate(q, soMap) }))
    .filter(({ ad }) => (!from || ad >= from) && ad <= to)
    .sort((a, b) => (a.ad < b.ad ? -1 : a.ad > b.ad ? 1 : 0)) // ordem do pedido: a primeira cotação de uma SO leva o valor
    .map(({ q, ad }) => {
      const st = stOf(q);
      /* Sold = Amount (Type Booked) do pedido da cotação no Sales Report, pelo Sales ID — igual a filtrar o
         Sales Report pelos Sales IDs das cotações. Pedido que não está no Sales Report = sem Sold.
         Sem Sales Report importado, usa o Sales Amt da cotação. Com uma linha escolhida, só a parte da linha. */
      const sv = st !== "won" ? { v: 0, src: "" } : soMap ? soldBooked(q, soMap, used) : { v: q.sa || 0, src: "quote" };
      return {
        id: q.id,
        d: q.d,
        ad,
        c: q.c,
        nn: q.nn,
        conta: q.conta,
        st,
        qb: q.qb,
        sp: q.sp,
        quoted: q.a,
        open: st === "open" ? q.a : 0,
        lostV: st === "lost" ? q.a : 0,
        sold: sv.v,
        soldSrc: sv.src,
        sos: (q.so || []).map((s) => ({ id: s, d: (soMap && soMap[s] && soMap[s].d) || "", v: soMap && soMap[s] ? soMap[s].v : null })),
        manual: st !== quoteState(q), // fechada à mão (o Quote Report ainda diz Open)
      };
    });
}

/* ===== campos manuais ===== */
export const FUP = {
  /* Escala de urgência: A (mais urgente) em vermelho → D (mais folgado) em verde.
     "every" diz como o prazo é contado: day = próximo dia útil · twice5 = 2º e 4º dia útil seguintes
     (duas datas) · 2w = 14 dias · month = 1 mês. "urg" é o peso da letra na prioridade. */
  A: { label: "Next day", long: "A · Next business day", every: "day", urg: 0.3, rank: 4, color: "#B42318", bg: "#FEE4E2" },
  B: { label: "2× in 5 days", long: "B · Twice in 5 days (2nd and 4th business day)", every: "twice5", urg: 0.2, rank: 3, color: "#B54708", bg: "#FEF0C7" },
  C: { label: "2 weeks", long: "C · Every 2 weeks", every: "2w", urg: 0.1, rank: 2, color: "#175CD3", bg: "#DCEBFE" },
  D: { label: "1 month", long: "D · Every month", every: "month", urg: 0.05, rank: 1, color: "#067647", bg: "#DCFAE6" },
  /* E = Ongoing: entra sozinho ao marcar Done — a próxima data é a que você escolhe no Done (meta.at).
     Não aparece na lista do FUP para escolher à mão; para sair do E, escolha outra letra. */
  E: { label: "Ongoing", long: "E · Ongoing (next date chosen at Done)", every: "date", urg: 0.15, rank: 2.5, color: "#5B52C9", bg: "#ECEAFB" },
};
export const FUP_KEYS = ["A", "B", "C", "D", "E"]; // todas as letras (filtros, legenda)
export const FUP_PICK = ["A", "B", "C", "D"]; // letras que se escolhem à mão (o E vem do Done)
export const WIN_VALS = [25, 50, 75, 99];
export const WIN_META = {
  25: { color: "#64748B", bg: "#EEF2F5" },
  50: { color: "#B7791F", bg: "#FDF0D9" },
  75: { color: "#2F6FDB", bg: "#E6EEFB" },
  99: { color: "#1F8A5B", bg: "#DDF3E8" },
};
export const TEMP_KEYS = ["HOT", "WARM", "COOL", "COLD"];
export const TEMP = {
  HOT: { label: "Hot", color: "#C2410C", bg: "#FDE6DC", dot: "#E9706B", rank: 4 },
  WARM: { label: "Warm", color: "#B7791F", bg: "#FDF0D9", dot: "#F2B04C", rank: 3 },
  COOL: { label: "Cool", color: "#0F766E", bg: "#D5F5F0", dot: "#3CBFAE", rank: 2 },
  COLD: { label: "Cold", color: "#2B6CB0", bg: "#E1EFFA", dot: "#5AA9D6", rank: 1 },
};

/* ===== Thermometer automático (por cliente) =====
   Hit Rate = valor vendido ÷ valor cotado, das cotações do cliente dos últimos THERMO_HR_MONTHS meses
   (a partir de hoje), na classificação e linha escolhidas; canceladas ficam de fora. Vendido = Booked do pedido no Sales Report,
   ou o Sales Amt da cotação (mesma regra da coluna Sold).

   PASSO 1 · nível pela janela longa (últimos THERMO_WINDOW dias úteis até a data de referência):
   • Hot  = Hit Rate ≥ THERMO_HR_MIN e cotou em THERMO_DAYS_MIN dias ou mais (sempre cota e compra);
   • Warm = Hit Rate ≥ THERMO_HR_MIN e cotou em menos dias (não cota sempre, mas compra quando cota);
   • Cool = Hit Rate < THERMO_HR_MIN e cotou em THERMO_DAYS_MIN dias ou mais (cota, mas nem sempre compra);
   • Cold = Hit Rate < THERMO_HR_MIN e cotou em menos dias (quase não cota e quase não compra).

   PASSO 2 · atividade recente: se nos últimos THERMO_RECENT_WINDOW dias úteis o cliente cotou em
   menos de THERMO_RECENT_MIN dias, ele desce um nível (Hot → Warm, Cool → Cold). Assim, quem cotou
   muito no começo da janela e parou esfria.

   Data de referência = a última data do Quote Report importado (veja thermoRefDate). Com o relatório
   do dia, é o mesmo que "hoje"; com um relatório de dias atrás, os clientes não esfriam só porque o
   arquivo está velho. Cliente que parou de cotar perde dias na janela e esfria sozinho.

   Sábado e domingo não contam no tamanho das janelas. Feriados contam como dia útil. */

// ---------------- CONFIGURAÇÃO DO THERMOMETER (altere aqui) ----------------
export const THERMO_HR_MONTHS = 6; //       Hit Rate = cotações dos últimos N meses, contados a partir de HOJE
export const THERMO_HR_MIN = 0.5; //         Hit Rate mínimo para Hot/Warm nos clientes OEM (0.5 = 50%); abaixo disso, Cool/Cold
export const THERMO_HR_MIN_OTHER = 0.3; //   o mesmo mínimo para as outras classificações (PFP, Distributor, End User, LATAM…): 30%
// mínimo de Hit Rate pela classificação do cliente
export const thermoHrMin = (cls) => (norm(cls) === "oem" ? THERMO_HR_MIN : THERMO_HR_MIN_OTHER);
export const THERMO_WINDOW = 30; //          janela longa, em dias úteis até a data de referência
export const THERMO_DAYS_MIN = 12; //        dias com cotação, na janela longa, para Hot (se Hit Rate ≥ mínimo) ou Cool
export const THERMO_RECENT_WINDOW = 10; //   janela recente, em dias úteis até a data de referência
export const THERMO_RECENT_MIN = 3; //       dias com cotação na janela recente; abaixo disso, desce um nível
export const THERMO_STEP_DOWN = { HOT: "WARM", COOL: "COLD" }; // o que acontece ao descer um nível
// ---------------------------------------------------------------------------

/* Data de referência das janelas: a última data do Quote Report (ou hoje, sem relatório).
   Para contar sempre a partir de hoje, troque por: export const thermoRefDate = (rep, today) => today; */
export const thermoRefDate = (rep, today) => (rep && rep.maxD && rep.maxD < today ? rep.maxD : today);

// passo 1 (janela longa) + passo 2 (atividade recente)
export const thermoOf = (hr, days, recentDays, hrMin = THERMO_HR_MIN) => {
  const base = hr >= hrMin ? (days >= THERMO_DAYS_MIN ? "HOT" : "WARM") : days >= THERMO_DAYS_MIN ? "COOL" : "COLD";
  if (recentDays != null && recentDays < THERMO_RECENT_MIN && THERMO_STEP_DOWN[base]) return THERMO_STEP_DOWN[base];
  return base;
};
/* Início da janela: volta a partir da data de referência até juntar N dias úteis (contando o próprio
   dia de referência quando ele é dia útil). */
export function businessWindowStart(last, n) {
  let d = last,
    count = isWeekend(last) ? 0 : 1;
  while (count < n) {
    d = addDays(d, -1);
    if (!isWeekend(d)) count++;
  }
  return d;
}
/* refDate = data de referência (thermoRefDate). Devolve, por cliente: temp, base (nível antes do passo 2),
   hr, days (janela longa), recent (janela recente) e as datas das janelas. */
export function customerThermo(quotes, sales, cls, refDate, today, closeOf) {
  const hrFrom = addMonths(today || todayISO(), -THERMO_HR_MONTHS); // início do período do Hit Rate
  const soMap = (sales && sales.so) || null;
  const used = new Set();
  const by = {};
  (quotes || [])
    .filter((q) => q.nn && clsMatch(q.cls, cls) && quoteStateM(q, closeOf ? closeOf(q.id) : "") !== "cancelled")
    .map((q) => ({ q, ad: anchorDate(q, soMap) }))
    .sort((a, b) => (a.ad < b.ad ? -1 : a.ad > b.ad ? 1 : 0)) // mesma ordem da tabela: o pedido vai para a 1ª cotação
    .forEach(({ q }) => {
      const o = by[q.nn] || (by[q.nn] = { quoted: 0, sold: 0, last: "", dates: [], cls: "" });
      if (q.cls && q.d >= (o.clsD || "")) {
        o.cls = q.cls; // classificação do cliente = a da cotação mais recente
        o.clsD = q.d;
      }
      // Hit Rate: só as cotações dos últimos meses; os dias com cotação usam as janelas abaixo
      if (q.d >= hrFrom) {
        o.quoted += q.a;
        if (quoteWon(q)) o.sold += soldValue(q, soMap, used).v;
      }
      if (q.d > o.last) o.last = q.d;
      o.dates.push(q.d);
    });
  const ref = refDate || todayISO();
  const from = businessWindowStart(ref, THERMO_WINDOW);
  const recentFrom = businessWindowStart(ref, THERMO_RECENT_WINDOW);
  const out = {};
  for (const nn in by) {
    const o = by[nn];
    const days = new Set(o.dates.filter((d) => d >= from && d <= ref)).size;
    const recent = new Set(o.dates.filter((d) => d >= recentFrom && d <= ref)).size;
    const hr = o.quoted ? o.sold / o.quoted : 0;
    const hrMin = thermoHrMin(o.cls);
    const base = thermoOf(hr, days, null, hrMin);
    out[nn] = { temp: thermoOf(hr, days, recent, hrMin), base, hr, hrMin, cls: o.cls, days, recent, from, recentFrom, ref, hrFrom, last: o.last, quoted: r2(o.quoted), sold: r2(o.sold) };
  }
  return out;
}

/* ===== agenda de follow-up =====
   A partir da "âncora" (dia em que o FUP foi definido, ou do último ciclo concluído com Done):
   A = o próximo dia útil · C = 14 dias depois · D = 1 mês depois (C e D no fim de semana vão para a segunda)
   B = DUAS datas: o 2º e o 4º dia útil depois da âncora. Ex.: marcado na segunda → quarta e sexta;
       marcado na terça → quinta e segunda (o fim de semana não conta).
   fupCycle devolve as datas do ciclo atual; meta.step = quantos follow-ups do ciclo já foram feitos. */
// soma N dias úteis (sábado, domingo e feriado não contam)
export function addBusinessDays(iso, n) {
  let d = iso,
    k = 0;
  while (k < n) {
    d = addDays(d, 1);
    if (!isOffDay(d)) k++;
  }
  return d;
}
export function fupCycle(code, anchor) {
  if (!code || !FUP[code] || !anchor || !/^\d{4}-\d{2}-\d{2}$/.test(anchor)) return [];
  const every = FUP[code].every;
  if (every === "date") return []; // E: a data vem de meta.at (fupPending)
  if (every === "twice5") return [addBusinessDays(anchor, 2), addBusinessDays(anchor, 4)];
  if (every === "day") return [addBusinessDays(anchor, 1)];
  if (every === "2w") return [toBusinessDay(addDays(anchor, 14))];
  return [toBusinessDay(addMonths(anchor, 1))];
}
// datas do ciclo entre from e to (usado nos testes e em telas que precisam do ciclo inteiro)
export function fupDates(code, anchor, from, to) {
  return fupCycle(code, anchor).filter((d) => d >= from && d <= to);
}
// follow-ups ainda pendentes do ciclo (o B tem duas datas; as já feitas saem). O E tem só a data escolhida.
export const fupPending = (meta) =>
  !meta || !meta.fup || meta.close ? [] : meta.fup === "E" ? (/^\d{4}-\d{2}-\d{2}$/.test(meta.at || "") ? [toBusinessDay(meta.at)] : []) : meta.since ? fupCycle(meta.fup, meta.since).slice(meta.step || 0) : [];
/* Data que o ciclo daria para o próximo follow-up se o Done fosse feito em "day" — é a sugestão que
   aparece no Done. (B no 1º Done = a 2ª data do ciclo; E = 1 semana depois.) */
export function fupSuggest(meta, day) {
  if (!meta || !meta.fup) return addBusinessDays(day, 1);
  if (meta.fup === "E") return toBusinessDay(addDays(day, 7));
  const nx = nextFup(fupDone(meta, day));
  return nx && nx > day ? nx : addBusinessDays(day, 1);
}
// próximo follow-up pendente (pode estar no passado = atrasado)
export function nextFup(meta) {
  return fupPending(meta)[0] || "";
}
/* Done: o B tem duas datas no ciclo — o 1º Done passa para a 2ª data; o 2º começa um ciclo novo a
   partir do dia do Done. As outras letras começam o ciclo novo direto. Devolve a meta atualizada. */
export function fupDone(meta, day, nextDate) {
  const m = { ...meta, log: [...((meta && meta.log) || []), day].slice(-50) };
  /* Done com data escolhida: a cotação passa para E · Ongoing, com o próximo follow-up nessa data */
  if (nextDate) {
    m.fup = "E";
    m.at = nextDate;
    m.since = day;
    delete m.step;
    return m;
  }
  const cyc = fupCycle(m.fup, m.since);
  const st = (m.step || 0) + 1;
  // passa para a próxima data do ciclo só se ela ainda não passou; senão o ciclo recomeça no dia do Done
  if (st < cyc.length && cyc[st] >= day) m.step = st;
  else {
    m.since = day;
    delete m.step;
  }
  return m;
}

/* ===== Conserto dos campos manuais a partir do histórico =====
   Cada mudança de FUP, % Win, requisição, Done e troca de data vai para o Activities History. Se o FUP ou o
   % Win de uma cotação ficou diferente do último registro dela no histórico (ex.: outro computador, com uma
   versão antiga do programa ou aberto desde antes, salvou por cima com dados velhos), o histórico vale:
   o campo volta para o que o último registro diz. Devolve { meta, fixed: [ids] }. */
export const REPAIR_E_FROM = "2026-10-23";
export function repairMeta(meta, acts) {
  const last = {};
  for (const a of acts || []) {
    if (!a || !a.id) continue;
    const o = last[a.id] || (last[a.id] = { dones: new Set() });
    const later = (x) => !x || String(a.t) >= String(x.t);
    if ((a.k === "done" || a.k === "resched" || a.k === "fup") && later(o.f)) o.f = a;
    if (a.k === "done") o.dones.add(String(a.t).slice(0, 10));
    if (a.k === "win" && later(o.w)) o.w = a;
    if (a.k === "req" && later(o.r)) o.r = a;
    if ((a.k === "close" || a.k === "reopen") && later(o.c)) o.c = a;
    if ((a.k === "close" || a.k === "why") && later(o.y)) o.y = a;
  }
  const out = { ...(meta || {}) };
  const fixed = [];
  for (const id in last) {
    const o = last[id];
    let cur = { ...(out[id] || {}) };
    let ch = false;
    const f = o.f;
    if (f) {
      const day = String(f.t).slice(0, 10);
      const iso = /^\d{4}-\d{2}-\d{2}$/;
      // Done só passa a cotação para E desde 23/10/2026 (antes o Done seguia o ciclo da letra)
      if ((f.k === "done" || f.k === "resched") && iso.test(f.n || "") && day >= REPAIR_E_FROM) {
        if (cur.fup !== "E" || cur.at !== f.n) {
          cur = { ...cur, fup: "E", at: f.n, since: day };
          delete cur.step;
          ch = true;
        }
      } else if (f.k === "fup") {
        const want = f.n || "";
        if ((cur.fup || "") !== want) {
          if (want) {
            cur.fup = want;
            cur.since = day;
            delete cur.step;
            if (want !== "E") delete cur.at;
          } else {
            delete cur.fup;
            delete cur.since;
            delete cur.step;
            delete cur.at;
          }
          ch = true;
        }
      }
    }
    if (o.w) {
      const want = +o.w.n || 0;
      if ((+cur.win || 0) !== want) {
        if (want) cur.win = want;
        else delete cur.win;
        ch = true;
      }
    }
    if (o.r) {
      const want = String(o.r.n || "");
      if ((cur.req || "") !== want) {
        if (want) cur.req = want;
        else delete cur.req;
        ch = true;
      }
    }
    if (o.c) {
      const want = o.c.k === "close" ? o.c.n || "" : "";
      if ((cur.close || "") !== want) {
        if (want) cur.close = want;
        else delete cur.close;
        ch = true;
      }
    }
    if (o.y) {
      const want = normWhy(String((o.y.k === "close" ? o.y.txt : o.y.n) || ""));
      if ((cur.why || "") !== want) {
        if (want) cur.why = want;
        else delete cur.why;
        ch = true;
      }
    }
    // dias com Done (a visão semanal mostra os Done por eles)
    const ds = [...o.dones].sort().slice(-50);
    const lg = new Set(cur.log || []);
    const miss = ds.filter((d) => !lg.has(d));
    if (miss.length) {
      cur.log = [...lg, ...miss].sort().slice(-50);
      ch = true;
    }
    if (ch) {
      out[id] = { ...cur, name: cur.name || (o.f || o.w || o.r || o.c || o.y || {}).c || "" };
      fixed.push(id);
    }
  }
  return { meta: fixed.length ? out : meta, fixed };
}

/* Cotações em aberto candidatas à meta, com os campos manuais de cada cotação. maxAge = idade máxima da
   cotação em dias (0 = qualquer idade). */
export function openCandidates(quotes, cls, maxAge, metaOf, today, thermo) {
  // metaOf(id) = campos manuais da COTAÇÃO
  return (quotes || [])
    .filter((q) => q.nn && clsMatch(q.cls, cls) && quoteState(q) === "open" && !(metaOf(q.id) || {}).close && q.a > 0)
    .map((q) => {
      const m = metaOf(q.id) || {};
      return { id: q.id, d: q.d, age: daysBetween(q.d, today), c: q.c, nn: q.nn, v: q.a, qb: q.qb, sp: q.sp, win: m.win || 0, temp: (thermo && thermo[q.nn] && thermo[q.nn].temp) || "", fup: m.fup || "", meta: m };
    })
    .filter((x) => !maxAge || x.age <= maxAge);
}
/* Booked entre duas datas (inclusive), na classificação e na linha escolhidas. Usa o total por dia
   do Sales Report; sem Sales Report, as cotações ganhas com âncora no intervalo. */
export function bookedRange(from, to, quotes, sales, cls, line) {
  if (sales && sales.days) {
    let v = 0;
    for (const d in sales.days) {
      if (d < from || d > to) continue;
      const dy = sales.days[d];
      for (const k in dy) {
        const [c, ln] = k.split("|");
        // Sales Report sem coluna de classificação: não dá para recortar, conta tudo
        if (sales.hasClass && !clsMatch(c, cls)) continue;
        if (line && ln !== line) continue;
        v += dy[k];
      }
    }
    return { v: r2(v), src: "sales" };
  }
  const used = new Set();
  let v = 0;
  (quotes || []).forEach((q) => {
    if (!clsMatch(q.cls, cls) || !quoteWon(q)) return;
    const ad = anchorDate(q, null);
    if (ad < from || ad > to) return;
    v += soldValue(q, null, used).v;
  });
  return { v: r2(v), src: quotes && quotes.length ? "quotes" : "none" };
}
// dias úteis (segunda a sexta, sem feriados) entre duas datas, inclusive
export function businessDays(from, to) {
  let n = 0;
  for (let d = from; d <= to; d = addDays(d, 1)) if (!isOffDay(d)) n++;
  return n;
}
/* Meta do dia: o que falta da meta do mês (antes desse dia) dividido pelos dias úteis que sobram,
   contando o próprio dia. Em fim de semana conta como 1 dia, para não dividir por zero. */
export function dailyTarget(goal, bookedBefore, day) {
  const left = Math.max(1, businessDays(day, lastDayOfMonth(day.slice(0, 7))));
  return { v: r2(Math.max(0, goal - bookedBefore) / left), left };
}

/* ===== planejamento dos dias =====
   O FUP é o PRAZO MÁXIMO de cada cotação, contado do dia em que o FUP foi definido (ou do último Done):
   A = o próximo dia útil · B = a próxima segunda ou quinta · C = 14 dias · D = 1 mês (nextFup).
   Para cada dia, de hoje em diante:
   1. entram as cotações cujo prazo vence nesse dia (no dia de hoje, também as atrasadas);
   2. se o valor dessas cotações não cobre a meta do dia, o programa puxa cotações em aberto SEM FUP,
      pela prioridade, até cobrir a meta. Cotação com FUP fica sempre nas datas do FUP dela (o B nas duas).
   Cada cotação fica em um dia só. As que não foram necessárias até o fim do horizonte ficam "sem dia".
   A prioridade (0 a 100) vem só do % Win e do Thermometer — o FUP não entra (ver PRIO_* abaixo). */
/* PREENCHIMENTO DA META (passo 2 acima): desligado por enquanto — cada dia mostra só as cotações
   com FUP naquela data. Para voltar a puxar cotações sem FUP até cobrir a meta, troque para true. */
export const AUTO_FILL = false;
const CHANCE_NO_WIN = { HOT: 0.45, WARM: 0.3, COOL: 0.2, COLD: 0.12 };
export const chanceOf = (x) => (x.win ? x.win / 100 : CHANCE_NO_WIN[x.temp] || 0.08);
// ---------------- CONFIGURAÇÃO DA PRIORIDADE (altere aqui) ----------------
// Priority (0 a 100) = PRIO_WIN × % Win + PRIO_TEMP × Thermometer. O FUP NÃO entra na conta.
export const PRIO_WIN = 0.7; // peso do % Win (25% → 0,25 … 99% → 0,99; sem % Win = 0)
export const PRIO_TEMP = 0.3; // peso do Thermometer
export const PRIO_TEMP_VAL = { HOT: 1, WARM: 0.67, COOL: 0.33, COLD: 0 }; // nota de cada nível (sem Thermometer = 0)
// Ex.: 75% + Hot = 0,7×0,75 + 0,3×1 = 83 · 75% + Warm = 73 · 99% + Cold = 69 · 50% + Hot = 65 · sem % Win + Hot = 30
export function priorityOn(x) {
  const w = x.win ? Math.min(1, x.win / 100) : 0;
  const t = PRIO_TEMP_VAL[x.temp] || 0;
  return { score: Math.round(100 * (PRIO_WIN * w + PRIO_TEMP * t)), chance: chanceOf(x), pw: w, pt: t };
}
/* cands = openCandidates(...); goalOf(dia) = meta do dia; bookedToday = já vendido hoje;
   strategy = ordem do preenchimento: "priority" | "likely" | "biggest". */
export function planDays(cands, { today, horizon, goalOf, bookedToday, strategy, excluded }) {
  const tr = (t) => (TEMP[t] ? TEMP[t].rank : 0);
  const live = cands
    .filter((c) => !(excluded && excluded.has(c.id)))
    .map((c) => {
      const dl = c.fup ? nextFup(c.meta || {}) : "";
      return { ...c, dl, overdue: !!dl && dl < today };
    });
  const byDay = {},
    dayOf = {};
  for (let d = today; d <= horizon; d = addDays(d, 1)) {
    const list = [];
    let planned = 0;
    const put = (x, kind) => {
      const p = priorityOn(x, d);
      list.push({ ...x, kind, ...p, exp: r2(x.v * p.chance) });
      dayOf[x.id] = d;
      planned += x.v;
    };
    // 1) prazo do FUP vence neste dia (hoje: também os atrasados)
    for (const x of live) if (!dayOf[x.id] && x.dl && (x.dl === d || (d === today && x.dl < today))) put(x, x.dl < today ? "overdue" : "deadline");
    // 2) completa a meta do dia com as de maior prioridade
    const goal = goalOf(d);
    const need = r2(Math.max(0, goal - (d === today ? bookedToday : 0)));
    if (AUTO_FILL && planned < need) {
      /* FUP definido (ou Done) nesse mesmo dia: o contato já foi feito hoje, então a cotação espera o
         prazo dela (um D marcado hoje fica para amanhã) e não é puxada para hoje. */
      const doneOn = (x) => !!(x.fup && x.meta && (x.meta.since === d || (x.meta.log || []).includes(d)));
      // só cotações SEM FUP são puxadas: quem tem FUP fica sempre nas datas do FUP dele
      const pool = live
        .filter((x) => !x.fup && !dayOf[x.id] && !doneOn(x))
        .map((x) => ({ x, p: priorityOn(x, d) }))
        .sort((a, b) =>
          strategy === "biggest"
            ? b.x.v - a.x.v || b.p.score - a.p.score
            : strategy === "likely"
            ? b.p.chance - a.p.chance || tr(b.x.temp) - tr(a.x.temp) || b.x.v - a.x.v
            : b.p.score - a.p.score || b.x.v - a.x.v,
        );
      for (const { x } of pool) {
        if (planned >= need) break;
        put(x, "fill");
      }
    }
    const order = { overdue: 0, deadline: 1, again: 2, fill: 3 };
    list.sort((a, b) => order[a.kind] - order[b.kind] || b.score - a.score || b.v - a.v);
    let cum = 0;
    list.forEach((x) => {
      cum += x.v;
      x.cum = r2(cum);
    });
    byDay[d] = { goal, need, planned: r2(planned), list, expected: r2(list.reduce((s, x) => s + x.exp, 0)) };
  }
  /* B (2× em 5 dias): a 2ª data do ciclo também aparece no calendário e na lista do dia ("again"),
     sem somar de novo o valor no plano — é a mesma cotação. */
  for (const x of live) {
    const pend = fupPending(x.meta || {});
    pend.slice(1).forEach((d) => {
      if (d < today || d > horizon || !byDay[d] || dayOf[x.id] === d) return;
      const p = priorityOn(x, d);
      byDay[d].list.push({ ...x, kind: "again", ...p, exp: 0, cum: null });
      byDay[d].hasAgain = true;
    });
  }
  const order = { overdue: 0, deadline: 1, again: 2, fill: 3 };
  for (const d in byDay) if (byDay[d].hasAgain) byDay[d].list.sort((a, b) => order[a.kind] - order[b.kind] || b.score - a.score || b.v - a.v);
  return { byDay, dayOf, live };
}

/* ===== Relatório de FUP por vendedor (Sales Person) =====
   Para o período from–to, nas cotações do recorte da tela:
   • Atualizada      = a cotação teve Done no período (vale a hora do 1º Done do período);
   • Após 11h        = esse 1º Done foi às LATE_HOUR (11:00) ou depois;
   • Para atualizar  = cotação em aberto, com FUP, sem Done no período e com o follow-up vencido até o fim do
                       período (no período atual, também o que ainda vence até o fim dele);
   • Qtd Cotações    = atualizadas + para atualizar (cada cotação conta uma vez).
   "groupOf" agrupa as grafias do Sales Person num nome só (como no filtro de vendedor). */
export const LATE_HOUR = 11;
export function sellerFupReport(quotes, meta, acts, from, to, groupOf) {
  const qmap = new Map((quotes || []).map((q) => [q.id, q]));
  const firstDone = {};
  for (const a of acts || []) {
    if (a.k !== "done" || !qmap.has(a.id)) continue;
    const d = String(a.t).slice(0, 10);
    if (d < from || d > to) continue;
    if (!firstDone[a.id] || a.t < firstDone[a.id].t) firstDone[a.id] = a;
  }
  const by = {};
  const row = (name) => by[name] || (by[name] = { name, total: 0, upd: 0, late: 0, pend: 0, items: [] });
  for (const q of quotes || []) {
    const dn = firstDone[q.id];
    const m = (meta && meta[q.id]) || {};
    let item = null;
    if (dn) {
      const hh = +String(dn.t).slice(11, 13);
      item = { id: q.id, kind: "upd", t: dn.t, late: hh >= LATE_HOUR, next: dn.n || "" };
    } else if (quoteState(q) === "open" && m.fup && !m.close) {
      const dl = nextFup(m);
      if (dl && dl <= to) item = { id: q.id, kind: "pend", dl, overdue: dl < from };
    }
    if (!item) continue;
    const r = row(groupOf ? groupOf(q.sp) : sellerDisplay(q.sp) || "");
    r.total++;
    if (item.kind === "upd") {
      r.upd++;
      if (item.late) r.late++;
    } else r.pend++;
    r.items.push({ ...item, c: q.c, a: q.a, fup: m.fup || "", win: m.win || 0 });
  }
  return Object.values(by)
    .map((r) => ({ ...r, pctUpd: r.total ? r.upd / r.total : 0, pctPend: r.total ? r.pend / r.total : 0 }))
    .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, "pt-BR"));
}

/* ===== relatório de follow-ups feitos (Done) num período =====
   Uma linha por cotação que teve Done entre "from" e "to" (o último Done do período):
   • comentário novo = a nota gravada com esse Done (o Done sempre grava uma nota);
   • comentário antigo = a nota anterior a ela;
   • follow-up date = a data do próximo follow-up escolhida nesse Done.
   "quotes" já vem no recorte da tela (classificação, linha, vendedor). */
export function fupReportRows(acts, quotes, meta, thermo, from, to) {
  const qmap = new Map((quotes || []).map((q) => [q.id, q]));
  const ord = (acts || []).map((a, i) => ({ ...a, _i: i })).sort((a, b) => (a.t < b.t ? -1 : a.t > b.t ? 1 : a._i - b._i));
  const notes = {},
    last = {};
  for (const a of ord) {
    if (a.k === "note") (notes[a.id] = notes[a.id] || []).push(a);
    if (a.k === "done") {
      const d = String(a.t).slice(0, 10);
      if (d >= from && d <= to && qmap.has(a.id)) last[a.id] = a;
    }
  }
  return Object.values(last)
    .map((dn) => {
      const q = qmap.get(dn.id);
      const ns = notes[dn.id] || [];
      let k = -1;
      ns.forEach((n, i) => {
        if (n.t < dn.t || (n.t === dn.t && n._i < dn._i)) k = i;
      });
      const m = (meta && meta[dn.id]) || {};
      const t = thermo && thermo[q.nn];
      return {
        id: q.id,
        amount: q.a,
        c: q.c,
        nw: k >= 0 ? ns[k] : null,
        old: k > 0 ? ns[k - 1] : null,
        req: m.req || "",
        win: m.win || 0,
        temp: t ? t.temp : "",
        next: dn.n && /^\d{4}-\d{2}-\d{2}$/.test(dn.n) ? dn.n : "",
        done: dn.t,
        by: dn.by || "",
        qb: q.qb || "",
        sp: q.sp || "",
      };
    })
    .sort((a, b) => sellerDisplay(a.sp).localeCompare(sellerDisplay(b.sp), "pt-BR") || a.c.localeCompare(b.c, "pt-BR") || (a.id < b.id ? -1 : 1));
}

/* ===== recortes: classificação, linha de produto, vendedor interno e externo =====
   • line = PT / CP / MH: entra a cotação que tem item dessa linha, e os valores passam a ser só os
     da linha (o Sales Amt é rateado pela parte da linha no valor cotado);
   • who = vendedor: interno (coluna "Quoted By") ou externo (coluna "Sales Person"), com as grafias
     agrupadas como no pulz. */
/* who = vendedor escolhido: { qb: [grafias no Quoted By], sp: [grafias no Sales Person] }. A cotação
   entra quando bate em QUALQUER uma das duas colunas — assim quem é interno e também aparece como
   Sales Person traz todas as cotações dele, sem contar nenhuma duas vezes. */
export function scopeQuotes(quotes, { cls, line, who }) {
  const qbs = who && who.qb && who.qb.length ? new Set(who.qb) : null,
    sps = who && who.sp && who.sp.length ? new Set(who.sp) : null;
  const byWho = (q) => !who || (qbs && qbs.has(clean(q.qb))) || (sps && sps.has(clean(q.sp)));
  return (quotes || [])
    .filter((q) => clsMatch(q.cls, cls) && byWho(q) && (!line || (q.al && q.al[line] > 0)))
    .map((q) => {
      if (!line) return q;
      const v = q.al[line];
      const share = q.a ? v / q.a : 0;
      return { ...q, a: v, sa: r2((q.sa || 0) * share) };
    });
}
// pedidos do Sales Report só com o valor da linha escolhida (sem a coluna de linha no relatório, fica o total)
/* Booked por dia só do vendedor escolhido (coluna Sales Responsible/Name do Sales Report). O valor das
   cotações (Sold, pelo Sales ID) não muda: só o booked de cada dia (calendário e meta diária).
   Sales Report guardado por uma versão antiga (sem o vendedor): não dá para recortar e fica como está. */
export const spCode = (s) => ((String(s || "").match(/^\s*(SP\d+)/i) || [])[1] || "").toUpperCase();
export function sellerMatcher(list) {
  const codes = new Set((list || []).map(spCode).filter(Boolean));
  const memo = new Map();
  return (s) => {
    if (!s) return false;
    if (memo.has(s)) return memo.get(s);
    const c = spCode(s);
    const ok = (c && codes.has(c)) || (list || []).some((x) => sellerSame(x, s));
    memo.set(s, ok);
    return ok;
  };
}
export function salesForSeller(sales, who) {
  if (!sales || !who || !(who.sp || []).length || !sales.hasSeller) return sales;
  const match = sellerMatcher(who.sp);
  const days = {};
  for (const r of sales.rows || []) {
    if (!match(r.sp)) continue;
    const dy = days[r.d] || (days[r.d] = {});
    const k = r.cls + "|" + r.ln;
    dy[k] = (dy[k] || 0) + r.v;
  }
  return { ...sales, days, bySeller: true };
}
export function salesForLine(sales, line) {
  if (!sales || !line || !sales.hasLine) return sales;
  const so = {};
  for (const k in sales.so) so[k] = { ...sales.so[k], v: sales.so[k][line] || 0 };
  return { ...sales, so };
}
/* ===== vendedores (mesma regra do pulz) =====
   Interno = coluna "Quoted By" e Externo = coluna "Sales Person", ambas do Quote Report.
   O Quote Report escreve o vendedor como "SP0178 - Julia Pulz Campo". Como no pulz, o nome é limpo
   (tira o código SP, os acentos e as partículas da/de/dos) e duas grafias são a MESMA pessoa quando têm
   o mesmo primeiro nome e um sobrenome igual, contido no outro ou com uma letra de diferença
   ("Matheus Galo" → Gallo). Nome de uma palavra só precisa ser exatamente igual. */
const SELLER_SKIP = new Set(["DA", "DE", "DO", "DOS", "DAS", "E"]);
const sellerTokens = (s) =>
  String(s || "")
    .replace(/^SP\d+\s*[-–]\s*/i, "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Z ]+/g, " ")
    .split(/\s+/)
    .filter((t) => t.length > 1 && !SELLER_SKIP.has(t));
function near1(a, b) {
  if (a === b) return true;
  if (Math.abs(a.length - b.length) > 1) return false;
  let i = 0,
    j = 0,
    diff = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) {
      i++;
      j++;
      continue;
    }
    if (++diff > 1) return false;
    if (a.length > b.length) i++;
    else if (b.length > a.length) j++;
    else {
      i++;
      j++;
    }
  }
  return diff + (a.length - i) + (b.length - j) <= 1;
}
export function sellerSame(a, b) {
  const p = sellerTokens(a),
    t = sellerTokens(b);
  if (!p.length || !t.length || p[0] !== t[0]) return false;
  if (p.length === 1 || t.length === 1) return p.length === t.length;
  const pl = p[p.length - 1],
    tl = t[t.length - 1];
  return pl === tl || p.every((x) => t.includes(x)) || t.every((x) => p.includes(x)) || near1(pl, tl);
}
// nome para mostrar: sem o código SP e com espaços normais
export const sellerDisplay = (v) =>
  clean(v)
    .replace(/^SP\d+\s*[-–]\s*/i, "")
    .replace(/\s+/g, " ");
/* Vendedores de uma coluna ("qb" = Quoted By, "sp" = Sales Person), já agrupando as grafias da mesma
   pessoa. Cada grupo: { name (a grafia mais usada, sem o código SP), n (cotações), raw (textos exatos
   do relatório que pertencem a ele) }. */
/* Lista do seletor Salesperson. Quem está nas duas colunas aparece só em Internal, e o número dele conta
   as cotações das duas colunas (sem repetir). */
export function sellerLists(quotes) {
  const ints = sellerOptions(quotes, "qb"),
    exts = sellerOptions(quotes, "sp");
  const external = [];
  exts.forEach((e) => {
    const i = ints.find((g) => e.raw.some((r) => sellerSame(g.name, r)));
    if (i) i.spRaw = [...(i.spRaw || []), ...e.raw];
    else external.push({ ...e, spRaw: e.raw, raw: [] });
  });
  const count = (g) => {
    const a = new Set(g.raw),
      b = new Set(g.spRaw || []);
    return (quotes || []).filter((q) => a.has(clean(q.qb)) || b.has(clean(q.sp))).length;
  };
  const internal = ints.map((g) => ({ ...g, spRaw: g.spRaw || [], n: g.spRaw ? count(g) : g.n }));
  return { internal, external };
}
export function sellerOptions(quotes, field) {
  const cnt = {};
  (quotes || []).forEach((q) => {
    const k = clean(q[field]);
    if (k) cnt[k] = (cnt[k] || 0) + 1;
  });
  const groups = [];
  Object.entries(cnt)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "pt-BR"))
    .forEach(([raw, n]) => {
      const g = groups.find((x) => sellerSame(x.name, raw));
      if (g) {
        g.n += n;
        g.raw.push(raw);
      } else groups.push({ v: sellerDisplay(raw), name: sellerDisplay(raw), n, raw: [raw] });
    });
  // nome mostrado: a grafia mais usada que não está toda em maiúsculas; se todas estiverem, em "Nome Sobrenome"
  const title = (t) => t.toLowerCase().replace(/(^|\s)\S/g, (m) => m.toUpperCase());
  groups.forEach((g) => {
    const nice = g.raw.find((r) => /[a-zà-ú]/.test(sellerDisplay(r)));
    g.name = g.v = nice ? sellerDisplay(nice) : title(sellerDisplay(g.raw[0]));
  });
  return groups.sort((a, b) => a.name.localeCompare(b.name, "pt-BR", { sensitivity: "base" }));
}

/* ===== data e hora local (Activities History) ===== */
export const nowStamp = () => {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
};
export const fmtStamp = (t) => {
  const m = String(t || "").match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]} ${m[4]}:${m[5]}` : String(t || "");
};

/* ===== junção do Quote Report novo com as cotações guardadas =====
   Quote ID novo entra; Quote ID que já existia é atualizado e as mudanças vão para o Activities
   History (situação, valor cotado, pedido novo). Cotação que não veio no relatório novo continua
   guardada (o relatório novo pode cobrir outro período). "fs" = quando a cotação entrou no programa. */
export const STATE_LABEL = { open: "Open", won: "Won", lost: "Lost", cancelled: "Cancelled" };
export function mergeQuotes(oldQuotes, incoming, stamp, dropped) {
  const by = new Map((oldQuotes || []).map((q) => [q.id, q]));
  const acts = [];
  let added = 0,
    updated = 0,
    removed = 0;
  // cotação que, no relatório novo, só tem Product Group que não conta: sai da base
  for (const id of dropped || []) if (by.delete(id)) removed++;
  for (const q of incoming || []) {
    const o = by.get(q.id);
    if (!o) {
      by.set(q.id, { ...q, fs: stamp, up: stamp });
      added++;
      continue;
    }
    const ch = [];
    const so = quoteState(o),
      sn = quoteState(q);
    if (so !== sn) ch.push({ k: "status", o: STATE_LABEL[so], n: STATE_LABEL[sn] });
    if (Math.abs((o.a || 0) - (q.a || 0)) > 0.005) ch.push({ k: "amount", o: o.a, n: q.a });
    const newSO = (q.so || []).filter((x) => !(o.so || []).includes(x));
    if (newSO.length) ch.push({ k: "order", o: "", n: newSO.join(", ") });
    // "up" = última vez que a cotação veio num Quote Report: a versão mais nova sempre ganha na junção com o arquivo
    by.set(q.id, { ...q, fs: o.fs || stamp, up: stamp });
    if (ch.length) {
      updated++;
      ch.forEach((c) => acts.push({ t: stamp, id: q.id, c: q.c, ...c }));
    }
  }
  const quotes = [...by.values()];
  let minD = "",
    maxD = "";
  quotes.forEach((q) => {
    if (!minD || q.d < minD) minD = q.d;
    if (!maxD || q.d > maxD) maxD = q.d;
  });
  return { quotes, added, updated, removed, acts, minD, maxD };
}

/* ===== arquivo de dados (Excel) =====
   Guarda tudo o que o programa precisa para não depender do cache do navegador. Abas:
   Quotes · Sales · Follow-up · Activities · Goals · Info. O programa reescreve o arquivo inteiro a cada
   mudança; as anotações (Notes) e o histórico nunca são alterados nem apagados pelo programa. */
export const DATA_APP = "Quote Follow Up";
const DATA_APP_OLD = "Quote Pipeline"; // nome antigo do programa (arquivos gravados antes da troca)
/* versão do programa (data): vai para a aba Info dos arquivos. Um programa mais velho que o arquivo não
   deixa salvar (pede para recarregar a página), para não gravar por cima com regras antigas. */
export const APP_VERSION = "2026-11-01";

/* ===== Usuários e senhas (tela de login) =====
   A senha nunca é guardada: o arquivo de dados guarda só o "hash" (SHA-256 repetido PASS_ROUNDS vezes,
   com um "sal" aleatório por usuário). Para conferir, o programa calcula o hash da senha digitada e compara.
   users = { [nome]: { role: "admin" | "seller", salt, hash, at } } · hash vazio = ainda vai criar a senha. */
export const ADMIN_USER = "Admin";
export const PASS_ROUNDS = 5000;
const SHA_K = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f,
  0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
  0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7,
  0xc67178f2,
];
// SHA-256 de um texto (UTF-8) → 64 caracteres hexadecimais
export function sha256(text) {
  const bytes = new TextEncoder().encode(text);
  const l = bytes.length;
  const n = (((l + 9 + 63) >> 6) << 6);
  const m = new Uint8Array(n);
  m.set(bytes);
  m[l] = 0x80;
  const bits = l * 8;
  m[n - 4] = (bits >>> 24) & 255;
  m[n - 3] = (bits >>> 16) & 255;
  m[n - 2] = (bits >>> 8) & 255;
  m[n - 1] = bits & 255;
  m[n - 5] = Math.floor(bits / 2 ** 32) & 255;
  const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const w = new Uint32Array(64);
  const rotr = (x, k) => (x >>> k) | (x << (32 - k));
  for (let o = 0; o < n; o += 64) {
    for (let i = 0; i < 16; i++) w[i] = (m[o + 4 * i] << 24) | (m[o + 4 * i + 1] << 16) | (m[o + 4 * i + 2] << 8) | m[o + 4 * i + 3];
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
    }
    let [a, b, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i++) {
      const t1 = (h + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + SHA_K[i] + w[i]) | 0;
      const t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      h = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0;
    H[1] = (H[1] + b) | 0;
    H[2] = (H[2] + c) | 0;
    H[3] = (H[3] + d) | 0;
    H[4] = (H[4] + e) | 0;
    H[5] = (H[5] + f) | 0;
    H[6] = (H[6] + g) | 0;
    H[7] = (H[7] + h) | 0;
  }
  return H.map((x) => (x >>> 0).toString(16).padStart(8, "0")).join("");
}
export function hashPassword(pwd, salt) {
  let h = sha256(salt + ":" + pwd);
  for (let i = 1; i < PASS_ROUNDS; i++) h = sha256(h + salt);
  return h;
}
export const newSalt = () => {
  const a = new Uint8Array(16);
  if (typeof crypto !== "undefined" && crypto.getRandomValues) crypto.getRandomValues(a);
  else for (let i = 0; i < 16; i++) a[i] = (Math.random() * 256) | 0;
  return Array.from(a, (b) => b.toString(16).padStart(2, "0")).join("");
};
/* Cadastro de usuários gravado CIFRADO no arquivo (aba "Config", uma coluna): quem abrir o Excel não vê
   nomes, papéis nem senhas, nem sabe qual linha é de quem. Cifra de fluxo com SHA-256 (chave do programa +
   vetor aleatório a cada gravação) e um código de conferência: se alguém mexer no texto, ele é rejeitado. */
const CFG_SECRET = "qp·cfg·7Hq2#vN9!x4Lr·" + DATA_APP_OLD; // fixa: não muda com o nome do programa
const hexBytes = (h) => Uint8Array.from(h.match(/../g), (x) => parseInt(x, 16));
const toB64 = (u8) => {
  let s = "";
  for (let i = 0; i < u8.length; i += 8192) s += String.fromCharCode.apply(null, u8.subarray(i, i + 8192));
  return btoa(s);
};
const fromB64 = (b) => Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
const keystreamXor = (bytes, iv) => {
  const key = sha256(CFG_SECRET);
  const out = new Uint8Array(bytes.length);
  for (let o = 0, blk = 0; o < bytes.length; o += 32, blk++) {
    const ks = hexBytes(sha256(key + ":" + iv + ":" + blk));
    for (let i = 0; i < 32 && o + i < bytes.length; i++) out[o + i] = bytes[o + i] ^ ks[i];
  }
  return out;
};
export function sealText(txt) {
  const iv = newSalt();
  const ct = toB64(keystreamXor(new TextEncoder().encode(txt), iv));
  const mac = sha256(sha256(CFG_SECRET) + "|" + iv + "|" + ct).slice(0, 32);
  return `q1.${iv}.${mac}.${ct}`;
}
export function openText(sealed) {
  const m = /^q1\.([0-9a-f]{32})\.([0-9a-f]{32})\.([A-Za-z0-9+/=]*)$/.exec(String(sealed || ""));
  if (!m) return null;
  const [, iv, mac, ct] = m;
  if (sha256(sha256(CFG_SECRET) + "|" + iv + "|" + ct).slice(0, 32) !== mac) return null;
  try {
    return new TextDecoder().decode(keystreamXor(fromB64(ct), iv));
  } catch {
    return null;
  }
}
// junta dois cadastros de usuários: por nome, vale o mais recente (criar, trocar ou zerar a senha)
export function mergeUsers(a, b) {
  const out = { ...(a || {}) };
  Object.entries(b || {}).forEach(([k, u]) => {
    if (!out[k] || String(u.at || "") > String(out[k].at || "")) out[k] = u;
  });
  return out;
}
export const ACT_TYPES = {
  note: "Note",
  fup: "FUP",
  win: "% Win",
  done: "Follow-up done",
  resched: "Next follow-up changed",
  req: "Customer requisition",
  plan: "Plan",
  status: "Status (Quote Report)",
  amount: "Amount (Quote Report)",
  order: "Order (Quote Report)",
  close: "Closed by hand",
  reopen: "Reopened",
  why: "Close reason",
  dcs: "DCS / Value Added form", // o formulário inteiro (JSON) fica na nota; n = Value Added / DCS / Deslealdade
};
const ACT_BY_LABEL = Object.fromEntries(Object.entries(ACT_TYPES).map(([k, v]) => [v, k]));
const sheetFrom = (rows, head, widths) => {
  const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
  ws["!cols"] = widths.map((w) => ({ wch: w }));
  return ws;
};
export function buildDataWorkbook(st) {
  const wb = XLSX.utils.book_new();
  const quotes = (st.quotesRep && st.quotesRep.quotes) || [];
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom(
      quotes.map((q) => [q.id, q.d, q.c, q.conta || "", q.cls || "", q.st || "", q.qb || "", q.sp || "", q.a || 0, q.sa || 0, (q.so || []).join(", "), (q.al && q.al.PT) || 0, (q.al && q.al.CP) || 0, (q.al && q.al.MH) || 0, q.fs || "", q.up || "", q.lr || (q.al && Object.keys(q.al).length) ? "yes" : "", q.pg && Object.keys(q.pg).length ? JSON.stringify(Object.fromEntries(Object.entries(q.pg).map(([k, v]) => [k, Math.round(v * 100) / 100]))) : ""]),
      ["Quote ID", "Quote date", "Customer", "Account", "Classification", "Status", "Quoted by", "Sales person", "Amount", "Sales amt", "Sales IDs", "PT", "CP", "MH", "First seen", "Updated", "Line read", "Product groups"],
      [14, 11, 40, 10, 14, 12, 18, 18, 13, 13, 22, 12, 12, 12, 18, 18, 10, 40],
    ),
    "Quotes",
  );
  const sales = (st.salesRep && st.salesRep.rows) || [];
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom(sales.map((r) => [r.so, r.d, r.c, r.cls, r.ln, r.v, r.sp || ""]), ["Sales ID", "Date", "Customer", "Classification", "Line", "Booked", "Salesperson"], [14, 11, 40, 14, 6, 13, 30]),
    "Sales",
  );
  const meta = st.meta || {},
    excl = new Set(st.excl || []);
  const fids = [...new Set([...Object.keys(meta), ...excl])];
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom(
      fids.map((id) => {
        const m = meta[id] || {};
        return [id, m.name || "", m.fup || "", m.since || "", m.win || "", excl.has(id) ? "Yes" : "", (m.log || []).join(", "), m.step || "", m.at || "", m.req || "", m.close || "", m.why || ""];
      }),
      ["Quote ID", "Customer", "FUP", "FUP since", "% Win", "Left out of plan", "Follow-ups done", "Cycle step", "Next date (E)", "Customer requisition", "Closed as", "Close reason"],
      [14, 40, 6, 11, 7, 15, 30, 10, 13, 30, 11, 22],
    ),
    "Follow-up",
  );
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom(
      (st.acts || []).map((a) => [a.t, a.id, a.c || "", ACT_TYPES[a.k] || a.k, a.o ?? "", a.n ?? "", a.txt || "", a.by || ""]),
      ["Date & time", "Quote ID", "Customer", "Type", "From", "To", "Note", "By"],
      [19, 14, 40, 22, 16, 16, 60, 18],
    ),
    "Activities",
  );
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom([["Daily goal", st.dailyGoal || 0], ...Object.entries(st.dayGoals || {}).sort().map(([d, v]) => [d, v])], ["Day (or 'Daily goal')", "Goal"], [20, 14]),
    "Goals",
  );
  // usuários (só no arquivo "data"): tudo cifrado na aba "Config" — nome, papel e senha não aparecem
  if (st.users) {
    const sealed = sealText(JSON.stringify(st.users));
    const parts = [];
    for (let i = 0; i < sealed.length; i += 30000) parts.push([sealed.slice(i, i + 30000)]);
    XLSX.utils.book_append_sheet(wb, sheetFrom(parts, ["Data"], [120]), "Config");
  }
  const qr = st.quotesRep || {},
    sr = st.salesRep || {};
  XLSX.utils.book_append_sheet(
    wb,
    sheetFrom(
      [
        ["App", DATA_APP],
        ["Version", 1],
        ["Saved at", nowStamp()],
        ["Program version", APP_VERSION],
        ["Last Quote Report", qr.fileName || ""],
        ["Quote Report imported at", qr.at || ""],
        ["Last Sales Report", sr.fileName || ""],
        ["Sales Report imported at", sr.at || ""],
        ["Note", "Saved by the program. Do not edit by hand: the Notes and the Activities History are kept as they were written."],
      ],
      ["Item", "Value"],
      [26, 100],
    ),
    "Info",
  );
  return wb;
}
const rowsOf = (wb, name) => (wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" }).slice(1) : []);
const str = (v) => (v == null ? "" : String(v).trim());
const isoOf = (v) => (typeof v === "number" ? toISO(v) : str(v).slice(0, 10));
export function isDataWorkbook(wb) {
  const info = rowsOf(wb, "Info");
  return info.some((r) => str(r[0]) === "App" && (str(r[1]) === DATA_APP || str(r[1]) === DATA_APP_OLD)) || (wb.SheetNames.includes("Quotes") && wb.SheetNames.includes("Activities"));
}
// usuários do arquivo: aba "Config" cifrada; arquivos antigos tinham a aba "Users" aberta (convertida no próximo Save)
// (texto cifrado alterado ou corrompido → lista vazia e usersBad, para o programa avisar e manter os usuários que já conhece)
function readUsers(wb) {
  if (wb.Sheets.Config) {
    const txt = openText(rowsOf(wb, "Config").map((r) => str(r[0])).join(""));
    try {
      const u = txt && JSON.parse(txt);
      if (u && typeof u === "object") return { users: u, usersBad: false };
    } catch {}
    return { users: {}, usersBad: true };
  }
  const users = Object.fromEntries(
    rowsOf(wb, "Users")
      .filter((r) => str(r[0]))
      .map((r) => [str(r[0]), { role: str(r[1]) === "admin" ? "admin" : "seller", salt: str(r[2]), hash: str(r[3]).toLowerCase(), at: str(r[4]) }]),
  );
  return { users, usersBad: false };
}
export function parseDataWorkbook(wb) {
  const qHead = wb.Sheets.Quotes ? (XLSX.utils.sheet_to_json(wb.Sheets.Quotes, { header: 1, raw: true, defval: "" })[0] || []).map((h) => str(h)) : [];
  const lri = qHead.indexOf("Line read");
  const pgi = qHead.indexOf("Product groups");
  const quotes = rowsOf(wb, "Quotes")
    .filter((r) => str(r[0]))
    .map((r) => {
      const al = {};
      [["PT", 11], ["CP", 12], ["MH", 13]].forEach(([k, i]) => {
        const v = toNum(r[i]);
        if (v) al[k] = v;
      });
      const c = str(r[2]);
      // "Line read" = yes: a linha dos itens foi lida do Quote Report (arquivo antigo não tem a coluna)
      const lr = lri >= 0 && str(r[lri]).toLowerCase() === "yes";
      let pg;
      if (pgi >= 0 && str(r[pgi])) {
        try {
          pg = JSON.parse(str(r[pgi]));
        } catch {}
      }
      return { id: str(r[0]).toUpperCase(), d: isoOf(r[1]), c, conta: str(r[3]), cls: str(r[4]), st: str(r[5]), qb: str(r[6]), sp: str(r[7]), a: toNum(r[8]), sa: toNum(r[9]), so: str(r[10]) ? str(r[10]).split(/\s*,\s*/) : [], al, lr, ...(pg ? { pg } : {}), fs: str(r[14]), up: str(r[15]), nn: normName(c) };
    });
  let minD = "",
    maxD = "";
  quotes.forEach((q) => {
    if (!minD || q.d < minD) minD = q.d;
    if (!maxD || q.d > maxD) maxD = q.d;
  });
  const info = Object.fromEntries(rowsOf(wb, "Info").map((r) => [str(r[0]), str(r[1])]));
  const salesRows = rowsOf(wb, "Sales")
    .filter((r) => isoOf(r[1]))
    .map((r) => ({ so: str(r[0]), d: isoOf(r[1]), c: str(r[2]), cls: str(r[3]), ln: str(r[4]), v: toNum(r[5]), sp: str(r[6]) }));
  const meta = {},
    excl = [];
  rowsOf(wb, "Follow-up").forEach((r) => {
    const id = str(r[0]).toUpperCase();
    if (!id) return;
    const m = {};
    if (str(r[1])) m.name = str(r[1]);
    if (FUP[str(r[2])]) m.fup = str(r[2]);
    if (isoOf(r[3])) m.since = isoOf(r[3]);
    if (toNum(r[4])) m.win = toNum(r[4]);
    if (str(r[6])) m.log = str(r[6]).split(/\s*,\s*/);
    if (toNum(r[7])) m.step = toNum(r[7]);
    if (isoOf(r[8])) m.at = isoOf(r[8]);
    if (str(r[9])) m.req = str(r[9]);
    if (/^(cancelled|lost)$/.test(str(r[10]))) m.close = str(r[10]);
    if (str(r[11])) m.why = normWhy(str(r[11]));
    if (Object.keys(m).some((k) => k !== "name")) meta[id] = m;
    if (/^y/i.test(str(r[5]))) excl.push(id);
  });
  const acts = rowsOf(wb, "Activities")
    .filter((r) => str(r[0]) && str(r[1]))
    .map((r) => {
      const a = { t: str(r[0]), id: str(r[1]).toUpperCase(), c: str(r[2]), k: ACT_BY_LABEL[str(r[3])] || str(r[3]) };
      if (str(r[4]) !== "") a.o = r[4];
      if (str(r[5]) !== "") a.n = r[5];
      if (str(r[6])) a.txt = str(r[6]);
      if (str(r[7])) a.by = str(r[7]);
      return a;
    });
  let dailyGoal = 0;
  const dayGoals = {};
  rowsOf(wb, "Goals").forEach((r) => {
    const k = str(r[0]);
    if (k === "Daily goal") dailyGoal = toNum(r[1]);
    else if (/^\d{4}-\d{2}-\d{2}$/.test(isoOf(r[0]))) dayGoals[isoOf(r[0])] = toNum(r[1]);
  });
  return {
    quotesRep: quotes.length ? { v: 2, quotes, minD, maxD, fileName: info["Last Quote Report"] || "", at: info["Quote Report imported at"] || "", hasClass: quotes.some((q) => q.cls), hasLine: quotes.some((q) => Object.keys(q.al).length) } : null,
    salesRows,
    salesInfo: { fileName: info["Last Sales Report"] || "", at: info["Sales Report imported at"] || "" },
    ver: info["Program version"] || "",
    ...readUsers(wb),
    meta,
    excl,
    acts,
    dailyGoal,
    dayGoals,
  };
}

/* ===== juntar o arquivo de dados com o que está no navegador =====
   Usado ao conectar/reconectar o arquivo, para nada se perder:
   • cotações: as duas listas juntas; com o mesmo Quote ID, fica a que veio no Quote Report mais recente
     ("Updated"); empatando, a que tem a linha de produto;
   • vendas: as linhas do arquivo + as que só existem no navegador;
   • FUP / % Win: vale o arquivo; o que só existe no navegador entra também;
   • histórico e anotações: os dois juntos, sem repetir;
   • metas: vale o arquivo (o navegador só completa o que faltar). */
/* preferLocal = junção feita antes de gravar (o arquivo mudou por fora durante a sessão): nos campos
   manuais (FUP, % Win, requisição, metas) vale o que está no navegador; o resto se soma. */
/* JSON com as chaves em ordem (para comparar o conteúdo de dois objetos) */
const sortedJson = (v) =>
  v && typeof v === "object" && !Array.isArray(v)
    ? "{" + Object.keys(v).sort().map((k) => JSON.stringify(k) + ":" + sortedJson(v[k])).join(",") + "}"
    : JSON.stringify(v === undefined ? null : v);
const metaKey = (m) => {
  if (!m) return "";
  const { name, ...rest } = m;
  return sortedJson(rest);
};
export const actKey = (a) => [a.t, a.id, a.k, a.txt || "", a.n ?? ""].join("|");
/* Junta duas versões dos campos manuais olhando para a versão "base" (o que este navegador leu ou gravou
   por último no arquivo): o que ESTE navegador mudou desde a base fica; o resto vem do arquivo, que pode
   ter mudanças de outra pessoa. Assim um Save não desfaz o que o colega salvou antes. */
export function merge3(file, local, base) {
  const fm = file.meta || {},
    lm = local.meta || {},
    bm = base.meta || {};
  const meta = {};
  for (const id of new Set([...Object.keys(fm), ...Object.keys(lm), ...Object.keys(bm)])) {
    const mine = metaKey(lm[id]) !== metaKey(bm[id]);
    const v = mine ? lm[id] : fm[id];
    if (v) meta[id] = v;
  }
  const fe = new Set(file.excl || []),
    le = new Set(local.excl || []),
    be = new Set(base.excl || []);
  le.forEach((id) => !be.has(id) && fe.add(id));
  be.forEach((id) => !le.has(id) && fe.delete(id));
  const fg = file.dayGoals || {},
    lg = local.dayGoals || {},
    bg = base.dayGoals || {};
  const dayGoals = {};
  for (const d of new Set([...Object.keys(fg), ...Object.keys(lg), ...Object.keys(bg)])) {
    const v = lg[d] !== bg[d] ? lg[d] : fg[d];
    if (v !== undefined && v !== null) dayGoals[d] = v;
  }
  const dailyGoal = (local.dailyGoal || 0) !== (base.dailyGoal || 0) ? local.dailyGoal || 0 : file.dailyGoal || 0;
  return { meta, excl: [...fe], dayGoals, dailyGoal };
}
export function mergeData(file, local, preferLocal, base) {
  const fq = (file.quotesRep && file.quotesRep.quotes) || [],
    lq = (local.quotesRep && local.quotesRep.quotes) || [];
  const by = new Map(fq.map((q) => [q.id, q]));
  const hasLine = (q) => !!(q && q.al && Object.keys(q.al).length);
  const hasPg = (q) => !!(q && q.pg && Object.keys(q.pg).length);
  let pgFromLocal = 0;
  lq.forEach((q) => {
    const o = by.get(q.id);
    // fica a mais nova; empatando, a que tem a linha de produto dos itens
    if (!o || (q.up || "") > (o.up || "") || ((q.up || "") === (o.up || "") && hasLine(q) && !hasLine(o))) by.set(q.id, !hasPg(q) && hasPg(o) ? { ...q, pg: o.pg } : q);
    // os Product Groups não se perdem: se a versão que ficou não tem e a outra tem, eles vêm junto
    else if (!hasPg(o) && hasPg(q)) {
      by.set(q.id, { ...o, pg: q.pg });
      pgFromLocal++;
    }
  });
  const quotes = [...by.values()];
  let minD = "",
    maxD = "";
  quotes.forEach((q) => {
    if (!minD || q.d < minD) minD = q.d;
    if (!maxD || q.d > maxD) maxD = q.d;
  });
  const rk = (r) => [r.so, r.d, r.cls, r.ln, r.c, r.sp || ""].join("|");
  const rowBase = (r) => [r.so, r.d, r.cls, r.ln, r.c].join("|");
  const fr = file.salesRows || [],
    seen = new Set(fr.map(rk));
  let salesRows = [...fr, ...((local.salesRep && local.salesRep.rows) || []).filter((r) => !seen.has(rk(r)))];
  // a mesma linha guardada por uma versão antiga (sem o vendedor) e pela nova (com o vendedor): fica a nova
  const withSp = new Set(salesRows.filter((r) => r.sp).map(rowBase));
  salesRows = salesRows.filter((r) => r.sp || !withSp.has(rowBase(r)));
  const ak = actKey;
  const fa = file.acts || [],
    aseen = new Set(fa.map(ak));
  const acts = [...fa, ...(local.acts || []).filter((a) => !aseen.has(ak(a)))].sort((a, b) => (a.t < b.t ? -1 : a.t > b.t ? 1 : 0));
  const qi = file.quotesRep || local.quotesRep || {};
  const m3 = base ? merge3(file, { meta: local.meta, excl: local.excl, dayGoals: local.dayGoals, dailyGoal: local.dailyGoal }, base) : null;
  const out = {
    quotesRep: quotes.length ? { ...(local.quotesRep || {}), ...qi, v: 2, quotes, minD, maxD, hasClass: quotes.some((q) => q.cls), hasLine: quotes.some((q) => q.al && Object.keys(q.al).length) } : null,
    salesRows,
    salesInfo: file.salesInfo && file.salesInfo.fileName ? file.salesInfo : { fileName: (local.salesRep && local.salesRep.fileName) || "", at: (local.salesRep && local.salesRep.at) || "" },
    meta: preferLocal ? { ...(file.meta || {}), ...(local.meta || {}) } : { ...(local.meta || {}), ...(file.meta || {}) },
    excl: [...new Set([...(file.excl || []), ...(local.excl || [])])],
    acts,
    dailyGoal: (preferLocal ? local.dailyGoal || file.dailyGoal : file.dailyGoal || local.dailyGoal) || 0,
    dayGoals: preferLocal ? { ...(file.dayGoals || {}), ...(local.dayGoals || {}) } : { ...(local.dayGoals || {}), ...(file.dayGoals || {}) },
    users: mergeUsers(file.users, local.users),
  };
  out.pgFromLocal = pgFromLocal; // cotações em que só este navegador tinha os Product Groups (o arquivo precisa ser regravado)
  return m3 ? { ...out, ...m3 } : out;
}
