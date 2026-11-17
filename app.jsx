/* =====================================================================================
   QUOTE FOLLOW UP — interface (React)
   Baseado na tabela "Customer portfolio" do BRAZIL OEM PLAN (pulz).
   • Quote portfolio: uma linha por COTAÇÃO com QT | FUP | % Win | Thermometer | Quoted | Open |
     Lost | Sold. FUP, % Win e Thermometer são preenchidos à mão, por cotação; o resto vem dos relatórios.
   • Calendar: meta diária, agenda de follow-ups e, ao clicar num dia, as cotações para bater a meta do dia.
   As regras (parsers, cálculos, agenda) ficam em core.js.
   ===================================================================================== */
import React, { useState, useEffect, useMemo, useRef } from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx";
import {
  norm,
  clean,
  money,
  moneyShort,
  fmtDate,
  fmtDay,
  todayISO,
  addDays,
  dowOf,
  daysBetween,
  lastDayOfMonth,
  monthLabel,
  MONTHS_EN,
  addMonths,
  toNum,
  isQuoteWorkbook,
  parseQuoteWorkbook,
  parseSalesWorkbook,
  buildQuoteRows,
  clsOptions,
  clsMatch,
  quoteState,
  periodGroups,
  periodValid,
  periodLabelOf,
  MON_SHORT,
  periodRange,
  FUP,
  AUTO_FILL,
  FUP_KEYS,
  FUP_PICK,
  fupSuggest,
  toBusinessDay,
  addBusinessDays,
  prevBusinessDay,
  holidayOf,
  isOffDay,
  WIN_VALS,
  WIN_META,
  TEMP,
  TEMP_KEYS,
  fupDates,
  fupPending,
  fupDone,
  nextFup,
  scopeQuotes,
  salesForLine,
  salesForSeller,
  sellerLists,
  sellerOptions,
  quoteCancelled,
  sellerDisplay,
  sellerSame,
  mergeQuotes,
  mergeSalesRows,
  buildSalesIndex,
  buildDataWorkbook,
  parseDataWorkbook,
  isDataWorkbook,
  nowStamp,
  fmtStamp,
  ACT_TYPES,
  STATE_LABEL,
  ALL_LINES,
  THERMO_HR_MONTHS,
  mergeData,
  actKey,
  ADMIN_USER,
  hashPassword,
  newSalt,
  fillQuotes,
  CLOSE_REASONS,
  CLOSE_LABEL,
  quoteStateM,
  normWhy,
  repairMeta,
  APP_VERSION,
  openCandidates,
  planDays,
  customerThermo,
  THERMO_WINDOW,
  THERMO_RECENT_WINDOW,
  THERMO_RECENT_MIN,
  THERMO_HR_MIN,
  THERMO_HR_MIN_OTHER,
  THERMO_DAYS_MIN,
  thermoRefDate,
  priorityOn,
  PRIO_WIN,
  PRIO_TEMP,
  bookedRange,
  businessDays,
  dailyTarget,
  fupReportRows,
  spCode,
  sellerFupReport,
  LATE_HOUR,
  mergeUsers,
} from "./core.js";

import {
  Upload,
  FileSpreadsheet,
  Copy,
  ClipboardList,
  Printer,
  CalendarDays,
  Table2,
  Search,
  ChevronRight,
  ChevronLeft,
  ChevronDown,
  Flame,
  Sun,
  Snowflake,
  Wind,
  Target,
  Check,
  X as XIcon,
  Download,
  MoreHorizontal,
  Filter,
  ArrowUp,
  ArrowDown,
  ArrowUpDown,
  Trash2,
  Save,
  RotateCcw,
  CheckCircle2,
  Clock,
  AlertTriangle,
  Info,
  Crosshair,
  Percent,
  DollarSign,
  Wallet,
  TrendingUp,
  XCircle,
  Hourglass,
  ListChecks,
  LayoutGrid,
  List as ListIcon,
  CalendarCheck,
  Database,
  MessageSquare,
  Lock,
  History,
  StickyNote,
  ArrowRight,
  RefreshCw,
  Link2,
  Unlink,
  FolderOpen,
  FilePlus,
  Layers,
  UserRound,
  Users,
  PieChart,
  KeyRound,
  ShieldCheck,
  LogOut,
  MapPin,
  Headphones,
  SlidersHorizontal,
  UserPlus,
  Ban,
} from "lucide-react";

/* ---------- paleta (mesma do Customer portfolio do pulz) ---------- */
const X = {
  bg: "#F4F7F9",
  line: "#E4E9EE",
  ink: "#1D2B36",
  sub: "#6B7A86",
  teal: "#1F9A84",
  tealDk: "#15806B",
  tealSoft: "#DDF1EC",
  green: "#1E7B4F",
  greenSoft: "#E8F6EE",
  amber: "#B7791F",
  amberSoft: "#FDF0D9",
  red: "#EF8A8A",
  redDk: "#D64545",
  redSoft: "#FDE3E3",
  track: "#EDF1F4",
  head: "#F3F6F8",
  violet: "#5B52C9",
  violetSoft: "#ECEAFB",
};
const card = { background: "#fff", border: `1px solid ${X.line}`, borderRadius: 12, boxShadow: "0 1px 3px rgba(16,24,40,.05)" };
const h2 = { fontSize: 17, fontWeight: 700, color: X.ink, letterSpacing: -0.2 };
const inp = { height: 36, border: `1px solid ${X.line}`, borderRadius: 8, padding: "0 10px", fontSize: 12.5, color: X.ink, background: "#fff", outline: "none" };
const btn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 7,
  height: 36,
  padding: "0 14px",
  borderRadius: 9,
  border: `1px solid ${X.line}`,
  background: "#fff",
  color: X.ink,
  fontSize: 13,
  fontWeight: 600,
  cursor: "pointer",
  whiteSpace: "nowrap",
};
const dash = <span style={{ color: "#C7CFD6" }}>—</span>;
const pill = (txt, fg, bg, extra) => (
  <span style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, fontWeight: 600, color: fg, background: bg, padding: "3px 10px", borderRadius: 999, whiteSpace: "nowrap", ...extra }}>{txt}</span>
);
const ST_PILL = { won: ["Won", "#1F8A5B", "#DDF3E8"], lost: ["Lost", X.redDk, X.redSoft], open: ["Open", X.amber, X.amberSoft], cancelled: ["Cancelled", "#64748B", "#EEF2F5"] };

/* ---------- mobile first ----------
   A tela é pensada primeiro para o celular (uma coluna, cartões no lugar das tabelas, abas embaixo) e
   cresce para o computador a partir de 761 px. useMobile() = tela estreita (celular ou janela pequena). */
const MOBILE_Q = "(max-width: 760px)";
function useMedia(q) {
  const get = () => (typeof window !== "undefined" && window.matchMedia ? window.matchMedia(q).matches : false);
  const [m, setM] = useState(get);
  useEffect(() => {
    if (!window.matchMedia) return;
    const mq = window.matchMedia(q);
    const f = () => setM(mq.matches);
    f();
    if (mq.addEventListener) mq.addEventListener("change", f);
    else mq.addListener(f);
    return () => (mq.removeEventListener ? mq.removeEventListener("change", f) : mq.removeListener(f));
  }, [q]);
  return m;
}
const useMobile = () => useMedia(MOBILE_Q);
// estilos globais (mobile first: o padrão é o celular; o @media acrescenta o computador)
const GLOBAL_CSS = `
/* campos com largura 100% contam o padding dentro da largura (senão passam da tela do celular) */
input:not([type="checkbox"]):not([type="radio"]), textarea, select { box-sizing: border-box; }
.qp-shell { padding: 10px 12px 96px; gap: 12px; }
.qp-hscroll { overflow-x: auto; scrollbar-width: none; -webkit-overflow-scrolling: touch; }
.qp-hscroll::-webkit-scrollbar { display: none; }
.qp-card-btn { -webkit-tap-highlight-color: transparent; }
.qp-card-btn:active { background: #F4FBF9 !important; }
@keyframes qpspin { to { transform: rotate(360deg); } }
.qp-spin { animation: qpspin 1s linear infinite; }
@media (min-width: 761px) {
  .qp-shell { padding: 18px 24px 40px; gap: 14px; }
}`;

/* ---------- papéis ----------
   O Admin escolhe o papel de cada usuário (aba Data da gerência). Cada papel tem a sua tela:
   • outside = vendedor externo (coluna Sales Person): Hoje e Minhas cotações, pensado para o celular;
   • inside  = vendedor interno (coluna Quoted By): Carteira e Calendário com o plano do dia;
   • manager = gerência: Análise, FUP por vendedor, meta diária, relatórios e usuários (todos os vendedores).
   Sem papel (usuário novo ou o antigo "seller"), a pessoa entra mas só vê o aviso para pedir o acesso. */
const ROLES = {
  outside: { label: "Outside sales", icon: MapPin, fg: X.violet, bg: X.violetSoft, col: "sp", colTxt: "Sales Person" },
  inside: { label: "Inside sales", icon: Headphones, fg: X.tealDk, bg: X.tealSoft, col: "qb", colTxt: "Quoted By" },
  manager: { label: "Management", icon: ShieldCheck, fg: "#2F6FDB", bg: "#E6EEFB", col: "", colTxt: "" },
};
const ROLE_KEYS = ["outside", "inside", "manager"];
const ROLE_TABS = {
  outside: [
    ["today", "Today", CalendarCheck],
    ["mine", "Quotes", ListIcon],
  ],
  inside: [
    ["today", "Today", CalendarCheck],
    ["portfolio", "Quotes", ListIcon],
    ["calendar", "Calendar", CalendarDays],
  ],
  manager: [
    ["analysis", "Analysis", PieChart],
    ["sellers", "Salespeople", Users],
    ["goal", "Goal", Target],
    ["data", "Data", Database],
  ],
};
const TEMP_ICON = { HOT: Flame, WARM: Sun, COOL: Wind, COLD: Snowflake };

const FupPill = ({ v, short }) => (FUP[v] ? pill(short ? v : `${v} · ${FUP[v].label}`, FUP[v].color, FUP[v].bg, { fontWeight: 700 }) : null);
const WinPill = ({ v }) => (WIN_META[v] ? pill(`${v}%`, WIN_META[v].color, WIN_META[v].bg, { fontWeight: 700, minWidth: 46, justifyContent: "center" }) : null);
const TempPill = ({ v }) => {
  const t = TEMP[v];
  if (!t) return null;
  const Ic = TEMP_ICON[v];
  return pill(
    <>
      <Ic size={12} /> {t.label}
    </>,
    t.color,
    t.bg,
  );
};

/* ---------- armazenamento local ----------
   Campos manuais, metas e preferências: localStorage (pequenos).
   Relatórios importados: IndexedDB, com reserva no localStorage e, por último, na memória
   (o Chrome bloqueia o IndexedDB quando a página é aberta por file://, como já tratado no pulz). */
const K = { users: "qp_users", user: "qp_user", meta: "qp_qmeta", dailyGoal: "qp_dailygoal", dayGoals: "qp_daygoals", excl: "qp_excl", acts: "qp_acts", prefs: "qp_prefs", quotes: "qp_quotes", sales: "qp_sales" };
const ls = {
  get(k, def) {
    try {
      const s = localStorage.getItem(k);
      return s == null ? def : JSON.parse(s);
    } catch {
      return def;
    }
  },
  set(k, v) {
    try {
      localStorage.setItem(k, JSON.stringify(v));
      return true;
    } catch {
      return false;
    }
  },
};
const idb = {
  db: null,
  async open() {
    if (this.db) return this.db;
    this.db = await new Promise((res, rej) => {
      try {
        const r = indexedDB.open("quotepipeline", 1);
        r.onupgradeneeded = () => r.result.createObjectStore("kv");
        r.onsuccess = () => res(r.result);
        r.onerror = () => rej(r.error);
      } catch (e) {
        rej(e);
      }
    });
    return this.db;
  },
  async run(mode, fn) {
    const db = await this.open();
    return new Promise((res, rej) => {
      const tx = db.transaction("kv", mode);
      const req = fn(tx.objectStore("kv"));
      tx.oncomplete = () => res(req && req.result);
      tx.onerror = () => rej(tx.error);
    });
  },
};
const mem = {};
const kv = {
  async get(k) {
    try {
      const v = await idb.run("readonly", (s) => s.get(k));
      if (v !== undefined) return v;
    } catch {}
    const l = ls.get(k, undefined);
    return l !== undefined ? l : mem[k];
  },
  async set(k, v) {
    try {
      await idb.run("readwrite", (s) => s.put(v, k));
      try {
        localStorage.removeItem(k);
      } catch {}
      return "idb";
    } catch {}
    if (ls.set(k, v)) return "ls";
    mem[k] = v;
    return "mem";
  },
  async del(k) {
    try {
      await idb.run("readwrite", (s) => s.delete(k));
    } catch {}
    try {
      localStorage.removeItem(k);
    } catch {}
    delete mem[k];
  },
};
function useStored(key, init) {
  const [v, setV] = useState(() => ls.get(key, init));
  useEffect(() => {
    ls.set(key, v);
  }, [key, v]);
  return [v, setV];
}
const downloadBlob = (blob, name) => {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(a.href);
    a.remove();
  }, 500);
};

/* ---------- menu flutuante (portal) ----------
   Fica preso à tela (position fixed) e fecha ao clicar fora, rolar ou redimensionar. Os cliques dentro
   dele param a propagação: o React propaga eventos do portal para o componente pai (a linha da tabela). */
function usePop(width = 200, maxH = 340) {
  const [pos, setPos] = useState(null);
  const anchor = useRef(null),
    menu = useRef(null);
  useEffect(() => {
    if (!pos) return;
    const close = (e) => {
      const t = e && e.target;
      if (t && ((menu.current && menu.current.contains(t)) || (anchor.current && anchor.current.contains(t)))) return;
      // a página rola sozinha quando o teclado do celular abre para a busca do menu: não fecha
      if (e && e.type === "scroll" && menu.current && menu.current.contains(document.activeElement)) return;
      setPos(null);
    };
    // no celular o teclado que abre muda só a altura da janela: isso não fecha o menu
    const w0 = window.innerWidth;
    const onResize = () => window.innerWidth !== w0 && setPos(null);
    window.addEventListener("mousedown", close);
    window.addEventListener("touchstart", close, { passive: true });
    window.addEventListener("resize", onResize);
    window.addEventListener("scroll", close, true);
    return () => {
      window.removeEventListener("mousedown", close);
      window.removeEventListener("touchstart", close);
      window.removeEventListener("resize", onResize);
      window.removeEventListener("scroll", close, true);
    };
  }, [pos]);
  const toggle = (e) => {
    e.stopPropagation();
    if (pos) return setPos(null);
    const rc = anchor.current.getBoundingClientRect();
    // no celular o menu nunca passa da largura da tela
    const w = Math.min(width, window.innerWidth - 16);
    const x = Math.max(8, Math.min(rc.left, window.innerWidth - w - 8));
    const up = rc.bottom + Math.min(280, maxH) > window.innerHeight && rc.top > 300;
    const h = Math.min(maxH, Math.max(160, (up ? rc.top : window.innerHeight - rc.bottom) - 16));
    setPos(up ? { x, w, h, bottom: window.innerHeight - rc.top + 6 } : { x, w, h, top: rc.bottom + 6 });
  };
  const render = (children) =>
    pos &&
    createPortal(
      <div
        ref={menu}
        onClick={(e) => e.stopPropagation()}
        style={{ position: "fixed", left: pos.x, top: pos.top, bottom: pos.bottom, width: pos.w || width, zIndex: 2100, ...card, borderRadius: 10, padding: 5, boxShadow: "0 10px 30px rgba(16,24,40,.16)", maxHeight: pos.h || maxH, overflow: "auto" }}
      >
        {children}
      </div>,
      document.body,
    );
  return { pos, setPos, anchor, toggle, render };
}
const optS = { display: "flex", alignItems: "center", gap: 8, padding: "7px 9px", borderRadius: 7, cursor: "pointer", fontSize: 12.5, color: X.ink };

/* Célula editável (FUP, % Win, Thermometer): mostra a pílula, clique abre as opções. */
function PickCell({ value, options, onChange, render, placeholder = "Set", title, width = 210, describe }) {
  const p = usePop(width);
  return (
    <>
      <button
        ref={p.anchor}
        onClick={p.toggle}
        title={title}
        style={{ border: "none", background: "transparent", padding: 0, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 3 }}
      >
        {value ? render(value) : <span style={{ fontSize: 11.5, fontWeight: 600, color: "#94A3B8", border: "1px dashed #CBD5E1", padding: "2px 9px", borderRadius: 999 }}>{placeholder}</span>}
        <ChevronDown size={12} color="#A0AEC0" />
      </button>
      {p.render(
        <>
          {options.map((o) => (
            <div
              key={o}
              className="qp-opt"
              onClick={() => {
                onChange(o);
                p.setPos(null);
              }}
              style={optS}
            >
              {render(o)}
              {describe && <span style={{ fontSize: 11.5, color: X.sub }}>{describe(o)}</span>}
              {o === value && <Check size={14} color={X.tealDk} style={{ marginLeft: "auto" }} />}
            </div>
          ))}
          {value ? (
            <div
              className="qp-opt"
              onClick={() => {
                onChange("");
                p.setPos(null);
              }}
              style={{ ...optS, color: X.sub, borderTop: `1px solid ${X.line}`, marginTop: 4, borderRadius: 0 }}
            >
              <XIcon size={13} /> Clear
            </div>
          ) : null}
        </>,
      )}
    </>
  );
}
const NONE = "__none__";
// FUP: as letras A–D e, no fim, os fechamentos à mão (abrem o diálogo do motivo)
const FUP_OPTS = [...FUP_PICK, "lost", "cancelled"];
const fupOptRender = (v) => (CLOSE_LABEL[v] ? <span style={{ fontSize: 11.5, fontWeight: 700, color: CLOSE_COLOR[v][0], background: CLOSE_COLOR[v][1], padding: "2px 9px", borderRadius: 999 }}>{CLOSE_LABEL[v]}</span> : <FupPill v={v} />);
const fupOptDescribe = (v) => (CLOSE_LABEL[v] ? `${CLOSE_REASONS[v].join(", ")}` : FUP[v].long.split("· ")[1]);
/* Thermometer automático: pílula + balão com a conta (Hit Rate em valor e dias com cotação). */
function ThermoCell({ t }) {
  if (!t) return dash;
  const pct = (t.hr * 100).toFixed(0) + "%";
  const down = t.temp !== t.base; // desceu um nível pela falta de atividade recente
  const tip =
    `Automatic · customer hit rate ${pct} in the last ${THERMO_HR_MONTHS} months (${money(t.sold)} sold of ${money(t.quoted)} quoted since ${fmtDate(t.hrFrom)})` +
    ` · minimum for Hot/Warm: ${Math.round((t.hrMin ?? THERMO_HR_MIN) * 100)}%${t.cls ? ` (${t.cls} customer)` : ""}` +
    ` · quoted on ${t.days} day${t.days === 1 ? "" : "s"} in the last ${THERMO_WINDOW} business days (${fmtDate(t.from)} to ${fmtDate(t.ref)})` +
    ` · ${t.recent} day${t.recent === 1 ? "" : "s"} in the last ${THERMO_RECENT_WINDOW} business days (since ${fmtDate(t.recentFrom)})` +
    (down ? ` · was ${TEMP[t.base].label}, dropped one level: fewer than ${THERMO_RECENT_MIN} recent days` : "") +
    (t.last ? ` · last quote ${fmtDate(t.last)}` : "");
  return (
    <span title={tip} style={{ display: "inline-flex", alignItems: "center", gap: 6, cursor: "help", whiteSpace: "nowrap" }}>
      <TempPill v={t.temp} />
      <span style={{ fontSize: 11, color: X.sub, fontVariantNumeric: "tabular-nums" }}>
        {pct} · {t.days}d
      </span>
      {down && <ArrowDown size={12} color={X.redDk} />}
    </span>
  );
}
/* Filtro de coluna com várias escolhas. */
/* Filtro de coluna: o próprio título da coluna abre a lista (várias escolhas). Com muitas opções
   (clientes), aparece uma busca no topo. O número ao lado do funil = quantas opções estão marcadas. */
/* row = uma linha inteira (nome à esquerda, "All" / "2 selected" à direita), usada na folha de filtros do celular */
function ColFilter({ label, options, selected, onChange, width = 240, iconOnly, row }) {
  const p = usePop(width);
  const [q, setQ] = useState("");
  const on = selected.length > 0;
  const toggle = (v) => onChange(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]);
  const ql = norm(q);
  const list = ql ? options.filter((o) => norm(o.text || o.v).includes(ql)) : options;
  return (
    <>
      <span
        ref={p.anchor}
        onClick={(e) => {
          setQ("");
          p.toggle(e);
        }}
        title={`Filter by ${label}`}
        style={
          row
            ? { display: "flex", alignItems: "center", gap: 8, width: "100%", padding: "13px 2px", cursor: "pointer", borderBottom: `1px solid ${X.line}` }
            : { display: "inline-flex", alignItems: "center", gap: 5, cursor: "pointer", padding: "3px 6px", margin: "-3px -6px", borderRadius: 6, background: on ? X.tealSoft : "transparent" }
        }
      >
        {row ? (
          <>
            <span style={{ flex: 1, fontSize: 14, fontWeight: 600, color: X.ink }}>{label}</span>
            <span style={{ fontSize: 13, color: on ? X.tealDk : X.sub, fontWeight: on ? 700 : 500 }}>{on ? `${selected.length} selected` : "All"}</span>
            <ChevronRight size={16} color="#A0AEC0" />
          </>
        ) : (
          <>
            {!iconOnly && label}
            <Filter size={11.5} color={on ? X.tealDk : "#9AA5B1"} />
            {on && <span style={{ fontSize: 10, fontWeight: 800, color: "#fff", background: X.teal, borderRadius: 999, padding: "0 5px", lineHeight: "15px" }}>{selected.length}</span>}
          </>
        )}
      </span>
      {p.render(
        <>
          {options.length > 8 && (
            <div style={{ position: "relative", padding: "2px 2px 6px" }}>
              <Search size={13} style={{ position: "absolute", left: 10, top: 11, color: X.sub, pointerEvents: "none" }} />
              <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder={`Search ${label.toLowerCase()}…`} style={{ ...inp, height: 30, width: "100%", paddingLeft: 28, fontSize: 12 }} />
            </div>
          )}
          {list.map((o) => (
            <label key={o.v} className="qp-opt" style={{ ...optS, textTransform: "none", letterSpacing: 0, fontWeight: 500 }}>
              <input type="checkbox" checked={selected.includes(o.v)} onChange={() => toggle(o.v)} style={{ accentColor: X.teal, flexShrink: 0 }} />
              <span style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={o.text}>
                {o.node}
              </span>
              <span style={{ marginLeft: "auto", fontSize: 11, color: X.sub, flexShrink: 0 }}>{o.n}</span>
            </label>
          ))}
          {!list.length && <div style={{ ...optS, color: X.sub, cursor: "default" }}>Nothing found</div>}
          {on && (
            <div className="qp-opt" onClick={() => onChange([])} style={{ ...optS, color: X.sub, borderTop: `1px solid ${X.line}`, marginTop: 4, borderRadius: 0 }}>
              <XIcon size={13} /> Clear filter
            </div>
          )}
        </>,
      )}
    </>
  );
}

/* ===== folha (sheet) =====
   Tudo o que não precisa ficar à vista (filtros, ordem, exportar) fica numa folha que sobe de baixo no celular
   (no computador, uma janela no centro). A tela fica só com a lista e a ação principal. */
function Sheet({ open, title, onClose, children, footer }) {
  const mob = useMobile();
  useEffect(() => {
    if (!open) return;
    const k = (e) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", k);
    const ov = document.body.style.overflow;
    document.body.style.overflow = "hidden"; // a página de trás não rola junto
    return () => {
      window.removeEventListener("keydown", k);
      document.body.style.overflow = ov;
    };
  }, [open, onClose]);
  if (!open) return null;
  return createPortal(
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 1600, background: "rgba(15,23,42,.38)", display: "flex", alignItems: mob ? "flex-end" : "center", justifyContent: "center" }}>
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: "#fff", width: mob ? "100%" : "min(480px, 94vw)", maxHeight: "88vh", borderRadius: mob ? "20px 20px 0 0" : 16, display: "flex", flexDirection: "column", boxShadow: "0 -10px 40px rgba(15,23,42,.25)", paddingBottom: mob ? "env(safe-area-inset-bottom)" : 0 }}
      >
        {mob && <div style={{ width: 40, height: 4, borderRadius: 9, background: "#D5DCE3", margin: "8px auto 0", flexShrink: 0 }} />}
        <div style={{ display: "flex", alignItems: "center", padding: "10px 16px 6px", flexShrink: 0 }}>
          <div style={{ fontSize: 17, fontWeight: 800, flex: 1 }}>{title}</div>
          <button onClick={onClose} title="Close" style={{ ...btn, width: 36, height: 36, padding: 0, justifyContent: "center", borderRadius: 999, border: "none", background: X.head }}>
            <XIcon size={17} />
          </button>
        </div>
        <div style={{ overflowY: "auto", padding: "6px 16px 18px", display: "flex", flexDirection: "column", gap: 18 }}>{children}</div>
        {footer && <div style={{ padding: "10px 16px 14px", borderTop: `1px solid ${X.line}`, display: "flex", gap: 8, flexShrink: 0 }}>{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}
function SheetField({ label, children }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div style={{ fontSize: 11.5, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</div>
      {children}
    </div>
  );
}
/* escolha única em botões (no celular é mais rápido que um select) */
function Chips({ value, options, onChange }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
      {options.map(([v, l]) => {
        const on = value === v;
        return (
          <button
            key={v || "all"}
            onClick={() => onChange(v)}
            style={{ border: `1px solid ${on ? X.teal : X.line}`, background: on ? X.teal : "#fff", color: on ? "#fff" : X.ink, borderRadius: 999, padding: "9px 14px", fontSize: 13, fontWeight: on ? 700 : 500, cursor: "pointer" }}
          >
            {l}
          </button>
        );
      })}
    </div>
  );
}
/* botão redondo de ícone (topo do celular) */
const iconBtn = (extra) => ({ ...btn, width: 40, height: 40, padding: 0, justifyContent: "center", borderRadius: 12, position: "relative", flexShrink: 0, ...extra });
const dotBadge = (n) =>
  n > 0 ? (
    <span style={{ position: "absolute", top: -4, right: -4, minWidth: 18, height: 18, padding: "0 5px", borderRadius: 999, background: X.teal, color: "#fff", fontSize: 10.5, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center", border: "2px solid #fff" }}>{n}</span>
  ) : null;

function Tile({ icon: Ic, label, value, caption, color = X.ink, tint = X.tealSoft, ink = X.tealDk, title }) {
  const mob = useMobile(); // celular: sem o ícone, para sobrar espaço para o número
  return (
    <div title={title} style={{ ...card, padding: mob ? "10px 11px" : "13px 15px", display: "flex", gap: 12, alignItems: "center", minWidth: 0, ...(mob ? { borderTop: `3px solid ${tint}` } : null) }}>
      {!mob && (
        <span style={{ width: 36, height: 36, borderRadius: 10, background: tint, color: ink, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Ic size={18} />
        </span>
      )}
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.35, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{label}</div>
        <div style={{ fontSize: 18, fontWeight: 800, color, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap", letterSpacing: -0.3, overflow: "hidden", textOverflow: "ellipsis" }}>{value}</div>
        {caption && <div style={{ fontSize: 11.5, color: X.sub, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{caption}</div>}
      </div>
    </div>
  );
}

const th = {
  position: "sticky",
  top: 0,
  zIndex: 2,
  padding: "10px 8px",
  background: X.head,
  boxShadow: `inset 0 -1px 0 ${X.line}`,
  borderTop: `1px solid ${X.line}`,
  color: "#4B5563",
  fontSize: 10.5,
  fontWeight: 700,
  textTransform: "uppercase",
  letterSpacing: 0.3,
  whiteSpace: "nowrap",
  textAlign: "left",
  userSelect: "none",
};
const td = { padding: "10px 8px", borderTop: `1px solid ${X.line}`, color: X.ink, fontSize: 12.5, verticalAlign: "middle" };
const tdN = { ...td, textAlign: "right", fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" };

/* ===== registro da cotação: detalhes + Activities History + anotações =====
   Aparece na seta da linha (Quote portfolio) e na gaveta lateral (plano do dia).
   As anotações entram no histórico com data e hora e NÃO têm botão de editar nem de excluir. */
/* Cada tipo de registro: ícone (círculo), título e etiqueta, no estilo do Activities / History do pulz. */
const ACT_TAG = {
  note: ["Note", X.tealDk, X.tealSoft],
  follow: ["Follow-up", X.violet, X.violetSoft],
  plan: ["Plan", X.amber, X.amberSoft],
  report: ["Quote Report", "#2F6FDB", "#E6EEFB"],
  import: ["Import", "#64748B", "#EEF2F5"],
};
const ACT_META = {
  note: { icon: MessageSquare, title: "Note", tag: "note", color: X.tealDk, bg: X.tealSoft },
  fup: { icon: CalendarDays, title: "FUP changed", tag: "follow", color: X.violet, bg: X.violetSoft },
  win: { icon: Percent, title: "% Win changed", tag: "follow", color: X.violet, bg: X.violetSoft },
  done: { icon: CheckCircle2, title: "Follow-up done", tag: "follow", color: X.violet, bg: X.violetSoft },
  resched: { icon: CalendarDays, title: "Next follow-up changed", tag: "follow", color: X.violet, bg: X.violetSoft },
  req: { icon: ClipboardList, title: "Customer requisition", tag: "follow", color: X.violet, bg: X.violetSoft },
  plan: { icon: ListChecks, title: "Daily plan", tag: "plan", color: X.amber, bg: X.amberSoft },
  close: { icon: XCircle, title: "Closed by hand", tag: "follow", color: X.redDk, bg: X.redSoft },
  reopen: { icon: RotateCcw, title: "Reopened", tag: "follow", color: X.tealDk, bg: X.tealSoft },
  why: { icon: XCircle, title: "Close reason", tag: "follow", color: "#64748B", bg: "#EEF2F5" },
  status: { icon: FileSpreadsheet, title: "Status changed", tag: "report", color: "#2F6FDB", bg: "#E6EEFB" },
  amount: { icon: FileSpreadsheet, title: "Quoted value changed", tag: "report", color: "#2F6FDB", bg: "#E6EEFB" },
  order: { icon: FileSpreadsheet, title: "New order", tag: "report", color: "#2F6FDB", bg: "#E6EEFB" },
  _seen: { icon: Database, title: "Entered the program", tag: "import", color: "#64748B", bg: "#EEF2F5" },
  _created: { icon: FileSpreadsheet, title: "Quote created", tag: "report", color: "#64748B", bg: "#EEF2F5" },
};
const fupTxt = (v) => (FUP[v] ? `${v} · ${FUP[v].label}` : "not set");
// grupos do filtro do histórico
const HIST_FILTERS = [
  ["all", "All"],
  ["note", "Notes"],
  ["follow", "Follow-up"],
  ["report", "Quote Report"],
];
const histGroup = (a) => (a.k === "note" ? "note" : ["fup", "win", "done", "resched", "req", "plan", "close", "reopen", "why"].includes(a.k) ? "follow" : "report");
// "Today", "Yesterday" ou o dia da semana com a data
const dayLabel = (d) => {
  const t = todayISO();
  if (d === t) return "Today";
  if (d === addDays(t, -1)) return "Yesterday";
  return `${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][dowOf(d)]}, ${fmtDate(d)}`;
};
const winTxt = (v) => (v ? `${v}%` : "not set");
const arrow = <ArrowRight size={12} style={{ verticalAlign: -2, margin: "0 3px" }} />;
// texto do registro (o título e a etiqueta vêm de ACT_META)
function actText(a) {
  switch (a.k) {
    case "fup":
      return <>{fupTxt(a.o)} {arrow} <b>{fupTxt(a.n)}</b></>;
    case "win":
      return <>{winTxt(a.o)} {arrow} <b>{winTxt(a.n)}</b></>;
    case "done":
      return (
        <>
          {a.o ? (
            <>
              {fupTxt(a.o)} {arrow} <b>{fupTxt("E")}</b> ·{" "}
            </>
          ) : null}
          {a.n ? (
            <>
              Next follow-up <b>{fmtDate(a.n)}</b>
            </>
          ) : (
            "Follow-up made"
          )}
        </>
      );
    case "req":
      return a.o ? (
        <>
          {a.o} {arrow} <b>{a.n || "removed"}</b>
        </>
      ) : (
        <b>{a.n}</b>
      );
    case "resched":
      return (
        <>
          {a.o ? fmtDate(a.o) : "no date"} {arrow} <b>{fmtDate(a.n)}</b> · {fupTxt("E")}
        </>
      );
    case "plan":
      return a.n === "out" ? "Left out of the daily plan" : "Put back in the daily plan";
    case "close":
      return (
        <>
          Marked as <b>{CLOSE_LABEL[a.n] || a.n}</b>
          {a.txt ? <> · reason <b>{a.txt}</b></> : null}
        </>
      );
    case "reopen":
      return <>Back to <b>Open</b> (the close by hand was undone)</>;
    case "why":
      return a.o ? (
        <>
          {a.o} {arrow} <b>{a.n || "removed"}</b>
        </>
      ) : (
        <b>{a.n}</b>
      );
    case "status":
      return <>{a.o} {arrow} <b>{a.n}</b></>;
    case "amount":
      return <>{money(a.o)} {arrow} <b>{money(a.n)}</b></>;
    case "order":
      return <b>{a.n}</b>;
    case "_seen":
      return "Imported from the Quote Report";
    case "_created":
      return "Create Date in the Quote Report";
    default:
      return a.txt || a.k;
  }
}
const recBox = { background: "#fff", border: `1px solid ${X.line}`, borderRadius: 12, padding: "10px 14px" };
const recHead = { display: "flex", alignItems: "center", gap: 7, fontSize: 11, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.45, paddingBottom: 6, marginBottom: 2, borderBottom: `1px solid ${X.line}` };
const recRow = { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "7px 0", borderBottom: "1px dashed #E9EEF3", minHeight: 36 };
const recKey = { fontSize: 12, color: X.sub };
const recLbl = { fontSize: 10.5, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.45 };
const LINE_COLOR = { PT: ["#2F6FDB", "#E6EEFB"], CP: ["#B7791F", "#FDF0D9"], MH: ["#7A6AD8", "#EEEAFB"] };
const initials = (n) =>
  String(n || "")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w[0])
    .join("")
    .toUpperCase();
/* título de cada passo obrigatório do Done: número, nome e selo Required / OK */
function StepHead({ n, title, sub, ok }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
      <span style={{ width: 22, height: 22, borderRadius: 999, background: ok ? X.teal : "#F0A43A", color: "#fff", fontSize: 12, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        {ok ? <Check size={13} strokeWidth={3} /> : n}
      </span>
      <span style={{ fontSize: 13, fontWeight: 700, color: X.ink }}>{title}</span>
      {sub && <span style={{ fontSize: 11.5, color: X.sub }}>{sub}</span>}
      <span
        style={{
          marginLeft: "auto",
          fontSize: 10.5,
          fontWeight: 800,
          textTransform: "uppercase",
          letterSpacing: 0.3,
          padding: "2px 8px",
          borderRadius: 999,
          color: ok ? X.green : "#B45309",
          background: ok ? X.greenSoft : "#FEF3C7",
        }}
      >
        {ok ? "OK" : "Required"}
      </span>
    </div>
  );
}
function QuoteRecord({ q, sales, acts, thermo, onNote, onDone, onResched, onReq, onAskClose, onReopen, onWhy, focusDone, compact, m, user, onUser, headRight }) {
  const mob = useMobile(); // celular: tudo numa coluna só, com uma rolagem para a tela inteira
  const [txt, setTxt] = useState("");
  // requisição do cliente (ex.: "RODRIGO - POLIAS"): campo da cotação, vai para o relatório de FUP
  const curReq = (m && m.req) || "";
  const [req, setReq] = useState(curReq);
  useEffect(() => setReq(curReq), [q && q.id, curReq]);
  const [nd, setNd] = useState(""); // data do próximo follow-up escolhida (começa vazia: tem que ser marcada)
  // % Win revisado no Done (0 = ainda não olhou): tem que confirmar ou trocar a cada Done
  const [wv, setWv] = useState(0);
  useEffect(() => setWv(0), [q && q.id]);
  const noteRef = useRef(null);
  // aberto pelo botão Done do plano ou da semana: vai direto para a nota + data
  useEffect(() => {
    if (focusDone && noteRef.current) {
      noteRef.current.scrollIntoView({ block: "center" });
      noteRef.current.focus();
    }
  }, [focusDone]);
  const [hf, setHf] = useState("all"); // filtro do histórico: all | note | follow | report
  const [showDet, setShowDet] = useState(false); // celular: "Details" recolhido
  const [showHist, setShowHist] = useState(false); // celular: "History" recolhido
  if (!q) return <div style={{ padding: 16, color: X.sub }}>Quote not found in the stored data.</div>;
  // situação considerando o fechamento à mão (FUP → Cancelled / Lost)
  const st = quoteStateM(q, (m || {}).close);
  const byHand = st !== quoteState(q);
  const s = ST_PILL[st] || ST_PILL.open;
  const soMap = (sales && sales.so) || {};
  // mais novo em cima; no mesmo segundo, o que foi feito por último fica em cima
  const list = [
    { t: q.d + " 00:00:00", k: "_created", dateOnly: true },
    ...(q.fs ? [{ t: q.fs, k: "_seen" }] : []),
    ...(acts || []).filter((a) => a.id === q.id),
  ]
    .map((a, i) => ({ ...a, _i: i }))
    .sort((a, b) => (a.t < b.t ? 1 : a.t > b.t ? -1 : b._i - a._i));
  const shownH = hf === "all" ? list : list.filter((a) => histGroup(a) === hf);
  // agrupa por dia, do mais novo para o mais antigo
  const dayGroups = [];
  shownH.forEach((a) => {
    const d = a.t.slice(0, 10);
    const g = dayGroups.length && dayGroups[dayGroups.length - 1][0] === d ? dayGroups[dayGroups.length - 1] : null;
    if (g) g[1].push(a);
    else dayGroups.push([d, [a]]);
  });
  const add = () => {
    const t = txt.trim();
    if (!t) return;
    onNote(q.id, q.c, t);
    setTxt("");
  };
  const lines = ["PT", "CP", "MH"].filter((k) => q.al && q.al[k]);
  const mm = m || {};
  const today = todayISO();
  const age = Math.max(0, daysBetween(q.d, today));
  const next = st === "open" && mm.fup ? nextFup(mm) : "";
  const nextIn = next ? daysBetween(today, next) : 0;
  const nextLate = !!next && nextIn < 0;
  // nota + próximo follow-up (Done / Change date)
  const open = st === "open" && !!(onDone && onResched);
  const ndFix = /^\d{4}-\d{2}-\d{2}$/.test(nd) ? toBusinessDay(nd) : "";
  const ndOk = !!ndFix && ndFix > today;
  const curWin = mm.win || 0;
  const winOk = wv > 0;
  const canDone = open && ndOk && !!txt.trim() && winOk;
  const canResched = open && ndOk && ndFix !== next;
  const quick = [
    ["Tomorrow", addBusinessDays(today, 1)],
    ["+2 days", addBusinessDays(today, 2)],
    ["+1 week", toBusinessDay(addDays(today, 7))],
    ["+2 weeks", toBusinessDay(addDays(today, 14))],
    ["+1 month", toBusinessDay(addMonths(today, 1))],
  ];
  const done = () => {
    if (!canDone) return;
    onDone(q.id, q.c, ndFix, txt.trim(), wv);
    setTxt("");
    setNd("");
    setWv(0);
  };
  const resched = () => {
    if (!canResched) return;
    onResched(q.id, q.c, ndFix, txt.trim());
    setTxt("");
    setNd("");
  };
  const hasNote = !!txt.trim();
  // caixa de cada passo: âmbar enquanto falta, verde quando preenchido
  const stepBox = (ok) => ({
    display: "flex",
    flexDirection: "column",
    gap: 8,
    padding: open ? "10px 11px" : 0,
    borderRadius: 10,
    border: open ? `1px solid ${ok ? "#BFE3D9" : "#F5D9B5"}` : "none",
    borderLeft: open ? `4px solid ${ok ? X.teal : "#F0A43A"}` : "none",
    background: open ? (ok ? "#F6FBFA" : "#FFFBF5") : "transparent",
  });
  const subBtn = (on) => ({ ...btn, flex: mob ? "1 1 0" : "1 1 160px", minWidth: 0, justifyContent: "center", height: mob ? 38 : 32, fontSize: 12, color: on ? X.ink : "#A0AEC0", cursor: on ? "pointer" : "not-allowed", background: "#fff" });
  // vendido: valor Booked dos pedidos no Sales Report (sem eles, o Sales Amt da cotação)
  const soldV = st === "won" ? ((q.so || []).some((id) => soMap[id]) ? (q.so || []).reduce((t, id) => t + ((soMap[id] && soMap[id].v) || 0), 0) : q.sa || q.a) : null;
  // vendedores: interno (Quoted By) e externo (Sales Person); a mesma pessoa nas duas colunas aparece uma vez
  const people = [];
  if (clean(q.qb)) people.push({ name: sellerDisplay(q.qb), raw: q.qb, role: "Inside · Quoted By", fg: X.tealDk, bg: X.tealSoft });
  if (clean(q.sp)) {
    if (people.length && sellerSame(q.qb, q.sp)) people[0].role = "Inside and outside · Quoted By and Sales Person";
    else people.push({ name: sellerDisplay(q.sp), raw: q.sp, role: "Outside · Sales Person", fg: X.violet, bg: X.violetSoft });
  }
  /* ---------- painel da cotação: cabeçalho, faixa de dados, e embaixo o registro do follow-up (esquerda)
     ao lado do Activities / History (direita), cada lado com a própria rolagem ---------- */
  const factLbl = { fontSize: 10.5, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.45, whiteSpace: "nowrap" };
  const fact = (label, body, sub, extra) => (
    <div style={{ padding: mob ? "10px 14px" : "11px 16px", borderRight: `1px solid ${X.line}`, ...(mob ? { borderBottom: `1px solid ${X.line}` } : null), minWidth: 0, display: "flex", flexDirection: "column", gap: 6, ...extra, ...(mob && extra && extra.paddingLeft ? { paddingLeft: 14 } : null) }}>
      <div style={factLbl}>{label}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0, flexWrap: "wrap", fontSize: 13 }}>{body}</div>
      {sub ? <div style={{ fontSize: 11, color: X.sub, marginTop: -2 }}>{sub}</div> : null}
    </div>
  );
  const muted = (t) => <span style={{ color: "#A0AEC0", fontSize: 12.5 }}>{t}</span>;
  // passo do registro: círculo com o número (âmbar) ou o check (verde)
  const step = (n, ok, title, hint) => (
    <div style={{ display: "flex", alignItems: "center", gap: 9, minWidth: 0 }}>
      <span style={{ width: 22, height: 22, borderRadius: 999, flexShrink: 0, background: ok ? X.teal : "#FDF0D9", color: ok ? "#fff" : "#B45309", border: ok ? "none" : "1px solid #F2C48D", fontSize: 11.5, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
        {ok ? <Check size={13} strokeWidth={3} /> : n}
      </span>
      <span style={{ fontSize: 13.5, fontWeight: 700, color: X.ink, whiteSpace: "nowrap" }}>{title}</span>
      {hint ? <span style={{ fontSize: 12, color: X.sub, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: mob ? "normal" : "nowrap" }}>{hint}</span> : null}
    </div>
  );
  const nDone = [hasNote, ndOk, winOk].filter(Boolean).length;
  const sect = { padding: mob ? "12px 12px" : "14px 18px", borderTop: `1px solid ${X.line}`, display: "flex", flexDirection: "column", gap: 10 };
  const panelCard = { background: "#fff", border: `1px solid ${X.line}`, borderRadius: 14, boxShadow: "0 1px 3px rgba(16,24,40,.05)" };
  const noteBox = (
    <textarea
      ref={noteRef}
      value={txt}
      onChange={(e) => setTxt(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) open ? null : add();
      }}
      placeholder={open ? "e.g. Talked to the buyer, he asked for a 5% discount. Call again after the meeting…" : "Write a note about this quote… (Ctrl+Enter saves)"}
      rows={4}
      style={{
        ...inp,
        height: "auto",
        width: "100%",
        padding: "10px 12px",
        resize: "vertical",
        fontSize: 13,
        lineHeight: 1.5,
        fontFamily: "inherit",
        borderRadius: 10,
        borderColor: open && !hasNote ? "#F2C48D" : hasNote ? "#9FD6C9" : X.line,
        background: open && !hasNote ? "#FFFCF7" : "#fff",
      }}
    />
  );
  const factsStrip = (
    <div style={{ display: "grid", gridTemplateColumns: mob ? "repeat(2, minmax(0, 1fr))" : "repeat(auto-fit, minmax(150px, 1fr))", background: "#fff", borderBottom: `1px solid ${X.line}`, flexShrink: 0 }}>
      {fact("FUP", mm.fup ? <FupPill v={mm.fup} /> : muted("Not set"), null, { paddingLeft: 22 })}
      {fact(
        "Next follow-up",
        next ? <b style={{ color: nextLate ? X.redDk : X.ink }}>{fmtDate(next)}</b> : muted(st === "open" ? "—" : "closed"),
        next ? <span style={{ color: nextLate ? X.redDk : nextIn === 0 ? X.amber : X.sub, fontWeight: nextLate ? 700 : 500 }}>{nextIn === 0 ? "today" : nextLate ? `${-nextIn} day${nextIn === -1 ? "" : "s"} late` : `in ${nextIn} day${nextIn === 1 ? "" : "s"}`}</span> : null,
        nextLate ? { background: "#FFF7F7" } : null,
      )}
      {fact("% Win", mm.win ? <WinPill v={mm.win} /> : muted("Not set"))}
      {fact("Thermometer", thermo ? <ThermoCell t={thermo} /> : muted("—"))}
      {fact(
        "Salespeople",
        people.length ? (
          people.map((p, i) => (
            <span key={i} title={`${p.raw} · ${p.role}`} style={{ display: "inline-flex", alignItems: "center", gap: 6, minWidth: 0, maxWidth: "100%" }}>
              <span style={{ width: 24, height: 24, borderRadius: 999, background: p.bg, color: p.fg, fontSize: 10, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{initials(p.name)}</span>
              <span style={{ fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</span>
            </span>
          ))
        ) : (
          muted("—")
        ),
        people.length === 1 ? people[0].role.split(" · ")[0] : null,
        { gridColumn: "span 2" },
      )}
      {fact(
        "Orders (SO)",
        q.so && q.so.length ? (
          q.so.map((id) => (
            <span key={id} title={soMap[id] ? `booked ${fmtDate(soMap[id].d)} · ${money(soMap[id].v)}` : "not in the Sales Report yet"} style={{ fontWeight: 700, fontSize: 12.5 }}>
              {id}
            </span>
          ))
        ) : (
          muted(st === "won" ? "no Sales Id" : "No order yet")
        ),
        q.so && q.so.length ? (q.so.some((id) => soMap[id]) ? <span style={{ color: X.green, fontWeight: 700 }}>booked {moneyShort(q.so.reduce((t, id) => t + ((soMap[id] && soMap[id].v) || 0), 0))}</span> : "not in the Sales Report yet") : null,
        { borderRight: "none" },
      )}
    </div>
  );
  const reqBox = onReq && (
    <div style={{ ...(mob ? { padding: "12px 14px", borderTop: `1px solid ${X.line}` } : { ...panelCard, padding: "12px 16px" }), display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
      <ClipboardList size={16} color={X.violet} />
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: X.ink }}>Customer requisition</div>
        <div style={{ fontSize: 11, color: X.sub }}>who asked and what · goes to the FUP report</div>
      </div>
      <input
        value={req}
        onChange={(e) => setReq(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && req.trim() !== curReq && onReq(q.id, q.c, req.trim())}
        placeholder="e.g. RODRIGO - POLIAS"
        style={{ ...inp, height: 32, fontSize: 12.5, flex: "1 1 180px", minWidth: 0, borderRadius: 9 }}
      />
      {req.trim() !== curReq && (
        <>
          <button onClick={() => onReq(q.id, q.c, req.trim())} style={{ ...btn, height: 32, fontSize: 12, background: X.teal, color: "#fff", borderColor: X.teal }}>
            <Check size={13} /> Save
          </button>
          <button onClick={() => setReq(curReq)} title="Undo" style={{ ...btn, height: 32, fontSize: 12, padding: "0 8px" }}>
            <XIcon size={13} />
          </button>
        </>
      )}
    </div>
  );
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", minHeight: 0, background: X.bg, ...(mob ? { overflowY: "auto", WebkitOverflowScrolling: "touch" } : null) }}>
      {/* ---------- cabeçalho ---------- */}
      {mob ? (
        /* celular: QT e situação, o cliente, o valor e quando é o próximo follow-up */
        <div style={{ background: "#fff", padding: "10px 14px 12px", borderBottom: `1px solid ${X.line}`, flexShrink: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: X.sub }}>{q.id}</span>
            {pill(s[0] + (byHand ? " · by hand" : ""), s[1], s[2], { fontSize: 11 })}
            {headRight ? <span style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6 }}>{headRight}</span> : null}
          </div>
          <div style={{ fontSize: 18, fontWeight: 800, color: X.ink, marginTop: 6, lineHeight: 1.25 }}>{q.c}</div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginTop: 6, flexWrap: "wrap" }}>
            <span style={{ fontSize: 20, fontWeight: 800, letterSpacing: -0.4, fontVariantNumeric: "tabular-nums", color: st === "won" ? X.green : st === "lost" ? X.redDk : X.ink }}>{money(st === "won" && soldV != null ? soldV : q.a)}</span>
            {next ? (
              <span style={{ fontSize: 12.5, fontWeight: 700, color: nextLate ? X.redDk : nextIn === 0 ? X.amber : X.sub }}>
                {nextLate ? `follow-up ${-nextIn} day${nextIn === -1 ? "" : "s"} late` : nextIn === 0 ? "follow-up today" : `next follow-up ${fmtDay(next)}`}
              </span>
            ) : null}
          </div>
        </div>
      ) : (
      <div style={{ background: "#fff", padding: mob ? "12px 14px" : "18px 22px 16px", display: "flex", gap: mob ? 8 : 20, alignItems: "flex-start", borderBottom: `1px solid ${X.line}`, flexWrap: mob ? "wrap" : "nowrap", flexShrink: 0 }}>
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: mob ? 7 : 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: mob ? 19 : 22, fontWeight: 800, letterSpacing: -0.5, color: X.ink }}>{q.id}</span>
            {pill(s[0] + (byHand ? " · by hand" : ""), s[1], s[2], { fontSize: 11.5, padding: "3px 10px" })}
            {lines.map((k) => (
              <span key={k} style={{ fontSize: 11, fontWeight: 800, color: LINE_COLOR[k][0], background: LINE_COLOR[k][1], padding: "3px 8px", borderRadius: 6 }}>
                {k}
              </span>
            ))}
            {mob && headRight ? <span style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6 }}>{headRight}</span> : null}
          </div>
          <div style={{ fontSize: 14.5, fontWeight: 600, color: X.ink, marginTop: 5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={q.c}>
            {q.c}
          </div>
          <div style={{ fontSize: 12, color: X.sub, marginTop: 3 }}>
            {[q.conta, q.cls].filter(Boolean).join(" · ")}
            {q.conta || q.cls ? " · " : ""}created {fmtDate(q.d)} · {age === 0 ? "today" : `${age} day${age === 1 ? "" : "s"} ago`}
          </div>
        </div>
        <div style={mob ? { width: "100%", display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" } : { textAlign: "right", flexShrink: 0 }}>
          <div style={factLbl}>{st === "won" ? "Sold" : st === "lost" ? "Lost" : "Quoted"}</div>
          <div style={{ fontSize: mob ? 21 : 26, fontWeight: 800, letterSpacing: -0.6, fontVariantNumeric: "tabular-nums", marginTop: 2, color: st === "won" ? X.green : st === "lost" ? X.redDk : X.ink }}>
            {money(st === "won" && soldV != null ? soldV : q.a)}
          </div>
          {st === "won" && soldV != null && Math.abs(soldV - q.a) > 0.005 ? <div style={{ fontSize: 11.5, color: X.sub }}>quoted {money(q.a)}</div> : null}
        </div>
        {headRight && !mob ? <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0, paddingLeft: 6 }}>{headRight}</div> : null}
      </div>
      )}

      {!mob && factsStrip}

      {/* ---------- corpo: registro (esquerda) e histórico (direita) ---------- */}
      <div style={mob ? { display: "flex", flexDirection: "column", gap: 12, padding: "12px 10px 24px" } : { flex: 1, minHeight: 0, display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 16, padding: 16 }}>
        {/* registro do follow-up */}
        <div style={mob ? { display: "flex", flexDirection: "column", gap: 12 } : { minHeight: 0, overflow: "auto", display: "flex", flexDirection: "column", gap: 12, paddingRight: 2 }} onClick={(e) => e.stopPropagation()}>
          {(st === "lost" || st === "cancelled") && onAskClose && (
            <div style={{ ...panelCard, padding: "14px 18px", display: "flex", flexDirection: "column", gap: 10, borderLeft: `4px solid ${CLOSE_COLOR[st][0]}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <XCircle size={17} color={CLOSE_COLOR[st][0]} />
                <span style={{ fontSize: 14, fontWeight: 800, color: X.ink }}>{CLOSE_LABEL[st]}</span>
                <span style={{ fontSize: 12, color: X.sub }}>{byHand ? "marked by hand · the Quote Report still says Open" : "in the Quote Report"}</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                <span style={{ fontSize: 12.5, color: "#475569" }}>Reason:</span>
                {CLOSE_REASONS[st].map((r) => {
                  const on = mm.why === r;
                  const [fg, bg] = CLOSE_COLOR[st];
                  return (
                    <button
                      key={r}
                      onClick={() => !on && (onWhy ? onWhy(q.id, q.c, r) : onAskClose(q.id, q.c, st, byHand ? "close" : "why"))}
                      style={{ ...btn, height: 30, fontSize: 12.5, borderRadius: 999, fontWeight: on ? 700 : 500, background: on ? bg : "#fff", color: on ? fg : X.ink, borderColor: on ? fg : X.line, cursor: on ? "default" : "pointer" }}
                    >
                      {on && <Check size={13} />} {r}
                    </button>
                  );
                })}
                {!mm.why && <span style={{ fontSize: 12, color: X.redDk, fontWeight: 600 }}>no reason yet</span>}
                {byHand && onReopen && (
                  <button onClick={() => onReopen(q.id, q.c)} style={{ ...btn, height: 30, fontSize: 12, marginLeft: "auto" }} title="Undo the close by hand: the quote goes back to Open">
                    <RotateCcw size={13} /> Reopen
                  </button>
                )}
              </div>
            </div>
          )}
          <div style={panelCard}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "14px 18px" }}>
              <span style={{ width: 30, height: 30, borderRadius: 9, background: X.violetSoft, color: X.violet, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                {open ? <CheckCircle2 size={17} /> : <StickyNote size={16} />}
              </span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 15, fontWeight: 800, color: X.ink }}>{open ? "Register a follow-up" : "Add a note"}</div>
                {(!mob || !open) && <div style={{ fontSize: 12, color: X.sub }}>{open ? "Note, next date and % Win — the three are needed to mark Done" : `This quote is ${s[0].toLowerCase()}: only notes can be added`}</div>}
              </div>
              {open && (
                <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 5 }} title={`${nDone} of 3 steps filled in`}>
                  {[hasNote, ndOk, winOk].map((ok, i) => (
                    <span key={i} style={{ width: 22, height: 6, borderRadius: 4, background: ok ? X.teal : "#E3E9EE" }} />
                  ))}
                  <span style={{ fontSize: 11.5, fontWeight: 700, color: nDone === 3 ? X.tealDk : X.sub, marginLeft: 4 }}>{nDone}/3</span>
                </div>
              )}
            </div>
            {/* 1 · nota */}
            <div style={sect}>
              {open ? step(1, hasNote, "What happened?", "note of this contact") : null}
              {noteBox}
              {!open && (
                <button onClick={add} disabled={!hasNote} style={{ ...btn, alignSelf: "flex-end", height: 34, background: hasNote ? X.teal : "#fff", color: hasNote ? "#fff" : X.sub, borderColor: hasNote ? X.teal : X.line, cursor: hasNote ? "pointer" : "not-allowed" }}>
                  <StickyNote size={14} /> Add note
                </button>
              )}
            </div>
            {open && (
              <>
                {/* 2 · próxima data */}
                <div style={sect}>
                  {step(2, ndOk, "Next follow-up", next ? <>now {fmtDate(next)}{mm.fup ? ` · ${fupTxt(mm.fup)}` : ""}</> : "no date yet")}
                  <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                    <input
                      type="date"
                      value={nd}
                      min={addDays(today, 1)}
                      onChange={(e) => setNd(e.target.value)}
                      style={{ ...inp, height: 32, fontSize: 13, width: 138, borderRadius: 9, borderColor: !ndOk ? "#F2C48D" : "#9FD6C9", background: !ndOk ? "#FFFCF7" : "#fff", fontWeight: ndOk ? 700 : 400 }}
                    />
                    {quick.map(([l, v]) => {
                      const on = ndFix === v;
                      return (
                        <button
                          key={l}
                          onClick={() => setNd(v)}
                          title={`${dowName(v)}, ${fmtDate(v)}`}
                          style={{ ...btn, height: 30, padding: "0 9px", fontSize: 11.5, borderRadius: 999, background: on ? X.tealSoft : "#fff", color: on ? X.tealDk : X.ink, borderColor: on ? "#9FD6C9" : X.line, fontWeight: on ? 700 : 600 }}
                        >
                          {l}
                        </button>
                      );
                    })}
                  </div>
                  {nd && !ndOk ? (
                    <div style={{ fontSize: 12, color: X.redDk, fontWeight: 600 }}>Choose a date after {fmtDate(today)}.</div>
                  ) : ndOk ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", fontSize: 12, color: X.sub }}>
                      <ArrowRight size={13} color={X.violet} />
                      <b style={{ color: X.ink }}>
                        {dowName(ndFix)}, {fmtDate(ndFix)}
                      </b>
                      · in {daysBetween(today, ndFix)} day{daysBetween(today, ndFix) === 1 ? "" : "s"}
                      {ndFix !== nd ? ` (${holidayOf(nd) && holidayOf(nd).block ? `${fmtDay(nd)} is a holiday: ${holidayOf(nd).name}` : "weekend"} → next business day)` : ""} · moves to <FupPill v="E" />
                    </div>
                  ) : null}
                </div>
                {/* 3 · % Win */}
                <div style={sect}>
                  {step(3, winOk, "% Win", curWin ? `is it still ${curWin}%? confirm or change` : "not set yet: choose the chance")}
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 8 }}>
                    {WIN_VALS.map((v) => {
                      const on = wv === v;
                      const wm = WIN_META[v];
                      const isCur = curWin === v;
                      return (
                        <button
                          key={v}
                          onClick={() => setWv(v)}
                          title={isCur ? `Keep ${v}%` : `Change to ${v}%`}
                          style={{
                            ...btn,
                            height: 46,
                            flexDirection: "column",
                            gap: 1,
                            justifyContent: "center",
                            borderRadius: 10,
                            border: `2px solid ${on ? wm.color : isCur ? wm.color + "66" : X.line}`,
                            background: on ? wm.bg : "#fff",
                            color: wm.color,
                            boxShadow: on ? `0 0 0 3px ${wm.color}22` : "none",
                          }}
                        >
                          <span style={{ fontSize: 15, fontWeight: 800, lineHeight: 1.1 }}>{v}%</span>
                          <span style={{ fontSize: 10, fontWeight: 700, color: on ? wm.color : X.sub, lineHeight: 1.1 }}>{isCur ? (on ? "kept" : "now") : on ? (curWin ? `was ${curWin}%` : "chosen") : " "}</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
                {/* ações */}
                <div style={{ ...sect, background: "#FAFBFC", borderRadius: "0 0 14px 14px" }}>
                  <button
                    onClick={done}
                    disabled={!canDone}
                    title={canDone ? "Saves the note, records the follow-up as done, updates the % Win and schedules the next one" : "Fill in the note, the next follow-up date and the % Win"}
                    style={{
                      ...btn,
                      width: "100%",
                      justifyContent: "center",
                      height: 44,
                      fontSize: 14,
                      fontWeight: 700,
                      borderRadius: 11,
                      background: canDone ? X.violet : "#EEF1F4",
                      color: canDone ? "#fff" : "#8A97A3",
                      borderColor: canDone ? X.violet : X.line,
                      cursor: canDone ? "pointer" : "not-allowed",
                      boxShadow: canDone ? "0 3px 10px rgba(91,82,201,.3)" : "none",
                    }}
                  >
                    <CheckCircle2 size={17} />
                    {canDone ? `Mark Done · ${wv}% · next follow-up ${fmtDay(ndFix)}` : "Done · missing: " + [!hasNote && "note", !ndOk && "date", !winOk && "% Win"].filter(Boolean).join(", ")}
                  </button>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <button onClick={add} disabled={!hasNote} style={subBtn(hasNote)} title="Save only the note (no Done)">
                      <StickyNote size={14} /> Save only the note
                    </button>
                    <button onClick={resched} disabled={!canResched} style={subBtn(canResched)} title="Change the next follow-up date without marking Done (the quote moves to E · Ongoing; the note is optional)">
                      <CalendarDays size={14} /> Only change the date
                    </button>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span style={{ fontSize: 11, color: X.sub, display: "inline-flex", alignItems: "center", gap: 4 }}>
                      <Lock size={11} /> notes can't be edited or deleted
                    </span>
                    {onAskClose && (
                      <span style={{ marginLeft: "auto", display: "inline-flex", alignItems: "center", gap: 6 }}>
                        <span style={{ fontSize: 11.5, color: X.sub }}>Closing it?</span>
                        {["lost", "cancelled"].map((k) => (
                          <button key={k} onClick={() => onAskClose(q.id, q.c, k)} style={{ ...btn, height: 28, fontSize: 11.5, padding: "0 10px", color: CLOSE_COLOR[k][0], borderColor: X.line }} title={`Mark as ${CLOSE_LABEL[k]} (with the reason)`}>
                            <XCircle size={12} /> {CLOSE_LABEL[k]}
                          </button>
                        ))}
                      </span>
                    )}
                  </div>
                </div>
              </>
            )}
          </div>
          {!mob && reqBox}
          {/* celular: os dados da cotação e a requisição ficam recolhidos em "Details" */}
          {mob && (
            <div style={{ ...panelCard, overflow: "hidden" }}>
              <div onClick={() => setShowDet((v) => !v)} className="qp-card-btn" style={{ display: "flex", alignItems: "center", gap: 9, padding: "13px 14px", cursor: "pointer" }}>
                <Info size={16} color={X.ink} />
                <span style={{ fontWeight: 800, fontSize: 14.5 }}>Details</span>
                <span style={{ fontSize: 12, color: X.sub, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 }}>FUP, % Win, salespeople, orders</span>
                <ChevronDown size={17} color={X.sub} style={{ marginLeft: "auto", flexShrink: 0, transform: showDet ? "rotate(180deg)" : "none", transition: "transform .14s" }} />
              </div>
              {showDet && (
                <>
                  <div style={{ borderTop: `1px solid ${X.line}` }}>{factsStrip}</div>
                  <div style={{ padding: "10px 14px", fontSize: 12, color: X.sub, display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                    {lines.map((k) => (
                      <span key={k} style={{ fontSize: 11, fontWeight: 800, color: LINE_COLOR[k][0], background: LINE_COLOR[k][1], padding: "2px 7px", borderRadius: 6 }}>
                        {k}
                      </span>
                    ))}
                    {[q.conta, q.cls].filter(Boolean).join(" · ")}
                    {q.conta || q.cls ? " · " : ""}created {fmtDate(q.d)} · {age === 0 ? "today" : `${age} day${age === 1 ? "" : "s"} ago`}
                  </div>
                  {reqBox}
                </>
              )}
            </div>
          )}
        </div>

        {/* Activities / History */}
        <div style={{ ...panelCard, minHeight: 0, display: "flex", flexDirection: "column", overflow: mob ? "visible" : "hidden" }}>
          <div
            onClick={mob ? () => setShowHist((v) => !v) : undefined}
            className={mob ? "qp-card-btn" : undefined}
            style={{ display: "flex", alignItems: "center", gap: 9, padding: mob ? "13px 14px" : "13px 16px", borderBottom: !mob || showHist ? `1px solid ${X.line}` : "none", flexWrap: "wrap", cursor: mob ? "pointer" : "default" }}
          >
            <Clock size={mob ? 16 : 17} color={X.ink} />
            <span style={{ fontWeight: 800, fontSize: 14.5 }}>{mob ? "History" : "Activities / History"}</span>
            {mob && <span style={{ fontSize: 12, fontWeight: 700, color: X.sub, background: X.head, borderRadius: 999, padding: "1px 8px" }}>{list.length}</span>}
            {mob && <ChevronDown size={17} color={X.sub} style={{ marginLeft: "auto", transform: showHist ? "rotate(180deg)" : "none", transition: "transform .14s" }} />}
            {(!mob || showHist) && (
            <div onClick={(e) => e.stopPropagation()} className="qp-hscroll" style={{ marginLeft: mob ? 0 : "auto", display: "inline-flex", gap: 3, background: "#EEF2F5", borderRadius: 9, padding: 3, ...(mob ? { width: "100%" } : null) }}>
              {HIST_FILTERS.map(([k, l]) => {
                const on = hf === k;
                const n = k === "all" ? list.length : list.filter((a) => histGroup(a) === k).length;
                return (
                  <button
                    key={k}
                    onClick={(e) => {
                      e.stopPropagation();
                      setHf(k);
                    }}
                    style={{ border: "none", borderRadius: 7, padding: "4px 10px", fontSize: 11.5, fontWeight: on ? 700 : 500, cursor: "pointer", background: on ? "#fff" : "transparent", color: on ? X.tealDk : X.sub, boxShadow: on ? "0 1px 2px rgba(16,24,40,.08)" : "none" }}
                  >
                    {l} <span style={{ opacity: 0.7 }}>{n}</span>
                  </button>
                );
              })}
            </div>
            )}
          </div>
          {(!mob || showHist) && (
          <div style={mob ? { padding: "4px 12px 10px" } : { flex: 1, minHeight: 0, overflow: "auto", padding: "4px 16px 10px" }}>
            {!shownH.length && <div style={{ fontSize: 12.5, color: X.sub, padding: "16px 0" }}>Nothing of this type yet.</div>}
            {dayGroups.map(([day, items]) => (
              <div key={day}>
                <div style={{ position: "sticky", top: 0, zIndex: 2, background: "#fff", padding: "10px 0 6px", display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 11, fontWeight: 800, color: X.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{dayLabel(day)}</span>
                  <span style={{ flex: 1, height: 1, background: X.line }} />
                  <span style={{ fontSize: 11, color: X.sub }}>{items.length}</span>
                </div>
                {items.map((a, i) => {
                  const mt = ACT_META[a.k] || ACT_META._seen;
                  const tg = ACT_TAG[mt.tag];
                  const isNote = a.k === "note";
                  const last = i === items.length - 1;
                  return (
                    <div key={a._i} style={{ display: "flex", gap: 12 }}>
                      <div style={{ position: "relative", width: 34, flexShrink: 0, display: "flex", justifyContent: "center" }}>
                        <span style={{ width: 34, height: 34, borderRadius: 999, background: mt.bg, color: mt.color, display: "inline-flex", alignItems: "center", justifyContent: "center", position: "relative", zIndex: 1 }}>
                          {React.createElement(mt.icon, { size: 16 })}
                        </span>
                        {!last && <span style={{ position: "absolute", top: 38, bottom: 4, left: "50%", width: 2, marginLeft: -1, background: "#E3E8EF", borderRadius: 2 }} />}
                      </div>
                      <div style={{ minWidth: 0, flex: 1, paddingBottom: last ? 6 : 16 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginTop: 1 }}>
                          <span style={{ fontSize: 13.5, fontWeight: 800, color: X.ink }}>{mt.title}</span>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11, fontWeight: 600, color: tg[1], background: tg[2], padding: "2px 9px", borderRadius: 999 }}>
                            {tg[0]}
                            {isNote && <Lock size={10} />}
                          </span>
                          <span style={{ marginLeft: "auto", fontSize: 11, color: X.sub, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" }} title={a.dateOnly ? fmtDate(a.t) : fmtStamp(a.t)}>
                            {a.dateOnly ? "" : a.t.slice(11, 16)}
                            {a.by ? ` · ${a.by}` : ""}
                          </span>
                        </div>
                        <div
                          style={{
                            fontSize: 12.5,
                            color: "#475569",
                            marginTop: 4,
                            lineHeight: 1.5,
                            whiteSpace: isNote ? "pre-wrap" : "normal",
                            wordBreak: "break-word",
                            ...(isNote ? { background: "#F7FBFA", border: `1px solid ${X.line}`, borderLeft: `3px solid ${X.teal}`, borderRadius: 8, padding: "8px 10px", color: X.ink } : null),
                          }}
                        >
                          {isNote ? a.txt : actText(a)}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
          )}
        </div>
      </div>
    </div>
  );
}
/* contador de notas da cotação (na tabela e no plano); passando o mouse, mostra a última nota */
function NoteBadge({ info, onClick }) {
  if (!info || !info.n) return null;
  const l = info.last;
  return (
    <span
      onClick={onClick}
      title={`${info.n} note${info.n === 1 ? "" : "s"} · last ${fmtStamp(l.t)}${l.by ? " by " + l.by : ""}:\n${l.txt}`}
      style={{ display: "inline-flex", alignItems: "center", gap: 3, marginLeft: 6, fontSize: 10.5, fontWeight: 700, color: X.tealDk, background: X.tealSoft, padding: "1px 6px", borderRadius: 999, cursor: onClick ? "pointer" : "help", verticalAlign: 1 }}
    >
      <MessageSquare size={11} /> {info.n}
    </span>
  );
}
/* Fechamento: pílula "Lost · Price" / "Cancelled · Budget Only" (sem motivo = tracejada) */
const CLOSE_COLOR = { lost: [X.redDk, X.redSoft], cancelled: ["#475569", "#E9EDF2"] };
function ClosePill({ kind, why, onClick, title }) {
  const [fg, bg] = CLOSE_COLOR[kind] || CLOSE_COLOR.cancelled;
  return (
    <span
      onClick={onClick}
      title={title}
      style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11.5, fontWeight: 700, color: fg, background: why ? bg : "#fff", border: why ? "1px solid transparent" : `1px dashed ${fg}`, padding: "2px 9px", borderRadius: 999, cursor: onClick ? "pointer" : "default", whiteSpace: "nowrap" }}
    >
      {CLOSE_LABEL[kind] || kind}
      <span style={{ fontWeight: 500, opacity: 0.85 }}>· {why || "no reason"}</span>
    </span>
  );
}
/* Diálogo para marcar Cancelled / Lost (com o motivo, obrigatório) ou só escolher o motivo de uma cotação
   que já veio fechada no Quote Report. ask = { id, c, kind, mode: "close" | "why", cur, manual } */
function CloseDialog({ ask, onConfirm, onReopen, onCancel }) {
  const [why, setWhy] = useState(ask.cur || "");
  useEffect(() => {
    const k = (e) => e.key === "Escape" && onCancel();
    window.addEventListener("keydown", k);
    return () => window.removeEventListener("keydown", k);
  }, [onCancel]);
  const [fg, bg] = CLOSE_COLOR[ask.kind] || CLOSE_COLOR.cancelled;
  const reasons = CLOSE_REASONS[ask.kind] || [];
  return createPortal(
    <div onClick={onCancel} style={{ position: "fixed", inset: 0, background: "rgba(16,24,40,.35)", zIndex: 2000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ ...card, width: "min(440px, 100%)", borderRadius: 14, padding: 20, display: "flex", flexDirection: "column", gap: 14, boxShadow: "0 20px 50px rgba(16,24,40,.25)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ width: 34, height: 34, borderRadius: 10, background: bg, color: fg, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <XCircle size={18} />
          </span>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 15.5, fontWeight: 800, color: X.ink }}>{ask.mode === "why" ? `Why is it ${CLOSE_LABEL[ask.kind].toLowerCase()}?` : `Mark as ${CLOSE_LABEL[ask.kind]}`}</div>
            <div style={{ fontSize: 12, color: X.sub, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {ask.id} · {ask.c}
            </div>
          </div>
        </div>
        {ask.mode !== "why" && (
          <div style={{ fontSize: 12.5, color: "#475569", lineHeight: 1.5 }}>
            The quote leaves the follow-up calendar and moves to <b>{CLOSE_LABEL[ask.kind]}</b> {ask.kind === "lost" ? "(its value goes to Lost)" : "(it leaves the totals, like the cancelled ones)"}, even before the Quote Report changes.
          </div>
        )}
        <div>
          <div style={{ fontSize: 11, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Reason · required</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {reasons.map((r) => {
              const on = why === r;
              return (
                <button
                  key={r}
                  onClick={() => setWhy(r)}
                  style={{ ...btn, height: 40, justifyContent: "flex-start", gap: 10, borderRadius: 10, fontSize: 13.5, fontWeight: on ? 700 : 500, borderColor: on ? fg : X.line, background: on ? bg : "#fff", color: on ? fg : X.ink, borderWidth: on ? 2 : 1 }}
                >
                  <span style={{ width: 16, height: 16, borderRadius: 999, border: `2px solid ${on ? fg : "#CBD5E1"}`, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    {on && <span style={{ width: 8, height: 8, borderRadius: 999, background: fg }} />}
                  </span>
                  {r}
                </button>
              );
            })}
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {ask.manual && onReopen && (
            <button onClick={onReopen} style={{ ...btn, height: 36, fontSize: 12.5 }} title="Undo the close by hand: the quote goes back to Open">
              <RotateCcw size={14} /> Reopen
            </button>
          )}
          <button onClick={onCancel} style={{ ...btn, height: 36, marginLeft: "auto" }}>
            Cancel
          </button>
          <button
            onClick={() => why && onConfirm(why)}
            disabled={!why}
            style={{ ...btn, height: 36, fontWeight: 700, background: why ? fg : "#EEF1F4", borderColor: why ? fg : X.line, color: why ? "#fff" : "#8A97A3", cursor: why ? "pointer" : "not-allowed" }}
          >
            <Check size={14} /> {ask.mode === "why" ? "Save reason" : `Mark as ${CLOSE_LABEL[ask.kind]}`}
          </button>
        </div>
      </div>
    </div>,
    document.body,
  );
}
/* gaveta lateral com o registro da cotação (usada no plano do dia) */
/* Painel lateral da cotação (abre por cima da tela, do lado direito). nav = { i, n, go(delta) } quando veio
   da tabela: setas para a cotação anterior / próxima da lista (também com Alt+↑ / Alt+↓). */
function QuoteDrawer({ q, m, sales, acts, thermo, onNote, onDone, onResched, onReq, onAskClose, onReopen, onWhy, focusDone, onClose, user, onUser, nav }) {
  const mob = useMobile(); // celular: o painel ocupa a tela inteira
  useEffect(() => {
    const k = (e) => {
      if (e.key === "Escape") onClose();
      if (nav && e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
        e.preventDefault();
        nav.go(e.key === "ArrowUp" ? -1 : 1);
      }
    };
    window.addEventListener("keydown", k);
    return () => window.removeEventListener("keydown", k);
  }, [onClose, nav]);
  return createPortal(
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(16,24,40,.28)", zIndex: 1500, display: "flex", justifyContent: "flex-end" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: mob ? "100vw" : "min(1180px, 96vw)", height: "100%", background: X.bg, boxShadow: "-10px 0 30px rgba(16,24,40,.18)", display: "flex", flexDirection: "column" }}>
        <QuoteRecord
          key={q ? q.id : ""}
          q={q}
          m={m}
          sales={sales}
          acts={acts}
          thermo={thermo}
          onNote={onNote}
          onDone={onDone}
          onResched={onResched}
          onReq={onReq}
          onAskClose={onAskClose}
          onReopen={onReopen}
          onWhy={onWhy}
          focusDone={focusDone}
          user={user}
          onUser={onUser}
          compact
          headRight={
            <>
              {nav && nav.n > 1 && (
                <>
                  <button onClick={() => nav.go(-1)} disabled={nav.i <= 0} style={{ ...btn, padding: "0 9px", opacity: nav.i <= 0 ? 0.4 : 1 }} title="Previous quote of the table (Alt+↑)">
                    <ChevronLeft size={16} />
                  </button>
                  <span style={{ fontSize: 12, color: X.sub, whiteSpace: "nowrap", minWidth: mob ? 0 : 64, textAlign: "center" }}>
                    {nav.i + 1}{mob ? "/" : " of "}{nav.n.toLocaleString("pt-BR")}
                  </span>
                  <button onClick={() => nav.go(1)} disabled={nav.i >= nav.n - 1} style={{ ...btn, padding: "0 9px", opacity: nav.i >= nav.n - 1 ? 0.4 : 1 }} title="Next quote of the table (Alt+↓)">
                    <ChevronRight size={16} />
                  </button>
                </>
              )}
              <button onClick={onClose} style={{ ...btn, padding: "0 10px", marginLeft: 4 }} title="Close (Esc)">
                <XIcon size={16} />
              </button>
            </>
          }
        />
      </div>
    </div>,
    document.body,
  );
}

/* ===== Quote portfolio: uma linha por cotação ===== */
const FUP_RANK = { A: 4, B: 3, C: 2, D: 1 }; // A = mais urgente
/* Colunas da carteira. Categorias filtram (Customer, FUP, % Win, Thermometer); data e valores
   ordenam (Date, Quoted, Open, Lost, Sold). QT não precisa de nenhum dos dois: a busca acha a cotação. */
const PCOLS = [
  { id: "id", label: "QT", tip: "Quote ID · use the search to find a quote", w: 150 },
  { id: "c", label: "Customer", sort: true, filter: true, fw: 320, tip: "Customer · click the name to sort A–Z / Z–A · the funnel filters" },
  { id: "d", label: "Date", sort: true, tip: "Quote date", w: 110 },
  { id: "fup", label: "FUP", filter: true, manual: true, tip: "Follow-up deadline · A = next day · B = twice a week · C = 2 weeks · D = 1 month", w: 190 },
  { id: "win", label: "% Win", filter: true, manual: true, tip: "Chance of closing this quote", w: 110 },
  { id: "temp", label: "Thermometer", filter: true, tip: `Automatic, by customer · hit rate of the last ${THERMO_HR_MONTHS} months · minimum hit rate: ${THERMO_HR_MIN * 100}% for OEM customers, ${THERMO_HR_MIN_OTHER * 100}% for the other classifications · Hot = hit rate ≥ minimum and quotes on ${THERMO_DAYS_MIN}+ of the last ${THERMO_WINDOW} business days · Warm = ≥ minimum and fewer days · Cool = below the minimum and ${THERMO_DAYS_MIN}+ days · Cold = below the minimum and fewer days · fewer than ${THERMO_RECENT_MIN} quote days in the last ${THERMO_RECENT_WINDOW} business days drops one level (↓)`, w: 170 },
  { id: "quoted", label: "Quoted", sort: true, right: true, w: 140 },
  { id: "open", label: "Open", sort: true, right: true, w: 140 },
  { id: "lostV", label: "Lost", sort: true, right: true, w: 140 },
  { id: "sold", label: "Sold", sort: true, right: true, w: 150, tip: "Booked value of the order (Sales Report) · without it, the quote's Sales Amt" },
];
const sortVal = (r, id) =>
  id === "id" || id === "c" || id === "d" ? r[id] || "" : id === "fup" ? FUP_RANK[r.fup] || 0 : id === "temp" ? (TEMP[r.temp] ? TEMP[r.temp].rank : 0) : r[id] || 0;
const fKey = (r, id) => (id === "c" ? r.c || NONE : id === "fup" ? r.fup || NONE : id === "win" ? (r.win ? String(r.win) : NONE) : id === "temp" ? r.temp || NONE : "");
const STATUS_TABS = [
  ["", "All"],
  ["open", "Open"],
  ["won", "Won"],
  ["lost", "Lost"],
  ["cancelled", "Cancelled"],
];
const ROW_CAP = 400; // linhas desenhadas de uma vez; o resto vem no "Show all"

/* cards = cartões no lugar da tabela (celular e a tela do vendedor externo); mesma busca, abas e filtros */
function QuoteTable({ onAskClose, rows, cancelRows, meta, onMeta, today, periodLabel, thermo, qIndex, sales, acts, onNote, onDone, onResched, onReq, user, onUser, notesBy, openQ, onOpenQuote, cards, title, filtersExtra, filtersExtraN }) {
  const [srch, setSrch] = useState("");
  const [sheet, setSheet] = useState(false); // cartões: folha de filtros
  const [sort, setSort] = useState({ id: "c", dir: 1 }); // padrão: cliente A–Z (e a cotação mais nova primeiro)
  const [flt, setFlt] = useState({});
  const [stF, setStF] = useState(cards ? "open" : ""); // cartões: começa nas abertas (o que ainda dá para fechar)
  const [all, setAll] = useState(false);
  const enriched = useMemo(
    () =>
      [...rows, ...(cancelRows || [])].map((r) => {
        const m = meta[r.id] || {};
        const t = thermo[r.nn] || null;
        const pend = r.st === "open" ? fupPending(m) : [];
        return { ...r, fup: m.fup || "", win: m.win || 0, why: m.why || "", temp: t ? t.temp : "", next: pend[0] || "", next2: pend[1] || "" };
      }),
    [rows, cancelRows, meta, thermo],
  );
  const searched = useMemo(() => {
    const s = norm(srch);
    return s ? enriched.filter((r) => norm(r.id + " " + r.c + " " + r.conta).includes(s)) : enriched;
  }, [enriched, srch]);
  const stCount = useMemo(() => {
    // "All" = todas menos as canceladas (as canceladas só aparecem na aba Cancelled e não entram nos totais)
    const m = { "": searched.filter((r) => r.st !== "cancelled").length };
    searched.forEach((r) => (m[r.st] = (m[r.st] || 0) + 1));
    return m;
  }, [searched]);
  const shown = useMemo(() => {
    const act = Object.entries(flt).filter(([, v]) => v && v.length);
    const f = searched.filter((r) => (stF ? r.st === stF : r.st !== "cancelled") && act.every(([id, v]) => v.includes(fKey(r, id))));
    return f.sort((a, b) => {
      const va = sortVal(a, sort.id),
        vb = sortVal(b, sort.id);
      const c = typeof va === "string" ? va.localeCompare(vb, "pt-BR") : va - vb;
      return c * sort.dir || (a.d < b.d ? 1 : a.d > b.d ? -1 : 0) || a.id.localeCompare(b.id);
    });
  }, [searched, flt, sort, stF]);
  const fltOpts = (id) => {
    const cnt = {};
    searched.forEach((r) => {
      if (stF ? r.st !== stF : r.st === "cancelled") return;
      const k = fKey(r, id);
      cnt[k] = (cnt[k] || 0) + 1;
    });
    if (id === "c") {
      // clientes: os que têm cotação no recorte, mais os já marcados; os de mais cotações primeiro
      const sel = flt.c || [];
      const ks = [...new Set([...Object.keys(cnt), ...sel])].filter((k) => k !== NONE);
      return ks
        .sort((a, b) => sel.includes(b) - sel.includes(a) || a.localeCompare(b, "pt-BR", { sensitivity: "base" }))
        .map((k) => ({ v: k, n: cnt[k] || 0, text: k, node: k }));
    }
    const keys = id === "fup" ? FUP_KEYS : id === "win" ? WIN_VALS.map(String) : TEMP_KEYS;
    return [
      ...keys.map((k) => ({ v: k, n: cnt[k] || 0, text: id === "fup" ? `${k} ${FUP[k].label}` : id === "win" ? k + "%" : TEMP[k].label, node: id === "fup" ? <FupPill v={k} /> : id === "win" ? <WinPill v={+k} /> : <TempPill v={k} /> })),
      // o Thermometer é automático e sempre tem valor: sem a opção "Not set"
      ...(id === "temp" ? [] : [{ v: NONE, n: cnt[NONE] || 0, node: <span style={{ color: X.sub }}>Not set</span> }]),
    ];
  };
  // 1º clique: cliente A–Z, nos outros o maior/mais recente primeiro; 2º clique inverte
  const clickSort = (c) => setSort((s) => (s.id === c.id ? { id: c.id, dir: -s.dir } : { id: c.id, dir: c.id === "c" ? 1 : -1 }));
  const sum = (k) => shown.reduce((a, r) => a + (r[k] || 0), 0);
  const cnt = (st) => shown.filter((r) => r.st === st).length;
  const tot = { quoted: sum("quoted"), open: sum("open"), lostV: sum("lostV"), sold: sum("sold"), qt: shown.length, openN: cnt("open"), won: cnt("won"), lost: cnt("lost") };
  const fltOn = Object.values(flt).some((v) => v && v.length) || !!stF;
  // cartões: menos de uma vez (no celular cada cartão é maior que uma linha da tabela)
  const vis = all ? shown : shown.slice(0, cards ? 60 : ROW_CAP);

  const exportXlsx = () => {
    const data = shown.map((r) => ({
      QT: r.id,
      Customer: r.c,
      Account: r.conta,
      Date: fmtDate(r.d),
      Result: (ST_PILL[r.st] || ST_PILL.open)[0] + (r.manual ? " (by hand)" : ""),
      "Close reason": r.why || "",
      FUP: r.fup ? `${r.fup} - ${FUP[r.fup].label}` : "",
      "Next FUP": r.next ? fmtDate(r.next) : "",
      "% Win": r.win ? r.win / 100 : "",
      Thermometer: r.temp ? TEMP[r.temp].label : "",
      Quoted: r.quoted,
      Open: r.open,
      Lost: r.lostV,
      Sold: r.sold,
      "Order (SO)": r.sos.map((s) => s.id).join(", "),
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 13 }, { wch: 44 }, { wch: 10 }, { wch: 11 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 11 }, { wch: 8 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = 1; R <= range.e.r; R++) {
      const pc = ws[XLSX.utils.encode_cell({ r: R, c: 8 })];
      if (pc && typeof pc.v === "number") pc.z = "0%";
      for (let C = 10; C <= 13; C++) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell) cell.z = '"R$" #,##0.00';
      }
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Quotes");
    XLSX.writeFile(wb, `Quotes ${today}.xlsx`);
  };

  /* ---------- cartões (mobile first) ----------
     Na tela: busca, Open / Won / Lost / All e a lista. Ordem, filtros por coluna, período, classificação,
     linha e exportar ficam na folha de filtros (botão ao lado da busca). Cada cartão mostra só o essencial:
     cliente, QT, data, valor e a situação (próximo follow-up, atrasado, ganho, perdido). Tocar abre a cotação. */
  if (cards) {
    const SORTS = [
      ["c:1", "Customer A–Z"],
      ["d:-1", "Newest"],
      ["fup:-1", "Most urgent FUP"],
      ["open:-1", "Biggest open"],
      ["quoted:-1", "Biggest quoted"],
      ["sold:-1", "Biggest sold"],
    ];
    const sortV = `${sort.id}:${sort.dir}`;
    const ids = () => shown.map((x) => x.id);
    const nCol = Object.values(flt).filter((v) => v && v.length).length;
    const nAll = nCol + (filtersExtraN || 0);
    const TABS = [
      ["open", "Open"],
      ["won", "Won"],
      ["lost", "Lost"],
      ["", "All"],
    ];
    const sumTxt = stF === "open" ? `${moneyShort(tot.open)} open` : stF === "won" ? `${moneyShort(tot.sold)} sold` : stF === "lost" ? `${moneyShort(tot.lostV)} lost` : stF === "cancelled" ? `${moneyShort(tot.quoted)} cancelled` : `${moneyShort(tot.quoted)} quoted`;
    const situ = (r) => {
      if (r.st === "won") return ["Won", X.green];
      if (r.st === "lost") return [r.why ? `Lost · ${r.why}` : "Lost", X.redDk];
      if (r.st === "cancelled") return [r.why ? `Cancelled · ${r.why}` : "Cancelled", "#64748B"];
      if (!r.fup) return ["No FUP", X.amber];
      if (r.next && r.next < today) return [`Overdue · ${fmtDay(r.next)}`, X.redDk];
      if (r.next === today) return ["Follow-up today", X.amber];
      return [r.next ? `Next ${fmtDay(r.next)}` : FUP[r.fup].label, X.sub];
    };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {/* busca + filtros */}
        <div style={{ display: "flex", gap: 8 }}>
          <div style={{ position: "relative", flex: 1, minWidth: 0 }}>
            <Search size={16} style={{ position: "absolute", left: 13, top: "50%", transform: "translateY(-50%)", color: X.sub, pointerEvents: "none" }} />
            <input value={srch} onChange={(e) => setSrch(e.target.value)} placeholder="Search customer or QT…" style={{ ...inp, width: "100%", height: 44, paddingLeft: 38, paddingRight: srch ? 38 : 10, fontSize: 15, borderRadius: 12 }} />
            {srch && (
              <button onClick={() => setSrch("")} title="Clear" style={{ position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)", border: "none", background: "transparent", color: X.sub, cursor: "pointer", padding: 6 }}>
                <XIcon size={16} />
              </button>
            )}
          </div>
          <button onClick={() => setSheet(true)} title="Filters, order and export" style={iconBtn({ width: 44, height: 44, ...(nAll ? { background: X.tealSoft, borderColor: "#9FD6C9", color: X.tealDk } : null) })}>
            <SlidersHorizontal size={18} />
            {dotBadge(nAll)}
          </button>
        </div>
        {/* Open / Won / Lost / All */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 3, background: X.head, borderRadius: 12, padding: 3 }}>
          {TABS.map(([k, l]) => {
            const on = stF === k;
            return (
              <button
                key={k || "all"}
                onClick={() => setStF(k)}
                style={{ border: "none", borderRadius: 9, padding: "8px 2px", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 1, background: on ? "#fff" : "transparent", color: on ? X.tealDk : X.sub, boxShadow: on ? "0 1px 2px rgba(16,24,40,.1)" : "none" }}
              >
                <span style={{ fontSize: 13, fontWeight: on ? 800 : 600 }}>{l}</span>
                <span style={{ fontSize: 11, fontWeight: 600, opacity: 0.8 }}>{(stCount[k] || 0).toLocaleString("pt-BR")}</span>
              </button>
            );
          })}
        </div>
        <div style={{ fontSize: 12.5, color: X.sub, padding: "0 2px", display: "flex", gap: 6, alignItems: "center" }}>
          <span>
            <b style={{ color: X.ink }}>{shown.length.toLocaleString("pt-BR")}</b> {shown.length === 1 ? "quote" : "quotes"} · {sumTxt}
          </span>
          {stF === "cancelled" && (
            <button onClick={() => setStF("")} style={{ border: "none", background: "transparent", color: X.tealDk, fontWeight: 700, cursor: "pointer", fontSize: 12.5, marginLeft: "auto" }}>
              back to All
            </button>
          )}
        </div>
        {/* cartões */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(min(100%, 320px), 1fr))", gap: 8 }}>
          {vis.map((r) => {
            const [sTxt, sCol] = situ(r);
            const late = r.st === "open" && r.next && r.next < today;
            const v = r.st === "won" ? r.sold : r.st === "lost" ? r.lostV : r.st === "open" ? r.open || r.quoted : r.quoted;
            return (
              <div
                key={r.id}
                className="qp-card-btn"
                onClick={() => onOpenQuote(r.id, false, ids())}
                style={{ ...card, padding: "12px 13px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", minWidth: 0, boxShadow: openQ === r.id ? `0 0 0 2px ${X.teal}` : "none", ...(late ? { borderLeft: `4px solid ${X.redDk}`, paddingLeft: 10 } : null) }}
              >
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: X.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.c}>
                    {r.c}
                  </div>
                  <div style={{ fontSize: 12, color: X.sub, marginTop: 3, display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap" }}>
                    {r.id} · {fmtDay(r.d)}
                    <NoteBadge info={notesBy[r.id]} />
                  </div>
                </div>
                <div style={{ textAlign: "right", flexShrink: 0, maxWidth: "48%" }}>
                  <div style={{ fontSize: 14, fontWeight: 800, color: r.st === "won" ? X.green : r.st === "lost" ? X.redDk : X.ink, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" }}>{money(v)}</div>
                  <div style={{ fontSize: 11.5, fontWeight: 700, color: sCol, marginTop: 3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{sTxt}</div>
                </div>
              </div>
            );
          })}
        </div>
        {!shown.length && <div style={{ padding: "36px 14px", textAlign: "center", color: X.sub, fontSize: 13.5 }}>{rows.length ? "No quote found." : "No quote in this period."}</div>}
        {shown.length > vis.length && (
          <button onClick={() => setAll(true)} style={{ ...btn, height: 44, justifyContent: "center", borderRadius: 12 }}>
            Show all {shown.length.toLocaleString("pt-BR")}
          </button>
        )}

        {/* folha: filtros, ordem e exportar */}
        <Sheet
          open={sheet}
          title="Filters"
          onClose={() => setSheet(false)}
          footer={
            <>
              {(nCol > 0 || stF === "cancelled") && (
                <button
                  onClick={() => {
                    setFlt({});
                    if (stF === "cancelled") setStF("open");
                  }}
                  style={{ ...btn, height: 44, flex: 1, justifyContent: "center", borderRadius: 12 }}
                >
                  Clear
                </button>
              )}
              <button onClick={() => setSheet(false)} style={{ ...btn, height: 44, flex: 2, justifyContent: "center", borderRadius: 12, background: X.teal, borderColor: X.teal, color: "#fff" }}>
                Show {shown.length.toLocaleString("pt-BR")} {shown.length === 1 ? "quote" : "quotes"}
              </button>
            </>
          }
        >
          <SheetField label="Order">
            <Chips
              value={sortV}
              options={SORTS}
              onChange={(v) => {
                const [id, d] = v.split(":");
                setSort({ id, dir: +d });
              }}
            />
          </SheetField>
          <SheetField label="Show only">
            <div>
              {["c", "fup", "win", "temp"].map((id) => {
                const c = PCOLS.find((x) => x.id === id);
                return <ColFilter key={id} row label={c.label} width={c.fw || 260} options={fltOpts(id)} selected={flt[id] || []} onChange={(v) => setFlt((f) => ({ ...f, [id]: v }))} />;
              })}
              <div
                onClick={() => {
                  setStF(stF === "cancelled" ? "" : "cancelled");
                  setSheet(false);
                }}
                style={{ display: "flex", alignItems: "center", gap: 8, padding: "13px 2px", cursor: "pointer", borderBottom: `1px solid ${X.line}` }}
              >
                <span style={{ flex: 1, fontSize: 14, fontWeight: 600 }}>Cancelled quotes</span>
                <span style={{ fontSize: 13, color: stF === "cancelled" ? X.tealDk : X.sub, fontWeight: stF === "cancelled" ? 700 : 500 }}>{stF === "cancelled" ? "showing" : (stCount.cancelled || 0).toLocaleString("pt-BR")}</span>
                <ChevronRight size={16} color="#A0AEC0" />
              </div>
            </div>
          </SheetField>
          {filtersExtra}
          <button onClick={exportXlsx} style={{ ...btn, height: 44, justifyContent: "center", borderRadius: 12 }}>
            <Download size={16} /> Download this list (Excel)
          </button>
        </Sheet>
      </div>
    );
  }

  const manualBg = "#EEF7F4";
  return (
    <>
      <div style={{ ...card, overflow: "hidden" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "16px 16px 12px", flexWrap: "wrap" }}>
          <div style={h2}>Quote portfolio</div>
          <span style={{ fontSize: 11, fontWeight: 600, color: X.tealDk, background: X.tealSoft, padding: "3px 10px", borderRadius: 999 }}>
            {shown.length.toLocaleString("pt-BR")} {shown.length === 1 ? "quote" : "quotes"}
          </span>
          <div style={{ display: "inline-flex", gap: 3, background: X.head, borderRadius: 9, padding: 3 }}>
            {STATUS_TABS.map(([k, l]) => (
              <button
                key={k}
                onClick={() => setStF(k)}
                style={{ border: "none", borderRadius: 7, padding: "5px 11px", fontSize: 12, fontWeight: stF === k ? 700 : 500, cursor: "pointer", background: stF === k ? "#fff" : "transparent", color: stF === k ? X.tealDk : X.sub, boxShadow: stF === k ? "0 1px 2px rgba(16,24,40,.08)" : "none" }}
              >
                {l} <span style={{ fontSize: 11, opacity: 0.75 }}>{(stCount[k] || 0).toLocaleString("pt-BR")}</span>
              </button>
            ))}
          </div>
          <span style={{ fontSize: 11.5, color: X.sub }}>FUP and % Win are filled in per quote · Thermometer is automatic (customer hit rate) · the arrow shows the details, the Activities History and the notes</span>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            {fltOn && (
              <button
                onClick={() => {
                  setFlt({});
                  setStF("");
                }}
                style={{ ...btn, height: 32, fontSize: 12 }}
              >
                <XIcon size={14} /> Clear filters
              </button>
            )}
            <div style={{ position: "relative" }}>
              <Search size={14} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: X.sub, pointerEvents: "none" }} />
              <input value={srch} onChange={(e) => setSrch(e.target.value)} placeholder="Search quote, customer or account…" style={{ ...inp, width: 250, paddingLeft: 30 }} />
            </div>
            <button onClick={exportXlsx} style={btn} title="Download this table as an Excel file">
              <Download size={15} /> Export
            </button>
          </div>
        </div>
        <div style={{ overflow: "auto", maxHeight: "min(680px, calc(120vh - 300px))" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <colgroup>
              {PCOLS.map((c) => (
                <col key={c.id} style={c.w ? { width: c.w } : undefined} />
              ))}
            </colgroup>
            <thead>
              <tr>
                {PCOLS.map((c) => {
                  const sOn = sort.id === c.id;
                  const fOn = (flt[c.id] || []).length > 0;
                  return (
                    <th
                      key={c.id}
                      title={c.tip}
                      style={{
                        ...th,
                        color: sOn || fOn ? X.tealDk : th.color,
                        ...(c.right ? { textAlign: "right" } : null),
                        ...(c.id === "id" ? { paddingLeft: 44 } : null),
                        ...(c.id === "quoted" ? { borderLeft: `1px solid ${X.line}` } : null),
                        ...(c.id === "fup" ? { borderLeft: `1px solid ${X.line}` } : null),
                        ...(c.manual ? { background: manualBg } : null),
                        ...(c.id === "sold" ? { paddingRight: 18 } : null),
                      }}
                    >
                      {c.sort ? (
                        <span
                          onClick={() => clickSort(c)}
                          title={`${c.tip ? c.tip + " · " : ""}click to sort${sOn ? " · click again to invert" : ""}`}
                          style={{ display: "inline-flex", alignItems: "center", gap: 4, cursor: "pointer" }}
                        >
                          {c.label}
                          {React.createElement(sOn ? (sort.dir === 1 ? ArrowUp : ArrowDown) : ArrowUpDown, { size: 12, color: sOn ? X.tealDk : "#B8BDC4" })}
                          {c.filter && (
                            <span onClick={(e) => e.stopPropagation()} style={{ marginLeft: 6, display: "inline-flex" }}>
                              <ColFilter iconOnly label={c.label} width={c.fw || 240} options={fltOpts(c.id)} selected={flt[c.id] || []} onChange={(v) => setFlt((f) => ({ ...f, [c.id]: v }))} />
                            </span>
                          )}
                        </span>
                      ) : c.filter ? (
                        <ColFilter label={c.label} width={c.fw || 240} options={fltOpts(c.id)} selected={flt[c.id] || []} onChange={(v) => setFlt((f) => ({ ...f, [c.id]: v }))} />
                      ) : (
                        c.label
                      )}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {vis.map((r, i) => {
                const isOpen = openQ === r.id;
                const closed = r.st !== "open";
                const tdx = { ...td, background: isOpen ? "#F1F8F6" : i % 2 ? "#FBFCFD" : "#fff" };
                const num = { ...tdN, background: tdx.background };
                const overdue = r.next && r.next < today;
                // abre a cotação no painel lateral (com as setas para ir à anterior / próxima da tabela)
                const toggle = () => onOpenQuote(r.id, false, shown.map((x) => x.id));
                const name = r.c + (r.conta ? " · " + r.conta : "");
                return (
                  <React.Fragment key={r.id}>
                    <tr className="qp-row" onClick={toggle} style={{ cursor: "pointer" }} title="Click to open the quote: details, orders and Activities / History">
                      <td style={{ ...tdx, paddingLeft: 14, whiteSpace: "nowrap", boxShadow: isOpen ? `inset 3px 0 0 ${X.teal}` : undefined }}>
                        <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ flexShrink: 0, width: 22, height: 22, borderRadius: 6, display: "inline-flex", alignItems: "center", justifyContent: "center", background: isOpen ? X.tealSoft : "#EEF2F5" }}>
                            <ChevronRight size={14} style={{ color: isOpen ? X.tealDk : "#4B5563" }} />
                          </span>
                          <span style={{ fontWeight: 700 }}>{r.id}</span>
                          <NoteBadge info={notesBy[r.id]} />
                        </span>
                      </td>
                      <td style={tdx}>
                        <span title={name} style={{ display: "block", fontWeight: 500, maxWidth: 330, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {r.c}
                        </span>
                      </td>
                      <td style={{ ...tdx, whiteSpace: "nowrap", color: X.sub }}>{fmtDate(r.d)}</td>
                      <td style={{ ...tdx, borderLeft: `1px solid ${X.line}`, whiteSpace: "nowrap" }}>
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                          {r.st === "lost" || r.st === "cancelled" ? (
                            <ClosePill
                              kind={r.st}
                              why={r.why}
                              onClick={(e) => {
                                e.stopPropagation();
                                onAskClose(r.id, r.c, r.st, r.manual ? "close" : "why");
                              }}
                              title={r.manual ? "Closed by hand (the Quote Report still says Open) · click to change the reason or reopen" : `${CLOSE_LABEL[r.st]} in the Quote Report · click to choose the reason`}
                            />
                          ) : (
                            <PickCell
                              value={r.fup}
                              options={r.st === "open" ? FUP_OPTS : FUP_PICK}
                              render={fupOptRender}
                              describe={fupOptDescribe}
                              onChange={(v) => (CLOSE_LABEL[v] ? onAskClose(r.id, r.c, v) : onMeta(r.id, { fup: v }, r.c))}
                              title={closed ? "This quote is closed: it does not appear in the follow-up calendar" : "Follow-up cadence · Lost / Cancelled closes the quote (with the reason)"}
                              width={300}
                            />
                          )}
                          {r.next && (
                            <span title={overdue ? `Follow-up overdue since ${fmtDate(r.next)}` : `Next follow-up ${fmtDate(r.next)}`} style={{ fontSize: 11, color: overdue ? X.redDk : X.sub, fontWeight: overdue ? 700 : 500 }}>
                              {overdue ? "overdue " : r.next === today ? "" : "next "}
                              {r.next === today ? "today" : fmtDay(r.next)}
                              {r.next2 ? ` & ${fmtDay(r.next2)}` : ""}
                            </span>
                          )}
                        </span>
                      </td>
                      <td style={tdx}>
                        <PickCell value={r.win} options={WIN_VALS} render={(v) => <WinPill v={v} />} onChange={(v) => onMeta(r.id, { win: v || 0 }, r.c)} title="Chance of winning this quote" width={140} />
                      </td>
                      <td style={tdx}>
                        <ThermoCell t={thermo[r.nn]} />
                      </td>
                      <td style={{ ...num, borderLeft: `1px solid ${X.line}` }}>{money(r.quoted)}</td>
                      <td style={{ ...num, color: r.open ? X.amber : undefined }}>{r.open ? money(r.open) : dash}</td>
                      <td style={{ ...num, color: r.lostV ? X.redDk : undefined }}>{r.lostV ? money(r.lostV) : dash}</td>
                      <td style={{ ...num, paddingRight: 12 }}>
                        {r.st === "won" ? <span style={{ color: X.green, fontWeight: 700, background: X.greenSoft, padding: "3px 6px", borderRadius: 6 }}>{money(r.sold)}</span> : dash}
                      </td>
                    </tr>
                  </React.Fragment>
                );
              })}
              {!shown.length && (
                <tr>
                  <td colSpan={PCOLS.length} style={{ ...td, textAlign: "center", color: X.sub, padding: "28px 10px" }}>
                    {rows.length ? "No quote matches the search or filters." : "No quote in this period and classification."}
                  </td>
                </tr>
              )}
              {shown.length > vis.length && (
                <tr>
                  <td colSpan={PCOLS.length} style={{ ...td, textAlign: "center", padding: 12 }}>
                    <button onClick={() => setAll(true)} style={{ ...btn, height: 30, fontSize: 12, margin: "0 auto" }}>
                      Show all {shown.length.toLocaleString("pt-BR")} quotes
                    </button>
                  </td>
                </tr>
              )}
            </tbody>
            {shown.length > 0 && (
              <tfoot>
                <tr>
                  {PCOLS.map((c) => {
                    const ft = { position: "sticky", bottom: 0, zIndex: 1, background: "#F3F6F8", borderTop: `2px solid ${X.line}`, padding: "11px 8px", fontSize: 12.5, fontWeight: 700, color: X.ink, fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" };
                    if (c.id === "id")
                      return (
                        <td key={c.id} style={{ ...ft, paddingLeft: 44 }} colSpan={1}>
                          Total · {shown.length.toLocaleString("pt-BR")} {shown.length === 1 ? "quote" : "quotes"}
                        </td>
                      );
                    const col = { quoted: X.ink, open: X.amber, lostV: X.redDk, sold: X.green }[c.id];
                    if (col)
                      return (
                        <td key={c.id} style={{ ...ft, textAlign: "right", color: col, ...(c.id === "quoted" ? { borderLeft: `1px solid ${X.line}` } : null), ...(c.id === "sold" ? { paddingRight: 18 } : null) }}>
                          {money(tot[c.id])}
                        </td>
                      );
                    return <td key={c.id} style={{ ...ft, ...(c.id === "fup" ? { borderLeft: `1px solid ${X.line}` } : null) }} />;
                  })}
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </>
  );
}

/* ===== campo de dinheiro (meta) ===== */
function MoneyInput({ value, onCommit }) {
  const fmt = (v) => (v ? Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "");
  const [txt, setTxt] = useState(fmt(value));
  const [focus, setFocus] = useState(false);
  useEffect(() => {
    if (!focus) setTxt(fmt(value));
  }, [value, focus]);
  const commit = () => {
    const n = Math.max(0, Math.round(toNum(txt) * 100) / 100);
    // só grava se o valor mudou (antes, clicar no campo e sair gravava 0 como meta daquele dia)
    if (n !== Math.round((Number(value) || 0) * 100) / 100) onCommit(n);
    setTxt(fmt(n));
  };
  return (
    <div style={{ display: "flex", alignItems: "center", border: `2px solid ${focus ? X.teal : "#CFE6E0"}`, borderRadius: 10, background: "#fff", height: 46, padding: "0 12px", gap: 6, minWidth: 0 }}>
      <span style={{ fontSize: 15, fontWeight: 700, color: X.sub }}>R$</span>
      <input
        value={txt}
        placeholder="0,00"
        inputMode="decimal"
        onFocus={() => setFocus(true)}
        onBlur={() => {
          setFocus(false);
          commit();
        }}
        onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
        onChange={(e) => setTxt(e.target.value)}
        style={{ border: "none", outline: "none", fontSize: 20, fontWeight: 800, color: X.ink, width: "100%", minWidth: 0, fontVariantNumeric: "tabular-nums", background: "transparent" }}
      />
    </div>
  );
}

/* ===== Calendar: meta diária → agenda de follow-up → plano do dia escolhido ===== */
const DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const DOW_LONG = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const dowName = (iso) => DOW_LONG[(dowOf(iso) + 6) % 7];
const REASON = {
  overdue: [X.redDk, X.redSoft],
  due: [X.tealDk, X.tealSoft],
  fill: ["#2F6FDB", "#E6EEFB"],
  other: ["#64748B", "#EEF2F5"],
  likely: ["#1F8A5B", "#DDF3E8"],
  hot: ["#C2410C", "#FDE6DC"],
  aging: [X.amber, X.amberSoft],
  nowin: ["#64748B", "#EEF2F5"],
};
/* Botão Done do plano e da semana: abre a cotação direto na nota + data do próximo follow-up
   (o Done fica no Activities / History e só confirma com nota e data). */
function DoneButton({ id, onOpenQuote, small }) {
  return (
    <button
      onClick={(e) => {
        e.stopPropagation();
        onOpenQuote(id, true);
      }}
      title="Mark Done in the Activities / History of the quote: write a note and choose the date of the next follow-up"
      style={
        small
          ? { ...btn, marginLeft: "auto", height: 22, padding: "0 7px", fontSize: 10.5, gap: 3, color: X.tealDk, borderColor: "#BFE3D9", borderRadius: 6 }
          : { ...btn, height: 28, padding: "0 10px", fontSize: 11.5, color: X.tealDk, borderColor: "#BFE3D9" }
      }
    >
      <Check size={small ? 11 : 13} /> Done
    </button>
  );
}
/* ===== Relatório de FUP (PDF) =====
   Follow-ups feitos (Done) no período escolhido, no recorte da tela (classificação, linha, vendedor).
   Abre uma página pronta para imprimir: no diálogo de impressão, escolha "Salvar como PDF". */
const esc = (v) =>
  String(v == null ? "" : v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
/* ===== Relatórios em Excel e para colar no e-mail =====
   copyRich: copia uma tabela HTML com estilos na própria tag (o Outlook/Gmail mantém as cores ao colar). */
/* ===== Planilha Excel com formatação (cores, bordas, negrito, R$, %) =====
   O SheetJS gratuito não grava cores; por isso aqui o .xlsx é montado à mão (XML do Office) e só a
   compactação (zip) usa o SheetJS. sheets = [{ title, cols: [larguras], rows: [[célula…]…], freeze, filter }]
   Célula: valor simples ou { v, st: { b, color, fill, fmt: "money" | "pct" | "int", h, wrap, size, border } }.
   freeze = nº de linhas fixas no topo; filter = nº da linha (1 = primeira) do cabeçalho com filtro. */
const xmlEsc = (v) =>
  String(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
const colName = (i) => {
  let s = "";
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
};
const argb = (hex) => "FF" + String(hex || "#000000").replace("#", "").toUpperCase();
function xlsxStyled(name, sheets) {
  // estilos usados (cada combinação vira um "xf")
  const fonts = ['<font><sz val="10"/><color rgb="FF1D2B36"/><name val="Calibri"/><family val="2"/></font>'];
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const idx = (arr, x) => {
    const i = arr.indexOf(x);
    return i >= 0 ? i : arr.push(x) - 1;
  };
  const xfOf = {};
  const styleId = (st) => {
    if (!st) return 0;
    const k = JSON.stringify(st);
    if (xfOf[k] != null) return xfOf[k];
    const f = idx(fonts, `<font>${st.b ? "<b/>" : ""}${st.i ? "<i/>" : ""}<sz val="${st.size || 10}"/><color rgb="${argb(st.color || "#1D2B36")}"/><name val="Calibri"/><family val="2"/></font>`);
    const fl = st.fill ? idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const bc = st.border === false ? null : argb(st.border || "#C9D3DC");
    const bd = bc ? idx(borders, `<border><left style="thin"><color rgb="${bc}"/></left><right style="thin"><color rgb="${bc}"/></right><top style="thin"><color rgb="${bc}"/></top><bottom style="thin"><color rgb="${bc}"/></bottom><diagonal/></border>`) : 0;
    const nf = st.fmt === "money" ? 164 : st.fmt === "pct" ? 9 : st.fmt === "int" ? 1 : 0;
    const al = `<alignment vertical="center"${st.h ? ` horizontal="${st.h}"` : ""}${st.wrap ? ' wrapText="1"' : ""}/>`;
    xfOf[k] = xfs.push(`<xf numFmtId="${nf}" fontId="${f}" fillId="${fl}" borderId="${bd}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`) - 1;
    return xfOf[k];
  };
  const sheetXml = (sh) => {
    const rowsXml = sh.rows
      .map((row, r) => {
        const cells = row
          .map((cell, c) => {
            const o = cell != null && typeof cell === "object" ? cell : { v: cell };
            const ref = colName(c) + (r + 1);
            const s = styleId(o.st);
            const sa = s ? ` s="${s}"` : "";
            if (o.v == null || o.v === "") return s ? `<c r="${ref}"${sa}/>` : "";
            if (typeof o.v === "number" && isFinite(o.v)) return `<c r="${ref}"${sa}><v>${o.v}</v></c>`;
            return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(o.v)}</t></is></c>`;
          })
          .join("");
        const ht = sh.heights && sh.heights[r] ? ` ht="${sh.heights[r]}" customHeight="1"` : "";
        return `<row r="${r + 1}"${ht}>${cells}</row>`;
      })
      .join("");
    const ncol = Math.max(1, ...sh.rows.map((r) => r.length));
    const pane = sh.freeze
      ? `<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="${sh.freeze}" topLeftCell="A${sh.freeze + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${sh.freeze + 1}" sqref="A${sh.freeze + 1}"/></sheetView></sheetViews>`
      : `<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>`;
    const cols = sh.cols ? `<cols>${sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const last = sh.filterTo || sh.rows.length;
    const af = sh.filter ? `<autoFilter ref="A${sh.filter}:${colName(ncol - 1)}${Math.max(sh.filter, last)}"/>` : "";
    const merges = sh.merges && sh.merges.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>${pane}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowsXml}</sheetData>${af}${merges}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  };
  const sheetFiles = sheets.map(sheetXml); // primeiro as abas (registram os estilos)
  const names = sheets.map((s, i) => xmlEsc(String(s.title || `Sheet${i + 1}`).replace(/[\\/?*[\]:]/g, " ").slice(0, 31)));
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;R$&quot;\\ #,##0.00"/></numFmts><fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills><borders count="${borders.length}">${borders.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${names.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets><definedNames>${sheets
      .map((s, i) => (s.filter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${names[i].replace(/'/g, "''")}'!$A$${s.filter}:$${colName(Math.max(1, ...s.rows.map((r) => r.length)) - 1)}$${Math.max(s.filter, s.filterTo || s.rows.length)}</definedName>` : ""))
      .join("")}</definedNames></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": styles,
  };
  sheetFiles.forEach((x, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = x));
  const zip = XLSX.CFB.utils.cfb_new();
  const enc = new TextEncoder();
  Object.entries(files).forEach(([p, x]) => XLSX.CFB.utils.cfb_add(zip, p, enc.encode(x)));
  return XLSX.CFB.write(zip, { fileType: "zip", type: "array", compression: true });
}

const xlsxSave = (name, sheets) => downloadBlob(new Blob([new Uint8Array(xlsxStyled(name, sheets))], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), name);
// estilos dos relatórios (as mesmas cores do PDF)
const XS = {
  title: { b: true, size: 14, border: false },
  sub: { color: "#64748B", border: false },
  head: (fill) => ({ b: true, color: "#FFFFFF", fill: fill || "#00BFFF", h: "center", wrap: true, border: "#FFFFFF" }),
  cell: (k, extra) => ({ ...(k % 2 ? { fill: "#F6F9FB" } : null), ...extra }),
  total: (extra) => ({ b: true, fill: "#E3E9EE", ...extra }),
};
const copyRich = (html) => {
  const box = document.createElement("div");
  box.innerHTML = html;
  box.style.cssText = "position:fixed;left:-99999px;top:0;background:#fff";
  document.body.appendChild(box);
  const range = document.createRange();
  range.selectNodeContents(box);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch {}
  sel.removeAllRanges();
  box.remove();
  return ok;
};
// tabela HTML com estilo em cada célula (para colar no e-mail)
const mailTable = (head, rows, { headBg = {}, align = {}, foot } = {}) => {
  const f = "font-family:Segoe UI,Arial,sans-serif;font-size:11px;";
  const th = (h, i) => `<th style="${f}background:${headBg[i] || "#00BFFF"};color:#fff;font-weight:bold;padding:6px 8px;border:1px solid #ffffff;text-align:center">${esc(h)}</th>`;
  const td = (v, i, extra = "") => `<td style="${f}padding:5px 8px;border:1px solid #C9D3DC;vertical-align:middle;text-align:${align[i] || "left"};${extra}">${v == null || v === "" ? "—" : esc(String(v)).replace(/\n/g, "<br>")}</td>`;
  return `<table style="border-collapse:collapse;${f}"><thead><tr>${head.map(th).join("")}</tr></thead><tbody>${rows
    .map((r, k) => `<tr>${r.map((v, i) => td(v, i, k % 2 ? "background:#F6F9FB;" : "")).join("")}</tr>`)
    .join("")}</tbody>${foot ? `<tfoot><tr>${foot.map((v, i) => td(v, i, "font-weight:bold;background:#EEF2F5;")).join("")}</tr></tfoot>` : ""}</table>`;
};
const noteTxt = (n) => (n ? `${fmtDate(String(n.t).slice(0, 10))}: ${n.txt}` : "");
const FUP_HEAD = ["Quotation", "New net amount", "Name", "Comentário antigo", "Comentário novo", "Requisição do cliente", "Termômetro", "Follow-up date", "Quoted by", "Vendedor"];
// linhas do relatório de FUP (Done) para Excel (números de verdade) e para o e-mail (texto)
const fupXlsSheet = (rows, { periodTxt, filtersTxt }) => {
  const total = rows.reduce((a, r) => a + (r.amount || 0), 0);
  const C = XS.cell;
  return {
    title: "FUP report",
    cols: [13, 15, 38, 48, 48, 22, 12, 13, 10, 22],
    freeze: 3,
    filter: 3,
    filterTo: rows.length + 3,
    merges: ["A1:J1", "A2:J2"],
    heights: { 0: 22, 2: 30 },
    rows: [
      [{ v: "Follow-up report", st: XS.title }],
      [{ v: `${periodTxt} · ${filtersTxt} · ${rows.length} follow-up${rows.length === 1 ? "" : "s"} · generated ${fmtStamp(nowStamp())}`, st: XS.sub }],
      FUP_HEAD.map((h, i) => ({ v: h, st: XS.head(i === 4 ? "#FF4081" : i === 6 ? "#008000" : null) })),
      ...rows.map((r, k) => [
        { v: r.id, st: C(k, { b: true, h: "center" }) },
        { v: r.amount || 0, st: C(k, { fmt: "money", h: "right", b: true }) },
        { v: r.c, st: C(k, { wrap: true }) },
        { v: noteTxt(r.old), st: C(k, { wrap: true, color: "#475569" }) },
        { v: noteTxt(r.nw), st: { wrap: true, fill: "#FFF1F6" } },
        { v: r.req || "", st: C(k, { wrap: true, h: "center" }) },
        { v: r.win ? r.win / 100 : "", st: C(k, { fmt: "pct", h: "center", b: true, color: "#008000" }) },
        { v: r.next ? fmtDate(r.next) : "", st: C(k, { h: "center", b: true }) },
        { v: spCode(r.qb) || sellerDisplay(r.qb), st: C(k, { h: "center" }) },
        { v: sellerDisplay(r.sp), st: C(k, { h: "center", wrap: true }) },
      ]),
      [{ v: "Total", st: XS.total({ h: "center" }) }, { v: total, st: XS.total({ fmt: "money", h: "right" }) }, { v: `${rows.length} follow-ups`, st: XS.total() }, ...Array(7).fill({ v: "", st: XS.total() })],
    ],
  };
};
const fupMailHtml = (rows) =>
  mailTable(
    FUP_HEAD,
    rows.map((r) => [r.id, money(r.amount), r.c, noteTxt(r.old), noteTxt(r.nw), r.req, r.win ? r.win + "%" : "", r.next ? fmtDate(r.next) : "", spCode(r.qb) || sellerDisplay(r.qb), sellerDisplay(r.sp)]),
    { headBg: { 4: "#FF4081", 6: "#008000" }, align: { 0: "center", 1: "right", 6: "center", 7: "center", 8: "center", 9: "center" }, foot: ["Total", money(rows.reduce((a, r) => a + (r.amount || 0), 0)), `${rows.length} follow-ups`, "", "", "", "", "", "", ""] },
  );
function fupReportHtml(rows, { title, periodTxt, filtersTxt, generated }) {
  const total = rows.reduce((a, r) => a + (r.amount || 0), 0);
  const note = (n) => (n ? `<div class="nt">${esc(n.txt).replace(/\n/g, "<br>")}</div><div class="dt">${esc(fmtDate(String(n.t).slice(0, 10)))}${n.by ? " · " + esc(n.by) : ""}</div>` : `<span class="na">—</span>`);
  const body = rows
    .map(
      (r) => `<tr>
  <td class="c b">${esc(r.id)}</td>
  <td class="r b">${esc(money(r.amount))}</td>
  <td class="c">${esc(r.c)}</td>
  <td>${note(r.old)}</td>
  <td class="new">${note(r.nw)}</td>
  <td class="c">${r.req ? esc(r.req) : '<span class="na">—</span>'}</td>
  <td class="c">${r.win ? `<b>${r.win}%</b>` : '<span class="na">—</span>'}</td>
  <td class="c b">${r.next ? esc(fmtDate(r.next)) : '<span class="na">—</span>'}<div class="dt">done ${esc(fmtDate(String(r.done).slice(0, 10)))}</div></td>
  <td class="c">${esc(spCode(r.qb) || sellerDisplay(r.qb))}</td>
  <td class="c">${esc(sellerDisplay(r.sp))}</td>
</tr>`,
    )
    .join("");
  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${esc(title)}</title>
<style>
  @page { size: A4 landscape; margin: 10mm; }
  * { box-sizing: border-box; }
  body { font-family: "Segoe UI", Arial, sans-serif; color: #1D2B36; margin: 0; padding: 16px; font-size: 10px; background: #fff; }
  .bar { display: flex; gap: 8px; align-items: center; margin-bottom: 12px; }
  .bar button { font: inherit; font-size: 13px; font-weight: 600; padding: 8px 14px; border-radius: 8px; border: 1px solid #1F9A84; background: #1F9A84; color: #fff; cursor: pointer; }
  .bar span { color: #6B7A86; font-size: 12px; }
  h1 { font-size: 17px; margin: 0 0 3px; }
  .sub { color: #475569; font-size: 11px; margin-bottom: 10px; }
  .sub b { color: #1D2B36; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; }
  thead { display: table-header-group; }
  th { background: #00BFFF; color: #fff; font-size: 10px; font-weight: 700; text-decoration: underline; padding: 7px 5px; border: 1px solid #fff; text-align: center; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  th.new { background: #FF4081; }
  th.thermo { background: #008000; }
  td { border: 1px solid #C9D3DC; padding: 6px 5px; vertical-align: middle; word-wrap: break-word; line-height: 1.35; }
  tr { break-inside: avoid; page-break-inside: avoid; }
  tbody tr:nth-child(even) td { background: #F6F9FB; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  td.new { background: #FFF1F6 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .c { text-align: center; } .r { text-align: right; } .b { font-weight: 700; }
  .nt { white-space: normal; } .dt { color: #6B7A86; font-size: 9px; margin-top: 2px; } .na { color: #A0AEC0; }
  tfoot td { font-weight: 700; background: #EEF2F5; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .empty { padding: 30px; text-align: center; color: #6B7A86; font-size: 13px; }
  @media print { .bar { display: none; } body { padding: 0; } }
</style></head><body>
<div class="bar"><button onclick="window.print()">Print / Save as PDF</button><span>In the print window, choose "Save as PDF" (or "Microsoft Print to PDF").</span></div>
<h1>${esc(title)}</h1>
<div class="sub"><b>${esc(periodTxt)}</b> · ${esc(filtersTxt)} · ${rows.length} follow-up${rows.length === 1 ? "" : "s"} · ${esc(money(total))} · generated ${esc(generated)}</div>
${
  rows.length
    ? `<table><colgroup><col style="width:8%"><col style="width:8%"><col style="width:12%"><col style="width:17%"><col style="width:17%"><col style="width:10%"><col style="width:6%"><col style="width:8%"><col style="width:6%"><col style="width:8%"></colgroup>
<thead><tr><th>Quotation</th><th>New net amount</th><th>Name</th><th>Comentário antigo</th><th class="new">Comentário novo</th><th>Requisição do cliente</th><th class="thermo">Termômetro</th><th>Follow-up date</th><th>Quoted by</th><th>Vendedor</th></tr></thead>
<tbody>${body}</tbody>
<tfoot><tr><td class="b">Total</td><td class="r">${esc(money(total))}</td><td colspan="8">${rows.length} follow-up${rows.length === 1 ? "" : "s"}</td></tr></tfoot></table>`
    : `<div class="empty">No follow-up marked Done in this period with the current filters.</div>`
}
</body></html>`;
}
function FupReportButton({ acts, quotes, meta, thermo, today, filtersTxt }) {
  const p = usePop(320);
  const mon = addDays(today, -((dowOf(today) + 6) % 7));
  const RANGES = [
    ["today", "Today", today, today],
    ["yesterday", "Previous business day", prevBusinessDay(today), prevBusinessDay(today)],
    ["week", "This week", mon, addDays(mon, 6)],
    ["lastweek", "Last week", addDays(mon, -7), addDays(mon, -1)],
    ["month", "This month", today.slice(0, 8) + "01", lastDayOfMonth(today.slice(0, 7))],
    ["custom", "Custom", "", ""],
  ];
  const [rk, setRk] = useState("today");
  const [cf, setCf] = useState(today);
  const [ct, setCt] = useState(today);
  const r = RANGES.find((x) => x[0] === rk) || RANGES[0];
  const from = rk === "custom" ? cf : r[2],
    to = rk === "custom" ? ct : r[3];
  const ok = /^\d{4}-\d{2}-\d{2}$/.test(from) && /^\d{4}-\d{2}-\d{2}$/.test(to) && from <= to;
  const rows = useMemo(() => (p.pos && ok ? fupReportRows(acts, quotes, meta, thermo, from, to) : []), [p.pos, ok, acts, quotes, meta, thermo, from, to]);
  const periodTxt = from === to ? `${dowName(from)}, ${fmtDate(from)}` : `${fmtDate(from)} – ${fmtDate(to)}`;
  const [copied, setCopied] = useState(false);
  const fileTag = from === to ? from : `${from} a ${to}`;
  const toExcel = () => {
    xlsxSave(`FUP report ${fileTag}.xlsx`, [fupXlsSheet(rows, { periodTxt, filtersTxt })]);
  };
  const toMail = () => {
    if (copyRich(fupMailHtml(rows))) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } else window.alert("The browser did not allow copying. Use Download Excel instead.");
  };
  const open = () => {
    const html = fupReportHtml(rows, { title: "Follow-up report", periodTxt, filtersTxt, generated: fmtStamp(nowStamp()) });
    const w = window.open("", "_blank");
    if (!w) {
      window.alert("The browser blocked the report window. Allow pop-ups for this page and try again.");
      return;
    }
    w.document.open();
    w.document.write(html);
    w.document.close();
    w.focus();
    setTimeout(() => {
      try {
        w.print();
      } catch {}
    }, 400);
    p.setPos(null);
  };
  return (
    <>
      <button ref={p.anchor} onClick={p.toggle} style={{ ...btn, height: 32, fontSize: 12, padding: "0 11px" }} title="Report (Excel, e-mail or PDF) of the follow-ups marked Done in a day, week or period (follows the filters on the screen)">
        <FileSpreadsheet size={14} /> FUP report
      </button>
      {p.render(
        <div style={{ padding: 8, display: "flex", flexDirection: "column", gap: 9 }}>
          <div style={{ fontSize: 13, fontWeight: 700 }}>Follow-up report</div>
          <div style={{ fontSize: 11.5, color: X.sub }}>Follow-ups marked Done in the period · {filtersTxt}</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {RANGES.map(([k, l]) => {
              const on = rk === k;
              return (
                <button key={k} onClick={() => setRk(k)} style={{ ...btn, height: 28, padding: "0 10px", fontSize: 11.5, borderRadius: 999, background: on ? X.tealSoft : "#fff", color: on ? X.tealDk : X.ink, borderColor: on ? "#9FD6C9" : X.line, fontWeight: on ? 700 : 600 }}>
                  {l}
                </button>
              );
            })}
          </div>
          {rk === "custom" && (
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input type="date" value={cf} onChange={(e) => setCf(e.target.value)} style={{ ...inp, height: 30, fontSize: 12, flex: 1 }} />
              <span style={{ color: X.sub }}>to</span>
              <input type="date" value={ct} onChange={(e) => setCt(e.target.value)} style={{ ...inp, height: 30, fontSize: 12, flex: 1 }} />
            </div>
          )}
          <div style={{ fontSize: 12, color: ok ? X.ink : X.redDk }}>{ok ? <><b>{periodTxt}</b> · {rows.length} follow-up{rows.length === 1 ? "" : "s"} done</> : "Choose a valid period"}</div>
          <button onClick={toExcel} disabled={!ok} style={{ ...btn, height: 34, justifyContent: "center", background: ok ? X.teal : "#C7D2D9", borderColor: ok ? X.teal : "#C7D2D9", color: "#fff", cursor: ok ? "pointer" : "not-allowed" }}>
            <FileSpreadsheet size={14} /> Download Excel
          </button>
          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={toMail} disabled={!ok || !rows.length} style={{ ...btn, flex: 1, height: 32, fontSize: 12, justifyContent: "center" }} title="Copy the table with its colors: paste it straight into the e-mail (Ctrl+V)">
              <Copy size={13} /> {copied ? "Copied!" : "Copy for e-mail"}
            </button>
            <button onClick={open} disabled={!ok} style={{ ...btn, flex: 1, height: 32, fontSize: 12, justifyContent: "center" }} title="Open the report ready to print / save as PDF">
              <Printer size={13} /> PDF
            </button>
          </div>
        </div>,
      )}
    </>
  );
}
/* feriado do dia no calendário: vermelho = sem expediente (nenhum FUP cai nele); âmbar = ponto facultativo */
function HolidayTag({ d, full }) {
  const h = holidayOf(d);
  if (!h) return null;
  const fg = h.block ? X.redDk : "#B45309",
    bg = h.block ? X.redSoft : "#FEF3C7";
  const kind = h.kind === "municipal" ? "municipal holiday" : h.kind === "state" ? "state holiday (SP)" : h.kind === "optional" ? "optional day off (ponto facultativo) · follow-ups allowed" : "national holiday";
  return (
    <span
      title={`${h.name} · ${kind}${h.block ? " · no follow-up on this day" : ""}`}
      style={{ display: "block", fontSize: full ? 11 : 9.5, fontWeight: 700, color: fg, background: bg, borderRadius: 5, padding: full ? "2px 7px" : "1px 4px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0, marginTop: full ? 4 : 0 }}
    >
      {h.block ? "Holiday · " : ""}
      {h.name}
    </span>
  );
}
/* ===== Analysis: card grande no centro da tela (o resto fica desfocado) =====
   Resultado das cotações (Open / Won / Lost / Cancelled), motivos de Lost e Cancelled e os Top 10 clientes e
   Product Groups com mais Lost + Cancelled. Segue os filtros da tela; dá para escolher um cliente. QTs / Value
   troca tudo entre quantidade e valor. */
const ST_ORDER = ["open", "won", "lost", "cancelled"];
const ST_COL = { open: ["#B7791F", "#FDF0D9"], won: ["#1F8A5B", "#DDF3E8"], lost: ["#C53030", "#FDE3E3"], cancelled: ["#64748B", "#E9EDF2"] };
const REASON_COL = {
  "Lead Time": "#E07B39",
  Price: "#C53030",
  "Budget Only": "#7A6AD8",
  "Project Discontinued": "#475569",
  "No Response": "#D69E2E",
  "Stalled project": "#2F6FDB",
};
const pctOf = (a, b) => (b ? a / b : 0);
const pct0 = (x) => (x > 0 && x < 0.005 ? "<1%" : `${Math.round(x * 100)}%`);
// rosca em SVG: fatias = [{ k, v, color }]
function Donut({ parts, size = 120, stroke = 18, center, sub }) {
  const tot = parts.reduce((t, p) => t + p.v, 0);
  const r = (size - stroke) / 2,
    c = 2 * Math.PI * r;
  let acc = 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="#EEF2F5" strokeWidth={stroke} />
      {tot > 0 &&
        parts
          .filter((p) => p.v > 0)
          .map((p) => {
            const len = (p.v / tot) * c;
            const el = (
              <circle key={p.k} cx={size / 2} cy={size / 2} r={r} fill="none" stroke={p.color} strokeWidth={stroke} strokeDasharray={`${Math.max(0, len - 1.5)} ${c}`} strokeDashoffset={-acc} transform={`rotate(-90 ${size / 2} ${size / 2})`}>
                <title>{`${p.k}: ${pct0(p.v / tot)}`}</title>
              </circle>
            );
            acc += len;
            return el;
          })}
      <text x="50%" y="49%" textAnchor="middle" style={{ fontSize: 20, fontWeight: 800, fill: X.ink }}>
        {center}
      </text>
      <text x="50%" y="64%" textAnchor="middle" style={{ fontSize: 10.5, fill: X.sub }}>
        {sub}
      </text>
    </svg>
  );
}
/* inline = página da gerência (aba Analysis), sem a janela por cima */
function AnalysisModal({ rows, cancelRows, meta, qIndex, groupOf, periodTxt, filtersTxt, sellerTxt, onOpenQuote, onClose, blocked, inline }) {
  const mob = useMobile();
  const [by, setBy] = useState("n"); // n = QTs · v = valor
  const [cust, setCust] = useState(""); // cliente escolhido ("" = todos)
  const [sell, setSell] = useState(""); // vendedor (Sales Person) escolhido ("" = todos)
  useEffect(() => {
    const k = (e) => e.key === "Escape" && !blocked && !inline && onClose && onClose();
    window.addEventListener("keydown", k);
    return () => window.removeEventListener("keydown", k);
  }, [onClose, blocked]);
  const base0 = useMemo(() => [...rows, ...cancelRows].map((r) => ({ ...r, why: normWhy(((meta || {})[r.id] || {}).why || ""), seller: groupOf(r.sp) || "(no salesperson)" })), [rows, cancelRows, meta, groupOf]);
  const sellers = useMemo(() => {
    const m = {};
    base0.forEach((r) => (m[r.seller] = (m[r.seller] || 0) + 1));
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0], "pt-BR"));
  }, [base0]);
  const base = useMemo(() => (sell ? base0.filter((r) => r.seller === sell) : base0), [base0, sell]);
  const customers = useMemo(() => {
    const m = {};
    base.forEach((r) => (m[r.c] = (m[r.c] || 0) + 1));
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0], "pt-BR"));
  }, [base]);
  // trocou o vendedor e o cliente escolhido não tem QTs dele: volta para todos os clientes
  useEffect(() => {
    if (cust && !customers.some(([c]) => c === cust)) setCust("");
  }, [customers, cust]);
  const all = cust ? base.filter((r) => r.c === cust) : base;
  // valor de cada QT: ganha = Sold (Booked do pedido, como a coluna Sold); as outras = valor cotado
  const money1 = (r) => (r.st === "won" ? r.sold || 0 : r.quoted || 0);
  const val = (r) => (by === "v" ? money1(r) : 1);
  const fmtX = (x) => (by === "v" ? moneyShort(x) : x.toLocaleString("pt-BR"));
  const tot = all.reduce((t, r) => t + val(r), 0);
  const st = {};
  ST_ORDER.forEach((k) => (st[k] = { n: 0, v: 0, x: 0, q: 0 }));
  all.forEach((r) => {
    const o = st[r.st] || st.open;
    o.n++;
    o.v += money1(r);
    o.x += val(r);
    o.q += r.quoted || 0;
  });
  /* Hit rate = ganho ÷ total cotado (Open + Won + Lost; as canceladas ficam fora, como no total Quoted da
     carteira). Em QTs: QTs ganhas ÷ QTs cotadas. Em Value: Sold ÷ valor cotado. */
  const quotedN = st.open.n + st.won.n + st.lost.n;
  const quotedV = st.open.q + st.won.q + st.lost.q;
  const hit = by === "v" ? (quotedV ? st.won.v / quotedV : null) : quotedN ? st.won.n / quotedN : null;
  // motivos de um tipo (lost / cancelled)
  const reasons = (kind) => {
    const list = all.filter((r) => r.st === kind);
    const m = Object.fromEntries(CLOSE_REASONS[kind].map((k) => [k, { n: 0, x: 0 }]));
    const none = [];
    list.forEach((r) => {
      if (m[r.why]) {
        m[r.why].n++;
        m[r.why].x += val(r);
      } else none.push(r.id);
    });
    const t = Object.values(m).reduce((s, o) => s + o.x, 0);
    return { m, t, n: list.length, none };
  };
  // Top 10: os 10 com mais lost + cancelled; a barra e o total contam todas as situações (open, won, lost, cancelled)
  const topBy = (keysOf) => {
    const g = {};
    all.forEach((r) => {
      keysOf(r).forEach(([k, w]) => {
        const o = g[k] || (g[k] = { k, open: 0, won: 0, lost: 0, cancelled: 0, why: {} });
        const x = by === "v" ? w : 1;
        o[r.st] = (o[r.st] || 0) + x;
        if ((r.st === "lost" || r.st === "cancelled") && r.why) o.why[r.why] = (o.why[r.why] || 0) + 1;
      });
    });
    return Object.values(g)
      .map((o) => ({ ...o, lc: o.lost + o.cancelled, t: o.open + o.won + o.lost + o.cancelled }))
      .filter((o) => o.lc > 0)
      .sort((a, b) => b.lc - a.lc || b.t - a.t)
      .slice(0, 10);
  };
  const topCust = topBy((r) => [[r.c, money1(r)]]);
  const topPg = topBy((r) => {
    const q = qIndex.get(r.id) || {};
    // valor dos itens do grupo; na ganha, proporcional ao Sold
    const f = r.quoted ? money1(r) / r.quoted : 0;
    return q.pg && Object.keys(q.pg).length ? Object.entries(q.pg).map(([k, v]) => [k, v * f]) : [["(no Product Group)", money1(r)]];
  });
  // cliente escolhido: todas as QTs perdidas / canceladas dele (filtro Lost / Cancelled), maiores primeiro
  const [lcF, setLcF] = useState("");
  const lcAll = cust ? all.filter((r) => r.st === "lost" || r.st === "cancelled").sort((a, b) => (b.quoted || 0) - (a.quoted || 0)) : [];
  const lostQts = lcF ? lcAll.filter((r) => r.st === lcF) : lcAll;
  const lcMax = Math.max(0, ...lostQts.map((r) => r.quoted || 0));

  const panel = { border: `1px solid ${X.line}`, borderRadius: 14, padding: 16, background: "#fff", minWidth: 0, display: "flex", flexDirection: "column", gap: 12 };
  const ttl = { fontSize: 13.5, fontWeight: 800, color: X.ink, display: "flex", alignItems: "center", gap: 8 };
  const muted = { fontSize: 11.5, color: X.sub, fontWeight: 500 };
  const seg = (opts, cur, set) => (
    <div style={{ display: "inline-flex", gap: 3, background: X.head, borderRadius: 9, padding: 3 }}>
      {opts.map(([k, l]) => (
        <button key={k} onClick={() => set(k)} style={{ border: "none", borderRadius: 7, padding: "5px 12px", fontSize: 12, fontWeight: cur === k ? 700 : 500, cursor: "pointer", background: cur === k ? "#fff" : "transparent", color: cur === k ? X.tealDk : X.sub, boxShadow: cur === k ? "0 1px 2px rgba(16,24,40,.08)" : "none" }}>
          {l}
        </button>
      ))}
    </div>
  );
  const reasonPanel = (kind) => {
    const R = reasons(kind);
    const nR = R.n - R.none.length;
    return (
      <div style={panel}>
        <div style={ttl}>
          <span style={{ width: 10, height: 10, borderRadius: 3, background: ST_COL[kind][0] }} /> {CLOSE_LABEL[kind]} · reasons
          <span style={{ ...muted, marginLeft: "auto" }}>{R.n.toLocaleString("pt-BR")} QTs</span>
        </div>
        <div style={{ display: "flex", gap: 18, alignItems: "center", flexWrap: mob ? "wrap" : "nowrap", justifyContent: "center" }}>
          <Donut parts={CLOSE_REASONS[kind].map((k) => ({ k, v: R.m[k].x, color: REASON_COL[k] || "#94A3B8" }))} center={nR.toLocaleString("pt-BR")} sub="with reason" />
          <div style={{ flex: "1 1 200px", minWidth: 0, display: "flex", flexDirection: "column", gap: 8 }}>
            {CLOSE_REASONS[kind].map((k) => {
              const sh = pctOf(R.m[k].x, R.t);
              return (
                <div key={k}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: REASON_COL[k] || "#94A3B8", flexShrink: 0 }} />
                    <span style={{ fontWeight: 600 }}>{k}</span>
                    <span style={{ ...muted, marginLeft: "auto" }}>{R.m[k].n}</span>
                    <b style={{ width: 40, textAlign: "right" }}>{pct0(sh)}</b>
                  </div>
                  <div style={{ height: 5, borderRadius: 4, background: X.track, marginTop: 4, overflow: "hidden" }}>
                    <div style={{ height: 5, width: sh * 100 + "%", background: REASON_COL[k] || "#94A3B8", borderRadius: 4 }} />
                  </div>
                </div>
              );
            })}
            {R.none.length > 0 && (
              <span onClick={() => onOpenQuote(R.none[0], false, R.none)} title="Open them to choose the reason" style={{ alignSelf: "flex-start", fontSize: 11.5, fontWeight: 600, color: X.redDk, background: X.redSoft, borderRadius: 999, padding: "2px 9px", cursor: "pointer" }}>
                {R.none.length.toLocaleString("pt-BR")} without reason ›
              </span>
            )}
          </div>
        </div>
      </div>
    );
  };
  const topWhy = (o) => {
    const e = Object.entries(o).sort((a, b) => b[1] - a[1])[0];
    return e ? e[0] : "";
  };
  const topPanel = (title, list, onPick) => {
    const max = Math.max(0, ...list.map((o) => o.t));
    return (
      <div style={panel}>
        <div style={{ ...ttl, flexWrap: "wrap", rowGap: 4 }}>
          {title}
          <span style={{ ...muted, marginLeft: "auto", display: "inline-flex", gap: 9, flexWrap: "wrap" }}>
            {ST_ORDER.map((k) => (
              <span key={k} style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                <span style={{ width: 8, height: 8, borderRadius: 2, background: ST_COL[k][0] }} />
                {STATE_LABEL[k]}
              </span>
            ))}
          </span>
        </div>
        {!list.length && <div style={muted}>Nothing lost or cancelled.</div>}
        <div style={{ display: "flex", flexDirection: "column" }}>
          {list.map((o, i) => (
            <div
              key={o.k}
              onClick={onPick ? () => onPick(o.k) : undefined}
              title={onPick ? "Show only this customer" : undefined}
              className={onPick ? "qp-opt" : undefined}
              style={{ display: "grid", gridTemplateColumns: "22px minmax(0, 1fr) auto", alignItems: "center", columnGap: 12, padding: "8px 6px", borderRadius: 8, borderTop: i ? `1px solid #F0F3F6` : "none", cursor: onPick ? "pointer" : "default" }}
            >
              <span style={{ width: 22, height: 22, borderRadius: 999, background: i < 3 ? "#FDE3E3" : X.head, color: i < 3 ? ST_COL.lost[0] : X.sub, fontSize: 11, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>{i + 1}</span>
              <span style={{ minWidth: 0, display: "flex", flexDirection: "column", gap: 5 }}>
                <span style={{ display: "flex", alignItems: "baseline", gap: 6, minWidth: 0 }}>
                  <span style={{ fontSize: 12.5, fontWeight: 700, color: X.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={o.k}>
                    {o.k}
                  </span>
                  {topWhy(o.why) && <span style={{ fontSize: 10.5, color: X.sub, whiteSpace: "nowrap" }}>· {topWhy(o.why)}</span>}
                </span>
                <span style={{ display: "flex", height: 7, borderRadius: 4, overflow: "hidden", background: X.track }} title={ST_ORDER.map((k) => `${STATE_LABEL[k]}: ${fmtX(o[k])}`).join(" · ")}>
                  {ST_ORDER.map((k) => (o[k] > 0 ? <span key={k} style={{ width: pctOf(o[k], max) * 100 + "%", background: ST_COL[k][0] }} /> : null))}
                </span>
              </span>
              <span style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap", lineHeight: 1.3 }}>
                <b style={{ display: "block", fontSize: 13, color: ST_COL.lost[0] }}>{fmtX(o.lc)}</b>
                <span style={{ fontSize: 10.5, color: X.sub }}>
                  {pct0(pctOf(o.lc, o.t))} of {fmtX(o.t)}
                </span>
              </span>
            </div>
          ))}
        </div>
      </div>
    );
  };
  const body = (
      <div onClick={(e) => e.stopPropagation()} style={inline ? { ...card, overflow: "hidden", background: X.bg, display: "flex", flexDirection: "column" } : { width: "min(1160px, 96vw)", maxHeight: "92vh", overflowY: "auto", overflowX: "hidden", background: X.bg, borderRadius: 18, boxShadow: "0 30px 80px rgba(15,23,42,.35)", display: "flex", flexDirection: "column" }}>
        {/* cabeçalho */}
        <div style={{ position: inline ? "static" : "sticky", top: 0, zIndex: 2, background: "#fff", borderBottom: `1px solid ${X.line}`, padding: mob ? "12px 12px" : "14px 18px", display: "flex", alignItems: "center", gap: mob ? 8 : 12, flexWrap: "wrap", borderRadius: inline ? 0 : "18px 18px 0 0" }}>
          {/* celular: o topo já diz "Analysis"; fica só a linha do recorte */}
          {!mob && (
          <span style={{ width: 36, height: 36, borderRadius: 10, background: X.tealSoft, color: X.tealDk, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <PieChart size={19} />
          </span>
          )}
          <div style={{ minWidth: 0 }}>
            {!mob && <div style={{ fontSize: 16.5, fontWeight: 800, color: X.ink }}>Quote analysis</div>}
            <div style={{ ...muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: mob ? "normal" : "nowrap" }}>
              {periodTxt} · {sell || sellerTxt} · {filtersTxt}
            </div>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", ...(mob ? { width: "100%" } : null) }}>
            <select value={sell} onChange={(e) => setSell(e.target.value)} style={{ ...inp, height: 34, maxWidth: mob ? "none" : 220, flex: mob ? "1 1 140px" : "none", minWidth: 0, fontWeight: sell ? 700 : 500, borderColor: sell ? "#9FD6C9" : X.line, background: sell ? X.tealSoft : "#fff" }} title="Salesperson (Sales Person)">
              <option value="">All salespeople</option>
              {sellers.map(([c, n]) => (
                <option key={c} value={c}>
                  {c} ({n})
                </option>
              ))}
            </select>
            <select value={cust} onChange={(e) => setCust(e.target.value)} style={{ ...inp, height: 34, maxWidth: mob ? "none" : 280, flex: mob ? "1 1 140px" : "none", minWidth: 0, fontWeight: cust ? 700 : 500, borderColor: cust ? "#9FD6C9" : X.line, background: cust ? X.tealSoft : "#fff" }} title="Customer">
              <option value="">All customers</option>
              {customers.map(([c, n]) => (
                <option key={c} value={c}>
                  {c} ({n})
                </option>
              ))}
            </select>
            {seg(
              [
                ["n", "QTs"],
                ["v", "Value"],
              ],
              by,
              setBy,
            )}
            {!inline && (
              <button onClick={onClose} style={{ ...btn, padding: "0 10px", height: 34 }} title="Close (Esc)">
                <XIcon size={16} />
              </button>
            )}
          </div>
        </div>

        <div style={{ padding: mob ? 10 : 18, display: "flex", flexDirection: "column", gap: mob ? 10 : 14 }}>
          {/* resultado: as 4 situações + o hit rate; embaixo, a barra com a divisão */}
          <div style={{ ...panel, gap: 14 }}>
            <div style={{ display: "grid", gridTemplateColumns: mob ? "repeat(2, minmax(0, 1fr))" : "repeat(auto-fit, minmax(170px, 1fr))", gap: mob ? 8 : 12 }}>
              {ST_ORDER.map((k) => {
                const o = st[k];
                const [fg, bg] = ST_COL[k];
                return (
                  <div key={k} style={{ borderRadius: 12, background: bg, padding: "12px 14px", minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 800, color: fg }}>{STATE_LABEL[k]}</div>
                    <div style={{ fontSize: mob ? 24 : 30, fontWeight: 800, color: fg, letterSpacing: -0.8, lineHeight: 1.15 }}>{pct0(pctOf(o.x, tot))}</div>
                    <div style={{ fontSize: 12, color: X.ink, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      {o.n.toLocaleString("pt-BR")} QTs · {moneyShort(o.v)}
                    </div>
                  </div>
                );
              })}
              <div style={{ borderRadius: 12, border: `2px solid ${ST_COL.won[1]}`, background: "#fff", padding: "10px 14px", display: "flex", alignItems: "center", gap: 12, minWidth: 0, ...(mob ? { gridColumn: "1 / -1" } : null) }}>
                <Donut parts={[{ k: "Hit rate", v: hit || 0, color: ST_COL.won[0] }, { k: "rest", v: 1 - (hit || 0), color: "#EEF2F5" }]} size={64} stroke={9} center="" sub="" />
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 800, color: X.tealDk }}>Hit rate</div>
                  <div style={{ fontSize: 28, fontWeight: 800, color: ST_COL.won[0], letterSpacing: -0.8, lineHeight: 1.15 }}>{hit != null ? pct0(hit) : "—"}</div>
                  <div style={{ fontSize: 11.5, color: X.sub }}>won ÷ quoted</div>
                </div>
              </div>
            </div>
            <div style={{ display: "flex", height: 10, borderRadius: 6, overflow: "hidden", background: X.track }}>
              {ST_ORDER.map((k) => (st[k].x > 0 ? <div key={k} title={`${STATE_LABEL[k]} ${pct0(pctOf(st[k].x, tot))}`} style={{ width: pctOf(st[k].x, tot) * 100 + "%", background: ST_COL[k][0] }} /> : null))}
            </div>
          </div>

          {/* motivos */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(380px, 100%), 1fr))", gap: 14 }}>
            {reasonPanel("lost")}
            {reasonPanel("cancelled")}
          </div>

          {/* top 10 */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(380px, 100%), 1fr))", gap: 14 }}>
            {cust ? (
              <div style={{ ...panel, gap: 10 }}>
                <div style={{ ...ttl, flexWrap: "wrap", rowGap: 6 }}>
                  <span style={{ minWidth: 0 }}>
                    Lost & cancelled QTs
                    <span style={{ display: "block", ...muted }}>
                      {lostQts.length.toLocaleString("pt-BR")} QTs · {moneyShort(lostQts.reduce((t, r) => t + (r.quoted || 0), 0))}
                    </span>
                  </span>
                  <span style={{ marginLeft: "auto", display: "inline-flex", alignItems: "center", gap: 8 }}>
                    {seg(
                      [
                        ["", `All ${lcAll.length}`],
                        ["lost", `Lost ${lcAll.filter((r) => r.st === "lost").length}`],
                        ["cancelled", `Cancelled ${lcAll.filter((r) => r.st === "cancelled").length}`],
                      ],
                      lcF,
                      setLcF,
                    )}
                    <button onClick={() => setCust("")} title="Back to all customers" style={{ ...btn, height: 28, padding: "0 8px", fontSize: 11.5 }}>
                      <XIcon size={12} /> Customer
                    </button>
                  </span>
                </div>
                {!lostQts.length ? (
                  <div style={muted}>Nothing lost or cancelled.</div>
                ) : (
                  <div style={{ maxHeight: 360, overflow: "auto", margin: "0 -6px" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "112px 70px 84px minmax(0, 1fr) 96px", minWidth: 470, gap: 10, padding: "0 6px 6px", fontSize: 10.5, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.3, position: "sticky", top: 0, background: "#fff", zIndex: 1, borderBottom: `1px solid ${X.line}` }}>
                      <span>Quotation</span>
                      <span>Date</span>
                      <span>Status</span>
                      <span>Reason</span>
                      <span style={{ textAlign: "right" }}>Quoted</span>
                    </div>
                    {lostQts.map((r) => (
                      <div
                        key={r.id}
                        onClick={() => onOpenQuote(r.id, false, lostQts.map((x) => x.id))}
                        className="qp-opt"
                        title="Open the quote"
                        style={{ display: "grid", gridTemplateColumns: "112px 70px 84px minmax(0, 1fr) 96px", minWidth: 470, alignItems: "center", gap: 10, fontSize: 12.5, cursor: "pointer", padding: "7px 6px", borderBottom: `1px solid #F0F3F6` }}
                      >
                        <span style={{ fontWeight: 700, fontSize: 12, color: X.tealDk, background: X.tealSoft, borderRadius: 6, padding: "2px 7px", justifySelf: "start" }}>{r.id}</span>
                        <span style={{ color: X.sub, fontSize: 12 }}>{r.d ? `${r.d.slice(8, 10)}/${r.d.slice(5, 7)}/${r.d.slice(2, 4)}` : "—"}</span>
                        <span style={{ fontSize: 11, fontWeight: 700, color: ST_COL[r.st][0], background: ST_COL[r.st][1], padding: "2px 8px", borderRadius: 999, justifySelf: "start" }}>{STATE_LABEL[r.st]}</span>
                        {r.why ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                            <span style={{ width: 8, height: 8, borderRadius: 2, background: REASON_COL[r.why] || "#94A3B8", flexShrink: 0 }} />
                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 600 }}>{r.why}</span>
                          </span>
                        ) : (
                          <span style={{ justifySelf: "start", fontSize: 11, fontWeight: 600, color: X.redDk, border: `1px dashed ${X.redDk}`, borderRadius: 999, padding: "1px 8px" }}>set reason ›</span>
                        )}
                        <span style={{ textAlign: "right" }}>
                          <b style={{ display: "block", fontVariantNumeric: "tabular-nums" }}>{moneyShort(r.quoted)}</b>
                          <span style={{ display: "block", height: 4, borderRadius: 3, background: X.track, marginTop: 3, overflow: "hidden" }}>
                            <span style={{ display: "block", height: 4, marginLeft: "auto", width: pctOf(r.quoted || 0, lcMax) * 100 + "%", background: ST_COL[r.st][0], borderRadius: 3 }} />
                          </span>
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ) : (
              topPanel("Top 10 customers · lost + cancelled", topCust, setCust)
            )}
            {topPanel("Top 10 product groups · lost + cancelled", topPg)}
          </div>
        </div>
      </div>
  );
  if (inline) return body;
  return createPortal(
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 1400, background: "rgba(15,23,42,.32)", backdropFilter: "blur(6px)", WebkitBackdropFilter: "blur(6px)", display: "flex", alignItems: "center", justifyContent: "center", padding: mob ? 0 : 20 }}>
      {body}
    </div>,
    document.body,
  );
}
/* ===== Relatório de FUP por vendedor (aba Salespeople) =====
   Uma linha por Responsável (Sales Person): cotações com follow-up no período, quantas foram atualizadas
   (Done), quantas depois das 11h e quantas faltam. A seta abre os detalhes do vendedor: linhas em que ele
   atua, Product Groups que ele cota e a lista dos follow-ups do período. */
const periodRanges = (today) => {
  const mon = addDays(today, -((dowOf(today) + 6) % 7));
  const prev = prevBusinessDay(today);
  return [
    ["today", "Today", today, today],
    ["yesterday", "Previous business day", prev, prev],
    ["week", "This week", mon, addDays(mon, 6)],
    ["lastweek", "Last week", addDays(mon, -7), addDays(mon, -1)],
    ["month", "This month", today.slice(0, 8) + "01", lastDayOfMonth(today.slice(0, 7))],
    ["custom", "Custom", "", ""],
  ];
};
const pctTxt = (x) => (x > 0 && x < 0.005 ? "<1%" : `${Math.round(x * 100)}%`);
function PctBar({ v, color }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 8, justifyContent: "flex-end", width: "100%" }}>
      <span style={{ width: 70, height: 6, borderRadius: 4, background: X.track, overflow: "hidden", display: "inline-block" }}>
        <span style={{ display: "block", height: 6, width: Math.round(v * 100) + "%", background: color, borderRadius: 4 }} />
      </span>
      <span style={{ fontWeight: 700, minWidth: 38, textAlign: "right" }}>{pctTxt(v)}</span>
    </span>
  );
}
function SellerReport({ quotes, periodRows, meta, acts, groupOf, today, filtersTxt, periodTxtMain, onOpenQuote }) {
  const RANGES = periodRanges(today);
  const [rk, setRk] = useState("today");
  const [cf, setCf] = useState(today);
  const [ct, setCt] = useState(today);
  const [open, setOpen] = useState("");
  const r = RANGES.find((x) => x[0] === rk) || RANGES[0];
  const from = rk === "custom" ? cf : r[2],
    to = rk === "custom" ? ct : r[3];
  const ok = /^\d{4}-\d{2}-\d{2}$/.test(from) && /^\d{4}-\d{2}-\d{2}$/.test(to) && from <= to;
  const periodTxt = !ok ? "" : from === to ? `${dowName(from)}, ${fmtDate(from)}` : `${fmtDate(from)} – ${fmtDate(to)}`;
  const rep = useMemo(() => (ok ? sellerFupReport(quotes, meta, acts, from, to, groupOf) : []), [ok, quotes, meta, acts, from, to, groupOf]);
  const tot = rep.reduce((t, x) => ({ total: t.total + x.total, upd: t.upd + x.upd, late: t.late + x.late, pend: t.pend + x.pend }), { total: 0, upd: 0, late: 0, pend: 0 });
  // detalhes: linhas e Product Groups das cotações do vendedor no período da tela (filtro de período do topo)
  const qmap = useMemo(() => new Map((quotes || []).map((q) => [q.id, q])), [quotes]);
  const detail = useMemo(() => {
    if (!open) return null;
    const mine = (periodRows || []).filter((x) => groupOf(x.sp) === open).map((x) => qmap.get(x.id)).filter(Boolean);
    const lines = { PT: 0, CP: 0, MH: 0 },
      lineN = { PT: 0, CP: 0, MH: 0 };
    const pgs = {};
    let withPg = 0;
    mine.forEach((q) => {
      ["PT", "CP", "MH"].forEach((k) => {
        if (q.al && q.al[k]) {
          lines[k] += q.al[k];
          lineN[k]++;
        }
      });
      if (q.pg && Object.keys(q.pg).length) {
        withPg++;
        Object.entries(q.pg).forEach(([k, v]) => {
          const o = pgs[k] || (pgs[k] = { v: 0, n: 0 });
          o.v += v;
          o.n++;
        });
      }
    });
    const lt = lines.PT + lines.CP + lines.MH;
    const pgT = Object.values(pgs).reduce((t, o) => t + o.v, 0);
    return { n: mine.length, lines, lineN, lt, pgs: Object.entries(pgs).sort((a, b) => b[1].v - a[1].v), pgT, withPg };
  }, [open, periodRows, qmap, groupOf]);

  const SR_HEAD = ["Responsável", "Qtd Cotações", "Qtd Atualizada", `Qtd Atualizada após ${LATE_HOUR}h`, "% Atualizada", "Qtd para atualizar", "% Faltou atualizar"];
  const srRows = rep.map((x) => [x.name || "(no salesperson)", x.total, x.upd, x.late, x.pctUpd, x.pend, x.pctPend]);
  const srTot = ["Total", tot.total, tot.upd, tot.late, tot.total ? tot.upd / tot.total : 0, tot.pend, tot.total ? tot.pend / tot.total : 0];
  const [copied, setCopied] = useState(false);
  const toExcel = () => {
    const C = XS.cell;
    const main = {
      title: "FUP by salesperson",
      cols: [30, 13, 15, 22, 13, 18, 18],
      freeze: 3,
      filter: 3,
      filterTo: rep.length + 3,
      merges: ["A1:G1", "A2:G2"],
      heights: { 0: 22, 2: 30 },
      rows: [
        [{ v: "FUP by salesperson", st: XS.title }],
        [{ v: `${periodTxt} · ${filtersTxt} · after ${LATE_HOUR}h = first Done of the period at ${LATE_HOUR}:00 or later`, st: XS.sub }],
        SR_HEAD.map((h) => ({ v: h, st: XS.head() })),
        ...srRows.map((r, k) => [
          { v: r[0], st: C(k, { b: true }) },
          { v: r[1], st: C(k, { h: "right", b: true }) },
          { v: r[2], st: C(k, { h: "right", color: "#1F8A5B" }) },
          { v: r[3], st: C(k, { h: "right", color: r[3] ? "#B45309" : "#94A3B8" }) },
          { v: r[4], st: { fmt: "pct", h: "center", b: true, color: "#1F8A5B", fill: r[4] >= 0.8 ? "#DDF3E8" : r[4] >= 0.5 ? "#FDF0D9" : "#FDE3E3" } },
          { v: r[5], st: C(k, { h: "right", color: r[5] ? "#C53030" : "#94A3B8" }) },
          { v: r[6], st: { fmt: "pct", h: "center", b: true, color: "#C53030", fill: r[6] <= 0.2 ? "#DDF3E8" : r[6] <= 0.5 ? "#FDF0D9" : "#FDE3E3" } },
        ]),
        srTot.map((v, i) => ({ v, st: XS.total({ ...(i === 4 || i === 6 ? { fmt: "pct", h: "center" } : i ? { h: "right" } : null) }) })),
      ],
    };
    const det = [];
    rep.forEach((x) =>
      [...x.items]
        .sort((a, b) => (a.kind === b.kind ? 0 : a.kind === "pend" ? -1 : 1))
        .forEach((it) => det.push({ x, it })),
    );
    const stOf = (it) => (it.kind === "upd" ? (it.late ? ["Updated after " + LATE_HOUR + "h", "#B45309", "#FDF0D9"] : ["Updated", "#1F8A5B", "#DDF3E8"]) : it.overdue ? ["Overdue", "#C53030", "#FDE3E3"] : ["To update", "#C53030", "#FFF1F2"]);
    const details = {
      title: "Details",
      cols: [26, 13, 40, 14, 18, 17, 14, 8],
      freeze: 3,
      filter: 3,
      filterTo: det.length + 3,
      merges: ["A1:H1", "A2:H2"],
      heights: { 0: 22, 2: 30 },
      rows: [
        [{ v: "FUP by salesperson · quotes", st: XS.title }],
        [{ v: `${periodTxt} · ${filtersTxt}`, st: XS.sub }],
        ["Responsável", "Quotation", "Customer", "Amount", "Status", "Updated at", "Follow-up date", "% Win"].map((h) => ({ v: h, st: XS.head() })),
        ...det.map(({ x, it }, k) => {
          const [t, fg, bg] = stOf(it);
          return [
            { v: x.name || "(no salesperson)", st: C(k) },
            { v: it.id, st: C(k, { b: true, h: "center" }) },
            { v: it.c, st: C(k) },
            { v: it.a || 0, st: C(k, { fmt: "money", h: "right" }) },
            { v: t, st: { b: true, h: "center", color: fg, fill: bg } },
            { v: it.kind === "upd" ? `${fmtDate(it.t.slice(0, 10))} ${it.t.slice(11, 16)}` : "", st: C(k, { h: "center" }) },
            { v: it.kind === "pend" ? fmtDate(it.dl) : it.next && /^\d{4}/.test(it.next) ? fmtDate(it.next) : "", st: C(k, { h: "center" }) },
            { v: it.win ? it.win / 100 : "", st: C(k, { fmt: "pct", h: "center" }) },
          ];
        }),
      ],
    };
    xlsxSave(`FUP by salesperson ${from === to ? from : `${from} a ${to}`}.xlsx`, [main, details]);
  };
  const toMail = () => {
    const p = (v) => pctTxt(v);
    const html = mailTable(
      SR_HEAD,
      srRows.map((r) => [r[0], r[1], r[2], r[3], p(r[4]), r[5], p(r[6])]),
      { align: { 1: "right", 2: "right", 3: "right", 4: "right", 5: "right", 6: "right" }, foot: [srTot[0], srTot[1], srTot[2], srTot[3], p(srTot[4]), srTot[5], p(srTot[6])] },
    );
    if (copyRich(`<div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;margin-bottom:6px"><b>FUP by salesperson · ${esc(periodTxt)}</b></div>` + html)) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } else window.alert("The browser did not allow copying. Use Excel instead.");
  };
  const printPdf = () => {
    const rowsHtml = rep
      .map(
        (x) =>
          `<tr><td>${esc(x.name || "(no salesperson)")}</td><td class="r">${x.total}</td><td class="r">${x.upd}</td><td class="r">${x.late}</td><td class="r b">${pctTxt(x.pctUpd)}</td><td class="r">${x.pend}</td><td class="r b">${pctTxt(x.pctPend)}</td></tr>`,
      )
      .join("");
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>FUP by salesperson</title><style>
@page { size: A4 landscape; margin: 12mm; } body { font-family: "Segoe UI", Arial, sans-serif; color: #1D2B36; font-size: 11px; margin: 0; padding: 16px; }
.bar { margin-bottom: 12px; } .bar button { font: inherit; font-size: 13px; font-weight: 600; padding: 8px 14px; border-radius: 8px; border: 1px solid #1F9A84; background: #1F9A84; color: #fff; cursor: pointer; }
h1 { font-size: 17px; margin: 0 0 3px; } .sub { color: #475569; margin-bottom: 10px; }
table { width: 100%; border-collapse: collapse; } th { background: #00BFFF; color: #fff; padding: 7px 6px; border: 1px solid #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
td { border: 1px solid #C9D3DC; padding: 6px; } .r { text-align: right; } .b { font-weight: 700; } tfoot td { font-weight: 700; background: #EEF2F5; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
@media print { .bar { display: none; } body { padding: 0; } }</style></head><body>
<div class="bar"><button onclick="window.print()">Print / Save as PDF</button></div>
<h1>FUP by salesperson</h1><div class="sub"><b>${esc(periodTxt)}</b> · ${esc(filtersTxt)} · generated ${esc(fmtStamp(nowStamp()))} · after ${LATE_HOUR}h = first Done of the period at ${LATE_HOUR}:00 or later</div>
<table><thead><tr><th>Responsável</th><th>Qtd Cotações</th><th>Qtd Atualizada</th><th>Qtd Atualizada após ${LATE_HOUR}h</th><th>% Atualizada</th><th>Qtd para atualizar</th><th>% Faltou atualizar</th></tr></thead>
<tbody>${rowsHtml}</tbody><tfoot><tr><td>Total</td><td class="r">${tot.total}</td><td class="r">${tot.upd}</td><td class="r">${tot.late}</td><td class="r">${tot.total ? pctTxt(tot.upd / tot.total) : "—"}</td><td class="r">${tot.pend}</td><td class="r">${tot.total ? pctTxt(tot.pend / tot.total) : "—"}</td></tr></tfoot></table></body></html>`;
    const w = window.open("", "_blank");
    if (!w) return window.alert("The browser blocked the report window. Allow pop-ups for this page and try again.");
    w.document.open();
    w.document.write(html);
    w.document.close();
    w.focus();
    setTimeout(() => {
      try {
        w.print();
      } catch {}
    }, 400);
  };

  const th = { padding: "9px 12px", fontSize: 10.5, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.3, background: X.head, borderBottom: `1px solid ${X.line}`, textAlign: "right", whiteSpace: "nowrap" };
  const td = { padding: "10px 12px", fontSize: 13, borderBottom: `1px solid ${X.line}`, textAlign: "right", fontVariantNumeric: "tabular-nums" };
  // detalhes do vendedor: três quadros brancos com cabeçalho de tabela
  const box = { border: `1px solid ${X.line}`, borderRadius: 12, background: "#fff", padding: "14px 16px", display: "flex", flexDirection: "column", gap: 10, minWidth: 0, boxShadow: "0 1px 2px rgba(16,24,40,.04)" };
  const head = { fontSize: 12.5, fontWeight: 800, color: X.ink, textTransform: "uppercase", letterSpacing: 0.3, display: "flex", alignItems: "center", gap: 9, paddingBottom: 4 };
  const headSub = { fontWeight: 500, color: X.sub, textTransform: "none", letterSpacing: 0 };
  const colH = { fontSize: 11, fontWeight: 600, color: X.sub, padding: "0 0 7px", borderBottom: `1px solid ${X.line}` };
  const rowB = { borderBottom: `1px solid #EEF2F5` };
  const MiniBar = ({ v, color, w = 80, h = 6 }) => (
    <span style={{ display: "inline-block", width: w, height: h, borderRadius: 4, background: X.track, overflow: "hidden", verticalAlign: "middle" }}>
      <span style={{ display: "block", height: h, width: Math.max(v > 0 ? 3 : 0, Math.min(1, v) * 100) + "%", background: color, borderRadius: 4 }} />
    </span>
  );
  /* detalhes do vendedor (linhas, Product Groups e os follow-ups do período): na tabela e nos cartões do celular */
  const mob = useMobile();
  const [share, setShare] = useState(false); // celular: Excel / e-mail / PDF numa folha
  // celular: ao abrir um vendedor, só a lista dos follow-ups dele (a pendência primeiro); linhas e grupos ficam no computador
  const fupSort = (a, b) => (a.kind === b.kind ? (a.kind === "upd" ? (a.t < b.t ? -1 : 1) : a.dl < b.dl ? -1 : 1) : a.kind === "pend" ? -1 : 1);
  const mobFups = (x) => (
    <div>
      {[...x.items].sort(fupSort).map((it) => {
        const [txt, fg] = it.kind === "upd" ? [`Updated ${it.t.slice(11, 16)}${from !== to ? " · " + fmtDay(it.t.slice(0, 10)) : ""}`, it.late ? X.amber : X.green] : [it.overdue ? `Overdue · ${fmtDay(it.dl)}` : "To update", X.redDk];
        return (
          <div
            key={it.id}
            className="qp-opt"
            onClick={(e) => {
              e.stopPropagation();
              onOpenQuote(it.id, false, [...x.items].sort(fupSort).map((y) => y.id));
            }}
            style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 4px", borderTop: `1px solid ${X.line}`, cursor: "pointer", minWidth: 0, borderRadius: 0 }}
          >
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 13, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{it.c}</div>
              <div style={{ fontSize: 11.5, color: X.sub, marginTop: 2 }}>
                {it.id} · {moneyShort(it.a)}
              </div>
            </div>
            <span style={{ fontSize: 11.5, fontWeight: 700, color: fg, whiteSpace: "nowrap" }}>{txt}</span>
          </div>
        );
      })}
    </div>
  );
  const detailBlock = (x) => (
      <div style={{ display: "grid", gridTemplateColumns: mob ? "minmax(0, 1fr)" : "minmax(300px, 0.95fr) minmax(320px, 1fr) minmax(420px, 1.45fr)", gap: 12, alignItems: "stretch" }}>
        {/* linhas */}
        <div style={box}>
          <div style={head}>
            <Layers size={18} color={X.ink} /> Lines <span style={headSub}>· {periodTxtMain}</span>
          </div>
          {detail && detail.lt > 0 ? (
            /* grade que ocupa a altura toda do quadro: as três linhas (PT, CP, MH) dividem o espaço igualmente */
            <div style={{ flex: 1, display: "grid", gridTemplateColumns: mob ? "60px 44px minmax(0, 1fr) minmax(0, 1.2fr)" : "104px 72px minmax(80px, 1fr) minmax(140px, 1.6fr)", gridTemplateRows: "auto repeat(3, minmax(52px, 1fr))", fontSize: 13 }}>
              {["Product Line", "Quotes", "Value", "Share of line value"].map((h) => (
                <div key={h} style={{ ...colH, display: "flex", alignItems: "flex-end" }}>
                  {h}
                </div>
              ))}
              {["PT", "CP", "MH"].map((k) => {
                const sh = detail.lines[k] / detail.lt;
                const lc = LINE_COLOR[k];
                const cell = { ...rowB, display: "flex", alignItems: "center", minWidth: 0 };
                return (
                  <React.Fragment key={k}>
                    <div style={cell}>
                      <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 48, height: 32, borderRadius: 999, fontSize: 14, fontWeight: 800, color: lc[0], background: lc[1] }}>{k}</span>
                    </div>
                    <div style={{ ...cell, fontSize: 15, fontWeight: 800, color: X.ink }}>{detail.lineN[k]}</div>
                    <div style={{ ...cell, whiteSpace: "nowrap", color: X.ink }}>{moneyShort(detail.lines[k])}</div>
                    <div style={{ ...cell, gap: 8, whiteSpace: "nowrap" }}>
                      <span style={{ width: 40, fontWeight: 600, flexShrink: 0 }}>{pctTxt(sh)}</span>
                      <span style={{ flex: 1, maxWidth: 200, height: 7, borderRadius: 4, background: X.track, overflow: "hidden" }}>
                        <span style={{ display: "block", height: 7, width: Math.max(sh > 0 ? 3 : 0, sh * 100) + "%", background: lc[0], borderRadius: 4 }} />
                      </span>
                    </div>
                  </React.Fragment>
                );
              })}
            </div>
          ) : (
            <div style={{ fontSize: 12, color: X.sub }}>No quotes in the period of the screen.</div>
          )}
          {detail && (
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 11.5, color: X.sub }}>
              <span style={{ width: 14, height: 6, borderRadius: 4, background: "#2F6FDB" }} /> Share of line value · {detail.n} quotes
            </div>
          )}
        </div>
        {/* Product Groups */}
        <div style={box}>
          <div style={head}>
            <LayoutGrid size={18} color={X.ink} /> Product groups <span style={headSub}>· {periodTxtMain}</span>
          </div>
          {detail && detail.pgs.length ? (
            <div style={{ maxHeight: 330, overflow: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr>
                    <th style={{ ...colH, textAlign: "left", position: "sticky", top: 0, background: "#fff" }}>Group</th>
                    <th style={{ ...colH, textAlign: "left", position: "sticky", top: 0, background: "#fff", width: 70 }}>QTD</th>
                    <th style={{ ...colH, textAlign: "left", position: "sticky", top: 0, background: "#fff", width: 84 }}>Value</th>
                    <th style={{ ...colH, textAlign: "left", position: "sticky", top: 0, background: "#fff", width: 64 }}>%</th>
                  </tr>
                </thead>
                <tbody>
                  {detail.pgs.slice(0, 20).map(([k, o]) => {
                    const sh = detail.pgT ? o.v / detail.pgT : 0;
                    return (
                      <tr key={k} style={rowB}>
                        <td style={{ padding: "6px 8px 6px 0", maxWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: X.ink }} title={k}>
                          {k}
                        </td>
                        <td style={{ color: X.sub, whiteSpace: "nowrap" }}>{o.n} QTs</td>
                        <td style={{ fontWeight: 700, whiteSpace: "nowrap" }}>{moneyShort(o.v)}</td>
                        <td style={{ padding: "4px 0", lineHeight: 1.1 }}>
                          <div style={{ fontWeight: 600, fontSize: 11.5, marginBottom: 2 }}>{pctTxt(sh)}</div>
                          <MiniBar v={sh} color={X.teal} w={50} h={4} />
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {detail.pgs.length > 20 && <div style={{ fontSize: 11, color: X.sub, paddingTop: 6 }}>+ {detail.pgs.length - 20} more groups</div>}
            </div>
          ) : (
            <div style={{ fontSize: 12, color: X.sub }}>{detail && detail.n ? "The stored quotes have no Product Group yet: import the Quote Report again." : "No quotes in the period of the screen."}</div>
          )}
          {detail && detail.withPg > 0 && detail.withPg < detail.n && <div style={{ marginTop: "auto", fontSize: 11, color: X.sub }}>{detail.n - detail.withPg} quotes without Product Group (import the Quote Report again)</div>}
        </div>
        {/* follow-ups do período */}
        <div style={box}>
          <div style={head}>
            <ListIcon size={18} color={X.ink} /> Follow-ups <span style={headSub}>· {periodTxt}</span>
          </div>
          <div style={{ maxHeight: 330, overflow: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr>
                  {[["Quotation ID", 116], ["Customer", null], ["Status", 150], ["Amount", 90]].map(([h, w]) => (
                    <th key={h} style={{ ...colH, textAlign: h === "Amount" ? "right" : "left", width: w || undefined, position: "sticky", top: 0, background: "#fff", zIndex: 1 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {[...x.items]
                  .sort((a, b) => (a.kind === b.kind ? (a.kind === "upd" ? (a.t < b.t ? -1 : 1) : a.dl < b.dl ? -1 : 1) : a.kind === "pend" ? -1 : 1))
                  .map((it) => {
                    const [txt, fg, bg] =
                      it.kind === "upd"
                        ? [`Updated ${it.t.slice(11, 16)}${from !== to ? " · " + fmtDay(it.t.slice(0, 10)) : ""}`, it.late ? X.amber : X.green, it.late ? X.amberSoft : X.greenSoft]
                        : [it.overdue ? `Overdue · ${fmtDay(it.dl)}` : `To update · ${fmtDay(it.dl)}`, X.redDk, X.redSoft];
                    return (
                      <tr key={it.id} style={rowB}>
                        <td style={{ padding: "7px 0" }}>
                          <span
                            onClick={(e) => {
                              e.stopPropagation();
                              onOpenQuote(it.id);
                            }}
                            title="Open the quote"
                            style={{ display: "inline-block", fontWeight: 700, fontSize: 12, color: X.tealDk, background: X.tealSoft, borderRadius: 6, padding: "3px 8px", cursor: "pointer" }}
                          >
                            {it.id}
                          </span>
                        </td>
                        <td style={{ maxWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: "#5B7083", paddingRight: 8 }} title={it.c}>
                          {it.c}
                        </td>
                        <td>
                          <span style={{ display: "inline-block", fontSize: 11.5, fontWeight: 600, color: fg, background: bg, borderRadius: 6, padding: "3px 8px", whiteSpace: "nowrap" }}>{txt}</span>
                        </td>
                        <td style={{ textAlign: "right", fontWeight: 700, whiteSpace: "nowrap" }}>{moneyShort(it.a)}</td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
  );
  return (
    <div style={{ ...card, overflow: "hidden" }}>
      {mob ? (
        /* celular: o período num seletor, compartilhar num botão, e o total da equipe numa barra */
        <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 10, borderBottom: `1px solid ${X.line}` }}>
          <div style={{ display: "flex", gap: 8 }}>
            <select value={rk} onChange={(e) => setRk(e.target.value)} style={{ ...inp, flex: 1, minWidth: 0, height: 44, fontSize: 14, fontWeight: 600, borderRadius: 12 }}>
              {RANGES.map(([k, l]) => (
                <option key={k} value={k}>
                  {l}
                </option>
              ))}
            </select>
            <button onClick={() => setShare(true)} disabled={!ok} title="Excel, e-mail or PDF" style={iconBtn({ width: 44, height: 44 })}>
              <Download size={18} />
            </button>
          </div>
          {rk === "custom" && (
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input type="date" value={cf} onChange={(e) => setCf(e.target.value)} style={{ ...inp, height: 40, flex: 1, minWidth: 0 }} />
              <span style={{ color: X.sub }}>to</span>
              <input type="date" value={ct} onChange={(e) => setCt(e.target.value)} style={{ ...inp, height: 40, flex: 1, minWidth: 0 }} />
            </div>
          )}
          {ok && rep.length > 0 && (
            <div>
              <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
                <span style={{ fontSize: 24, fontWeight: 800, color: X.tealDk }}>{tot.upd}</span>
                <span style={{ fontSize: 13.5, color: X.sub, fontWeight: 600 }}>of {tot.total} updated</span>
                <span style={{ marginLeft: "auto", fontSize: 12.5, fontWeight: 700, color: tot.pend ? X.redDk : X.green }}>{tot.pend ? `${tot.pend} to update` : "all updated"}</span>
              </div>
              <div style={{ height: 8, borderRadius: 9, background: X.head, overflow: "hidden", marginTop: 7 }}>
                <div style={{ width: `${tot.total ? (tot.upd / tot.total) * 100 : 0}%`, height: "100%", background: X.teal, borderRadius: 9 }} />
              </div>
            </div>
          )}
          {!ok && <div style={{ fontSize: 12.5, color: X.redDk }}>Choose a valid period</div>}
          <Sheet open={share} title="Share the report" onClose={() => setShare(false)}>
            <button onClick={() => { toExcel(); setShare(false); }} style={{ ...btn, height: 48, justifyContent: "flex-start", borderRadius: 12, fontSize: 14 }}>
              <FileSpreadsheet size={18} /> Download as Excel
            </button>
            <button onClick={toMail} disabled={!rep.length} style={{ ...btn, height: 48, justifyContent: "flex-start", borderRadius: 12, fontSize: 14 }}>
              <Copy size={18} /> {copied ? "Copied! Paste it in the e-mail" : "Copy for e-mail"}
            </button>
            <button onClick={() => { printPdf(); setShare(false); }} style={{ ...btn, height: 48, justifyContent: "flex-start", borderRadius: 12, fontSize: 14 }}>
              <Printer size={18} /> PDF
            </button>
            <div style={{ fontSize: 12, color: X.sub }}>{periodTxt} · {filtersTxt}</div>
          </Sheet>
        </div>
      ) : (
      <div style={{ padding: 16, display: "flex", flexDirection: "column", gap: 12, borderBottom: `1px solid ${X.line}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ width: 38, height: 38, borderRadius: 10, background: X.tealSoft, color: X.tealDk, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <Users size={20} />
          </span>
          <div>
            <div style={h2}>FUP by salesperson</div>
            <div style={{ fontSize: 12, color: X.sub, marginTop: 2 }}>Follow-ups due and updated (Done) in the period · {filtersTxt}</div>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 6, flexWrap: "wrap" }}>
            <button onClick={toExcel} disabled={!ok} style={{ ...btn, height: 34, fontSize: 12.5, background: X.teal, borderColor: X.teal, color: "#fff" }} title="Download the table (and the quotes of each salesperson, in the Details sheet) as Excel">
              <FileSpreadsheet size={15} /> Excel
            </button>
            <button onClick={toMail} disabled={!ok || !rep.length} style={{ ...btn, height: 34, fontSize: 12.5 }} title="Copy the table with its colors: paste it straight into the e-mail (Ctrl+V)">
              <Copy size={15} /> {copied ? "Copied!" : "Copy for e-mail"}
            </button>
            <button onClick={printPdf} disabled={!ok} style={{ ...btn, height: 34, fontSize: 12.5 }} title="Open the table ready to print / save as PDF">
              <Printer size={15} /> PDF
            </button>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
          {RANGES.map(([k, l]) => {
            const on = rk === k;
            return (
              <button key={k} onClick={() => setRk(k)} style={{ ...btn, height: 30, padding: "0 12px", fontSize: 12, borderRadius: 999, background: on ? X.tealSoft : "#fff", color: on ? X.tealDk : X.ink, borderColor: on ? "#9FD6C9" : X.line, fontWeight: on ? 700 : 600 }}>
                {l}
              </button>
            );
          })}
          {rk === "custom" && (
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
              <input type="date" value={cf} onChange={(e) => setCf(e.target.value)} style={{ ...inp, height: 30, fontSize: 12 }} />
              <span style={{ color: X.sub }}>to</span>
              <input type="date" value={ct} onChange={(e) => setCt(e.target.value)} style={{ ...inp, height: 30, fontSize: 12 }} />
            </span>
          )}
          <span style={{ marginLeft: 6, fontSize: 12.5, color: ok ? X.ink : X.redDk }}>{ok ? <b>{periodTxt}</b> : "Choose a valid period"}</span>
        </div>
      </div>
      )}
      {mob ? (
        /* celular: um cartão por vendedor; tocar abre os detalhes embaixo */
        <div style={{ padding: 10, display: "flex", flexDirection: "column", gap: 8 }}>
          {!rep.length && <div style={{ textAlign: "center", color: X.sub, padding: 24, fontSize: 12.5 }}>No follow-up due or done in this period with the current filters.</div>}
          {rep.map((x) => {
            const isOpen = open === x.name;
            return (
              <div key={x.name} style={{ border: `1px solid ${isOpen ? "#9FD6C9" : X.line}`, borderRadius: 12, background: isOpen ? "#F7FBFA" : "#fff", overflow: "hidden" }}>
                <div className="qp-card-btn" onClick={() => setOpen(isOpen ? "" : x.name)} style={{ padding: "12px 13px", display: "flex", flexDirection: "column", gap: 7, cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, minWidth: 0 }}>
                    <b style={{ fontSize: 14.5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 }}>{x.name || "(no salesperson)"}</b>
                    <span style={{ marginLeft: "auto", fontSize: 12.5, color: X.sub, whiteSpace: "nowrap" }}>
                      <b style={{ color: X.tealDk, fontSize: 15 }}>{x.upd}</b> / {x.total}
                    </span>
                  </div>
                  <div style={{ height: 6, borderRadius: 9, background: X.head, overflow: "hidden" }}>
                    <div style={{ width: `${x.total ? (x.upd / x.total) * 100 : 0}%`, height: "100%", background: X.teal, borderRadius: 9 }} />
                  </div>
                  <div style={{ display: "flex", gap: 10, fontSize: 12, fontWeight: 600 }}>
                    <span style={{ color: x.pend ? X.redDk : X.green }}>{x.pend ? `${x.pend} to update` : "all updated"}</span>
                    {x.late ? <span style={{ color: X.amber }}>{x.late} after {LATE_HOUR}h</span> : null}
                    <ChevronDown size={15} color={X.sub} style={{ marginLeft: "auto", transform: isOpen ? "rotate(180deg)" : "none", transition: "transform .14s" }} />
                  </div>
                </div>
                {isOpen && <div style={{ padding: "0 10px 6px" }}>{mobFups(x)}</div>}
              </div>
            );
          })}
        </div>
      ) : (
      <div style={{ overflow: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
          <thead>
            <tr>
              <th style={{ ...th, width: 34 }} />
              <th style={{ ...th, textAlign: "left" }}>Responsável</th>
              <th style={th} title="Follow-ups of the period: updated + still to update">Qtd Cotações</th>
              <th style={th} title="Quotes marked Done in the period">Qtd Atualizada</th>
              <th style={th} title={`First Done of the period at ${LATE_HOUR}:00 or later`}>Qtd Atualizada após {LATE_HOUR}h</th>
              <th style={th}>% Atualizada</th>
              <th style={th} title="Open quotes with the follow-up due until the end of the period and no Done in it">Qtd para atualizar</th>
              <th style={th}>% Faltou atualizar</th>
            </tr>
          </thead>
          <tbody>
            {!rep.length && (
              <tr>
                <td colSpan={8} style={{ ...td, textAlign: "center", color: X.sub, padding: 30 }}>
                  No follow-up due or done in this period with the current filters.
                </td>
              </tr>
            )}
            {rep.map((x) => {
              const isOpen = open === x.name;
              return (
                <React.Fragment key={x.name}>
                  <tr className="qp-row" onClick={() => setOpen(isOpen ? "" : x.name)} style={{ cursor: "pointer", background: isOpen ? "#F1F8F6" : "#fff" }}>
                    <td style={{ ...td, textAlign: "center" }}>
                      <ChevronRight size={15} style={{ color: isOpen ? X.tealDk : X.sub, transform: isOpen ? "rotate(90deg)" : "none", transition: "transform .14s" }} />
                    </td>
                    <td style={{ ...td, textAlign: "left", fontWeight: 700 }}>{x.name || "(no salesperson)"}</td>
                    <td style={{ ...td, fontWeight: 700 }}>{x.total}</td>
                    <td style={{ ...td, color: X.green, fontWeight: 600 }}>{x.upd}</td>
                    <td style={{ ...td, color: x.late ? X.amber : X.sub }}>{x.late}</td>
                    <td style={td}>
                      <PctBar v={x.pctUpd} color={X.teal} />
                    </td>
                    <td style={{ ...td, color: x.pend ? X.redDk : X.sub, fontWeight: 600 }}>{x.pend}</td>
                    <td style={td}>
                      <PctBar v={x.pctPend} color={X.redDk} />
                    </td>
                  </tr>
                  {isOpen && (
                    <tr>
                      <td colSpan={8} style={{ padding: "12px 16px 16px 46px", background: "#F7FBFA", borderBottom: `1px solid ${X.line}` }}>
                        {detailBlock(x)}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
          {rep.length > 0 && (
            <tfoot>
              <tr>
                <td style={{ ...td, background: X.head }} />
                <td style={{ ...td, textAlign: "left", fontWeight: 800, background: X.head }}>Total · {rep.length} salespeople</td>
                <td style={{ ...td, fontWeight: 800, background: X.head }}>{tot.total}</td>
                <td style={{ ...td, fontWeight: 800, background: X.head, color: X.green }}>{tot.upd}</td>
                <td style={{ ...td, fontWeight: 800, background: X.head }}>{tot.late}</td>
                <td style={{ ...td, background: X.head }}>
                  <PctBar v={tot.total ? tot.upd / tot.total : 0} color={X.teal} />
                </td>
                <td style={{ ...td, fontWeight: 800, background: X.head, color: X.redDk }}>{tot.pend}</td>
                <td style={{ ...td, background: X.head }}>
                  <PctBar v={tot.total ? tot.pend / tot.total : 0} color={X.redDk} />
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
      )}
    </div>
  );
}

/* ===== Visão semanal =====
   Uma coluna por dia útil (segunda a sexta). No topo de cada dia: a meta e uma barra — nos dias que já
   passaram (e hoje) o que foi vendido (Sales Report); nos dias à frente, o valor das cotações com FUP no dia.
   Embaixo: os cartões das cotações do dia (clique abre a cotação; hoje tem o botão Done) e os follow-ups
   já feitos naquele dia (Done). Clicar no dia (fora de um cartão) abre o plano do dia, como no mês. */
const WEEK_LIST_H = 420; // altura máxima da lista de cada dia na visão semanal (px); passou disso, rola
const tag = (txt, fg, bg, extra) => (
  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 10.5, fontWeight: 700, color: fg, background: bg, padding: "1px 7px", borderRadius: 999, whiteSpace: "nowrap", ...extra }}>{txt}</span>
);
function WeekStat({ label, value, sub, ink, title }) {
  return (
    <div title={title} style={{ flex: "1 1 150px", minWidth: 0, padding: "8px 12px", borderRadius: 9, background: X.head }}>
      <div style={{ fontSize: 10.5, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.3 }}>{label}</div>
      <div style={{ fontSize: 17, fontWeight: 700, color: ink || X.ink, fontVariantNumeric: "tabular-nums", marginTop: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: X.sub, marginTop: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sub}</div>}
    </div>
  );
}
function WeekView({ days, plan, today, sel, pick, goalOf, dayBooked, partial, meta, cands, notesBy, onOpenQuote, onDone }) {
  const mob = useMobile(); // celular: um dia embaixo do outro
  const byId = useMemo(() => new Map(cands.map((c) => [c.id, c])), [cands]);
  // follow-ups feitos em cada dia da semana (o Done fica no histórico do FUP de cada cotação)
  const doneOn = useMemo(() => {
    const m = {};
    days.forEach((d) => (m[d] = []));
    for (const id in meta) {
      const x = meta[id];
      ((x && x.log) || []).forEach((d) => {
        if (m[d] && !m[d].some((y) => y.id === id)) m[d].push({ id, c: (byId.get(id) || {}).c || x.name || "", v: (byId.get(id) || {}).v });
      });
    }
    return m;
  }, [days.join(","), meta, byId]);
  const listOf = (d) => (plan.byDay[d] && plan.byDay[d].list) || [];
  // resumo da semana
  const wGoal = days.reduce((s, d) => s + goalOf(d), 0);
  const elapsed = days.filter((d) => d <= today);
  const goalSoFar = elapsed.reduce((s, d) => s + goalOf(d), 0);
  const wBooked = elapsed.reduce((s, d) => s + (dayBooked[d] || 0), 0);
  const ahead = days.filter((d) => d >= today);
  const fupAhead = ahead.reduce((s, d) => s + ((plan.byDay[d] && plan.byDay[d].planned) || 0), 0);
  const nItems = days.reduce((s, d) => s + listOf(d).length, 0);
  const nOver = days.includes(today) ? listOf(today).filter((e) => e.kind === "overdue").length : 0;
  const nDone = days.reduce((s, d) => s + doneOn[d].length, 0);
  const hit = elapsed.filter((d) => d < today && goalOf(d) > 0 && (dayBooked[d] || 0) >= goalOf(d)).length;
  const withG = elapsed.filter((d) => d < today && goalOf(d) > 0).length;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {/* resumo da semana */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <WeekStat label="Goal of the week" value={wGoal ? moneyShort(wGoal) : "—"} sub={wGoal ? `${days.filter((d) => goalOf(d) > 0).length} business days` : "no daily goal"} title={money(wGoal)} />
        <WeekStat
          label="Booked in the week"
          value={!elapsed.length ? "—" : partial ? "—" : moneyShort(wBooked)}
          sub={!elapsed.length ? "week still ahead" : partial ? "import the Sales Report again" : goalSoFar ? `${Math.round((wBooked / goalSoFar) * 100)}% of the goal until today · ${hit} of ${withG} days on target` : "Sales Report"}
          ink={elapsed.length && !partial ? (wBooked >= goalSoFar ? X.green : X.redDk) : X.ink}
          title={money(wBooked)}
        />
        <WeekStat label="With FUP from today" value={ahead.length ? moneyShort(fupAhead) : "—"} sub={ahead.length ? "value of the quotes with FUP on these days" : "week is over"} ink="#2F6FDB" title={money(fupAhead)} />
        <WeekStat label="Follow-ups" value={nItems || "—"} sub={nOver ? `${nOver} overdue, brought to today` : ahead.length ? "planned from today" : "—"} ink={nOver ? X.redDk : X.ink} />
        <WeekStat label="Done" value={nDone || "—"} sub="follow-ups marked Done this week" ink={X.tealDk} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: mob ? "minmax(0, 1fr)" : "repeat(5, minmax(0, 1fr))", gap: 8, alignItems: "stretch" }}>
        {days.map((d) => {
          const evs = listOf(d);
          const done = doneOn[d];
          const isT = d === today,
            past = d < today,
            isSel = d === sel;
          const goal = goalOf(d);
          const booked = dayBooked[d];
          const pd = plan.byDay[d];
          // barra do dia: vendido x meta (passado e hoje) ou valor com FUP x meta (à frente)
          const showBooked = !partial && booked != null && d <= today;
          const bv = showBooked ? booked : pd ? pd.planned : 0;
          const pct = goal ? Math.min(1, bv / goal) : 0;
          const barCol = showBooked ? (bv >= goal ? X.green : isT ? X.amber : X.redDk) : "#2F6FDB";
          return (
            <div
              key={d}
              className="qp-day"
              onClick={() => pick(d)}
              title={isSel ? "Click again to close the plan of this day" : "Click to see the plan of this day"}
              style={{
                borderRadius: 10,
                border: `1px solid ${isSel ? X.teal : isT ? "#9FD6C9" : X.line}`,
                boxShadow: isSel ? `0 0 0 3px ${X.tealSoft}` : "none",
                background: past ? "#FAFBFC" : "#fff",
                cursor: "pointer",
                display: "flex",
                flexDirection: "column",
                minWidth: 0,
                minHeight: mob ? 0 : 280,
                overflow: "hidden",
              }}
            >
              {/* cabeçalho do dia */}
              <div style={{ padding: "9px 11px 10px", background: isT ? "#EEF8F5" : isSel ? "#F4FBF9" : X.head, borderBottom: `1px solid ${X.line}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ fontSize: 10.5, fontWeight: 700, color: isT ? X.tealDk : X.sub, textTransform: "uppercase", letterSpacing: 0.35 }}>{dowName(d)}</span>
                  {isT && tag("Today", "#fff", X.teal)}
                  {evs.length > 0 && <span style={{ marginLeft: "auto", fontSize: 11, fontWeight: 700, color: X.sub }}>{evs.length} FUP</span>}
                </div>
                <div style={{ fontSize: 19, fontWeight: 700, color: past ? "#8A97A3" : X.ink, letterSpacing: -0.3, marginTop: 1 }}>{fmtDay(d)}</div>
                <HolidayTag d={d} full />
                {(goal > 0 || bv > 0) && (
                  <div style={{ marginTop: 7 }} title={showBooked ? `Booked ${money(bv)}${goal ? ` of ${money(goal)}` : ""}` : `Quotes with FUP on this day: ${money(bv)}${goal ? ` · goal ${money(goal)}` : ""}`}>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 4, fontSize: 11, fontVariantNumeric: "tabular-nums", flexWrap: "wrap" }}>
                      <span style={{ fontWeight: 600, color: X.sub, whiteSpace: "nowrap" }}>{showBooked ? "Booked" : "With FUP"}</span>
                      <span style={{ marginLeft: "auto", whiteSpace: "nowrap" }}>
                        <span style={{ fontWeight: 700, color: barCol }}>{moneyShort(bv)}</span>
                        {goal > 0 && <span style={{ color: X.sub }}> / {moneyShort(goal)}</span>}
                      </span>
                    </div>
                    {goal > 0 && (
                      <div style={{ height: 5, borderRadius: 4, background: "#E3E9EE", marginTop: 4, overflow: "hidden" }}>
                        <div style={{ height: 5, width: Math.max(bv > 0 ? 3 : 0, pct * 100) + "%", background: barCol, borderRadius: 4 }} />
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* cotações do dia: altura limitada, com barra de rolagem própria quando o dia tem muitas cotações */}
              <div style={{ padding: 7, display: "flex", flexDirection: "column", gap: 6, flex: 1, maxHeight: mob ? "none" : WEEK_LIST_H, overflowY: mob ? "visible" : "auto", scrollbarWidth: "thin", scrollbarColor: "#CBD5E1 transparent" }}>
                {evs.map((e) => {
                  const f = FUP[e.fup];
                  const ov = e.kind === "overdue";
                  const why = ov
                    ? tag(`Overdue · ${fmtDay(e.dl)}`, X.redDk, X.redSoft)
                    : e.kind === "again"
                    ? tag("2nd follow-up", f.color, f.bg)
                    : e.fup === "B"
                    ? tag(fupPending(e.meta || {}).length > 1 ? "1st follow-up" : "2nd follow-up", f.color, f.bg)
                    : e.kind === "fill"
                    ? tag("Pulled in", "#2F6FDB", "#E6EEFB")
                    : f
                    ? tag(f.label, f.color, f.bg)
                    : null;
                  return (
                    <div
                      key={e.id + e.kind}
                      onClick={(ev) => {
                        ev.stopPropagation();
                        onOpenQuote(e.id);
                      }}
                      title={`${e.id} · ${e.c} · ${money(e.v)} — click to open the record, the Activities History and the notes`}
                      style={{
                        border: `1px ${ov ? "dashed" : "solid"} ${ov ? X.redDk : X.line}`,
                        borderLeft: `3px solid ${ov ? X.redDk : f ? f.color : "#7DA2E8"}`,
                        borderRadius: 8,
                        padding: "6px 8px 7px",
                        background: "#fff",
                        display: "flex",
                        flexDirection: "column",
                        gap: 3,
                        minWidth: 0,
                        boxShadow: "0 1px 2px rgba(16,24,40,.04)",
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                        <span style={{ fontSize: 10.5, fontWeight: 800, color: f ? f.color : "#2F6FDB", background: f ? f.bg : "#E6EEFB", borderRadius: 5, padding: "1px 6px", flexShrink: 0 }}>{e.fup || "+"}</span>
                        <span style={{ fontSize: 12.5, fontWeight: 700, color: X.ink, whiteSpace: "nowrap", flexShrink: 0 }}>{e.id}</span>
                        <NoteBadge info={notesBy && notesBy[e.id]} />
                        <span style={{ marginLeft: "auto", fontSize: 11.5, fontWeight: 700, fontVariantNumeric: "tabular-nums", minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={money(e.v)}>
                          {moneyShort(e.v)}
                        </span>
                      </div>
                      <div style={{ fontSize: 11.5, color: X.sub, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.c}</div>
                      <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap", marginTop: 1 }}>
                        {why}
                        {e.win ? tag(`${e.win}%`, WIN_META[e.win] ? WIN_META[e.win].color : X.ink, WIN_META[e.win] ? WIN_META[e.win].bg : X.head) : tag("no % Win", "#8A97A3", X.head)}
                        {TEMP[e.temp] && tag(TEMP[e.temp].label, TEMP[e.temp].color, TEMP[e.temp].bg)}
                        {isT && e.fup && <DoneButton id={e.id} onOpenQuote={onOpenQuote} small />}
                      </div>
                    </div>
                  );
                })}
                {!evs.length && !done.length && (
                  <div style={{ margin: "auto 0", textAlign: "center", fontSize: 11.5, color: "#A0AEC0", padding: "18px 4px" }}>{past ? "No follow-up done" : "No follow-up"}</div>
                )}
                {done.length > 0 && (
                  <div style={{ marginTop: evs.length ? 4 : 0, paddingTop: evs.length ? 7 : 0, borderTop: evs.length ? `1px dashed ${X.line}` : "none", display: "flex", flexDirection: "column", gap: 3 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: X.tealDk, textTransform: "uppercase", letterSpacing: 0.3, display: "flex", alignItems: "center", gap: 4 }}>
                      <CheckCircle2 size={12} /> Done · {done.length}
                    </div>
                    {done.map((x) => (
                      <div
                        key={x.id}
                        onClick={(ev) => {
                          ev.stopPropagation();
                          onOpenQuote(x.id);
                        }}
                        title={`${x.id} · ${x.c} — follow-up done on ${fmtDate(d)}`}
                        style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: X.sub, padding: "3px 6px", borderRadius: 6, background: "#F1F8F6", minWidth: 0 }}
                      >
                        <Check size={11} color={X.teal} style={{ flexShrink: 0 }} />
                        <span style={{ fontWeight: 600, color: X.ink, flexShrink: 0 }}>{x.id}</span>
                        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{x.c}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ===== Hoje (tela do vendedor externo, pensada para o celular) =====
   No topo, o que fazer hoje: os follow-ups com prazo hoje e os atrasados (trazidos para hoje), cada um com o
   botão Done (abre a cotação direto na nota + data + % Win). Depois, o que já foi feito hoje e os próximos dias
   úteis. Usa o mesmo plano do Calendar, só que sem meta: entram apenas os prazos do FUP. */
function TodayView({ quotes, cls, meta, thermo, today, notesBy, excl, prefs, onOpenQuote, onGoMine }) {
  const maxAge = prefs.maxAge ?? 365;
  const cands = useMemo(() => openCandidates(quotes, cls, maxAge, (id) => meta[id], today, thermo), [quotes, cls, maxAge, meta, today, thermo]);
  const exclSet = useMemo(() => new Set(excl || []), [excl]);
  const horizon = useMemo(() => addBusinessDays(today, 7), [today]);
  const plan = useMemo(() => planDays(cands, { today, horizon, goalOf: () => 0, bookedToday: 0, strategy: "priority", excluded: exclSet }), [cands, today, horizon, exclSet]);
  const byDay = (plan && plan.byDay) || {};
  const keys = Object.keys(byDay)
    .filter((d) => d >= today)
    .sort();
  const first = keys[0] || today;
  const listOf = (d) => ((byDay[d] && byDay[d].list) || []).filter((e) => e.kind !== "fill");
  const todo = listOf(first);
  const nOver = todo.filter((e) => e.kind === "overdue").length;
  const todoV = todo.reduce((s, e) => s + (e.v || 0), 0);
  const inScope = useMemo(() => new Set((quotes || []).map((q) => q.id)), [quotes]);
  const byId = useMemo(() => new Map(cands.map((c) => [c.id, c])), [cands]);
  // follow-ups feitos hoje (o Done fica no histórico do FUP de cada cotação), só das cotações desta pessoa
  const doneToday = useMemo(() => {
    const out = [];
    for (const id in meta) {
      const x = meta[id];
      if (!x || !inScope.has(id) || !(x.log || []).includes(today)) continue;
      const c = byId.get(id) || {};
      out.push({ id, c: c.c || x.name || "", v: c.v, next: x.fup ? nextFup(x) : "" });
    }
    return out;
  }, [meta, today, byId, inScope]);
  const noFup = cands.filter((c) => !c.fup).length;
  const ahead = keys.slice(1, 6).map((d) => ({ d, list: listOf(d) }));
  const [openDay, setOpenDay] = useState("");
  const hol = holidayOf(first);

  const [showDone, setShowDone] = useState(false);
  const total = todo.length + doneToday.length;
  const pct = total ? doneToday.length / total : 0;
  const ids = todo.map((e) => e.id);
  const secT = { fontSize: 12, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4, padding: "6px 2px 0" };
  const why = (e) => {
    const f = FUP[e.fup];
    if (e.kind === "overdue") return [`Overdue since ${fmtDay(e.dl)}`, X.redDk];
    if (e.kind === "again" || (e.fup === "B" && fupPending(e.meta || {}).length <= 1)) return ["2nd follow-up of the week", f ? f.color : X.sub];
    return [f ? f.label : "", f ? f.color : X.sub];
  };
  const fold = (on, label, n, onClick) => (
    <div onClick={onClick} className="qp-card-btn" style={{ display: "flex", alignItems: "center", gap: 8, padding: "12px 2px", cursor: "pointer", fontSize: 13.5, fontWeight: 700, color: X.ink }}>
      {label}
      <span style={{ fontSize: 12, fontWeight: 700, color: X.sub, background: X.head, borderRadius: 999, padding: "1px 8px" }}>{n}</span>
      <ChevronDown size={16} color={X.sub} style={{ marginLeft: "auto", transform: on ? "rotate(180deg)" : "none", transition: "transform .14s" }} />
    </div>
  );
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {/* o dia numa linha: quantos faltam, quantos já foram */}
      <div style={{ ...card, padding: "14px 14px 12px" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 7 }}>
          <span style={{ fontSize: 28, fontWeight: 800, letterSpacing: -0.6, lineHeight: 1 }}>{todo.length}</span>
          <span style={{ fontSize: 14, fontWeight: 600 }}>to do {first === today ? "today" : "next"}</span>
          <span style={{ marginLeft: "auto", fontSize: 12.5, color: X.sub, whiteSpace: "nowrap" }}>
            {dowName(first).slice(0, 3)}, {fmtDay(first)}
          </span>
        </div>
        <div style={{ height: 8, borderRadius: 9, background: X.head, margin: "11px 0 9px", overflow: "hidden" }}>
          <div style={{ width: `${pct * 100}%`, height: "100%", background: X.teal, borderRadius: 9, transition: "width .3s" }} />
        </div>
        <div style={{ display: "flex", gap: 14, fontSize: 12.5, color: X.sub, flexWrap: "wrap" }}>
          <span>
            <b style={{ color: X.tealDk }}>{doneToday.length}</b> done
          </span>
          {nOver > 0 && (
            <span>
              <b style={{ color: X.redDk }}>{nOver}</b> overdue
            </span>
          )}
          {hol && <span>{hol.name}</span>}
          {todo.length > 0 && <span style={{ marginLeft: "auto", fontWeight: 600 }}>{moneyShort(todoV)} in quotes</span>}
        </div>
      </div>

      {/* a lista do dia: tocar abre a cotação; Done registra o follow-up */}
      {!todo.length ? (
        <div style={{ padding: "34px 14px", textAlign: "center", color: X.sub, fontSize: 14 }}>
          <CheckCircle2 size={34} color={X.teal} style={{ display: "block", margin: "0 auto 8px" }} />
          <b style={{ color: X.ink }}>{doneToday.length ? "All done for today" : "Nothing due today"}</b>
        </div>
      ) : (
        todo.map((e) => {
          const ov = e.kind === "overdue";
          const [wTxt, wCol] = why(e);
          return (
            <div
              key={e.id + e.kind}
              className="qp-card-btn"
              onClick={() => onOpenQuote(e.id, false, ids)}
              style={{ ...card, boxShadow: "none", padding: "12px 10px 12px 13px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", minWidth: 0, ...(ov ? { borderLeft: `4px solid ${X.redDk}`, paddingLeft: 10 } : null) }}
            >
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 14, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={e.c}>
                  {e.c}
                </div>
                <div style={{ fontSize: 12, color: X.sub, marginTop: 3, display: "flex", alignItems: "center", gap: 5, whiteSpace: "nowrap" }}>
                  {e.id} · <b style={{ color: X.ink }}>{money(e.v)}</b>
                  <NoteBadge info={notesBy && notesBy[e.id]} />
                </div>
                {wTxt && <div style={{ fontSize: 11.5, fontWeight: 700, color: wCol, marginTop: 3 }}>{wTxt}</div>}
              </div>
              {e.fup && (
                <button
                  onClick={(ev) => {
                    ev.stopPropagation();
                    onOpenQuote(e.id, true);
                  }}
                  title="Register the follow-up: note, next date and % Win"
                  style={{ ...btn, height: 42, padding: "0 14px", borderRadius: 12, background: X.teal, borderColor: X.teal, color: "#fff", fontWeight: 700, fontSize: 13.5, flexShrink: 0 }}
                >
                  <Check size={16} /> Done
                </button>
              )}
            </div>
          );
        })
      )}

      {noFup > 0 && (
        <div onClick={onGoMine} className="qp-card-btn" style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 2px", fontSize: 12.5, color: X.amber, cursor: onGoMine ? "pointer" : "default" }}>
          <AlertTriangle size={14} style={{ flexShrink: 0 }} />
          <span style={{ flex: 1 }}>
            {noFup.toLocaleString("pt-BR")} open quote{noFup === 1 ? " has" : "s have"} no FUP and never show up here
          </span>
          {onGoMine && <ChevronRight size={15} />}
        </div>
      )}

      {/* feitos hoje (recolhido) */}
      {doneToday.length > 0 && (
        <div style={{ borderTop: `1px solid ${X.line}` }}>
          {fold(showDone, "Done today", doneToday.length, () => setShowDone((v) => !v))}
          {showDone &&
            doneToday.map((x) => (
              <div key={x.id} className="qp-opt" onClick={() => onOpenQuote(x.id)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 4px", cursor: "pointer", fontSize: 13, minWidth: 0 }}>
                <CheckCircle2 size={16} color={X.teal} style={{ flexShrink: 0 }} />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0, flex: 1 }}>{x.c || x.id}</span>
                {x.next && <span style={{ fontSize: 12, color: X.sub, whiteSpace: "nowrap" }}>next {fmtDay(x.next)}</span>}
              </div>
            ))}
        </div>
      )}

      {/* próximos dias (cada dia abre a sua lista) */}
      {ahead.some((x) => x.list.length) && (
        <div style={{ borderTop: `1px solid ${X.line}` }}>
          <div style={secT}>Next days</div>
          {ahead.map(({ d, list }) => {
            if (!list.length) return null;
            const on = openDay === d;
            return (
              <div key={d}>
                {fold(
                  on,
                  <span>
                    {dowName(d).slice(0, 3)} {fmtDay(d)} <span style={{ fontWeight: 500, color: X.sub }}>· {moneyShort(list.reduce((s, e) => s + (e.v || 0), 0))}</span>
                  </span>,
                  list.length,
                  () => setOpenDay(on ? "" : d),
                )}
                {on &&
                  list.map((e) => (
                    <div key={e.id + e.kind} className="qp-opt" onClick={() => onOpenQuote(e.id, false, list.map((x) => x.id))} style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 4px", cursor: "pointer", fontSize: 13, minWidth: 0 }}>
                      <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1, minWidth: 0 }}>{e.c}</span>
                      <b style={{ whiteSpace: "nowrap", fontVariantNumeric: "tabular-nums" }}>{moneyShort(e.v)}</b>
                    </div>
                  ))}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* goalEdit = pode mudar a meta (gerência); o vendedor interno vê a meta, mas não muda */
function CalendarView({ onAskClose, notesBy, thermo, quotes, sales, cls, line, meta, onMeta, onDone, onToggleExcl, onOpenQuote, dailyGoal, setDailyGoal, dayGoals, setDayGoals, excl, prefs, setPref, today, goalEdit = true }) {
  const [ym, setYm] = useState(today.slice(0, 7));
  const [sel, setSel] = useState(""); // dia clicado; sem dia, o plano não aparece
  // visão do calendário: mensal (padrão) ou semanal (segunda a sexta)
  const weekly = prefs.calView === "week";
  const mob = useMobile(); // celular: mês compacto (só o número de follow-ups em cada dia) e o plano embaixo
  const [editGoal, setEditGoal] = useState(false); // celular: o campo da meta só aparece ao tocar em "Change"
  const mondayOf = (d) => addDays(d, -((dowOf(d) + 6) % 7));
  const [wk, setWk] = useState(() => mondayOf(today)); // segunda-feira da semana mostrada
  const planRef = useRef(null);
  const mFrom = ym + "-01",
    mTo = lastDayOfMonth(ym);
  const maxAge = prefs.maxAge ?? 365;
  const cands = useMemo(() => openCandidates(quotes, cls, maxAge, (id) => meta[id], today, thermo), [quotes, cls, maxAge, meta, today, thermo]);
  const weighted = cands.reduce((s, c) => s + (c.v * c.win) / 100, 0);
  /* Meta de um dia: o valor digitado para aquele dia, senão a meta diária padrão (só em dia útil;
     sábado e domingo ficam sem meta, a não ser que você digite uma). */
  const goalOf = (d) => (dayGoals[d] != null ? Math.max(0, dayGoals[d]) : isOffDay(d) ? 0 : dailyGoal || 0);
  const bookedToday = useMemo(() => bookedRange(today, today, quotes, sales, cls, line), [today, quotes, sales, cls, line]);
  const goalToday = goalOf(today);
  const missingToday = Math.max(0, Math.round((goalToday - bookedToday.v) * 100) / 100);
  const partial = bookedToday.src === "partial";

  // grade: mês inteiro (segunda a domingo) ou uma semana (segunda a sexta)
  // mês: de domingo a sábado · semana: de segunda a sexta
  const gStart = weekly ? wk : addDays(mFrom, -dowOf(mFrom));
  const gEnd = weekly ? addDays(wk, 4) : addDays(mTo, 6 - dowOf(mTo));
  const days = [];
  for (let d = gStart; d <= gEnd; d = addDays(d, 1)) days.push(d);

  /* Plano: de hoje até o fim do mês mostrado (no mínimo, até o fim do mês atual). O FUP é o prazo
     máximo de cada cotação; os dias sem cotação suficiente puxam outras pela prioridade até a meta. */
  const strategy = ["priority", "likely", "biggest"].includes(prefs.strategy) ? prefs.strategy : "priority";
  const exclSet = useMemo(() => new Set(excl), [excl]);
  const horizon = [gEnd, lastDayOfMonth(today.slice(0, 7))].sort()[1];
  const plan = useMemo(
    () => planDays(cands, { today, horizon, goalOf, bookedToday: partial ? 0 : bookedToday.v, strategy, excluded: exclSet }),
    // goalOf depende de dayGoals e dailyGoal
    [cands, today, horizon, dayGoals, dailyGoal, bookedToday.v, partial, strategy, exclSet],
  );
  const overdueN = plan.live.filter((x) => x.overdue).length;
  const monthCount = days.filter((d) => weekly || d.slice(0, 7) === ym).reduce((s, d) => s + ((plan.byDay[d] && plan.byDay[d].list.length) || 0), 0);
  const noFup = cands.filter((c) => !c.fup).length;

  const go = (k) => {
    if (weekly) {
      const nw = addDays(wk, 7 * k);
      setWk(nw);
      setYm(addDays(nw, 2).slice(0, 7)); // o mês da quarta-feira manda nos ladrilhos do mês
    } else setYm(addMonths(mFrom, k).slice(0, 7));
    setSel("");
  };
  const goToday = () => {
    setYm(today.slice(0, 7));
    setWk(mondayOf(today));
    setSel("");
  };
  const setView = (v) => {
    if (v === "week") {
      // abre a semana de hoje (no mês atual) ou a 1ª semana do mês mostrado
      const w = ym === today.slice(0, 7) ? mondayOf(sel || today) : mondayOf(sel || mFrom);
      setWk(w);
    }
    setPref("calView", v);
  };
  const atToday = weekly ? wk === mondayOf(today) : ym === today.slice(0, 7);
  const pick = (d) => {
    if (!weekly && d.slice(0, 7) !== ym) setYm(d.slice(0, 7));
    setSel(d === sel ? "" : d);
  };
  // ao escolher um dia, rola até o plano
  useEffect(() => {
    if (sel && planRef.current) planRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [sel]);

  /* Booked de cada dia do mês até hoje, para pintar o calendário (verde = bateu a meta do dia,
     vermelho = ficou abaixo) e contar os dias em que a meta foi batida. */
  const dayBooked = useMemo(() => {
    const m = {};
    if (partial) return m;
    const from = gStart < mFrom ? gStart : mFrom,
      to = gEnd > mTo ? gEnd : mTo;
    for (let d = from; d <= to && d <= today; d = addDays(d, 1)) m[d] = bookedRange(d, d, quotes, sales, cls, line).v;
    return m;
  }, [mFrom, mTo, gStart, gEnd, today, quotes, sales, cls, line, partial]);
  const withGoal = Object.keys(dayBooked).filter((d) => d >= mFrom && d <= mTo && goalOf(d) > 0 && d < today);
  const hitDays = withGoal.filter((d) => dayBooked[d] >= goalOf(d)).length;

  let msg, msgTone;
  if (!dailyGoal) {
    msg = "Type the daily goal. Then click a day in the calendar to see which quotes to close that day.";
    msgTone = "info";
  } else if (!goalToday) {
    msg = `No goal today (${holidayOf(today) && holidayOf(today).block ? "holiday: " + holidayOf(today).name : "weekend"}). Click a day in the calendar to see its plan.`;
    msgTone = "info";
  } else if (missingToday <= 0) {
    msg = `Today's goal reached: ${money(bookedToday.v)} booked (${money(bookedToday.v - goalToday)} above the goal).`;
    msgTone = "ok";
  } else {
    msg = `Today: ${money(bookedToday.v)} booked of ${money(goalToday)} — ${money(missingToday)} missing. Click today in the calendar to see which quotes to close.`;
    msgTone = "warn";
  }
  const tone = { info: [X.sub, X.head, Info], ok: [X.green, X.greenSoft, CheckCircle2], warn: [X.amber, X.amberSoft, Crosshair] }[msgTone];

  return (
    <>
      {/* ---------- meta diária ---------- */}
      {mob ? (
        /* celular: a meta numa barra (vendido hoje × meta), o resto fica no computador */
        <div style={{ ...card, padding: "14px 14px 12px", display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4 }}>Today's goal</span>
            {withGoal.length > 0 && (
              <span style={{ marginLeft: "auto", fontSize: 12, color: X.sub }}>
                <b style={{ color: X.amber }}>{hitDays}</b> of {withGoal.length} days on target
              </span>
            )}
          </div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
            <span style={{ fontSize: 26, fontWeight: 800, letterSpacing: -0.5, color: X.green, fontVariantNumeric: "tabular-nums" }}>{partial ? "—" : moneyShort(bookedToday.v)}</span>
            <span style={{ fontSize: 14, color: X.sub, fontWeight: 600 }}>of {goalToday ? moneyShort(goalToday) : "no goal"}</span>
            {goalEdit && (
              <button onClick={() => setEditGoal((v) => !v)} style={{ marginLeft: "auto", border: "none", background: "transparent", color: X.tealDk, fontWeight: 700, fontSize: 13, cursor: "pointer", padding: "4px 0" }}>
                {editGoal ? "Close" : "Change goal"}
              </button>
            )}
          </div>
          <div style={{ height: 8, borderRadius: 9, background: X.head, overflow: "hidden" }}>
            <div style={{ width: `${goalToday ? Math.min(100, (bookedToday.v / goalToday) * 100) : 0}%`, height: "100%", background: X.green, borderRadius: 9 }} />
          </div>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: !goalToday ? X.sub : missingToday ? X.redDk : X.green }}>
            {!goalToday ? "No goal for today" : missingToday ? `${moneyShort(missingToday)} missing today · tap today in the calendar` : "Goal reached today"}
          </div>
          {editGoal && goalEdit && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 4 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.35 }}>Goal per business day</div>
              <MoneyInput value={dailyGoal} onCommit={(n) => setDailyGoal(n)} />
            </div>
          )}
        </div>
      ) : (
      <div style={{ ...card, padding: 18, display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ width: 38, height: 38, borderRadius: 10, background: X.tealSoft, color: X.tealDk, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <Target size={20} />
          </span>
          <div>
            <div style={h2}>Daily goal</div>
            <div style={{ fontSize: 12, color: X.sub, marginTop: 2 }}>
              {partial
                ? "Import the Sales Report again to get the booked value of each day"
                : bookedToday.src === "sales"
                ? `Valid for every business day · booked comes from the Sales Report${cls && sales && sales.hasClass ? ` (${cls} only)` : ""}`
                : bookedToday.src === "quotes"
                ? "Valid for every business day · no Sales Report yet: booked = quotes won"
                : "Valid for every business day · import the reports to compare it with what is booked"}
            </div>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: mob ? "repeat(2, minmax(0, 1fr))" : "minmax(220px, 1.1fr) repeat(auto-fit, minmax(170px, 1fr))", gap: mob ? 8 : 12, alignItems: "stretch" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, justifyContent: "center", ...(mob ? { gridColumn: "1 / -1" } : null) }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.35 }}>Goal per business day</div>
            {goalEdit ? (
              <MoneyInput value={dailyGoal} onCommit={(n) => setDailyGoal(n)} />
            ) : (
              <div title="Set by the management" style={{ display: "flex", alignItems: "center", gap: 8, height: 46, padding: "0 12px", borderRadius: 10, background: X.head, fontSize: 20, fontWeight: 800, fontVariantNumeric: "tabular-nums" }}>
                {dailyGoal ? money(dailyGoal) : "—"}
                <Lock size={14} color={X.sub} style={{ marginLeft: "auto" }} />
              </div>
            )}
          </div>
          <Tile icon={Wallet} label="Booked today" value={partial ? "—" : moneyShort(bookedToday.v)} caption={goalToday ? `${((bookedToday.v / goalToday) * 100).toFixed(0)}% of today's goal` : "no goal today"} color={X.green} tint={X.greenSoft} ink={X.green} title={money(bookedToday.v)} />
          <Tile icon={Crosshair} label="Missing today" value={goalToday ? moneyShort(missingToday) : "—"} caption={goalToday ? (missingToday ? "to close in quotes" : "goal reached") : "no goal today"} color={missingToday ? X.redDk : X.green} tint={X.redSoft} ink={X.redDk} title={money(missingToday)} />
          <Tile
            icon={CalendarCheck}
            label={`Days on target · ${MONTHS_EN[+ym.slice(5) - 1]}`}
            value={withGoal.length ? `${hitDays} of ${withGoal.length}` : "—"}
            caption={withGoal.length ? "business days until yesterday" : "no past day with a goal"}
            color={X.amber}
            tint={X.amberSoft}
            ink={X.amber}
          />
          <Tile icon={TrendingUp} label="Expected (open × % Win)" value={moneyShort(weighted)} caption={`${cands.length} open quotes · ${noFup} without FUP`} tint="#E6EEFB" ink="#2F6FDB" title={money(weighted)} />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "10px 12px", borderRadius: 9, background: tone[1], color: tone[0], fontSize: 13, fontWeight: 600 }}>
          {React.createElement(tone[2], { size: 16, style: { flexShrink: 0 } })}
          {msg}
        </div>
      </div>
      )}

      {/* ---------- agenda de follow-up ---------- */}
      <div style={{ ...card, minWidth: 0, padding: mob ? 10 : 16, display: "flex", flexDirection: "column", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: mob ? 6 : 10, flexWrap: "wrap" }}>
          {/* mensal / semanal */}
          <div style={{ display: "inline-flex", gap: 3, background: X.head, borderRadius: 9, padding: 3 }}>
            {[
              ["month", "Month"],
              ["week", "Week"],
            ].map(([k, l]) => {
              const on = (k === "week") === weekly;
              return (
                <button
                  key={k}
                  onClick={() => setView(k)}
                  title={k === "week" ? "One week, Monday to Friday" : "The whole month"}
                  style={{ border: "none", borderRadius: 7, padding: "6px 12px", fontSize: 12, fontWeight: on ? 700 : 500, cursor: "pointer", background: on ? "#fff" : "transparent", color: on ? X.tealDk : X.sub, boxShadow: on ? "0 1px 2px rgba(16,24,40,.08)" : "none" }}
                >
                  {l}
                </button>
              );
            })}
          </div>
          <button onClick={() => go(-1)} style={{ ...btn, padding: "0 9px", height: 32 }} title={weekly ? "Previous week" : "Previous month"}>
            <ChevronLeft size={16} />
          </button>
          <div style={{ ...h2, minWidth: mob ? 0 : weekly ? 200 : 150, flex: mob ? 1 : "none", fontSize: mob ? 15 : h2.fontSize, textAlign: "center", whiteSpace: "nowrap" }}>
            {weekly ? `${fmtDay(gStart)} – ${fmtDay(gEnd)} ${gEnd.slice(0, 4)}` : monthLabel(ym)}
          </div>
          <button onClick={() => go(1)} style={{ ...btn, padding: "0 9px", height: 32 }} title={weekly ? "Next week" : "Next month"}>
            <ChevronRight size={16} />
          </button>
          {!atToday && (
            <button onClick={goToday} style={{ ...btn, height: 32 }}>
              Today
            </button>
          )}
          <span style={{ fontSize: 11, fontWeight: 600, color: X.tealDk, background: X.tealSoft, padding: "3px 10px", borderRadius: 999 }}>
            {monthCount} planned in the {weekly ? "week" : "month"}
          </span>
          {overdueN > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: X.redDk, background: X.redSoft, padding: "3px 10px", borderRadius: 999 }}>{overdueN} overdue</span>}
          {!mob && <span style={{ fontSize: 12, color: X.sub }}>{weekly ? "Click a day to see its plan · click a quote to open it" : "Click a day to see its plan · past days: booked / goal · today on: planned / goal"}</span>}
          {!mob && (
            <span style={{ marginLeft: "auto", display: "flex", gap: 6, flexWrap: "wrap" }}>
              {FUP_KEYS.map((k) => (
                <span key={k} title={FUP[k].long}>
                  <FupPill v={k} />
                </span>
              ))}
            </span>
          )}
        </div>
        {weekly ? (
          <WeekView days={days} plan={plan} today={today} sel={sel} pick={pick} goalOf={goalOf} dayBooked={dayBooked} partial={partial} meta={meta} cands={cands} notesBy={notesBy} onOpenQuote={onOpenQuote} onDone={onDone} />
        ) : (
        <div style={{ display: "grid", gridTemplateColumns: `repeat(${weekly ? 5 : 7}, minmax(0, 1fr))`, gap: weekly ? 8 : 4 }}>
          {(weekly ? DOW.slice(0, 5) : ["Sun", ...DOW.slice(0, 6)]).map((d, i) => (
            <div key={d} style={{ fontSize: 10.5, fontWeight: 700, color: !weekly && (i === 0 || i === 6) ? "#A0AEC0" : X.sub, textTransform: "uppercase", letterSpacing: 0.3, padding: "0 4px 2px" }}>
              {d}
            </div>
          ))}
          {days.map((d) => {
            const inM = weekly || d.slice(0, 7) === ym;
            const we = isOffDay(d);
            const pd = plan.byDay[d];
            const evs = (pd && pd.list) || [];
            const isSel = d === sel;
            const isT = d === today;
            const past = d < today;
            return (
              <div
                key={d}
                className="qp-day"
                onClick={() => pick(d)}
                title={isSel ? "Click again to close the plan of this day" : "Click to see the plan of this day"}
                style={{
                  minHeight: weekly ? 300 : mob ? 50 : 100,
                  borderRadius: 8,
                  border: `1px solid ${isSel ? X.teal : X.line}`,
                  boxShadow: isSel ? `0 0 0 3px ${X.tealSoft}` : "none",
                  background: isSel ? "#F4FBF9" : !inM ? "#FAFBFC" : we ? "#F8FAFB" : "#fff",
                  padding: mob ? 3 : 5,
                  cursor: "pointer",
                  display: "flex",
                  flexDirection: "column",
                  gap: 3,
                  minWidth: 0,
                  opacity: inM ? 1 : 0.55,
                }}
              >
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span
                    style={{
                      fontSize: 11.5,
                      fontWeight: isT ? 800 : 600,
                      color: isT ? "#fff" : past ? "#A0AEC0" : X.ink,
                      background: isT ? X.teal : "transparent",
                      borderRadius: 999,
                      minWidth: 21,
                      height: 21,
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      padding: "0 5px",
                    }}
                  >
                    {weekly ? fmtDay(d) : +d.slice(8)}
                  </span>
                  {evs.length > 0 && !mob && <span style={{ fontSize: 10, fontWeight: 700, color: X.sub }}>{evs.length}</span>}
                </div>
                {inM && !mob && <HolidayTag d={d} />}
                {/* celular: no lugar dos cartões, o número de follow-ups do dia (vermelho = tem atrasado) */}
                {mob && inM && (evs.length > 0 || holidayOf(d)) && (
                  <span style={{ marginTop: "auto", alignSelf: "center", display: "inline-flex", alignItems: "center", gap: 3 }}>
                    {evs.length > 0 && (
                      <span style={{ fontSize: 10.5, fontWeight: 800, minWidth: 18, textAlign: "center", padding: "1px 5px", borderRadius: 999, color: "#fff", background: evs.some((e) => e.kind === "overdue") ? X.redDk : X.teal }}>{evs.length}</span>
                    )}
                    {holidayOf(d) && <span title={holidayOf(d).name} style={{ width: 6, height: 6, borderRadius: 9, background: holidayOf(d).block ? X.redDk : "#F2B04C" }} />}
                  </span>
                )}
                {(mob ? [] : weekly ? evs : evs.slice(0, 4)).map((e) => (
                  <div
                    key={e.id}
                    title={`${e.id} · ${e.c} · ${money(e.v)}${e.fup ? ` · FUP ${e.fup}` : ""}${e.win ? ` · ${e.win}%` : ""} · ${
                      e.kind === "overdue" ? `FUP overdue since ${fmtDate(e.dl)}` : e.kind === "deadline" ? "FUP deadline on this day" : e.kind === "again" ? "2nd follow-up of the cycle (B)" : `pulled in to reach the goal${e.dl ? ` (deadline ${fmtDate(e.dl)})` : ""}`
                    }`}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 4,
                      fontSize: 10.5,
                      fontWeight: 600,
                      borderRadius: 5,
                      padding: "2px 4px",
                      // atrasado: branco com borda vermelha tracejada · prazo no dia: cor do FUP ·
                      // puxado para a meta: cinza-azulado com borda tracejada
                      background: e.kind === "overdue" ? "#fff" : e.kind === "deadline" || e.kind === "again" ? FUP[e.fup].bg : "#F1F5FB",
                      color: e.kind === "overdue" ? X.redDk : X.ink,
                      border: e.kind === "overdue" ? `1px dashed ${X.redDk}` : e.kind === "fill" ? "1px dashed #B9CBEA" : "1px solid transparent",
                      minWidth: 0,
                    }}
                  >
                    <span style={{ fontWeight: 800, color: e.fup ? FUP[e.fup].color : "#2F6FDB", flexShrink: 0 }}>{e.fup || "+"}</span>
                    {TEMP[e.temp] && <span style={{ width: 6, height: 6, borderRadius: 9, background: TEMP[e.temp].dot, flexShrink: 0 }} />}
                    {/* só o número da cotação; o cliente fica no balão ao passar o mouse */}
                    <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.id}</span>
                  </div>
                ))}
                {!weekly && !mob && evs.length > 4 && <div style={{ fontSize: 10.5, color: X.sub, fontWeight: 600, paddingLeft: 3 }}>+{evs.length - 4} more</div>}
                <div style={{ marginTop: "auto", display: mob ? "none" : "flex", flexDirection: "column", gap: 1, paddingLeft: 3, fontSize: 10.5, fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>
                  {inM && dayBooked[d] != null && (goalOf(d) > 0 || dayBooked[d] > 0) && (!pd || dayBooked[d] > 0) && (
                    <span
                      title={`Booked ${money(dayBooked[d])}${goalOf(d) ? ` of ${money(goalOf(d))}` : ""}`}
                      style={{ color: !goalOf(d) ? X.sub : dayBooked[d] >= goalOf(d) ? X.green : d === today ? X.amber : X.redDk }}
                    >
                      {pd ? "booked " : ""}
                      {moneyShort(dayBooked[d])}
                      {goalOf(d) > 0 && <span style={{ fontWeight: 500, color: X.sub }}> / {moneyShort(goalOf(d))}</span>}
                    </span>
                  )}
                  {inM && pd && (pd.goal > 0 || pd.planned > 0) && (
                    <span title={`Planned ${money(pd.planned)} for ${money(pd.need)} missing on this day`} style={{ color: pd.planned >= pd.need ? "#2F6FDB" : X.redDk }}>
                      plan {moneyShort(pd.planned)}
                      {pd.need > 0 && <span style={{ fontWeight: 500, color: X.sub }}> / {moneyShort(pd.need)}</span>}
                    </span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
        )}
      </div>

      {/* ---------- plano do dia (só com um dia escolhido) ---------- */}
      {sel && (
        <div ref={planRef} style={{ scrollMarginTop: 12 }}>
          <DayPlan
            key={sel}
            day={sel}
            today={today}
            dailyGoal={dailyGoal}
            quotes={quotes}
            sales={sales}
            cls={cls}
            cands={cands}
            plan={plan}
            notesBy={notesBy}
            thermo={thermo}
            strategy={strategy}
            exclSet={exclSet}
            onMeta={onMeta}
            onAskClose={onAskClose}
            onDone={onDone}
            dayGoals={dayGoals}
            setDayGoals={setDayGoals}
            excl={excl}
            line={line}
            onToggleExcl={onToggleExcl}
            onOpenQuote={onOpenQuote}
            prefs={prefs}
            setPref={setPref}
            onClose={() => setSel("")}
            goalEdit={goalEdit}
          />
        </div>
      )}
    </>
  );
}

/* Plano do dia escolhido: o que o planejamento colocou nesse dia.
   • prazo do FUP vencendo no dia (e, hoje, os atrasados);
   • cotações puxadas de outros dias/sem FUP para cobrir a meta do dia, pela prioridade.
   Dia que já passou mostra só o vendido x meta: o plano começa hoje. */
const reasonsOf = (x, day) => {
  const out = [];
  if (x.kind === "overdue") out.push({ k: "overdue", t: `FUP overdue since ${fmtDay(x.dl)}` });
  else if (x.kind === "deadline") out.push({ k: "due", t: x.fup === "B" && fupPending(x.meta || {}).length > 1 ? "1st follow-up (B)" : x.fup === "B" ? "2nd follow-up (B)" : "FUP deadline on this day" });
  else if (x.kind === "again") out.push({ k: "due", t: "2nd follow-up (B)" });
  else if (x.kind === "fill") out.push({ k: "fill", t: x.dl ? `pulled in · deadline ${fmtDay(x.dl)}` : "pulled in · no FUP" });
  else if (x.kind === "other") out.push({ k: "other", t: x.off ? "left out" : x.plannedFor ? `planned for ${fmtDay(x.plannedFor)}` : !AUTO_FILL && !x.fup ? "no FUP" : "no day yet" });
  if (x.win >= 75) out.push({ k: "likely", t: `${x.win}% win` });
  if (x.temp === "HOT") out.push({ k: "hot", t: "Hot" });
  if (x.age >= 90) out.push({ k: "aging", t: `${x.age} days old` });
  if (!x.win) out.push({ k: "nowin", t: "no % Win" });
  return out;
};
function DayPlan({ onAskClose, notesBy, thermo, day, today, dailyGoal, quotes, sales, cls, line, cands, plan, strategy, exclSet, onMeta, onDone, onToggleExcl, onOpenQuote, dayGoals, setDayGoals, excl, prefs, setPref, onClose, goalEdit = true }) {
  const mob = useMobile(); // celular: as cotações do dia em cartões
  const [showAll, setShowAll] = useState(false);
  const bookedOn = useMemo(() => bookedRange(day, day, quotes, sales, cls, line), [day, quotes, sales, cls, line]);
  const past = day < today,
    isToday = day === today,
    we = isOffDay(day);
  const hol = holidayOf(day);
  const partial = bookedOn.src === "partial";
  // meta do dia: a digitada para esse dia ou a meta diária padrão (fim de semana = sem meta)
  const auto = { v: we ? 0 : dailyGoal || 0 };
  const manual = dayGoals[day];
  const target = manual != null ? Math.max(0, manual) : auto.v;
  const pd = plan.byDay[day] || null;
  const maxAge = prefs.maxAge ?? 365;
  const list = (pd && pd.list) || [];
  const inDay = new Set(list.map((x) => x.id));
  // demais cotações abertas: onde foram planejadas, ou "sem dia" / "fora do plano"
  const others = useMemo(
    () =>
      showAll
        ? cands
            .filter((c) => !inDay.has(c.id))
            .map((c) => {
              const dl = c.fup ? nextFup(c.meta || {}) : "";
              const x = { ...c, dl, kind: "other", off: exclSet.has(c.id), plannedFor: plan.dayOf[c.id] || "" };
              return { ...x, ...priorityOn(x, day), cum: null };
            })
            .sort((a, b) => a.off - b.off || (a.plannedFor || "9") .localeCompare(b.plannedFor || "9") || b.score - a.score)
        : [],
    [showAll, cands, list, exclSet, plan, day],
  );
  const rows = [...list, ...others];
  /* Filtros e ordenação da lista do dia. Categorias filtram (Customer, FUP, % Win, Thermometer, Why);
     data, prioridade e valor ordenam. Sem ordenação escolhida, vale a ordem do plano (#). */
  const [pf, setPf] = useState({});
  const [ps, setPs] = useState(null);
  const WHY = { overdue: "FUP overdue", deadline: "FUP deadline on the day", again: "2nd follow-up (B)", fill: "Pulled in for the goal", other: "Other open quotes" };
  const pKey = (r, id) => (id === "c" ? r.c : id === "fup" ? r.fup || NONE : id === "win" ? (r.win ? String(r.win) : NONE) : id === "temp" ? r.temp || NONE : id === "why" ? r.kind : "");
  const planNo = new Map(rows.map((r, i) => [r.id, i + 1]));
  const pfOn = Object.values(pf).some((v) => v && v.length) || !!ps;
  const shownRows = (() => {
    const act = Object.entries(pf).filter(([, v]) => v && v.length);
    const f = rows.filter((r) => act.every(([id, v]) => v.includes(pKey(r, id))));
    if (!ps) return f;
    const val = (r) => (ps.id === "d" ? r.d : ps.id === "score" ? r.score : r.v);
    return [...f].sort((a, b) => {
      const va = val(a),
        vb = val(b);
      return (typeof va === "string" ? va.localeCompare(vb) : va - vb) * ps.dir || planNo.get(a.id) - planNo.get(b.id);
    });
  })();
  const pOpts = (id) => {
    const cnt = {};
    rows.forEach((r) => {
      const k = pKey(r, id);
      cnt[k] = (cnt[k] || 0) + 1;
    });
    if (id === "c") return Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a] || a.localeCompare(b, "pt-BR")).map((k) => ({ v: k, n: cnt[k], text: k, node: k }));
    if (id === "why") return Object.keys(WHY).filter((k) => cnt[k]).map((k) => ({ v: k, n: cnt[k], text: WHY[k], node: WHY[k] }));
    const keys = id === "fup" ? FUP_KEYS : id === "win" ? WIN_VALS.map(String) : TEMP_KEYS;
    return [
      ...keys.map((k) => ({ v: k, n: cnt[k] || 0, text: k, node: id === "fup" ? <FupPill v={k} /> : id === "win" ? <WinPill v={+k} /> : <TempPill v={k} /> })),
      ...(id === "temp" ? [] : [{ v: NONE, n: cnt[NONE] || 0, node: <span style={{ color: X.sub }}>Not set</span> }]),
    ];
  };
  const PH = [
    { id: "_chk", w: 34 },
    { id: "_no", label: "#", w: 36, tip: "Order in the plan of the day" },
    { id: "id", label: "Quote", tip: "Click a quote to open its record, Activities History and notes" },
    { id: "c", label: "Customer", filter: true, fw: 320 },
    { id: "d", label: "Date", sort: true },
    { id: "fup", label: "FUP", filter: true, manual: true },
    { id: "win", label: "% Win", filter: true, manual: true },
    { id: "temp", label: "Thermometer", filter: true },
    { id: "why", label: "Why", filter: true, fw: 260 },
    { id: "score", label: "Priority", sort: true, tip: `0 to 100: ${Math.round(PRIO_WIN * 100)}% from the % Win + ${Math.round(PRIO_TEMP * 100)}% from the Thermometer (the FUP does not count)` },
    { id: "v", label: "Value", sort: true, right: true },
    { id: "_act", w: 80 },
  ];
  const nDue = list.filter((x) => x.kind !== "fill").length,
    nFill = list.length - nDue;

  let msg, msgTone;
  if (past) {
    msg = `Past day: ${money(bookedOn.v)} booked${target ? ` of ${money(target)} — ${bookedOn.v >= target ? "goal reached" : `${money(target - bookedOn.v)} short`}` : ""}. The plan starts today: click today or a day ahead.`;
    msgTone = !target ? "info" : bookedOn.v >= target ? "ok" : "warn";
  } else if (!pd) {
    msg = "This day is outside the plan.";
    msgTone = "info";
  } else if (!pd.need && !list.length) {
    msg = target ? `Goal of the day already covered: ${money(bookedOn.v)} booked.` : "No goal and no FUP deadline on this day.";
    msgTone = target ? "ok" : "info";
  } else if (!pd.need) {
    msg = `${target ? "Goal of the day already covered" : "No goal on this day"}: only the ${list.length} quote${list.length === 1 ? "" : "s"} with the FUP deadline on this day ${list.length === 1 ? "is" : "are"} planned.`;
    msgTone = "ok";
  } else if (pd.planned >= pd.need) {
    msg = `${list.length} quote${list.length === 1 ? "" : "s"} planned (${money(pd.planned)}) for the ${money(pd.need)} of the day: ${nDue} with the FUP deadline here${nFill ? ` + ${nFill} pulled in by priority` : ""}. Weighted by % Win they are worth about ${money(pd.expected)}.`;
    msgTone = "ok";
  } else if (!AUTO_FILL) {
    msg = list.length
      ? `${list.length} quote${list.length === 1 ? "" : "s"} with FUP on this day (${money(pd.planned)}) for the ${money(pd.need)} of the day: ${money(pd.need - pd.planned)} still missing. The button at the end of the list shows the other open quotes.`
      : `No FUP on this day: ${money(pd.need)} missing for the goal. The button at the end of the list shows the other open quotes.`;
    msgTone = "warn";
  } else {
    msg = `Only ${money(pd.planned)} planned for the ${money(pd.need)} of the day: there is no other open quote left for it. Try quotes of any age or another classification.`;
    msgTone = "warn";
  }
  const tone = { info: [X.sub, X.head, Info], ok: [X.green, X.greenSoft, CheckCircle2], warn: [X.redDk, X.redSoft, AlertTriangle] }[msgTone];

  const exportPlan = () => {
    const data = rows.map((r, i) => ({
      "#": i + 1,
      Quote: r.id,
      Customer: r.c,
      "Quote date": fmtDate(r.d),
      "Age (days)": r.age,
      FUP: r.fup ? `${r.fup} - ${FUP[r.fup].label}` : "",
      "FUP deadline": r.dl ? fmtDate(r.dl) : "",
      "% Win": r.win ? r.win / 100 : "",
      Thermometer: r.temp ? TEMP[r.temp].label : "",
      Why: reasonsOf(r, day)
        .map((x) => x.t)
        .join(" · "),
      Priority: r.score,
      Value: r.v,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 5 }, { wch: 13 }, { wch: 44 }, { wch: 11 }, { wch: 10 }, { wch: 16 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 40 }, { wch: 9 }, { wch: 14 }];
    const head = XLSX.utils.aoa_to_sheet([
      ["Day", fmtDate(day)],
      ["Goal of the day", target],
      ["Booked on the day", bookedOn.v],
      ["Missing on the day", pd ? pd.need : 0],
      ["Planned quotes", list.length],
      ["Planned value", pd ? pd.planned : 0],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, head, "Day");
    XLSX.utils.book_append_sheet(wb, ws, "Quotes of the day");
    XLSX.writeFile(wb, `Plan ${day}.xlsx`);
  };

  const scoreCol = (s) => (s >= 65 ? X.teal : s >= 40 ? "#F2B04C" : "#CBD5E1");
  return (
    <div style={{ ...card, overflow: "hidden", borderColor: "#BFE3D9" }}>
      <div style={{ padding: mob ? "12px 12px 10px" : "16px 16px 14px", display: "flex", flexDirection: "column", gap: mob ? 10 : 14, background: "linear-gradient(180deg, #F4FBF9, #fff)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: mob ? "nowrap" : "wrap" }}>
          {!mob && (
          <span style={{ width: 38, height: 38, borderRadius: 10, background: X.teal, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <ListChecks size={20} />
          </span>
          )}
          <div style={{ minWidth: 0 }}>
            <div style={{ ...h2, ...(mob ? { fontSize: 16 } : null) }}>
              Plan for {isToday ? "today, " : ""}
              {dowName(day)}, {fmtDate(day)}
            </div>
            <div style={{ fontSize: 12, color: X.sub, marginTop: 2, display: mob ? "none" : "block" }}>
              {AUTO_FILL ? "FUP = deadline of each quote · days without enough quotes pull in others by priority until the daily goal" : "Quotes with the FUP on this day · the others stay out of the plan"}
              {hol && hol.block ? ` · holiday: ${hol.name}` : we ? " · weekend" : hol ? ` · ${hol.name} (optional day off)` : ""}
            </div>
          </div>
          <button onClick={onClose} style={{ ...btn, marginLeft: "auto", padding: "0 10px" }} title="Close the plan">
            <XIcon size={16} />
          </button>
        </div>
        {/* celular: sem os quadros; a frase de baixo já diz o que falta no dia */}
        <div style={{ display: mob ? "none" : "grid", gridTemplateColumns: "minmax(220px, 1.1fr) repeat(auto-fit, minmax(160px, 1fr))", gap: 12, alignItems: "stretch" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, justifyContent: "center" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.35 }}>Goal of the day</div>
            {!goalEdit ? (
              <div title="Set by the management" style={{ display: "flex", alignItems: "center", gap: 8, height: 46, padding: "0 12px", borderRadius: 10, background: X.head, fontSize: 20, fontWeight: 800, fontVariantNumeric: "tabular-nums" }}>
                {target ? money(target) : "—"}
                <Lock size={14} color={X.sub} style={{ marginLeft: "auto" }} />
              </div>
            ) : (
            <MoneyInput
              value={target}
              onCommit={(n) =>
                setDayGoals((g) => {
                  const x = { ...g };
                  if (n === auto.v) delete x[day]; // igual à meta padrão: o dia volta a seguir a meta diária
                  else x[day] = n || DAY_NO_GOAL; // 0 digitado = sem meta nesse dia
                  return x;
                })
              }
            />
            )}
            {goalEdit && <div style={{ fontSize: 11.5, color: X.sub }}>
              {manual != null ? (
                <span
                  onClick={() =>
                    setDayGoals((g) => {
                      const n = { ...g };
                      delete n[day];
                      return n;
                    })
                  }
                  style={{ fontWeight: 600, color: X.tealDk, cursor: "pointer" }}
                  title="Go back to the daily goal"
                >
                  Only this day · back to the daily goal ({money(auto.v)})
                </span>
              ) : auto.v ? (
                "The daily goal · type to change only this day"
              ) : (
                `${hol && hol.block ? "Holiday" : we ? "Weekend" : "No daily goal"} · type to set a goal for this day`
              )}
            </div>}
          </div>
          <Tile icon={Wallet} label="Booked on the day" value={partial ? "—" : day > today ? "—" : moneyShort(bookedOn.v)} caption={partial ? "import the Sales Report again" : day > today ? "day still ahead" : "Sales Report"} color={X.green} tint={X.greenSoft} ink={X.green} title={money(bookedOn.v)} />
          <Tile icon={Crosshair} label="Missing on the day" value={pd ? moneyShort(pd.need) : past ? moneyShort(Math.max(0, target - bookedOn.v)) : "—"} caption={pd ? (pd.need ? "to close in quotes" : "covered") : past ? "day is over" : "—"} color={X.redDk} tint={X.redSoft} ink={X.redDk} />
          <Tile icon={ListChecks} label="Planned" value={pd ? moneyShort(pd.planned) : "—"} caption={pd ? (AUTO_FILL ? `${list.length} quote${list.length === 1 ? "" : "s"} · ${nDue} FUP + ${nFill} pulled in` : `${list.length} quote${list.length === 1 ? "" : "s"} with FUP`) : "plan starts today"} color={X.amber} tint={X.amberSoft} ink={X.amber} title={pd ? money(pd.planned) : ""} />
          <Tile icon={TrendingUp} label="Expected (× % Win)" value={pd ? moneyShort(pd.expected) : "—"} caption="planned × chance" tint="#E6EEFB" ink="#2F6FDB" title={pd ? money(pd.expected) : ""} />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "10px 12px", borderRadius: 9, background: tone[1], color: tone[0], fontSize: 13, fontWeight: 600 }}>
          {React.createElement(tone[2], { size: 16, style: { flexShrink: 0 } })}
          {msg}
        </div>
      </div>

      {!past && (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px", flexWrap: "wrap", borderTop: `1px solid ${X.line}` }}>
            <div style={{ fontSize: 15, fontWeight: 700 }}>{mob ? `${list.length} quote${list.length === 1 ? "" : "s"} to work on` : "Quotes to work on"}</div>
            {pfOn && (
              <button onClick={() => { setPf({}); setPs(null); }} style={{ ...btn, height: 28, fontSize: 11.5, padding: "0 10px" }}>
                <XIcon size={13} /> Clear filters and sorting
              </button>
            )}
            {!mob && (
            <span style={{ fontSize: 11, fontWeight: 600, color: X.amber, background: X.amberSoft, padding: "3px 10px", borderRadius: 999 }}>
              {list.length} planned · {cands.length} open
            </span>
            )}
            <div style={{ marginLeft: "auto", display: mob ? "none" : "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
              {AUTO_FILL && <span style={{ fontSize: 11.5, color: X.sub }}>Pull in by</span>}
              {AUTO_FILL && <div style={{ display: "inline-flex", gap: 3, background: X.head, borderRadius: 9, padding: 3 }}>
                {[
                  ["priority", "Priority", "% Win + Thermometer"],
                  ["likely", "Most likely", "Highest % Win first"],
                  ["biggest", "Biggest", "Highest value first: the fewest quotes per day"],
                ].map(([k, l, t]) => (
                  <button
                    key={k}
                    title={t}
                    onClick={() => setPref("strategy", k)}
                    style={{ border: "none", borderRadius: 7, padding: "6px 11px", fontSize: 12, fontWeight: strategy === k ? 700 : 500, cursor: "pointer", background: strategy === k ? "#fff" : "transparent", color: strategy === k ? X.tealDk : X.sub, boxShadow: strategy === k ? "0 1px 2px rgba(16,24,40,.08)" : "none" }}
                  >
                    {l}
                  </button>
                ))}
              </div>}
              <select value={maxAge} onChange={(e) => setPref("maxAge", +e.target.value)} style={{ ...inp, height: 32 }} title="Leave out old quotes that will hardly close">
                <option value={90}>Quotes up to 90 days</option>
                <option value={180}>Quotes up to 180 days</option>
                <option value={365}>Quotes up to 1 year</option>
                <option value={0}>Quotes of any age</option>
              </select>
              <button onClick={exportPlan} style={{ ...btn, height: 32, fontSize: 12 }} title="Download the plan of the day as an Excel file">
                <Download size={14} /> Export
              </button>
            </div>
          </div>
          {mob ? (
            /* celular: a mesma lista do Today (cliente, QT, valor, motivo, Done); FUP e % Win mudam dentro da cotação */
            <div style={{ padding: "8px 10px 12px", display: "flex", flexDirection: "column", gap: 8, borderTop: `1px solid ${X.line}` }}>
              {shownRows.map((r, i) => {
                const other = r.kind === "other";
                const late = r.kind === "overdue";
                const rs = reasonsOf(r, day);
                return (
                  <React.Fragment key={r.id}>
                    {other && !ps && (i === 0 || shownRows[i - 1].kind !== "other") && <div style={{ fontSize: 11.5, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4, padding: "10px 2px 2px" }}>Other open quotes</div>}
                    <div
                      className="qp-card-btn"
                      onClick={() => onOpenQuote(r.id, false, shownRows.map((x) => x.id))}
                      style={{ ...card, boxShadow: "none", padding: "12px 10px 12px 13px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", minWidth: 0, opacity: r.off ? 0.5 : 1, ...(late ? { borderLeft: `4px solid ${X.redDk}`, paddingLeft: 10 } : null) }}
                    >
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 14, fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", textDecoration: r.off ? "line-through" : "none" }} title={r.c}>
                          {r.c}
                        </div>
                        <div style={{ fontSize: 12, color: X.sub, marginTop: 3, display: "flex", alignItems: "center", gap: 5, whiteSpace: "nowrap" }}>
                          {r.id} · <b style={{ color: X.ink }}>{money(r.v)}</b>
                          <NoteBadge info={notesBy && notesBy[r.id]} />
                        </div>
                        {rs.length > 0 && <div style={{ fontSize: 11.5, fontWeight: 700, color: REASON[rs[0].k][0], marginTop: 3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{rs.map((x) => x.t).join(" · ")}</div>}
                      </div>
                      {isToday && !other && r.fup && (
                        <button
                          onClick={(ev) => {
                            ev.stopPropagation();
                            onOpenQuote(r.id, true);
                          }}
                          title="Register the follow-up: note, next date and % Win"
                          style={{ ...btn, height: 42, padding: "0 14px", borderRadius: 12, background: X.teal, borderColor: X.teal, color: "#fff", fontWeight: 700, fontSize: 13.5, flexShrink: 0 }}
                        >
                          <Check size={16} /> Done
                        </button>
                      )}
                    </div>
                  </React.Fragment>
                );
              })}
              {!shownRows.length && (
                <div style={{ textAlign: "center", color: X.sub, padding: "22px 10px", fontSize: 13 }}>
                  {rows.length ? "No quote matches the filters." : !cands.length ? (quotes && quotes.length ? "No open quote in this classification and age." : "Import the Quote Report to see the open quotes.") : "Nothing planned for this day."}
                </div>
              )}
            </div>
          ) : (
          <div style={{ overflow: "auto", maxHeight: 640 }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {PH.map((c) => {
                    const sOn = ps && ps.id === c.id;
                    const fOn = (pf[c.id] || []).length > 0;
                    if (c.id === "_chk")
                      return (
                        <th key={c.id} style={{ ...th, width: c.w, textAlign: "center" }} title="Untick a quote you know will not close: it leaves the plan and the days are planned again">
                          <Check size={12} />
                        </th>
                      );
                    return (
                      <th
                        key={c.id}
                        title={c.tip}
                        style={{ ...th, width: c.w, color: sOn || fOn ? X.tealDk : th.color, ...(c.right ? { textAlign: "right" } : null), ...(c.manual ? { background: "#EEF7F4" } : null) }}
                      >
                        {c.sort ? (
                          <span
                            onClick={() => setPs((p) => (p && p.id === c.id ? (p.dir === -1 ? { id: c.id, dir: 1 } : null) : { id: c.id, dir: -1 }))}
                            title={`${c.tip ? c.tip + " · " : ""}click to sort · 3rd click goes back to the plan order`}
                            style={{ display: "inline-flex", alignItems: "center", gap: 4, cursor: "pointer" }}
                          >
                            {c.label}
                            {React.createElement(sOn ? (ps.dir === 1 ? ArrowUp : ArrowDown) : ArrowUpDown, { size: 12, color: sOn ? X.tealDk : "#B8BDC4" })}
                          </span>
                        ) : c.filter ? (
                          <ColFilter label={c.label} width={c.fw || 240} options={pOpts(c.id)} selected={pf[c.id] || []} onChange={(v) => setPf((f) => ({ ...f, [c.id]: v }))} />
                        ) : (
                          c.label || null
                        )}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {shownRows.map((r, i) => {
                  const other = r.kind === "other";
                  const bg = r.kind === "overdue" ? "#FFF6F6" : r.kind === "deadline" ? "#F4FBF9" : r.kind === "fill" ? "#F6F9FE" : "#fff";
                  const c = { ...td, background: bg, opacity: r.off ? 0.45 : 1, padding: "8px 8px" };
                  const cN = { ...tdN, background: bg, opacity: r.off ? 0.45 : 1, padding: "8px 8px" };
                  const bar = { overdue: X.redDk, deadline: X.teal, fill: "#7DA2E8" }[r.kind];
                  return (
                    <React.Fragment key={r.id}>
                      {other && !ps && (i === 0 || shownRows[i - 1].kind !== "other") && (
                        <tr>
                          <td colSpan={12} style={{ ...td, background: X.head, fontSize: 11, fontWeight: 700, color: X.sub, textTransform: "uppercase", letterSpacing: 0.35, padding: "8px 14px" }}>
                            Other open quotes
                          </td>
                        </tr>
                      )}
                      <tr>
                        <td style={{ ...c, opacity: 1, textAlign: "center", boxShadow: bar ? `inset 3px 0 0 ${bar}` : undefined }}>
                          <input
                            type="checkbox"
                            checked={!r.off}
                            onChange={() => onToggleExcl(r.id, r.c, !r.off)}
                            title={r.off ? "Put back in the plan" : "Leave out of the plan"}
                            style={{ accentColor: X.teal, cursor: "pointer" }}
                          />
                        </td>
                        <td style={{ ...cN, color: X.sub }}>{planNo.get(r.id)}</td>
                        <td style={{ ...c, fontWeight: 700, whiteSpace: "nowrap", textDecoration: r.off ? "line-through" : "none" }}>
                          <span onClick={() => onOpenQuote(r.id)} title="Open the record, Activities History and notes" style={{ cursor: "pointer", color: X.tealDk, borderBottom: `1px dashed ${X.teal}` }}>
                            {r.id}
                          </span>
                          <NoteBadge info={notesBy && notesBy[r.id]} onClick={() => onOpenQuote(r.id)} />
                        </td>
                        <td style={{ ...c, maxWidth: 240, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.c}>
                          {r.c}
                        </td>
                        <td style={{ ...c, whiteSpace: "nowrap" }}>
                          {fmtDate(r.d)} <span style={{ fontSize: 11, color: r.age > 180 ? X.redDk : X.sub }}>{r.age}d</span>
                        </td>
                        <td style={{ ...c, whiteSpace: "nowrap" }}>
                          <PickCell
                            value={r.fup}
                            options={FUP_OPTS}
                            render={fupOptRender}
                            describe={fupOptDescribe}
                            onChange={(v) => (CLOSE_LABEL[v] ? onAskClose(r.id, r.c, v) : onMeta(r.id, { fup: v }, r.c))}
                            title="FUP = deadline of this quote · Cancelled / Lost closes it (with the reason)"
                            width={300}
                          />
                        </td>
                        <td style={c}>
                          <PickCell value={r.win} options={WIN_VALS} render={(v) => <WinPill v={v} />} onChange={(v) => onMeta(r.id, { win: v || 0 }, r.c)} title="% Win of this quote" width={140} />
                        </td>
                        <td style={c}>
                          <ThermoCell t={thermo[r.nn]} />
                        </td>
                        <td style={{ ...c, minWidth: 190 }}>
                          <span style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                            {reasonsOf(r, day).map((x) => (
                              <span key={x.k} style={{ fontSize: 10.5, fontWeight: 700, color: REASON[x.k][0], background: REASON[x.k][1], padding: "2px 7px", borderRadius: 999, whiteSpace: "nowrap" }}>
                                {x.t}
                              </span>
                            ))}
                          </span>
                        </td>
                        <td style={{ ...c, whiteSpace: "nowrap" }}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 7 }}>
                            <span style={{ width: 54, height: 6, borderRadius: 4, background: X.track, overflow: "hidden", display: "inline-block" }}>
                              <span style={{ display: "block", height: 6, width: r.score + "%", background: scoreCol(r.score), borderRadius: 4 }} />
                            </span>
                            <span style={{ fontSize: 12, fontWeight: 700, fontVariantNumeric: "tabular-nums", minWidth: 20 }}>{r.score}</span>
                          </span>
                        </td>
                        <td style={{ ...cN, fontWeight: 600 }}>{money(r.v)}</td>
                        <td style={{ ...c, opacity: 1, textAlign: "right", paddingRight: 12 }}>
                          {isToday && !other && r.fup && <DoneButton id={r.id} onOpenQuote={onOpenQuote} />}
                        </td>
                      </tr>
                    </React.Fragment>
                  );
                })}
                {!shownRows.length && rows.length > 0 && (
                  <tr>
                    <td colSpan={12} style={{ ...td, textAlign: "center", color: X.sub, padding: "22px 10px" }}>
                      No quote matches the filters.{" "}
                      <span onClick={() => { setPf({}); setPs(null); }} style={{ color: X.tealDk, cursor: "pointer", fontWeight: 600 }}>
                        Clear filters
                      </span>
                    </td>
                  </tr>
                )}
                {!rows.length && (
                  <tr>
                    <td colSpan={12} style={{ ...td, textAlign: "center", color: X.sub, padding: "28px 10px" }}>
                      {!cands.length ? (quotes && quotes.length ? "No open quote in this classification and age." : "Import the Quote Report to see the open quotes.") : "Nothing planned for this day."}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          )}
          <div style={{ padding: 10, textAlign: "center", borderTop: `1px solid ${X.line}` }}>
            <button onClick={() => setShowAll((v) => !v)} style={{ ...btn, height: 30, fontSize: 12 }}>
              {showAll ? "Show only the plan of the day" : `Show the other ${cands.length - list.length} open quotes`}
            </button>
          </div>
        </>
      )}
    </div>
  );
}

/* ===== botão de importar ===== */
function ImportButton({ label, icon: Ic, info, onFile, accent }) {
  const ref = useRef(null);
  return (
    <>
      <button
        onClick={() => ref.current && ref.current.click()}
        style={{ ...btn, background: accent ? X.teal : "#fff", color: accent ? "#fff" : X.ink, borderColor: accent ? X.teal : X.line }}
        title={info || `Import the ${label}`}
      >
        <Ic size={15} /> {label}
        {info && <Check size={14} color={accent ? "#fff" : X.tealDk} />}
      </button>
      <input
        ref={ref}
        type="file"
        accept=".xlsx,.xls,.xlsm,.csv"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files && e.target.files[0];
          e.target.value = "";
          if (f) onFile(f);
        }}
      />
    </>
  );
}

function DataMenu({ items }) {
  const p = usePop(260);
  return (
    <>
      <button ref={p.anchor} onClick={p.toggle} style={{ ...btn, padding: "0 10px" }} title="Data and backup">
        <MoreHorizontal size={17} />
      </button>
      {p.render(
        items.map((it, i) =>
          it === "-" ? (
            <div key={i} style={{ height: 1, background: X.line, margin: "4px 0" }} />
          ) : (
            <div
              key={it.label}
              className="qp-opt"
              onClick={() => {
                p.setPos(null);
                it.onClick();
              }}
              style={{ ...optS, color: it.danger ? X.redDk : X.ink, opacity: it.disabled ? 0.45 : 1, pointerEvents: it.disabled ? "none" : "auto" }}
            >
              {React.createElement(it.icon, { size: 15 })}
              {it.label}
            </div>
          ),
        ),
      )}
    </>
  );
}

/* ===== Lista do que ainda não foi salvo (abre ao clicar no aviso amarelo do arquivo de dados) =====
   detail = { quotes: [{ id, c, items: [registro do histórico | { k: "_meta" }] }], goals: [{ d, o, n }], report: [...] } */
const goalTxt = (v, isDay) => (v == null ? (isDay ? "daily goal" : "—") : v < 0 ? "no goal" : money(v));
function UnsavedList({ detail, onOpenQuote, onSave, saving }) {
  const today = todayISO();
  const when = (t) => (t ? `${String(t).slice(0, 10) === today ? "" : fmtDay(String(t).slice(0, 10)) + " "}${String(t).slice(11, 16)}` : "");
  const repQ = new Set((detail.report || []).map((a) => a.id)).size;
  const head = { fontSize: 10.5, fontWeight: 800, color: X.sub, textTransform: "uppercase", letterSpacing: 0.4, padding: "8px 8px 4px" };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <div style={{ padding: "6px 8px 4px", display: "flex", alignItems: "center", gap: 8 }}>
        <div>
          <div style={{ fontSize: 13.5, fontWeight: 800, color: X.ink }}>Changes not saved yet</div>
          <div style={{ fontSize: 11.5, color: X.sub }}>Kept in this browser until you click Save · click a QT to open it</div>
        </div>
        <button onClick={onSave} disabled={saving} style={{ ...btn, marginLeft: "auto", height: 32, fontSize: 12.5, fontWeight: 700, background: X.teal, borderColor: X.teal, color: "#fff" }}>
          <Save size={14} /> Save
        </button>
      </div>
      {detail.quotes.length > 0 && <div style={head}>Quotes · {detail.quotes.length}</div>}
      {detail.quotes.map((q) => (
        <div key={q.id} style={{ border: `1px solid ${X.line}`, borderRadius: 9, padding: "7px 9px", margin: "0 4px 6px", background: "#fff" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
            <span onClick={() => onOpenQuote(q.id)} title="Open the quote" style={{ fontWeight: 800, fontSize: 12.5, color: X.tealDk, cursor: "pointer", borderBottom: `1px dashed ${X.teal}` }}>
              {q.id}
            </span>
            <span style={{ fontSize: 12, color: X.sub, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 }}>{q.c}</span>
          </div>
          {q.items.map((a, i) => {
            const am = ACT_META[a.k] || (a.k === "_repair" ? { icon: History, title: "Restored from the history", color: X.violet, bg: X.violetSoft } : { icon: Database, title: "Follow-up data changed", color: X.sub, bg: X.head });
            return (
              <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 7, fontSize: 12, marginTop: 5 }}>
                <span style={{ width: 20, height: 20, borderRadius: 999, background: am.bg, color: am.color, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{React.createElement(am.icon, { size: 11 })}</span>
                <div style={{ minWidth: 0, flex: 1, lineHeight: 1.4 }}>
                  <b style={{ color: am.color }}>{am.title}</b>
                  {a.k === "_repair" ? <span style={{ color: X.sub }}> · FUP, % Win and date put back as the last entries of the Activities History say (someone had saved over them with old data)</span> : a.k === "_meta" ? <span style={{ color: X.sub }}> · FUP, % Win or date (changed and changed back, or from a Done)</span> : a.k === "note" ? <div style={{ color: X.ink, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{String(a.txt || "").slice(0, 220)}{String(a.txt || "").length > 220 ? "…" : ""}</div> : <span style={{ color: X.ink }}> · {actText(a)}</span>}
                </div>
                <span style={{ fontSize: 11, color: X.sub, whiteSpace: "nowrap" }}>{when(a.t)}</span>
              </div>
            );
          })}
        </div>
      ))}
      {detail.users && (
        <div style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, padding: "6px 10px" }}>
          <KeyRound size={13} color={X.violet} />
          <b>Users and passwords</b>
          <span style={{ color: X.sub }}>changed</span>
        </div>
      )}
      {detail.goals.length > 0 && <div style={head}>Goals · {detail.goals.length}</div>}
      {detail.goals.map((g) => (
        <div key={g.d || "daily"} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, padding: "4px 10px" }}>
          <Target size={13} color={X.amber} />
          <b>{g.d ? `Goal of ${dowName(g.d).slice(0, 3)}, ${fmtDate(g.d)}` : "Daily goal"}</b>
          <span style={{ color: X.sub, marginLeft: "auto", whiteSpace: "nowrap" }}>
            {goalTxt(g.o, !!g.d)} {arrow} <b style={{ color: X.ink }}>{goalTxt(g.n, !!g.d)}</b>
          </span>
        </div>
      ))}
      {repQ > 0 && (
        <>
          <div style={head}>Quote Report import</div>
          <div style={{ fontSize: 12, color: X.sub, padding: "2px 10px 6px" }}>
            {detail.report.length} history entries (status, value, new orders) on {repQ} quotes, written when the report was imported. They don't count in the number on the Save button.
          </div>
        </>
      )}
      {!detail.quotes.length && !detail.goals.length && !repQ && <div style={{ fontSize: 12, color: X.sub, padding: "6px 10px 10px" }}>Only the order of the history changed (e.g. something marked and unmarked).</div>}
    </div>
  );
}
/* ===== Login: escolher o usuário (Sales Person ou Admin) e a senha =====
   Na primeira vez de cada usuário o programa pede para criar a senha. O arquivo de dados guarda só o hash
   (veja hashPassword em core.js); por isso criar a senha exige a pasta de dados conectada. */
function LoginScreen({ names, users, connected, dfName, dfSt, onOpen, onReconnect, onLogin, onCreate, onLoadFile, busy }) {
  const fileRef = useRef(null);
  const [who, setWho] = useState(() => {
    try {
      const w = localStorage.getItem("qp_lastuser") || "";
      return w === ADMIN_USER ? "" : w;
    } catch {
      return "";
    }
  });
  // Admin é "secreto": não aparece na lista. Para mostrar: Ctrl+Shift+A ou duplo clique no ícone do programa
  const [adminOn, setAdminOn] = useState(false);
  const toggleAdmin = () => {
    setAdminOn((on) => {
      setWho(on ? "" : ADMIN_USER);
      setPwd("");
      setPwd2("");
      setErr("");
      return !on;
    });
  };
  useEffect(() => {
    const onKey = (e) => {
      if (e.ctrlKey && e.shiftKey && (e.key === "A" || e.key === "a")) {
        e.preventDefault();
        toggleAdmin();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);
  const [pwd, setPwd] = useState("");
  const [pwd2, setPwd2] = useState("");
  const [err, setErr] = useState("");
  const [checking, setChecking] = useState(false);
  const u = who ? users[who] : null;
  const hasPwd = !!(u && u.hash);
  const valid = who && ((who === ADMIN_USER && adminOn) || names.includes(who));
  const submit = (e) => {
    e && e.preventDefault();
    setErr("");
    if (!valid) return setErr("Choose your user.");
    if (hasPwd) {
      if (!pwd) return setErr("Type your password.");
      setChecking(true);
      setTimeout(() => {
        const ok = hashPassword(pwd, u.salt) === u.hash;
        setChecking(false);
        if (!ok) return setErr("Wrong password.");
        try {
          localStorage.setItem("qp_lastuser", who === ADMIN_USER ? "" : who);
        } catch {}
        onLogin(who);
      }, 20);
    } else {
      if (!connected) return setErr(dfSt === "nofsa" ? "The first password must be created on a computer with Chrome or Edge, with the data folder open." : "Open the data folder first: the password is saved in the shared data file.");
      if (pwd.length < 4) return setErr("The password needs at least 4 characters.");
      if (pwd !== pwd2) return setErr("The two passwords are not the same.");
      setChecking(true);
      setTimeout(() => {
        const salt = newSalt();
        // o papel (a tela) continua o que o Admin escolheu; quem é novo fica sem acesso até o Admin escolher
        onCreate(who, { role: who === ADMIN_USER ? "admin" : (u && u.role) || "", salt, hash: hashPassword(pwd, salt), at: nowStamp() });
        setChecking(false);
        try {
          localStorage.setItem("qp_lastuser", who === ADMIN_USER ? "" : who);
        } catch {}
      }, 20);
    }
  };
  const field = { ...inp, height: 42, fontSize: 14, width: "100%", borderRadius: 10 };
  const lbl = { fontSize: 12, fontWeight: 700, color: X.sub, marginBottom: 6, display: "block" };
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 14, background: `radial-gradient(1200px 600px at 20% 0%, #DDF1EC 0%, transparent 60%), radial-gradient(900px 500px at 100% 100%, #E6EEFB 0%, transparent 55%), ${X.bg}` }}>
      <form onSubmit={submit} style={{ ...card, width: "min(420px, 100%)", borderRadius: 18, padding: "24px 20px 20px", display: "flex", flexDirection: "column", gap: 16, boxShadow: "0 24px 60px rgba(16,24,40,.12)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span onDoubleClick={toggleAdmin} style={{ width: 44, height: 44, borderRadius: 12, background: X.teal, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", userSelect: "none" }}>
            <TrendingUp size={24} />
          </span>
          <div>
            <div style={{ fontSize: 20, fontWeight: 800, letterSpacing: -0.4 }}>Quote Follow Up</div>
            <div style={{ fontSize: 12.5, color: X.sub }}>Sign in to continue</div>
          </div>
        </div>
        <div>
          <label style={lbl}>User</label>
          <select
            value={valid ? who : ""}
            onChange={(e) => {
              setWho(e.target.value);
              setPwd("");
              setPwd2("");
              setErr("");
            }}
            style={{ ...field, fontWeight: who ? 600 : 400 }}
          >
            <option value="">Choose your name…</option>
            {adminOn && <option value={ADMIN_USER}>Admin · all salespeople</option>}
            {names.map((n) => (
              <option key={n} value={n}>
                {n}
              </option>
            ))}
          </select>
          {!names.length && <div style={{ fontSize: 11.5, color: X.sub, marginTop: 6 }}>The salespeople appear after the data folder is opened.</div>}
        </div>
        {valid && hasPwd && (
          <div>
            <label style={lbl}>Password</label>
            <input type="password" autoFocus value={pwd} onChange={(e) => setPwd(e.target.value)} style={field} autoComplete="current-password" />
          </div>
        )}
        {valid && !hasPwd && (
          <>
            <div style={{ display: "flex", gap: 10, alignItems: "flex-start", background: X.tealSoft, color: X.tealDk, borderRadius: 10, padding: "10px 12px", fontSize: 12.5, lineHeight: 1.45 }}>
              <KeyRound size={16} style={{ flexShrink: 0, marginTop: 2 }} />
              <span>
                <b>First access.</b> Create your password — only you will know it.
              </span>
            </div>
            <div>
              <label style={lbl}>New password</label>
              <input type="password" autoFocus value={pwd} onChange={(e) => setPwd(e.target.value)} style={field} autoComplete="new-password" />
            </div>
            <div>
              <label style={lbl}>Repeat the password</label>
              <input type="password" value={pwd2} onChange={(e) => setPwd2(e.target.value)} style={field} autoComplete="new-password" />
            </div>
          </>
        )}
        {err && <div style={{ fontSize: 12.5, fontWeight: 600, color: X.redDk, background: X.redSoft, borderRadius: 8, padding: "8px 10px" }}>{err}</div>}
        <button type="submit" disabled={checking || busy || !valid} style={{ ...btn, height: 44, justifyContent: "center", fontSize: 14, fontWeight: 700, borderRadius: 11, background: valid ? X.teal : "#C7D2D9", borderColor: valid ? X.teal : "#C7D2D9", color: "#fff", cursor: valid ? "pointer" : "not-allowed" }}>
          {checking || busy ? "Checking…" : valid && !hasPwd ? "Create password and sign in" : "Sign in"}
        </button>
        <div style={{ borderTop: `1px solid ${X.line}`, paddingTop: 12, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", fontSize: 12, color: X.sub }}>
          <Database size={14} />
          {connected ? (
            <span>
              Data folder <b style={{ color: X.ink }}>{dfName}</b> connected
            </span>
          ) : dfSt === "ask" ? (
            <button type="button" onClick={onReconnect} style={{ ...btn, height: 30, fontSize: 12 }}>
              Reconnect the data folder
            </button>
          ) : dfSt === "opening" ? (
            <span>Opening the data folder…</span>
          ) : dfSt === "nofsa" ? (
            /* celular e navegadores sem acesso a pastas: carregar o arquivo de dados à mão */
            <>
              <button type="button" onClick={() => fileRef.current && fileRef.current.click()} style={{ ...btn, height: 34, fontSize: 12 }}>
                <FolderOpen size={13} /> Load the data file…
              </button>
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx"
                multiple
                style={{ display: "none" }}
                onChange={async (e) => {
                  const fs = [...(e.target.files || [])];
                  e.target.value = "";
                  for (const f of fs) if (onLoadFile) await onLoadFile(f);
                }}
              />
              <span style={{ width: "100%", fontSize: 11, lineHeight: 1.4 }}>This browser can't keep the data folder connected (phones, Firefox, Safari): load the two files {DATA_FILE} and {REP_FILE}.</span>
            </>
          ) : (
            <button type="button" onClick={onOpen} style={{ ...btn, height: 30, fontSize: 12 }}>
              <FolderOpen size={13} /> Open the data folder…
            </button>
          )}
        </div>
        <div style={{ fontSize: 11, color: X.sub, lineHeight: 1.45 }}>The password is saved only as an encrypted code (hash) in the data file. Forgot it? Ask the admin to reset it.</div>
      </form>
    </div>
  );
}
/* ===== Admin: usuários e acesso =====
   O Admin escolhe o papel (a tela) de cada pessoa: vendedor externo, vendedor interno, gerência ou sem acesso.
   A lista vem das colunas Sales Person e Quoted By do Quote Report, mais as pessoas da gerência que o Admin
   acrescenta pelo nome (quem não aparece nos relatórios). Tirar o acesso de alguém acrescentado à mão tira a
   pessoa da lista. Reset = a pessoa cria outra senha no próximo login. */
function UsersPanel({ names, colOf, users, onRole, onAdd, onReset }) {
  const mob = useMobile();
  const [ask, setAsk] = useState("");
  const [q, setQ] = useState("");
  const [nf, setNf] = useState(""); // filtro: "" todos · "none" sem acesso · papel
  const [add, setAdd] = useState("");
  const roleOfU = (n) => (ROLES[(users[n] || {}).role] ? users[n].role : "");
  const cnt = { "": names.length, none: names.filter((n) => !roleOfU(n)).length };
  ROLE_KEYS.forEach((k) => (cnt[k] = names.filter((n) => roleOfU(n) === k).length));
  const ql = norm(q);
  const list = [ADMIN_USER, ...names].filter((n) => (!ql || norm(n).includes(ql)) && (!nf || (n !== ADMIN_USER && (nf === "none" ? !roleOfU(n) : roleOfU(n) === nf))));
  const addName = add.trim().replace(/\s+/g, " ");
  const addOk = !!addName && addName !== ADMIN_USER && ![...names].some((n) => norm(n) === norm(addName));
  const doAdd = () => {
    if (!addOk) return;
    onAdd(addName);
    setAdd("");
  };
  const chip = (k, l, fg) => {
    const on = nf === k;
    return (
      <button key={k || "all"} onClick={() => setNf(k)} style={{ ...btn, height: 30, padding: "0 10px", fontSize: 12, borderRadius: 999, background: on ? X.tealSoft : "#fff", color: on ? X.tealDk : fg || X.ink, borderColor: on ? "#9FD6C9" : X.line, fontWeight: on ? 700 : 600 }}>
        {l} <span style={{ opacity: 0.7 }}>{cnt[k] || 0}</span>
      </button>
    );
  };
  const where = (n) => {
    const c = colOf(n);
    return c.sp && c.qb ? "Sales Person and Quoted By" : c.sp ? "Sales Person" : c.qb ? "Quoted By" : "added by hand";
  };
  return (
    <div style={{ ...card, overflow: "hidden" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: mob ? "12px 12px" : "16px 18px", borderBottom: `1px solid ${X.line}` }}>
        <span style={{ width: 36, height: 36, borderRadius: 10, background: X.tealSoft, color: X.tealDk, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Users size={18} />
        </span>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 15.5, fontWeight: 800 }}>Users and access</div>
          <div style={{ fontSize: 12, color: X.sub }}>Choose the screen of each person · saved in the data file right away</div>
        </div>
      </div>
      <div style={{ padding: mob ? "10px 12px" : "12px 18px", display: "flex", flexDirection: "column", gap: 10, borderBottom: `1px solid ${X.line}` }}>
        <div className="qp-hscroll" style={{ display: "flex", gap: 6 }}>
          {chip("", "All")}
          {chip("none", "No access", cnt.none ? X.redDk : X.sub)}
          {ROLE_KEYS.map((k) => chip(k, ROLES[k].label))}
        </div>
        <div style={{ position: "relative" }}>
          <Search size={14} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: X.sub, pointerEvents: "none" }} />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search a person…" style={{ ...inp, width: "100%", height: 38, paddingLeft: 30 }} />
        </div>
      </div>
      <div>
        {list.map((n) => {
          const u = users[n];
          const has = !!(u && u.hash);
          const isA = n === ADMIN_USER;
          const r = roleOfU(n);
          const R = ROLES[r];
          const hand = !isA && where(n) === "added by hand";
          const roleSel = isA ? (
            <span style={{ fontSize: 12, fontWeight: 700, color: ROLES.manager.fg }}>Admin · Management screen</span>
          ) : (
            <select
              value={r}
              onChange={(e) => onRole(n, e.target.value)}
              title="Screen this person sees after signing in"
              style={{ ...inp, height: 36, width: "100%", fontWeight: 700, color: R ? R.fg : X.redDk, background: R ? R.bg : "#FFF7F7", borderColor: R ? "transparent" : "#F3B4B4" }}
            >
              <option value="">{hand ? "No access (remove)" : "No access"}</option>
              {ROLE_KEYS.map((k) => (
                <option key={k} value={k}>
                  {ROLES[k].label}
                </option>
              ))}
            </select>
          );
          const pwd = <span style={{ fontSize: 12, color: has ? X.green : X.sub, whiteSpace: "nowrap" }}>{has ? "password set" : "no password yet"}</span>;
          const reset =
            has &&
            (ask === n ? (
              <button
                onClick={() => {
                  onReset(n);
                  setAsk("");
                }}
                style={{ ...btn, height: 32, fontSize: 12, background: X.redDk, borderColor: X.redDk, color: "#fff" }}
              >
                Confirm reset
              </button>
            ) : (
              <button onClick={() => setAsk(n)} style={{ ...btn, height: 32, fontSize: 12 }} title="The person creates a new password at the next sign-in">
                <RotateCcw size={13} /> Reset
              </button>
            ));
          const who = (
            <span style={{ display: "flex", alignItems: "center", gap: 9, minWidth: 0 }}>
              <span style={{ width: 30, height: 30, borderRadius: 999, background: isA ? X.violetSoft : R ? R.bg : "#F1F4F7", color: isA ? X.violet : R ? R.fg : X.sub, fontSize: 10.5, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                {isA ? <ShieldCheck size={14} /> : initials(n)}
              </span>
              <span style={{ minWidth: 0 }}>
                <span style={{ display: "block", fontWeight: 700, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{n}</span>
                <span style={{ display: "block", fontSize: 11, color: X.sub }}>{isA ? "secret user · Ctrl+Shift+A on the sign-in" : hand ? "added by hand (not in the reports)" : `in the report as ${where(n)}`}</span>
              </span>
            </span>
          );
          return mob ? (
            <div key={n} style={{ padding: "11px 12px", borderBottom: `1px solid #F0F3F6`, display: "flex", flexDirection: "column", gap: 8 }}>
              {who}
              {roleSel}
              <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "space-between" }}>
                {pwd}
                {reset}
              </div>
            </div>
          ) : (
            <div key={n} style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) 200px 120px 130px", alignItems: "center", gap: 12, padding: "9px 18px", borderBottom: `1px solid #F0F3F6` }}>
              {who}
              {roleSel}
              {pwd}
              <span style={{ textAlign: "right" }}>{reset}</span>
            </div>
          );
        })}
        {!list.length && <div style={{ padding: 20, textAlign: "center", color: X.sub, fontSize: 12.5 }}>Nobody found.</div>}
      </div>
      <div style={{ padding: mob ? "12px 12px" : "12px 18px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", background: "#FAFBFC" }}>
        <UserPlus size={16} color={X.sub} />
        <span style={{ fontSize: 12.5, fontWeight: 600, color: X.ink }}>Someone who is not in the reports?</span>
        <input value={add} onChange={(e) => setAdd(e.target.value)} onKeyDown={(e) => e.key === "Enter" && doAdd()} placeholder="Full name" style={{ ...inp, height: 36, flex: "1 1 180px", minWidth: 0 }} />
        <button onClick={doAdd} disabled={!addOk} style={{ ...btn, height: 36, background: addOk ? X.teal : "#fff", color: addOk ? "#fff" : X.sub, borderColor: addOk ? X.teal : X.line, cursor: addOk ? "pointer" : "not-allowed" }} title="Adds the person with the Management screen (change it in the list)">
          <UserPlus size={14} /> Add
        </button>
      </div>
    </div>
  );
}
/* Sem acesso: a pessoa entrou, mas o Admin ainda não escolheu a tela dela */
function NoAccessScreen({ name, connected, checking, onCheck, onLogout }) {
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, background: `radial-gradient(1200px 600px at 20% 0%, #DDF1EC 0%, transparent 60%), ${X.bg}` }}>
      <div style={{ ...card, width: "min(420px, 100%)", borderRadius: 18, padding: "26px 22px 20px", display: "flex", flexDirection: "column", gap: 14, textAlign: "center", alignItems: "center" }}>
        <span style={{ width: 52, height: 52, borderRadius: 14, background: X.amberSoft, color: X.amber, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
          <Ban size={26} />
        </span>
        <div>
          <div style={{ fontSize: 18, fontWeight: 800 }}>Hi, {String(name).split(" ")[0]}</div>
          <div style={{ fontSize: 13, color: X.sub, marginTop: 6, lineHeight: 1.5 }}>
            Your access is not set yet. Ask the admin to choose your screen: <b>outside sales</b>, <b>inside sales</b> or <b>management</b>.
          </div>
        </div>
        {connected ? (
          <button onClick={onCheck} disabled={checking} style={{ ...btn, height: 44, width: "100%", justifyContent: "center", fontSize: 14, background: X.teal, borderColor: X.teal, color: "#fff" }}>
            <RefreshCw size={15} /> {checking ? "Checking…" : "The admin set it · check again"}
          </button>
        ) : (
          <div style={{ fontSize: 12, color: X.sub }}>Open the data folder on this computer to receive the access.</div>
        )}
        <button onClick={onLogout} style={{ ...btn, height: 40, width: "100%", justifyContent: "center" }}>
          <LogOut size={15} /> Log out
        </button>
      </div>
    </div>
  );
}
/* Usuário logado (no topo): nome e papel; o menu tem "Users and access" (gerência) e "Log out" */
/* items = ações extras do menu (no celular: atualizar, salvar, pasta de dados); status = situação dos dados */
function UserChip({ name, admin, role, compact, onLogout, onUsers, items, status }) {
  const p = usePop(compact ? 280 : 230, 480);
  const R = ROLES[role];
  return (
    <>
      <button ref={p.anchor} onClick={p.toggle} style={{ ...btn, height: compact ? 38 : 40, padding: compact ? "0 6px" : "0 10px 0 6px", borderRadius: 11, gap: compact ? 4 : 8 }} title={`${admin ? "Admin" : name}${R ? " · " + R.label : ""}`}>
        <span style={{ width: 28, height: 28, borderRadius: 999, background: admin ? X.violetSoft : R ? R.bg : X.tealSoft, color: admin ? X.violet : R ? R.fg : X.tealDk, fontSize: 10.5, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
          {admin ? <ShieldCheck size={15} /> : initials(name)}
        </span>
        {!compact && (
          <span style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", lineHeight: 1.15 }}>
            <span style={{ fontSize: 12.5, fontWeight: 700, maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{admin ? "Admin" : name.split(" ")[0]}</span>
            <span style={{ fontSize: 10.5, color: X.sub }}>{R ? R.label : "no access"}</span>
          </span>
        )}
        <ChevronDown size={13} color="#A0AEC0" />
      </button>
      {p.render(
        <>
          <div style={{ padding: "8px 9px 6px", borderBottom: `1px solid ${X.line}`, marginBottom: 4 }}>
            <div style={{ fontSize: 13, fontWeight: 700 }}>{admin ? "Admin" : name}</div>
            <div style={{ fontSize: 11.5, color: R ? R.fg : X.sub, fontWeight: 600 }}>{R ? R.label : "no access"}</div>
            {status && <div style={{ fontSize: 11.5, color: X.sub, marginTop: 5, lineHeight: 1.35 }}>{status}</div>}
          </div>
          {(items || []).filter(Boolean).map((it) => (
            <div
              key={it.label}
              className="qp-opt"
              onClick={() => {
                p.setPos(null);
                it.onClick();
              }}
              style={{ ...optS, padding: "10px 9px", fontSize: 13.5, color: it.color || X.ink }}
            >
              <it.icon size={16} /> {it.label}
            </div>
          ))}
          {onUsers && (
            <div
              className="qp-opt"
              onClick={() => {
                p.setPos(null);
                onUsers();
              }}
              style={{ ...optS, ...(compact ? { padding: "10px 9px", fontSize: 13.5 } : null) }}
            >
              <Users size={15} /> Users and access…
            </div>
          )}
          <div
            className="qp-opt"
            onClick={() => {
              p.setPos(null);
              onLogout();
            }}
            style={{ ...optS, ...(compact ? { padding: "10px 9px", fontSize: 13.5, borderTop: `1px solid ${X.line}`, borderRadius: 0, marginTop: 4 } : null) }}
          >
            <LogOut size={15} /> Log out
          </div>
        </>,
      )}
    </>
  );
}
/* ===== arquivo de dados (Excel) conectado =====
   Usa a File System Access API do Chrome/Edge (a mesma do pulz). A referência do arquivo fica no
   navegador (IndexedDB) para reconectar com um clique; o arquivo é reescrito a cada mudança. */
const FSA = typeof window !== "undefined" && "showDirectoryPicker" in window;
const HANDLE_KEY = "qp_datafile";
const DATA_FILE = "Quote Follow Up data.xlsx"; // FUP, % Win, requisição, notas, histórico, metas
const REP_FILE = "Quote Follow Up reports.xlsx"; // cotações e vendas dos relatórios
// nomes antigos (o programa se chamava Quote Pipeline): ao abrir a pasta, os arquivos são renomeados para os novos
const OLD_FILES = { data: "Quote Pipeline data.xlsx", rep: "Quote Pipeline reports.xlsx" };
const XLSX_TYPE = [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }];
const SAVE_DELAY = 1500; // ms depois da última mudança
/* Gravação do arquivo de dados: compactado (o arquivo fica ~3× menor e grava mais rápido). */
const DATA_WRITE = { type: "array", bookType: "xlsx", compression: true, bookSST: true };
/* Erro do arquivo de dados em palavras: o que o navegador disse + o que fazer. */
const fileErr = (e) => {
  const name = (e && e.name) || "Error";
  const msg = (e && e.message) || String(e);
  const hint =
    /InvalidState|NotReadable/.test(name)
      ? "Windows refused the last step of the save (replacing the file). This happens mostly with files on a network drive; it can also be the file open in Excel or a sync/antivirus program. Nothing is lost: the data stays in this browser. Save the data file in a local folder, e.g. Documents (button on the right)."
      : /NoModificationAllowed|InvalidModification/.test(name)
      ? "The file is locked by another program (Excel, OneDrive/SharePoint sync, antivirus or the network drive). Close it in Excel; if it keeps happening, save the data file in a local folder (button on the right)."
      : /NotAllowed|Security/.test(name)
      ? "The browser did not allow saving in this folder. Pick a local folder (e.g. Documents) for the data file."
      : /NotFound/.test(name)
      ? "The file was moved, renamed or deleted. Open it again from the ⋯ menu."
      : /NotReadable/.test(name)
      ? "The browser could not read the file: it is locked by another program or the network drive is unavailable."
      : /QuotaExceeded/.test(name)
      ? "There is no free space to save the file."
      : "If the file is on a network drive, try keeping it in a local folder (e.g. Documents).";
  return { text: `${name}: ${msg}`, hint };
};
/* "Assinatura" do que vai para o arquivo: se não mudou desde a última gravação, não grava de novo. */
const dataSig = (st) => {
  if (!st) return "";
  const q = st.quotesRep,
    sr = st.salesRep,
    a = st.acts || [];
  return [
    q ? `${q.quotes.length}|${q.at || ""}|${q.minD || ""}|${q.maxD || ""}|${q.fillAt || ""}` : "-",
    sr ? `${(sr.rows || []).length}|${sr.at || ""}|${sr.minD || ""}|${sr.maxD || ""}` : "-",
    stableJson(st.meta || {}),
    [...(st.excl || [])].sort().join(","),
    `${a.length}|${a.reduce((m, x) => (x.t > m ? x.t : m), "")}`,
    st.dailyGoal || 0,
    stableJson(st.dayGoals || {}),
    stableJson(st.users || {}),
  ].join("#");
};
// JSON com as chaves em ordem (o mesmo conteúdo sempre dá o mesmo texto)
const stableJson = (v) =>
  v && typeof v === "object" && !Array.isArray(v)
    ? "{" + Object.keys(v).sort().map((k) => JSON.stringify(k) + ":" + stableJson(v[k])).join(",") + "}"
    : JSON.stringify(v === undefined ? null : v);
const DAY_NO_GOAL = -1; // meta do dia = "sem meta" digitada à mão num dia útil (0 digitado)
const PULL_EVERY = 60000; // ms entre as olhadas nos arquivos de dados (mudanças salvas por outra pessoa)
const RETRY_DELAY = 60000; // ms para tentar de novo depois de um erro (cada tentativa deixa um .crswap se o disco recusar)

/* ===== aplicação ===== */
export default function App() {
  const today = todayISO();
  const [quotesRep, setQuotesRep] = useState(null);
  const [salesRep, setSalesRep] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [busy, setBusy] = useState("");
  const [toast, setToast] = useState(null);
  const [meta, setMeta] = useStored(K.meta, {}); // { [Quote ID]: { fup, win, since, name, log } } · o Thermometer é automático
  const [dailyGoal, setDailyGoal] = useStored(K.dailyGoal, 0); // meta de cada dia útil
  const [dayGoals, setDayGoals] = useStored(K.dayGoals, {}); // meta diária digitada à mão { "AAAA-MM-DD": valor }
  /* Correção: versões anteriores gravavam meta 0 num dia só por clicar no campo da meta do dia e sair.
     Esses dias ficavam presos em 0 mesmo depois de digitar a meta diária. Agora "sem meta" num dia útil é
     gravado como -1 (DAY_NO_GOAL); os 0 que sobraram das versões anteriores são apagados (o dia volta a
     seguir a meta diária). */
  useEffect(() => {
    const bad = Object.keys(dayGoals || {}).filter((d) => dayGoals[d] === 0 || dayGoals[d] == null);
    if (!bad.length) return;
    setDayGoals((g) => {
      const x = { ...g };
      bad.forEach((d) => delete x[d]);
      return x;
    });
  }, [dayGoals]);
  const [users, setUsers] = useStored(K.users, {}); // usuários e o hash das senhas (vai para o arquivo de dados)
  const [excl, setExcl] = useStored(K.excl, []); // cotações tiradas do plano
  const [acts, setActs] = useStored(K.acts, []); // Activities History + anotações (só acrescenta, nunca altera)
  const [user, setUser] = useStored(K.user, ""); // nome de quem usa este navegador (vai em cada registro)
  const [prefs, setPrefs] = useStored(K.prefs, { tab: "portfolio", period: "fy:" + (today.slice(5, 7) >= "07" ? +today.slice(0, 4) : +today.slice(0, 4) - 1) });
  const [openQ, setOpenQ] = useState(""); // cotação aberta na gaveta lateral
  const [openDone, setOpenDone] = useState(false); // gaveta aberta pelo botão Done
  const [df, setDf] = useState({ st: FSA ? "none" : "nofsa" });
  const [dfErr, setDfErr] = useState(null); // erro do arquivo de dados que fica na tela até fechar // arquivo de dados: none | ask | on | saving | err | nofsa
  const setPref = (k, v) => setPrefs((p) => ({ ...p, [k]: v }));
  const restoreRef = useRef(null);
  const loadDataRef = useRef(null);
  const handleRef = useRef(null); // pasta dos arquivos de dados (FileSystemDirectoryHandle)
  const skipSave = useRef(0);
  const stateRef = useRef(null);
  /* Dois arquivos na pasta:
     • "data"    = Quote Follow Up data.xlsx    → FUP, % Win, requisição, notas, histórico, metas (pequeno; vai em todo Save)
     • "reports" = Quote Follow Up reports.xlsx → cotações e vendas dos relatórios (grande; só grava quando mudou = depois de importar) */
  const savedSig = useRef({ data: "", rep: "" }); // assinatura do que está em cada arquivo
  const lastMod = useRef({ data: 0, rep: 0 }); // data de modificação de cada arquivo na última leitura/gravação
  const savedActs = useRef(new Set()); // atividades (chaves) que havia no arquivo "data" na última leitura/gravação
  // o que está no arquivo "data" (para contar as cotações alteradas desde o último Save)
  const savedSnap = useRef({ meta: {}, excl: [], dailyGoal: 0, dayGoals: {}, users: {} });
  const needSplit = useRef(false); // arquivo "data" antigo (tudo junto): o próximo Save separa em dois
  /* espera (ms) antes de trocar o arquivo que funcionou (rede/antivírus), por arquivo. Fica guardada neste
     navegador: na próxima vez a gravação já começa com ela e passa de primeira, sem deixar .crswap. */
  const closeWait = useRef(
    (() => {
      const get = (k) => {
        try {
          return Math.min(30000, Math.max(0, +localStorage.getItem(k) || 0));
        } catch {
          return 0;
        }
      };
      return { data: get("qp_closewait_data"), rep: get("qp_closewait_rep") || get("qp_closewait") };
    })(),
  );
  const writing = useRef(false); // uma gravação por vez
  const syncOk = useRef(true);
  const again = useRef(false);
  stateRef.current = { quotesRep, salesRep, meta, excl, acts, dailyGoal, dayGoals, users };

  const say = (tone, text) => setToast({ tone, text });

  /* ---------- arquivos de dados: ler, juntar, gravar ---------- */
  const partOf = (st, kind) => (kind === "rep" ? { quotesRep: st.quotesRep, salesRep: st.salesRep } : { meta: st.meta, excl: st.excl, acts: st.acts, dailyGoal: st.dailyGoal, dayGoals: st.dayGoals, users: st.users || {} });
  const sigOf = (st, kind) => dataSig(partOf(st, kind));
  const fileName = (kind) => (kind === "rep" ? REP_FILE : DATA_FILE);
  const getFH = async (kind, create, dir) => {
    try {
      return await (dir || handleRef.current).getFileHandle(fileName(kind), { create: !!create });
    } catch (e) {
      if (e && (e.name === "NotFoundError" || e.name === "TypeMismatchError")) return null;
      throw e;
    }
  };
  const newerFile = useRef(""); // versão do programa que salvou o arquivo, quando é mais nova que esta
  const readFileObj = async (f) => {
    if (!f.size) return null;
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
    if (!isDataWorkbook(wb)) throw new Error(`${f.name} is not a Quote Follow Up data file`);
    const d = parseDataWorkbook(wb);
    if (d.ver && d.ver > APP_VERSION && !newerFile.current) {
      newerFile.current = d.ver;
      setDfErr({
        kind: "save",
        title: "A newer version of Quote Follow Up saved the data files",
        text: `${f.name} was saved by the version of ${d.ver}; this page runs the version of ${APP_VERSION}.`,
        hint: "Reload this page (F5) — or open the new Quote Follow Up — before saving, so the old version does not write over the new data. Saving is blocked on this page.",
      });
    }
    return d;
  };
  const readPart = async (kind, dir) => {
    const fh = await getFH(kind, false, dir);
    if (!fh) return null;
    const f = await fh.getFile();
    const d = await readFileObj(f);
    lastMod.current[kind] = f.lastModified;
    return d;
  };
  /* Lê os dois arquivos da pasta. O "data" de antes (com tudo junto) também serve: as cotações dele são
     usadas enquanto não existir o "reports"; no próximo Save elas passam para o "reports". */
  // pasta com os arquivos do nome antigo (Quote Pipeline …): renomeia para Quote Follow Up … (conteúdo igual)
  const renameOld = async (dir) => {
    for (const kind of ["data", "rep"]) {
      let oldH = null;
      try {
        oldH = await dir.getFileHandle(OLD_FILES[kind]);
      } catch {
        continue; // não existe o antigo
      }
      try {
        await dir.getFileHandle(fileName(kind));
        continue; // o novo já existe: não mexe em nada
      } catch {}
      try {
        if (typeof oldH.move === "function") {
          await oldH.move(fileName(kind));
        } else {
          // sem "renomear" no navegador: copia byte a byte e só então apaga o antigo
          const buf = await (await oldH.getFile()).arrayBuffer();
          const nh = await dir.getFileHandle(fileName(kind), { create: true });
          const w = await nh.createWritable();
          await w.write(buf);
          await w.close();
          if ((await nh.getFile()).size === buf.byteLength) await dir.removeEntry(OLD_FILES[kind]);
        }
        console.info(`Quote Follow Up · renamed ${OLD_FILES[kind]} → ${fileName(kind)}`);
      } catch (e) {
        console.warn(`Quote Follow Up · could not rename ${OLD_FILES[kind]}`, e);
      }
    }
  };
  const readFolder = async (dir) => {
    await renameOld(dir);
    const dd = await readPart("data", dir);
    const dr = await readPart("rep", dir);
    if (!dd && !dr) return null;
    const legacy = !!(dd && dd.quotesRep && !dr);
    const repSrc = dr || (legacy ? dd : null);
    const d = {
      meta: (dd && dd.meta) || {},
      excl: (dd && dd.excl) || [],
      acts: (dd && dd.acts) || [],
      dailyGoal: (dd && dd.dailyGoal) || 0,
      dayGoals: (dd && dd.dayGoals) || {},
      users: (dd && dd.users) || {},
      quotesRep: (repSrc && repSrc.quotesRep) || null,
      salesRows: (repSrc && repSrc.salesRows) || [],
      salesInfo: (repSrc && repSrc.salesInfo) || {},
    };
    return { d, dd, dr, legacy };
  };
  const applyData = async (d) => {
    skipSave.current = 1;
    const sales = d.salesRows && d.salesRows.length ? buildSalesIndex(d.salesRows, d.salesInfo) : null;
    setQuotesRep(d.quotesRep);
    setSalesRep(sales);
    setMeta(d.meta || {});
    setExcl(d.excl || []);
    setActs(d.acts || []);
    setDailyGoal(d.dailyGoal || 0);
    setDayGoals(d.dayGoals || {});
    setUsers(d.users || {});
    if (d.quotesRep) await kv.set(K.quotes, d.quotesRep);
    else await kv.del(K.quotes);
    if (sales) await kv.set(K.sales, sales);
    else await kv.del(K.sales);
  };
  /* Grava um dos arquivos. O navegador escreve primeiro num temporário (".crswap") e, no fim, troca pelo
     arquivo; se der erro, o temporário é descartado (abort). */
  const writeHandle = async (kind, st, wait) => {
    const data = partOf(st || stateRef.current, kind);
    const fh = await getFH(kind, true);
    // cronômetro de cada etapa (aparece no console): mostra onde a rede/antivírus demora
    const T = [];
    let t0 = performance.now();
    const lap = (name) => {
      const t = performance.now();
      T.push(`${name} ${((t - t0) / 1000).toFixed(1)}s`);
      t0 = t;
    };
    const buf = XLSX.write(buildDataWorkbook(data), DATA_WRITE);
    lap(`build Excel (${(buf.byteLength / 1e6).toFixed(2)} MB)`);
    const w = await fh.createWritable();
    lap("open temporary file");
    try {
      await w.write(buf);
      lap("write");
      if (wait) await new Promise((r) => setTimeout(r, wait));
      if (wait) lap("wait");
      try {
        await w.close();
      } finally {
        lap("replace the file (close)");
        console.info(`Quote Follow Up · ${fileName(kind)} · save steps: ` + T.join(" · "));
      }
    } catch (e) {
      try {
        await w.abort();
      } catch {}
      throw e;
    }
    savedSig.current[kind] = dataSig(data);
    if (kind === "data") {
      needSplit.current = false;
      savedActs.current = new Set((data.acts || []).map(actKey));
      savedSnap.current = { meta: data.meta || {}, excl: data.excl || [], dailyGoal: data.dailyGoal || 0, dayGoals: data.dayGoals || {}, users: data.users || {} };
    }
    try {
      lastMod.current[kind] = (await fh.getFile()).lastModified;
    } catch {}
  };
  /* Trazer as mudanças de outra pessoa: a cada minuto (e quando a janela volta ao foco) olha a data de
     modificação dos arquivos. Se alguém salvou, lê e junta: o que você ainda não salvou continua aqui
     (e continua contando no botão Save); o resto passa a ser o que está no arquivo. */
  const pulling = useRef(false);
  const [checkSt, setCheckSt] = useState({ busy: false, at: "" }); // botão Refresh: conferindo agora / hora da última conferida
  const pullIfChanged = async (manual) => {
    const dir = handleRef.current;
    if (!dir || dir.kind !== "directory" || !syncOk.current) return;
    if (writing.current || pulling.current) {
      if (manual === true) say("info", writing.current ? "Saving now: the files are checked again right after." : "Already checking the data files…");
      return;
    }
    pulling.current = true;
    if (manual === true) setCheckSt((x) => ({ ...x, busy: true }));
    try {
      const got = [];
      for (const kind of ["data", "rep"]) {
        const fh = await getFH(kind, false);
        if (!fh || writing.current) continue;
        const f = await fh.getFile();
        if (!lastMod.current[kind] || f.lastModified === lastMod.current[kind] || !f.size) continue;
        const d = await readFileObj(f);
        if (!d || writing.current) continue;
        const local = stateRef.current;
        if (kind === "data") {
          const fileD = { meta: d.meta || {}, excl: d.excl || [], acts: d.acts || [], dailyGoal: d.dailyGoal || 0, dayGoals: d.dayGoals || {}, users: d.users || {} };
          const base = savedSnap.current;
          const n = new Set([...Object.keys(fileD.meta), ...Object.keys(base.meta || {})].filter((id) => stableJson(fileD.meta[id] || null) !== stableJson((base.meta || {})[id] || null))).size;
          const mine = sigOf(local, "data") !== savedSig.current.data;
          const m = mine ? mergeData({ ...fileD, quotesRep: local.quotesRep, salesRows: [] }, local, true, base) : fileD;
          skipSave.current = 1;
          setMeta(m.meta || {});
          setExcl(m.excl || []);
          setActs(m.acts || []);
          setDailyGoal(m.dailyGoal || 0);
          setDayGoals(m.dayGoals || {});
          // usuários: sempre somados aos que este PC já conhece (apagar no arquivo não tira ninguém; o mais recente vale)
          const mu = mergeUsers(m.users || {}, local.users || {});
          setUsers(mu);
          stateRef.current = { ...local, meta: m.meta, excl: m.excl, acts: m.acts, dailyGoal: m.dailyGoal, dayGoals: m.dayGoals, users: mu };
          if (d.usersBad) setTimeout(() => say("warn", "The users list in the data file was changed outside the program and could not be read — kept the users this computer knows. Save to write it again."), 1500);
          savedSig.current.data = sigOf(fileD, "data");
          savedSnap.current = { meta: fileD.meta, excl: fileD.excl, dailyGoal: fileD.dailyGoal, dayGoals: fileD.dayGoals, users: fileD.users };
          savedActs.current = new Set(fileD.acts.map(actKey));
          needSplit.current = false;
          got.push(n ? `${n} quote${n > 1 ? "s" : ""} with new follow-up data` : "follow-up data");
        } else {
          const fileR = { quotesRep: d.quotesRep || null, salesRows: d.salesRows || [], salesInfo: d.salesInfo || {} };
          const mine = sigOf(local, "rep") !== savedSig.current.rep;
          let m;
          if (mine) m = mergeData({ ...fileR, meta: {}, acts: [] }, local, true);
          else {
            // sem mudança aqui: vale o arquivo, mas os Product Groups que só este navegador tem não se perdem
            const lp = new Map(((local.quotesRep && local.quotesRep.quotes) || []).filter((q) => q.pg && Object.keys(q.pg).length).map((q) => [q.id, q.pg]));
            let n = 0;
            const qs = ((fileR.quotesRep && fileR.quotesRep.quotes) || []).map((q) => (!(q.pg && Object.keys(q.pg).length) && lp.has(q.id) ? (n++, { ...q, pg: lp.get(q.id) }) : q));
            m = { ...fileR, quotesRep: fileR.quotesRep ? { ...fileR.quotesRep, quotes: qs } : null, pgFromLocal: n };
          }
          const sales = m.salesRows && m.salesRows.length ? buildSalesIndex(m.salesRows, m.salesInfo) : null;
          skipSave.current = 1;
          setQuotesRep(m.quotesRep);
          setSalesRep(sales);
          if (m.quotesRep) await kv.set(K.quotes, m.quotesRep);
          if (sales) await kv.set(K.sales, sales);
          stateRef.current = { ...stateRef.current, quotesRep: m.quotesRep, salesRep: sales };
          // arquivo sem os Product Groups que este navegador tem: fica "não salvo" para o próximo Save regravar com eles
          savedSig.current.rep = sigOf({ quotesRep: fileR.quotesRep, salesRep: fileR.salesRows.length ? buildSalesIndex(fileR.salesRows, fileR.salesInfo) : null }, "rep") + (m.pgFromLocal ? "|pg" : "");
          if (m.pgFromLocal) console.info(`Quote Follow Up · ${REP_FILE} had no Product Groups for ${m.pgFromLocal} quotes: kept the ones in this browser (Save writes them back)`);
          got.push("quotes and sales reports");
        }
        lastMod.current[kind] = f.lastModified;
      }
      if (got.length) {
        console.info("Quote Follow Up · loaded changes saved by someone else: " + got.join(", "));
        say("info", `Someone else saved the data files: ${got.join(" and ")} loaded. Your unsaved changes are still here.`);
        setDf((x) => ({ ...x, openedAt: nowStamp() }));
      } else if (manual === true) say("ok", "You already have the latest data: nobody saved the data files since you opened or saved them.");
      setCheckSt({ busy: false, at: nowStamp() });
    } catch (e) {
      console.warn("Quote Follow Up · could not check the data files for changes", e);
      if (manual === true) say("err", `Could not read the data files: ${(e && e.message) || e}`);
      setCheckSt((x) => ({ ...x, busy: false }));
    } finally {
      pulling.current = false;
    }
  };
  const pullRef = useRef(null);
  pullRef.current = pullIfChanged;
  useEffect(() => {
    const run = () => pullRef.current && pullRef.current(false);
    const iv = setInterval(run, PULL_EVERY);
    const vis = () => document.visibilityState === "visible" && run();
    window.addEventListener("focus", run);
    document.addEventListener("visibilitychange", vis);
    return () => {
      clearInterval(iv);
      window.removeEventListener("focus", run);
      document.removeEventListener("visibilitychange", vis);
    };
  }, []);
  /* Antes de gravar: se o arquivo mudou por fora (outra aba, outro computador, Excel), lê e junta com o
     navegador (vale o que está no navegador nos campos manuais) para não apagar nada de ninguém. */
  const joinIfChanged = async (kind) => {
    try {
      const fh = await getFH(kind, false);
      if (!fh) return null;
      const f = await fh.getFile();
      if (!lastMod.current[kind] || f.lastModified === lastMod.current[kind] || !f.size) return null;
      console.info(`Quote Follow Up · ${f.name} changed outside the program: reading it before saving`);
      const d = await readFileObj(f);
      if (!d) return null;
      const local = stateRef.current;
      // "data": vale o que ESTE navegador mudou desde a última leitura; o resto vem do arquivo (mudanças de outra pessoa)
      const m = mergeData(d, local, true, kind === "data" ? savedSnap.current : null);
      await applyData(m);
      say("info", `${f.name} was saved by someone else since you opened it. Their changes were kept and joined with yours before saving.`);
      return { ...local, ...m, salesRep: m.salesRows.length ? buildSalesIndex(m.salesRows, m.salesInfo) : null };
    } catch (e) {
      console.warn("Quote Follow Up · could not check the data file before saving", e);
      return null;
    }
  };
  /* Grava com novas tentativas. Em unidade de rede o Windows pode recusar a troca do arquivo (o Chrome
     mostra InvalidStateError) enquanto o antivírus examina o arquivo recém-escrito. Por isso, a partir da
     2ª tentativa o programa escreve e ESPERA (3 s, 8 s, 15 s, 25 s) antes de pedir a troca. */
  const RETRY_WAITS = [1500, 3000, 6000];
  const CLOSE_WAITS = [3000, 8000, 15000, 25000];
  const writeSafe = async (kind, st) => {
    const first = closeWait.current[kind];
    const waits = [first, ...CLOSE_WAITS.filter((x) => x > first)].slice(0, 4);
    while (waits.length < 3) waits.push(waits[waits.length - 1]);
    for (let i = 0; ; i++) {
      try {
        await writeHandle(kind, st, waits[i]);
        if (i) console.info(`Quote Follow Up · ${fileName(kind)} saved on attempt ${i + 1} (waited ${waits[i] / 1000}s before replacing the file)`);
        if (waits[i] !== closeWait.current[kind]) {
          closeWait.current[kind] = waits[i];
          try {
            localStorage.setItem(kind === "rep" ? "qp_closewait_rep" : "qp_closewait_data", String(waits[i]));
          } catch {}
        }
        return;
      } catch (e) {
        const retry = /InvalidState|NotReadable|NoModificationAllowed|InvalidModification/.test((e && e.name) || "");
        if (!retry || i >= waits.length - 1) throw e;
        console.info(`Quote Follow Up · ${fileName(kind)} · save attempt ${i + 1} refused by the drive (${e.name}); next attempt in ${RETRY_WAITS[Math.min(i, RETRY_WAITS.length - 1)] / 1000}s, waiting ${waits[i + 1] / 1000}s before replacing the file`);
        await new Promise((r) => setTimeout(r, RETRY_WAITS[Math.min(i, RETRY_WAITS.length - 1)]));
        st = (await joinIfChanged(kind)) || st;
      }
    }
  };
  /* Save: grava o "data" (pequeno) e, só se as cotações/vendas mudaram (importação), o "reports".
     force = grava o "data" mesmo sem mudança (clique em Saved). */
  const saveNow = async (force) => {
    if (!handleRef.current || writing.current) return;
    if (newerFile.current) {
      say("err", `Saving is blocked: the data files were saved by a newer version of Quote Follow Up (${newerFile.current}). Reload this page (F5).`);
      return;
    }
    const st0 = stateRef.current;
    const needD = sigOf(st0, "data") !== savedSig.current.data || force === true;
    const needR = sigOf(st0, "rep") !== savedSig.current.rep;
    if (!needD && !needR) {
      setDf((d) => (d.st === "saving" || d.st === "err" ? { ...d, st: "on", err: "" } : d));
      return;
    }
    writing.current = true;
    const tAll = performance.now();
    let cur = "";
    try {
      for (const kind of ["data", "rep"]) {
        if (kind === "data" ? !needD : !needR) continue;
        cur = kind;
        setDf((d) => ({ ...d, st: "saving", part: kind }));
        await writeSafe(kind, (await joinIfChanged(kind)) || stateRef.current);
      }
      console.info(`Quote Follow Up · saved in ${((performance.now() - tAll) / 1000).toFixed(1)}s in total`);
      setDf((d) => ({ ...d, st: "on", part: "", savedAt: nowStamp(), err: "" }));
      setDfErr((x) => (x && x.kind === "save" ? null : x));
    } catch (e) {
      const fe = fileErr(e);
      console.error("Quote Follow Up · save failed", e);
      setDf((d) => ({ ...d, st: "err", part: "", err: fe.text }));
      setDfErr({ kind: "save", title: `Could not save in ${fileName(cur)}`, ...fe });
    } finally {
      writing.current = false;
    }
  };
  /* Conectar à pasta: lê os dois arquivos e junta com o navegador. Como no pulz, NÃO grava ao abrir: a
     gravação só acontece no Save (botão ou Ctrl+S). write = true para criar os arquivos numa pasta nova. */
  const syncFrom = async (dir, r, write) => {
    const local = stateRef.current;
    const m = r ? mergeData(r.d, local) : null;
    if (m) await applyData(m);
    handleRef.current = dir;
    try {
      await kv.set(HANDLE_KEY, dir);
    } catch {}
    const st = m ? { ...local, ...m, salesRep: m.salesRows.length ? buildSalesIndex(m.salesRows, m.salesInfo) : null } : local;
    // o que está em cada arquivo agora: o que o navegador tiver a mais aparece como "não salvo"
    const dd = r && r.dd,
      dr = r && r.dr;
    savedSig.current = {
      // arquivo "data" antigo (com as cotações dentro) precisa ser regravado sem elas
      data: dd && !r.legacy ? sigOf({ meta: dd.meta, excl: dd.excl, acts: dd.acts, dailyGoal: dd.dailyGoal, dayGoals: dd.dayGoals, users: dd.users }, "data") : "",
      rep: dr ? sigOf({ quotesRep: dr.quotesRep, salesRep: dr.salesRows.length ? buildSalesIndex(dr.salesRows, dr.salesInfo) : null }, "rep") + (m && m.pgFromLocal ? "|pg" : "") : sigOf({ quotesRep: null, salesRep: null }, "rep"),
    };
    savedActs.current = new Set(dd ? (dd.acts || []).map(actKey) : []);
    needSplit.current = !!(r && r.legacy);
    savedSnap.current = dd ? { meta: dd.meta || {}, excl: dd.excl || [], dailyGoal: dd.dailyGoal || 0, dayGoals: dd.dayGoals || {}, users: dd.users || {} } : { meta: {}, excl: [], dailyGoal: 0, dayGoals: {}, users: {} };
    setDf({ st: "on", name: dir.name, openedAt: nowStamp() });
    if (dd && dd.usersBad) setTimeout(() => say("warn", "The users list in the data file was changed outside the program and could not be read — kept the users this computer knows. Save to write it again."), 1800);
    setDfErr((x) => (newerFile.current && x && x.kind === "save" ? x : null));
    syncOk.current = true;
    if (write) {
      stateRef.current = st;
      await saveNow(true);
    }
    return m;
  };
  const pickFolder = () => window.showDirectoryPicker({ id: "qp-data", mode: "readwrite" });
  const openFolder = async (dir, create) => {
    setDf((x) => ({ ...x, st: "opening", name: dir.name, prev: x.st }));
    setDfErr(null);
    await new Promise((r) => setTimeout(r, 30)); // deixa o aviso aparecer antes da leitura
    let r;
    try {
      r = await readFolder(dir);
    } catch (e) {
      const fe = fileErr(e);
      console.error("Quote Follow Up · open failed", e);
      setDf((x) => ({ st: x.prev && x.prev !== "opening" ? x.prev : "none", name: x.prev === "on" ? x.name : "" }));
      setDfErr({ kind: "open", title: `Could not read the data files in ${dir.name}`, ...fe });
      return null;
    }
    if (!r && !create) {
      setDf((x) => ({ st: x.prev && x.prev !== "opening" ? x.prev : "none", name: x.prev === "on" ? x.name : "" }));
      setDfErr({ kind: "open", title: `No Quote Follow Up data files in ${dir.name}`, text: `${DATA_FILE} was not found in this folder.`, hint: "Pick the folder that has the data file, or use “Create the data files in a folder…”." });
      return null;
    }
    const m = await syncFrom(dir, r, create);
    return { m, r };
  };
  const connectNew = async () => {
    try {
      const dir = await pickFolder();
      const res = await openFolder(dir, true);
      if (res) say("ok", `Data files ${res.r ? "updated" : "created"} in ${dir.name}: ${DATA_FILE} (follow-ups, notes, history) and ${REP_FILE} (quotes and sales).`);
    } catch (e) {
      if (e && e.name === "AbortError") return;
      setDf((x) => (x.st === "opening" ? { st: "none" } : x));
      setDfErr({ kind: "open", title: "Could not create the data files", ...fileErr(e) });
    }
  };
  const connectOpen = async () => {
    try {
      const dir = await pickFolder();
      const res = await openFolder(dir, false);
      if (res) {
        const m = res.m;
        say(
          "ok",
          `Connected to the folder ${dir.name}${m ? `: ${((m.quotesRep && m.quotesRep.quotes.length) || 0).toLocaleString("pt-BR")} quotes, ${m.acts.length} activities` : ""}.${res.r && res.r.legacy ? ` On the next Save the quotes move to ${REP_FILE} and ${DATA_FILE} becomes small.` : ""} Changes stay in this browser until you click Save (or Ctrl+S).`,
        );
      }
    } catch (e) {
      if (e && e.name === "AbortError") return;
      setDf((x) => (x.st === "opening" ? { st: "none" } : x));
      setDfErr({ kind: "open", title: "Could not open the data folder", ...fileErr(e) });
    }
  };
  const reconnect = async () => {
    const dir = handleRef.current;
    if (!dir || dir.kind !== "directory") return connectOpen();
    try {
      if ((await dir.requestPermission({ mode: "readwrite" })) !== "granted") throw new Error("permission to save in the folder was not given");
      if (await openFolder(dir, false)) say("ok", `Reconnected to ${dir.name}.`);
    } catch (e) {
      say("err", `Could not reconnect: ${(e && e.message) || e}. Use "Open the data folder" to pick it again.`);
    }
  };
  // sem File System Access (Firefox etc.): baixar / carregar o arquivo à mão (tudo num arquivo só)
  const downloadData = () => {
    const buf = XLSX.write(buildDataWorkbook(stateRef.current), DATA_WRITE);
    downloadBlob(new Blob([buf], { type: XLSX_TYPE[0].accept && Object.keys(XLSX_TYPE[0].accept)[0] }), `Quote Follow Up data ${today}.xlsx`);
  };
  const loadDataFile = async (file) => {
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      if (!isDataWorkbook(wb)) throw new Error("not a Quote Follow Up data file");
      const m = mergeData(parseDataWorkbook(wb), stateRef.current);
      await applyData(m);
      skipSave.current = 0;
      say("ok", `Data file loaded: ${((m.quotesRep && m.quotesRep.quotes.length) || 0).toLocaleString("pt-BR")} quotes.`);
    } catch (e) {
      say("err", `Could not load ${file.name}: ${(e && e.message) || e}`);
    }
  };

  /* ---------- carregar: cache do navegador e, se houver, a pasta dos arquivos de dados ---------- */
  useEffect(() => {
    (async () => {
      const [q, s] = await Promise.all([kv.get(K.quotes), kv.get(K.sales)]);
      if (q) setQuotesRep(q);
      if (s && s.v === 2) setSalesRep(s);
      else if (s) {
        await kv.del(K.sales);
        say("warn", "The Sales Report was saved by an older version: import it again.");
      }
      let h = null;
      try {
        h = FSA ? await kv.get(HANDLE_KEY) : null;
      } catch {}
      if (h && h.kind === "file") {
        // conexão antiga (um arquivo só): agora o programa usa a pasta (dois arquivos)
        setDf({ st: "ask", name: "the data folder", legacyFile: h.name });
      } else if (h && h.queryPermission) {
        handleRef.current = h;
        try {
          if ((await h.queryPermission({ mode: "readwrite" })) === "granted") {
            stateRef.current = { ...stateRef.current, quotesRep: q || null, salesRep: s && s.v === 2 ? s : null };
            await openFolder(h, false);
          } else setDf({ st: "ask", name: h.name });
        } catch (e) {
          setDf({ st: "ask", name: h.name, err: (e && e.message) || String(e) });
        }
      }
      setLoaded(true);
    })();
  }, []);
  /* Como no pulz: sem gravação automática. As mudanças ficam no navegador (não se perdem) e o botão Save
     mostra quantas ainda não estão nos arquivos; Save (ou Ctrl+S) grava. */
  const connected = !!handleRef.current && handleRef.current.kind === "directory" && ["on", "err", "saving"].includes(df.st);
  const dirtyD = connected && loaded && sigOf(stateRef.current, "data") !== savedSig.current.data;
  const dirtyR = connected && loaded && sigOf(stateRef.current, "rep") !== savedSig.current.rep;
  const unsaved = dirtyD || dirtyR;
  /* Quantas mudanças: cotações cujo FUP, % Win, requisição, Done ou nota estão diferentes do arquivo (cada
     cotação conta uma vez; desfazer tira da conta) + metas alteradas + cotações tiradas/recolocadas no plano.
     0 com "unsaved" = só o histórico mudou (ex.: marcou e desmarcou). */
  const unsavedN = useMemo(() => {
    if (!dirtyD) return 0;
    const sv = savedSnap.current;
    const ids = new Set();
    const clean = (m) => {
      if (!m) return "";
      const { name, ...rest } = m;
      return stableJson(rest);
    };
    for (const id of new Set([...Object.keys(meta || {}), ...Object.keys(sv.meta || {})])) if (clean((meta || {})[id]) !== clean((sv.meta || {})[id])) ids.add(id);
    (acts || []).forEach((a) => a.k === "note" && a.id && !savedActs.current.has(actKey(a)) && ids.add(a.id));
    const e0 = new Set(sv.excl || []),
      e1 = new Set(excl || []);
    [...e0, ...e1].forEach((id) => e0.has(id) !== e1.has(id) && ids.add(id));
    let n = ids.size;
    if ((dailyGoal || 0) !== (sv.dailyGoal || 0)) n++;
    for (const d of new Set([...Object.keys(dayGoals || {}), ...Object.keys(sv.dayGoals || {})])) if ((dayGoals || {})[d] !== (sv.dayGoals || {})[d]) n++;
    if (stableJson(users || {}) !== stableJson(sv.users || {})) n++;
    return n;
  }, [dirtyD, meta, acts, excl, dailyGoal, dayGoals, users, df.savedAt, df.openedAt]);
  const repaired = useRef(new Set()); // cotações consertadas pelo histórico (ver o efeito com repairMeta)
  // motivos gravados com nomes antigos passam para os nomes atuais (Budget Only, Price…)
  useEffect(() => {
    const old = Object.entries(meta || {}).filter(([, m]) => m && m.why && normWhy(m.why) !== m.why);
    if (!old.length) return;
    setMeta((all) => {
      const out = { ...all };
      old.forEach(([id]) => out[id] && (out[id] = { ...out[id], why: normWhy(out[id].why) }));
      return out;
    });
  }, [meta]);
  // o que mudou desde o último Save (lista do aviso amarelo): registros novos do histórico, por cotação, e metas
  const unsavedDetail = useMemo(() => {
    if (!dirtyD) return null;
    const sv = savedSnap.current;
    const byQ = new Map();
    const report = [];
    const qOf = (id, c) => byQ.get(id) || (byQ.set(id, { id, c: c || ((meta || {})[id] || {}).name || "", items: [], t: "" }), byQ.get(id));
    (acts || []).forEach((a) => {
      if (!a.id || savedActs.current.has(actKey(a))) return;
      if (["status", "amount", "order"].includes(a.k)) return report.push(a);
      const q = qOf(a.id, a.c);
      q.items.push(a);
      if (a.t > q.t) q.t = a.t;
    });
    const cleanM = (m) => {
      if (!m) return "";
      const { name, ...rest } = m;
      return stableJson(rest);
    };
    for (const id of new Set([...Object.keys(meta || {}), ...Object.keys(sv.meta || {})]))
      if (cleanM((meta || {})[id]) !== cleanM((sv.meta || {})[id]) && !byQ.has(id)) qOf(id).items.push({ k: repaired.current.has(id) ? "_repair" : "_meta" });
    const e0 = new Set(sv.excl || []),
      e1 = new Set(excl || []);
    [...e0, ...e1].forEach((id) => e0.has(id) !== e1.has(id) && !byQ.has(id) && qOf(id).items.push({ k: "plan", n: e1.has(id) ? "out" : "in" }));
    const goals = [];
    if ((dailyGoal || 0) !== (sv.dailyGoal || 0)) goals.push({ d: "", o: sv.dailyGoal || 0, n: dailyGoal || 0 });
    for (const d of [...new Set([...Object.keys(dayGoals || {}), ...Object.keys(sv.dayGoals || {})])].sort())
      if ((dayGoals || {})[d] !== (sv.dayGoals || {})[d]) goals.push({ d, o: (sv.dayGoals || {})[d], n: (dayGoals || {})[d] });
    const quotes = [...byQ.values()].sort((a, b) => (a.t < b.t ? 1 : a.t > b.t ? -1 : 0));
    quotes.forEach((q) => q.items.sort((a, b) => ((a.t || "") < (b.t || "") ? 1 : -1)));
    return { quotes, goals, report, users: stableJson(users || {}) !== stableJson(sv.users || {}) };
  }, [dirtyD, meta, acts, excl, dailyGoal, dayGoals, users, df.savedAt, df.openedAt]);
  const unsavedPop = usePop(440, 520);
  /* Conserto: se o FUP / % Win / data de uma cotação ficou diferente do último registro do histórico (alguém
     salvou por cima com dados velhos), volta para o que o histórico diz. Ver repairMeta em core.js. */
  useEffect(() => {
    if (!loaded) return;
    const r = repairMeta(meta, acts);
    if (!r.fixed.length) return;
    r.fixed.forEach((id) => repaired.current.add(id));
    setMeta(r.meta);
    console.info("Quote Follow Up · restored from the Activities History: " + r.fixed.join(", "));
    const n = r.fixed.length;
    // depois do aviso de abertura da pasta, para não ser coberto por ele
    setTimeout(() => say("info", `${n} quote${n > 1 ? "s had" : " had"} FUP / % Win / follow-up date older than their Activities History (someone saved over them with old data). Restored from the history — click Save to write it.`), 1500);
  }, [loaded, meta, acts]);
  // segundos desde o início da gravação (mostrado no botão "Saving…")
  const [saveSecs, setSaveSecs] = useState(0);
  useEffect(() => {
    if (df.st !== "saving") return;
    setSaveSecs(0);
    const t0 = Date.now();
    const iv = setInterval(() => setSaveSecs(Math.round((Date.now() - t0) / 1000)), 1000);
    return () => clearInterval(iv);
  }, [df.st]);
  const unsavedRef = useRef(false);
  unsavedRef.current = unsaved;
  useEffect(() => {
    const key = (e) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === "s" || e.key === "S")) {
        if (!handleRef.current) return;
        e.preventDefault();
        saveNow();
      }
    };
    const leave = (e) => {
      if (!unsavedRef.current) return;
      e.preventDefault();
      e.returnValue = "";
    };
    window.addEventListener("keydown", key);
    window.addEventListener("beforeunload", leave);
    return () => {
      window.removeEventListener("keydown", key);
      window.removeEventListener("beforeunload", leave);
    };
  }, []);
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), toast.tone === "err" ? 9000 : 5500);
    return () => clearTimeout(t);
  }, [toast]);

  /* ---------- recortes ---------- */
  const quotes = (quotesRep && quotesRep.quotes) || [];
  const qIndex = useMemo(() => new Map(quotes.map((q) => [q.id, q])), [quotes]);
  const clsList = useMemo(() => clsOptions(quotes), [quotes]);
  /* ---------- login e papel ----------
     Quem pode entrar: os nomes das colunas Sales Person e Quoted By do Quote Report, mais as pessoas que o
     Admin acrescentou à mão (gerência). A tela depende do papel que o Admin escolheu (users[nome].role):
     externo = só as cotações em que ele é o Sales Person · interno = só as que ele cotou (Quoted By) ·
     gerência (e o Admin) = todas, com o filtro de vendedor livre.
     Filtro de vendedor da gerência: "sp:Nome" (Sales Person) ou "qb:Nome" (Quoted By). A escolha antiga
     "ext:Nome" continua valendo como "sp:Nome". */
  const [session, setSession] = useState(() => {
    try {
      return sessionStorage.getItem("qp_session") || "";
    } catch {
      return "";
    }
  });
  // classificação padrão: OEM para a gerência; o vendedor (externo e interno) começa vendo todas as dele
  const sellerRole = !!session && session !== ADMIN_USER && ["outside", "inside"].includes(((users || {})[session] || {}).role);
  const cls = prefs.cls !== undefined ? prefs.cls : sellerRole ? "" : (clsList.find((c) => norm(c.v) === "oem") || { v: "" }).v;
  const line = ALL_LINES.includes(prefs.line) ? prefs.line : "";
  const isAdmin = session === ADMIN_USER;
  const liveQuotes = useMemo(() => quotes.filter((q) => !quoteCancelled(q)), [quotes]);
  const allSellers = useMemo(() => sellerOptions(liveQuotes, "sp"), [liveQuotes]);
  const allInside = useMemo(() => sellerOptions(liveQuotes, "qb"), [liveQuotes]);
  const spNameSet = useMemo(() => new Set(allSellers.map((g) => g.name)), [allSellers]);
  const qbNameSet = useMemo(() => new Set(allInside.map((g) => g.name)), [allInside]);
  // pessoas acrescentadas pelo Admin (não aparecem nos relatórios): ficam enquanto tiverem um papel
  const handNames = useMemo(() => Object.keys(users || {}).filter((n) => n !== ADMIN_USER && !spNameSet.has(n) && !qbNameSet.has(n) && ROLES[(users[n] || {}).role]), [users, spNameSet, qbNameSet]);
  const sellerNames = useMemo(() => [...new Set([...spNameSet, ...qbNameSet, ...handNames])].sort((a, b) => a.localeCompare(b, "pt-BR", { sensitivity: "base" })), [spNameSet, qbNameSet, handNames]);
  const colOf = (n) => ({ sp: spNameSet.has(n), qb: qbNameSet.has(n) });
  // papel de quem entrou: o Admin usa a tela da gerência; sem papel válido = sem acesso
  const role = !session ? "" : isAdmin ? "manager" : ROLES[((users || {})[session] || {}).role] ? users[session].role : "";
  const lockSeller = role === "outside" || role === "inside" ? session : "";
  const seller = lockSeller
    ? (role === "inside" ? "qb:" : "sp:") + lockSeller
    : /^(sp|qb):./.test(prefs.seller || "")
    ? prefs.seller
    : /^ext:./.test(prefs.seller || "")
    ? "sp:" + prefs.seller.slice(4)
    : "";
  const inCls = useMemo(() => scopeQuotes(quotes, { cls }), [quotes, cls]);
  /* vendedores da coluna Sales Person, agrupados como no pulz (sem o código SP; a mesma pessoa com grafias
     diferentes = um nome só). O número = cotações não canceladas (todas as datas), como na tabela. */
  const spList = useMemo(() => sellerOptions(inCls.filter((q) => !quoteCancelled(q)), "sp"), [inCls]);
  const qbList = useMemo(() => sellerOptions(inCls.filter((q) => !quoteCancelled(q)), "qb"), [inCls]);
  // nome do vendedor (grafias agrupadas como no filtro) a partir do Sales Person da cotação
  const groupOf = useMemo(() => {
    const m = new Map();
    spList.forEach((g) => g.raw.forEach((r) => m.set(r, g.name)));
    return (raw) => m.get(clean(raw)) || sellerDisplay(raw) || "";
  }, [spList]);
  const who = useMemo(() => {
    if (!seller) return null;
    const name = seller.slice(3);
    if (seller.startsWith("qb:")) {
      // vendedor interno: as cotações em que ele aparece no Quoted By
      const g = qbList.find((x) => x.name === name) || allInside.find((x) => x.name === name);
      return { qb: g ? g.raw : [name], sp: [] };
    }
    const g = spList.find((x) => x.name === name) || allSellers.find((x) => x.name === name);
    return { qb: [], sp: g ? g.raw : [name] };
  }, [seller, spList, allSellers, qbList, allInside]);
  // classificação + linha (Thermometer) e + vendedores (tabela e plano)
  const inLine = useMemo(() => scopeQuotes(quotes, { cls, line }), [quotes, cls, line]);
  // cotações guardadas sem a linha dos itens (importadas por uma versão antiga ou sem a coluna Product Group)
  const noLineList = useMemo(() => inCls.filter((q) => !q.lr && (!q.al || !["PT", "CP", "MH"].some((k) => q.al[k]))), [inCls]);
  const noLine = noLineList.length;
  const scoped = useMemo(() => scopeQuotes(quotes, { cls, line, who }), [quotes, cls, line, who]);
  // Hoje: todos os follow-ups da pessoa, sem classificação nem linha (um follow-up vencido não some por filtro)
  const scopedToday = useMemo(() => scopeQuotes(quotes, { who }), [quotes, who]);
  // Sales Report no recorte da tela: booked por dia só do vendedor e da linha escolhidos
  const salesView = useMemo(() => salesForLine(salesForSeller(salesRep, who), line), [salesRep, who, line]);
  // período como no pulz: ano fiscal inteiro ou um mês dele (padrão: o ano fiscal atual)
  const pGroups = useMemo(() => periodGroups(inCls, salesRep), [inCls, salesRep]);
  const curFyY = (pGroups.find((g) => g.current) || pGroups[0] || { y: 0 }).y;
  const period = periodValid(prefs.period, pGroups) ? prefs.period : "fy:" + curFyY;
  const range = periodRange(period);
  // Thermometer: por cliente, na classificação e linha escolhidas (vendedor não muda o comportamento do cliente)
  const thermoRef = thermoRefDate(quotesRep, today);
  // cotações fechadas à mão (FUP → Cancelled / Lost): chave em texto para só recalcular quando elas mudam
  const closeKey = useMemo(
    () =>
      Object.entries(meta || {})
        .filter(([, m]) => m && m.close)
        .map(([id, m]) => id + ":" + m.close)
        .sort()
        .join("|"),
    [meta],
  );
  const closeOf = useMemo(() => {
    const m = new Map(closeKey ? closeKey.split("|").map((x) => x.split(":")) : []);
    return (id) => m.get(id) || "";
  }, [closeKey]);
  const thermo = useMemo(() => customerThermo(inLine, salesView, cls, thermoRef, today, closeOf), [inLine, salesView, cls, thermoRef, today, closeOf]);
  const rows = useMemo(() => buildQuoteRows(scoped, salesView, { ...range, cls, closeOf }), [scoped, salesView, range.from, range.to, cls, closeOf]);
  const cancelRows = useMemo(() => buildQuoteRows(scoped, salesView, { ...range, cls, cancelled: true, closeOf }), [scoped, salesView, range.from, range.to, cls, closeOf]);

  /* ---------- campos manuais e Activities History ---------- */
  const log = (a) => setActs((x) => [...x, { t: nowStamp(), ...a, ...(user ? { by: user } : null) }]);
  // notas por cotação (contador e última nota, para a tabela e o plano)
  const notesBy = useMemo(() => {
    const m = {};
    acts.forEach((a) => {
      if (a.k !== "note") return;
      const o = m[a.id] || (m[a.id] = { n: 0, last: null });
      o.n++;
      if (!o.last || a.t >= o.last.t) o.last = a;
    });
    return m;
  }, [acts]);
  /* Campos manuais por COTAÇÃO (chave = Quote ID). Ao escolher o FUP pela primeira vez, a agenda
     começa a contar de hoje. Cada mudança vai para o Activities History. */
  const onMeta = (id, patch, name) => {
    const old = meta[id] || {};
    if ("fup" in patch && (patch.fup || "") !== (old.fup || "")) log({ id, c: name, k: "fup", o: old.fup || "", n: patch.fup || "" });
    if ("win" in patch && (patch.win || 0) !== (old.win || 0)) log({ id, c: name, k: "win", o: old.win || "", n: patch.win || "" });
    if ("req" in patch && (patch.req || "") !== (old.req || "")) log({ id, c: name, k: "req", o: old.req || "", n: patch.req || "" });
    setMeta((all) => {
      const cur = { ...(all[id] || {}) };
      if ("fup" in patch) {
        // FUP novo ou trocado: o ciclo começa hoje (o B agenda o 2º e o 4º dia útil a partir de hoje)
        if (patch.fup && patch.fup !== cur.fup) {
          cur.since = today;
          delete cur.step;
          if (patch.fup !== "E") delete cur.at; // a data escolhida só vale no E
        }
        if (!patch.fup) {
          delete cur.since;
          delete cur.step;
          delete cur.at;
        }
      }
      const next = { ...cur, ...patch, name: name || cur.name };
      Object.keys(next).forEach((k) => (next[k] === "" || next[k] === 0 || next[k] == null) && delete next[k]);
      const out = { ...all };
      if (Object.keys(next).filter((k) => k !== "name").length) out[id] = next;
      else delete out[id];
      return out;
    });
  };
  // follow-up feito: o próximo prazo passa a contar desse dia
  /* Done com a data do próximo follow-up escolhida: a cotação passa para E · Ongoing nessa data.
     Sem data (nextDate vazio), segue o ciclo da letra — B: o 1º Done passa para a 2ª data do ciclo. */
  /* Done (no Activities / History): só com nota e data do próximo follow-up. Grava a nota, o Done e
     passa a cotação para E · Ongoing nessa data. */
  const onDone = (id, name, nextDate, note, win) => {
    if (!nextDate || !(note || "").trim() || !win) return;
    const cur = meta[id] || {};
    const c = name || cur.name || (qIndex.get(id) || {}).c || "";
    log({ id, c, k: "note", txt: note.trim() });
    if (win !== (cur.win || 0)) log({ id, c, k: "win", o: cur.win || "", n: win });
    log({ id, c, k: "done", o: cur.fup !== "E" ? cur.fup || "none" : "", n: nextDate });
    setMeta((all) => ({ ...all, [id]: { ...fupDone(all[id] || {}, today, nextDate), win, name: c } }));
  };
  // Change date: só troca a data do próximo follow-up (sem Done); a cotação passa para E · Ongoing
  const onResched = (id, name, nextDate, note) => {
    if (!nextDate) return;
    const cur = meta[id] || {};
    const c = name || cur.name || (qIndex.get(id) || {}).c || "";
    if ((note || "").trim()) log({ id, c, k: "note", txt: note.trim() });
    log({ id, c, k: "resched", o: nextFup(cur) || "", n: nextDate });
    setMeta((all) => {
      const m = { ...(all[id] || {}), fup: "E", at: nextDate, since: today, name: c };
      delete m.step;
      return { ...all, [id]: m };
    });
  };
  // requisição do cliente (campo da cotação; a mudança vai para o Activities History)
  const onReq = (id, name, v) => onMeta(id, { req: v }, name || (qIndex.get(id) || {}).c || "");
  /* Fechar à mão (FUP → Cancelled / Lost, com motivo), reabrir, e motivo de cotação que já veio fechada.
     Tudo vai para o Activities History. */
  const [closeAsk, setCloseAsk] = useState(null);
  const askClose = (id, name, kind, mode) => {
    const m = meta[id] || {};
    setCloseAsk({ id, c: name || m.name || (qIndex.get(id) || {}).c || "", kind, mode: mode || "close", cur: m.why || "", manual: m.close === kind });
  };
  const setMetaOf = (id, c, fn) =>
    setMeta((all) => {
      const x = fn({ ...(all[id] || {}) });
      x.name = c || x.name;
      const out = { ...all };
      if (Object.keys(x).some((k) => k !== "name")) out[id] = x;
      else delete out[id];
      return out;
    });
  const doClose = (id, c, kind, why) => {
    log({ id, c, k: "close", n: kind, txt: why });
    setMetaOf(id, c, (x) => ({ ...x, close: kind, why }));
  };
  const doWhy = (id, c, why) => {
    const old = (meta[id] || {}).why || "";
    if (old === why) return;
    log({ id, c, k: "why", o: old, n: why });
    setMetaOf(id, c, (x) => ({ ...x, why }));
  };
  const doReopen = (id, c) => {
    log({ id, c, k: "reopen" });
    setMetaOf(id, c, (x) => {
      delete x.close;
      delete x.why;
      return x;
    });
  };
  // abre a gaveta da cotação; done = veio do botão Done (vai direto para a nota + data)
  const [navIds, setNavIds] = useState(null); // lista da tabela quando a cotação foi aberta por ela (setas do painel)
  const openQuote = (id, done, list) => {
    setOpenQ(id);
    setOpenDone(!!done);
    setNavIds(Array.isArray(list) ? list : null);
  };
  const onToggleExcl = (id, name, out) => {
    log({ id, c: name, k: "plan", n: out ? "out" : "in" });
    setExcl((e) => (out ? [...new Set([...e, id])] : e.filter((x) => x !== id)));
  };
  // anotação: entra no histórico com data e hora; não há como editar nem excluir
  const onNote = (id, name, txt) => log({ id, c: name, k: "note", txt });

  /* ---------- importar relatórios: as cotações novas se juntam às guardadas ---------- */
  const importFile = async (file, want) => {
    setBusy(want);
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      if (isDataWorkbook(wb)) {
        await loadDataFile(file);
        return;
      }
      const isQ = isQuoteWorkbook(wb);
      const kind = isQ ? "quote" : "sales";
      const at = nowStamp();
      const note = kind !== want ? ` (this file looks like the ${kind === "quote" ? "Quote" : "Sales"} Report, so it went there)` : "";
      if (kind === "quote") {
        const r = parseQuoteWorkbook(wb);
        if (!r || !r.quotes.length) throw new Error("No quote found. The file needs the columns Quote ID, Customer and Create Date.");
        /* relatório mais velho que o guardado (a última cotação dele é anterior à última guardada): só completa
           Product Groups e linhas que faltam; não muda situação nem valores (senão uma ganha voltaria a Open) */
        if (quotesRep && quotesRep.maxD && r.maxD && r.maxD < quotesRep.maxD) {
          const f = fillQuotes(quotes, r.quotes);
          const rep = { ...quotesRep, quotes: f.quotes, fillAt: at, hasLine: true };
          setQuotesRep(rep);
          await kv.set(K.quotes, rep);
          say(
            "ok",
            `Older Quote Report (until ${fmtDate(r.maxD)}; the stored quotes go until ${fmtDate(quotesRep.maxD)}): used only to fill what was missing. ${f.filled.toLocaleString("pt-BR")} quote${f.filled === 1 ? "" : "s"} got the Product Groups / lines. Status, values and orders did not change${f.outside ? ` · ${f.outside.toLocaleString("pt-BR")} quotes of the file are not stored and were left out` : ""}.`,
          );
          return;
        }
        const m = mergeQuotes(quotes, r.quotes, at, r.dropped);
        const rep = { v: 2, quotes: m.quotes, minD: m.minD, maxD: m.maxD, fileName: file.name, at, hasClass: r.hasClass || (quotesRep && quotesRep.hasClass), hasLine: r.hasLine || (quotesRep && quotesRep.hasLine) };
        setQuotesRep(rep);
        if (m.acts.length) setActs((x) => [...x, ...m.acts]);
        await kv.set(K.quotes, rep);
        say(
          r.hasLine ? "ok" : "warn",
          `Quote Report: ${r.quotes.length.toLocaleString("pt-BR")} quotes read · ${m.added.toLocaleString("pt-BR")} new · ${m.updated.toLocaleString("pt-BR")} changed · ${m.quotes.length.toLocaleString("pt-BR")} stored${note}.${r.skipped ? ` ${r.skipped.toLocaleString("pt-BR")} item line${r.skipped === 1 ? "" : "s"} left out (Product Group outside PT / CP / MH: ${money(r.skippedAmt)})${m.removed ? `, ${m.removed} quote${m.removed === 1 ? "" : "s"} removed` : ""}.` : ""}${r.hasLine ? "" : " No Product Group column: the PT / CP / MH filter will not find these quotes."}`,
        );
      } else {
        let r = parseSalesWorkbook(wb);
        if (!r || !r.booked) throw new Error("No booked line found. The file needs a date column, a customer column and an amount.");
        /* relatório mais velho que o guardado: não substitui o período (as vendas guardadas são mais novas);
           só entram os dias anteriores ao que já está guardado */
        if (salesRep && salesRep.maxD && r.maxD < salesRep.maxD) {
          const before = r.rows.filter((x) => x.d < salesRep.minD);
          if (!before.length) {
            say("warn", `Older Sales Report (until ${fmtDate(r.maxD)}; the stored sales go until ${fmtDate(salesRep.maxD)}): nothing was changed — the stored sales are newer.`);
            return;
          }
          const ds = before.map((x) => x.d).sort();
          r = { ...r, rows: before, minD: ds[0], maxD: ds[ds.length - 1], booked: before.length };
        }
        const rowsAll = mergeSalesRows(salesRep && salesRep.rows, r);
        const rep = buildSalesIndex(rowsAll, { fileName: file.name, at });
        setSalesRep(rep);
        await kv.set(K.sales, rep);
        say(
          r.hasSO ? "ok" : "warn",
          `Sales Report: ${r.booked.toLocaleString("pt-BR")} booked lines from ${fmtDate(r.minD)} to ${fmtDate(r.maxD)} (this period was replaced; the rest of the history stays)${note}.${r.hasSO ? "" : " No Sales Id column: Sold will use the Sales Amt of the quotes."}`,
        );
      }
    } catch (e) {
      say("err", `Could not read ${file.name}: ${e.message || e}`);
    } finally {
      setBusy("");
    }
  };

  const backup = () => {
    const blob = new Blob([JSON.stringify({ app: "quote-follow-up", v: 1, at: new Date().toISOString(), meta, dailyGoal, dayGoals, excl, acts }, null, 1)], { type: "application/json" });
    downloadBlob(blob, `quote-follow-up-backup-${today}.json`);
  };
  const restore = async (file) => {
    try {
      const j = JSON.parse(await file.text());
      if (!j || (j.app !== "quote-follow-up" && j.app !== "quote-pipeline")) throw new Error("not a Quote Follow Up backup");
      const m = mergeData({ meta: j.meta, dailyGoal: j.dailyGoal, dayGoals: j.dayGoals, excl: j.excl, acts: j.acts }, stateRef.current);
      setMeta(m.meta);
      setDailyGoal(m.dailyGoal);
      setDayGoals(m.dayGoals);
      setExcl(m.excl);
      setActs(m.acts);
      say("ok", `Backup restored: ${Object.keys(j.meta || {}).length} quotes with FUP / % Win.`);
    } catch (e) {
      say("err", `Could not restore: ${e.message || e}`);
    }
  };

  const login = (name) => {
    setSession(name);
    setUser(name); // vai em cada registro do Activities History
    try {
      sessionStorage.setItem("qp_session", name);
    } catch {}
  };
  const logout = () => {
    setSession("");
    setOpenQ("");
    try {
      sessionStorage.removeItem("qp_session");
    } catch {}
  };
  // senha criada ou zerada: grava logo no arquivo de dados (senão outro computador não acharia a senha)
  const pendingSave = useRef(false);
  useEffect(() => {
    if (!pendingSave.current) return;
    pendingSave.current = false;
    if (handleRef.current) saveNow(true);
  }, [users]);
  // o papel escolhido pelo Admin nunca é apagado pela senha (criar ou zerar)
  const keepRole = (name, u) => (name === ADMIN_USER ? "admin" : ((u || {})[name] || {}).role || "");
  const createPassword = (name, rec) => {
    pendingSave.current = true;
    setUsers((u) => ({ ...u, [name]: { ...((u || {})[name] || {}), ...rec, role: keepRole(name, u) || rec.role || "" } }));
    login(name);
  };
  const resetPassword = (name) => {
    pendingSave.current = true;
    setUsers((u) => ({ ...u, [name]: { ...((u || {})[name] || {}), role: keepRole(name, u), salt: "", hash: "", at: nowStamp() } }));
    say("ok", `Password of ${name} reset: a new one is created at the next sign-in.`);
  };
  // Admin escolhe a tela de alguém (grava logo, como as senhas; "at" novo = vale a escolha mais recente)
  const setRole = (name, r) => {
    pendingSave.current = true;
    setUsers((u) => ({ ...u, [name]: { salt: "", hash: "", ...((u || {})[name] || {}), role: ROLES[r] ? r : "", at: nowStamp() } }));
    say("ok", ROLES[r] ? `${name} now sees the ${ROLES[r].label} screen.` : `${name} has no access now.`);
  };
  const addUser = (name) => setRole(name, "manager");
  // precisa entrar: ninguém logado, ou quem entrou não existe mais (nem nos relatórios nem acrescentado pelo Admin)
  const needLogin = !session || (!isAdmin && sellerNames.length > 0 && !sellerNames.includes(session));
  const noAccess = !needLogin && !role;

  /* abas: cada papel tem as suas (ROLE_TABS); a aba escolhida fica guardada por papel */
  const tabs = ROLE_TABS[role] || [];
  const tabKey = "tab_" + (role || "none");
  const tab = tabs.some((t) => t[0] === prefs[tabKey]) ? prefs[tabKey] : (tabs[0] || [""])[0];
  const setTab = (k) => setPref(tabKey, k);
  const mob = useMobile();
  const [showFilters, setShowFilters] = useState(false); // celular: filtros recolhidos atrás de um botão
  // celular: a página precisa da meta viewport para não abrir "encolhida"
  useEffect(() => {
    if (document.querySelector('meta[name="viewport"]')) return;
    const m = document.createElement("meta");
    m.name = "viewport";
    m.content = "width=device-width, initial-scale=1, viewport-fit=cover";
    document.head.appendChild(m);
  }, []);
  const repInfo = (r, kind) =>
    r ? `${kind === "q" ? r.quotes.length.toLocaleString("pt-BR") + " quotes stored" : r.rows.length.toLocaleString("pt-BR") + " booked rows stored"} · ${fmtDate(r.minD)}–${fmtDate(r.maxD)} · last file ${r.fileName || "—"}${r.at ? " imported " + fmtStamp(r.at) : ""}` : "";
  const selS = { ...inp, background: "#FEF6EA", borderColor: "#E9C48A", color: "#8A5A12", fontWeight: 600, maxWidth: 230 };

  /* ---------- botão Save (ao lado das abas) ---------- */
  const saveBtn = (() => {
    const base = { ...btn, height: mob ? 38 : 40, padding: mob ? "0 11px" : "0 14px", fontSize: 13, fontWeight: 700, borderRadius: 11 };
    if (!FSA)
      return (
        <button onClick={downloadData} title="This browser cannot keep a file connected: download the data file" style={base}>
          <Save size={16} /> Save
        </button>
      );
    if (!connected && df.st !== "saving")
      return (
        <button onClick={df.st === "ask" ? reconnect : connectNew} title={df.st === "ask" ? "Reconnect the data folder first" : "No data folder connected: choose a folder to create the data files (or open one from the ⋯ menu)"} style={base}>
          <Save size={16} /> Save
        </button>
      );
    if (df.st === "saving")
      return (
        <button disabled title={df.part === "rep" ? `Writing ${REP_FILE} (quotes and sales). It is big: on the network drive this can take a while.` : `Writing ${DATA_FILE}.`} style={{ ...base, cursor: "wait", color: X.sub, background: "#EEF2F5" }}>
          <Save size={16} /> {df.part === "rep" ? "Saving quotes…" : "Saving…"} {saveSecs}s
        </button>
      );
    const hot = unsaved || df.st === "err";
    return (
      <button
        onClick={() => saveNow(!hot)}
        title={hot ? `Write the changes (Ctrl+S)${dirtyR ? ` · the quotes and sales changed too (${REP_FILE}, bigger: takes longer)` : ""}` : `Everything is saved in the folder ${df.name} · click (or Ctrl+S) to save again`}
        style={hot ? { ...base, background: X.teal, borderColor: X.teal, color: "#fff", boxShadow: "0 2px 6px rgba(31,154,132,.3)" } : { ...base, color: X.tealDk, borderColor: "#BFE3D9" }}
      >
        <Save size={16} /> {hot ? "Save" : "Saved"}
        {unsavedN > 0 && <span style={{ fontSize: 11, fontWeight: 800, background: "#fff", color: X.tealDk, padding: "1px 7px", borderRadius: 999 }}>{unsavedN.toLocaleString("pt-BR")}</span>}
      </button>
    );
  })();
  /* ---------- chip do arquivo de dados ---------- */
  const dfChip = (() => {
    const base = { display: "inline-flex", alignItems: "center", gap: 7, height: 36, padding: "0 12px", borderRadius: 9, fontSize: 12.5, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap", border: "1px solid" };
    if (df.st === "opening")
      return (
        <span title="Reading the data file and joining it with this browser. A big file (or a network drive) can take a few seconds." style={{ ...base, cursor: "wait", background: "#EEF2F5", borderColor: X.line, color: X.sub }}>
          <Database size={15} /> Opening {df.name}…
        </span>
      );
    if (df.st === "saving")
      return (
        <span title="Writing the data files. On a network drive this can take a while." style={{ ...base, cursor: "wait", background: "#EEF2F5", borderColor: X.line, color: X.sub }}>
          <Database size={15} /> {df.name} · saving {df.part === "rep" ? "quotes and sales" : "follow-ups"}…
        </span>
      );
    if (df.st === "on" && unsaved)
      return (
        <span ref={unsavedPop.anchor} onClick={unsavedPop.toggle} title={`Click to see what changed. The changes are kept in this browser; Save (or Ctrl+S) writes them in the folder ${df.name}${dirtyR ? ` (including ${REP_FILE}: quotes and sales)` : ""}`} style={{ ...base, background: "#FEF0C7", borderColor: "#E9C48A", color: "#8A5A12" }}>
          <Database size={15} /> {df.name}{" "}
          <span style={{ fontWeight: 500 }}>
            · {unsavedN ? `${unsavedN.toLocaleString("pt-BR")} unsaved change${unsavedN === 1 ? "" : "s"}` : needSplit.current ? "click Save to split into 2 files" : dirtyD ? "history only, not saved" : dirtyR ? "imported reports not saved" : "not saved"}
            {dirtyR && (unsavedN || dirtyD) && !needSplit.current ? " + reports" : ""}
          </span>
        </span>
      );
    if (df.st === "on")
      return (
        <span title={`Data folder ${df.name}: ${DATA_FILE} and ${REP_FILE} · everything is saved`} style={{ ...base, cursor: "default", background: X.tealSoft, borderColor: "#BFE3D9", color: X.tealDk }}>
          <Database size={15} /> {df.name} <span style={{ fontWeight: 500, color: X.sub }}>{df.savedAt ? `· saved ${df.savedAt.slice(11, 16)}` : df.openedAt ? `· opened ${df.openedAt.slice(11, 16)}` : ""}</span>
        </span>
      );
    if (df.st === "err")
      return (
        <span onClick={() => saveNow()} title={`Could not save in ${df.name}: ${df.err}. The changes are kept in this browser. Click to try again.`} style={{ ...base, background: X.redSoft, borderColor: "#F3B4B4", color: X.redDk }}>
          <AlertTriangle size={15} /> Not saved · {df.name} · click to try again
        </span>
      );
    if (df.st === "ask")
      return (
        <span
          onClick={reconnect}
          title={df.legacyFile ? `The data is now kept in two files in a folder. Click and pick the folder where ${df.legacyFile} is.` : "The browser asks for permission again after the page is reopened. Click to reconnect the data folder."}
          style={{ ...base, background: "#FEF0C7", borderColor: "#E9C48A", color: "#8A5A12" }}
        >
          <Link2 size={15} /> {df.legacyFile ? "Choose the data folder" : `Reconnect ${df.name || "the data folder"}`}
        </span>
      );
    if (df.st === "nofsa")
      return (
        <span onClick={downloadData} title="This browser cannot keep a file connected (use Chrome or Edge). Click to download the data file; load it again from the ⋯ menu." style={{ ...base, background: "#fff", borderColor: X.line, color: X.sub }}>
          <Download size={15} /> Download data file
        </span>
      );
    return (
      <span onClick={connectNew} title="The data is only in this browser: clearing the cache erases it. Click to create an Excel data file (or open one from the ⋯ menu)." style={{ ...base, background: "#FEF0C7", borderColor: "#E9C48A", color: "#8A5A12" }}>
        <AlertTriangle size={15} /> No data file · create one
      </span>
    );
  })();

  /* ---------- aviso (toast): no celular fica em cima da barra de abas ---------- */
  const toastEl = toast && (
    <div
      style={{
        position: "fixed",
        ...(mob ? { left: 10, right: 10, bottom: needLogin || noAccess ? 12 : 76 } : { right: 20, bottom: 20, maxWidth: 480 }),
        zIndex: 2200,
        ...card,
        padding: "12px 14px",
        display: "flex",
        gap: 10,
        alignItems: "flex-start",
        borderLeft: `4px solid ${toast.tone === "err" ? X.redDk : toast.tone === "warn" ? "#F2B04C" : X.teal}`,
        boxShadow: "0 10px 30px rgba(16,24,40,.18)",
        fontSize: 13,
      }}
    >
      {React.createElement(toast.tone === "err" ? AlertTriangle : toast.tone === "warn" || toast.tone === "info" ? Info : CheckCircle2, { size: 17, color: toast.tone === "err" ? X.redDk : toast.tone === "warn" ? X.amber : toast.tone === "info" ? X.sub : X.tealDk, style: { flexShrink: 0, marginTop: 1 } })}
      <span style={{ flex: 1 }}>{toast.text}</span>
      <button onClick={() => setToast(null)} style={{ border: "none", background: "transparent", cursor: "pointer", color: X.sub, padding: 0 }}>
        <XIcon size={15} />
      </button>
    </div>
  );
  const css = <style>{GLOBAL_CSS}</style>;

  if (needLogin)
    return (
      <>
        {css}
        <LoginScreen
          names={sellerNames}
          users={users || {}}
          connected={connected}
          dfName={df.name}
          dfSt={df.st}
          onOpen={connectOpen}
          onReconnect={reconnect}
          onLogin={login}
          onCreate={createPassword}
          onLoadFile={loadDataFile}
          busy={!loaded}
        />
        {toastEl}
      </>
    );
  if (noAccess)
    return (
      <>
        {css}
        <NoAccessScreen name={session} connected={connected} checking={checkSt.busy} onCheck={() => pullIfChanged(true)} onLogout={logout} />
        {toastEl}
      </>
    );

  const R = ROLES[role];
  const filtersTxt = `${seller ? seller.slice(3) : "All salespeople"} · ${cls || "All classifications"} · ${line ? "line " + line : "all lines"}`;
  const showPeriod = ["portfolio", "mine", "analysis", "sellers"].includes(tab);
  const canFupReport = role === "inside" || role === "manager";
  const nFilters = [cls, line, !lockSeller && seller].filter(Boolean).length;

  /* ---------- filtros (período, classificação, linha, vendedor) ---------- */
  const fLbl = { fontSize: 12, color: X.sub, fontWeight: 600 };
  const periodOpts = (
    <>
    <option value="all">All periods</option>
    {pGroups.map((g) => (
      <optgroup key={g.y} label={`${g.label} · ${g.range}`}>
        <option value={"fy:" + g.y}>
          {g.label} · {g.range} (whole year){g.current ? " · current" : ""}
        </option>
        {g.months.map((m) => (
          <option key={m} value={"m:" + m}>
            {"   "}
            {MON_SHORT[+m.slice(5) - 1]}/{m.slice(0, 4)}
          </option>
        ))}
      </optgroup>
    ))}
    </>
  );
  const periodSel = showPeriod && (
    <div style={{ position: "relative", display: "flex", alignItems: "center", minWidth: 0, ...(mob ? { flex: 1 } : null) }} title="Period: all, a whole fiscal year or one month">
      <CalendarDays size={15} color="#8A5A12" style={{ position: "absolute", left: 11, pointerEvents: "none" }} />
      <select value={period} onChange={(e) => setPref("period", e.target.value)} style={{ ...selS, maxWidth: "none", minWidth: mob ? 0 : 250, width: mob ? "100%" : undefined, paddingLeft: 34, height: mob ? 40 : 36 }}>
        {periodOpts}
      </select>
    </div>
  );
  const clsSel = clsList.length > 0 && (
    <label style={{ display: "flex", flexDirection: mob ? "column" : "row", alignItems: mob ? "stretch" : "center", gap: mob ? 5 : 7, ...fLbl }}>
      Classification
      <select value={cls} onChange={(e) => setPref("cls", e.target.value)} style={{ ...selS, ...(mob ? { maxWidth: "none", width: "100%", height: 40 } : null) }}>
        <option value="">All</option>
        {clsList.map((o) => (
          <option key={o.v} value={o.v}>
            {o.v} ({o.n.toLocaleString("pt-BR")})
          </option>
        ))}
      </select>
    </label>
  );
  const lineSeg = (
    <div title="Product line (Product Group): quoted, open, lost, sold and booked values only count the line" style={{ display: "flex", flexDirection: mob ? "column" : "row", alignItems: mob ? "stretch" : "center", gap: mob ? 5 : 6, ...fLbl }}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
        <Layers size={14} /> Line{mob ? "" : ":"}
      </span>
      <div style={{ display: "flex", gap: 2, padding: 2, borderRadius: 9, border: `1px solid ${line ? "#E9C48A" : X.line}`, background: line ? "#FFFBEB" : "#fff" }}>
        {["", ...ALL_LINES].map((v) => {
          const on = line === v;
          return (
            <button
              key={v || "all"}
              onClick={() => setPref("line", v)}
              style={{ flex: mob ? 1 : "none", border: "none", borderRadius: 7, padding: mob ? "9px 11px" : "5px 11px", fontSize: 12, fontWeight: on ? 700 : 500, cursor: "pointer", background: on ? (v ? "#FDE9CF" : X.tealSoft) : "transparent", color: on ? (v ? "#8A5A12" : X.tealDk) : X.sub }}
            >
              {v || "All"}
            </button>
          );
        })}
      </div>
    </div>
  );
  const noLineTag = line && noLine > 0 && (
    <span
      title={`These quotes were stored by an older version of the program (or the Quote Report had no Product Group column), so the program does not know their line. Import the Quote Report that has them again: the lines are filled in and nothing else changes.\n\n${noLineList
        .slice(0, 12)
        .map((q) => `${q.id} · ${q.c} · ${fmtDate(q.d)} · ${q.cls || "no classification"}${q.up ? ` · last import ${fmtStamp(q.up)}` : ""}`)
        .join("\n")}${noLine > 12 ? `\n… and ${noLine - 12} more` : ""}`}
      style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: X.redDk, background: X.redSoft, padding: "5px 10px", borderRadius: 8 }}
    >
      <AlertTriangle size={14} /> {noLine.toLocaleString("pt-BR")} quote{noLine === 1 ? "" : "s"} without product line · import the Quote Report again
    </span>
  );
  // gerência: vendedor livre (externo = Sales Person, interno = Quoted By) · externo e interno: preso no próprio nome
  const sellerOpts = (
    <>
    <option value="">All salespeople</option>
    {spList.length > 0 && (
      <optgroup label="Outside sales · Sales Person">
        {spList.map((o) => (
          <option key={"sp" + o.v} value={`sp:${o.v}`} title={"Sales Person in the Quote Report as: " + o.raw.join(" · ")}>
            {o.v} ({o.n.toLocaleString("pt-BR")})
          </option>
        ))}
      </optgroup>
    )}
    {qbList.length > 0 && (
      <optgroup label="Inside sales · Quoted By">
        {qbList.map((o) => (
          <option key={"qb" + o.v} value={`qb:${o.v}`} title={"Quoted By in the Quote Report as: " + o.raw.join(" · ")}>
            {o.v} ({o.n.toLocaleString("pt-BR")})
          </option>
        ))}
      </optgroup>
    )}
    </>
  );
  const sellerSel = lockSeller ? (
    <span title={`Signed in as ${lockSeller}: you see only the quotes where you are the ${R.colTxt}`} style={{ display: "inline-flex", alignItems: "center", gap: 7, fontSize: 12.5, fontWeight: 600, color: R.fg, background: R.bg, padding: "7px 11px", borderRadius: 9 }}>
      {React.createElement(R.icon, { size: 14 })} Only my quotes · {R.colTxt}
    </span>
  ) : (
    <label style={{ display: "flex", flexDirection: mob ? "column" : "row", alignItems: mob ? "stretch" : "center", gap: mob ? 5 : 8, fontSize: 12.5, fontWeight: 600, color: X.ink }}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 7, ...(mob ? fLbl : null) }}>
        <UserRound size={15} color={X.sub} /> Salesperson
      </span>
      <select value={seller} onChange={(e) => setPref("seller", e.target.value)} title="Outside = Sales Person column · Inside = Quoted By column" style={{ ...selS, maxWidth: "none", minWidth: mob ? 0 : 230, width: mob ? "100%" : undefined, height: mob ? 40 : 36, background: seller ? "#FEF6EA" : "#FDE9CF" }}>
        {sellerOpts}
      </select>
    </label>
  );
  const clearBtn = (line || (!lockSeller && seller)) && (
    <button onClick={() => setPrefs((p) => ({ ...p, line: "", ...(lockSeller ? null : { seller: "" }) }))} style={{ ...btn, height: mob ? 38 : 32, fontSize: 12, padding: "0 10px", justifyContent: "center" }} title="Clear Line and Salesperson">
      <XIcon size={13} /> Clear
    </button>
  );
  const fupRep = canFupReport && <FupReportButton acts={acts} quotes={scoped} meta={meta} thermo={thermo} today={today} filtersTxt={filtersTxt} />;
  /* ---------- celular: os filtros ficam numa folha ----------
     Na aba Quotes a folha abre pelo botão ao lado da busca (junto com ordem e filtros por coluna);
     nas outras abas, pelo botão de filtro no topo. O Today não tem filtro: mostra todos os follow-ups da pessoa. */
  const sheetFilters = (
    <>
      {showPeriod && (
        <SheetField label="Period">
          <select value={period} onChange={(e) => setPref("period", e.target.value)} style={{ ...inp, width: "100%", height: 46, fontSize: 14, borderRadius: 12 }}>
            {periodOpts}
          </select>
        </SheetField>
      )}
      {clsList.length > 0 && (
        <SheetField label="Classification">
          <Chips value={cls} options={[["", "All"], ...clsList.map((o) => [o.v, o.v])]} onChange={(v) => setPref("cls", v)} />
        </SheetField>
      )}
      <SheetField label="Product line">
        <Chips value={line} options={[["", "All"], ...ALL_LINES.map((v) => [v, v])]} onChange={(v) => setPref("line", v)} />
      </SheetField>
      {!lockSeller && (
        <SheetField label="Salesperson">
          <select value={seller} onChange={(e) => setPref("seller", e.target.value)} style={{ ...inp, width: "100%", height: 46, fontSize: 14, borderRadius: 12 }}>
            {sellerOpts}
          </select>
        </SheetField>
      )}
      {noLineTag}
      {fupRep && (
        <SheetField label="Report">
          <span>{fupRep}</span>
        </SheetField>
      )}
    </>
  );
  // quantos filtros estão ligados (o período não conta: sempre tem um)
  const nSheet = [cls, line, !lockSeller && seller].filter(Boolean).length;
  const filterTabs = ["portfolio", "calendar", "goal", "analysis", "sellers"];
  const hasTopFilter = mob && quotesRep && filterTabs.includes(tab) && tab !== "portfolio"; // Quotes: o filtro fica ao lado da busca
  const filterSheet = mob && (
    <Sheet
      open={showFilters}
      title="Filters"
      onClose={() => setShowFilters(false)}
      footer={
        <>
          {nSheet > 0 && (
            <button onClick={() => setPrefs((p) => ({ ...p, cls: "", line: "", ...(lockSeller ? null : { seller: "" }) }))} style={{ ...btn, height: 44, flex: 1, justifyContent: "center", borderRadius: 12 }}>
              Clear
            </button>
          )}
          <button onClick={() => setShowFilters(false)} style={{ ...btn, height: 44, flex: 2, justifyContent: "center", borderRadius: 12, background: X.teal, borderColor: X.teal, color: "#fff" }}>
            Done
          </button>
        </>
      }
    >
      {sheetFilters}
    </Sheet>
  );
  const filterBar =
    mob || !quotesRep || tab === "data" || tab === "today" ? null : (
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        {periodSel}
        {clsSel}
        {lineSeg}
        {noLineTag}
        {sellerSel}
        {clearBtn}
        <span style={{ fontSize: 11.5, color: X.sub, display: "inline-flex", alignItems: "center", gap: 6 }}>
          <Database size={13} />
          {salesRep ? `Sales ${fmtDate(salesRep.minD)}–${fmtDate(salesRep.maxD)}` : "No Sales Report: Sold uses the quote's Sales Amt"}
          {seller && salesRep && (salesRep.hasSeller ? ` · booked: only ${seller.slice(3)}` : " · booked does not follow the salesperson yet: import the Sales Report again")}
        </span>
        {fupRep && <span style={{ marginLeft: "auto" }}>{fupRep}</span>}
      </div>
    );

  /* ---------- abas: no computador no topo; no celular, fixas embaixo ---------- */
  const tabBadge = (k) => (k === "portfolio" || k === "mine" ? rows.length : null);
  const tabsTop = !mob && tabs.length > 1 && (
    <div style={{ display: "inline-flex", gap: 4, background: "#fff", border: `1px solid ${X.line}`, borderRadius: 12, padding: 4, boxShadow: "0 1px 3px rgba(16,24,40,.05)" }}>
      {tabs.map(([k, l, Ico]) => {
        const on = tab === k;
        const n = tabBadge(k);
        return (
          <button
            key={k}
            onClick={() => setTab(k)}
            style={{ display: "inline-flex", alignItems: "center", gap: 8, border: "none", borderRadius: 9, padding: "7px 14px", cursor: "pointer", fontSize: 13, fontWeight: on ? 700 : 500, background: on ? X.tealSoft : "transparent", color: on ? X.tealDk : X.ink }}
          >
            <Ico size={16} /> {l}
            {n != null && <span style={{ fontSize: 11, fontWeight: 700, background: on ? "#fff" : X.head, color: on ? X.tealDk : X.sub, padding: "1px 8px", borderRadius: 999 }}>{n.toLocaleString("pt-BR")}</span>}
          </button>
        );
      })}
    </div>
  );
  const tabsBottom = mob && tabs.length > 1 && (
    <nav
      className="no-print"
      style={{ position: "fixed", left: 0, right: 0, bottom: 0, zIndex: 1200, background: "#fff", borderTop: `1px solid ${X.line}`, boxShadow: "0 -4px 16px rgba(16,24,40,.06)", display: "flex", padding: "6px 6px calc(6px + env(safe-area-inset-bottom))" }}
    >
      {tabs.map(([k, l, Ico]) => {
        const on = tab === k;
        return (
          <button
            key={k}
            onClick={() => {
              setTab(k);
              window.scrollTo(0, 0);
            }}
            className="qp-card-btn"
            style={{ flex: 1, minWidth: 0, border: "none", background: "transparent", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "5px 2px", cursor: "pointer", color: on ? X.tealDk : X.sub, fontSize: 11, fontWeight: on ? 800 : 600, borderRadius: 10 }}
          >
            <span style={{ width: 44, height: 26, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: on ? X.tealSoft : "transparent" }}>
              <Ico size={19} />
            </span>
            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "100%" }}>{l}</span>
          </button>
        );
      })}
    </nav>
  );

  /* ---------- botões de dados (topo e aba Data) ---------- */
  const refreshBtn = connected && df.st !== "saving" && (
    <button
      onClick={() => pullIfChanged(true)}
      disabled={checkSt.busy}
      title={`Bring the changes other people saved in the data files (this also happens by itself every minute and when you come back to this window)${checkSt.at ? ` · last check ${checkSt.at.slice(11, 16)}` : ""}`}
      style={{ ...btn, height: mob ? 38 : 40, padding: mob ? "0 10px" : "0 12px", fontSize: 13, fontWeight: 700, borderRadius: 11, color: X.tealDk, cursor: checkSt.busy ? "wait" : "pointer" }}
    >
      <RefreshCw size={15} /> {mob ? "" : checkSt.busy ? "Checking…" : "Refresh"}
    </button>
  );
  const imports = (
    <>
      <ImportButton label={busy === "quote" ? "Reading…" : "Quote Report"} icon={FileSpreadsheet} info={repInfo(quotesRep, "q")} onFile={(f) => importFile(f, "quote")} accent={!quotesRep} />
      <ImportButton label={busy === "sales" ? "Reading…" : "Sales Report"} icon={Upload} info={repInfo(salesRep, "s")} onFile={(f) => importFile(f, "sales")} accent={!!quotesRep && !salesRep} />
    </>
  );
  /* celular: um ícone diz se está tudo salvo (✓), se falta salvar (tocar salva), ou se a pasta pede atenção */
  const warnS = { background: "#FEF0C7", borderColor: "#E9C48A", color: "#8A5A12" };
  const syncBtn =
    df.st === "saving" || df.st === "opening" ? (
      <button disabled title={df.st === "saving" ? "Saving…" : "Opening the data folder…"} style={iconBtn({ color: X.sub, cursor: "wait" })}>
        <RefreshCw size={17} className="qp-spin" />
      </button>
    ) : !FSA ? (
      <button onClick={downloadData} title="This browser can't keep the data folder: download the data file" style={iconBtn(unsaved ? warnS : null)}>
        <Download size={17} />
      </button>
    ) : df.st === "ask" ? (
      <button onClick={reconnect} title="Reconnect the data folder" style={iconBtn(warnS)}>
        <Link2 size={17} />
      </button>
    ) : !connected ? (
      <button onClick={connectNew} title="No data folder: the changes are only in this browser" style={iconBtn(warnS)}>
        <AlertTriangle size={17} />
      </button>
    ) : df.st === "err" ? (
      <button onClick={() => saveNow()} title="Not saved · tap to try again" style={iconBtn({ background: X.redSoft, borderColor: "#F3B4B4", color: X.redDk })}>
        <AlertTriangle size={17} />
      </button>
    ) : unsaved ? (
      <button onClick={() => saveNow()} title="Save now" style={iconBtn({ background: X.teal, borderColor: X.teal, color: "#fff" })}>
        <Save size={17} />
        {unsavedN > 0 && (
          <span style={{ position: "absolute", top: -4, right: -4, minWidth: 18, height: 18, padding: "0 5px", borderRadius: 999, background: "#F2B04C", color: "#fff", fontSize: 10.5, fontWeight: 800, display: "inline-flex", alignItems: "center", justifyContent: "center", border: "2px solid #fff" }}>{unsavedN}</span>
        )}
      </button>
    ) : (
      <button onClick={() => pullIfChanged(true)} title="Everything is saved · tap to bring the changes of the other people" style={iconBtn({ color: X.tealDk, borderColor: "#BFE3D9" })}>
        {checkSt.busy ? <RefreshCw size={17} className="qp-spin" /> : <CheckCircle2 size={18} />}
      </button>
    );
  const dfStatusTxt =
    df.st === "on"
      ? unsaved
        ? `${unsavedN ? `${unsavedN} change${unsavedN === 1 ? "" : "s"}` : "Changes"} not saved yet in ${df.name}`
        : `All saved in ${df.name}${df.savedAt ? ` · ${df.savedAt.slice(11, 16)}` : ""}`
      : df.st === "err"
      ? `Not saved in ${df.name}: ${df.err}`
      : df.st === "ask"
      ? `Reconnect ${df.name || "the data folder"} to save`
      : df.st === "nofsa"
      ? "This browser can't keep the data folder: download the data file to keep the changes"
      : df.st === "saving"
      ? "Saving…"
      : "No data folder: the changes are only in this browser";
  const menuItems = [
    connected && { label: checkSt.busy ? "Checking…" : "Bring the latest changes", icon: RefreshCw, onClick: () => pullIfChanged(true) },
    connected && { label: "Save now", icon: Save, onClick: () => saveNow(!unsaved) },
    ...(FSA
      ? [
          { label: "Open the data folder…", icon: FolderOpen, onClick: connectOpen },
          { label: "Create the data files in a folder…", icon: FilePlus, onClick: connectNew },
        ]
      : [
          { label: "Load a data file…", icon: FolderOpen, onClick: () => loadDataRef.current && loadDataRef.current.click() },
          { label: "Download the data file", icon: Download, onClick: downloadData },
        ]),
  ];
  const dataMenu = (
    <DataMenu
      items={
        FSA
          ? [
              { label: "Open the data folder…", icon: FolderOpen, onClick: connectOpen },
              { label: "Create the data files in a folder…", icon: FilePlus, onClick: connectNew },
            ]
          : [
              { label: "Load a data file…", icon: FolderOpen, onClick: () => loadDataRef.current && loadDataRef.current.click() },
              { label: "Download the data file", icon: Download, onClick: downloadData },
            ]
      }
    />
  );

  /* ---------- aba Data (gerência): arquivos de dados, relatórios e usuários ---------- */
  const secCard = { ...card, padding: mob ? 12 : 16, display: "flex", flexDirection: "column", gap: 12 };
  const secHead = (Ic, t, sub) => (
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      <span style={{ width: 36, height: 36, borderRadius: 10, background: X.tealSoft, color: X.tealDk, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Ic size={18} />
      </span>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 15.5, fontWeight: 800 }}>{t}</div>
        <div style={{ fontSize: 12, color: X.sub }}>{sub}</div>
      </div>
    </div>
  );
  const dataTab = (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={secCard}>
        {secHead(Database, "Data files", `${DATA_FILE} (follow-ups, notes, history, goals, users) and ${REP_FILE} (quotes and sales)`)}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          {dfChip}
          {refreshBtn}
          {saveBtn}
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {FSA ? (
            <>
              <button onClick={connectOpen} style={btn}>
                <FolderOpen size={15} /> Open the data folder…
              </button>
              <button onClick={connectNew} style={btn}>
                <FilePlus size={15} /> Create the data files in a folder…
              </button>
            </>
          ) : (
            <>
              <button onClick={() => loadDataRef.current && loadDataRef.current.click()} style={btn}>
                <FolderOpen size={15} /> Load a data file…
              </button>
              <button onClick={downloadData} style={btn}>
                <Download size={15} /> Download the data file
              </button>
            </>
          )}
        </div>
      </div>
      <div style={secCard}>
        {secHead(FileSpreadsheet, "Reports", "Import the Quote Report and the Sales Report: the quotes and sales are joined with the ones stored")}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>{imports}</div>
        <div style={{ fontSize: 12, color: X.sub, display: "flex", flexDirection: "column", gap: 4 }}>
          <span>
            <b style={{ color: X.ink }}>Quote Report</b> · {quotesRep ? repInfo(quotesRep, "q") : "not imported yet"}
          </span>
          <span>
            <b style={{ color: X.ink }}>Sales Report</b> · {salesRep ? repInfo(salesRep, "s") : "not imported yet"}
          </span>
        </div>
      </div>
      {isAdmin ? (
        <UsersPanel names={sellerNames} colOf={colOf} users={users || {}} onRole={setRole} onAdd={addUser} onReset={resetPassword} />
      ) : (
        <div style={{ ...secCard, flexDirection: "row", alignItems: "center", fontSize: 12.5, color: X.sub }}>
          <Lock size={15} /> Only the Admin chooses the screen of each person and resets passwords.
        </div>
      )}
    </div>
  );

  /* ---------- conteúdo de cada aba ---------- */
  const tableProps = { onAskClose: askClose, thermo, rows, cancelRows, meta, onMeta, today, periodLabel: periodLabelOf(period), qIndex, sales: salesView, acts, onNote, onDone, onResched, onReq, user, onUser: setUser, notesBy, openQ, onOpenQuote: openQuote };
  const calProps = { goalEdit: role === "manager", thermo, quotes: scoped, sales: salesView, cls, line, meta, onMeta, onAskClose: askClose, onDone, onToggleExcl, onOpenQuote: openQuote, notesBy, dailyGoal, setDailyGoal, dayGoals, setDayGoals, excl, prefs, setPref, today };
  const content = !loaded ? (
    <div style={{ padding: 60, textAlign: "center", color: X.sub }}>Loading…</div>
  ) : tab === "data" ? (
    dataTab
  ) : !quotesRep ? (
    role === "manager" ? (
      <EmptyState onFile={importFile} />
    ) : (
      <div style={{ ...card, padding: "32px 18px", textAlign: "center", display: "flex", flexDirection: "column", gap: 8, alignItems: "center" }}>
        <FileSpreadsheet size={26} color={X.sub} />
        <div style={{ fontSize: 16, fontWeight: 800 }}>No quotes yet</div>
        <div style={{ fontSize: 13, color: X.sub, maxWidth: 420 }}>The management imports the Quote Report into the data folder. {connected ? "Then click Refresh." : "Open the data folder on this computer (⋯ menu)."}</div>
      </div>
    )
  ) : tab === "today" ? (
    <TodayView quotes={scopedToday} cls="" meta={meta} thermo={thermo} today={today} notesBy={notesBy} excl={excl} prefs={prefs} onOpenQuote={openQuote} onGoMine={() => setTab("mine")} />
  ) : tab === "mine" ? (
    <QuoteTable {...tableProps} cards title="Quotes" filtersExtra={mob ? sheetFilters : null} filtersExtraN={mob ? nSheet : 0} />
  ) : tab === "calendar" || tab === "goal" ? (
    <CalendarView {...calProps} />
  ) : tab === "sellers" ? (
    <SellerReport quotes={scoped} periodRows={rows} meta={meta} acts={acts} groupOf={groupOf} today={today} filtersTxt={filtersTxt} periodTxtMain={periodLabelOf(period)} onOpenQuote={openQuote} />
  ) : tab === "analysis" ? (
    <AnalysisModal
      inline
      rows={rows}
      cancelRows={cancelRows}
      meta={meta}
      qIndex={qIndex}
      groupOf={groupOf}
      periodTxt={periodLabelOf(period)}
      sellerTxt={seller ? seller.slice(3) : "All salespeople"}
      filtersTxt={`${cls || "All classifications"} · ${line ? "line " + line : "all lines"}`}
      onOpenQuote={openQuote}
    />
  ) : (
    <QuoteTable {...tableProps} cards={mob} filtersExtra={mob ? sheetFilters : null} filtersExtraN={mob ? nSheet : 0} />
  );

  return (
    <div className="qp-shell" style={{ minHeight: "100%", background: X.bg, display: "flex", flexDirection: "column", color: X.ink }}>
      {css}
      {/* ---------- erro do arquivo de dados (fica até fechar) ---------- */}
      {dfErr && (
        <div className="no-print" style={{ display: "flex", gap: 12, alignItems: "flex-start", flexWrap: "wrap", padding: "12px 14px", borderRadius: 10, background: X.redSoft, border: "1px solid #F3B4B4", color: X.redDk }}>
          <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 1 }} />
          <div style={{ minWidth: 0, flex: "1 1 220px", display: "flex", flexDirection: "column", gap: 3 }}>
            <div style={{ fontSize: 13.5, fontWeight: 800 }}>{dfErr.title}</div>
            <div style={{ fontSize: 12.5, color: X.ink }}>{dfErr.hint}</div>
            <div style={{ fontSize: 11.5, color: X.sub, fontFamily: "ui-monospace, Consolas, monospace", wordBreak: "break-word" }}>Browser message: {dfErr.text}</div>
          </div>
          {FSA && (
            <button
              onClick={connectNew}
              title="Pick another folder: the data files are written there and used from now on"
              style={{ ...btn, height: 30, fontSize: 12, background: "#fff", borderColor: "#F3B4B4", color: X.redDk }}
            >
              <Database size={14} /> Save the data files in another folder…
            </button>
          )}
          {dfErr.kind === "save" && (
            <button onClick={() => saveNow()} title="Try to save again now" style={{ ...btn, height: 30, fontSize: 12, background: "#fff", borderColor: "#F3B4B4", color: X.redDk }}>
              Try again
            </button>
          )}
          <button onClick={() => setDfErr(null)} title="Close" style={{ ...btn, height: 30, padding: "0 8px", background: "transparent", borderColor: "#F3B4B4", color: X.redDk }}>
            <XIcon size={15} />
          </button>
        </div>
      )}
      {/* ---------- topo ---------- */}
      <div
        style={
          mob
            ? { position: "sticky", top: 0, zIndex: 1100, margin: "-10px -12px 0", padding: "10px 12px 8px", background: "rgba(244,247,249,.96)", backdropFilter: "blur(8px)", WebkitBackdropFilter: "blur(8px)", borderBottom: `1px solid ${X.line}`, display: "flex", flexDirection: "column", gap: 8 }
            : { display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }
        }
      >
        {mob ? (
          /* celular: o nome da aba, e só três botões (filtro, salvamento, menu) */
          <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
            <div style={{ minWidth: 0, flex: 1 }}>
              <div style={{ fontSize: 21, fontWeight: 800, letterSpacing: -0.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", lineHeight: 1.15 }}>{(tabs.find((t) => t[0] === tab) || [, "Quote Follow Up"])[1]}</div>
              <div style={{ fontSize: 12, color: R ? R.fg : X.sub, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                {lockSeller ? String(session).split(" ")[0] + " · " : ""}
                {R ? R.label : ""}
              </div>
            </div>
            <span className="no-print" style={{ display: "flex", alignItems: "center", gap: 7, flexShrink: 0 }}>
              {hasTopFilter && (
                <button onClick={() => setShowFilters(true)} title="Filters" style={iconBtn(nSheet ? { background: X.tealSoft, borderColor: "#9FD6C9", color: X.tealDk } : null)}>
                  <SlidersHorizontal size={18} />
                  {dotBadge(nSheet)}
                </button>
              )}
              {syncBtn}
              <UserChip compact name={session} admin={isAdmin} role={role} onLogout={logout} onUsers={isAdmin ? () => setTab("data") : null} items={menuItems} status={dfStatusTxt} />
            </span>
          </div>
        ) : (
        <div style={{ display: "flex", alignItems: "center", gap: mob ? 8 : 12, minWidth: 0 }}>
          <span style={{ width: mob ? 34 : 40, height: mob ? 34 : 40, borderRadius: 11, background: X.teal, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <TrendingUp size={mob ? 18 : 21} />
          </span>
          <div style={{ minWidth: 0, flex: mob ? 1 : "none" }}>
            <div style={{ fontSize: mob ? 16 : 20, fontWeight: 800, letterSpacing: -0.4, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Quote Follow Up</div>
            <div style={{ fontSize: 12, color: R ? R.fg : X.sub, fontWeight: 600, display: "flex", alignItems: "center", gap: 5, whiteSpace: "nowrap" }}>
              {R && React.createElement(R.icon, { size: 12 })}
              {R ? R.label : ""}
              {!mob && <span style={{ color: X.sub, fontWeight: 400 }}>· {role === "outside" ? "today's follow-ups and my quotes" : role === "inside" ? "portfolio, calendar and the plan of the day" : "results, salespeople, goal and data"}</span>}
            </div>
          </div>
        </div>
        )}
        {mob ? null : (
          <div className="no-print" style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            {dfChip}
            {role === "manager" && (
              <>
                <div style={{ width: 1, height: 26, background: X.line, margin: "0 2px" }} />
                {imports}
              </>
            )}
            <div style={{ width: 1, height: 26, background: X.line, margin: "0 2px" }} />
            {tabsTop}
            {refreshBtn}
            {saveBtn}
            <UserChip name={session} admin={isAdmin} role={role} onLogout={logout} onUsers={isAdmin ? () => setTab("data") : null} />
            {dataMenu}
          </div>
        )}
        {unsavedPop.render(
          unsavedDetail && (
            <UnsavedList
              detail={unsavedDetail}
              saving={df.st === "saving"}
              onSave={() => {
                unsavedPop.setPos(null);
                saveNow();
              }}
              onOpenQuote={(id) => {
                unsavedPop.setPos(null);
                openQuote(id);
              }}
            />
          ),
        )}
        <input
          ref={restoreRef}
          type="file"
          accept=".json"
          style={{ display: "none" }}
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            e.target.value = "";
            if (f) restore(f);
          }}
        />
        <input
          ref={loadDataRef}
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            e.target.value = "";
            if (f) loadDataFile(f);
          }}
        />
      </div>

      {df.st === "ask" && (
        <div style={{ ...card, padding: "10px 14px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", background: "#FFFBEB", borderColor: "#F3D08A", fontSize: 13 }}>
          <Link2 size={16} color="#8A5A12" />
          <span style={{ color: "#8A5A12", flex: "1 1 200px" }}>
            The browser needs your permission again to save in <b>{df.name}</b>. Reconnect before importing or editing, so nothing is left out of the file.
          </span>
          <button onClick={reconnect} style={{ ...btn, marginLeft: "auto", height: 32, background: X.teal, color: "#fff", borderColor: X.teal }}>
            <Link2 size={14} /> Reconnect
          </button>
        </div>
      )}

      {filterBar}
      {content}
      {tabsBottom}
      {filterSheet}

      {closeAsk && (
        <CloseDialog
          key={closeAsk.id + closeAsk.kind + closeAsk.mode}
          ask={closeAsk}
          onCancel={() => setCloseAsk(null)}
          onConfirm={(why) => {
            if (closeAsk.mode === "why" || closeAsk.manual) doWhy(closeAsk.id, closeAsk.c, why);
            else doClose(closeAsk.id, closeAsk.c, closeAsk.kind, why);
            setCloseAsk(null);
          }}
          onReopen={() => {
            doReopen(closeAsk.id, closeAsk.c);
            setCloseAsk(null);
          }}
        />
      )}
      {openQ && (
        <QuoteDrawer
          q={qIndex.get(openQ)}
          m={meta[openQ]}
          sales={salesView}
          acts={acts}
          thermo={qIndex.get(openQ) ? thermo[qIndex.get(openQ).nn] : null}
          onNote={onNote}
          onDone={onDone}
          onResched={onResched}
          onReq={onReq}
          onAskClose={askClose}
          onReopen={doReopen}
          onWhy={doWhy}
          focusDone={openDone}
          user={user}
          onUser={setUser}
          onClose={() => setOpenQ("")}
          nav={
            navIds && navIds.includes(openQ)
              ? {
                  i: navIds.indexOf(openQ),
                  n: navIds.length,
                  go: (d) => {
                    const j = navIds.indexOf(openQ) + d;
                    if (j >= 0 && j < navIds.length) {
                      setOpenQ(navIds[j]);
                      setOpenDone(false);
                    }
                  },
                }
              : null
          }
        />
      )}

      {toastEl}
    </div>
  );
}

function EmptyState({ onFile }) {
  const step = (n, title, text) => (
    <div style={{ display: "flex", gap: 12, alignItems: "flex-start", textAlign: "left" }}>
      <span style={{ width: 26, height: 26, borderRadius: 99, background: X.tealSoft, color: X.tealDk, fontWeight: 800, fontSize: 13, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{n}</span>
      <div>
        <div style={{ fontWeight: 700, fontSize: 13.5 }}>{title}</div>
        <div style={{ fontSize: 12.5, color: X.sub, marginTop: 2 }}>{text}</div>
      </div>
    </div>
  );
  return (
    <div style={{ ...card, padding: "40px 24px", display: "flex", justifyContent: "center" }}>
      <div style={{ maxWidth: 560, display: "flex", flexDirection: "column", gap: 18 }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 19, fontWeight: 800 }}>Start with the Quote Report</div>
          <div style={{ fontSize: 13, color: X.sub, marginTop: 6 }}>The files stay on this computer: nothing is uploaded anywhere.</div>
        </div>
        {step(1, "Quote Report", "Needs Quote ID, Customer and Create Date. Status, Amount, Sales Amt, Sales Id and Classification are used when present.")}
        {step(2, "Sales Report (optional)", "Booked lines with Sales Id: gives the real Sold value of each won quote and the booked total of the month for the goal.")}
        {step(3, "Fill in FUP and % Win", "Straight in the table (the Thermometer is automatic). Then open the Calendar, type the daily goal and click a day to see its plan.")}
        <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 4, flexWrap: "wrap" }}>
          <ImportButton label="Import Quote Report" icon={FileSpreadsheet} onFile={(f) => onFile(f, "quote")} accent />
          <ImportButton label="Import Sales Report" icon={Upload} onFile={(f) => onFile(f, "sales")} />
        </div>
      </div>
    </div>
  );
}