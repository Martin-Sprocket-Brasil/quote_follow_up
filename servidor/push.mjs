/* =====================================================================================
   Avisos no celular / PC (Web Push)
   -------------------------------------------------------------------------------------
   • Cada aparelho que tocou em "Ativar avisos" fica na tabela push_subs (um registro por aparelho).
   • Quem manda é este servidor, pelos serviços de aviso do próprio navegador (Google, Apple, Microsoft,
     Mozilla). Chega com a tela apagada e o navegador fechado. Não precisa de conta: as chaves (VAPID)
     são criadas na primeira vez e ficam na tabela settings.
   • Quando avisa: importação do Quote Report (QTs novas / ganhas de cada pessoa), meta diária nova,
     anotação de outra pessoa numa QT sua, formulário DCS novo (gerência) e o lembrete da manhã
     (atrasados, os de hoje e os que vencem no próximo dia útil) nos dias úteis.
   • Quem fez a ação não recebe aviso dela.
   ===================================================================================== */
import https from "https";
import * as webpush from "web-push";

const RX_HOUR = /^([01]\d|2[0-3]):[0-5]\d$/;
// só os serviços de aviso dos navegadores (o servidor não manda nada para outro endereço)
const PUSH_HOSTS = [/(^|\.)googleapis\.com$/, /(^|\.)push\.apple\.com$/, /(^|\.)notify\.windows\.com$/, /(^|\.)push\.services\.mozilla\.com$/];
const TEST = process.env.QFU_PUSH_TEST === "1"; // testes automáticos: aceita um serviço falso em 127.0.0.1
const testAgent = TEST ? new https.Agent({ rejectUnauthorized: false }) : null;

export function okEndpoint(u) {
  try {
    const x = new URL(String(u));
    if (TEST && x.protocol === "https:" && x.hostname === "127.0.0.1") return true;
    return x.protocol === "https:" && !x.port && PUSH_HOSTS.some((rx) => rx.test(x.hostname));
  } catch {
    return false;
  }
}

export function createPush({ pool, core, scopeOf, MGR, ADMIN_USER, proxy = "", log = console }) {
  let keys = null; // { pub, priv }
  let ready = false;
  let subject = "";
  const st = { hour: "08:00" };

  const getSetting = async (name) => {
    const r = await pool.query("SELECT value FROM settings WHERE name = ?", [name]);
    return r.length ? r[0].value : null;
  };
  const putSetting = (name, value) => pool.query("INSERT INTO settings (name, value) VALUES (?, ?) ON DUPLICATE KEY UPDATE value = VALUES(value)", [name, String(value).slice(0, 255)]);

  async function init() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS push_subs (
        id          INT UNSIGNED NOT NULL AUTO_INCREMENT,
        endpoint    VARCHAR(700) NOT NULL,
        name        VARCHAR(100) NOT NULL,
        p256dh      VARCHAR(200) NOT NULL,
        auth        VARCHAR(60)  NOT NULL,
        user_agent  VARCHAR(200) NOT NULL DEFAULT '',
        created_at  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
        last_ok     DATETIME     NULL,
        PRIMARY KEY (id),
        UNIQUE KEY ux_push_endpoint (endpoint),
        KEY ix_push_name (name)
      ) ENGINE=InnoDB`);
    } catch {
      const has = await pool.query("SHOW TABLES LIKE 'push_subs'");
      if (!has.length) {
        log.warn("\n[AVISO] Avisos no celular desligados: a tabela push_subs não existe e o usuário do banco não pode criá-la.\n        Rode o schema.sql novo como root: mariadb -u root -p < banco\\schema.sql\n");
        return;
      }
    }
    let pub = await getSetting("vapid_public"),
      priv = await getSetting("vapid_private");
    if (!pub || !priv) {
      const k = webpush.generateVAPIDKeys();
      pub = k.publicKey;
      priv = k.privateKey;
      await putSetting("vapid_public", pub);
      await putSetting("vapid_private", priv);
      log.info("[avisos] chaves dos avisos criadas (ficam no banco, tabela settings)");
    }
    keys = { pub, priv };
    subject = (await getSetting("push_subject")) || "";
    const h = await getSetting("push_hour");
    if (h != null) st.hour = h === "off" || RX_HOUR.test(h) ? h : "08:00";
    ready = true;
    setInterval(() => morning().catch((e) => log.error("[avisos] lembrete da manhã:", e.message || e)), 60000).unref();
  }

  /* ---------- aparelhos ---------- */
  async function subscribe(name, sub, ua, origin) {
    if (!ready) throw Object.assign(new Error("Push is off on the server."), { status: 503, code: "push-off" });
    const ep = String((sub && sub.endpoint) || ""),
      k = (sub && sub.keys) || {};
    if (!okEndpoint(ep) || ep.length > 700 || !/^[A-Za-z0-9_-]{20,200}$/.test(String(k.p256dh || "")) || !/^[A-Za-z0-9_-]{8,60}$/.test(String(k.auth || "")))
      throw Object.assign(new Error("Invalid subscription."), { status: 400, code: "bad-request" });
    await pool.query(
      "INSERT INTO push_subs (endpoint, name, p256dh, auth, user_agent) VALUES (?, ?, ?, ?, ?) ON DUPLICATE KEY UPDATE name = VALUES(name), p256dh = VALUES(p256dh), auth = VALUES(auth), user_agent = VALUES(user_agent)",
      [ep, name, k.p256dh, k.auth, String(ua || "").slice(0, 200)],
    );
    // endereço do site (o "dono" dos avisos para a Apple / Google): o do túnel, guardado na primeira vez
    if (origin && /^https:\/\/[a-z0-9.-]+$/i.test(origin) && origin !== subject && !/localhost|127\.0\.0\.1/.test(origin)) {
      subject = origin;
      await putSetting("push_subject", origin);
    }
  }
  const unsubscribe = (endpoint, name) => pool.query("DELETE FROM push_subs WHERE endpoint = ?" + (name ? " AND name = ?" : ""), name ? [String(endpoint || ""), name] : [String(endpoint || "")]);
  const unsubscribeUser = (name) => pool.query("DELETE FROM push_subs WHERE name = ?", [name]);
  const count = async (name) => (await pool.query("SELECT COUNT(*) AS n FROM push_subs WHERE name = ?", [name]))[0].n;

  /* ---------- envio ---------- */
  async function sendTo(names, payload) {
    if (!ready || !names.length) return 0;
    const subs = await pool.query("SELECT id, endpoint, p256dh, auth, name FROM push_subs WHERE name IN (?)", [names]);
    if (!subs.length) return 0;
    const vapid = { subject: subject || "https://quote-follow-up.example", publicKey: keys.pub, privateKey: keys.priv };
    const body = JSON.stringify({ title: "Quote Follow Up", url: "/", ...payload });
    let ok = 0;
    await Promise.all(
      subs.map(async (s) => {
        try {
          await webpush.sendNotification({ endpoint: s.endpoint, keys: { p256dh: s.p256dh, auth: s.auth } }, body, { vapidDetails: vapid, TTL: payload.ttl || 6 * 3600, urgency: "normal", timeout: 15000, ...(testAgent ? { agent: testAgent } : proxy ? { proxy } : null) });
          ok++;
          pool.query("UPDATE push_subs SET last_ok = NOW() WHERE id = ?", [s.id]).catch(() => {});
        } catch (e) {
          // 404 / 410 = o aparelho desativou ou trocou de navegador: sai da lista
          if (e && (e.statusCode === 404 || e.statusCode === 410)) await pool.query("DELETE FROM push_subs WHERE id = ?", [s.id]).catch(() => {});
          else log.warn(`[avisos] não foi para um aparelho de ${s.name}: ${(e && (e.statusCode || e.code || e.message)) || e}`);
        }
      }),
    );
    return ok;
  }

  // pessoas que têm aparelho cadastrado, com o papel (o Admin conta como gerência)
  async function people(except) {
    const r = await pool.query("SELECT DISTINCT p.name, u.role FROM push_subs p JOIN users u ON u.name = p.name");
    return r
      .map((x) => ({ name: x.name, role: x.name === ADMIN_USER ? "admin" : x.role }))
      .filter((x) => x.role && x.name !== except)
      .map((x) => ({ ...x, mgr: MGR.has(x.role) }));
  }
  const names3 = (list) => {
    const n = [...new Set(list.filter(Boolean))];
    return n.slice(0, 3).join(", ") + (n.length > 3 ? ` +${n.length - 3}` : "");
  };
  const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

  /* depois de cada gravação (sync): importação, meta, anotação, formulário */
  async function afterSync({ me, rep, newIds, acts: actsIn, quotes, goalBefore, goalAfter }) {
    if (!ready) return;
    const ppl = await people(me.name);
    if (!ppl.length) return;
    const allActs = (actsIn || []).filter((a) => a && typeof a === "object");
    const scopes = new Map(ppl.map((p) => [p.name, p.mgr ? () => true : scopeOf({ name: p.name, role: p.role }, rep)]));
    const sends = [];
    // 1) Quote Report: QTs que entraram agora e as que viraram ganhas
    {
      const fresh = newIds && newIds.size ? (quotes || []).filter((q) => q && newIds.has(String(q.id).toUpperCase()) && !quoteCancelledSafe(q)) : [];
      const acts = allActs.filter((a) => a.k === "status" || a.k === "order");
      const won = acts.filter((a) => a.k === "status" && /won/i.test(String(a.n || "")));
      const orders = acts.filter((a) => a.k === "order");
      for (const p of ppl) {
        const ok = scopes.get(p.name);
        const f = fresh.filter((q) => ok(String(q.id).toUpperCase()));
        const w = [...new Set(won.filter((a) => ok(String(a.id).toUpperCase())).map((a) => a.id))];
        const o = [...new Set(orders.filter((a) => ok(String(a.id).toUpperCase())).map((a) => a.id))];
        if (!f.length && !w.length && !o.length) continue;
        const title = f.length ? (p.mgr ? `Equipe: ${plural(f.length, "QT nova", "QTs novas")}` : plural(f.length, "QT nova", "QTs novas")) : w.length ? plural(w.length, "cotação ganha", "cotações ganhas") : plural(o.length, "pedido novo (SO)", "pedidos novos (SO)");
        const extra = [f.length && w.length && plural(w.length, "ganha", "ganhas"), (f.length || w.length) && o.length && plural(o.length, "pedido novo", "pedidos novos")].filter(Boolean);
        const bodyTxt = [f.length ? names3(f.map((q) => q.c)) : names3([...w, ...o].map((id) => (rep.byId.get(id) || {}).c)), ...extra, `importado por ${me.name}`].filter(Boolean).join(" · ");
        sends.push(sendTo([p.name], { title, body: bodyTxt, tag: "import" }));
      }
    }
    // 2) meta diária nova (total ou por linha)
    if (goalBefore && goalAfter && goalBefore.join() !== goalAfter.join()) {
      const lines = goalAfter.slice(1).some((v) => v > 0) ? ["PT", "CP", "MH"].map((k, i) => `${k} ${core.moneyShort(goalAfter[i + 1])}`).join(" · ") : "";
      sends.push(sendTo(ppl.map((p) => p.name), { title: `Nova meta diária: ${core.money(goalAfter[0])}`, body: [`antes ${core.money(goalBefore[0])}`, lines, `por ${me.name}`].filter(Boolean).join(" · "), tag: "goal" }));
    }
    // 3) anotação de outra pessoa numa QT sua · 4) formulário DCS novo (gerência)
    const notes = allActs.filter((a) => a.k === "note" && a.txt);
    if (notes.length)
      for (const p of ppl.filter((x) => !x.mgr)) {
        const mine = notes.filter((a) => scopes.get(p.name)(String(a.id).toUpperCase()));
        if (!mine.length) continue;
        const a = mine[mine.length - 1];
        sends.push(sendTo([p.name], { title: mine.length === 1 ? `Anotação de ${me.name}` : `${mine.length} anotações de ${me.name}`, body: `${a.id} · ${a.c || ""} · ${String(a.txt).slice(0, 120)}`, tag: "note-" + a.id }));
      }
    const dcs = allActs.filter((a) => a.k === "dcs");
    if (dcs.length) {
      const mg = ppl.filter((x) => x.mgr).map((x) => x.name);
      if (mg.length) sends.push(sendTo(mg, { title: dcs.length === 1 ? "Formulário DCS / VA novo" : `${dcs.length} formulários DCS / VA novos`, body: `${names3(dcs.map((a) => `${a.id} ${a.c || ""}`.trim()))} · por ${me.name}`, tag: "dcs" }));
    }
    await Promise.all(sends);
  }
  const quoteCancelledSafe = (q) => {
    try {
      return core.quoteCancelled({ st: q.st || "" });
    } catch {
      return false;
    }
  };

  /* ---------- lembrete da manhã (dias úteis, no horário escolhido pela gerência) ---------- */
  const localNow = () => {
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    return { day: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`, hm: `${p(d.getHours())}:${p(d.getMinutes())}` };
  };
  async function morning(force = false, onlyName = "") {
    if (!ready || (st.hour === "off" && !force)) return 0;
    const { day, hm } = localNow();
    if (!force) {
      if (core.isOffDay(day) || hm < st.hour || hm >= "18:00") return 0; // fim de semana, feriado, cedo ou tarde demais
      if ((await getSetting("push_morning_day")) === day) return 0;
      await putSetting("push_morning_day", day); // antes de mandar: nunca manda duas vezes no mesmo dia
    }
    const ppl = (await people("")).filter((p) => !onlyName || p.name === onlyName);
    if (!ppl.length) return 0;
    const rep = await getRepRef();
    const tom = core.addBusinessDays(day, 1);
    const [fups, qs] = await Promise.all([
      pool.query("SELECT quote_id, fup, fup_since, cycle_step, next_date, closed_as FROM followups WHERE fup IS NOT NULL AND closed_as IS NULL"),
      pool.query("SELECT quote_id, status FROM quotes"),
    ]);
    const stOf = new Map(qs.map((q) => [q.quote_id, q.status || ""]));
    const rows = [];
    for (const f of fups) {
      if (!stOf.has(f.quote_id) || core.quoteState({ st: stOf.get(f.quote_id) }) !== "open") continue;
      const m = { fup: f.fup };
      if (f.fup_since) m.since = String(f.fup_since).slice(0, 10);
      if (f.cycle_step) m.step = f.cycle_step;
      if (f.next_date) m.at = String(f.next_date).slice(0, 10);
      const n = core.fupPending(m)[0];
      if (!n) continue;
      rows.push({ id: f.quote_id, k: n < day ? "late" : n === day ? "due" : n === tom ? "soon" : "" });
    }
    const tomTxt = tom === core.addDays(day, 1) ? "amanhã" : "no próximo dia útil";
    let sent = 0;
    for (const p of ppl) {
      const ok = p.mgr ? () => true : scopeOf({ name: p.name, role: p.role }, rep);
      const c = { late: 0, due: 0, soon: 0 };
      rows.forEach((r) => r.k && ok(r.id) && c[r.k]++);
      if (!c.late && !c.due && !c.soon) continue;
      const parts = [c.late && plural(c.late, "atrasado", "atrasados"), c.due && `${c.due} para hoje`, c.soon && `${c.soon} ${c.soon === 1 ? "vence" : "vencem"} ${tomTxt} (prestes a atrasar)`].filter(Boolean);
      sent += await sendTo([p.name], { title: p.mgr ? "Equipe: follow-ups de hoje" : "Seus follow-ups de hoje", body: parts.join(" · "), tag: "morning", ttl: 10 * 3600 });
    }
    if (sent) log.info(`[avisos] lembrete da manhã enviado para ${sent} aparelho(s)`);
    return sent;
  }
  let getRepRef = async () => ({ byId: new Map(), spGroups: [], qbGroups: [] });

  return {
    init,
    get ready() {
      return ready;
    },
    get publicKey() {
      return keys ? keys.pub : "";
    },
    get hour() {
      return st.hour;
    },
    async setHour(h) {
      if (h !== "off" && !RX_HOUR.test(h)) throw Object.assign(new Error("Invalid time."), { status: 400, code: "bad-request" });
      st.hour = h;
      await putSetting("push_hour", h);
      // mudou para um horário que ainda não passou hoje: o lembrete de hoje volta a valer
      const { day, hm } = localNow();
      if (h !== "off" && h > hm) await pool.query("DELETE FROM settings WHERE name = 'push_morning_day' AND value = ?", [day]);
    },
    setRepGetter(fn) {
      getRepRef = fn;
    },
    subscribe,
    unsubscribe,
    unsubscribeUser,
    count,
    sendTo,
    afterSync,
    morning,
  };
}
