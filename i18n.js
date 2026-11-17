/* ===== Idioma (inglês / português) =====
   O código está em inglês; _t("texto") devolve a tradução do i18n-pt.js quando o idioma é português.
   _t("{0} quotes", n) troca {0}, {1}… pelos valores. A escolha fica salva neste navegador (menu do usuário ou
   tela de entrada) e a página recarrega para trocar tudo de uma vez. Sem escolha: o idioma do navegador. */
import { PT } from "./i18n-pt.js";
import { FUP, TEMP, CLOSE_LABEL, MON_SHORT, MONTHS_EN, HOLIDAYS_FIXED, HOLIDAYS_EASTER } from "./core.js";

const KEY = "qp_lang";
export const LANG = (() => {
  try {
    const v = localStorage.getItem(KEY);
    if (v === "en" || v === "pt") return v;
  } catch {}
  try {
    return /^pt/i.test(navigator.language || "") ? "pt" : "en";
  } catch {
    return "en";
  }
})();
const has = Object.prototype.hasOwnProperty;
export function _t(s, ...a) {
  let r = LANG === "pt" && has.call(PT, s) ? PT[s] : s;
  if (a.length) r = String(r).replace(/\{(\d+)\}/g, (m, i) => (a[+i] == null ? "" : String(a[+i])));
  return r;
}
// motivo de fechamento: gravado em inglês ("Price"); na tela, no idioma escolhido
export const whyTxt = (v) => (v ? _t(v) : v);
export function setLang(v) {
  try {
    localStorage.setItem(KEY, v);
  } catch {}
  location.reload();
}

/* rótulos que vêm do core.js (só a tela; os valores gravados — A, HOT, lost — não mudam) */
if (LANG === "pt") {
  const fup = {
    A: ["Próximo dia", "A · Próximo dia útil"],
    B: ["2× em 5 dias", "B · Duas vezes em 5 dias (2º e 4º dia útil)"],
    C: ["2 semanas", "C · A cada 2 semanas"],
    D: ["1 mês", "D · Todo mês"],
    E: ["Contínuo", "E · Contínuo (a próxima data é escolhida no Feito)"],
  };
  for (const k in fup) if (FUP[k]) [FUP[k].label, FUP[k].long] = fup[k];
  const temp = { HOT: "Quente", WARM: "Morno", COOL: "Frio", COLD: "Gelado" };
  for (const k in temp) if (TEMP[k]) TEMP[k].label = temp[k];
  CLOSE_LABEL.cancelled = "Cancelada";
  CLOSE_LABEL.lost = "Perdida";
  ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"].forEach((m, i) => (MON_SHORT[i] = m));
  ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"].forEach((m, i) => (MONTHS_EN[i] = m));
  const hol = {
    "01-01": "Ano Novo",
    "04-10": "Aniversário de Artur Nogueira",
    "04-21": "Tiradentes",
    "05-01": "Dia do Trabalhador",
    "07-09": "Revolução Constitucionalista (SP)",
    "09-07": "Independência",
    "09-15": "N. Sra. das Dores (Artur Nogueira)",
    "10-12": "N. Sra. Aparecida",
    "11-02": "Finados",
    "11-15": "Proclamação da República",
    "11-20": "Consciência Negra",
    "12-25": "Natal",
  };
  for (const k in hol) if (HOLIDAYS_FIXED[k]) HOLIDAYS_FIXED[k].name = hol[k];
  const easter = { "-48": "Segunda de Carnaval", "-47": "Terça de Carnaval", "-2": "Sexta-feira Santa", 60: "Corpus Christi" };
  HOLIDAYS_EASTER.forEach((h) => easter[h.off] && (h.name = easter[h.off]));
}
