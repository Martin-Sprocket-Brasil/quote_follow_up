/* ===== Formulário DCS / Value Added (Martin) =====
   O programa preenche o PRÓPRIO formulário oficial (dcs-va-template.pdf), então quem recebe vê o mesmo documento
   de sempre — e os campos continuam editáveis no Acrobat.
   O modelo é o "VA_DCS Formulário DCS e VA - 26.01.26.pdf" com uma correção: no original, "Economia (monetizado
   ou não)" e "Documentado" eram O MESMO campo (mudar um mudava o outro); aqui são dois campos separados.
   Cada formulário salvo fica no Histórico de Atividades da cotação (tipo "dcs"), com todos os campos em JSON. */
import { PDFDocument, PDFName, StandardFonts } from "pdf-lib";

export const DCS_MAIL = "GKrebski@martinsprocket.com"; // para onde o botão ENVIAR do formulário original manda

// tipo → [rótulo oficial, escolha no grupo "Value Added / DCS / Deslealdade", escolha no grupo "Tipos de Economia"]
export const DCS_KINDS = {
  va: { pt: "Value Added", en: "Value Added", g1: "Choice1", g5: "Choice1" },
  dcs: { pt: "DCS", en: "DCS", g1: "Choice2", g5: "Choice2" },
  desl: { pt: "Deslealdade", en: "Disloyalty", g1: "Choice3", g5: null },
};
// subgrupos (Tipos de Economia): [chave, nome oficial no formulário, nome em inglês, caixa no PDF]
export const DCS_ITEMS = {
  va: [
    ["pass", "Passar negócio ao distribuidor", "Pass the business to the distributor", "Check Box4"],
    ["discount", "Descontos", "Discounts", "Check Box5"],
    ["return", "Receber devolução sem custo", "Free-of-charge return", "Check Box6"],
    ["freight", "Frete cortesia", "Courtesy freight", "Check Box7"],
    ["commission", "Comissão", "Commission", "Check Box8"],
    ["leadtime", "Prazo de entrega", "Lead time", "Check Box9"],
    ["payment", "Condição de pagamento especial", "Special payment terms", "Check Box10"],
  ],
  dcs: [
    ["downAvoid", "Tempo inativo evitado", "Downtime avoided", "Check Box11"],
    ["downReduce", "Tempo inativo reduzido", "Downtime reduced", "Check Box12"],
    ["emergency", "Atendimento de emergência", "Emergency service", "Check Box13"],
    ["afterHours", "Atendimento fora do horário", "After-hours service", "Check Box14"],
    ["productivity", "Aumento da produtividade", "Productivity increase", "Check Box15"],
    ["labor", "Redução trabalhista", "Labor reduction", "Check Box16"],
    ["oem", "Conversão de peças OEM", "OEM part conversion", "Check Box17"],
    ["priceDiff", "Diferenciação de preços", "Price differentiation", "Check Box18"],
    ["rebuild", "Reconstrução/retrabalho de produto", "Product rebuild / rework", "Check Box19"],
    ["energy", "Economia de energia", "Energy savings", "Check Box20"],
    ["freightSave", "Economia de frete", "Freight savings", "Check Box21"],
    ["invReduce", "Redução de Inventário", "Inventory reduction", "Check Box22"],
    ["invCommit", "Inventário Comprometido", "Committed inventory", "Check Box23"],
    ["maintTrain", "Treinamento de Manutenção", "Maintenance training", "Check Box24"],
    ["engAssist", "Assistência da Engenharia", "Engineering assistance", "Check Box25"],
    ["install", "Assistência com Instalação", "Installation assistance", "Check Box26"],
    ["reverse", "Engenharia Reversa", "Reverse engineering", "Check Box27"],
    ["prodTrain", "Treinamento em Produto", "Product training", "Check Box28"],
    ["field", "Pesquisa em Campo", "Field survey", "Check Box29"],
  ],
};
export const DCS_ITEM = Object.fromEntries(Object.values(DCS_ITEMS).flat().map(([k, pt, en, box]) => [k, { pt, en, box }]));
export const UFS = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"];

// o que é obrigatório no formulário (*) — devolve a lista do que falta
export function dcsMissing(f) {
  const miss = [];
  if (!DCS_KINDS[f.kind]) miss.push("kind");
  if ((f.kind === "va" || f.kind === "dcs") && !(f.items || []).length) miss.push("items");
  if (!f.qt) miss.push("qt");
  if (!f.monetized) miss.push("monetized");
  if (!String(f.inside || "").trim()) miss.push("inside");
  if (!String(f.endUser || "").trim()) miss.push("endUser");
  if (!String(f.city || "").trim()) miss.push("city");
  if (!String(f.desc || "").trim()) miss.push("desc");
  if (!f.documented) miss.push("documented");
  return miss;
}

// a fonte padrão do PDF (Helvetica) só tem o alfabeto latino: o resto sai do texto em vez de quebrar o PDF
const pdfSafe = (s) =>
  String(s == null ? "" : s)
    .replace(/\r\n?/g, "\n")
    .replace(/[‐-‒]/g, "-")
    .replace(/[^\n\x20-\x7E -ÿ–—‘’‚“”„†‡•…‰€™]/g, "");

/* preenche o formulário oficial com f = { kind, items[], so, qt, inside, outside, endUser, distributor, city, state,
   desc, monetized: "Sim"|"Não", value, currency: "BRL"|"USD", documented: "Sim"|"Não" } → bytes do PDF */
export async function fillDcsPdf(template, f) {
  const pdf = await PDFDocument.load(template);
  const form = pdf.getForm();
  const text = (name, v) => {
    const fld = form.getTextField(name);
    fld.setText(pdfSafe(v).trim() || undefined);
  };
  const kind = DCS_KINDS[f.kind];
  text("Text2", f.so);
  text("Text3", f.qt);
  text("Text4", f.inside);
  text("Text5", f.outside);
  text("Text6", f.endUser);
  text("Text1", f.distributor);
  text("Text7", f.city);
  text("Text8", f.state);
  const desc = form.getTextField("Text9");
  desc.enableMultiline();
  desc.acroField.setDefaultAppearance("0 g /Helv 10 Tf"); // no original o campo não tem tamanho de letra
  text("Text9", f.desc);
  text("Text10", f.value);
  const pick = (name, v) => {
    const d = form.getDropdown(name);
    if (v && d.getOptions().includes(v)) d.select(v);
    else d.clear();
  };
  pick("Economia", f.monetized);
  pick("Documentado", f.documented);
  pick("Moeda", f.value ? f.currency || "BRL" : "");
  /* bolinhas e caixas: só o valor, com o desenho do próprio formulário (o ✔ e a bolinha originais). Pelo caminho
     normal o pdf-lib redesenharia o ✔ com uma fonte de símbolos de tamanho zero, que sai como um borrão. */
  const mark = (name, on) => {
    const af = form.getField(name).acroField;
    const v = on ? PDFName.of(on) : PDFName.of("Off");
    af.dict.set(PDFName.of("V"), v);
    af.getWidgets().forEach((w) => {
      const st = w.getOnValue && w.getOnValue();
      w.setAppearanceState(on && st && st.asString() === "/" + on ? PDFName.of(on) : PDFName.of("Off"));
    });
  };
  mark("Group1", kind ? kind.g1 : "");
  mark("Group5", kind && kind.g5 ? kind.g5 : "");
  const on = new Set(f.kind === "va" || f.kind === "dcs" ? (f.items || []).filter((k) => DCS_ITEMS[f.kind].some((x) => x[0] === k)) : []);
  Object.values(DCS_ITEMS)
    .flat()
    .forEach(([k, , , box]) => mark(box, on.has(k) ? "Yes" : ""));
  const helv = await pdf.embedFont(StandardFonts.Helvetica);
  form.updateFieldAppearances(helv);
  pdf.setTitle(`Formulário DCS / Value Added · ${pdfSafe(f.qt)} · ${pdfSafe(f.endUser)}`);
  pdf.setSubject(kind ? kind.pt : "");
  pdf.setCreator("Quote Follow Up");
  pdf.setProducer("Quote Follow Up");
  return pdf.save();
}
// nome do arquivo: Formulario DCS-VA QT00197476 MARCHER BRASIL.pdf
export const dcsFileName = (f) =>
  `Formulario ${f.kind === "desl" ? "Deslealdade" : f.kind === "dcs" ? "DCS" : "VA"} ${f.qt || ""} ${String(f.endUser || "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 50)}.pdf`.replace(/\s+/g, " ");
