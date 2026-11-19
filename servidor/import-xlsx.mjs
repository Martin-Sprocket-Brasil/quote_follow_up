/* =====================================================================================
   QUOTE FOLLOW UP — importação das planilhas para o MariaDB
   Lê "Quote Follow Up data.xlsx" e "Quote Follow Up reports.xlsx" com o PRÓPRIO core.js do programa
   (mesmas regras de leitura, inclusive a aba Config cifrada dos usuários) e gera o arquivo import.sql.

   Como usar (na pasta do programa — rode o build.bat uma vez antes; ou dê dois cliques no importar-banco.bat):
     node servidor/import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx"
     mariadb -u root -p < banco/import.sql

   Pode rodar quantas vezes quiser: o import.sql atualiza o que já existe e não duplica nada.
   O histórico (activities) só ganha os registros que ainda não estão no banco.
   ===================================================================================== */
import fs from "fs";
import path from "path";
import { fileURLToPath, pathToFileURL } from "url";
import * as XLSX from "xlsx";

const here = path.dirname(fileURLToPath(import.meta.url));
const [dataPath, repPath, outArg] = process.argv.slice(2);
if (!dataPath || !repPath) {
  console.error('Uso: node servidor/import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx" [import.sql]');
  process.exit(1);
}
const OUT = outArg || path.join(here, "..", "banco", "import.sql");

/* core.js é ESM com extensão .js: uma cópia temporária .mjs garante que o Node o carregue como módulo
   em qualquer versão (a cópia fica na mesma pasta para achar o node_modules/xlsx). */
const tmpCore = path.join(here, ".core-import.tmp.mjs");
fs.copyFileSync(path.join(here, "..", "codigo", "core.js"), tmpCore);
let core;
try {
  core = await import(pathToFileURL(tmpCore).href);
} finally {
  fs.rmSync(tmpCore, { force: true });
}

const readWb = (p) => XLSX.read(fs.readFileSync(p), { type: "buffer" });
const { buildLegacyImport } = await import(pathToFileURL(path.join(here, "legacy-import.mjs")).href);
// mesmas regras da importação pela página (aba Data → Importar da versão anterior); aqui o arquivo vale ("replace")
const { statements, counts, warnings } = buildLegacyImport(core, readWb(dataPath), readWb(repPath), { mode: "replace" });
if (warnings.includes("users-bad")) console.warn("ATENÇÃO: a aba Config (usuários) foi alterada por fora e não pôde ser lida: os usuários ficam de fora.");

const out = [
  "-- Gerado por import-xlsx.mjs em " + core.nowStamp(),
  `-- data: ${path.basename(dataPath)} · reports: ${path.basename(repPath)}`,
  "SET NAMES utf8mb4;",
  "USE quote_follow_up;",
  "START TRANSACTION;",
  ...statements.map((x) => x.sql + ";"),
  "COMMIT;",
];
fs.writeFileSync(OUT, out.join("\n\n") + "\n");

console.log(`Pronto! ${path.relative(process.cwd(), OUT) || OUT} (${(fs.statSync(OUT).size / 1024 / 1024).toFixed(1)} MB)`);
console.log(`  cotações ${counts.quotes} · pedidos ${counts.orders} · product groups ${counts.pgs} · vendas ${counts.sales}`);
console.log(`  follow-ups ${counts.followups} · dias com Done ${counts.done} · histórico ${counts.acts} · metas por dia ${counts.goals}`);
console.log(`  usuários ${counts.users} (${Object.entries(counts.roles).map(([k, v]) => `${k === "none" ? "sem acesso" : k}: ${v}`).join(", ")})`);
if (counts.winFixed) console.log(`  ${counts.winFixed} % Win fora de 25/50/75/99 ficaram vazios`);
console.log("  Dica: dá para importar estes arquivos também pela página (Admin → aba Data → Importar da versão anterior).");
