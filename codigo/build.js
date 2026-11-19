/* Compila codigo/app.jsx -> pagina/QuoteFollowUp.html (arquivo unico, offline). Rode pelo build.bat. */
const esbuild = require("esbuild");
const fs = require("fs");
const path = require("path");

const SRC = __dirname; // pasta codigo/
const OUT = path.join(SRC, "..", "pagina", "QuoteFollowUp.html");
const CSS = path.join(SRC, "styles.css");
const style = fs.existsSync(CSS)
  ? fs.readFileSync(CSS, "utf8")
  : "html,body,#root{margin:0;height:100%}";

console.log("Empacotando o codigo (esbuild)...");
const res = esbuild.buildSync({
  entryPoints: [path.join(SRC, "entry.jsx")],
  bundle: true,
  minify: true,
  format: "iife",
  loader: { ".jsx": "jsx", ".pdf": "binary" }, // o formulário DCS / Value Added (PDF) vai dentro da página
  write: false,
});
/* "</script", "</body", "</html" e "</head" dentro do código (a biblioteca de Excel tem "</body></html>" num texto)
   viram "<\/…": para o JavaScript é a mesma coisa, mas assim nem o navegador nem o Live Server confundem esse
   texto com o fim da página (o Live Server injeta o script dele antes do primeiro "</body>" que encontra). */
const dist = res.outputFiles[0].text.replace(/<\/(script|body|html|head)/gi, "<\\/$1");

const html =
`<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Quote Follow Up</title>
    <style>${style}</style>
  </head>
  <body>
    <div id="root"><div id="boot">Carregando o programa...</div></div>
    <script>${dist}</script>
  </body>
</html>`;

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, html);
console.log("Pronto! Gerado pagina/QuoteFollowUp.html (" + (Buffer.byteLength(html) / 1024 / 1024).toFixed(2) + " MB)");
