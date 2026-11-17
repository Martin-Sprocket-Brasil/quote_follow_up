/* Compila app.jsx -> QuoteFollowUp.html (arquivo unico, offline). */
const esbuild = require("esbuild");
const fs = require("fs");

const OUT = "QuoteFollowUp.html";
const style = fs.existsSync("styles.css")
  ? fs.readFileSync("styles.css", "utf8")
  : "html,body,#root{margin:0;height:100%}";

console.log("Empacotando o codigo (esbuild)...");
const res = esbuild.buildSync({
  entryPoints: ["entry.jsx"],
  bundle: true,
  minify: true,
  format: "iife",
  loader: { ".jsx": "jsx" },
  write: false,
});
const dist = res.outputFiles[0].text.replace(/<\/script>/g, "<\\/script>");

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

fs.writeFileSync(OUT, html);
console.log("Pronto! Gerado " + OUT + " (" + (Buffer.byteLength(html) / 1024 / 1024).toFixed(2) + " MB)");
