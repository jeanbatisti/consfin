const fs = require("node:fs");
const path = require("node:path");
const model = require("./app.js");
const reportEngine = require("./report-engine.js");

const outputArg = process.argv[2];
if (!outputArg) {
  throw new Error("Informe o caminho do HTML de saída.");
}

const outputPath = path.resolve(outputArg);
let styles = ["styles.css", "report-print.css"]
  .map((fileName) => fs.readFileSync(path.join(__dirname, fileName), "utf8"))
  .join("\n");
const fontDirectory = path.join(__dirname, "assets", "fonts");
for (const fontName of fs.readdirSync(fontDirectory).filter((name) => name.endsWith(".woff2"))) {
  const fontData = fs.readFileSync(path.join(fontDirectory, fontName)).toString("base64");
  styles = styles.replaceAll(
    `assets/fonts/${fontName}`,
    `data:font/woff2;base64,${fontData}`,
  );
}
const result = model.calculateScenario(model.defaultInputs);
const report = reportEngine.buildDecisionReport(result, {
  clientName: "Mariana Alves",
  plannerName: "Jean Batisti",
  priority: "wealth",
  source: "standalone",
  generatedAt: "2026-07-22",
});

const html = `<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Estratégias de Crédito - Mariana Alves</title>
    <style>${styles}</style>
  </head>
  <body>
    <section class="print-report">${reportEngine.renderReportHtml(report)}</section>
  </body>
</html>`;

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, html, "utf8");
console.log(outputPath);
