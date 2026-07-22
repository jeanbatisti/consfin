const assert = require("node:assert/strict");
const fs = require("node:fs");
const engine = require("./report-engine.js");
const model = require("./app.js");

const fixedMeta = {
  clientName: "Marina Alves",
  plannerName: "Jean Batisti",
  generatedAt: "2026-07-22",
};

function strategy(id, overrides = {}) {
  const has = (key) => Object.prototype.hasOwnProperty.call(overrides, key);
  const isEnabled = has("isEnabled") ? overrides.isEnabled : true;
  const isViable = has("isViable") ? overrides.isViable : true;
  const isAffordable = has("isAffordable") ? overrides.isAffordable : isEnabled ? isViable : true;
  const lifetimeContractualOutflow = has("contractualOutflowsLifetime")
    ? undefined
    : has("lifetimeContractualOutflow")
      ? overrides.lifetimeContractualOutflow
      : has("totalOutputs")
        ? overrides.totalOutputs
        : 50_000;

  return {
    id,
    name: `Estratégia ${id.toUpperCase()}`,
    short: id.toUpperCase(),
    isEnabled,
    isAffordable,
    isViable,
    adjustedToday: 100_000,
    totalOutputs: 50_000,
    lifetimeContractualOutflow,
    acquisitionMonth: 1,
    payoffMonth: 60,
    debt: 0,
    ...overrides,
  };
}

function scenario(strategies, inputOverrides = {}) {
  return {
    inputs: {
      assetType: "Imovel",
      usage: "Aquisicao",
      propertyValue: 500_000,
      reserve: 120_000,
      monthlyCash: 8_000,
      horizonMonths: 240,
      financingSystem: "SAC",
      annualReturn: 0.08,
      inflationPct: 0.045,
      ...inputOverrides,
    },
    strategies,
  };
}

{
  const report = engine.buildDecisionReport({}, fixedMeta);
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.INVALID);
  assert.equal(report.schemaVersion, engine.SCHEMA_VERSION);
  assert.equal(report.engineVersion, engine.VERSION);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { isViable: false, adjustedToday: null, totalOutputs: null }),
      strategy("b", { isViable: false, adjustedToday: null, totalOutputs: null }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.NO_VIABLE);
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.doesNotMatch(report.viewModel.tradeoffSummary, /alternativa indicada/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, totalOutputs: 90_000 }),
      strategy("b", { isViable: false }),
    ]),
    { ...fixedMeta, priority: "compare" },
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.SINGLE_VIABLE);
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.match(report.viewModel.summary, /única opção viável/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, totalOutputs: 90_000 }),
      strategy("b", { isViable: false }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.SINGLE_VIABLE);
  assert.equal(report.outcome.recommendedStrategyId, "a");
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, totalOutputs: 70_000 }),
      strategy("b", { adjustedToday: 100_000, totalOutputs: 120_000 }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.SAME_LEADER);
  assert.equal(report.facts.leaders.wealth, "a");
  assert.equal(report.facts.leaders.cost, "a");
}

{
  const input = scenario([
    strategy("wealth", { adjustedToday: 300_000, totalOutputs: 180_000, payoffMonth: 180 }),
    strategy("cost", { adjustedToday: 200_000, totalOutputs: 90_000, payoffMonth: 36 }),
  ]);
  const wealthReport = engine.buildDecisionReport(input, fixedMeta);
  const costReport = engine.buildDecisionReport(input, { ...fixedMeta, priority: "cost" });

  assert.equal(wealthReport.meta.priority, "wealth", "wealth deve ser a prioridade padrão");
  assert.equal(wealthReport.outcome.state, engine.OUTCOME_STATES.SPLIT_LEADERS);
  assert.equal(wealthReport.outcome.recommendedStrategyId, "wealth");
  assert.equal(costReport.outcome.recommendedStrategyId, "cost");
  assert.match(wealthReport.viewModel.tradeoffSummary, /acrescenta/i);
  assert.match(costReport.viewModel.tradeoffSummary, /reduz o custo/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 100_000, totalOutputs: 80_000 }),
      strategy("b", { adjustedToday: 99_500, totalOutputs: 120_000 }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.NEAR_TIE);
  assert.equal(report.outcome.nearTieMetric, "wealth");
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", {
        adjustedToday: 300_000,
        totalOutputs: 10_000,
        lifetimeContractualOutflow: 300_000,
      }),
      strategy("b", {
        adjustedToday: 100_000,
        totalOutputs: 20_000,
        contractualOutflowsLifetime: 200_000,
      }),
    ]),
    fixedMeta,
  );
  const normalizedA = report.facts.strategies.find((item) => item.id === "a");
  assert.equal(report.facts.leaders.cost, "b", "o custo contratual deve prevalecer");
  assert.equal(normalizedA.cost, 300_000);
  assert.equal(normalizedA.costSource, "contractual_lifetime");
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", {
        adjustedToday: 300_000,
        totalOutputs: 100_000,
        lifetimeContractualOutflow: null,
      }),
      strategy("b", {
        adjustedToday: 200_000,
        totalOutputs: 80_000,
        lifetimeContractualOutflow: null,
      }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.INVALID);
  assert.ok(report.facts.errors.some((error) => error.startsWith("missing_cost:")));
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("card", {
        name: "Carta contemplada",
        adjustedToday: 300_000,
        totalOutputs: 100_000,
        payoffMonth: 300,
      }),
      strategy("amort", { adjustedToday: 100_000, totalOutputs: 150_000 }),
    ]),
    fixedMeta,
  );
  const card = report.facts.strategies.find((item) => item.id === "card");
  assert.ok(card.alerts.some((alert) => alert.code === "card_validation"));
  assert.ok(card.alerts.some((alert) => alert.code === "payments_after_horizon"));
  assert.ok(report.viewModel.alerts.some((alert) => alert.code === "card_validation"));
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, totalOutputs: 100_000 }),
      strategy("b", { adjustedToday: 200_000, totalOutputs: 80_000 }),
    ]),
    { ...fixedMeta, priority: "compare" },
  );
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.match(report.viewModel.recommendation.heading, /sem recomendação/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, lifetimeContractualOutflow: 100_000 }),
      strategy("b", { adjustedToday: 300_000, lifetimeContractualOutflow: 80_000 }),
    ]),
    { ...fixedMeta, priority: "wealth" },
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.NEAR_TIE);
  assert.equal(report.outcome.exactTieMetric, "wealth");
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.equal(report.facts.leaders.wealth, null);
  assert.deepEqual(report.facts.coLeaders.wealth, ["a", "b"]);
  assert.match(report.viewModel.recommendation.heading, /empate/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, lifetimeContractualOutflow: 80_000 }),
      strategy("b", { adjustedToday: 200_000, lifetimeContractualOutflow: 80_000 }),
    ]),
    { ...fixedMeta, priority: "cost" },
  );
  assert.equal(report.outcome.exactTieMetric, "cost");
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.equal(report.facts.leaders.cost, null);
}

{
  const input = scenario([
    strategy("a", { adjustedToday: 300_000, lifetimeContractualOutflow: 80_000, payoffMonth: 12 }),
    strategy("b", { adjustedToday: 200_000, lifetimeContractualOutflow: 90_000, payoffMonth: 12 }),
  ]);
  const payoffReport = engine.buildDecisionReport(input, { ...fixedMeta, priority: "payoff" });
  const acquisitionReport = engine.buildDecisionReport(input, { ...fixedMeta, priority: "acquisition" });
  assert.equal(payoffReport.outcome.exactTieMetric, "payoff");
  assert.equal(payoffReport.outcome.recommendedStrategyId, null);
  assert.equal(acquisitionReport.outcome.exactTieMetric, "acquisition");
  assert.equal(acquisitionReport.outcome.recommendedStrategyId, null);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { payoffMonth: null, acquisitionMonth: null, adjustedToday: 300_000 }),
      strategy("b", { payoffMonth: null, acquisitionMonth: null, adjustedToday: 200_000 }),
    ]),
    { ...fixedMeta, priority: "payoff" },
  );
  assert.equal(report.outcome.metricAvailable, false);
  assert.equal(report.outcome.recommendedStrategyId, null);
  assert.match(report.viewModel.recommendation.heading, /indisponível/i);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("leader", {
        name: "Líder patrimonial",
        adjustedToday: 300_000,
        lifetimeContractualOutflow: 70_000,
        payoffMonth: 60,
      }),
      strategy("fast", {
        name: "Quitação rápida",
        adjustedToday: 200_000,
        lifetimeContractualOutflow: 90_000,
        payoffMonth: 12,
      }),
    ]),
    { ...fixedMeta, priority: "payoff" },
  );
  assert.equal(report.outcome.recommendedStrategyId, "fast");
  assert.match(report.viewModel.tradeoffSummary, /Líder patrimonial lidera/i);
  assert.match(report.viewModel.tradeoffSummary, /Quitação rápida atende/i);
  assert.equal(report.viewModel.tradeoffCards.length, 3);
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("card", {
        name: "Carta contemplada",
        isEnabled: false,
        isViable: false,
        isAffordable: true,
        lifetimeContractualOutflow: 0,
      }),
      strategy("amort", { adjustedToday: 200_000, lifetimeContractualOutflow: 90_000 }),
    ]),
    { ...fixedMeta, priority: "compare" },
  );
  const card = report.facts.strategies.find((item) => item.id === "card");
  const cardChart = report.viewModel.charts.cost.find((item) => item.id === "card");
  assert.deepEqual(card.alerts, []);
  assert.equal(cardChart.formattedValue, "Não considerada");
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { isViable: true, isAffordable: false }),
      strategy("b", { adjustedToday: 200_000, lifetimeContractualOutflow: 90_000 }),
    ]),
    fixedMeta,
  );
  assert.equal(report.outcome.state, engine.OUTCOME_STATES.INVALID);
  assert.ok(report.facts.errors.includes("inconsistent_status:a"));
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("iq", { adjustedToday: 300_000, lifetimeContractualOutflow: 100_000 }),
      strategy("cons", {
        adjustedToday: 200_000,
        lifetimeContractualOutflow: 90_000,
        acquisitionMonth: 24,
      }),
    ]),
    { ...fixedMeta, priority: "compare" },
  );
  assert.ok(report.viewModel.alerts.length > 0);
  assert.ok(report.viewModel.alerts.every((alert) => alert.text.includes(":")));
}

{
  const report = engine.buildDecisionReport(
    scenario([
      strategy("a", { adjustedToday: 300_000, totalOutputs: 100_000 }),
      strategy("b", { adjustedToday: 200_000, totalOutputs: 80_000 }),
    ]),
    { ...fixedMeta, priority: "not-a-priority" },
  );
  assert.equal(report.meta.priority, engine.DEFAULT_PRIORITY);
}

{
  const input = scenario([
    strategy("a", { name: "Opção <segura>", adjustedToday: 300_000, totalOutputs: 100_000 }),
    strategy("b", { adjustedToday: 200_000, totalOutputs: 80_000 }),
  ]);
  const report = engine.buildDecisionReport(input, {
    ...fixedMeta,
    clientName: "Ana & Bruno",
  });
  const html = engine.renderReportHtml(report);

  assert.equal((html.match(/<section class="report-page(?:\s|")/g) || []).length, 4, "o relatório deve ter quatro páginas");
  assert.match(html, /Menor custo/);
  assert.doesNotMatch(html, /menor desembolso/i);
  assert.match(html, /Ana &amp; Bruno/);
  assert.match(html, /Opção &lt;segura&gt;/);
  assert.match(html, /data-report-version="1\.0\.0"/);
  assert.match(html, /22 de julho de 2026/);
  assert.match(html, /Imóvel/);
  assert.match(html, /rentabilidade líquida de 8,0% a\.a\./);
  assert.match(html, /inflação de 4,5% a\.a\./);
}

{
  const input = scenario([
    strategy("a", { adjustedToday: 300_000, totalOutputs: 100_000 }),
    strategy("b", { adjustedToday: 200_000, totalOutputs: 80_000 }),
  ]);
  const first = engine.buildDecisionReport(input, fixedMeta);
  const second = engine.buildDecisionReport(input, fixedMeta);
  assert.deepEqual(second, first, "o mesmo cenário e a mesma meta devem gerar o mesmo relatório");
}

{
  const realResult = model.calculateScenario(model.defaultInputs);
  const report = engine.buildDecisionReport(realResult, fixedMeta);
  assert.equal(report.facts.strategies.length, realResult.strategies.length);
  assert.equal(report.facts.viableCount, realResult.viableStrategies.length);
  assert.ok(report.facts.leaders.wealth, "o cenário real deve ter líder patrimonial");
  assert.ok(report.facts.leaders.cost, "o cenário real deve ter líder de custo");
  assert.equal((engine.renderReportHtml(report).match(/<section class="report-page(?:\s|")/g) || []).length, 4);
}

{
  const indexHtml = fs.readFileSync(require.resolve("./index.html"), "utf8");
  const appSource = fs.readFileSync(require.resolve("./app.js"), "utf8");
  const styles = fs.readFileSync(require.resolve("./styles.css"), "utf8");
  const reportPrintStyles = fs.readFileSync(require.resolve("./report-print.css"), "utf8");
  assert.match(indexHtml, /value="wealth"\s+checked/);
  assert.match(indexHtml, /<span>Gerar relatório<\/span>/);
  assert.ok(indexHtml.indexOf("report-engine.js") < indexHtml.indexOf("app.js"));
  assert.match(indexHtml, /report-print\.css/);
  assert.doesNotMatch(indexHtml, /https?:\/\//, "a ferramenta distribuída não deve depender de CDN");
  assert.match(appSource, /engine\.buildDecisionReport\(result, reportContext\)/);
  assert.match(appSource, /waitForReportLayout\(\(\) => window\.print\(\)\)/);
  assert.match(appSource, /document\.fonts\.ready/);
  assert.doesNotMatch(appSource, /function renderPrintReport/);
  assert.match(styles, /SpaceGrotesk-Medium\.woff2/);
  assert.match(styles, /SpaceGrotesk-Bold\.woff2/);
  assert.match(reportPrintStyles, /\.cover-page/);
  assert.match(reportPrintStyles, /\.analysis-verdict/);
  assert.match(reportPrintStyles, /\.planner-reco-header/);
}

console.log("Report engine: regras determinísticas, conteúdo e integração com calculateScenario validados.");
