const assert = require("node:assert/strict");
const model = require("./app.js");

function near(actual, expected, tolerance = 0.01, label = "valor") {
  assert.ok(Number.isFinite(actual), `${label}: valor nao numerico`);
  assert.ok(Math.abs(actual - expected) <= tolerance, `${label}: esperado ${expected}, atual ${actual}`);
}

const preset = model.calculateScenario(model.defaultInputs);
assert.equal(preset.strategies.length, 5, "preset deve conter cinco estrategias");
near(preset.schedules.direct.totalBidPct, 0.25, 1e-9, "preset lance total");
near(preset.schedules.direct.embedded, 0.25, 1e-9, "preset lance embutido");
near(preset.schedules.direct.clientBid, 0, 0.01, "preset desembolso externo");
near(
  preset.schedules.direct.netCredit,
  preset.schedules.direct.acquisitionPrice,
  0.01,
  "preset credito liquido cobre preco futuro",
);

const scenarioResults = [];
for (const assetType of ["Imovel", "Veiculo"]) {
  for (const financingSystem of ["SAC", "PRICE"]) {
    for (const horizonMonths of [60, 120, 240, 420]) {
      const result = model.calculateScenario({
        ...model.defaultInputs,
        assetType,
        financingSystem,
        horizonMonths,
      });
      const financing = result.schedules.financing;
      near(financing[result.inputs.finMonths - 1].balance, 0, 0.01, `${assetType}/${financingSystem}/${horizonMonths}: saldo final`);
      near(result.derived.financingPrice, result.inputs.propertyValue, 0.01, "financiamento compra pelo valor atual");
      assert.ok(result.sharedRows[horizonMonths - 1].propertyValueMonth >= 0, "valor residual negativo");
      assert.ok(result.strategies.every((strategy) => Number.isFinite(strategy.grossWorth)), "patrimonio nao numerico");
      assert.equal(result.inputs.selectedAssetRate, assetType === "Veiculo" ? -0.15 : 0.06);
      if (assetType === "Veiculo") {
        assert.equal(result.inputs.usage, "Aquisicao", "veiculo deve ignorar finalidade de renda");
        for (const strategy of result.strategies) {
          near(strategy.totalInputs, 0, 0.01, `${strategy.id}: veiculo sem renda`);
          near(strategy.rows.slice(0, horizonMonths).reduce((sum, row) => sum + row.rentPaid, 0), 0, 0.01, `${strategy.id}: veiculo sem aluguel`);
        }
      }
      scenarioResults.push({ assetType, financingSystem, horizonMonths });
    }
  }
}

const priceFuture = model.calculateScenario({
  ...model.defaultInputs,
  assetType: "Imovel",
  directMonth: 60,
  cardPurchaseMonth: 60,
  cardUseMonth: 60,
});
near(priceFuture.derived.financingPrice, priceFuture.inputs.propertyValue, 0.01, "financiamento no valor atual");
assert.ok(priceFuture.derived.directAcquisitionPrice > priceFuture.inputs.propertyValue, "direto nao usou preco futuro");
assert.ok(priceFuture.derived.cardAcquisitionPrice > priceFuture.inputs.propertyValue, "carta nao usou preco futuro");

const freeBid = model.calculateScenario({
  ...model.defaultInputs,
  useFixedBid: false,
  useFreeBid: true,
  freeBidPct: 0.35,
  useEmbeddedBid: true,
  embeddedBidPct: 0.1,
});
near(freeBid.schedules.direct.totalBidPct, 0.35, 1e-9, "lance livre total");
near(freeBid.schedules.direct.clientBidPct, 0.25, 1e-9, "lance livre com embutido");

const clampedBid = model.calculateScenario({
  ...model.defaultInputs,
  fixedBidPct: 0.2,
  embeddedBidPct: 0.4,
});
near(clampedBid.inputs.embeddedBidPct, 0.2, 1e-9, "embutido limitado ao total");
near(clampedBid.schedules.direct.clientBid, 0, 0.01, "embutido igual ao total");

const mutuallyExclusive = model.normalizeInputs({
  ...model.defaultInputs,
  useFixedBid: true,
  useFreeBid: true,
});
assert.equal(mutuallyExclusive.useFixedBid, true, "lance fixo deveria prevalecer no conflito");
assert.equal(mutuallyExclusive.useFreeBid, false, "lance livre deveria ser desativado no conflito");

const bid25 = model.calculateScenario({ ...model.defaultInputs, fixedBidPct: 0.25, embeddedBidPct: 0.1 });
const bid35 = model.calculateScenario({ ...model.defaultInputs, fixedBidPct: 0.35, embeddedBidPct: 0.1 });
assert.notEqual(
  bid25.strategies.find((strategy) => strategy.id === "cons").totalOutputs,
  bid35.strategies.find((strategy) => strategy.id === "cons").totalOutputs,
  "alterar lance deve alterar outputs",
);

const surplus = model.calculateScenario({
  ...model.defaultInputs,
  cardGrossCredit: 500000,
  cardRestrictions: 0,
  cardPrice: 45000,
  cardTransferFee: 2000,
});
assert.ok(surplus.schedules.card.ignoredSurplus > 0, "cenario deveria ter credito excedente");
near(surplus.schedules.card.releasedSurplus, 0, 0.01, "excedente nao vira caixa");
near(surplus.schedules.card.initialCost, 47000, 0.01, "excedente nao reduz custo inicial");

const longTailCardInput = {
  ...model.defaultInputs,
  horizonMonths: 480,
  cardEnabled: true,
  cardPurchaseMonth: 480,
  cardUseMonth: 480,
  cardRemainingMonths: 480,
  cardInstallment: 1000,
  annualReturn: 0.12,
  cardAdjustPct: 0.06,
};
const longTailCard = model.calculateScenario(longTailCardInput);
const longTailSchedule = longTailCard.schedules.card;
const longTailStrategy = longTailCard.strategies.find((strategy) => strategy.id === "card");
const longTailLastMonth =
  longTailCardInput.cardPurchaseMonth + longTailCardInput.cardRemainingMonths - 1;
const horizonCardRow = longTailSchedule.rows[longTailCardInput.horizonMonths - 1];
const monthlyReturn = Math.pow(1 + longTailCardInput.annualReturn, 1 / 12) - 1;
let expectedHorizonObligation = 0;
let expectedEconomicCost =
  longTailCardInput.cardPrice +
  longTailCardInput.cardTransferFee -
  longTailCardInput.cardGrossCredit;

for (let month = longTailCardInput.cardPurchaseMonth; month <= longTailLastMonth; month += 1) {
  const adjustmentFactor = Math.pow(
    1 + longTailCardInput.cardAdjustPct,
    Math.floor((month - longTailCardInput.cardPurchaseMonth) / 12),
  );
  const installment = longTailCardInput.cardInstallment * adjustmentFactor;
  expectedEconomicCost += installment / Math.pow(1 + monthlyReturn, month);
  if (month > longTailCardInput.horizonMonths) {
    expectedHorizonObligation += installment / Math.pow(1 + monthlyReturn, month - longTailCardInput.horizonMonths);
  }
}

assert.equal(longTailSchedule.rows.length, longTailLastMonth, "cronograma da carta precisa cobrir todas as parcelas");
assert.equal(horizonCardRow.month, longTailCardInput.horizonMonths, "linha do horizonte incorreta");
assert.ok(horizonCardRow.obligation > 0, "parcelas depois do horizonte devem permanecer como obrigacao");
assert.ok(longTailSchedule.rows[longTailSchedule.rows.length - 1].installment > 0, "ultima parcela da carta ausente");
near(horizonCardRow.obligation, expectedHorizonObligation, 0.01, "obrigacao futura da carta");
near(longTailStrategy.debt, horizonCardRow.obligation, 0.01, "divida da carta no horizonte");
near(longTailSchedule.economicCost, expectedEconomicCost, 0.01, "custo economico com cauda longa");

const priceDefault = model.calculateScenario({ ...model.defaultInputs, financingSystem: "PRICE" });
near(priceDefault.inputs.annualTR, 0, 1e-9, "PRICE com indexador padrao zero");
const priceIndexed = model.calculateScenario({ ...model.defaultInputs, financingSystem: "PRICE", annualTR: 0.05 });
assert.notEqual(
  priceDefault.schedules.financing[11].payment,
  priceIndexed.schedules.financing[11].payment,
  "indexador alterado deve impactar PRICE",
);

console.log(JSON.stringify({
  scenarioCount: scenarioResults.length,
  horizons: [60, 120, 240, 420],
  priceFuture: "OK",
  bids: "OK",
  cardSurplus: "ignored",
  cardLongTail: "OK",
  priceIndexer: "OK",
}));
