const MAX_MONTHS = 480;

const defaultInputs = {
  usage: "Investimento para aluguel",
  propertyValue: 200000,
  reserve: 90000,
  monthlyCash: 7000,
  horizonMonths: 420,
  entryPct: 0.2,
  finMonths: 420,
  annualInterest: 0.12,
  annualTR: 0.02,
  insurance: 150,
  consAdminPct: 0.24,
  consReservePct: 0.02,
  consInsurancePct: 0.00035,
  consMonths: 240,
  consAdjustPct: 0.06,
  iqMonth: 60,
  directMonth: 60,
  useEmbeddedBid: true,
  embeddedBidPct: 0.2,
  useFreeBid: false,
  freeBidPct: 0,
  useFixedBid: true,
  fixedBidPct: 0.2,
  availabilityMonth: 1,
  rentPaidPct: 0.0035,
  rentReceivedPct: 0.004,
  rentAdjustPct: 0.06,
  rentalDeductionsPct: 0.2,
  annualReturn: 0.12,
  appreciationPct: 0.06,
  inflationPct: 0.06,
  acquisitionFinMonth: 1,
};

const percentFields = new Set([
  "entryPct",
  "annualInterest",
  "annualTR",
  "consAdminPct",
  "consReservePct",
  "consInsurancePct",
  "consAdjustPct",
  "embeddedBidPct",
  "freeBidPct",
  "fixedBidPct",
  "rentPaidPct",
  "rentReceivedPct",
  "rentAdjustPct",
  "rentalDeductionsPct",
  "annualReturn",
  "appreciationPct",
  "inflationPct",
]);

const integerFields = new Set([
  "horizonMonths",
  "finMonths",
  "consMonths",
  "iqMonth",
  "directMonth",
  "availabilityMonth",
]);

const booleanFields = new Set(["useEmbeddedBid", "useFreeBid", "useFixedBid"]);

const currencyFields = new Set(["propertyValue", "reserve", "monthlyCash", "insurance"]);

const strategyColors = {
  fin: "#111111",
  amort: "#115e59",
  iq: "#244d62",
  cons: "#8f3d49",
};

const strategyNames = {
  fin: "Financiamento sem amortização",
  amort: "Financiamento com amortização",
  iq: "Financiamento + consórcio IQ",
  cons: "Consórcio",
};

const shortNames = {
  fin: "Fin. sem amort.",
  amort: "Fin. amort.",
  iq: "Fin. + IQ",
  cons: "Consórcio",
};

const observations = {
  fin: "Compra por financiamento; aluguel depende do uso e das chaves.",
  amort: "Pagamento extra usa a sobra entre caixa livre mensal e prestação SAC do mês.",
  iq: "Financiamento compra no início; carta quita no mês de contemplação.",
  cons: "Sem entrada inicial; compra e uso/aluguel só após contemplação e chaves.",
};

function monthlyRate(annualRate) {
  return Math.pow(1 + annualRate, 1 / 12) - 1;
}

function annualStepFactor(annualRate, month) {
  return Math.pow(1 + annualRate, Math.floor((month - 1) / 12));
}

function clampNumber(value, min, max) {
  const number = Number.isFinite(value) ? value : min;
  return Math.min(max, Math.max(min, number));
}

function normalizeInputs(raw) {
  const p = { ...defaultInputs, ...raw };

  p.propertyValue = Math.max(0, Number(p.propertyValue) || 0);
  p.reserve = Math.max(0, Number(p.reserve) || 0);
  p.monthlyCash = Math.max(0, Number(p.monthlyCash) || 0);
  p.insurance = Math.max(0, Number(p.insurance) || 0);

  for (const key of percentFields) {
    p[key] = Math.max(0, Number(p[key]) || 0);
  }

  for (const key of integerFields) {
    p[key] = Math.round(clampNumber(Number(p[key]) || defaultInputs[key], 1, MAX_MONTHS));
  }

  p.horizonMonths = Math.round(clampNumber(p.horizonMonths, 12, MAX_MONTHS));
  p.finMonths = Math.round(clampNumber(p.finMonths, 1, MAX_MONTHS));
  p.consMonths = Math.round(clampNumber(p.consMonths, 1, MAX_MONTHS));
  p.iqMonth = Math.round(clampNumber(p.iqMonth, 1, Math.min(MAX_MONTHS, p.consMonths)));
  p.directMonth = Math.round(clampNumber(p.directMonth, 1, Math.min(MAX_MONTHS, p.consMonths)));
  p.availabilityMonth = Math.round(clampNumber(p.availabilityMonth, 1, MAX_MONTHS));
  p.acquisitionFinMonth = 1;

  p.entryPct = clampNumber(p.entryPct, 0, 1);
  p.rentalDeductionsPct = clampNumber(p.rentalDeductionsPct, 0, 1);
  p.embeddedBidPct = clampNumber(p.embeddedBidPct, 0, 0.9);
  p.freeBidPct = clampNumber(p.freeBidPct, 0, 0.9);
  p.fixedBidPct = clampNumber(p.fixedBidPct, 0, 0.9);
  p.usage = p.usage === "Moradia propria" ? "Moradia propria" : "Investimento para aluguel";

  for (const key of booleanFields) {
    p[key] = Boolean(p[key]);
  }

  return p;
}

function buildFinancingSchedule(p) {
  const interestMonthly = monthlyRate(p.annualInterest);
  const trMonthly = monthlyRate(p.annualTR);
  const entry = p.propertyValue * p.entryPct;
  const financed = Math.max(0, p.propertyValue - entry);
  const rows = [];
  let balance = financed;

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const active = month <= p.finMonths && balance > 0.005;
    if (!active) {
      rows.push({
        month,
        initial: 0,
        correction: 0,
        corrected: 0,
        amortization: 0,
        interest: 0,
        insurance: 0,
        payment: 0,
        balance: 0,
      });
      balance = 0;
      continue;
    }

    const correction = balance * trMonthly;
    const corrected = balance + correction;
    const remaining = Math.max(1, p.finMonths - month + 1);
    const amortization = Math.min(corrected, corrected / remaining);
    const interest = corrected * interestMonthly;
    const insurance = p.insurance;
    const payment = amortization + interest + insurance;
    const finalBalance = Math.max(0, corrected - amortization);

    rows.push({
      month,
      initial: balance,
      correction,
      corrected,
      amortization,
      interest,
      insurance,
      payment,
      balance: finalBalance,
    });

    balance = finalBalance;
  }

  return rows;
}

function buildAmortizationSchedule(p, financingRows) {
  const interestMonthly = monthlyRate(p.annualInterest);
  const trMonthly = monthlyRate(p.annualTR);
  const entry = p.propertyValue * p.entryPct;
  const financed = Math.max(0, p.propertyValue - entry);
  const rows = [];
  let balance = financed;
  let payoffMonth = null;

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const active = month <= p.finMonths && balance > 0.005;
    if (!active) {
      rows.push({
        month,
        ordinaryPayment: 0,
        extraPayment: 0,
        payment: 0,
        balance: 0,
      });
      balance = 0;
      continue;
    }

    const correction = balance * trMonthly;
    const corrected = balance + correction;
    const ordinaryAmort = Math.min(corrected, financingRows[month - 1]?.amortization || corrected);
    const interest = corrected * interestMonthly;
    const insurance = p.insurance;
    const ordinaryPayment = ordinaryAmort + interest + insurance;
    const roomForExtra = Math.max(0, p.monthlyCash - ordinaryPayment);
    const extraPayment = Math.min(Math.max(0, corrected - ordinaryAmort), roomForExtra);
    const payment = ordinaryPayment + extraPayment;
    const finalBalance = Math.max(0, corrected - ordinaryAmort - extraPayment);

    if (!payoffMonth && finalBalance <= 0.01) {
      payoffMonth = month;
    }

    rows.push({
      month,
      ordinaryPayment,
      extraPayment,
      payment,
      balance: finalBalance,
    });

    balance = finalBalance;
  }

  return { rows, payoffMonth: payoffMonth || p.finMonths };
}

function buildIqSchedule(p, financingRows) {
  const totalFees = p.consAdminPct + p.consReservePct + p.consInsurancePct;
  const factorAtContemplation = annualStepFactor(p.consAdjustPct, p.iqMonth);
  const balanceAtContemplation = financingRows[p.iqMonth - 1]?.balance || 0;
  const letterInitial = factorAtContemplation > 0 ? balanceAtContemplation / factorAtContemplation : 0;
  const creditUpdated = letterInitial * factorAtContemplation;
  const installmentInitial = p.consMonths > 0 ? (letterInitial * (1 + totalFees)) / p.consMonths : 0;
  const rows = [];

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const factor = annualStepFactor(p.consAdjustPct, month);
    const consInstallment = month <= p.consMonths ? installmentInitial * factor : 0;
    const finPayment = month <= p.iqMonth && month <= p.finMonths ? financingRows[month - 1]?.payment || 0 : 0;
    const payment = consInstallment + finPayment;
    const debt = month < p.iqMonth && month <= p.finMonths ? financingRows[month - 1]?.balance || 0 : 0;

    rows.push({
      month,
      consInstallment,
      finPayment,
      payment,
      debt,
    });
  }

  return {
    rows,
    letterInitial,
    creditUpdated,
    installmentInitial,
    balanceAtContemplation,
  };
}

function buildDirectConsortiumSchedule(p) {
  const totalFees = p.consAdminPct + p.consReservePct + p.consInsurancePct;
  const embedded = p.useEmbeddedBid ? p.embeddedBidPct : 0;
  const freeFixed = (p.useFreeBid ? p.freeBidPct : 0) + (p.useFixedBid ? p.fixedBidPct : 0);
  const clientBidPct = Math.max(0, freeFixed - embedded);
  const factorAtContemplation = annualStepFactor(p.consAdjustPct, p.directMonth);
  const denominator = Math.max(0.01, factorAtContemplation * (1 - embedded));
  const letterInitial = p.propertyValue / denominator;
  const grossCredit = letterInitial * factorAtContemplation;
  const netCredit = grossCredit * (1 - embedded);
  const clientBid = grossCredit * clientBidPct;
  const installmentInitial = p.consMonths > 0 ? (letterInitial * (1 + totalFees)) / p.consMonths : 0;
  const rows = [];

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const factor = annualStepFactor(p.consAdjustPct, month);
    const installment = month <= p.consMonths ? installmentInitial * factor : 0;
    const bid = month === p.directMonth ? clientBid : 0;

    rows.push({
      month,
      installment,
      bid,
      payment: installment,
      debt: 0,
    });
  }

  return {
    rows,
    embedded,
    freeFixed,
    clientBidPct,
    letterInitial,
    grossCredit,
    netCredit,
    clientBid,
    installmentInitial,
  };
}

function buildSharedRows(p) {
  const returnMonthly = monthlyRate(p.annualReturn);
  const rows = [];
  let baseFv = 0;

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const cashFree = month <= p.horizonMonths ? p.monthlyCash : 0;
    const rentStep = annualStepFactor(p.rentAdjustPct, month);
    const rentPaidBase = p.propertyValue * p.rentPaidPct * rentStep;
    const rentReceivedBase = p.propertyValue * p.rentReceivedPct * (1 - p.rentalDeductionsPct) * rentStep;
    const propertyValueMonth = p.propertyValue * Math.pow(1 + p.appreciationPct, month / 12);
    baseFv = month === 1 ? p.reserve + cashFree : baseFv * (1 + returnMonthly) + cashFree;

    rows.push({
      month,
      cashFree,
      rentPaidBase,
      rentReceivedBase,
      propertyValueMonth,
      baseFv,
    });
  }

  return rows;
}

function buildStrategyRows(p, sharedRows, schedules, strategyId) {
  const returnMonthly = monthlyRate(p.annualReturn);
  const rows = [];
  let invested = 0;
  let capitalVf = 0;
  const entry = p.propertyValue * p.entryPct;
  const finAvailability = Math.max(p.acquisitionFinMonth, p.availabilityMonth);
  const directAvailability = Math.max(p.directMonth, p.availabilityMonth);

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const shared = sharedRows[month - 1];
    let entryOrBid = 0;
    let payment = 0;
    let debt = 0;
    let acquiredMonth = p.acquisitionFinMonth;
    let availability = finAvailability;

    if (strategyId === "fin") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = schedules.financing[month - 1]?.payment || 0;
      debt = month <= p.finMonths ? schedules.financing[month - 1]?.balance || 0 : 0;
    }

    if (strategyId === "amort") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = schedules.amortization.rows[month - 1]?.payment || 0;
      debt = month <= p.finMonths ? schedules.amortization.rows[month - 1]?.balance || 0 : 0;
    }

    if (strategyId === "iq") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = schedules.iq.rows[month - 1]?.payment || 0;
      debt = schedules.iq.rows[month - 1]?.debt || 0;
    }

    if (strategyId === "cons") {
      entryOrBid = month === p.directMonth ? schedules.direct.clientBid : 0;
      payment = schedules.direct.rows[month - 1]?.payment || 0;
      debt = 0;
      acquiredMonth = p.directMonth;
      availability = directAvailability;
    }

    const rentPaid =
      p.usage === "Moradia propria" && month < availability ? shared.rentPaidBase : 0;
    const rentReceived =
      p.usage === "Investimento para aluguel" && month >= availability ? shared.rentReceivedBase : 0;
    const outputs = entryOrBid + payment + rentPaid;
    const inputs = rentReceived;
    const freeFlow = shared.cashFree + inputs - outputs;
    const beforeAdjustment = month === 1 ? p.reserve + freeFlow : invested * (1 + returnMonthly) + freeFlow;
    const capitalMonth = Math.max(0, -beforeAdjustment);
    capitalVf = month === 1 ? capitalMonth : capitalVf * (1 + returnMonthly) + capitalMonth;
    invested = Math.max(0, beforeAdjustment);

    const property = month >= acquiredMonth ? shared.propertyValueMonth : 0;
    const grossWorth = invested + property - debt;
    const adjustedWorth = grossWorth - capitalVf;

    rows.push({
      month,
      entryOrBid,
      payment,
      rentPaid,
      rentReceived,
      outputs,
      inputs,
      freeFlow,
      invested,
      capitalMonth,
      capitalVf,
      debt,
      property,
      grossWorth,
      adjustedWorth,
    });
  }

  return rows;
}

function summarizeStrategy(p, sharedRows, strategyRows, schedules, id) {
  const horizon = p.horizonMonths;
  const final = strategyRows[horizon - 1];
  const totalOutputs = strategyRows.slice(0, horizon).reduce((sum, row) => sum + row.outputs, 0);
  const totalInputs = strategyRows.slice(0, horizon).reduce((sum, row) => sum + row.inputs, 0);
  const inflationFactor = Math.pow(1 + p.inflationPct, horizon / 12);
  const baseFv = sharedRows[horizon - 1].baseFv;
  const gapFv = final.adjustedWorth - baseFv;
  const entry = p.propertyValue * p.entryPct;

  const acquisitionMonth = id === "cons" ? p.directMonth : p.acquisitionFinMonth;
  const payoffMonth =
    id === "fin"
      ? p.finMonths
      : id === "amort"
        ? schedules.amortization.payoffMonth
        : id === "iq"
          ? p.iqMonth
          : p.directMonth;

  const entryOrInitial =
    id === "cons" ? schedules.direct.clientBid : entry;

  return {
    id,
    name: strategyNames[id],
    short: shortNames[id],
    observation: observations[id],
    acquisitionMonth,
    payoffMonth,
    entryOrInitial,
    totalOutputs,
    totalInputs,
    netFlow: totalInputs - totalOutputs,
    capitalVf: final.capitalVf,
    grossWorth: final.grossWorth,
    grossToday: final.grossWorth / inflationFactor,
    adjustedWorth: final.adjustedWorth,
    adjustedToday: final.adjustedWorth / inflationFactor,
    gapFv,
    invested: final.invested,
    debt: final.debt,
    rows: strategyRows,
  };
}

function calculateScenario(rawInputs) {
  const p = normalizeInputs(rawInputs);
  const financing = buildFinancingSchedule(p);
  const amortization = buildAmortizationSchedule(p, financing);
  const iq = buildIqSchedule(p, financing);
  const direct = buildDirectConsortiumSchedule(p);
  const sharedRows = buildSharedRows(p);
  const schedules = { financing, amortization, iq, direct };

  const strategies = ["fin", "amort", "iq", "cons"].map((id) => {
    const rows = buildStrategyRows(p, sharedRows, schedules, id);
    return summarizeStrategy(p, sharedRows, rows, schedules, id);
  });

  const sorted = [...strategies].sort((a, b) => b.adjustedToday - a.adjustedToday);
  const best = sorted[0];
  const second = sorted[1];
  const entry = p.propertyValue * p.entryPct;
  const totalFees = p.consAdminPct + p.consReservePct + p.consInsurancePct;

  return {
    inputs: p,
    sharedRows,
    schedules,
    strategies,
    best,
    second,
    derived: {
      entry,
      financed: Math.max(0, p.propertyValue - entry),
      monthlyInterest: monthlyRate(p.annualInterest),
      monthlyTR: monthlyRate(p.annualTR),
      monthlyReturn: monthlyRate(p.annualReturn),
      totalConsortiumFees: totalFees,
      rentPaidInitial: p.propertyValue * p.rentPaidPct,
      rentReceivedInitial: p.propertyValue * p.rentReceivedPct * (1 - p.rentalDeductionsPct),
      baseFv: sharedRows[p.horizonMonths - 1].baseFv,
    },
  };
}

function formatCurrency(value, options = {}) {
  const digits = options.digits ?? 0;
  const notation = options.compact ? "compact" : "standard";
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: options.minimumDigits ?? digits,
    maximumFractionDigits: digits,
    notation,
  }).format(Number.isFinite(value) ? value : 0).replace(/\u00a0/g, " ");
}

function formatPatrimony(value) {
  return formatCurrency(value, { compact: true, digits: 2 });
}

function formatPercent(value, digits = 1) {
  return new Intl.NumberFormat("pt-BR", {
    style: "percent",
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(Number.isFinite(value) ? value : 0);
}

function formatMonth(month) {
  return `mês ${Math.round(month || 0)}`;
}

function formatSignedCurrency(value) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${formatCurrency(value)}`;
}

function parseCurrencyValue(value) {
  const text = String(value ?? "").trim();
  if (!text) return 0;

  const normalized = text.replace(/[^\d,.-]/g, "");
  const isNegative = normalized.includes("-");
  const unsigned = normalized.replace(/-/g, "");
  const lastComma = unsigned.lastIndexOf(",");
  const lastDot = unsigned.lastIndexOf(".");
  let integerPart = unsigned;
  let decimalPart = "";

  if (lastComma >= 0) {
    integerPart = unsigned.slice(0, lastComma);
    decimalPart = unsigned.slice(lastComma + 1);
  } else if (lastDot >= 0 && unsigned.indexOf(".") === lastDot) {
    const possibleDecimals = unsigned.slice(lastDot + 1);
    if (possibleDecimals.length > 0 && possibleDecimals.length <= 2) {
      integerPart = unsigned.slice(0, lastDot);
      decimalPart = possibleDecimals;
    }
  }

  const integerDigits = integerPart.replace(/[^\d]/g, "") || "0";
  const decimalDigits = decimalPart.replace(/[^\d]/g, "").slice(0, 2);
  const parsed = Number(`${isNegative ? "-" : ""}${integerDigits}${decimalDigits ? `.${decimalDigits}` : ""}`);
  return Number.isFinite(parsed) ? parsed : 0;
}

function inputValueForDisplay(key, value) {
  if (currencyFields.has(key)) return formatCurrency(value, { digits: 2 });
  if (percentFields.has(key)) return Number((value * 100).toFixed(4));
  return value;
}

function parseInputValue(key, element) {
  if (booleanFields.has(key)) return element.checked;
  if (currencyFields.has(key)) return parseCurrencyValue(element.value);
  const raw = Number(element.value);
  const value = Number.isFinite(raw) ? raw : 0;
  return percentFields.has(key) ? value / 100 : value;
}

function getDomState() {
  const state = { ...defaultInputs };
  document.querySelectorAll("[data-field]").forEach((element) => {
    const key = element.dataset.field;
    state[key] = parseInputValue(key, element);
  });
  const activeUsage = document.querySelector('[data-segment="usage"].is-active');
  if (activeUsage) state.usage = activeUsage.dataset.value;
  return state;
}

function setDomState(state) {
  const normalized = normalizeInputs(state);
  document.querySelectorAll("[data-field]").forEach((element) => {
    const key = element.dataset.field;
    if (booleanFields.has(key)) {
      element.checked = Boolean(normalized[key]);
    } else {
      element.value = inputValueForDisplay(key, normalized[key]);
    }
  });

  document.querySelectorAll('[data-segment="usage"]').forEach((button) => {
    button.classList.toggle("is-active", button.dataset.value === normalized.usage);
  });
}

function renderHero(result) {
  const best = result.best;
  document.getElementById("bestStrategyName").textContent = best.short;
  document.getElementById("bestAdjustedToday").textContent = formatPatrimony(best.adjustedToday);
  document.getElementById("bestNominalCost").textContent = formatCurrency(best.totalOutputs);

  const gapToSecond = best.adjustedToday - result.second.adjustedToday;

  const insight = `${best.short} lidera por ${formatCurrency(gapToSecond)} em Patrimônio Acumulado contra a segunda alternativa. O custo estimado dessa estratégia é ${formatCurrency(best.totalOutputs)} ao longo do horizonte analisado.`;
  document.getElementById("executiveInsight").textContent = insight;
}

function renderStrategyCards(result) {
  const container = document.getElementById("strategyCards");
  container.innerHTML = "";

  result.strategies.forEach((strategy, index) => {
    const card = document.createElement("article");
    card.className = `strategy-card${strategy.id === result.best.id ? " is-best" : ""}`;

    card.innerHTML = `
      <div class="strategy-title">
        <span class="badge ${strategy.id === result.best.id ? "best" : ""}">
          ${strategy.id === result.best.id ? "Indicada" : `Cenário ${index + 1}`}
        </span>
        <h3>${strategy.name}</h3>
      </div>
      <div class="metric-list">
        <div class="metric">
          <span>Patrimônio Acumulado</span>
          <strong>${formatPatrimony(strategy.adjustedToday)}</strong>
        </div>
        <div class="metric">
          <span>Custo</span>
          <strong>${formatCurrency(strategy.totalOutputs)}</strong>
        </div>
        <div class="metric">
          <span>Entrada ou lance</span>
          <strong>${formatCurrency(strategy.entryOrInitial)}</strong>
        </div>
        <div class="metric">
          <span>Aquisição</span>
          <strong>${formatMonth(strategy.acquisitionMonth)}</strong>
        </div>
      </div>
    `;

    container.appendChild(card);
  });
}

function renderBarChart(result) {
  const container = document.getElementById("barChart");
  container.innerHTML = "";
  const values = result.strategies.map((strategy) => strategy.adjustedToday);
  const maxAbs = Math.max(1, ...values.map((value) => Math.abs(value)));

  result.strategies.forEach((strategy) => {
    const row = document.createElement("div");
    row.className = "bar-row";
    const width = Math.max(2, Math.abs(strategy.adjustedToday) / maxAbs * 100);
    const isBest = strategy.id === result.best.id;
    const isNegative = strategy.adjustedToday < 0;

    row.innerHTML = `
      <div class="bar-label">${strategy.short}</div>
      <div class="bar-track">
        <div
          class="bar-fill ${isBest ? "is-best" : ""} ${isNegative ? "is-negative" : ""}"
          style="width: ${width}%"
        ></div>
      </div>
      <div class="bar-value">${formatPatrimony(strategy.adjustedToday)}</div>
    `;
    container.appendChild(row);
  });
}

function svgEl(name, attrs = {}) {
  const element = document.createElementNS("http://www.w3.org/2000/svg", name);
  Object.entries(attrs).forEach(([key, value]) => element.setAttribute(key, value));
  return element;
}

function renderLineChart(result) {
  const svg = document.getElementById("lineChart");
  svg.innerHTML = "";
  const width = 760;
  const height = 300;
  const padding = { top: 22, right: 22, bottom: 34, left: 54 };
  const horizon = result.inputs.horizonMonths;
  const sampleStep = Math.max(1, Math.floor(horizon / 72));
  const allPoints = [];

  result.strategies.forEach((strategy) => {
    for (let index = 0; index < horizon; index += sampleStep) {
      allPoints.push(strategy.rows[index].adjustedWorth);
    }
    allPoints.push(strategy.rows[horizon - 1].adjustedWorth);
  });

  let min = Math.min(...allPoints);
  let max = Math.max(...allPoints);
  if (min === max) {
    min -= 1;
    max += 1;
  }

  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;
  const xScale = (month) => padding.left + ((month - 1) / Math.max(1, horizon - 1)) * plotWidth;
  const yScale = (value) => padding.top + (1 - (value - min) / (max - min)) * plotHeight;

  for (let i = 0; i <= 4; i += 1) {
    const y = padding.top + (plotHeight / 4) * i;
    svg.appendChild(svgEl("line", {
      x1: padding.left,
      x2: width - padding.right,
      y1: y,
      y2: y,
      class: "chart-grid",
    }));
  }

  svg.appendChild(svgEl("line", {
    x1: padding.left,
    x2: padding.left,
    y1: padding.top,
    y2: height - padding.bottom,
    class: "chart-axis",
  }));
  svg.appendChild(svgEl("line", {
    x1: padding.left,
    x2: width - padding.right,
    y1: height - padding.bottom,
    y2: height - padding.bottom,
    class: "chart-axis",
  }));

  const yTop = svgEl("text", {
    x: 6,
    y: padding.top + 4,
    fill: "#77746b",
    "font-size": "12",
  });
  yTop.textContent = formatPatrimony(max);
  svg.appendChild(yTop);

  const yBottom = svgEl("text", {
    x: 6,
    y: height - padding.bottom,
    fill: "#77746b",
    "font-size": "12",
  });
  yBottom.textContent = formatPatrimony(min);
  svg.appendChild(yBottom);

  const xStart = svgEl("text", {
    x: padding.left,
    y: height - 10,
    fill: "#77746b",
    "font-size": "12",
  });
  xStart.textContent = "mês 1";
  svg.appendChild(xStart);

  const xEnd = svgEl("text", {
    x: width - padding.right - 54,
    y: height - 10,
    fill: "#77746b",
    "font-size": "12",
  });
  xEnd.textContent = `mês ${horizon}`;
  svg.appendChild(xEnd);

  result.strategies.forEach((strategy) => {
    const points = [];
    for (let index = 0; index < horizon; index += sampleStep) {
      const month = index + 1;
      points.push(`${xScale(month).toFixed(2)},${yScale(strategy.rows[index].adjustedWorth).toFixed(2)}`);
    }
    points.push(`${xScale(horizon).toFixed(2)},${yScale(strategy.rows[horizon - 1].adjustedWorth).toFixed(2)}`);

    svg.appendChild(svgEl("polyline", {
      points: points.join(" "),
      fill: "none",
      stroke: strategyColors[strategy.id],
      "stroke-width": strategy.id === result.best.id ? "4" : "2.5",
      "stroke-linejoin": "round",
      "stroke-linecap": "round",
    }));
  });

  const legend = document.getElementById("lineLegend");
  legend.innerHTML = "";
  result.strategies.forEach((strategy) => {
    const item = document.createElement("span");
    item.className = "legend-item";
    item.innerHTML = `
      <span class="legend-swatch" style="background: ${strategyColors[strategy.id]}"></span>
      ${strategy.short}
    `;
    legend.appendChild(item);
  });
}

function renderComparisonRows(result) {
  const tbody = document.getElementById("comparisonRows");
  tbody.innerHTML = "";

  result.strategies.forEach((strategy) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><strong>${strategy.name}</strong><br><span class="muted-text">${strategy.observation}</span></td>
      <td>${formatCurrency(strategy.entryOrInitial)}</td>
      <td>${formatMonth(strategy.acquisitionMonth)}</td>
      <td>${formatCurrency(strategy.totalOutputs)}</td>
      <td>${formatCurrency(strategy.adjustedToday)}</td>
    `;
    tbody.appendChild(row);
  });
}

function renderAssumptions(result) {
  const p = result.inputs;
  const items = [
    ["Imóvel", formatCurrency(p.propertyValue)],
    ["Reserva", formatCurrency(p.reserve)],
    ["Caixa mensal", formatCurrency(p.monthlyCash)],
    ["Horizonte", `${p.horizonMonths} meses`],
    ["Financiamento", `${p.finMonths} meses a ${formatPercent(p.annualInterest)}`],
    ["Consórcio", `${p.consMonths} meses, reajuste ${formatPercent(p.consAdjustPct)}`],
    ["Uso", p.usage === "Moradia propria" ? "Moradia" : "Investimento"],
    ["Rentabilidade", formatPercent(p.annualReturn)],
  ];

  const container = document.getElementById("assumptionChips");
  container.innerHTML = "";
  items.forEach(([label, value]) => {
    const chip = document.createElement("span");
    chip.className = "assumption-chip";
    chip.innerHTML = `${label}<strong>${value}</strong>`;
    container.appendChild(chip);
  });
}

function renderProposal(result) {
  const best = result.best;
  const second = result.second;
  const p = result.inputs;
  const gap = best.adjustedToday - second.adjustedToday;
  const conclusion = `${best.name} aparece como a alternativa mais eficiente no cenário informado, com ${formatCurrency(best.adjustedToday)} de Patrimônio Acumulado e vantagem de ${formatCurrency(gap)} sobre ${second.short}. A decisão deve considerar capacidade de caixa, risco de contemplação, regras do grupo e validação de crédito.`;
  document.getElementById("proposalConclusion").textContent = conclusion;

  const assumptions = [
    `Valor do imóvel: ${formatCurrency(p.propertyValue)}.`,
    `Entrada do financiamento: ${formatPercent(p.entryPct)} (${formatCurrency(result.derived.entry)}).`,
    `Caixa livre mensal considerado: ${formatCurrency(p.monthlyCash)}.`,
    `Horizonte comparativo: ${p.horizonMonths} meses.`,
    `Rentabilidade líquida anual: ${formatPercent(p.annualReturn)}.`,
    `Finalidade do imóvel: ${p.usage === "Moradia propria" ? "moradia própria" : "investimento para aluguel"}.`,
  ];
  const list = document.getElementById("proposalAssumptions");
  list.innerHTML = "";
  assumptions.forEach((text) => {
    const li = document.createElement("li");
    li.textContent = text;
    list.appendChild(li);
  });
}

function render(result) {
  renderHero(result);
  renderStrategyCards(result);
  renderBarChart(result);
  renderLineChart(result);
  renderComparisonRows(result);
  renderAssumptions(result);
  renderProposal(result);
}

function update() {
  const state = getDomState();
  const result = calculateScenario(state);
  render(result);
}

function init() {
  setDomState(defaultInputs);
  render(calculateScenario(defaultInputs));

  document.querySelectorAll("[data-field]").forEach((element) => {
    element.addEventListener("input", update);
    element.addEventListener("change", update);
    if (currencyFields.has(element.dataset.field)) {
      element.addEventListener("blur", () => {
        const key = element.dataset.field;
        element.value = inputValueForDisplay(key, parseInputValue(key, element));
        update();
      });
    }
  });

  document.querySelectorAll('[data-segment="usage"]').forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll('[data-segment="usage"]').forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      update();
    });
  });

  document.getElementById("resetButton").addEventListener("click", () => {
    setDomState(defaultInputs);
    render(calculateScenario(defaultInputs));
  });

  document.getElementById("printButton").addEventListener("click", () => window.print());

  if (window.lucide) {
    window.lucide.createIcons();
  } else {
    window.addEventListener("load", () => window.lucide?.createIcons());
  }
}

if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", init);
}

if (typeof module !== "undefined") {
  module.exports = {
    defaultInputs,
    normalizeInputs,
    calculateScenario,
    buildFinancingSchedule,
    buildAmortizationSchedule,
    buildIqSchedule,
    buildDirectConsortiumSchedule,
  };
}
