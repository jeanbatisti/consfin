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
  let totalUncoveredDeficit = 0;
  let firstDeficitMonth = null;
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
    if (capitalMonth > 0.01 && firstDeficitMonth === null) {
      firstDeficitMonth = month;
    }
    totalUncoveredDeficit += capitalMonth;
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
      firstDeficitMonth,
      totalUncoveredDeficit,
    });
  }

  return rows;
}

function getInitialPaymentLines(id, schedules) {
  if (id === "fin") {
    return [{ label: "Financiamento", value: schedules.financing[0]?.payment || 0 }];
  }

  if (id === "amort") {
    const first = schedules.amortization.rows[0] || {};
    const lines = [{ label: "Financiamento", value: first.ordinaryPayment || 0 }];
    if ((first.extraPayment || 0) > 0.01) {
      lines.push({ label: "Amortização extra", value: first.extraPayment || 0 });
    }
    return lines;
  }

  if (id === "iq") {
    return [
      { label: "Financiamento", value: schedules.financing[0]?.payment || 0 },
      { label: "Consórcio IQ", value: schedules.iq.installmentInitial || 0 },
    ];
  }

  return [{ label: "Consórcio", value: schedules.direct.installmentInitial || 0 }];
}

function summarizeStrategy(p, sharedRows, strategyRows, schedules, id) {
  const horizon = p.horizonMonths;
  const final = strategyRows[horizon - 1];
  const totalOutputs = strategyRows.slice(0, horizon).reduce((sum, row) => sum + row.outputs, 0);
  const totalInputs = strategyRows.slice(0, horizon).reduce((sum, row) => sum + row.inputs, 0);
  const horizonRows = strategyRows.slice(0, horizon);
  const firstDeficitRow = horizonRows.find((row) => row.capitalMonth > 0.01);
  const totalUncoveredDeficit = horizonRows.reduce((sum, row) => sum + row.capitalMonth, 0);
  const inflationFactor = Math.pow(1 + p.inflationPct, horizon / 12);
  const baseFv = sharedRows[horizon - 1].baseFv;
  const gapFv = final.adjustedWorth - baseFv;
  const entry = p.propertyValue * p.entryPct;
  const isViable = final.capitalVf <= 1;

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
    initialPaymentLines: getInitialPaymentLines(id, schedules),
    totalOutputs,
    totalInputs,
    netFlow: totalInputs - totalOutputs,
    capitalVf: final.capitalVf,
    isViable,
    firstDeficitMonth: firstDeficitRow?.month ?? null,
    totalUncoveredDeficit,
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

  const viableStrategies = strategies.filter((strategy) => strategy.isViable);
  const sorted = [...viableStrategies].sort((a, b) => b.adjustedToday - a.adjustedToday);
  const best = sorted[0] || null;
  const second = sorted[1] || null;
  const entry = p.propertyValue * p.entryPct;
  const totalFees = p.consAdminPct + p.consReservePct + p.consInsurancePct;

  return {
    inputs: p,
    sharedRows,
    schedules,
    strategies,
    viableStrategies,
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

function formatStrategyPatrimony(strategy) {
  return strategy.isViable ? formatPatrimony(strategy.adjustedToday) : "Inviável";
}

function formatInitialPaymentLines(strategy) {
  return strategy.initialPaymentLines
    .map((line) => `${line.label}: ${formatCurrency(line.value)}`)
    .join("<br>");
}

function formatCashStatus(strategy) {
  if (strategy.isViable) return "Viável";
  return `Inviável no mês ${strategy.firstDeficitMonth || 1}`;
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
  if (!best) {
    document.getElementById("bestStrategyName").textContent = "Nenhuma viável";
    document.getElementById("bestAdjustedToday").textContent = "Inviável";
    document.getElementById("bestNominalCost").textContent = "-";
    document.getElementById("executiveInsight").textContent =
      "Com o caixa livre mensal e a reserva informados, nenhuma alternativa se sustenta sem déficit não coberto. Aumente a reserva, o caixa mensal ou reduza os desembolsos antes de comparar patrimônio.";
    return;
  }

  document.getElementById("bestStrategyName").textContent = best.short;
  document.getElementById("bestAdjustedToday").textContent = formatPatrimony(best.adjustedToday);
  document.getElementById("bestNominalCost").textContent = formatCurrency(best.totalOutputs);

  const gapToSecond = result.second ? best.adjustedToday - result.second.adjustedToday : 0;

  const insight = result.second
    ? `${best.short} lidera por ${formatCurrency(gapToSecond)} em Patrimônio Acumulado entre as alternativas viáveis. O custo estimado dessa estratégia é ${formatCurrency(best.totalOutputs)} ao longo do horizonte analisado.`
    : `${best.short} é a única alternativa viável com o caixa livre mensal e a reserva informados. O custo estimado dessa estratégia é ${formatCurrency(best.totalOutputs)} ao longo do horizonte analisado.`;
  document.getElementById("executiveInsight").textContent = insight;
}

function renderStrategyCards(result) {
  const container = document.getElementById("strategyCards");
  container.innerHTML = "";

  result.strategies.forEach((strategy, index) => {
    const card = document.createElement("article");
    const isBest = result.best && strategy.id === result.best.id;
    card.className = `strategy-card${isBest ? " is-best" : ""}${strategy.isViable ? "" : " is-infeasible"}`;
    const badgeClass = strategy.isViable ? (isBest ? "best" : "") : "infeasible";
    const badgeText = strategy.isViable ? (isBest ? "Indicada" : `Cenário ${index + 1}`) : "Inviável";

    card.innerHTML = `
      <div class="strategy-title">
        <span class="badge ${badgeClass}">${badgeText}</span>
        <h3>${strategy.name}</h3>
      </div>
      <div class="metric-list">
        <div class="metric">
          <span>Patrimônio Acumulado</span>
          <strong class="${strategy.isViable ? "" : "infeasible"}">${formatStrategyPatrimony(strategy)}</strong>
        </div>
        <div class="metric">
          <span>Parcela inicial</span>
          <strong class="multi-line">${formatInitialPaymentLines(strategy)}</strong>
        </div>
        <div class="metric">
          <span>Entrada ou lance</span>
          <strong>${formatCurrency(strategy.entryOrInitial)}</strong>
        </div>
        <div class="metric">
          <span>Status de caixa</span>
          <strong class="${strategy.isViable ? "" : "infeasible"}">${formatCashStatus(strategy)}</strong>
        </div>
        <div class="metric">
          <span>Déficit não coberto</span>
          <strong>${strategy.isViable ? "-" : formatCurrency(strategy.totalUncoveredDeficit)}</strong>
        </div>
      </div>
    `;

    container.appendChild(card);
  });
}

function renderBarChart(result) {
  const container = document.getElementById("barChart");
  container.innerHTML = "";
  const values = result.strategies.filter((strategy) => strategy.isViable).map((strategy) => strategy.adjustedToday);
  const maxAbs = Math.max(1, ...values.map((value) => Math.abs(value)));

  result.strategies.forEach((strategy) => {
    const row = document.createElement("div");
    row.className = `bar-row${strategy.isViable ? "" : " is-infeasible"}`;
    const width = strategy.isViable ? Math.max(2, Math.abs(strategy.adjustedToday) / maxAbs * 100) : 0;
    const isBest = result.best && strategy.id === result.best.id;
    const isNegative = strategy.adjustedToday < 0;

    row.innerHTML = `
      <div class="bar-label">${strategy.short}</div>
      <div class="bar-track">
        <div
          class="bar-fill ${isBest ? "is-best" : ""} ${isNegative ? "is-negative" : ""} ${strategy.isViable ? "" : "is-infeasible"}"
          style="width: ${width}%"
        ></div>
      </div>
      <div class="bar-value">${formatStrategyPatrimony(strategy)}</div>
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
  const drawableStrategies = result.strategies.filter((strategy) => strategy.isViable);
  const legend = document.getElementById("lineLegend");
  legend.innerHTML = "";

  if (drawableStrategies.length === 0) {
    const notice = svgEl("text", {
      x: 54,
      y: 150,
      fill: "#77746b",
      "font-size": "15",
    });
    notice.textContent = "Nenhuma estratégia viável com o caixa e a reserva informados.";
    svg.appendChild(notice);
    legend.innerHTML = '<span class="legend-item">Ajuste caixa mensal, reserva ou desembolsos.</span>';
    return;
  }

  const allPoints = [];

  drawableStrategies.forEach((strategy) => {
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

  drawableStrategies.forEach((strategy) => {
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
      "stroke-width": result.best && strategy.id === result.best.id ? "4" : "2.5",
      "stroke-linejoin": "round",
      "stroke-linecap": "round",
    }));
  });

  drawableStrategies.forEach((strategy) => {
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
      <td>${formatInitialPaymentLines(strategy)}</td>
      <td>${formatMonth(strategy.acquisitionMonth)}</td>
      <td>${formatCurrency(strategy.totalOutputs)}</td>
      <td>${formatStrategyPatrimony(strategy)}</td>
      <td>${formatCashStatus(strategy)}</td>
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
  const conclusion = best
    ? second
      ? `${best.name} aparece como a alternativa mais eficiente entre as opções viáveis, com ${formatCurrency(best.adjustedToday)} de Patrimônio Acumulado e vantagem de ${formatCurrency(best.adjustedToday - second.adjustedToday)} sobre ${second.short}. A decisão deve considerar capacidade de caixa, risco de contemplação, regras do grupo e validação de crédito.`
      : `${best.name} é a única alternativa viável no cenário informado, com ${formatCurrency(best.adjustedToday)} de Patrimônio Acumulado. As demais exigem déficit não coberto e não devem ser recomendadas com esse caixa/reserva.`
    : "Nenhuma alternativa é viável com o caixa livre mensal e a reserva informados. Como a pessoa não pode ficar negativa, o cenário exige aumentar caixa/reserva, reduzir desembolsos ou rever prazo/valor antes de recomendar uma estratégia.";
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

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function printMetric(label, value) {
  return `
    <div class="print-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${value}</strong>
    </div>
  `;
}

function getAlternativeDescription(result, strategy) {
  const p = result.inputs;
  const financed = result.derived.financed;
  const entry = result.derived.entry;

  if (strategy.id === "fin") {
    return `Financiamento de ${formatCurrency(financed)}, com entrada de ${formatCurrency(entry)}, sem amortização extraordinária até o final do prazo de ${p.finMonths} meses.`;
  }

  if (strategy.id === "amort") {
    return `Financiamento de ${formatCurrency(financed)}, com entrada de ${formatCurrency(entry)} e amortização mensal usando a disponibilidade de caixa prevista até a quitação.`;
  }

  if (strategy.id === "iq") {
    return `Financiamento de ${formatCurrency(financed)}, com entrada de ${formatCurrency(entry)}, combinado com consórcio de aproximadamente ${formatCurrency(result.schedules.iq.creditUpdated)} para atuar como interveniente quitante a partir do mês ${p.iqMonth}. Carta inicial estimada de ${formatCurrency(result.schedules.iq.letterInitial)}, prazo de ${p.consMonths} meses e reajuste anual de ${formatPercent(p.consAdjustPct)}.`;
  }

  const bidParts = [];
  if (p.useEmbeddedBid) bidParts.push(`lance embutido de ${formatPercent(p.embeddedBidPct)}`);
  if (p.useFreeBid) bidParts.push(`lance livre de ${formatPercent(p.freeBidPct)}`);
  if (p.useFixedBid) bidParts.push(`lance fixo de ${formatPercent(p.fixedBidPct)}`);
  const bidText = bidParts.length ? `, com ${bidParts.join(" e ")}` : "";
  const clientBidText = result.schedules.direct.clientBid > 0.01
    ? ` Lance adicional de recursos próprios estimado em ${formatCurrency(result.schedules.direct.clientBid)}.`
    : " Sem lance adicional de recursos próprios no cenário base.";

  return `Consórcio para aquisição de ${formatCurrency(p.propertyValue)}, com carta bruta estimada de ${formatCurrency(result.schedules.direct.grossCredit)}${bidText}, contemplação simulada no mês ${p.directMonth}.${clientBidText}`;
}

function renderPrintAlternativeCards(result) {
  return result.strategies
    .map((strategy, index) => `
      <article class="print-alt-card" style="--strategy-color: ${strategyColors[strategy.id]}">
        <div class="print-alt-number">${index + 1}</div>
        <div class="print-alt-content">
          <h3>${escapeHtml(strategy.name)}</h3>
          <p>${escapeHtml(getAlternativeDescription(result, strategy))}</p>
          <div class="print-alt-metrics">
            ${printMetric("Aquisição", escapeHtml(formatMonth(strategy.acquisitionMonth)))}
            ${printMetric("Entrada/lance", escapeHtml(formatCurrency(strategy.entryOrInitial)))}
            ${printMetric("Custo", escapeHtml(formatCurrency(strategy.totalOutputs)))}
            ${printMetric("Patrimônio Acumulado", escapeHtml(formatStrategyPatrimony(strategy)))}
          </div>
        </div>
      </article>
    `)
    .join("");
}

function renderPrintBarRows(result, valueKey) {
  const viable = result.strategies.filter((strategy) => strategy.isViable);
  const maxValue = Math.max(1, ...viable.map((strategy) => Math.abs(strategy[valueKey])));

  return result.strategies
    .map((strategy) => {
      const isViable = strategy.isViable;
      const width = isViable ? Math.max(2, Math.abs(strategy[valueKey]) / maxValue * 100) : 0;
      const value = valueKey === "adjustedToday"
        ? formatStrategyPatrimony(strategy)
        : isViable ? formatCurrency(strategy[valueKey]) : "Inviável";
      const color = valueKey === "adjustedToday" ? strategyColors[strategy.id] : "#766300";

      return `
        <div class="print-chart-row${isViable ? "" : " is-infeasible"}">
          <span>${escapeHtml(strategy.short)}</span>
          <div class="print-chart-track">
            <div class="print-chart-fill" style="width: ${width}%; background: ${color};"></div>
          </div>
          <strong>${escapeHtml(value)}</strong>
        </div>
      `;
    })
    .join("");
}

function renderPrintAssumptions(result) {
  const p = result.inputs;
  const items = [
    ["Imóvel", formatCurrency(p.propertyValue)],
    ["Reserva", formatCurrency(p.reserve)],
    ["Caixa mensal", formatCurrency(p.monthlyCash)],
    ["Horizonte", `${p.horizonMonths} meses`],
    ["Finalidade", p.usage === "Moradia propria" ? "Moradia" : "Investimento para aluguel"],
    ["Rentabilidade", formatPercent(p.annualReturn)],
    ["Valorização", formatPercent(p.appreciationPct)],
    ["Inflação", formatPercent(p.inflationPct)],
  ];

  return items
    .map(([label, value]) => printMetric(label, escapeHtml(value)))
    .join("");
}

function renderPrintRecommendation(result) {
  if (!result.best) {
    return `
      <section class="print-card print-recommendation-card">
        <h2>Nenhuma alternativa viável no cenário</h2>
        <p>Com o caixa livre mensal e a reserva informados, nenhuma alternativa se sustenta sem déficit não coberto. Ajuste as premissas antes de emitir uma recomendação.</p>
      </section>
    `;
  }

  const gap = result.second ? result.best.adjustedToday - result.second.adjustedToday : 0;
  const secondText = result.second
    ? `${formatCurrency(gap)} acima da segunda alternativa.`
    : "Única alternativa viável no cenário informado.";

  return `
    <section class="print-card print-recommendation-card">
      <h2>${escapeHtml(result.best.name)} aparece como a melhor alternativa neste cenário.</h2>
      <div class="print-why-grid">
        <article>
          <i></i>
          <h3>1. Maior patrimônio</h3>
          <p>Patrimônio Acumulado de ${escapeHtml(formatStrategyPatrimony(result.best))}.</p>
        </article>
        <article>
          <i></i>
          <h3>2. Vantagem clara</h3>
          <p>${escapeHtml(secondText)}</p>
        </article>
        <article>
          <i></i>
          <h3>3. Custo estimado</h3>
          <p>Custo de ${escapeHtml(formatCurrency(result.best.totalOutputs))} no horizonte analisado.</p>
        </article>
      </div>
      <div class="print-conclusion">
        <h3>Conclusão</h3>
        <p>A alternativa ${escapeHtml(result.best.name)} se destaca por entregar o maior Patrimônio Acumulado entre as opções viáveis analisadas, mantendo um custo estimado compatível com as premissas de caixa e contemplação adotadas no cenário.</p>
      </div>
    </section>
  `;
}

function renderPrintReport(result) {
  const container = document.getElementById("printReport");
  if (!container) return;

  const bestName = result.best ? result.best.name : "Nenhuma viável";
  const bestPatrimony = result.best ? formatStrategyPatrimony(result.best) : "Inviável";

  container.innerHTML = `
    <section class="print-page print-page-light">
      <header class="print-header">
        <div class="print-wordmark" aria-label="Grupo Ável"><span>grupo</span><i></i><strong>ável.</strong></div>
        <b></b>
      </header>
      <h1>Alternativas analisadas</h1>
      <p class="print-lead">Quatro caminhos foram analisados para a aquisição do imóvel. Cada alternativa abaixo considera as mesmas premissas de valor, prazo, caixa disponível e rentabilidade.</p>
      <div class="print-alt-list">
        ${renderPrintAlternativeCards(result)}
      </div>
      <footer>Ável - Relatório comparativo de aquisição <strong>Página 1</strong></footer>
    </section>

    <section class="print-page print-page-dark">
      <header class="print-header">
        <div class="print-wordmark" aria-label="Grupo Ável"><span>grupo</span><i></i><strong>ável.</strong></div>
        <b></b>
      </header>
      <div class="print-recommendation-hero">
        <div>
          <h1>Recomendação<br>do cenário</h1>
          <p>Com as premissas informadas, a indicação considera Patrimônio Acumulado, custo estimado, momento de aquisição e capacidade de manter o plano.</p>
        </div>
        <aside>
          ${printMetric("Estratégia indicada", escapeHtml(bestName))}
          ${printMetric("Patrimônio Acumulado", escapeHtml(bestPatrimony))}
        </aside>
      </div>
      ${renderPrintRecommendation(result)}
      <footer>Ável - Relatório comparativo de aquisição <strong>Página 2</strong></footer>
    </section>

    <section class="print-page print-page-light">
      <header class="print-header">
        <div class="print-wordmark" aria-label="Grupo Ável"><span>grupo</span><i></i><strong>ável.</strong></div>
        <b></b>
      </header>
      <h1>Gráficos e premissas</h1>
      <section class="print-chart-card">
        <h2>Patrimônio Acumulado</h2>
        ${renderPrintBarRows(result, "adjustedToday")}
      </section>
      <section class="print-chart-card">
        <h2>Custo</h2>
        ${renderPrintBarRows(result, "totalOutputs")}
      </section>
      <section class="print-assumptions-card">
        <h2>Premissas usadas</h2>
        <div>${renderPrintAssumptions(result)}</div>
      </section>
      <footer>Ável - Relatório comparativo de aquisição <strong>Página 3</strong></footer>
    </section>
  `;
}

function render(result) {
  renderHero(result);
  renderStrategyCards(result);
  renderBarChart(result);
  renderLineChart(result);
  renderComparisonRows(result);
  renderAssumptions(result);
  renderProposal(result);
  renderPrintReport(result);
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

  const manualButton = document.getElementById("manualButton");
  const manualDialog = document.getElementById("manualDialog");
  const manualCloseButton = document.getElementById("manualCloseButton");

  manualButton.addEventListener("click", () => {
    manualDialog.querySelector(".manual-content").scrollTop = 0;
    manualDialog.showModal();
  });

  manualCloseButton.addEventListener("click", () => manualDialog.close());
  manualDialog.addEventListener("click", (event) => {
    if (event.target === manualDialog) manualDialog.close();
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
