const MAX_MONTHS = 480;

const defaultInputs = {
  assetType: "Imovel",
  financingSystem: "SAC",
  usage: "Investimento",
  propertyValue: 200000,
  reserve: 90000,
  monthlyCash: 7000,
  horizonMonths: 420,
  entryPct: 0.2,
  finMonths: 420,
  annualInterest: 0.12,
  annualTR: 0,
  insurance: 150,
  consAdminPct: 0.24,
  consReservePct: 0.02,
  consInsurancePct: 0.00035,
  consMonths: 240,
  consAdjustPct: 0.06,
  iqMonth: 60,
  directMonth: 60,
  useEmbeddedBid: true,
  embeddedBidPct: 0.25,
  useFreeBid: false,
  freeBidPct: 0,
  useFixedBid: true,
  fixedBidPct: 0.25,
  availabilityMonth: 1,
  rentPaidPct: 0.0035,
  rentReceivedPct: 0.004,
  rentAdjustPct: 0.06,
  rentalDeductionsPct: 0.2,
  annualReturn: 0.12,
  propertyAppreciationPct: 0.06,
  vehicleDepreciationPct: -0.15,
  inflationPct: 0.06,
  acquisitionFinMonth: 1,
  cardEnabled: true,
  cardPurchaseMonth: 1,
  cardUseMonth: 1,
  cardGrossCredit: 200000,
  cardRestrictions: 0,
  cardPrice: 45000,
  cardTransferFee: 2000,
  cardRemainingMonths: 180,
  cardInstallment: 1200,
  cardAdjustPct: 0.06,
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
  "propertyAppreciationPct",
  "inflationPct",
  "cardAdjustPct",
]);

const integerFields = new Set([
  "horizonMonths",
  "finMonths",
  "consMonths",
  "iqMonth",
  "directMonth",
  "availabilityMonth",
  "acquisitionFinMonth",
  "cardPurchaseMonth",
  "cardUseMonth",
]);

const booleanFields = new Set([
  "useEmbeddedBid",
  "useFreeBid",
  "useFixedBid",
  "cardEnabled",
]);

const currencyFields = new Set([
  "propertyValue",
  "reserve",
  "monthlyCash",
  "insurance",
  "cardGrossCredit",
  "cardRestrictions",
  "cardPrice",
  "cardTransferFee",
  "cardInstallment",
]);

const strategyColors = {
  fin: "#111111",
  amort: "#115e59",
  iq: "#244d62",
  cons: "#8f3d49",
  card: "#766300",
};

const strategyNames = {
  fin: "Financiamento sem amortização",
  amort: "Financiamento com amortização",
  iq: "Financiamento + consórcio IQ",
  cons: "Consórcio direto",
  card: "Compra de carta contemplada",
};

const shortNames = {
  fin: "Fin. sem amort.",
  amort: "Fin. amort.",
  iq: "Fin. + IQ",
  cons: "Consórcio direto",
  card: "Carta contemplada",
};

const observations = {
  fin: "Compra por financiamento no mês inicial pelo valor atual do bem.",
  amort: "Pagamento extra usa a sobra depois da prestação ordinária SAC ou PRICE.",
  iq: "A carta quita o financiamento, mas as parcelas futuras do consórcio permanecem como obrigação.",
  cons: "Antes da contemplação reconhece direito da cota e obrigação; depois reconhece o bem.",
  card: "Compra carta já contemplada e considera custo inicial, complemento e parcelas restantes.",
};

function monthlyRate(annualRate) {
  return Math.pow(1 + annualRate, 1 / 12) - 1;
}

function annualStepFactor(annualRate, month) {
  return Math.pow(1 + annualRate, Math.floor((month - 1) / 12));
}

function assetValueAtMonth(p, month) {
  return Math.max(0, p.propertyValue * Math.pow(1 + p.selectedAssetRate, (month - 1) / 12));
}

function addRemainingObligations(rows, paymentKey, discountRate) {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (index === rows.length - 1) {
      rows[index].obligation = 0;
    } else {
      rows[index].obligation =
        (rows[index + 1].obligation + (rows[index + 1][paymentKey] || 0)) /
        (1 + discountRate);
    }
  }
  return rows;
}

function clampNumber(value, min, max) {
  const number = Number.isFinite(value) ? value : min;
  return Math.min(max, Math.max(min, number));
}

function normalizeInputs(raw) {
  const p = { ...defaultInputs, ...raw };

  if (raw && raw.appreciationPct !== undefined) {
    if ((raw.assetType || p.assetType) === "Veiculo" && raw.vehicleDepreciationPct === undefined) {
      p.vehicleDepreciationPct = Number(raw.appreciationPct);
    }
    if ((raw.assetType || p.assetType) !== "Veiculo" && raw.propertyAppreciationPct === undefined) {
      p.propertyAppreciationPct = Number(raw.appreciationPct);
    }
  }

  p.propertyValue = Math.max(0, Number(p.propertyValue) || 0);
  p.reserve = Math.max(0, Number(p.reserve) || 0);
  p.monthlyCash = Math.max(0, Number(p.monthlyCash) || 0);
  p.insurance = Math.max(0, Number(p.insurance) || 0);
  p.cardGrossCredit = Math.max(0, Number(p.cardGrossCredit) || 0);
  p.cardRestrictions = Math.max(0, Number(p.cardRestrictions) || 0);
  p.cardPrice = Math.max(0, Number(p.cardPrice) || 0);
  p.cardTransferFee = Math.max(0, Number(p.cardTransferFee) || 0);
  p.cardInstallment = Math.max(0, Number(p.cardInstallment) || 0);

  for (const key of percentFields) {
    p[key] = Math.max(0, Number(p[key]) || 0);
  }
  p.vehicleDepreciationPct = clampNumber(Number(p.vehicleDepreciationPct), -0.999999, 10);

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
  p.cardPurchaseMonth = Math.round(clampNumber(p.cardPurchaseMonth, 1, MAX_MONTHS));
  p.cardUseMonth = Math.round(clampNumber(p.cardUseMonth, 1, MAX_MONTHS));
  p.cardRemainingMonths = Math.round(clampNumber(Number(p.cardRemainingMonths) || 0, 0, MAX_MONTHS));

  p.entryPct = clampNumber(p.entryPct, 0, 1);
  p.rentalDeductionsPct = clampNumber(p.rentalDeductionsPct, 0, 1);
  p.embeddedBidPct = clampNumber(p.embeddedBidPct, 0, 0.9);
  p.freeBidPct = clampNumber(p.freeBidPct, 0, 0.9);
  p.fixedBidPct = clampNumber(p.fixedBidPct, 0, 0.9);
  p.assetType = p.assetType === "Veiculo" ? "Veiculo" : "Imovel";
  p.financingSystem = p.financingSystem === "PRICE" ? "PRICE" : "SAC";
  const usageAliases = {
    "Moradia propria": "Moradia",
    "Uso proprio": "Moradia",
    "Geracao de renda": "Investimento",
    "Investimento para aluguel": "Investimento",
    "Uso proprio sem aluguel": "Aquisicao",
    "Apenas aquisicao": "Aquisicao",
  };
  p.usage = usageAliases[p.usage] || p.usage;
  if (!["Moradia", "Investimento", "Aquisicao"].includes(p.usage)) {
    p.usage = "Investimento";
  }

  for (const key of booleanFields) {
    p[key] = Boolean(p[key]);
  }
  if (p.useFixedBid && p.useFreeBid) {
    p.useFreeBid = false;
  }
  const totalBidPct = p.useFixedBid ? p.fixedBidPct : p.useFreeBid ? p.freeBidPct : 0;
  if (totalBidPct <= 0) {
    p.useEmbeddedBid = false;
    p.embeddedBidPct = 0;
  } else {
    p.embeddedBidPct = Math.min(p.embeddedBidPct, totalBidPct);
  }
  if (p.assetType === "Veiculo") {
    p.usage = "Aquisicao";
  }

  p.selectedAssetRate = p.assetType === "Veiculo"
    ? p.vehicleDepreciationPct
    : p.propertyAppreciationPct;
  return p;
}

function parseR2Payload(hash) {
  if (!hash || typeof hash !== "string") return null;
  try {
    const params = new URLSearchParams(hash.replace(/^#/, ""));
    const rawPayload = params.get("r2");
    if (!rawPayload) return null;
    const payload = JSON.parse(rawPayload);
    const inputs = payload && payload.inputs ? payload.inputs : payload;
    if (!inputs || typeof inputs !== "object" || Array.isArray(inputs)) return null;
    return {
      source: payload.source || "avel-r2-vantagens-financeiras",
      clientName: payload.clientName || "",
      goalId: payload.goalId || "",
      goalName: payload.goalName || "",
      returnToken: payload.returnToken || "",
      referenceStrategyId: payload.referenceStrategyId || "fin",
      inputs: normalizeInputs(inputs),
    };
  } catch (error) {
    return null;
  }
}

function buildFinancingSchedule(p) {
  const interestMonthly = monthlyRate(p.annualInterest);
  const indexerMonthly = monthlyRate(p.annualTR);
  const acquisitionPrice = p.propertyValue;
  const entry = acquisitionPrice * p.entryPct;
  const financed = Math.max(0, acquisitionPrice - entry);
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

    const correction = balance * indexerMonthly;
    const corrected = balance + correction;
    const remaining = Math.max(1, p.finMonths - month + 1);
    const interest = corrected * interestMonthly;
    const financialPayment = interestMonthly === 0
      ? corrected / remaining
      : corrected * interestMonthly / (1 - Math.pow(1 + interestMonthly, -remaining));
    const amortization = p.financingSystem === "PRICE"
      ? Math.max(0, Math.min(corrected, financialPayment - interest))
      : Math.min(corrected, corrected / remaining);
    const insurance = p.insurance;
    const payment = p.financingSystem === "PRICE"
      ? financialPayment + insurance
      : amortization + interest + insurance;
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
  const indexerMonthly = monthlyRate(p.annualTR);
  const acquisitionPrice = p.propertyValue;
  const entry = acquisitionPrice * p.entryPct;
  const financed = Math.max(0, acquisitionPrice - entry);
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

    const correction = balance * indexerMonthly;
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
  const contemplationContractMonth = p.iqMonth - p.acquisitionFinMonth + 1;
  const balanceAtContemplation = contemplationContractMonth > 0
    ? financingRows[contemplationContractMonth - 1]?.balance || 0
    : 0;
  const letterInitial = factorAtContemplation > 0 ? balanceAtContemplation / factorAtContemplation : 0;
  const creditUpdated = letterInitial * factorAtContemplation;
  const installmentInitial = p.consMonths > 0 ? (letterInitial * (1 + totalFees)) / p.consMonths : 0;
  const returnMonthly = monthlyRate(p.annualReturn);
  const rows = [];

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const factor = annualStepFactor(p.consAdjustPct, month);
    const consInstallment = month <= p.consMonths ? installmentInitial * factor : 0;
    const contractMonth = month - p.acquisitionFinMonth + 1;
    const financingActive = contractMonth >= 1 && contractMonth <= p.finMonths && month <= p.iqMonth;
    const finRow = financingActive ? financingRows[contractMonth - 1] : null;
    const finPayment = finRow?.payment || 0;
    const financingBalance = finRow?.balance || 0;
    const payment = consInstallment + finPayment;

    rows.push({
      month,
      consInstallment,
      finPayment,
      payment,
      financingBalance,
    });
  }

  addRemainingObligations(rows, "consInstallment", returnMonthly);
  rows.forEach((row) => {
    row.right = row.month < p.iqMonth
      ? creditUpdated / Math.pow(1 + returnMonthly, p.iqMonth - row.month)
      : 0;
    const residualFinancing = row.month < p.iqMonth
      ? row.financingBalance
      : row.month === p.iqMonth
        ? Math.max(0, row.financingBalance - creditUpdated)
        : 0;
    row.debt = residualFinancing + row.obligation;
  });

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
  const acquisitionPrice = assetValueAtMonth(p, p.directMonth);
  const totalBidPct = p.useFixedBid ? p.fixedBidPct : p.useFreeBid ? p.freeBidPct : 0;
  const embedded = p.useEmbeddedBid ? Math.min(p.embeddedBidPct, totalBidPct) : 0;
  const clientBidPct = Math.max(0, totalBidPct - embedded);
  const factorAtContemplation = annualStepFactor(p.consAdjustPct, p.directMonth);
  const denominator = Math.max(0.01, factorAtContemplation * (1 - embedded));
  const letterInitial = acquisitionPrice / denominator;
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
    });
  }

  const returnMonthly = monthlyRate(p.annualReturn);
  addRemainingObligations(rows, "installment", returnMonthly);
  rows.forEach((row) => {
    row.right = row.month < p.directMonth
      ? netCredit / Math.pow(1 + returnMonthly, p.directMonth - row.month)
      : 0;
    row.debt = row.obligation;
  });

  return {
    rows,
    acquisitionPrice,
    embedded,
    totalBidPct,
    clientBidPct,
    letterInitial,
    grossCredit,
    netCredit,
    clientBid,
    installmentInitial,
  };
}

function buildContemplatedCardSchedule(p) {
  const acquisitionMonth = Math.max(p.cardPurchaseMonth, p.cardUseMonth);
  const acquisitionPrice = assetValueAtMonth(p, acquisitionMonth);
  const netCredit = Math.max(0, p.cardGrossCredit - p.cardRestrictions);
  const complement = Math.max(0, acquisitionPrice - netCredit);
  const ignoredSurplus = Math.max(0, netCredit - acquisitionPrice);
  const releasedSurplus = 0;
  const initialCost = Math.max(0, p.cardPrice + p.cardTransferFee + complement);
  const lastInstallmentMonth =
    p.cardEnabled && p.cardRemainingMonths > 0
      ? p.cardPurchaseMonth + p.cardRemainingMonths - 1
      : 0;
  const scheduleMonths = Math.max(MAX_MONTHS, lastInstallmentMonth);
  const rows = [];

  for (let month = 1; month <= scheduleMonths; month += 1) {
    const inInstallmentWindow =
      p.cardEnabled &&
      month >= p.cardPurchaseMonth &&
      month < p.cardPurchaseMonth + p.cardRemainingMonths;
    const installmentFactor = month < p.cardPurchaseMonth
      ? 1
      : annualStepFactor(p.cardAdjustPct, month - p.cardPurchaseMonth + 1);
    const installment = inInstallmentWindow ? p.cardInstallment * installmentFactor : 0;

    rows.push({
      month,
      installment,
      initialCost: p.cardEnabled && month === p.cardPurchaseMonth ? initialCost : 0,
    });
  }

  const returnMonthly = monthlyRate(p.annualReturn);
  addRemainingObligations(rows, "installment", returnMonthly);
  rows.forEach((row) => {
    row.debt = p.cardEnabled && row.month >= p.cardPurchaseMonth ? row.obligation : 0;
  });

  return {
    rows,
    acquisitionMonth,
    acquisitionPrice,
    netCredit,
    complement,
    ignoredSurplus,
    releasedSurplus,
    initialCost,
    lastInstallmentMonth,
    scheduleMonths,
    economicCost:
      p.cardPrice +
      p.cardTransferFee +
      rows.reduce((sum, row, index) => sum + row.installment / Math.pow(1 + returnMonthly, index + 1), 0) -
      netCredit,
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
    const propertyValueMonth = assetValueAtMonth(p, month);
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
  const financingPrice = p.propertyValue;
  const entry = financingPrice * p.entryPct;
  const finAvailability = Math.max(p.acquisitionFinMonth, p.availabilityMonth);
  const directAvailability = Math.max(p.directMonth, p.availabilityMonth);
  const cardAvailability = Math.max(schedules.card.acquisitionMonth, p.availabilityMonth);

  for (let month = 1; month <= MAX_MONTHS; month += 1) {
    const shared = sharedRows[month - 1];
    const contractMonth = month - p.acquisitionFinMonth + 1;
    const contractIndex = contractMonth - 1;
    let entryOrBid = 0;
    let payment = 0;
    let debt = 0;
    let acquiredMonth = p.acquisitionFinMonth;
    let availability = finAvailability;
    let economicRight = 0;
    let strategyActive = true;

    if (strategyId === "fin") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = contractMonth >= 1 && contractMonth <= p.finMonths
        ? schedules.financing[contractIndex]?.payment || 0
        : 0;
      debt = contractMonth >= 1 && contractMonth <= p.finMonths
        ? schedules.financing[contractIndex]?.balance || 0
        : 0;
    }

    if (strategyId === "amort") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = contractMonth >= 1 && contractMonth <= p.finMonths
        ? schedules.amortization.rows[contractIndex]?.payment || 0
        : 0;
      debt = contractMonth >= 1 && contractMonth <= p.finMonths
        ? schedules.amortization.rows[contractIndex]?.balance || 0
        : 0;
    }

    if (strategyId === "iq") {
      entryOrBid = month === p.acquisitionFinMonth ? entry : 0;
      payment = schedules.iq.rows[month - 1]?.payment || 0;
      debt = schedules.iq.rows[month - 1]?.debt || 0;
      economicRight = schedules.iq.rows[month - 1]?.right || 0;
    }

    if (strategyId === "cons") {
      entryOrBid = month === p.directMonth ? schedules.direct.clientBid : 0;
      payment = schedules.direct.rows[month - 1]?.payment || 0;
      debt = schedules.direct.rows[month - 1]?.debt || 0;
      economicRight = schedules.direct.rows[month - 1]?.right || 0;
      acquiredMonth = p.directMonth;
      availability = directAvailability;
    }

    if (strategyId === "card") {
      strategyActive = p.cardEnabled;
      entryOrBid = schedules.card.rows[month - 1]?.initialCost || 0;
      payment = schedules.card.rows[month - 1]?.installment || 0;
      debt = schedules.card.rows[month - 1]?.debt || 0;
      acquiredMonth = schedules.card.acquisitionMonth;
      availability = cardAvailability;
    }

    const rentPaid =
      strategyActive &&
      p.assetType === "Imovel" &&
      p.usage === "Moradia" &&
      month < availability
        ? shared.rentPaidBase
        : 0;
    const rentReceived =
      strategyActive &&
      p.assetType === "Imovel" &&
      p.usage === "Investimento" &&
      month >= availability
        ? shared.rentReceivedBase
        : 0;
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

    const property = strategyActive && month >= acquiredMonth ? shared.propertyValueMonth : 0;
    const assetPosition = property + economicRight;
    const grossWorth = invested + assetPosition - debt;
    const adjustedWorth = grossWorth;

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
      property: assetPosition,
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

  if (id === "card") {
    const firstInstallment = schedules.card.rows.find((row) => row.installment > 0)?.installment || 0;
    return [{ label: "Carta contemplada", value: firstInstallment }];
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
  const isEnabled = id !== "card" || p.cardEnabled;
  const isViable = isEnabled && final.capitalVf <= 1;

  const acquisitionMonth =
    id === "cons"
      ? p.directMonth
      : id === "card"
        ? schedules.card.acquisitionMonth
        : p.acquisitionFinMonth;
  const payoffMonth =
    id === "fin"
      ? p.acquisitionFinMonth + p.finMonths - 1
      : id === "amort"
        ? p.acquisitionFinMonth + schedules.amortization.payoffMonth - 1
        : id === "iq"
          ? p.consMonths
          : id === "cons"
            ? p.consMonths
            : p.cardPurchaseMonth + p.cardRemainingMonths - 1;

  const entryOrInitial =
    id === "cons"
      ? schedules.direct.clientBid
      : id === "card"
        ? schedules.card.initialCost
        : entry;

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
    vplVsBase: gapFv / Math.pow(1 + monthlyRate(p.annualReturn), horizon),
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
  const card = buildContemplatedCardSchedule(p);
  const sharedRows = buildSharedRows(p);
  const schedules = { financing, amortization, iq, direct, card };

  const strategies = ["fin", "amort", "iq", "cons", "card"].map((id) => {
    const rows = buildStrategyRows(p, sharedRows, schedules, id);
    return summarizeStrategy(p, sharedRows, rows, schedules, id);
  });

  const viableStrategies = strategies.filter((strategy) => strategy.isViable);
  const sorted = [...viableStrategies].sort((a, b) => b.adjustedToday - a.adjustedToday);
  const best = sorted[0] || null;
  const second = sorted[1] || null;
  const financingPrice = p.propertyValue;
  const entry = financingPrice * p.entryPct;
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
      financingPrice,
      financed: Math.max(0, financingPrice - entry),
      monthlyInterest: monthlyRate(p.annualInterest),
      monthlyIndexer: monthlyRate(p.annualTR),
      monthlyTR: monthlyRate(p.annualTR),
      monthlyReturn: monthlyRate(p.annualReturn),
      totalConsortiumFees: totalFees,
      rentPaidInitial: p.propertyValue * p.rentPaidPct,
      rentReceivedInitial: p.propertyValue * p.rentReceivedPct * (1 - p.rentalDeductionsPct),
      selectedAssetRate: p.selectedAssetRate,
      directAcquisitionPrice: direct.acquisitionPrice,
      cardAcquisitionPrice: card.acquisitionPrice,
      baseFv: sharedRows[p.horizonMonths - 1].baseFv,
    },
  };
}

function summarizeR2Strategy(strategy) {
  if (!strategy) return null;
  return {
    id: strategy.id,
    name: strategy.name,
    acquisitionMonth: strategy.acquisitionMonth,
    payoffMonth: strategy.payoffMonth,
    entryOrInitial: strategy.entryOrInitial,
    initialPaymentLines: strategy.initialPaymentLines,
    totalOutputs: strategy.totalOutputs,
    totalInputs: strategy.totalInputs,
    netFlow: strategy.netFlow,
    isViable: strategy.isViable,
    firstDeficitMonth: strategy.firstDeficitMonth,
    totalUncoveredDeficit: strategy.totalUncoveredDeficit,
    adjustedToday: strategy.adjustedToday,
    gapFv: strategy.gapFv,
    vplVsBase: strategy.vplVsBase,
    debt: strategy.debt,
  };
}

function buildR2AcquisitionAnalysis(result, options = {}) {
  const best = result.best || null;
  const requestedReferenceId = options.referenceStrategyId || "fin";
  const requestedReference = result.strategies.find((strategy) => strategy.id === requestedReferenceId) || null;
  let reference = requestedReference && requestedReference.isViable ? requestedReference : null;
  let referenceFallback = false;

  if (!reference && best) {
    reference = result.viableStrategies.find((strategy) => strategy.id !== best.id) || null;
    referenceFallback = Boolean(reference);
  }

  const advantage = best && reference
    ? Math.max(0, best.adjustedToday - reference.adjustedToday)
    : null;

  return {
    status: options.status || "reviewed",
    source: "calculadora-aquisicao",
    calculatedAt: new Date().toISOString(),
    goalId: options.goalId || "",
    goalName: options.goalName || "",
    basis: "patrimonio_acumulado_valor_presente",
    horizonMonths: result.inputs.horizonMonths,
    inputs: result.inputs,
    requestedReferenceStrategyId: requestedReferenceId,
    referenceFallback,
    viableCount: result.viableStrategies.length,
    bestStrategy: summarizeR2Strategy(best),
    referenceStrategy: summarizeR2Strategy(reference),
    advantagePresentValue: advantage,
    strategies: result.strategies.map(summarizeR2Strategy),
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
  document.querySelectorAll("[data-segment].is-active").forEach((button) => {
    state[button.dataset.segment] = button.dataset.value;
  });
  return state;
}

function updateConditionalFields(state) {
  const normalized = normalizeInputs(state);
  document.querySelectorAll("[data-asset-only]").forEach((element) => {
    element.hidden = element.dataset.assetOnly !== normalized.assetType;
  });
  document.querySelectorAll("[data-real-estate-only]").forEach((element) => {
    element.hidden = normalized.assetType !== "Imovel";
  });
  document.querySelectorAll("[data-purpose-only]").forEach((element) => {
    element.hidden =
      normalized.assetType !== "Imovel" ||
      element.dataset.purposeOnly !== normalized.usage;
  });
  document.querySelectorAll("[data-bid-only]").forEach((element) => {
    const active = element.dataset.bidOnly === "fixed"
      ? normalized.useFixedBid
      : normalized.useFreeBid;
    element.hidden = !active;
  });
  document.querySelectorAll("[data-embedded-only]").forEach((element) => {
    element.hidden = !normalized.useEmbeddedBid;
  });
}

function synchronizeBidControls(changedKey) {
  const fixed = document.querySelector('[data-field="useFixedBid"]');
  const free = document.querySelector('[data-field="useFreeBid"]');
  const embedded = document.querySelector('[data-field="useEmbeddedBid"]');
  if (!fixed || !free || !embedded) return;

  if (changedKey === "useFixedBid" && fixed.checked) free.checked = false;
  if (changedKey === "useFreeBid" && free.checked) fixed.checked = false;
  if (!fixed.checked && !free.checked) embedded.checked = false;

  const totalField = fixed.checked
    ? document.querySelector('[data-field="fixedBidPct"]')
    : free.checked
      ? document.querySelector('[data-field="freeBidPct"]')
      : null;
  const embeddedField = document.querySelector('[data-field="embeddedBidPct"]');
  if (embedded.checked && totalField && embeddedField) {
    embeddedField.max = totalField.value || "0";
    if (Number(embeddedField.value) > Number(totalField.value)) {
      embeddedField.value = totalField.value;
    }
  }
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

  document.querySelectorAll("[data-segment]").forEach((button) => {
    const key = button.dataset.segment;
    button.classList.toggle("is-active", button.dataset.value === normalized[key]);
  });
  updateConditionalFields(normalized);
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
    [p.assetType === "Veiculo" ? "Veículo" : "Imóvel", formatCurrency(p.propertyValue)],
    ["Reserva", formatCurrency(p.reserve)],
    ["Caixa mensal", formatCurrency(p.monthlyCash)],
    ["Horizonte", `${p.horizonMonths} meses`],
    ["Financiamento", `${p.financingSystem}, ${p.finMonths} meses a ${formatPercent(p.annualInterest)}`],
    ["Consórcio", `${p.consMonths} meses, reajuste ${formatPercent(p.consAdjustPct)}`],
    ["Rentabilidade", formatPercent(p.annualReturn)],
  ];
  if (p.assetType === "Imovel") items.splice(6, 0, ["Finalidade", p.usage]);

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
    `Tipo e valor do bem: ${p.assetType === "Veiculo" ? "veículo" : "imóvel"} de ${formatCurrency(p.propertyValue)}.`,
    `Entrada do financiamento: ${formatPercent(p.entryPct)} (${formatCurrency(result.derived.entry)}).`,
    `Sistema de financiamento: ${p.financingSystem}; indexador anual de ${formatPercent(p.annualTR)}.`,
    "Preço de aquisição: financiamento pelo valor atual no mês 1; consórcios pelo valor futuro no mês de uso do crédito.",
    `Caixa livre mensal considerado: ${formatCurrency(p.monthlyCash)}.`,
    `Horizonte comparativo: ${p.horizonMonths} meses.`,
    `Rentabilidade líquida anual: ${formatPercent(p.annualReturn)}.`,
  ];
  if (p.assetType === "Imovel") assumptions.push(`Finalidade do imóvel: ${p.usage.toLowerCase()}.`);
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
    return `Financiamento ${p.financingSystem} de ${formatCurrency(financed)}, com entrada de ${formatCurrency(entry)}, sem amortização extraordinária até o final do prazo de ${p.finMonths} meses.`;
  }

  if (strategy.id === "amort") {
    return `Financiamento ${p.financingSystem} de ${formatCurrency(financed)}, com entrada de ${formatCurrency(entry)} e amortização mensal usando a disponibilidade de caixa prevista até a quitação.`;
  }

  if (strategy.id === "iq") {
    return `Financiamento ${p.financingSystem} de ${formatCurrency(financed)}, combinado com consórcio de aproximadamente ${formatCurrency(result.schedules.iq.creditUpdated)} no mês ${p.iqMonth}. O patrimônio desconta o VP das parcelas futuras até o mês ${p.consMonths}.`;
  }

  if (strategy.id === "card") {
    return `Carta contemplada com crédito líquido de ${formatCurrency(result.schedules.card.netCredit)}, preço de compra de ${formatCurrency(p.cardPrice)}, taxa de transferência de ${formatCurrency(p.cardTransferFee)} e ${p.cardRemainingMonths} parcelas restantes.`;
  }

  const bidParts = [];
  if (p.useEmbeddedBid) bidParts.push(`lance embutido de ${formatPercent(result.schedules.direct.embedded)}`);
  if (p.useFreeBid) bidParts.push(`lance livre de ${formatPercent(p.freeBidPct)}`);
  if (p.useFixedBid) bidParts.push(`lance fixo de ${formatPercent(p.fixedBidPct)}`);
  const bidText = bidParts.length ? `, com ${bidParts.join(" e ")}` : "";
  const clientBidText = result.schedules.direct.clientBid > 0.01
    ? ` Lance adicional de recursos próprios estimado em ${formatCurrency(result.schedules.direct.clientBid)}.`
    : " Sem lance adicional de recursos próprios no cenário base.";

  return `Consórcio direto para preço de aquisição de ${formatCurrency(result.schedules.direct.acquisitionPrice)}, com carta bruta estimada de ${formatCurrency(result.schedules.direct.grossCredit)}${bidText}, contemplação simulada no mês ${p.directMonth}.${clientBidText}`;
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
    [p.assetType === "Veiculo" ? "Veículo" : "Imóvel", formatCurrency(p.propertyValue)],
    ["Reserva", formatCurrency(p.reserve)],
    ["Caixa mensal", formatCurrency(p.monthlyCash)],
    ["Horizonte", `${p.horizonMonths} meses`],
    ["Sistema", p.financingSystem],
    ["Rentabilidade", formatPercent(p.annualReturn)],
    ["Taxa do bem", formatPercent(p.selectedAssetRate)],
    ["Inflação", formatPercent(p.inflationPct)],
  ];
  if (p.assetType === "Imovel") items.splice(4, 0, ["Finalidade", p.usage]);

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
      <p class="print-lead">Cinco caminhos foram analisados para a aquisição do bem. Cada alternativa usa as mesmas premissas de valor, prazo, caixa, preço de aquisição e rentabilidade.</p>
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
  const importedScenario = parseR2Payload(window.location.hash);
  const startupInputs = importedScenario ? importedScenario.inputs : defaultInputs;
  setDomState(startupInputs);
  render(calculateScenario(startupInputs));

  document.querySelectorAll("[data-field]").forEach((element) => {
    element.addEventListener("input", () => {
      synchronizeBidControls(element.dataset.field);
      updateConditionalFields(getDomState());
      update();
    });
    element.addEventListener("change", () => {
      synchronizeBidControls(element.dataset.field);
      updateConditionalFields(getDomState());
      update();
    });
    if (currencyFields.has(element.dataset.field)) {
      element.addEventListener("blur", () => {
        const key = element.dataset.field;
        element.value = inputValueForDisplay(key, parseInputValue(key, element));
        update();
      });
    }
  });

  document.querySelectorAll("[data-segment]").forEach((button) => {
    button.addEventListener("click", () => {
      document
        .querySelectorAll(`[data-segment="${button.dataset.segment}"]`)
        .forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      if (button.dataset.segment === "financingSystem" && button.dataset.value === "PRICE") {
        const indexer = document.querySelector('[data-field="annualTR"]');
        if (indexer) indexer.value = 0;
      }
      updateConditionalFields(getDomState());
      update();
    });
  });

  document.getElementById("resetButton").addEventListener("click", () => {
    setDomState(startupInputs);
    render(calculateScenario(startupInputs));
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

  const applyR2Button = document.getElementById("applyR2Button");
  if (importedScenario && importedScenario.returnToken && window.opener) {
    applyR2Button.hidden = false;
    applyR2Button.addEventListener("click", () => {
      const result = calculateScenario(getDomState());
      const analysis = buildR2AcquisitionAnalysis(result, {
        status: "reviewed",
        goalId: importedScenario.goalId,
        goalName: importedScenario.goalName,
        referenceStrategyId: importedScenario.referenceStrategyId,
      });
      window.opener.postMessage({
        type: "avel-r2-acquisition-result",
        returnToken: importedScenario.returnToken,
        goalId: importedScenario.goalId,
        analysis,
      }, "*");
      applyR2Button.querySelector("span").textContent = "Aplicado à R2";
      applyR2Button.classList.add("is-applied");
    });
  }

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
    parseR2Payload,
    calculateScenario,
    buildR2AcquisitionAnalysis,
    buildFinancingSchedule,
    buildAmortizationSchedule,
    buildIqSchedule,
    buildDirectConsortiumSchedule,
    buildContemplatedCardSchedule,
  };
}
