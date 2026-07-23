(function attachReportEngine(root, factory) {
  "use strict";

  const api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }

  if (root) {
    root.ConsfinReportEngine = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function createReportEngine() {
  "use strict";

  const VERSION = "1.0.0";
  const SCHEMA_VERSION = "consfin-decision-report/v1";
  const DEFAULT_PRIORITY = "wealth";
  const PRIORITIES = Object.freeze(["wealth", "cost", "payoff", "acquisition", "compare"]);
  const OUTCOME_STATES = Object.freeze({
    INVALID: "invalid",
    NO_VIABLE: "no_viable",
    SINGLE_VIABLE: "single_viable",
    SAME_LEADER: "same_leader",
    SPLIT_LEADERS: "split_leaders",
    NEAR_TIE: "near_tie",
  });
  const NEAR_TIE_RATIO = 0.01;

  function isFiniteNumber(value) {
    return typeof value === "number" && Number.isFinite(value);
  }

  function asFiniteNumber(value, fallback = null) {
    return isFiniteNumber(value) ? value : fallback;
  }

  function cleanText(value, fallback) {
    const normalized = typeof value === "string" ? value.trim() : "";
    return normalized || fallback;
  }

  function normalizeMeta(meta) {
    const source = meta && typeof meta === "object" ? meta : {};
    const requestedPriority = source.priority === "comparison" ? "compare" : source.priority;
    const priority = PRIORITIES.includes(requestedPriority) ? requestedPriority : DEFAULT_PRIORITY;

    return {
      clientName: cleanText(source.clientName, "Cliente"),
      plannerName: cleanText(source.plannerName, "Não informado"),
      priority,
      source: source.source === "r2" ? "r2" : "standalone",
      generatedAt: cleanText(source.generatedAt, ""),
    };
  }

  function getCost(strategy) {
    if (isFiniteNumber(strategy.lifetimeContractualOutflow)) {
      return {
        value: strategy.lifetimeContractualOutflow,
        source: "contractual_lifetime",
      };
    }

    if (isFiniteNumber(strategy.contractualOutflowsLifetime)) {
      return {
        value: strategy.contractualOutflowsLifetime,
        source: "contractual_lifetime",
      };
    }

    return {
      value: null,
      source: "missing",
    };
  }

  function uniqueAlerts(alerts) {
    const seen = new Set();
    return alerts.filter((alert) => {
      if (!alert || seen.has(alert.code)) return false;
      seen.add(alert.code);
      return true;
    });
  }

  function strategyAlerts(strategy, horizonMonths) {
    const alerts = [];

    if (!strategy.isEnabled) return alerts;

    if (strategy.isEnabled && !strategy.isViable) {
      alerts.push({
        code: "cash_deficit",
        severity: "attention",
        text: strategy.firstDeficitMonth
          ? `O caixa fica insuficiente a partir do mês ${strategy.firstDeficitMonth}.`
          : "O caixa informado não sustenta esta alternativa.",
      });
    }

    if (
      strategy.isViable &&
      isFiniteNumber(strategy.payoffMonth) &&
      isFiniteNumber(horizonMonths) &&
      strategy.payoffMonth > horizonMonths
    ) {
      alerts.push({
        code: "payments_after_horizon",
        severity: "attention",
        text: `As parcelas seguem após os ${horizonMonths} meses analisados.`,
      });
    }

    if (strategy.id === "card") {
      alerts.push({
        code: "card_validation",
        severity: "attention",
        text: "A carta exige validação de crédito, transferência e parcelas.",
      });
    } else if (strategy.id === "cons" && isFiniteNumber(strategy.acquisitionMonth)) {
      alerts.push({
        code: "consortium_timing",
        severity: "neutral",
        text: `A contemplação foi considerada no mês ${strategy.acquisitionMonth}.`,
      });
    } else if (strategy.id === "iq") {
      alerts.push({
        code: "combined_contracts",
        severity: "neutral",
        text: "A estratégia depende da coordenação entre dois contratos.",
      });
    } else if (strategy.id === "amort") {
      alerts.push({
        code: "monthly_contributions",
        severity: "neutral",
        text: "A quitação antecipada depende dos aportes mensais previstos.",
      });
    } else if (strategy.id === "fin" && isFiniteNumber(strategy.debt) && strategy.debt > 0.01) {
      alerts.push({
        code: "remaining_debt",
        severity: "neutral",
        text: "O saldo devedor permanece até o prazo contratado.",
      });
    }

    return uniqueAlerts(alerts);
  }

  function normalizeStrategy(strategy, index, horizonMonths) {
    const source = strategy && typeof strategy === "object" ? strategy : {};
    const cost = getCost(source);
    const id = cleanText(source.id, `strategy-${index + 1}`);
    const name = cleanText(source.name, cleanText(source.short, `Alternativa ${index + 1}`));
    const short = cleanText(source.short, name);
    const isEnabled = source.isEnabled !== false;
    const hasAffordability = typeof source.isAffordable === "boolean";
    const isAffordable = hasAffordability ? source.isAffordable : source.isViable === true;
    const declaredViable = source.isViable === true;
    const normalized = {
      id,
      name,
      short,
      isEnabled,
      isAffordable,
      isViable: isEnabled && isAffordable && declaredViable,
      statusConflict:
        typeof source.isViable === "boolean" &&
        hasAffordability &&
        source.isViable !== (isEnabled && isAffordable),
      wealth: asFiniteNumber(source.adjustedToday),
      cost: cost.value,
      costSource: cost.source,
      acquisitionMonth: asFiniteNumber(source.acquisitionMonth),
      payoffMonth: asFiniteNumber(source.payoffMonth),
      entryOrInitial: asFiniteNumber(source.entryOrInitial),
      firstDeficitMonth: asFiniteNumber(source.firstDeficitMonth),
      totalUncoveredDeficit: asFiniteNumber(source.totalUncoveredDeficit),
      debt: asFiniteNumber(source.debt),
    };

    normalized.alerts = strategyAlerts(normalized, horizonMonths);
    return normalized;
  }

  function compareById(a, b) {
    return a.id.localeCompare(b.id, "pt-BR");
  }

  function rankStrategies(strategies, key, direction) {
    return strategies
      .filter((strategy) => isFiniteNumber(strategy[key]))
      .slice()
      .sort((a, b) => {
        const difference = direction === "desc" ? b[key] - a[key] : a[key] - b[key];
        return Math.abs(difference) > Number.EPSILON ? difference : compareById(a, b);
      });
  }

  function exactTieTolerance(key) {
    return key === "payoffMonth" || key === "acquisitionMonth" ? 0 : 0.005;
  }

  function coLeaderIds(ranking, key) {
    const leader = ranking[0] || null;
    if (!leader) return [];
    const tolerance = exactTieTolerance(key);
    return ranking
      .filter((strategy) => Math.abs(strategy[key] - leader[key]) <= tolerance)
      .map((strategy) => strategy.id);
  }

  function tieDetails(ranking, key) {
    const leader = ranking[0] || null;
    const runnerUp = ranking[1] || null;
    const exactLeaders = coLeaderIds(ranking, key);

    if (!leader || !runnerUp) {
      return {
        isNear: false,
        isExact: false,
        leaderIds: exactLeaders,
        leaderId: leader ? leader.id : null,
        runnerUpId: null,
        gap: null,
        ratio: null,
      };
    }

    const gap = Math.abs(leader[key] - runnerUp[key]);
    const scale = Math.max(Math.abs(leader[key]), Math.abs(runnerUp[key]), 1);
    const ratio = gap / scale;

    return {
      isNear: ratio <= NEAR_TIE_RATIO,
      isExact: exactLeaders.length > 1,
      leaderIds: exactLeaders,
      leaderId: leader.id,
      runnerUpId: runnerUp.id,
      gap,
      ratio,
    };
  }

  function createLeader(ranking, key) {
    const ids = coLeaderIds(ranking, key);
    return ids.length === 1 ? ids[0] : null;
  }

  function extractFacts(result) {
    const source = result && typeof result === "object" ? result : {};
    const inputs = source.inputs && typeof source.inputs === "object" ? source.inputs : {};
    const horizonMonths = asFiniteNumber(inputs.horizonMonths);
    const errors = [];

    if (!Array.isArray(source.strategies) || !source.strategies.length) {
      errors.push("missing_strategies");
    }

    const strategies = Array.isArray(source.strategies)
      ? source.strategies.map((strategy, index) => normalizeStrategy(strategy, index, horizonMonths))
      : [];
    const viable = strategies.filter((strategy) => strategy.isViable);

    viable.forEach((strategy) => {
      if (!isFiniteNumber(strategy.wealth)) errors.push(`missing_wealth:${strategy.id}`);
      if (!isFiniteNumber(strategy.cost)) errors.push(`missing_cost:${strategy.id}`);
    });

    strategies.forEach((strategy) => {
      if (strategy.statusConflict) errors.push(`inconsistent_status:${strategy.id}`);
    });

    const wealthRanking = rankStrategies(viable, "wealth", "desc");
    const costRanking = rankStrategies(viable, "cost", "asc");
    const payoffRanking = rankStrategies(viable, "payoffMonth", "asc");
    const acquisitionRanking = rankStrategies(viable, "acquisitionMonth", "asc");
    const coLeaders = {
      wealth: coLeaderIds(wealthRanking, "wealth"),
      cost: coLeaderIds(costRanking, "cost"),
      payoff: coLeaderIds(payoffRanking, "payoffMonth"),
      acquisition: coLeaderIds(acquisitionRanking, "acquisitionMonth"),
    };
    const leaders = {
      wealth: createLeader(wealthRanking, "wealth"),
      cost: createLeader(costRanking, "cost"),
      payoff: createLeader(payoffRanking, "payoffMonth"),
      acquisition: createLeader(acquisitionRanking, "acquisitionMonth"),
    };
    const nearTies = {
      wealth: tieDetails(wealthRanking, "wealth"),
      cost: tieDetails(costRanking, "cost"),
    };
    const wealthLeader = viable.find((strategy) => strategy.id === leaders.wealth) || null;
    const costLeader = viable.find((strategy) => strategy.id === leaders.cost) || null;
    const tradeoff = wealthLeader && costLeader
      ? {
          wealthLeaderId: wealthLeader.id,
          costLeaderId: costLeader.id,
          wealthAdvantage: wealthLeader.wealth - costLeader.wealth,
          extraCost: wealthLeader.cost - costLeader.cost,
          payoffDifferenceMonths:
            isFiniteNumber(wealthLeader.payoffMonth) && isFiniteNumber(costLeader.payoffMonth)
              ? wealthLeader.payoffMonth - costLeader.payoffMonth
              : null,
          acquisitionDifferenceMonths:
            isFiniteNumber(wealthLeader.acquisitionMonth) && isFiniteNumber(costLeader.acquisitionMonth)
              ? wealthLeader.acquisitionMonth - costLeader.acquisitionMonth
              : null,
        }
      : null;

    return {
      schemaVersion: SCHEMA_VERSION,
      errors: [...new Set(errors)],
      horizonMonths,
      strategies,
      viableCount: viable.length,
      viableStrategyIds: viable.map((strategy) => strategy.id),
      leaders,
      coLeaders,
      rankings: {
        wealth: wealthRanking.map((strategy) => strategy.id),
        cost: costRanking.map((strategy) => strategy.id),
        payoff: payoffRanking.map((strategy) => strategy.id),
        acquisition: acquisitionRanking.map((strategy) => strategy.id),
      },
      nearTies,
      tradeoff,
      assumptions: {
        assetType: cleanText(inputs.assetType, "Bem"),
        usage: cleanText(inputs.usage, ""),
        propertyValue: asFiniteNumber(inputs.propertyValue),
        reserve: asFiniteNumber(inputs.reserve),
        monthlyCash: asFiniteNumber(inputs.monthlyCash),
        horizonMonths,
        financingSystem: cleanText(inputs.financingSystem, ""),
        annualReturn: asFiniteNumber(inputs.annualReturn),
        inflationPct: asFiniteNumber(inputs.inflationPct),
      },
    };
  }

  function leaderForPriority(facts, priority) {
    if (priority === "compare") return null;
    return facts.leaders[priority] || null;
  }

  function classifyOutcome(facts, meta) {
    const normalizedMeta = normalizeMeta(meta);
    const priority = normalizedMeta.priority;
    const chosenCoLeaders = priority === "compare" ? [] : facts?.coLeaders?.[priority] || [];
    const hasExactTie = chosenCoLeaders.length > 1;
    const hasNearTie =
      (priority === "wealth" && facts?.nearTies?.wealth?.isNear) ||
      (priority === "cost" && facts?.nearTies?.cost?.isNear);
    const metricAvailable = priority === "compare" || chosenCoLeaders.length > 0;
    let state;

    if (!facts || facts.errors.length) {
      state = OUTCOME_STATES.INVALID;
    } else if (facts.viableCount === 0) {
      state = OUTCOME_STATES.NO_VIABLE;
    } else if (facts.viableCount === 1) {
      state = OUTCOME_STATES.SINGLE_VIABLE;
    } else if (hasExactTie || hasNearTie) {
      state = OUTCOME_STATES.NEAR_TIE;
    } else if (facts.leaders.wealth && facts.leaders.wealth === facts.leaders.cost) {
      state = OUTCOME_STATES.SAME_LEADER;
    } else {
      state = OUTCOME_STATES.SPLIT_LEADERS;
    }

    return {
      state,
      priority,
      metricAvailable,
      exactTieMetric: hasExactTie ? priority : null,
      recommendedStrategyId:
        state === OUTCOME_STATES.INVALID || state === OUTCOME_STATES.NO_VIABLE
          ? null
          : leaderForPriority(facts, priority),
      wealthLeaderId: facts && facts.leaders ? facts.leaders.wealth : null,
      costLeaderId: facts && facts.leaders ? facts.leaders.cost : null,
      nearTieMetric:
        state === OUTCOME_STATES.NEAR_TIE ? priority : null,
    };
  }

  function formatMoney(value, compact) {
    if (!isFiniteNumber(value)) return "—";

    if (compact && Math.abs(value) >= 1_000_000) {
      return `R$ ${(value / 1_000_000).toLocaleString("pt-BR", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })} mi`;
    }

    if (compact && Math.abs(value) >= 1_000) {
      return `R$ ${(value / 1_000).toLocaleString("pt-BR", {
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      })} mil`;
    }

    return value.toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function formatPercent(value) {
    if (!isFiniteNumber(value)) return "—";
    return value.toLocaleString("pt-BR", {
      style: "percent",
      minimumFractionDigits: 1,
      maximumFractionDigits: 2,
    });
  }

  function formatDuration(months) {
    if (!isFiniteNumber(months)) return "—";
    const rounded = Math.round(Math.abs(months));
    if (rounded === 0) return "Mesmo prazo";
    if (rounded < 12) return `${rounded} ${rounded === 1 ? "mês" : "meses"}`;

    const years = Math.floor(rounded / 12);
    const remainingMonths = rounded % 12;
    const yearsText = `${years} ${years === 1 ? "ano" : "anos"}`;
    if (!remainingMonths) return yearsText;
    return `${yearsText} e ${remainingMonths} ${remainingMonths === 1 ? "mês" : "meses"}`;
  }

  function formatMonth(month) {
    if (!isFiniteNumber(month)) return "—";
    return `Mês ${Math.round(month)}`;
  }

  function formatDate(value) {
    if (!value) return "Data não informada";
    const plainDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    if (plainDate) return `${plainDate[3]}/${plainDate[2]}/${plainDate[1]}`;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", { timeZone: "America/Sao_Paulo" }).format(date);
  }

  function strategyById(facts, id) {
    return facts.strategies.find((strategy) => strategy.id === id) || null;
  }

  function priorityLabel(priority) {
    const labels = {
      wealth: "Preservar mais patrimônio",
      cost: "Menor custo",
      payoff: "Quitar antes",
      acquisition: "Adquirir antes",
      compare: "Comparar sem recomendação",
    };
    return labels[priority] || labels.wealth;
  }

  function outcomeCopy(facts, outcome) {
    const wealthLeader = strategyById(facts, outcome.wealthLeaderId);
    const costLeader = strategyById(facts, outcome.costLeaderId);
    const recommended = strategyById(facts, outcome.recommendedStrategyId);

    if (outcome.state === OUTCOME_STATES.INVALID) {
      return {
        headline: "Não foi possível concluir a análise",
        summary: "Os dados desta comparação estão incompletos.",
      };
    }

    if (outcome.state === OUTCOME_STATES.NO_VIABLE) {
      return {
        headline: "As premissas precisam ser revistas",
        summary: "Nenhuma alternativa se sustenta com o caixa informado.",
      };
    }

    if (outcome.priority !== "compare" && !outcome.metricAvailable) {
      return {
        headline: "Este critério não está disponível",
        summary: "A simulação não trouxe os dados necessários para concluir esta prioridade.",
      };
    }

    if (outcome.state === OUTCOME_STATES.SINGLE_VIABLE) {
      const onlyViable = strategyById(facts, facts.viableStrategyIds[0]);
      return {
        headline: "Há uma alternativa financeiramente viável",
        summary: `${(recommended || onlyViable).name} é a única opção viável neste cenário.`,
      };
    }

    if (outcome.state === OUTCOME_STATES.NEAR_TIE) {
      if (outcome.exactTieMetric) {
        return {
          headline: "Há um empate no critério escolhido",
          summary: "Nenhuma alternativa deve ser indicada apenas por este critério.",
        };
      }
      return {
        headline: "Os resultados estão próximos",
        summary: "A diferença é pequena. Prazo e segurança de execução ganham mais peso nesta escolha.",
      };
    }

    if (outcome.state === OUTCOME_STATES.SAME_LEADER) {
      return {
        headline: "Uma alternativa lidera nos dois critérios",
        summary: `${wealthLeader.name} reúne o maior patrimônio e o menor custo.`,
      };
    }

    if (wealthLeader && costLeader) {
      return {
        headline: "A prioridade define o melhor caminho",
        summary: `${wealthLeader.name} lidera em patrimônio. ${costLeader.name} apresenta o menor custo.`,
      };
    }

    return {
      headline: "Os critérios não apontam uma escolha única",
      summary: "Há empate em pelo menos um dos resultados comparados.",
    };
  }

  function recommendationCopy(facts, outcome) {
    const strategy = strategyById(facts, outcome.recommendedStrategyId);
    if (!strategy) {
      if (outcome.priority === "compare" && facts.viableCount > 0) {
        return {
          heading: "Comparação sem recomendação",
          text:
            facts.viableCount === 1
              ? "A única alternativa viável está apresentada sem uma recomendação automática."
              : "Os resultados de cada critério estão apresentados para sua decisão.",
          condition: "A escolha final depende da prioridade definida para a aquisição.",
        };
      }

      if (!outcome.metricAvailable) {
        return {
          heading: `${priorityLabel(outcome.priority)} indisponível`,
          text: "A simulação não contém os dados necessários para aplicar este critério.",
          condition: "Revise as premissas e gere uma nova comparação.",
        };
      }

      if (outcome.exactTieMetric) {
        return {
          heading: `Empate em ${priorityLabel(outcome.priority).toLowerCase()}`,
          text: "Duas ou mais alternativas entregam o mesmo resultado neste critério.",
          condition: "Use os demais critérios e as condições de contratação para decidir.",
        };
      }

      return {
        heading: "Parecer não concluído",
        text:
          outcome.state === OUTCOME_STATES.INVALID
            ? "A análise depende da correção dos dados incompletos."
            : "A análise depende da revisão das premissas de caixa.",
        condition: "Uma nova comparação deve ser feita após os ajustes.",
      };
    }

    const texts = {
      wealth: `Para preservar mais patrimônio, ${strategy.name} é a alternativa indicada.`,
      cost: `Para buscar o menor custo, ${strategy.name} é a alternativa indicada.`,
      payoff: `Para encerrar as parcelas antes, ${strategy.name} é a alternativa indicada.`,
      acquisition: `Para adquirir o bem mais cedo, ${strategy.name} é a alternativa indicada.`,
    };
    const conditionByStrategy = {
      card: "A indicação depende da validação da carta e da transferência.",
      cons: `A indicação considera a contemplação no mês ${strategy.acquisitionMonth || "informado"}.`,
      iq: "A indicação considera a execução coordenada dos dois contratos.",
      amort: "A indicação considera a manutenção dos aportes mensais.",
      fin: "A indicação considera as condições atuais do financiamento.",
    };

    if (outcome.state === OUTCOME_STATES.NEAR_TIE) {
      return {
        heading: priorityLabel(outcome.priority),
        text: `A diferença é pequena. Pelo critério escolhido, ${strategy.name} é a referência deste cenário.`,
        condition: conditionByStrategy[strategy.id] || "Confirme as premissas antes da decisão.",
      };
    }

    return {
      heading: priorityLabel(outcome.priority),
      text: texts[outcome.priority] || texts.wealth,
      condition: conditionByStrategy[strategy.id] || "A indicação considera as premissas apresentadas.",
    };
  }

  function tradeoffCopy(facts, outcome) {
    const tradeoff = facts.tradeoff;
    const recommended = strategyById(facts, outcome.recommendedStrategyId);
    const wealthLeader = strategyById(facts, facts.leaders.wealth);
    const costLeader = strategyById(facts, facts.leaders.cost);

    if (outcome.state === OUTCOME_STATES.INVALID) {
      return "A comparação será concluída após a correção dos dados.";
    }

    if (outcome.state === OUTCOME_STATES.NO_VIABLE) {
      return "Reserva, caixa ou prazos precisam ser revistos.";
    }

    if (outcome.state === OUTCOME_STATES.SINGLE_VIABLE) {
      return "A análise encontrou uma única alternativa viável.";
    }

    if (outcome.exactTieMetric) {
      return "O critério escolhido terminou empatado. Os demais resultados devem orientar a decisão.";
    }

    if (outcome.priority !== "compare" && !outcome.metricAvailable) {
      return "Este critério não pode ser aplicado com os dados disponíveis.";
    }

    if (wealthLeader && costLeader && wealthLeader.id === costLeader.id) {
      if (recommended && recommended.id !== wealthLeader.id) {
        return `${wealthLeader.name} lidera em patrimônio e custo. ${recommended.name} atende ao critério de ${priorityLabel(
          outcome.priority,
        ).toLowerCase()}.`;
      }
      return `${wealthLeader.name} lidera em patrimônio e custo.`;
    }

    if (!tradeoff) {
      return "Há empate em pelo menos um critério. Compare também prazo e condições de contratação.";
    }

    if (outcome.priority === "cost") {
      return `A escolha reduz o custo em ${formatMoney(Math.max(0, tradeoff.extraCost), true)}, com ${formatMoney(
        Math.max(0, tradeoff.wealthAdvantage),
        true,
      )} a menos de patrimônio.`;
    }

    if (outcome.priority === "wealth") {
      return `A escolha acrescenta ${formatMoney(Math.max(0, tradeoff.wealthAdvantage), true)} ao patrimônio e exige ${formatMoney(
        Math.max(0, tradeoff.extraCost),
        true,
      )} a mais.`;
    }

    if (outcome.priority === "payoff" && recommended) {
      return `${recommended.name} encerra as parcelas em ${formatMonth(recommended.payoffMonth).toLowerCase()}. Patrimônio e custo podem apontar para alternativas diferentes.`;
    }

    if (outcome.priority === "acquisition" && recommended) {
      return `${recommended.name} permite a aquisição em ${formatMonth(recommended.acquisitionMonth).toLowerCase()}. Patrimônio e custo podem apontar para alternativas diferentes.`;
    }

    if (outcome.priority === "compare" && wealthLeader && costLeader) {
      return `${wealthLeader.name} lidera em patrimônio; ${costLeader.name}, em menor custo.`;
    }

    return "Patrimônio, custo e prazo apontam para escolhas diferentes.";
  }

  function metricRows(facts, key, leaderIds) {
    const values = facts.strategies
      .filter((strategy) => strategy.isViable && isFiniteNumber(strategy[key]))
      .map((strategy) => Math.abs(strategy[key]));
    const scale = Math.max(1, ...values);

    return facts.strategies.map((strategy) => {
      const value = strategy[key];
      const isViable = strategy.isViable && isFiniteNumber(value);
      return {
        id: strategy.id,
        name: strategy.short,
        isViable,
        isLeader: isViable && leaderIds.includes(strategy.id),
        value,
        formattedValue:
          !strategy.isEnabled
            ? "Não considerada"
            : key === "wealth"
              ? isViable ? formatMoney(value, true) : "Inviável"
              : isViable ? formatMoney(value, true) : "Inviável",
        width: isViable ? Math.max(4, Math.min(100, (Math.abs(value) / scale) * 100)) : 0,
      };
    });
  }

  function metricLeaderView(facts, key) {
    const ids = facts.coLeaders[key] || [];
    if (!ids.length) return null;
    const strategies = ids.map((id) => strategyById(facts, id)).filter(Boolean);
    if (!strategies.length) return null;
    const first = strategies[0];
    const value = key === "wealth" ? first.wealth : first.cost;
    return {
      ids,
      name:
        strategies.length === 1
          ? first.name
          : strategies.length === 2
            ? `${strategies[0].short} e ${strategies[1].short} empatam`
            : `${strategies.length} alternativas empatam`,
      value: formatMoney(value, true),
    };
  }

  function decisionCards(facts, outcome, recommended) {
    const tradeoff = facts.tradeoff;

    if (recommended && (outcome.priority === "payoff" || outcome.priority === "acquisition")) {
      return [
        {
          tone: "positive",
          label: "Patrimônio estimado",
          value: formatMoney(recommended.wealth, true),
        },
        {
          tone: "attention",
          label: "Custo contratual",
          value: formatMoney(recommended.cost, true),
        },
        {
          tone: "neutral",
          label: outcome.priority === "acquisition" ? "Aquisição" : "Quitação",
          value: formatMonth(outcome.priority === "acquisition" ? recommended.acquisitionMonth : recommended.payoffMonth),
        },
      ];
    }

    if (!tradeoff || tradeoff.wealthLeaderId === tradeoff.costLeaderId) return [];

    if (outcome.priority === "wealth") {
      return [
        {
          tone: "positive",
          label: "Patrimônio adicional",
          value: formatMoney(Math.max(0, tradeoff.wealthAdvantage), true),
        },
        {
          tone: "attention",
          label: "Custo adicional",
          value: formatMoney(Math.max(0, tradeoff.extraCost), true),
        },
        {
          tone: "neutral",
          label: "Quitação",
          value: recommended ? formatMonth(recommended.payoffMonth) : "—",
        },
      ];
    }

    if (outcome.priority === "cost") {
      return [
        {
          tone: "positive",
          label: "Economia contratual",
          value: formatMoney(Math.max(0, tradeoff.extraCost), true),
        },
        {
          tone: "attention",
          label: "Patrimônio menor",
          value: formatMoney(Math.max(0, tradeoff.wealthAdvantage), true),
        },
        {
          tone: "neutral",
          label: "Quitação",
          value: recommended ? formatMonth(recommended.payoffMonth) : "—",
        },
      ];
    }

    return [
      {
        tone: "positive",
        label: "Diferença patrimonial",
        value: formatMoney(Math.max(0, tradeoff.wealthAdvantage), true),
      },
      {
        tone: "attention",
        label: "Diferença de custo",
        value: formatMoney(Math.max(0, tradeoff.extraCost), true),
      },
    ];
  }

  function buildViewModel(facts, outcome, meta) {
    const copy = outcomeCopy(facts, outcome);
    const recommendation = recommendationCopy(facts, outcome);
    const recommended = strategyById(facts, outcome.recommendedStrategyId);
    const selectedAlerts = recommended
      ? recommended.alerts
      : facts.strategies
          .filter((strategy) => strategy.isEnabled)
          .flatMap((strategy) => strategy.alerts.map((alert) => ({
            ...alert,
            code: `${strategy.id}:${alert.code}`,
            text: `${strategy.short}: ${alert.text}`,
          })))
          .slice(0, 3);
    const alerts = selectedAlerts.length
      ? selectedAlerts
      : [{
          code: "confirm_assumptions",
          severity: "neutral",
          text: "A contratação depende da confirmação das premissas.",
        }];

    return {
      documentTitle: "Estratégias de Crédito",
      reportLabel: "RELATÓRIO DE AQUISIÇÃO",
      tagline: "Clareza para construir seu patrimônio.",
      clientName: meta.clientName,
      plannerName: meta.plannerName,
      generatedAt: formatDate(meta.generatedAt),
      priorityLabel: priorityLabel(outcome.priority),
      headline: copy.headline,
      summary: copy.summary,
      recommendation: {
        ...recommendation,
        strategyId: recommended ? recommended.id : null,
        strategyName: recommended ? recommended.name : null,
      },
      tradeoffSummary: tradeoffCopy(facts, outcome),
      criteriaTitle:
        facts.viableCount === 1 || (facts.leaders.wealth && facts.leaders.wealth === facts.leaders.cost)
          ? "Uma alternativa lidera nos dois critérios"
          : facts.leaders.wealth && facts.leaders.cost
            ? "Dois critérios, dois líderes"
            : facts.viableCount > 0
              ? "Resultados por critério"
              : "Comparação indisponível",
      leaders: {
        wealth: metricLeaderView(facts, "wealth"),
        cost: metricLeaderView(facts, "cost"),
      },
      charts: {
        wealth: metricRows(facts, "wealth", facts.coLeaders.wealth),
        cost: metricRows(facts, "cost", facts.coLeaders.cost),
      },
      tradeoffCards: decisionCards(facts, outcome, recommended),
      scenario: [
        {
          label: ({ Imovel: "Imóvel", Veiculo: "Veículo" })[facts.assumptions.assetType] || facts.assumptions.assetType,
          value: formatMoney(facts.assumptions.propertyValue, true),
        },
        { label: "Reserva", value: formatMoney(facts.assumptions.reserve, true) },
        { label: "Caixa mensal", value: formatMoney(facts.assumptions.monthlyCash, true) },
        {
          label: "Horizonte",
          value: isFiniteNumber(facts.assumptions.horizonMonths)
            ? formatDuration(facts.assumptions.horizonMonths)
            : "—",
        },
      ],
      methodNote: `Patrimônio corrigido com rentabilidade líquida de ${formatPercent(
        facts.assumptions.annualReturn,
      )} a.a. e inflação de ${formatPercent(facts.assumptions.inflationPct)} a.a.`,
      alerts: alerts.slice(0, 3),
      selectedStrategy: recommended
        ? {
            acquisition: formatMonth(recommended.acquisitionMonth),
            payoff: formatMonth(recommended.payoffMonth),
            wealth: formatMoney(recommended.wealth, true),
            cost: formatMoney(recommended.cost, true),
          }
        : null,
    };
  }

  function buildDecisionReport(result, meta) {
    const normalizedMeta = normalizeMeta(meta);
    const facts = extractFacts(result);
    const outcome = classifyOutcome(facts, normalizedMeta);
    const viewModel = buildViewModel(facts, outcome, normalizedMeta);

    return {
      schemaVersion: SCHEMA_VERSION,
      engineVersion: VERSION,
      meta: normalizedMeta,
      facts,
      outcome,
      viewModel,
    };
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  const REPORT_SHORT_NAMES = Object.freeze({
    fin: "Financiamento sem amort.",
    amort: "Financiamento amortizado",
    iq: "Financiamento + consórcio",
    cons: "Consórcio direto",
    card: "Carta contemplada",
  });

  function reportShortName(strategy) {
    if (!strategy) return "Revisar condições";
    return REPORT_SHORT_NAMES[strategy.id] || strategy.short || strategy.name;
  }

  function formatReportMoney(value) {
    if (!isFiniteNumber(value)) return "—";
    return value.toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL",
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    });
  }

  function formatReportDate(value) {
    if (!value) return "Data não informada";
    const plainDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    const date = plainDate
      ? new Date(Date.UTC(Number(plainDate[1]), Number(plainDate[2]) - 1, Number(plainDate[3]), 12))
      : new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", {
      day: "numeric",
      month: "long",
      year: "numeric",
      timeZone: "America/Sao_Paulo",
    }).format(date);
  }

  function formatReportDuration(months) {
    if (!isFiniteNumber(months)) return "—";
    const total = Math.max(0, Math.round(Math.abs(months)));
    const years = Math.floor(total / 12);
    const remainingMonths = total % 12;
    const parts = [];
    if (years) parts.push(`${years} ${years === 1 ? "ano" : "anos"}`);
    if (remainingMonths) parts.push(`${remainingMonths} ${remainingMonths === 1 ? "mês" : "meses"}`);
    return parts.join(" e ") || "menos de um mês";
  }

  function formatReportMonth(month) {
    if (!isFiniteNumber(month)) return "momento não informado";
    const rounded = Math.max(1, Math.round(month));
    if (rounded === 1) return "primeiro mês";
    if (rounded % 12 === 0) return `mês ${rounded}, após ${rounded / 12} ${rounded === 12 ? "ano" : "anos"}`;
    return `mês ${rounded}`;
  }

  function formatReportHorizon(months) {
    if (!isFiniteNumber(months)) return "—";
    const rounded = Math.max(0, Math.round(months));
    return rounded % 12 === 0 ? `${rounded / 12} ${rounded === 12 ? "ano" : "anos"}` : formatReportDuration(rounded);
  }

  function reportStrategyByRanking(facts, metric) {
    const id = facts.rankings[metric]?.[0];
    return strategyById(facts, id);
  }

  function reportWordmark() {
    return `
      <div class="wordmark" aria-label="Ável Planejamento">
        <strong>ável.</strong><i></i><span>Planejamento</span>
      </div>`;
  }

  function reportPageNumber(number) {
    return `<footer class="page-number">${String(number).padStart(2, "0")}</footer>`;
  }

  function reportRunningHead(section) {
    return `
      <header class="report-running-head">
        <span>Estratégias de Crédito</span>
        <span>${escapeHtml(section)}</span>
      </header>`;
  }

  function reportContext(report) {
    const facts = report.facts;
    const outcome = report.outcome;
    const recommended = strategyById(facts, outcome.recommendedStrategyId);
    const wealthLeader = reportStrategyByRanking(facts, "wealth");
    const costLeader = reportStrategyByRanking(facts, "cost");
    const hasDistinctLeaders = Boolean(wealthLeader && costLeader && wealthLeader.id !== costLeader.id);
    const wealthAdvantage = hasDistinctLeaders ? Math.max(0, wealthLeader.wealth - costLeader.wealth) : 0;
    const extraCost = hasDistinctLeaders ? Math.max(0, wealthLeader.cost - costLeader.cost) : 0;
    const payoffDifference =
      hasDistinctLeaders && isFiniteNumber(wealthLeader.payoffMonth) && isFiniteNumber(costLeader.payoffMonth)
        ? wealthLeader.payoffMonth - costLeader.payoffMonth
        : null;

    return {
      facts,
      outcome,
      recommended,
      wealthLeader,
      costLeader,
      hasDistinctLeaders,
      wealthAdvantage,
      extraCost,
      payoffDifference,
    };
  }

  function reportOverviewCopy(context) {
    const { facts, outcome, wealthLeader, costLeader, hasDistinctLeaders } = context;
    if (outcome.state === OUTCOME_STATES.INVALID) return "Os dados precisam ser revistos antes de concluir a análise.";
    if (outcome.state === OUTCOME_STATES.NO_VIABLE) return "Reserva, valor mensal ou prazos precisam ser revistos antes de indicar um caminho.";
    if (facts.viableCount === 1) {
      return `${escapeHtml(reportShortName(wealthLeader))} é a única alternativa viável nas condições informadas.`;
    }
    if (outcome.exactTieMetric) {
      return "O critério escolhido terminou empatado. Patrimônio, custo, prazo e condições da oferta devem ser lidos em conjunto.";
    }
    if (hasDistinctLeaders) {
      const payoffText =
        isFiniteNumber(costLeader.payoffMonth) && isFiniteNumber(wealthLeader.payoffMonth) && costLeader.payoffMonth < wealthLeader.payoffMonth
          ? " e encerra as parcelas antes"
          : "";
      return `${escapeHtml(wealthLeader.name)} preserva mais patrimônio; ${escapeHtml(costLeader.name)} apresenta o menor custo${payoffText}.`;
    }
    if (wealthLeader) return `${escapeHtml(wealthLeader.name)} lidera nos principais critérios analisados.`;
    return "Os critérios não apontam uma escolha única neste cenário.";
  }

  function reportVerdict(context) {
    const { facts, outcome, recommended, wealthLeader, costLeader, wealthAdvantage, extraCost } = context;
    if (!recommended) {
      if (outcome.priority === "compare" && facts.viableCount > 0) {
        return {
          title: "Os critérios estão apresentados sem uma recomendação automática.",
          text: "A decisão final depende do peso que você dá ao patrimônio, ao custo, ao prazo e ao momento da aquisição.",
          asideLabel: "Alternativas viáveis",
          asideValue: String(facts.viableCount),
          asideCaption: "caminhos sustentáveis no cenário",
        };
      }
      return {
        title: outcome.exactTieMetric ? "Há um empate no critério escolhido." : "As condições precisam ser revistas antes de indicar um caminho.",
        text: outcome.exactTieMetric
          ? "Use os demais resultados e confirme as condições das ofertas antes de decidir."
          : "Uma nova comparação deve ser feita depois dos ajustes.",
        asideLabel: "Parecer",
        asideValue: outcome.exactTieMetric ? "Empate" : "Revisar",
        asideCaption: "sem indicação exclusiva",
      };
    }

    const objectiveByPriority = {
      wealth: "preservar mais patrimônio",
      cost: "reduzir o custo contratual",
      payoff: "encerrar as parcelas antes",
      acquisition: "adquirir o bem mais cedo",
    };
    const detailByPriority = {
      wealth: `O bem fica disponível no ${formatReportMonth(recommended.acquisitionMonth)} e o patrimônio corrigido chega a ${formatReportMoney(recommended.wealth)} no período analisado.`,
      cost: `O custo contratual estimado é de ${formatReportMoney(recommended.cost)} e as parcelas se encerram em ${formatReportDuration(recommended.payoffMonth)}.`,
      payoff: `As parcelas se encerram em ${formatReportDuration(recommended.payoffMonth)}, com patrimônio corrigido de ${formatReportMoney(recommended.wealth)}.`,
      acquisition: `O bem fica disponível no ${formatReportMonth(recommended.acquisitionMonth)}, com custo contratual de ${formatReportMoney(recommended.cost)}.`,
    };
    const asideByPriority = {
      wealth: {
        label: "Vantagem patrimonial",
        value: formatReportMoney(wealthAdvantage),
        caption: costLeader ? `acima de ${reportShortName(costLeader).toLowerCase()}` : "no cenário analisado",
      },
      cost: {
        label: "Economia contratual",
        value: formatReportMoney(extraCost),
        caption: wealthLeader ? `em relação a ${reportShortName(wealthLeader).toLowerCase()}` : "no cenário analisado",
      },
      payoff: {
        label: "Parcelas encerradas em",
        value: formatReportDuration(recommended.payoffMonth),
        caption: "prazo estimado da alternativa",
      },
      acquisition: {
        label: "Bem disponível",
        value: formatReportMonth(recommended.acquisitionMonth),
        caption: "momento estimado da aquisição",
      },
    };
    const aside = asideByPriority[outcome.priority] || asideByPriority.wealth;
    return {
      title: `${escapeHtml(recommended.name)} é a escolha indicada para ${objectiveByPriority[outcome.priority] || objectiveByPriority.wealth}.`,
      text: detailByPriority[outcome.priority] || detailByPriority.wealth,
      asideLabel: aside.label,
      asideValue: aside.value,
      asideCaption: aside.caption,
    };
  }

  function reportDecisionCopy(context) {
    const { wealthLeader, costLeader, hasDistinctLeaders, wealthAdvantage, extraCost, payoffDifference } = context;
    if (!hasDistinctLeaders) {
      if (wealthLeader) return `${escapeHtml(wealthLeader.name)} lidera os critérios disponíveis. Confirme as condições finais antes de contratar.`;
      return "Revise as condições do cenário antes de decidir.";
    }
    let payoffText = "";
    if (isFiniteNumber(payoffDifference) && payoffDifference > 0) {
      payoffText = ` e mantém parcelas por ${formatReportDuration(payoffDifference)} adicionais`;
    } else if (isFiniteNumber(payoffDifference) && payoffDifference < 0) {
      payoffText = ` e encerra as parcelas ${formatReportDuration(payoffDifference)} antes`;
    }
    return `No cenário analisado, ${escapeHtml(reportShortName(wealthLeader))} termina com ${formatReportMoney(wealthAdvantage)} a mais de patrimônio, porém custa ${formatReportMoney(extraCost)} a mais${payoffText}.`;
  }

  function renderCoverPage(report) {
    return `
      <section class="report-page cover-page">
        ${reportWordmark()}
        <div class="cover-rule" aria-hidden="true"></div>
        <div class="cover-copy">
          <span class="report-kicker">Relatório de aquisição</span>
          <h1>Estratégias<br>de Crédito</h1>
          <p>${escapeHtml(report.meta.clientName)}</p>
        </div>
        <div class="cover-meta">
          <strong class="cover-date">${escapeHtml(formatReportDate(report.meta.generatedAt))}</strong>
          <div class="cover-advisor">
            <span>Planejador financeiro</span>
            <strong>${escapeHtml(report.meta.plannerName)}</strong>
          </div>
        </div>
        <p class="cover-thesis">Clareza para construir seu patrimônio.</p>
        ${reportPageNumber(1)}
      </section>`;
  }

  function renderAnalysisPage(report) {
    const context = reportContext(report);
    const verdict = reportVerdict(context);
    const { facts, wealthLeader, costLeader } = context;
    const wealthName = report.viewModel.leaders.wealth?.name || "Resultado indisponível";
    const costName = report.viewModel.leaders.cost?.name || "Resultado indisponível";
    const wealthValue = wealthLeader ? formatReportMoney(wealthLeader.wealth) : "—";
    const costValue = costLeader ? formatReportMoney(costLeader.cost) : "—";
    const wealthNote = wealthLeader ? `Bem disponível no ${formatReportMonth(wealthLeader.acquisitionMonth)}.` : "Revise as premissas.";
    const costNote = costLeader ? `Parcelas encerradas em ${formatReportDuration(costLeader.payoffMonth)}.` : "Revise as premissas.";

    return `
      <section class="report-page">
        ${reportRunningHead("Leitura do planejador")}
        <div class="title-block">
          <span>Análise</span>
          <h2>A prioridade define o melhor caminho</h2>
          <p>${reportOverviewCopy(context)}</p>
        </div>

        <section class="analysis-verdict">
          <div class="analysis-verdict-copy">
            <span>Nossa leitura</span>
            <h3>${verdict.title}</h3>
            <p>${verdict.text}</p>
          </div>
          <aside>
            <span>${escapeHtml(verdict.asideLabel)}</span>
            <strong>${escapeHtml(verdict.asideValue)}</strong>
            <small>${escapeHtml(verdict.asideCaption)}</small>
          </aside>
        </section>

        <section class="priority-grid">
          <article class="priority-card priority-wealth">
            <span>Maior patrimônio</span>
            <strong>${wealthValue}</strong>
            <h3>${escapeHtml(wealthName)}</h3>
            <p>${escapeHtml(wealthNote)}</p>
          </article>
          <article class="priority-card priority-cost">
            <span>Menor custo</span>
            <strong>${costValue}</strong>
            <h3>${escapeHtml(costName)}</h3>
            <p>${escapeHtml(costNote)}</p>
          </article>
        </section>

        <section class="planner-tradeoff">
          <strong>O ponto de decisão</strong>
          <p>${reportDecisionCopy(context)}</p>
        </section>

        <section class="scenario-line" aria-label="Premissas centrais do estudo">
          <div><span>${facts.assumptions.assetType === "Veiculo" ? "Veículo" : "Imóvel"}</span><strong>${formatReportMoney(facts.assumptions.propertyValue)}</strong></div>
          <div><span>Reserva</span><strong>${formatReportMoney(facts.assumptions.reserve)}</strong></div>
          <div><span>Valor mensal disponível</span><strong>${formatReportMoney(facts.assumptions.monthlyCash)}</strong></div>
          <div><span>Período analisado</span><strong>${formatReportHorizon(facts.assumptions.horizonMonths)}</strong></div>
        </section>
        <p class="scenario-method-note">O custo considera o contrato completo. O patrimônio foi corrigido com rentabilidade líquida de ${formatPercent(facts.assumptions.annualReturn)} a.a. e inflação de ${formatPercent(facts.assumptions.inflationPct)} a.a.</p>
        ${reportPageNumber(2)}
      </section>`;
  }

  function renderMetricComparisonRows(facts, metric, ascending, zoomDomain) {
    const strategies = facts.strategies
      .filter((strategy) => strategy.isViable && isFiniteNumber(strategy[metric]))
      .slice()
      .sort((a, b) => ascending ? a[metric] - b[metric] : b[metric] - a[metric]);
    if (!strategies.length) return '<p class="empty-comparison">Nenhuma alternativa disponível para comparação.</p>';
    const highest = Math.max(...strategies.map((strategy) => strategy[metric]), 1);
    const lowest = Math.min(...strategies.map((strategy) => strategy[metric]));
    const range = highest - lowest;
    const leaderIds = new Set(facts.coLeaders[metric] || []);

    return strategies.map((strategy) => {
      const width = zoomDomain && range > 0
        ? 14 + ((strategy[metric] - lowest) / range) * 86
        : Math.max(7, (strategy[metric] / highest) * 100);
      const isLeader = leaderIds.has(strategy.id);
      return `
        <article class="metric-comparison-row ${isLeader ? "is-criterion-leader" : ""}">
          <div class="metric-comparison-name">
            <strong>${escapeHtml(reportShortName(strategy))}</strong>
            ${isLeader ? "<small>Melhor resultado</small>" : ""}
          </div>
          <div class="metric-comparison-track"><i style="--bar-width:${width.toFixed(1)}%; --bar-color:${isLeader ? "#1dc077" : "#8e8e93"}"></i></div>
          <strong class="metric-comparison-value">${formatReportMoney(strategy[metric])}</strong>
        </article>`;
    }).join("");
  }

  function renderComparisonPage(report) {
    const context = reportContext(report);
    const { facts, wealthLeader, costLeader, hasDistinctLeaders } = context;
    const heading = hasDistinctLeaders
      ? "Dois critérios, dois líderes"
      : wealthLeader && costLeader
        ? "Uma alternativa lidera os dois critérios"
        : "Resultados por critério";
    const comparisonCopy = hasDistinctLeaders
      ? `${escapeHtml(wealthLeader.name)} encerra o período com o maior patrimônio. ${escapeHtml(costLeader.name)} apresenta o menor custo.`
      : wealthLeader
        ? `${escapeHtml(wealthLeader.name)} lidera os critérios disponíveis.`
        : "Não há alternativas viáveis para comparar.";

    return `
      <section class="report-page">
        ${reportRunningHead("Comparação")}
        <div class="title-block">
          <span>Alternativas</span>
          <h2>${heading}</h2>
          <p>${comparisonCopy}</p>
        </div>
        <div class="comparison-legend"><i></i><span>Verde identifica o melhor resultado em cada critério. Cinza mostra as demais alternativas.</span></div>
        <section class="planner-chart-panel">
          <header><div><span>Critério 01</span><h3>Patrimônio corrigido</h3></div><p>Escala ampliada entre as alternativas.</p></header>
          ${renderMetricComparisonRows(facts, "wealth", false, true)}
        </section>
        <section class="planner-chart-panel cost-chart-panel">
          <header><div><span>Critério 02</span><h3>Custo contratual</h3></div><p>Quanto menor, melhor.</p></header>
          ${renderMetricComparisonRows(facts, "cost", true, false)}
        </section>
        <section class="planner-insight">
          <strong>Leitura do planejador</strong>
          <p>A escolha depende do peso que você dá ao patrimônio, ao custo e ao prazo. O custo reúne os pagamentos do contrato; o patrimônio corrigido mostra o resultado em valores de hoje.</p>
        </section>
        ${reportPageNumber(3)}
      </section>`;
  }

  function recommendationVisual(context) {
    const { facts, outcome, recommended, wealthLeader, costLeader, hasDistinctLeaders, wealthAdvantage, extraCost, payoffDifference } = context;
    if (!recommended) {
      const isComparison = outcome.priority === "compare" && facts.viableCount > 0;
      return {
        kicker: isComparison ? "Leitura comparativa" : outcome.exactTieMetric ? "Critério empatado" : "Parecer em aberto",
        title: isComparison ? "Comparação sem recomendação" : outcome.exactTieMetric ? "Empate no critério escolhido" : "Revise as condições",
        intro: isComparison
          ? "Os resultados estão organizados para que você decida qual prioridade deve orientar a aquisição."
          : outcome.exactTieMetric
            ? "Duas ou mais alternativas entregam o mesmo resultado no critério escolhido."
            : "As premissas precisam ser revistas antes de recomendar uma alternativa.",
        cards: [
          { className: "tradeoff-gain", label: "Maior patrimônio", value: wealthLeader ? formatReportMoney(wealthLeader.wealth) : "—", text: wealthLeader ? reportShortName(wealthLeader) : "Resultado indisponível" },
          { className: "tradeoff-cost", label: "Menor custo", value: costLeader ? formatReportMoney(costLeader.cost) : "—", text: costLeader ? reportShortName(costLeader) : "Resultado indisponível" },
          { className: "tradeoff-time", label: "Parecer", value: isComparison ? "Sem recomendação" : "Empate", text: "Use os demais critérios para decidir." },
        ],
        decisionTitle: isComparison ? "A prioridade do cliente deve definir a escolha final." : "O empate pede uma leitura dos demais critérios.",
        decisionText: "Confirme custo, prazo, momento da aquisição e condições da oferta antes de seguir.",
      };
    }

    const commonMetricCards = [
      { className: "tradeoff-gain", label: "Patrimônio corrigido", value: formatReportMoney(recommended.wealth), text: "Resultado estimado em valores de hoje." },
      { className: "tradeoff-cost", label: "Custo contratual", value: formatReportMoney(recommended.cost), text: "Pagamentos previstos no contrato completo." },
      { className: "tradeoff-time", label: "Quitação", value: formatReportDuration(recommended.payoffMonth), text: "Prazo estimado das parcelas." },
    ];
    const visuals = {
      wealth: {
        kicker: "Nossa recomendação",
        title: reportShortName(recommended),
        intro: `${recommended.name} é indicada quando sua prioridade é preservar mais patrimônio sem adiar a aquisição do bem.`,
        cards: hasDistinctLeaders ? [
          { className: "tradeoff-gain", label: "Patrimônio adicional", value: formatReportMoney(wealthAdvantage), text: "Em relação à alternativa de menor custo." },
          { className: "tradeoff-cost", label: "Custo adicional", value: formatReportMoney(extraCost), text: "Valor extra para alcançar o maior patrimônio." },
          { className: "tradeoff-time", label: payoffDifference >= 0 ? "Prazo adicional" : "Prazo reduzido", value: formatReportDuration(payoffDifference), text: payoffDifference >= 0 ? "Tempo a mais de parcelas." : "Tempo a menos de parcelas." },
        ] : commonMetricCards,
        decisionTitle: hasDistinctLeaders ? "Esta escolha privilegia o patrimônio final, não o menor esforço financeiro." : "A alternativa lidera os principais critérios analisados.",
        decisionText: hasDistinctLeaders
          ? `Recomendamos ${recommended.name} se o custo adicional couber no plano sem reduzir sua reserva de segurança e se a oferta confirmar as condições analisadas. Se a prioridade for o menor custo, ${costLeader.name} é a escolha mais eficiente.`
          : `Confirme crédito, prazo, parcelas e reajustes de ${recommended.name} antes da contratação.`,
      },
      cost: {
        kicker: "Nossa recomendação",
        title: reportShortName(recommended),
        intro: `${recommended.name} é indicada quando sua prioridade é assumir o menor custo contratual.`,
        cards: hasDistinctLeaders ? [
          { className: "tradeoff-gain", label: "Economia contratual", value: formatReportMoney(extraCost), text: "Em relação à alternativa de maior patrimônio." },
          { className: "tradeoff-cost", label: "Patrimônio menor", value: formatReportMoney(wealthAdvantage), text: "Diferença patrimonial estimada ao fim do período." },
          { className: "tradeoff-time", label: "Prazo das parcelas", value: formatReportDuration(recommended.payoffMonth), text: "Tempo estimado até a quitação." },
        ] : commonMetricCards,
        decisionTitle: "Esta escolha privilegia o menor custo, não o maior patrimônio final.",
        decisionText: `Recomendamos ${recommended.name} se a prioridade for reduzir os pagamentos do contrato. Confirme a taxa, o prazo, os seguros e as condições finais da proposta.`,
      },
      payoff: {
        kicker: "Nossa recomendação",
        title: reportShortName(recommended),
        intro: `${recommended.name} é indicada quando sua prioridade é encerrar as parcelas antes.`,
        cards: commonMetricCards,
        decisionTitle: "Esta escolha privilegia a quitação antecipada.",
        decisionText: "A decisão deve considerar se os aportes mensais previstos cabem no plano sem comprometer a reserva de segurança.",
      },
      acquisition: {
        kicker: "Nossa recomendação",
        title: reportShortName(recommended),
        intro: `${recommended.name} é indicada quando sua prioridade é ter o bem disponível mais cedo.`,
        cards: [
          { className: "tradeoff-gain", label: "Bem disponível", value: formatReportMonth(recommended.acquisitionMonth), text: "Momento estimado da aquisição." },
          { className: "tradeoff-cost", label: "Custo contratual", value: formatReportMoney(recommended.cost), text: "Pagamentos previstos no contrato completo." },
          { className: "tradeoff-time", label: "Patrimônio corrigido", value: formatReportMoney(recommended.wealth), text: "Resultado estimado em valores de hoje." },
        ],
        decisionTitle: "Esta escolha privilegia o acesso ao bem.",
        decisionText: "Confirme se as condições de crédito, transferência ou contemplação sustentam o momento de aquisição usado na simulação.",
      },
    };
    return visuals[outcome.priority] || visuals.wealth;
  }

  function validationItems(strategy) {
    const byStrategy = {
      card: ["Crédito líquido e transferência", "Parcelas, prazo e reajustes", "Reserva preservada após a entrada"],
      cons: ["Regulamento e contemplação", "Lance e crédito líquido", "Reserva preservada após o lance"],
      iq: ["Coordenação dos contratos", "Contemplação e quitação", "Parcelas futuras do consórcio"],
      amort: ["Taxa e sistema de amortização", "Aportes mensais previstos", "Reserva preservada durante o plano"],
      fin: ["Taxa e sistema do financiamento", "Seguro e indexador", "Parcela compatível com o caixa"],
    };
    return byStrategy[strategy?.id] || ["Condições da proposta", "Prazo e pagamentos", "Reserva de segurança"];
  }

  function renderRecommendationPage(report) {
    const context = reportContext(report);
    const visual = recommendationVisual(context);
    const validations = validationItems(context.recommended);
    return `
      <section class="report-page dark-page">
        ${reportRunningHead("Parecer do planejador")}
        <div class="planner-reco-header">
          <span class="recommendation-kicker">${escapeHtml(visual.kicker)}</span>
          <h2>${escapeHtml(visual.title)}</h2>
          <p>${escapeHtml(visual.intro)}</p>
        </div>

        <section class="tradeoff-grid">
          ${visual.cards.map((card) => `<article class="${card.className}"><span>${escapeHtml(card.label)}</span><strong>${escapeHtml(card.value)}</strong><p>${escapeHtml(card.text)}</p></article>`).join("")}
        </section>

        <section class="decision-condition">
          <span>Ponto de decisão</span>
          <h3>${escapeHtml(visual.decisionTitle)}</h3>
          <p>${escapeHtml(visual.decisionText)}</p>
        </section>

        <section class="validation-grid">
          ${validations.map((item, index) => `<article><span>${String(index + 1).padStart(2, "0")}</span><strong>${escapeHtml(item)}</strong></article>`).join("")}
        </section>
        ${reportPageNumber(4)}
      </section>`;
  }

  function renderReportHtml(reportOrResult, meta) {
    const report =
      reportOrResult && reportOrResult.schemaVersion === SCHEMA_VERSION && reportOrResult.viewModel
        ? reportOrResult
        : buildDecisionReport(reportOrResult, meta);
    const statusClass = `cdr-status-${report.outcome.state.replace(/_/g, "-")}`;
    return `
      <article class="cdr-report report-executive ${statusClass} cdr-priority-${escapeHtml(report.outcome.priority)}" data-report-version="${escapeHtml(report.engineVersion)}">
        ${renderCoverPage(report)}
        ${renderAnalysisPage(report)}
        ${renderComparisonPage(report)}
        ${renderRecommendationPage(report)}
      </article>`;
  }

  return Object.freeze({
    VERSION,
    SCHEMA_VERSION,
    DEFAULT_PRIORITY,
    PRIORITIES,
    OUTCOME_STATES,
    normalizeMeta,
    extractFacts,
    classifyOutcome,
    buildViewModel,
    buildDecisionReport,
    renderReportHtml,
    formatMoney,
    formatDuration,
    formatMonth,
    formatPercent,
    escapeHtml,
  });
});
