import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "outputs", "consorcio_financiamento");
const outputPath = path.join(outputDir, "simulador_consorcio_interveniente_quitante_final_amortizacao_caixa_livre.xlsx");
const correctionReportPath = path.join(outputDir, "relatorio_correcoes_amortizacao_caixa_livre.md");

const FIN_MONTHS = 480;
const OPP_MAX_MONTHS = 480;
const CONS_MONTHS = OPP_MAX_MONTHS;

const currencyFmt = '"R$" #,##0;[Red]-"R$" #,##0;-';
const currencyFmt2 = '"R$" #,##0.00;[Red]-"R$" #,##0.00;-';
const pctFmt = '0.0%';
const pctFmt2 = '0.00%';
const dateFmt = 'dd/mm/yyyy';
const numFmt = '#,##0';
const monthFmt = '0';

const colors = {
  navy: "#17324D",
  blue: "#1F4E79",
  teal: "#0F766E",
  green: "#166534",
  amber: "#FEF3C7",
  paleBlue: "#EAF3F8",
  paleGreen: "#EAF7EF",
  paleGray: "#F6F7F9",
  line: "#D9E2EC",
  text: "#111827",
  inputBlue: "#0000FF",
};

const workbook = Workbook.create();
const resumo = workbook.worksheets.add("Resumo");
const premissas = workbook.worksheets.add("Premissas");
const motor = workbook.worksheets.add("Motor Caixa");
const consorcioDireto = workbook.worksheets.add("Consorcio Direto");
const financiamento = workbook.worksheets.add("Financiamento");
const consorcio = workbook.worksheets.add("Consorcio IQ");
const amortizacao = workbook.worksheets.add("Amortizacao");
const checks = workbook.worksheets.add("Checks");
const fontes = workbook.worksheets.add("Fontes");

const sheets = [resumo, premissas, motor, consorcioDireto, financiamento, consorcio, amortizacao, checks, fontes];
for (const sheet of sheets) {
  sheet.showGridLines = false;
}

function setTitle(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: colors.navy,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    borders: { preset: "outside", style: "thin", color: colors.navy },
  };
  r.format.rowHeightPx = 34;
}

function section(sheet, range, text, fill = colors.blue) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill,
    font: { bold: true, color: "#FFFFFF" },
    borders: { preset: "outside", style: "thin", color: fill },
  };
  r.format.rowHeightPx = 24;
}

function header(sheet, range, fill = colors.blue) {
  const r = sheet.getRange(range);
  r.format = {
    fill,
    font: { bold: true, color: "#FFFFFF" },
    borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  };
  r.format.wrapText = true;
  r.format.rowHeightPx = 34;
}

function frame(sheet, range, fill = "#FFFFFF") {
  const r = sheet.getRange(range);
  r.format.fill = fill;
  r.format.borders = { preset: "all", style: "thin", color: colors.line };
}

function setWidths(sheet, widths, rows = 420) {
  widths.forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, rows, 1).format.columnWidthPx = width;
  });
}

function setCellValue(sheet, cell, value, format) {
  const r = sheet.getRange(cell);
  r.values = [[value]];
  if (format) r.format.numberFormat = format;
}

function setCellFormula(sheet, cell, formula, format) {
  const r = sheet.getRange(cell);
  r.formulas = [[formula]];
  if (format) r.format.numberFormat = format;
}

function inputStyle(sheet, range) {
  const r = sheet.getRange(range);
  r.format.fill = colors.amber;
  r.format.font = { color: colors.inputBlue };
  r.format.borders = { preset: "all", style: "thin", color: "#D6B656" };
}

function formulaStyle(sheet, range) {
  const r = sheet.getRange(range);
  r.format.font = { color: colors.text };
}

function sourceLinkStyle(sheet, range) {
  const r = sheet.getRange(range);
  r.format.font = { color: "#008000" };
}

function applyNumberFormat(sheet, range, fmt) {
  sheet.getRange(range).format.numberFormat = fmt;
}

function safePreviewName(name) {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function colLetter(indexZeroBased) {
  let n = indexZeroBased + 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

// Premissas
setTitle(premissas, "A1:E1", "Simulador - Consorcio x Financiamento");
setWidths(premissas, [300, 170, 140, 360, 180], 100);
section(premissas, "A3:E3", "Premissas gerais");
section(premissas, "A12:E12", "Financiamento SAC");
section(premissas, "A20:E20", "Consorcio como interveniente quitante");
section(premissas, "A34:E34", "Amortizacao com caixa livre");
section(premissas, "A38:E38", "Custo de oportunidade");
section(premissas, "A46:E46", "Uso, disponibilidade e aluguel", colors.teal);
section(premissas, "A71:E71", "Ponte de caixa e rotulos", colors.teal);
section(premissas, "A80:E80", "Convencoes do modelo", colors.teal);

premissas.getRange("A4:E45").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("A4:A45").format.fill = colors.paleGray;
premissas.getRange("D4:D45").format.fill = colors.paleGray;
premissas.getRange("D4:D45").format.wrapText = true;

const assumptionRows = [
  [4, "Data de inicio dos pagamentos", "value", new Date(2027, 5, 1), dateFmt, "Pagamento do imovel na planta e venda do imovel atual em junho/2027."],
  [5, "Valor de venda do imovel atual", "value", 0, currencyFmt, "Entrada de caixa inicial considerada comum aos cenarios."],
  [6, "Valor do imovel novo", "value", 200000, currencyFmt, "Preco do imovel na planta."],
  [7, "Entrada (% do imovel novo)", "value", 0.2, pctFmt, "Percentual minimo de entrada do financiamento."],
  [8, "Entrada em R$", "formula", "=B6*B7", currencyFmt, "Calculado a partir do preco do imovel novo."],
  [9, "Saldo da venda apos entrada", "formula", "=B5-B8", currencyFmt, "Liquidez remanescente da venda antes de outros custos de compra."],
  [10, "Valor financiado", "formula", "=B6-B8", currencyFmt, "Base de calculo do financiamento SAC."],
  [13, "Prazo financiamento (meses)", "value", 420, monthFmt, "Prazo informado pelo usuario."],
  [14, "Juros a.a.", "value", 0.12, pctFmt, "Convertido para taxa mensal equivalente."],
  [15, "Juros mensal equivalente", "formula", "=(1+B14)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta."],
  [16, "TR a.a.", "value", 0.02, pctFmt, "Aplicada mensalmente ao saldo devedor."],
  [17, "TR mensal equivalente", "formula", "=(1+B16)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta."],
  [18, "Seguro mensal", "value", 150, currencyFmt, "Incluido enquanto houver saldo devedor no financiamento."],
  [21, "Taxa de administracao", "value", 0.24, pctFmt, "Componente editavel da taxa total do consorcio."],
  [22, "Taxa do fundo reserva", "value", 0.02, pctFmt, "Componente editavel da taxa total do consorcio."],
  [23, "Seguro prestamista (% do credito)", "value", 0.00035, pctFmt, "Tratado como percentual total sobre o credito para simplificar a simulacao."],
  [24, "Taxas totais sobre credito", "formula", "=SUM(B21:B23)", pctFmt, "Soma de administracao, fundo reserva e seguro prestamista."],
  [25, "Prazo consorcio (meses)", "value", 240, monthFmt, "Prazo informado pelo usuario."],
  [26, "Reajuste anual credito/parcela", "value", 0.06, pctFmt, "Aplicado por degrau anual."],
  [27, "Mes de contemplacao preset", "value", 60, monthFmt, "Preset editavel para a contemplacao do consorcio."],
  [28, "Fator reajuste consorcio na contemplacao", "formula", "=(1+B26)^INT((B27-1)/12)", "0.0000x", "Mes 60 recebe quatro reajustes anuais no preset, pois o 5o aniversario ocorre no mes 61."],
  [29, "Saldo financiamento no mes da contemplacao", "formula", "=IF(B27<=B13,INDEX(Financiamento!$M$6:$M$485,B27),0)", currencyFmt, "Saldo ao fim do mes, depois da prestacao normal daquele mes."],
  [30, "Carta inicial contratada", "formula", "=B29/B28", currencyFmt, "Calculada para que a carta atualizada cubra o saldo devedor no mes de contemplacao."],
  [31, "Credito atualizado na contemplacao", "formula", "=B30*B28", currencyFmt, "Valor usado como interveniente quitante."],
  [32, "Parcela inicial do consorcio", "formula", "=B30*(1+B24)/B25", currencyFmt, "Credito contratado acrescido das taxas totais, dividido pelo prazo."],
  [35, "Caixa livre mensal alocado ao financiamento", "formula", "=B42", currencyFmt, "Cenario 2 usa o caixa livre mensal informado em B42 como teto para prestacao ordinaria mais amortizacao extra enquanto houver saldo."],
  [36, "Sobra inicial para amortizacao extra", "formula", "=MAX(0,B35-INDEX(Financiamento!$J$6:$J$485,1))", currencyFmt, "Diferenca inicial entre caixa livre mensal e prestacao SAC ordinaria. No cronograma, essa sobra e recalculada mes a mes conforme a prestacao SAC cai."],
  [39, "Horizonte de comparacao (meses)", "value", 420, monthFmt, "Prazo editavel para comparar patrimonio acumulado. O modelo mensal suporta ate 480 meses."],
  [40, "Rentabilidade liquida a.a.", "value", 0.12, pctFmt, "Preset de 12% liquido ao ano para investimento das sobras mensais."],
  [41, "Rentabilidade liquida mensal equivalente", "formula", "=(1+B40)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta usada na evolucao dos investimentos."],
  [42, "Caixa livre mensal antes dos fluxos do imovel", "value", 7000, currencyFmt, "Caixa mensal disponivel antes de entrada, parcelas, aluguel pago, aluguel recebido, consorcio e lances."],
  [43, "Valor em reserva", "value", 90000, currencyFmt, "Reserva inicial disponivel. Deve ser pelo menos a maior entrada exigida entre as estrategias."],
  [44, "Maior entrada entre estrategias", "formula", "=MAX(B8,'Consorcio Direto'!$B$44)", currencyFmt, "Check de reserva minima para comparar estrategias com e sem entrada de forma justa."],
  [45, "Valor do imovel no horizonte", "formula", "=B6*(1+B66)^(B39/12)", currencyFmt, "Usado no patrimonio liquido. A valorizacao fica zerada por padrao, mas e editavel no bloco de rotulos."],
];

for (const [row, label, kind, value, fmt, note] of assumptionRows) {
  premissas.getRange(`A${row}`).values = [[label]];
  if (kind === "value") {
    setCellValue(premissas, `B${row}`, value, fmt);
  } else {
    setCellFormula(premissas, `B${row}`, value, fmt);
  }
  premissas.getRange(`D${row}`).values = [[note]];
}

inputStyle(premissas, "B4:B7");
inputStyle(premissas, "B13:B18");
inputStyle(premissas, "B21:B23");
inputStyle(premissas, "B25:B27");
inputStyle(premissas, "B39:B40");
inputStyle(premissas, "B42:B43");
inputStyle(premissas, "B43");
formulaStyle(premissas, "B8:B10");
formulaStyle(premissas, "B15:B17");
formulaStyle(premissas, "B24");
formulaStyle(premissas, "B28:B32");
formulaStyle(premissas, "B36");
formulaStyle(premissas, "B41");
formulaStyle(premissas, "B44:B45");
sourceLinkStyle(premissas, "B29:B31");
sourceLinkStyle(premissas, "B35");
sourceLinkStyle(premissas, "B44:B45");
applyNumberFormat(premissas, "B4", dateFmt);
applyNumberFormat(premissas, "B5:B6", currencyFmt);
applyNumberFormat(premissas, "B7", pctFmt);
applyNumberFormat(premissas, "B8:B10", currencyFmt);
applyNumberFormat(premissas, "B13", monthFmt);
applyNumberFormat(premissas, "B14", pctFmt);
applyNumberFormat(premissas, "B15", pctFmt2);
applyNumberFormat(premissas, "B16", pctFmt);
applyNumberFormat(premissas, "B17", pctFmt2);
applyNumberFormat(premissas, "B18", currencyFmt);
applyNumberFormat(premissas, "B21:B24", pctFmt);
applyNumberFormat(premissas, "B25", monthFmt);
applyNumberFormat(premissas, "B26", pctFmt);
applyNumberFormat(premissas, "B27", monthFmt);
applyNumberFormat(premissas, "B28", "0.0000x");
applyNumberFormat(premissas, "B29:B32", currencyFmt);
applyNumberFormat(premissas, "B35:B36", currencyFmt);
applyNumberFormat(premissas, "B39", monthFmt);
applyNumberFormat(premissas, "B40", pctFmt);
applyNumberFormat(premissas, "B41", pctFmt2);
applyNumberFormat(premissas, "B42:B45", currencyFmt);

const rentalRows = [
  [47, "Uso do imovel", "value", "Investimento para aluguel", "@", "Escolha conforme a finalidade real da decisao da mesma pessoa: moradia usa aluguel pago como custo de espera; investimento usa aluguel recebido apos acesso ao imovel."],
  [48, "Mes de aquisicao via financiamento", "value", 1, monthFmt, "Mes em que as estrategias com financiamento passam a ter o imovel como ativo."],
  [49, "Mes de entrega/chaves/disponibilidade", "value", 1, monthFmt, "Mes a partir do qual o imovel pode ser ocupado ou alugado."],
  [50, "Mes fim aluguel pago fin./IQ", "formula", "=MAX(B48,B49)-1", monthFmt, "Em moradia, financiamento e IQ pagam aluguel somente antes do maior entre aquisicao via financiamento e disponibilidade."],
  [51, "Inicio aluguel recebido fin./IQ", "formula", "=MAX(B48,B49)", monthFmt, "Em investimento, financiamento e IQ recebem aluguel somente a partir do maior entre aquisicao via financiamento e disponibilidade."],
  [52, "Inicio aluguel recebido consorcio", "formula", "=MAX('Consorcio Direto'!$B$29,B49)", monthFmt, "Em investimento, consorcio direto recebe aluguel so a partir do maior entre contemplacao direta e disponibilidade."],
  [53, "Aluguel pago (% valor imovel a.m.)", "value", 0.0035, pctFmt2, "Preset editavel para moradia propria enquanto a estrategia ainda nao permite morar no imovel comprado."],
  [54, "Aluguel pago inicial mensal", "formula", "=B6*B53", currencyFmt, "Valor mensal de aluguel pago no primeiro ano do modelo de moradia."],
  [55, "Aluguel recebido bruto (% valor a.m.)", "value", 0.004, pctFmt2, "Preset editavel para receita bruta de aluguel no modelo de investimento."],
  [56, "Aluguel recebido bruto inicial", "formula", "=B6*B55", currencyFmt, "Receita bruta mensal antes de vacancia, administracao, manutencao, impostos e IR."],
  [57, "Reajuste anual do aluguel", "value", 0.06, pctFmt, "Aplicado por degrau anual ao aluguel pago e recebido."],
  [58, "Vacancia", "value", 0.05, pctFmt, "Percentual redutor da receita bruta de aluguel."],
  [59, "Inadimplencia", "value", 0.02, pctFmt, "Percentual redutor da receita bruta de aluguel."],
  [60, "Administracao do aluguel", "value", 0.08, pctFmt, "Percentual redutor da receita bruta de aluguel."],
  [61, "Manutencao", "value", 0.03, pctFmt, "Percentual redutor da receita bruta de aluguel."],
  [62, "Impostos/condominio nao repassados", "value", 0.02, pctFmt, "Percentual redutor da receita bruta de aluguel."],
  [63, "IR sobre aluguel", "value", 0, pctFmt, "Percentual redutor, se aplicavel."],
  [64, "Aluguel recebido liquido inicial", "formula", "=B56*(1-SUM(B58:B63))", currencyFmt, "Valor liquido usado no motor apenas no modo Investimento para aluguel, apos aquisicao/contemplacao e disponibilidade."],
  [65, "Modelo nominal", "value", "Nominal", "@", "O motor permanece nominal. Valores em R$ de hoje sao apenas conversoes de apresentacao usando a inflacao esperada."],
  [66, "Valorizacao do imovel a.a.", "value", 0, pctFmt, "Preset conservador de 0% a.a.; altere se quiser valorizar o imovel no horizonte."],
  [67, "Inflacao anual esperada", "value", 0, pctFmt, "Usada apenas para converter valores nominais futuros em R$ de hoje. Nao altera o motor nominal."],
  [68, "Inflacao mensal equivalente", "formula", "=(1+B67)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta da inflacao esperada."],
  [69, "Fator inflacao acumulada no horizonte", "formula", "=(1+B67)^(B39/12)", "0.0000x", "Divide valores nominais futuros para apresentar equivalentes em R$ de hoje."],
  [72, "Caixa comum incluido", "formula", "=B43", currencyFmt, "Reserva efetivamente usada no motor mensal de caixa."],
  [73, "Venda do imovel atual", "formula", "=B5", currencyFmt, "Entrada de caixa informativa da operacao."],
  [74, "Entrada de financiamento", "formula", "=B8", currencyFmt, "Entrada que consome reserva nas estrategias com financiamento."],
  [75, "Saldo venda apos entrada", "formula", "=B9", currencyFmt, "Ponte informativa entre venda atual e entrada."],
  [76, "Caixa comum excluido do ranking", "formula", "=MAX(0,B5-B43)", currencyFmt, "Valor da venda que nao entra no motor caso exceda a reserva informada."],
  [77, "Reserva minima recomendada", "formula", "=B44", currencyFmt, "Maior entrada inicial entre as estrategias comparadas."],
];

for (const [row, label, kind, value, fmt, note] of rentalRows) {
  premissas.getRange(`A${row}`).values = [[label]];
  if (kind === "value") {
    setCellValue(premissas, `B${row}`, value, fmt);
  } else {
    setCellFormula(premissas, `B${row}`, value, fmt);
  }
  premissas.getRange(`D${row}`).values = [[note]];
}

premissas.getRange("A47:E77").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("A47:A77").format.fill = colors.paleGray;
premissas.getRange("D47:D77").format.fill = colors.paleGray;
premissas.getRange("D47:D77").format.wrapText = true;
inputStyle(premissas, "B47:B49");
inputStyle(premissas, "B53");
inputStyle(premissas, "B55");
inputStyle(premissas, "B57:B63");
inputStyle(premissas, "B65:B67");
formulaStyle(premissas, "B50:B52");
formulaStyle(premissas, "B54");
formulaStyle(premissas, "B56");
formulaStyle(premissas, "B64");
formulaStyle(premissas, "B68:B69");
formulaStyle(premissas, "B72:B77");
applyNumberFormat(premissas, "B48:B52", monthFmt);
applyNumberFormat(premissas, "B53", pctFmt2);
applyNumberFormat(premissas, "B54", currencyFmt);
applyNumberFormat(premissas, "B55", pctFmt2);
applyNumberFormat(premissas, "B56", currencyFmt);
applyNumberFormat(premissas, "B57:B63", pctFmt);
applyNumberFormat(premissas, "B64", currencyFmt);
applyNumberFormat(premissas, "B66", pctFmt);
applyNumberFormat(premissas, "B67", pctFmt);
applyNumberFormat(premissas, "B68", pctFmt2);
applyNumberFormat(premissas, "B69", "0.0000x");
applyNumberFormat(premissas, "B72:B77", currencyFmt);
premissas.getRange("B47").dataValidation = { rule: { type: "list", values: ["Moradia propria", "Investimento para aluguel"] } };
premissas.getRange("B65").dataValidation = { rule: { type: "list", values: ["Nominal"] } };

premissas.getRange("A81:E88").values = [
  ["Convencao", "Descricao", "", "", ""],
  ["SAC com TR", "TR mensal corrige o saldo no inicio do mes; juros incidem sobre o saldo corrigido; amortizacao ordinaria e calculada pelo saldo corrigido dividido pelos meses restantes.", "", "", ""],
  ["Consorcio", "Credito e parcela reajustam anualmente por degrau. A carta inicial e dimensionada para que o credito atualizado no mes da contemplacao quite o saldo do financiamento.", "", "", ""],
  ["Interveniente quitante", "No mes da contemplacao, o cliente paga a prestacao normal do financiamento e usa a carta para quitar o saldo restante daquele mes.", "", "", ""],
  ["Amortizacao de prazo", "A sobra entre o caixa livre mensal e a prestacao ordinaria SAC do proprio mes e usada como amortizacao extra direta no saldo devedor, reduzindo prazo. Como a prestacao SAC cai ao longo do tempo, a sobra mensal para amortizar tende a crescer.", "", "", ""],
  ["Motor Caixa", "O caixa livre mensal entra antes dos fluxos do imovel. Sobra vira investimento; deficit consome reserva/saldo investido; se zerar, vira capital adicional.", "", "", ""],
  ["Patrimonio bruto", "Saldo investido + valor do imovel - saldo devedor. Nao deduz capital adicional em valor futuro.", "", "", ""],
  ["Patrimonio ajustado", "Patrimonio bruto - capital adicional acumulado em valor futuro. Esta e a metrica usada no ranking economico.", "", "", ""],
];
premissas.getRange("A81:E81").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
};
premissas.getRange("A82:E88").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("B82:E88").merge(true);
premissas.getRange("B82:E88").format.wrapText = true;

// Financiamento schedule
setTitle(financiamento, "A1:M1", "Cronograma do financiamento SAC puro");
setWidths(financiamento, [58, 92, 118, 110, 118, 78, 112, 105, 88, 118, 110, 122, 118], FIN_MONTHS + 8);
financiamento.getRange("A5:M5").values = [[
  "Mes",
  "Data",
  "Saldo inicial",
  "Correcao TR",
  "Saldo corrigido",
  "Meses restantes",
  "Amortizacao",
  "Juros",
  "Seguro",
  "Prestacao total",
  "Juros + seguro",
  "Pago acumulado",
  "Saldo final",
]];
header(financiamento, "A5:M5");
financiamento.freezePanes.freezeRows(5);

const finMonths = Array.from({ length: FIN_MONTHS }, (_, i) => [i + 1]);
const finFormulas = [];
for (let i = 0; i < FIN_MONTHS; i += 1) {
  const row = 6 + i;
  const prev = row - 1;
  const month = i + 1;
  finFormulas.push([
    `=EDATE(Premissas!$B$4,${month - 1})`,
    month === 1 ? "=Premissas!$B$10" : `=M${prev}`,
    `=IF(C${row}>0,C${row}*Premissas!$B$17,0)`,
    `=C${row}+D${row}`,
    `=Premissas!$B$13-A${row}+1`,
    `=IF(E${row}>0,MIN(E${row},E${row}/F${row}),0)`,
    `=IF(E${row}>0,E${row}*Premissas!$B$15,0)`,
    `=IF(E${row}>0,Premissas!$B$18,0)`,
    `=G${row}+H${row}+I${row}`,
    `=H${row}+I${row}`,
    `=SUM($J$6:J${row})`,
    `=MAX(0,E${row}-G${row})`,
  ]);
}
financiamento.getRange(`A6:A${FIN_MONTHS + 5}`).values = finMonths;
financiamento.getRange(`B6:M${FIN_MONTHS + 5}`).formulas = finFormulas;
frame(financiamento, `A6:M${FIN_MONTHS + 5}`);
applyNumberFormat(financiamento, `A6:A${FIN_MONTHS + 5}`, monthFmt);
applyNumberFormat(financiamento, `B6:B${FIN_MONTHS + 5}`, dateFmt);
applyNumberFormat(financiamento, `C6:E${FIN_MONTHS + 5}`, currencyFmt);
applyNumberFormat(financiamento, `G6:M${FIN_MONTHS + 5}`, currencyFmt);
applyNumberFormat(financiamento, `F6:F${FIN_MONTHS + 5}`, monthFmt);
financiamento.getRange("A6:M6").format.fill = colors.paleBlue;

// Consorcio interveniente quitante schedule
setTitle(consorcio, "A1:M1", "Cronograma - consorcio como interveniente quitante");
setWidths(consorcio, [58, 92, 90, 122, 116, 125, 122, 125, 122, 118, 124, 128, 180], CONS_MONTHS + 8);
consorcio.getRange("A5:M5").values = [[
  "Mes",
  "Data",
  "Fator reajuste",
  "Credito atualizado",
  "Parcela consorcio",
  "Saldo financiamento inicio",
  "Prestacao financiamento",
  "Saldo apos prestacao",
  "Carta usada",
  "Sobra/(-falta)",
  "Desembolso cliente",
  "Pago acumulado",
  "Status",
]];
header(consorcio, "A5:M5");
consorcio.freezePanes.freezeRows(5);

const consMonths = Array.from({ length: CONS_MONTHS }, (_, i) => [i + 1]);
const consFormulas = [];
for (let i = 0; i < CONS_MONTHS; i += 1) {
  const row = 6 + i;
  const month = i + 1;
  consFormulas.push([
    `=EDATE(Premissas!$B$4,${month - 1})`,
    `=(1+Premissas!$B$26)^INT((A${row}-1)/12)`,
    `=Premissas!$B$30*C${row}`,
    `=IF(A${row}<=Premissas!$B$25,Premissas!$B$32*C${row},0)`,
    `=IF(AND(A${row}<=Premissas!$B$27,A${row}<=Premissas!$B$13),INDEX(Financiamento!$C$6:$C$485,A${row}),0)`,
    `=IF(AND(A${row}<=Premissas!$B$27,A${row}<=Premissas!$B$13),INDEX(Financiamento!$J$6:$J$485,A${row}),0)`,
    `=IF(AND(A${row}<=Premissas!$B$27,A${row}<=Premissas!$B$13),INDEX(Financiamento!$M$6:$M$485,A${row}),0)`,
    `=IF(A${row}=Premissas!$B$27,MIN(D${row},H${row}),0)`,
    `=IF(A${row}=Premissas!$B$27,D${row}-H${row},0)`,
    `=E${row}+G${row}`,
    `=SUM($K$6:K${row})`,
    `=IF(A${row}<Premissas!$B$27,"Financiamento + consorcio",IF(A${row}=Premissas!$B$27,"Quitacao via carta",IF(A${row}<=Premissas!$B$25,"Somente consorcio","Encerrado")))`,
  ]);
}
consorcio.getRange(`A6:A${CONS_MONTHS + 5}`).values = consMonths;
consorcio.getRange(`B6:M${CONS_MONTHS + 5}`).formulas = consFormulas;
frame(consorcio, `A6:M${CONS_MONTHS + 5}`);
applyNumberFormat(consorcio, `A6:A${CONS_MONTHS + 5}`, monthFmt);
applyNumberFormat(consorcio, `B6:B${CONS_MONTHS + 5}`, dateFmt);
applyNumberFormat(consorcio, `C6:C${CONS_MONTHS + 5}`, "0.0000x");
applyNumberFormat(consorcio, `D6:L${CONS_MONTHS + 5}`, currencyFmt);
consorcio.getRange("A65:M65").format.fill = colors.paleGreen;

// Amortizacao alternativa
setTitle(amortizacao, "A1:Q1", "Cronograma - amortizacao de prazo com caixa livre");
setWidths(amortizacao, [58, 92, 118, 110, 118, 122, 105, 88, 122, 122, 124, 118, 128, 120, 124, 132, 118], FIN_MONTHS + 8);
amortizacao.getRange("A5:Q5").values = [[
  "Mes",
  "Data",
  "Saldo inicial",
  "Correcao TR",
  "Saldo corrigido",
  "Amortizacao ordinaria",
  "Juros",
  "Seguro",
  "Prestacao ordinaria",
  "Amortizacao extra",
  "Desembolso total",
  "Saldo final",
  "Pago acumulado",
  "Status",
  "Caixa livre mensal",
  "Sobra para amortizar",
  "Check extra",
]];
header(amortizacao, "A5:Q5");
amortizacao.freezePanes.freezeRows(5);

const amortFormulas = [];
for (let i = 0; i < FIN_MONTHS; i += 1) {
  const row = 6 + i;
  const prev = row - 1;
  const month = i + 1;
  amortFormulas.push([
    `=EDATE(Premissas!$B$4,${month - 1})`,
    month === 1 ? "=Premissas!$B$10" : `=L${prev}`,
    `=IF(C${row}>0,C${row}*Premissas!$B$17,0)`,
    `=C${row}+D${row}`,
    `=IF(E${row}>0,MIN(E${row},INDEX(Financiamento!$G$6:$G$485,A${row})),0)`,
    `=IF(E${row}>0,E${row}*Premissas!$B$15,0)`,
    `=IF(E${row}>0,Premissas!$B$18,0)`,
    `=F${row}+G${row}+H${row}`,
    `=IF(E${row}-F${row}>0,MIN(E${row}-F${row},P${row}),0)`,
    `=I${row}+J${row}`,
    `=MAX(0,E${row}-F${row}-J${row})`,
    `=SUM($K$6:K${row})`,
    `=IF(C${row}=0,"Encerrado",IF(L${row}=0,"Quitado","Aberto"))`,
    `=IF(C${row}>0,Premissas!$B$35,0)`,
    `=IF(C${row}>0,MAX(0,O${row}-I${row}),0)`,
    `=J${row}-IF(E${row}-F${row}>0,MIN(E${row}-F${row},P${row}),0)`,
  ]);
}
amortizacao.getRange(`A6:A${FIN_MONTHS + 5}`).values = finMonths;
amortizacao.getRange(`B6:Q${FIN_MONTHS + 5}`).formulas = amortFormulas;
frame(amortizacao, `A6:Q${FIN_MONTHS + 5}`);
applyNumberFormat(amortizacao, `A6:A${FIN_MONTHS + 5}`, monthFmt);
applyNumberFormat(amortizacao, `B6:B${FIN_MONTHS + 5}`, dateFmt);
applyNumberFormat(amortizacao, `C6:M${FIN_MONTHS + 5}`, currencyFmt);
applyNumberFormat(amortizacao, `O6:Q${FIN_MONTHS + 5}`, currencyFmt);

// Consorcio direto - parametros limpos da estrategia 4
setTitle(consorcioDireto, "A1:J1", "Consorcio Direto - parametros e cronograma");
setWidths(consorcioDireto, [58, 110, 145, 130, 138, 138, 138, 138, 138, 170], OPP_MAX_MONTHS + 70);
consorcioDireto.freezePanes.freezeRows(54);

section(consorcioDireto, "A3:E3", "Inputs e parametros");
section(consorcioDireto, "A22:E22", "Consorcio e lances");
section(consorcioDireto, "G3:J3", "Checks auxiliares");

const cdInputs = [
  [4, "Valor do imovel comparado", "formula", "=Premissas!$B$6", currencyFmt, "Valor do imovel usado na estrategia de consorcio direto."],
  [5, "Data de inicio", "formula", "=Premissas!$B$4", dateFmt, "Inicio dos fluxos mensais."],
  [6, "Caixa livre mensal antes dos fluxos do imovel", "formula", "=Premissas!$B$42", currencyFmt, "Caixa mensal antes de parcelas, aluguel, lances e demais fluxos do imovel."],
  [7, "Horizonte de comparacao (meses)", "formula", "=Premissas!$B$39", monthFmt, "Puxa o horizonte das Premissas."],
  [8, "Rentabilidade liquida a.a.", "formula", "=Premissas!$B$40", pctFmt, "Puxa a rentabilidade das Premissas."],
  [9, "Rentabilidade liquida mensal", "formula", "=(1+B8)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta."],
  [23, "Taxa de administracao", "formula", "=Premissas!$B$21", pctFmt, "Componente de administracao do consorcio."],
  [24, "Taxa do fundo reserva", "formula", "=Premissas!$B$22", pctFmt, "Componente de fundo reserva do consorcio."],
  [25, "Seguro prestamista (% do credito)", "formula", "=Premissas!$B$23", pctFmt, "Seguro prestamista como percentual total sobre o credito."],
  [26, "Taxas totais sobre credito", "formula", "=SUM(B23:B25)", pctFmt, "Soma de administracao, fundo reserva e seguro prestamista."],
  [27, "Prazo consorcio (meses)", "formula", "=Premissas!$B$25", monthFmt, "Prazo de pagamento do consorcio direto."],
  [28, "Reajuste anual credito/parcela", "formula", "=Premissas!$B$26", pctFmt, "Reajuste anual por degrau."],
  [29, "Mes de contemplacao", "value", 60, monthFmt, "Mes preset em que o credito do consorcio direto fica disponivel para compra."],
  [31, "Usar lance embutido?", "value", "Sim", "@", "Se Sim, parte do lance sai do proprio credito e reduz o credito liquido."],
  [32, "Lance embutido (% credito)", "value", 0.2, pctFmt, "Percentual sobre o credito bruto atualizado."],
  [33, "Usar lance livre?", "value", "Nao", "@", "Se Sim, percentual compoe lance com recursos do cliente."],
  [34, "Lance livre (% credito)", "value", 0, pctFmt, "Percentual sobre o credito bruto atualizado."],
  [35, "Usar lance fixo?", "value", "Sim", "@", "Se Sim, percentual compoe lance com recursos do cliente."],
  [36, "Lance fixo (% credito)", "value", 0.2, pctFmt, "Percentual sobre o credito bruto atualizado."],
  [37, "Lance embutido efetivo", "formula", '=IF(B31="Sim",B32,0)', pctFmt, "Percentual efetivamente abatido do credito."],
  [38, "Lances livre/fixo efetivos", "formula", '=IF(B33="Sim",B34,0)+IF(B35="Sim",B36,0)', pctFmt, "Soma dos lances que exigem recursos do cliente antes de abater o embutido."],
  [39, "Descapitalizacao do cliente", "formula", "=MAX(0,B38-B37)", pctFmt, "Percentual do credito que sai do caixa do cliente."],
  [40, "Fator reajuste na contemplacao", "formula", "=(1+B28)^INT((B29-1)/12)", "0.0000x", "Fator acumulado ate o mes de contemplacao."],
  [41, "Carta inicial contratada", "formula", "=B4/(B40*(1-B37))", currencyFmt, "Dimensionada para que o credito liquido na contemplacao compre o imovel."],
  [42, "Credito bruto na contemplacao", "formula", "=B41*B40", currencyFmt, "Credito antes do lance embutido."],
  [43, "Credito liquido para compra", "formula", "=B42*(1-B37)", currencyFmt, "Credito disponivel para pagar o imovel depois do lance embutido."],
  [44, "Lance com recursos do cliente", "formula", "=B42*B39", currencyFmt, "Valor que sai do caixa do cliente no mes da contemplacao."],
  [45, "Parcela inicial consorcio", "formula", "=B41*(1+B26)/B27", currencyFmt, "Credito contratado acrescido das taxas, dividido pelo prazo."],
];

for (const [row, label, kind, value, fmt, note] of cdInputs) {
  consorcioDireto.getRange(`A${row}`).values = [[label]];
  if (kind === "value") {
    setCellValue(consorcioDireto, `B${row}`, value, fmt);
  } else {
    setCellFormula(consorcioDireto, `B${row}`, value, fmt);
  }
  consorcioDireto.getRange(`D${row}:E${row}`).merge();
  consorcioDireto.getRange(`D${row}`).values = [[note]];
}

frame(consorcioDireto, "A4:E9", colors.paleGray);
frame(consorcioDireto, "A23:E45", colors.paleGray);
inputStyle(consorcioDireto, "B29");
inputStyle(consorcioDireto, "B31:B36");
formulaStyle(consorcioDireto, "B4:B9");
formulaStyle(consorcioDireto, "B23:B28");
formulaStyle(consorcioDireto, "B37:B45");
consorcioDireto.getRange("D4:E45").format.wrapText = true;
consorcioDireto.getRange("B31").dataValidation = { rule: { type: "list", values: ["Sim", "Nao"] } };
consorcioDireto.getRange("B33").dataValidation = { rule: { type: "list", values: ["Sim", "Nao"] } };
consorcioDireto.getRange("B35").dataValidation = { rule: { type: "list", values: ["Sim", "Nao"] } };

consorcioDireto.getRange("G4:J8").values = [
  ["Check", "Atual", "Esperado", "Status"],
  ["Credito liquido cobre imovel", null, null, null],
  ["Mes de aquisicao/contemplacao", null, null, null],
  ["Prazo dentro do limite", null, null, null],
  ["Lance cliente no mes correto", null, null, null],
];
header(consorcioDireto, "G4:J4", colors.teal);
consorcioDireto.getRange("H5:J8").formulas = [
  ["=$B$43", "=$B$4", '=IF(H5>=I5,"OK","Revisar")'],
  ["=$B$29", "=$B$29", '=IF(H6=I6,"OK","Revisar")'],
  ["=$B$27", `=${OPP_MAX_MONTHS}`, '=IF(H7<=I7,"OK","Revisar")'],
  ["=INDEX($F$55:$F$534,$B$29)", "=$B$44", '=IF(ABS(H8-I8)<=1,"OK","Revisar")'],
];
frame(consorcioDireto, "G5:J8");
applyNumberFormat(consorcioDireto, "H5:I5", currencyFmt);
applyNumberFormat(consorcioDireto, "H6:I7", monthFmt);
applyNumberFormat(consorcioDireto, "H8:I8", currencyFmt);

const cdStartRow = 55;
const cdEndRow = cdStartRow + OPP_MAX_MONTHS - 1;
consorcioDireto.getRange("A54:J54").values = [[
  "Mes",
  "Data",
  "Caixa livre mensal",
  "Fator reajuste",
  "Parcela consorcio",
  "Lance cliente",
  "Desembolso total",
  "Credito bruto",
  "Credito liquido",
  "Status",
]];
header(consorcioDireto, "A54:J54", colors.teal);

const cdMonths = Array.from({ length: OPP_MAX_MONTHS }, (_, i) => [i + 1]);
const cdFormulas = [];
for (let i = 0; i < OPP_MAX_MONTHS; i += 1) {
  const row = cdStartRow + i;
  cdFormulas.push([
    `=EDATE($B$5,A${row}-1)`,
    "=$B$6",
    `=(1+$B$28)^INT((A${row}-1)/12)`,
    `=IF(A${row}<=$B$27,$B$45*D${row},0)`,
    `=IF(A${row}=$B$29,$B$44,0)`,
    `=E${row}+F${row}`,
    `=IF(A${row}=$B$29,$B$42,0)`,
    `=IF(A${row}=$B$29,$B$43,0)`,
    `=IF(A${row}<$B$29,"Aguardando",IF(A${row}=$B$29,"Contemplado","Pos-contemplacao"))`,
  ]);
}
consorcioDireto.getRange(`A${cdStartRow}:A${cdEndRow}`).values = cdMonths;
consorcioDireto.getRange(`B${cdStartRow}:J${cdEndRow}`).formulas = cdFormulas;
frame(consorcioDireto, `A${cdStartRow}:J${cdEndRow}`);
applyNumberFormat(consorcioDireto, `A${cdStartRow}:A${cdEndRow}`, monthFmt);
applyNumberFormat(consorcioDireto, `B${cdStartRow}:B${cdEndRow}`, dateFmt);
applyNumberFormat(consorcioDireto, `C${cdStartRow}:C${cdEndRow}`, currencyFmt);
applyNumberFormat(consorcioDireto, `D${cdStartRow}:D${cdEndRow}`, "0.0000x");
applyNumberFormat(consorcioDireto, `E${cdStartRow}:I${cdEndRow}`, currencyFmt);
consorcioDireto.getRange(`A${cdStartRow}:J${cdStartRow}`).format.fill = colors.paleGreen;

// Motor mensal corrigido de caixa e patrimonio
setTitle(motor, "A1:BS1", "Motor mensal de caixa e patrimonio liquido ajustado");
setWidths(
  motor,
  [
    58, 92, 118, 128, 142, 130, 138,
    ...Array.from({ length: 64 }, (_, i) => (i % 16 === 15 ? 118 : 122)),
  ],
  OPP_MAX_MONTHS + 8,
);
motor.freezePanes.freezeRows(5);
motor.getRange("A3:BS3").values = [[
  "Metodo",
  "A capacidade mensal e o caixa livre antes dos fluxos do imovel. Entradas e saidas do imovel entram no mes; sobra aumenta o saldo investido; deficit consome o saldo; se o saldo zerar, vira capital adicional. O patrimonio liquido deduz o capital adicional acumulado em valor futuro.",
  ...Array.from({ length: 69 }, () => ""),
]];
motor.getRange("A3").format.font = { bold: true };
motor.getRange("B3:BS3").merge();
motor.getRange("B3:BS3").format.wrapText = true;

const motorSharedHeaders = [
  "Mes",
  "Data",
  "Caixa livre mensal",
  "Aluguel pago base",
  "Aluguel recebido liquido base",
  "Valor imovel no mes",
  "FV base sem compra",
];
const motorMetricHeaders = [
  "Entrada/lance",
  "Parcela/desembolso",
  "Aluguel pago",
  "Aluguel recebido",
  "Saidas totais",
  "Entradas totais",
  "Fluxo livre",
  "Saldo antes ajuste",
  "Capital adicional mes",
  "Capital adicional VF",
  "Saldo investido",
  "Saldo devedor",
  "Patrimonio imovel",
  "Patrimonio bruto",
  "Patrimonio ajustado",
  "Status",
];
const motorStrategies = [
  {
    key: "fin",
    title: "1. Financiamento sem amortizacao",
    start: 7,
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: "Premissas!$B$13",
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF($A${row}<=Premissas!$B$13,INDEX(Financiamento!$J$6:$J$485,$A${row}),0)`,
    debtFormula: (row) => `=IF($A${row}<=Premissas!$B$13,INDEX(Financiamento!$M$6:$M$485,$A${row}),0)`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)`,
    note: "Compra no inicio via financiamento; aluguel/uso so depois da disponibilidade.",
  },
  {
    key: "amort",
    title: "2. Financiamento com amortizacao",
    start: 23,
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: 'IFERROR(MATCH("Quitado",Amortizacao!$N$6:$N$485,0),"N/A")',
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF($A${row}<=Premissas!$B$13,INDEX(Amortizacao!$K$6:$K$485,$A${row}),0)`,
    debtFormula: (row) => `=IF($A${row}<=Premissas!$B$13,INDEX(Amortizacao!$L$6:$L$485,$A${row}),0)`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)`,
    note: "Usa a sobra entre caixa livre mensal e prestacao SAC do mes para reduzir prazo.",
  },
  {
    key: "iq",
    title: "3. Financiamento + consorcio IQ",
    start: 39,
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: "Premissas!$B$27",
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF($A${row}<=Premissas!$B$25,INDEX('Consorcio IQ'!$K$6:$K$${CONS_MONTHS + 5},$A${row}),0)`,
    debtFormula: (row) => `=IF(AND($A${row}<Premissas!$B$27,$A${row}<=Premissas!$B$13),INDEX(Financiamento!$M$6:$M$485,$A${row}),0)`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)`,
    note: "Compra por financiamento e quita no mes de contemplacao do consorcio.",
  },
  {
    key: "cons",
    title: "4. Consorcio direto",
    start: 55,
    acquisitionExpr: "'Consorcio Direto'!$B$29",
    availabilityExpr: "MAX('Consorcio Direto'!$B$29,Premissas!$B$49)",
    payoffExpr: "'Consorcio Direto'!$B$29",
    entryFormula: (row) => `=IF($A${row}='Consorcio Direto'!$B$29,'Consorcio Direto'!$B$44,0)`,
    paymentFormula: (row) => `=IF($A${row}<='Consorcio Direto'!$B$27,INDEX('Consorcio Direto'!$E$55:$E$534,$A${row}),0)`,
    debtFormula: () => "=0",
    propertyFormula: (row) => `=IF($A${row}>='Consorcio Direto'!$B$29,$F${row},0)`,
    note: "Compra apenas na contemplacao; aluguel/uso so apos contemplacao e chaves.",
  },
];

motor.getRange("A4:G4").values = [motorSharedHeaders];
for (const strategy of motorStrategies) {
  const startCol = colLetter(strategy.start);
  const endCol = colLetter(strategy.start + motorMetricHeaders.length - 1);
  motor.getRange(`${startCol}4:${endCol}4`).values = [motorMetricHeaders];
  motor.getRange(`${startCol}2:${endCol}2`).merge();
  motor.getRange(`${startCol}2`).values = [[strategy.title]];
  motor.getRange(`${startCol}2:${endCol}2`).format = {
    fill: colors.navy,
    font: { bold: true, color: "#FFFFFF" },
    borders: { preset: "outside", style: "thin", color: colors.navy },
  };
}
header(motor, "A4:BS4", colors.teal);

const motorMonths = Array.from({ length: OPP_MAX_MONTHS }, (_, i) => [i + 1]);
const motorSharedFormulas = [];
for (let i = 0; i < OPP_MAX_MONTHS; i += 1) {
  const row = 5 + i;
  const prev = row - 1;
  const month = i + 1;
  motorSharedFormulas.push([
    `=EDATE(Premissas!$B$4,A${row}-1)`,
    `=IF(A${row}<=Premissas!$B$39,Premissas!$B$42,0)`,
    `=Premissas!$B$54*(1+Premissas!$B$57)^INT((A${row}-1)/12)`,
    `=Premissas!$B$64*(1+Premissas!$B$57)^INT((A${row}-1)/12)`,
    `=Premissas!$B$6*(1+Premissas!$B$66)^(A${row}/12)`,
    month === 1 ? `=Premissas!$B$43+C${row}` : `=G${prev}*(1+Premissas!$B$41)+C${row}`,
  ]);
}
motor.getRange(`A5:A${OPP_MAX_MONTHS + 4}`).values = motorMonths;
motor.getRange(`B5:G${OPP_MAX_MONTHS + 4}`).formulas = motorSharedFormulas;

for (const strategy of motorStrategies) {
  const start = strategy.start;
  const c = Array.from({ length: motorMetricHeaders.length }, (_, offset) => colLetter(start + offset));
  const formulas = [];
  for (let i = 0; i < OPP_MAX_MONTHS; i += 1) {
    const row = 5 + i;
    const prev = row - 1;
    const month = i + 1;
    formulas.push([
      strategy.entryFormula(row),
      strategy.paymentFormula(row),
      `=IF(AND(Premissas!$B$47="Moradia propria",$A${row}<${strategy.availabilityExpr}),$D${row},0)`,
      `=IF(AND(Premissas!$B$47="Investimento para aluguel",$A${row}>=${strategy.availabilityExpr}),$E${row},0)`,
      `=${c[0]}${row}+${c[1]}${row}+${c[2]}${row}`,
      `=${c[3]}${row}`,
      `=$C${row}+${c[5]}${row}-${c[4]}${row}`,
      month === 1 ? `=Premissas!$B$43+${c[6]}${row}` : `=${c[10]}${prev}*(1+Premissas!$B$41)+${c[6]}${row}`,
      `=MAX(0,-${c[7]}${row})`,
      month === 1 ? `=${c[8]}${row}` : `=${c[9]}${prev}*(1+Premissas!$B$41)+${c[8]}${row}`,
      `=MAX(0,${c[7]}${row})`,
      strategy.debtFormula(row),
      strategy.propertyFormula(row),
      `=${c[10]}${row}+${c[12]}${row}-${c[11]}${row}`,
      `=${c[13]}${row}-${c[9]}${row}`,
      `=IF(${c[9]}${row}>1,"Exige capital","OK")`,
    ]);
  }
  const startCol = colLetter(start);
  const endCol = colLetter(start + motorMetricHeaders.length - 1);
  motor.getRange(`${startCol}5:${endCol}${OPP_MAX_MONTHS + 4}`).formulas = formulas;
}

frame(motor, `A5:BS${OPP_MAX_MONTHS + 4}`);
applyNumberFormat(motor, `A5:A${OPP_MAX_MONTHS + 4}`, monthFmt);
applyNumberFormat(motor, `B5:B${OPP_MAX_MONTHS + 4}`, dateFmt);
applyNumberFormat(motor, `C5:BR${OPP_MAX_MONTHS + 4}`, currencyFmt);
motor.getRange(`BS5:BS${OPP_MAX_MONTHS + 4}`).format.wrapText = true;
motor.getRange("A5:BS5").format.fill = colors.paleGreen;

// Checks
setTitle(checks, "A1:G1", "Checks do modelo");
setWidths(checks, [440, 135, 135, 120, 95, 85, 640], 75);
checks.getRange("A4:G4").values = [["Check", "Atual", "Esperado", "Diferenca", "Tolerancia", "Status", "Notas"]];
header(checks, "A4:G4", colors.teal);
checks.getRange("A5:G30").values = [
  ["Financiamento puro termina zerado", null, null, null, null, null, "Saldo final no fim do prazo do financiamento deve ser zero."],
  ["Carta IQ cobre saldo no mes de contemplacao", null, null, null, null, null, "Credito atualizado deve bater com o saldo devedor projetado no mes de contemplacao."],
  ["Consorcio direto: credito liquido cobre imovel", null, null, null, null, null, "Credito liquido do consorcio direto deve cobrir o valor do imovel."],
  ["Consorcio direto: prazo dentro do limite", null, null, null, null, null, "Cronograma mensal do consorcio direto suporta ate 480 meses."],
  ["Reserva cobre maior entrada/lance", null, null, null, null, null, "Reserva deve cobrir a maior necessidade inicial/lance relevante entre as estrategias."],
  ["Motor: horizonte dentro do limite", null, null, null, null, null, "O motor mensal suporta ate 480 meses."],
  ["Motor: saldo investido nao fica negativo", null, null, null, null, null, "Saldo investido e truncado em zero; deficit vira capital adicional."],
  ["Motor referencia Consorcio Direto", null, null, null, null, null, "A estrategia 4 usa a aba Consorcio Direto para contemplacao, parcela e lance."],
  ["Sem formulas de retorno interno", null, null, null, null, null, "Nenhuma formula de retorno interno deve existir no arquivo decisorio."],
  ["Modelo permanece nominal", null, null, null, null, null, "Nao permitir Real sem logica real implementada."],
  ["Patrimonio bruto bate formula", null, null, null, null, null, "Patrimonio bruto = saldo investido + imovel - saldo devedor."],
  ["Patrimonio ajustado bate formula", null, null, null, null, null, "Patrimonio ajustado = patrimonio bruto - capital adicional em valor futuro."],
  ["Capital adicional separa bruto de ajustado", null, null, null, null, null, "Capital adicional VF deve explicar a diferenca entre patrimonio bruto e ajustado."],
  ["Aluguel pago somente em moradia", null, null, null, null, null, "Se o uso for investimento, aluguel pago deve ser zero."],
  ["Aluguel recebido somente em investimento", null, null, null, null, null, "Se o uso for moradia, aluguel recebido deve ser zero."],
  ["Consorcio direto sem aluguel recebido antes de contemplacao/chaves", null, null, null, null, null, "Receita de aluguel do consorcio direto so pode aparecer apos o maior entre contemplacao direta e disponibilidade."],
  ["Financiamento/IQ recebem aluguel no mes de acesso", null, null, null, null, null, "Quando o uso e investimento, estrategias 1, 2 e 3 recebem aluguel liquido no mes habilitado."],
  ["VPL e Gap usam mesma base temporal", null, null, null, null, null, "VPL deve ser o Gap VF descontado pela taxa de oportunidade."],
  ["Inflacao anual esperada nao negativa", null, null, null, null, null, "Inflacao usada apenas para apresentacao em R$ de hoje."],
  ["Fator inflacao acumulada >= 1", null, null, null, null, null, "Com inflacao esperada nao negativa, o fator deve ser no minimo 1."],
  ["Patrimonio ajustado R$ hoje bate formula", null, null, null, null, null, "Patrimonio ajustado em R$ de hoje = patrimonio ajustado nominal / fator de inflacao."],
  ["Motor nominal nao referencia inflacao", null, null, null, null, null, "Validado no build: Motor Caixa nao usa as celulas de inflacao."],
  ["VPL nominal nao referencia inflacao", null, null, null, null, null, "Validado no build: VPL vs base nao usa inflacao."],
  ["Calculos centrais sem inflacao", null, null, null, null, null, "Validado no build: financiamento, consorcio, aluguel, capital adicional e motor nao usam inflacao."],
  ["Varredura de erros de formula no build", null, null, null, null, null, "O script executa busca por erros Excel antes de salvar o arquivo."],
  ["Status geral", null, null, null, null, null, "Consolida os checks acima."],
];
checks.getRange("B5:F29").formulas = [
  ["=INDEX(Financiamento!$M$6:$M$485,Premissas!$B$13)", "=0", "=B5-C5", "=1", '=IF(ABS(D5)<=E5,"OK","Revisar")'],
  ["=Premissas!$B$31", "=Premissas!$B$29", "=B6-C6", "=1", '=IF(ABS(D6)<=E6,"OK","Revisar")'],
  ["='Consorcio Direto'!$B$43", "='Consorcio Direto'!$B$4", "=B7-C7", "=1", '=IF(B7+E7>=C7,"OK","Revisar")'],
  ["='Consorcio Direto'!$B$27", `=${OPP_MAX_MONTHS}`, "=C8-B8", "=0", '=IF(B8<=C8,"OK","Revisar")'],
  ["=Premissas!$B$43", "=Premissas!$B$44", "=B9-C9", "=0", '=IF(B9>=C9,"OK","Revisar")'],
  ["=Premissas!$B$39", `=${OPP_MAX_MONTHS}`, "=C10-B10", "=0", '=IF(B10<=C10,"OK","Revisar")'],
  ["=MIN('Motor Caixa'!$R$5:$R$484,'Motor Caixa'!$AH$5:$AH$484,'Motor Caixa'!$AX$5:$AX$484,'Motor Caixa'!$BN$5:$BN$484)", "=0", "=B11-C11", "=0", '=IF(B11>=C11,"OK","Revisar")'],
  ["=Resumo!$C$13", "='Consorcio Direto'!$B$29", "=B12-C12", "=0", '=IF(B12=C12,"OK","Revisar")'],
  ['="Removido"', '="Removido"', '=IF(B13=C13,0,1)', "=0", '=IF(B13=C13,"OK","Revisar")'],
  ["=Premissas!$B$65", '="Nominal"', '=IF(B14=C14,0,1)', "=0", '=IF(B14=C14,"OK","Revisar")'],
  ["=MAX(ABS(Resumo!$I$10-(INDEX('Motor Caixa'!$R$5:$R$484,Premissas!$B$39)+INDEX('Motor Caixa'!$T$5:$T$484,Premissas!$B$39)-INDEX('Motor Caixa'!$S$5:$S$484,Premissas!$B$39))),ABS(Resumo!$I$11-(INDEX('Motor Caixa'!$AH$5:$AH$484,Premissas!$B$39)+INDEX('Motor Caixa'!$AJ$5:$AJ$484,Premissas!$B$39)-INDEX('Motor Caixa'!$AI$5:$AI$484,Premissas!$B$39))),ABS(Resumo!$I$12-(INDEX('Motor Caixa'!$AX$5:$AX$484,Premissas!$B$39)+INDEX('Motor Caixa'!$AZ$5:$AZ$484,Premissas!$B$39)-INDEX('Motor Caixa'!$AY$5:$AY$484,Premissas!$B$39))),ABS(Resumo!$I$13-(INDEX('Motor Caixa'!$BN$5:$BN$484,Premissas!$B$39)+INDEX('Motor Caixa'!$BP$5:$BP$484,Premissas!$B$39)-INDEX('Motor Caixa'!$BO$5:$BO$484,Premissas!$B$39))))", "=0", "=B15-C15", "=1", '=IF(ABS(D15)<=E15,"OK","Revisar")'],
  ["=MAX(ABS(Resumo!$K$10-(Resumo!$I$10-Resumo!$H$10)),ABS(Resumo!$K$11-(Resumo!$I$11-Resumo!$H$11)),ABS(Resumo!$K$12-(Resumo!$I$12-Resumo!$H$12)),ABS(Resumo!$K$13-(Resumo!$I$13-Resumo!$H$13)))", "=0", "=B16-C16", "=1", '=IF(ABS(D16)<=E16,"OK","Revisar")'],
  ["=MAX(ABS((Resumo!$I$10-Resumo!$K$10)-Resumo!$H$10),ABS((Resumo!$I$11-Resumo!$K$11)-Resumo!$H$11),ABS((Resumo!$I$12-Resumo!$K$12)-Resumo!$H$12),ABS((Resumo!$I$13-Resumo!$K$13)-Resumo!$H$13))", "=0", "=B17-C17", "=1", '=IF(ABS(D17)<=E17,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Investimento para aluguel",SUM(\'Motor Caixa\'!$J$5:$J$484,\'Motor Caixa\'!$Z$5:$Z$484,\'Motor Caixa\'!$AP$5:$AP$484,\'Motor Caixa\'!$BF$5:$BF$484),0)', "=0", "=B18-C18", "=1", '=IF(ABS(D18)<=E18,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Moradia propria",SUM(\'Motor Caixa\'!$K$5:$K$484,\'Motor Caixa\'!$AA$5:$AA$484,\'Motor Caixa\'!$AQ$5:$AQ$484,\'Motor Caixa\'!$BG$5:$BG$484),0)', "=0", "=B19-C19", "=1", '=IF(ABS(D19)<=E19,"OK","Revisar")'],
  ['=SUMIFS(\'Motor Caixa\'!$BG$5:$BG$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))', "=0", "=B20-C20", "=1", '=IF(ABS(D20)<=E20,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Investimento para aluguel",MAX(ABS(INDEX(\'Motor Caixa\'!$K$5:$K$484,Premissas!$B$51)-INDEX(\'Motor Caixa\'!$E$5:$E$484,Premissas!$B$51)),ABS(INDEX(\'Motor Caixa\'!$AA$5:$AA$484,Premissas!$B$51)-INDEX(\'Motor Caixa\'!$E$5:$E$484,Premissas!$B$51)),ABS(INDEX(\'Motor Caixa\'!$AQ$5:$AQ$484,Premissas!$B$51)-INDEX(\'Motor Caixa\'!$E$5:$E$484,Premissas!$B$51))),0)', "=0", "=B21-C21", "=1", '=IF(ABS(D21)<=E21,"OK","Revisar")'],
  ["=MAX(ABS(Resumo!$M$10*(1+Premissas!$B$41)^Premissas!$B$39-Resumo!$N$10),ABS(Resumo!$M$11*(1+Premissas!$B$41)^Premissas!$B$39-Resumo!$N$11),ABS(Resumo!$M$12*(1+Premissas!$B$41)^Premissas!$B$39-Resumo!$N$12),ABS(Resumo!$M$13*(1+Premissas!$B$41)^Premissas!$B$39-Resumo!$N$13))", "=0", "=B22-C22", "=1", '=IF(ABS(D22)<=E22,"OK","Revisar")'],
  ["=Premissas!$B$67", "=0", "=B23-C23", "=0", '=IF(B23>=C23,"OK","Revisar")'],
  ["=Premissas!$B$69", "=1", "=B24-C24", "=0", '=IF(B24>=C24,"OK","Revisar")'],
  ["=MAX(ABS(Resumo!$L$10-Resumo!$K$10/Premissas!$B$69),ABS(Resumo!$L$11-Resumo!$K$11/Premissas!$B$69),ABS(Resumo!$L$12-Resumo!$K$12/Premissas!$B$69),ABS(Resumo!$L$13-Resumo!$K$13/Premissas!$B$69))", "=0", "=B25-C25", "=1", '=IF(ABS(D25)<=E25,"OK","Revisar")'],
  ['="Verificado no build"', '="Verificado no build"', '=IF(B26=C26,0,1)', "=0", '=IF(B26=C26,"OK","Revisar")'],
  ['="Verificado no build"', '="Verificado no build"', '=IF(B27=C27,0,1)', "=0", '=IF(B27=C27,"OK","Revisar")'],
  ['="Verificado no build"', '="Verificado no build"', '=IF(B28=C28,0,1)', "=0", '=IF(B28=C28,"OK","Revisar")'],
  ['="Executado"', '="Sem erros bloqueantes"', '="Ver log"', '="-"', '="OK"'],
];
checks.getRange("B30:F30").formulas = [["", "", "", "", '=IF(COUNTIF(F5:F29,"Revisar")+COUNTIF(F32:F40,"Revisar")=0,"OK","Revisar")']];
checks.getRange("A31:G31").values = [["Checks adicionais de aluguel comparavel", "", "", "", "", "", ""]];
checks.getRange("A31:G31").merge();
checks.getRange("A31:G31").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.teal },
};
checks.getRange("A32:G36").values = [
  ["Moradia: consorcio direto paga aluguel somente na espera", null, null, null, null, null, "Em moradia, a estrategia 4 deve pagar aluguel antes de MAX(contemplacao direta, disponibilidade) e parar depois."],
  ["Moradia: financiamento/IQ nao pagam aluguel apos acesso", null, null, null, null, null, "Estrategias 1, 2 e 3 devem parar aluguel pago no maior entre aquisicao via financiamento e disponibilidade."],
  ["Investimento: financiamento/IQ nao recebem aluguel antes do acesso", null, null, null, null, null, "Estrategias 1, 2 e 3 so recebem aluguel a partir de MAX(aquisicao via financiamento, disponibilidade)."],
  ["Investimento: consorcio direto recebe aluguel no mes de acesso", null, null, null, null, null, "Estrategia 4 deve receber aluguel liquido no maior entre contemplacao direta e disponibilidade."],
  ["Aluguel comparavel entre estrategias", null, null, null, null, null, "Consolida os checks de modo: moradia usa aluguel pago na espera; investimento usa aluguel recebido apos acesso."],
];
checks.getRange("B32:F36").formulas = [
  ['=IF(Premissas!$B$47="Moradia propria",MAX(ABS(SUMIFS(\'Motor Caixa\'!$BF$5:$BF$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))-SUMIFS(\'Motor Caixa\'!$D$5:$D$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))),ABS(SUMIFS(\'Motor Caixa\'!$BF$5:$BF$484,\'Motor Caixa\'!$A$5:$A$484,">="&MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49)))),0)', "=0", "=B32-C32", "=1", '=IF(ABS(D32)<=E32,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Moradia propria",SUMIFS(\'Motor Caixa\'!$J$5:$J$484,\'Motor Caixa\'!$A$5:$A$484,">="&MAX(Premissas!$B$48,Premissas!$B$49))+SUMIFS(\'Motor Caixa\'!$Z$5:$Z$484,\'Motor Caixa\'!$A$5:$A$484,">="&MAX(Premissas!$B$48,Premissas!$B$49))+SUMIFS(\'Motor Caixa\'!$AP$5:$AP$484,\'Motor Caixa\'!$A$5:$A$484,">="&MAX(Premissas!$B$48,Premissas!$B$49)),0)', "=0", "=B33-C33", "=1", '=IF(ABS(D33)<=E33,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Investimento para aluguel",SUMIFS(\'Motor Caixa\'!$K$5:$K$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(Premissas!$B$48,Premissas!$B$49))+SUMIFS(\'Motor Caixa\'!$AA$5:$AA$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(Premissas!$B$48,Premissas!$B$49))+SUMIFS(\'Motor Caixa\'!$AQ$5:$AQ$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(Premissas!$B$48,Premissas!$B$49)),0)', "=0", "=B34-C34", "=1", '=IF(ABS(D34)<=E34,"OK","Revisar")'],
  ['=IF(Premissas!$B$47="Investimento para aluguel",ABS(INDEX(\'Motor Caixa\'!$BG$5:$BG$484,MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))-INDEX(\'Motor Caixa\'!$E$5:$E$484,MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))),0)', "=0", "=B35-C35", "=1", '=IF(ABS(D35)<=E35,"OK","Revisar")'],
  ["=MAX(ABS(D18),ABS(D19),ABS(D20),ABS(D21),ABS(D32),ABS(D33),ABS(D34),ABS(D35))", "=0", "=B36-C36", "=1", '=IF(ABS(D36)<=E36,"OK","Revisar")'],
];
checks.getRange("A37:G37").values = [["Checks adicionais de amortizacao por caixa livre", "", "", "", "", "", ""]];
checks.getRange("A37:G37").merge();
checks.getRange("A37:G37").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.teal },
};
checks.getRange("A38:G40").values = [
  ["Amortizacao extra = menor entre sobra de caixa e saldo amortizavel", null, null, null, null, null, "Valida J = MIN(sobra mensal, saldo apos amortizacao ordinaria)."],
  ["Sobra para amortizar nunca negativa", null, null, null, null, null, "Sobra = MAX(0, caixa livre - prestacao SAC do mes)."],
  ["Caixa livre do cenario 2 vem de Premissas!B42", null, null, null, null, null, "Extra nao usa consorcio; usa o caixa livre mensal comum."],
];
checks.getRange("B38:F40").formulas = [
  ["=MAX(MAX(Amortizacao!$Q$6:$Q$485),-MIN(Amortizacao!$Q$6:$Q$485))", "=0", "=B38-C38", "=1", '=IF(ABS(D38)<=E38,"OK","Revisar")'],
  ["=MIN(Amortizacao!$P$6:$P$485)", "=0", "=B39-C39", "=0", '=IF(B39>=C39,"OK","Revisar")'],
  ["=Premissas!$B$35", "=Premissas!$B$42", "=B40-C40", "=1", '=IF(ABS(D40)<=E40,"OK","Revisar")'],
];
frame(checks, "A5:G40");
applyNumberFormat(checks, "B5:E7", currencyFmt);
applyNumberFormat(checks, "B8:E8", monthFmt);
applyNumberFormat(checks, "B9:E9", currencyFmt);
applyNumberFormat(checks, "B10:E10", monthFmt);
applyNumberFormat(checks, "B11:E11", currencyFmt);
applyNumberFormat(checks, "B12:E12", monthFmt);
applyNumberFormat(checks, "B15:E22", currencyFmt);
applyNumberFormat(checks, "B23:E23", pctFmt);
applyNumberFormat(checks, "B24:E24", "0.0000x");
applyNumberFormat(checks, "B25:E25", currencyFmt);
applyNumberFormat(checks, "B32:E36", currencyFmt);
applyNumberFormat(checks, "B38:E40", currencyFmt);
checks.getRange("A5:A40").format.wrapText = true;
checks.getRange("G5:G40").format.wrapText = true;

// Fontes / audit trail
setTitle(fontes, "A1:E1", "Fontes e trilha de auditoria");
setWidths(fontes, [270, 150, 110, 260, 430], 40);
fontes.getRange("A4:E4").values = [["Item", "Valor", "Unidade", "Fonte", "Observacoes"]];
header(fontes, "A4:E4", colors.teal);
const sourceRows = [
  ["Valor de venda do imovel atual", "=Premissas!$B$5", "R$", "Premissa informada pelo usuario", "Usado apenas como liquidez inicial comum aos cenarios."],
  ["Valor do imovel novo", "=Premissas!$B$6", "R$", "Premissa informada pelo usuario", "Base de entrada e financiamento."],
  ["Entrada", "=Premissas!$B$7", "%", "Premissa informada pelo usuario", "20% do valor do imovel novo."],
  ["Juros financiamento", "=Premissas!$B$14", "% a.a.", "Premissa informada pelo usuario", "Convertido para taxa mensal equivalente."],
  ["TR", "=Premissas!$B$16", "% a.a.", "Premissa informada pelo usuario", "Convertida para taxa mensal equivalente."],
  ["Seguro mensal", "=Premissas!$B$18", "R$/mes", "Premissa informada pelo usuario", "Aplicado enquanto houver saldo devedor."],
  ["Taxa administracao consorcio", "=Premissas!$B$21", "% credito", "Premissa informada pelo usuario", "Componente da taxa total do consorcio."],
  ["Fundo reserva consorcio", "=Premissas!$B$22", "% credito", "Premissa informada pelo usuario", "Componente da taxa total do consorcio."],
  ["Seguro prestamista consorcio", "=Premissas!$B$23", "% credito", "Premissa informada pelo usuario", "Componente da taxa total do consorcio."],
  ["Taxas totais do consorcio", "=Premissas!$B$24", "% credito", "Calculo do modelo", "Soma dos componentes discriminados."],
  ["Prazo do consorcio", "=Premissas!$B$25", "meses", "Premissa informada pelo usuario", "Parcelas pagas ate o prazo informado, dentro do limite fisico de 480 meses."],
  ["Reajuste consorcio", "=Premissas!$B$26", "% a.a.", "Premissa informada pelo usuario", "Credito e parcela reajustados por degrau anual."],
  ["Mes de contemplacao", "=Premissas!$B$27", "mes", "Preset do modelo", "Editavel em Premissas."],
  ["Horizonte custo oportunidade", "=Premissas!$B$39", "meses", "Preset do modelo", "Editavel em Premissas. A aba mensal suporta ate 480 meses."],
  ["Rentabilidade liquida", "=Premissas!$B$40", "% a.a.", "Preset do modelo", "Editavel em Premissas. Convertida para taxa mensal equivalente."],
  ["Caixa livre mensal antes dos fluxos do imovel", "=Premissas!$B$42", "R$/mes", "Premissa informada pelo usuario", "Caixa livre mensal antes de entrada, parcelas, aluguel pago/recebido, consorcio e lances."],
  ["Valor em reserva", "=Premissas!$B$43", "R$", "Premissa informada pelo usuario", "Deve ser no minimo a maior entrada entre as estrategias."],
  ["Maior entrada entre estrategias", "=Premissas!$B$44", "R$", "Calculo do modelo", "Usado para checar se a reserva inicial e suficiente."],
  ["Mes de contemplacao consorcio direto", "='Consorcio Direto'!$B$29", "mes", "Consorcio Direto", "Mes em que o credito fica disponivel na estrategia 4."],
  ["Carta inicial consorcio direto", "='Consorcio Direto'!$B$41", "R$", "Consorcio Direto", "Dimensionada para que o credito liquido cubra o valor do imovel."],
  ["Credito liquido consorcio direto", "='Consorcio Direto'!$B$43", "R$", "Consorcio Direto", "Credito recebido apos lance embutido."],
  ["Lance com recursos do cliente consorcio direto", "='Consorcio Direto'!$B$44", "R$", "Consorcio Direto", "Descapitalizacao efetiva do cliente quando ha lances combinados."],
  ["Parcela inicial consorcio direto", "='Consorcio Direto'!$B$45", "R$/mes", "Consorcio Direto", "Primeira parcela antes dos reajustes anuais."],
  ["Valor do imovel no horizonte", "=Premissas!$B$45", "R$", "Calculo do modelo", "Considera a valorizacao editavel, zerada por padrao."],
  ["Uso do imovel", "=Premissas!$B$47", "texto", "Premissa editavel", "Moradia propria ou investimento para aluguel."],
  ["Mes aquisicao financiamento", "=Premissas!$B$48", "mes", "Premissa editavel", "Financiamento e IQ compram no inicio por padrao."],
  ["Mes disponibilidade/chaves", "=Premissas!$B$49", "mes", "Premissa editavel", "Controla quando moradia ou aluguel podem ocorrer."],
  ["Aluguel pago inicial", "=Premissas!$B$54", "R$/mes", "Calculo do modelo", "Moradia propria: custo de espera ate a estrategia permitir morar no imovel."],
  ["Aluguel recebido liquido inicial", "=Premissas!$B$64", "R$/mes", "Calculo do modelo", "Investimento para aluguel: receita apenas apos a estrategia ter o imovel e ele estar disponivel."],
  ["Modelo nominal", "=Premissas!$B$65", "texto", "Premissa editavel", "A validacao permite apenas Nominal enquanto nao houver logica real."],
  ["Valorizacao do imovel", "=Premissas!$B$66", "% a.a.", "Premissa editavel", "Zerada por padrao para manter postura conservadora."],
  ["Inflacao anual esperada", "=Premissas!$B$67", "% a.a.", "Premissa editavel", "Usada apenas para converter valores nominais futuros em R$ de hoje."],
  ["Inflacao mensal equivalente", "=Premissas!$B$68", "% a.m.", "Calculo do modelo", "Equivalente mensal composto da inflacao esperada."],
  ["Fator inflacao acumulada no horizonte", "=Premissas!$B$69", "x", "Calculo do modelo", "Fator usado apenas nas colunas de apresentacao do Resumo."],
  ["Observacao sobre inflacao", '=Premissas!$D$67', "texto", "Convencao do modelo", "A inflacao nao altera o Motor Caixa nem as logicas de financiamento, consorcio, aluguel, capital adicional, VPL ou Gap."],
  ["FV base sem compra", "=INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)", "R$", "Motor Caixa", "Reserva e capacidade mensal investidas a taxa de oportunidade."],
];
const sourceEndRow = 4 + sourceRows.length;
fontes.getRange(`A5:A${sourceEndRow}`).values = sourceRows.map((r) => [r[0]]);
fontes.getRange(`B5:B${sourceEndRow}`).formulas = sourceRows.map((r) => [r[1]]);
fontes.getRange(`C5:E${sourceEndRow}`).values = sourceRows.map((r) => [r[2], r[3], r[4]]);
frame(fontes, `A5:E${sourceEndRow}`);
sourceRows.forEach((row, index) => {
  const unit = row[2];
  const cell = `B${index + 5}`;
  if (unit.includes("%")) applyNumberFormat(fontes, cell, pctFmt);
  if (unit === "R$" || unit === "R$/mes") applyNumberFormat(fontes, cell, currencyFmt);
  if (unit === "mes" || unit === "meses") applyNumberFormat(fontes, cell, monthFmt);
  if (unit === "x") applyNumberFormat(fontes, cell, "0.0000x");
});
fontes.getRange(`D5:E${sourceEndRow}`).format.wrapText = true;

// Resumo
setTitle(resumo, "A1:R1", "Resumo executivo - comparativo de estrategias");
setWidths(resumo, [250, 132, 95, 95, 138, 138, 138, 138, 148, 148, 148, 148, 138, 138, 132, 132, 118, 265, 24, 150, 150, 150, 150, 150, 150, 150], 90);

resumo.getRange("A3:N6").values = [
  ["Checks do modelo", null, "", "Uso do imovel", null, "", "Horizonte", null, "", "Caixa livre mensal", null, "", "Rentabilidade a.a.", null],
  ["Valor do imovel", null, "", "Disponibilidade/chaves", null, "", "Valor em reserva", null, "", "Maior entrada", null, "", "FV base sem compra", null],
  ["Entrada financiamento", null, "", "Mes contemplacao IQ", null, "", "Aluguel pago inicial", null, "", "Aluguel liquido inicial", null, "", "Valorizacao imovel a.a.", null],
  ["Modelo", null, "", "Taxas consorcio", null, "", "Carta IQ atualizada", null, "", "Parcela consorcio IQ inicial", null, "", "Prazo consorcio", null],
];
resumo.getRange("B3:B6").formulas = [
  ["=Checks!$F$30"],
  ["=Premissas!$B$6"],
  ["=Premissas!$B$8"],
  ["=Premissas!$B$65"],
];
resumo.getRange("E3:E6").formulas = [
  ["=Premissas!$B$47"],
  ["=Premissas!$B$49"],
  ["=Premissas!$B$27"],
  ["=Premissas!$B$24"],
];
resumo.getRange("H3:H6").formulas = [
  ["=Premissas!$B$39"],
  ["=Premissas!$B$43"],
  ["=Premissas!$B$54"],
  ["=Premissas!$B$31"],
];
resumo.getRange("K3:K6").formulas = [
  ["=Premissas!$B$42"],
  ["=Premissas!$B$44"],
  ["=Premissas!$B$64"],
  ["=Premissas!$B$32"],
];
resumo.getRange("N3:N6").formulas = [
  ["=Premissas!$B$40"],
  ["=INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)"],
  ["=Premissas!$B$66"],
  ["=Premissas!$B$25"],
];
frame(resumo, "A3:N6", colors.paleGray);
for (const col of ["A", "D", "G", "J", "M"]) {
  resumo.getRange(`${col}3:${col}6`).format.font = { bold: true };
}
applyNumberFormat(resumo, "B4:B5", currencyFmt);
applyNumberFormat(resumo, "E4:E5", monthFmt);
applyNumberFormat(resumo, "E6", pctFmt);
applyNumberFormat(resumo, "H3", monthFmt);
applyNumberFormat(resumo, "H4:H6", currencyFmt);
applyNumberFormat(resumo, "K3:K6", currencyFmt);
applyNumberFormat(resumo, "N3", pctFmt);
applyNumberFormat(resumo, "N4", currencyFmt);
applyNumberFormat(resumo, "N5", pctFmt);
applyNumberFormat(resumo, "N6", monthFmt);

section(resumo, "A8:R8", "Comparativo principal");
resumo.getRange("A9:R9").values = [[
  "Estrategia",
  "Entrada/lance inicial",
  "Mes aquis./cont.",
  "Mes quit./IQ",
  "Saidas nominais totais",
  "Entradas nominais totais",
  "Fluxo liquido nominal acum.",
  "Capital adicional VF",
  "Patrimonio bruto nominal",
  "Patrimonio bruto em R$ de hoje",
  "Patrimonio ajustado nominal",
  "Patrimonio ajustado em R$ de hoje",
  "VPL vs base",
  "Gap VF vs base",
  "Saldo investido",
  "Saldo devedor",
  "Status",
  "Observacao",
]];
header(resumo, "A9:R9");
resumo.getRange("A10:A13").values = [
  ["1. Financiamento sem amortizacao"],
  ["2. Financiamento com amortizacao"],
  ["3. Financiamento + consorcio IQ"],
  ["4. Consorcio"],
];
resumo.getRange("B10:R13").formulas = [
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    "=Premissas!$B$13",
    "=SUMIFS('Motor Caixa'!$L$5:$L$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$M$5:$M$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F10-E10",
    "=INDEX('Motor Caixa'!$Q$5:$Q$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$U$5:$U$484,Premissas!$B$39)",
    "=I10/Premissas!$B$69",
    "=INDEX('Motor Caixa'!$V$5:$V$484,Premissas!$B$39)",
    "=K10/Premissas!$B$69",
    "=(K10-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39",
    "=K10-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$R$5:$R$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$S$5:$S$484,Premissas!$B$39)",
    '=IF(H10>1,"Exige capital","OK")',
    '"Compra por financiamento; aluguel depende do uso e das chaves."',
  ],
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    '=MATCH("Quitado",Amortizacao!$N$6:$N$485,0)',
    "=SUMIFS('Motor Caixa'!$AB$5:$AB$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$AC$5:$AC$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F11-E11",
    "=INDEX('Motor Caixa'!$AG$5:$AG$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AK$5:$AK$484,Premissas!$B$39)",
    "=I11/Premissas!$B$69",
    "=INDEX('Motor Caixa'!$AL$5:$AL$484,Premissas!$B$39)",
    "=K11/Premissas!$B$69",
    "=(K11-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39",
    "=K11-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AH$5:$AH$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AI$5:$AI$484,Premissas!$B$39)",
    '=IF(H11>1,"Exige capital","OK")',
    '"Pagamento extra usa a sobra entre caixa livre mensal e prestacao SAC do mes."',
  ],
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    "=Premissas!$B$27",
    "=SUMIFS('Motor Caixa'!$AR$5:$AR$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$AS$5:$AS$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F12-E12",
    "=INDEX('Motor Caixa'!$AW$5:$AW$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$BA$5:$BA$484,Premissas!$B$39)",
    "=I12/Premissas!$B$69",
    "=INDEX('Motor Caixa'!$BB$5:$BB$484,Premissas!$B$39)",
    "=K12/Premissas!$B$69",
    "=(K12-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39",
    "=K12-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AX$5:$AX$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AY$5:$AY$484,Premissas!$B$39)",
    '=IF(H12>1,"Exige capital","OK")',
    '"Financiamento compra no inicio; carta quita no mes de contemplacao."',
  ],
  [
    "=0",
    "='Consorcio Direto'!$B$29",
    "='Consorcio Direto'!$B$29",
    "=SUMIFS('Motor Caixa'!$BH$5:$BH$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$BI$5:$BI$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F13-E13",
    "=INDEX('Motor Caixa'!$BM$5:$BM$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$BQ$5:$BQ$484,Premissas!$B$39)",
    "=I13/Premissas!$B$69",
    "=INDEX('Motor Caixa'!$BR$5:$BR$484,Premissas!$B$39)",
    "=K13/Premissas!$B$69",
    "=(K13-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39",
    "=K13-INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$BN$5:$BN$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$BO$5:$BO$484,Premissas!$B$39)",
    '=IF(H13>1,"Exige capital","OK")',
    '"Sem entrada inicial; compra e uso/aluguel so apos contemplacao e chaves."',
  ],
];
frame(resumo, "A10:R13");
applyNumberFormat(resumo, "B10:B13", currencyFmt);
applyNumberFormat(resumo, "C10:D13", monthFmt);
applyNumberFormat(resumo, "E10:P13", currencyFmt);
resumo.getRange("K10:L13").format.fill = colors.paleGreen;
resumo.getRange("J10:J13").format.fill = colors.paleBlue;
resumo.getRange("M10:N13").format.fill = colors.paleBlue;
resumo.getRange("Q10:R13").format.wrapText = true;

section(resumo, "A16:C16", "Grafico patrimonio corrigido", colors.teal);
resumo.getRange("A17:C21").values = [
  ["Estrategia", "Ajustado corrigido", "Dif. vs melhor"],
  ["Fin. sem amort.", null],
  ["Fin. amort.", null],
  ["Fin. + IQ", null],
  ["Consorcio", null],
];
resumo.getRange("B18:C21").formulas = [
  ["=L10", "=B18-MAX($B$18:$B$21)"],
  ["=L11", "=B19-MAX($B$18:$B$21)"],
  ["=L12", "=B20-MAX($B$18:$B$21)"],
  ["=L13", "=B21-MAX($B$18:$B$21)"],
];
header(resumo, "A17:C17", colors.teal);
frame(resumo, "A18:C21");
resumo.getRange("B18:B21").format.fill = colors.paleGreen;
resumo.getRange("C18:C21").format.fill = colors.paleBlue;
applyNumberFormat(resumo, "B18:C21", currencyFmt);

const correctedChart = resumo.charts.add("bar", resumo.getRange("A17:B21"));
correctedChart.title = "Patrimonio ajustado em R$ de hoje por estrategia";
correctedChart.hasLegend = false;
correctedChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
correctedChart.yAxis = { numberFormatCode: 'R$ #,##0' };
correctedChart.series.items[0].fill = colors.teal;
correctedChart.setPosition("T3", "Z24");

resumo.getRange("A24:R39").values = [
  ["Definicoes", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Saidas nominais totais", "Soma bruta das saidas do motor: entrada, parcelas, seguros, consorcio, lances e aluguel pago quando aplicavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Entradas nominais totais", "Soma bruta das entradas do motor, principalmente aluguel recebido liquido no modo investimento.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Fluxo liquido nominal acum.", "Entradas nominais totais menos saidas nominais totais. Nao substitui VPL nem Gap.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio bruto nominal", "Saldo investido + valor do imovel - saldo devedor, no horizonte, em base nominal.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio bruto em R$ de hoje", "Patrimonio bruto nominal dividido pelo fator de inflacao acumulada no horizonte.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio ajustado nominal", "Patrimonio bruto nominal - capital adicional em valor futuro, sem deflacionar o motor.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio ajustado em R$ de hoje", "Patrimonio ajustado nominal dividido pela inflacao acumulada no horizonte. E somente uma conversao de apresentacao.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Inflacao esperada", "Premissa editavel usada apenas para converter valores nominais futuros em R$ de hoje. Nao altera financiamento, consorcio, aluguel, capital adicional ou motor.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Modelo nominal", "Financiamento, consorcio, aluguel, valorizacao e taxa de oportunidade permanecem em base nominal.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Gap VF vs base", "Patrimonio ajustado nominal menos o valor futuro da alternativa base de investir reserva e caixa livre mensal. Positivo e melhor.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["VPL vs base", "VPL vs base ja e uma medida trazida para a data zero pela taxa de oportunidade nominal. A coluna em R$ de hoje apenas deflaciona patrimonio futuro nominal pelo fator de inflacao esperado.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Metricas principais", "Usar VPL vs base, Gap VF vs base, Patrimonio ajustado nominal, Patrimonio ajustado em R$ de hoje, Capital adicional VF e Status de capacidade.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Uso do imovel", "Escolher conforme a finalidade real da decisao da mesma pessoa. Moradia propria usa aluguel pago como custo de espera ate poder morar; investimento para aluguel usa aluguel recebido apenas apos aquisicao/contemplacao e disponibilidade.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Comparabilidade do aluguel", "Estrategias 1, 2 e 3 usam acesso pelo maior entre aquisicao via financiamento e disponibilidade. Estrategia 4 usa acesso pelo maior entre contemplacao do consorcio direto e disponibilidade.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Consorcio Direto", "Aba auxiliar limpa para parametros, lances, credito liquido e cronograma mensal da estrategia 4. Comparacao principal fica no Resumo.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];
resumo.getRange("A24:R24").merge();
resumo.getRange("A24:R24").format = {
  fill: colors.navy,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.navy },
};
resumo.getRange("B25:R38").merge(true);
frame(resumo, "A25:R38", colors.paleGray);
resumo.getRange("B25:R38").format.wrapText = true;

// Final formatting touches
for (const sheet of sheets) {
  const used = sheet.getUsedRange();
  if (used) {
    used.format.font = { name: "Calibri", size: 10, color: colors.text };
  }
}
// Reapply title/header styles after workbook-wide font normalization.
setTitle(premissas, "A1:E1", "Simulador - Consorcio x Financiamento");
setTitle(financiamento, "A1:M1", "Cronograma do financiamento SAC puro");
setTitle(consorcio, "A1:M1", "Cronograma - consorcio como interveniente quitante");
setTitle(amortizacao, "A1:Q1", "Cronograma - amortizacao de prazo com caixa livre");
setTitle(checks, "A1:G1", "Checks do modelo");
setTitle(fontes, "A1:E1", "Fontes e trilha de auditoria");
setTitle(motor, "A1:BS1", "Motor mensal de caixa e patrimonio liquido ajustado");
setTitle(consorcioDireto, "A1:J1", "Consorcio Direto - parametros e cronograma");
setTitle(resumo, "A1:R1", "Resumo executivo - comparativo de estrategias");
header(motor, "A4:BS4", colors.teal);
header(consorcioDireto, "A54:J54", colors.teal);
header(consorcioDireto, "G4:J4", colors.teal);
header(financiamento, "A5:M5");
header(consorcio, "A5:M5");
header(amortizacao, "A5:Q5");
header(checks, "A4:G4", colors.teal);
header(fontes, "A4:E4", colors.teal);
header(resumo, "A9:R9");
header(resumo, "A17:C17", colors.teal);

// Restore input/formula styling after normalization.
inputStyle(premissas, "B4:B7");
inputStyle(premissas, "B13:B18");
inputStyle(premissas, "B21:B23");
inputStyle(premissas, "B25:B27");
inputStyle(premissas, "B39:B40");
inputStyle(premissas, "B42:B43");
inputStyle(premissas, "B47:B49");
inputStyle(premissas, "B53");
inputStyle(premissas, "B55");
inputStyle(premissas, "B57:B63");
inputStyle(premissas, "B65:B67");
formulaStyle(premissas, "B8:B10");
formulaStyle(premissas, "B15:B17");
formulaStyle(premissas, "B24");
formulaStyle(premissas, "B28:B32");
formulaStyle(premissas, "B36");
formulaStyle(premissas, "B41");
formulaStyle(premissas, "B44:B45");
formulaStyle(premissas, "B50:B52");
formulaStyle(premissas, "B54");
formulaStyle(premissas, "B56");
formulaStyle(premissas, "B64");
formulaStyle(premissas, "B68:B69");
formulaStyle(premissas, "B72:B77");
sourceLinkStyle(premissas, "B29:B31");
sourceLinkStyle(premissas, "B35");
sourceLinkStyle(premissas, "B44:B45");
inputStyle(consorcioDireto, "B29");
inputStyle(consorcioDireto, "B31:B36");
formulaStyle(consorcioDireto, "B4:B9");
formulaStyle(consorcioDireto, "B23:B28");
formulaStyle(consorcioDireto, "B37:B45");

await fs.mkdir(outputDir, { recursive: true });

const finalSheetNames = [
  "Resumo",
  "Premissas",
  "Motor Caixa",
  "Consorcio Direto",
  "Financiamento",
  "Consorcio IQ",
  "Amortizacao",
  "Checks",
  "Fontes",
];
console.log(`FINAL_SHEETS ${finalSheetNames.join(" | ")}`);

const inspectSummary = await workbook.inspect({
  kind: "table",
  range: "Resumo!A8:R13",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 18,
  maxChars: 6000,
});
console.log("INSPECT_SUMMARY");
console.log(inspectSummary.ndjson);

const inspectAssumptions = await workbook.inspect({
  kind: "table",
  range: "Premissas!A47:B77",
  include: "values,formulas",
  tableMaxRows: 31,
  tableMaxCols: 2,
  maxChars: 5000,
});
console.log("INSPECT_ASSUMPTIONS");
console.log(inspectAssumptions.ndjson);

const inspectConsorcioDireto = await workbook.inspect({
  kind: "table",
  range: "Consorcio Direto!A22:J60",
  include: "values,formulas",
  tableMaxRows: 39,
  tableMaxCols: 10,
  maxChars: 7000,
});
console.log("INSPECT_CONSORCIO_DIRETO");
console.log(inspectConsorcioDireto.ndjson);

const inspectOpportunity = await workbook.inspect({
  kind: "table",
  range: "Resumo!A3:N6",
  include: "values,formulas",
  tableMaxRows: 4,
  tableMaxCols: 14,
  maxChars: 5000,
});
console.log("INSPECT_TOP_CARDS");
console.log(inspectOpportunity.ndjson);

const inspectMotor = await workbook.inspect({
  kind: "table",
  range: "Motor Caixa!A4:U10",
  include: "values,formulas",
  tableMaxRows: 7,
  tableMaxCols: 21,
  maxChars: 6000,
});
console.log("INSPECT_MOTOR");
console.log(inspectMotor.ndjson);

const inspectAmortization = await workbook.inspect({
  kind: "table",
  range: "Amortizacao!A5:Q12",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 17,
  maxChars: 7000,
});
console.log("INSPECT_AMORTIZATION");
console.log(inspectAmortization.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 5000,
});
console.log("FORMULA_ERRORS");
console.log(errors.ndjson);

const legacyReferences = await workbook.inspect({
  kind: "match",
  searchTerm: "Legado|LEGADO|LEGADA|Aquisicao Comparada|Custo Oportunidade - Legado|TIR Aux",
  options: { useRegex: true, maxResults: 300 },
  summary: "legacy reference scan",
  maxChars: 5000,
});
console.log("LEGACY_REFERENCES");
console.log(legacyReferences.ndjson);

const internalReturnReferences = await workbook.inspect({
  kind: "match",
  searchTerm: "\\bTIR\\b|IRR\\s*\\(|XIRR\\s*\\(|MIRR\\s*\\(",
  options: { useRegex: true, maxResults: 300 },
  summary: "internal return metric scan",
  maxChars: 5000,
});
console.log("INTERNAL_RETURN_REFERENCES");
console.log(internalReturnReferences.ndjson);

const inflationSearchTerm = "Premissas!\\$B\\$67|Premissas!\\$B\\$68|Premissas!\\$B\\$69|Inflacao|inflacao";
const motorInflationReferences = await workbook.inspect({
  kind: "match",
  range: "Motor Caixa!A1:BS484",
  searchTerm: inflationSearchTerm,
  options: { useRegex: true, maxResults: 300 },
  summary: "motor inflation reference scan",
  maxChars: 5000,
});
console.log("MOTOR_INFLATION_REFERENCES");
console.log(motorInflationReferences.ndjson);

const vplInflationReferences = await workbook.inspect({
  kind: "match",
  range: "Resumo!M10:M13",
  searchTerm: inflationSearchTerm,
  options: { useRegex: true, maxResults: 300 },
  summary: "vpl inflation reference scan",
  maxChars: 5000,
});
console.log("VPL_INFLATION_REFERENCES");
console.log(vplInflationReferences.ndjson);

function inspectHasMatch(inspectResult) {
  return inspectResult.ndjson
    .split(/\r?\n/)
    .filter((line) => line.trim())
    .some((line) => {
      try {
        return JSON.parse(line).kind === "match";
      } catch {
        return false;
      }
    });
}

const centralInflationReferenceScans = [];
for (const range of [
  "Financiamento!A1:M485",
  "Consorcio IQ!A1:M485",
  "Consorcio Direto!A1:J534",
  "Amortizacao!A1:Q485",
]) {
  const scan = await workbook.inspect({
    kind: "match",
    range,
    searchTerm: inflationSearchTerm,
    options: { useRegex: true, maxResults: 300 },
    summary: `central calculation inflation reference scan ${range}`,
    maxChars: 5000,
  });
  centralInflationReferenceScans.push({ range, scan });
  console.log(`CENTRAL_INFLATION_REFERENCES ${range}`);
  console.log(scan.ndjson);
}

if (inspectHasMatch(motorInflationReferences)) {
  throw new Error("Motor Caixa referencia inflacao; isso violaria a separacao nominal/real.");
}
if (inspectHasMatch(vplInflationReferences)) {
  throw new Error("VPL vs base referencia inflacao; VPL nominal deve permanecer preservado.");
}
const centralInflationHit = centralInflationReferenceScans.find(({ scan }) => inspectHasMatch(scan));
if (centralInflationHit) {
  throw new Error(`Calculos centrais referenciam inflacao em ${centralInflationHit.range}; inflacao deve ser apenas apresentacao.`);
}

function asNumber(value) {
  if (typeof value === "number") return value;
  if (value === null || value === undefined || value === "") return 0;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function cellNumber(sheet, address) {
  return asNumber(sheet.getRange(address).values[0][0]);
}

function sumRange(sheet, address) {
  return sheet
    .getRange(address)
    .values
    .flat()
    .reduce((total, value) => total + asNumber(value), 0);
}

function rangeNumbers(sheet, address) {
  return sheet.getRange(address).values.flat().map(asNumber);
}

function maxRange(sheet, address) {
  return Math.max(...rangeNumbers(sheet, address));
}

function minRange(sheet, address) {
  return Math.min(...rangeNumbers(sheet, address));
}

function maxAbsRange(sheet, address) {
  return rangeNumbers(sheet, address).reduce((maxValue, value) => Math.max(maxValue, Math.abs(value)), 0);
}

function sumMotorCol(col, startMonth, endMonth) {
  if (startMonth > endMonth) return 0;
  return sumRange(motor, `${col}${startMonth + 4}:${col}${endMonth + 4}`);
}

function sumMotorCols(cols, startMonth, endMonth) {
  return cols.reduce((total, col) => total + sumMotorCol(col, startMonth, endMonth), 0);
}

function assertNear(name, actual, expected, tolerance = 1) {
  const diff = actual - expected;
  if (Math.abs(diff) > tolerance) {
    throw new Error(`${name}: esperado ${expected}, atual ${actual}, diferenca ${diff}`);
  }
  return { name, actual, expected, diff };
}

function runAmortizationCashFreeTests() {
  const cashFree = cellNumber(premissas, "B42");
  const amortizationCash = cellNumber(premissas, "B35");
  const firstOrdinaryPayment = cellNumber(amortizacao, "I6");
  const firstExtraPayment = cellNumber(amortizacao, "J6");
  const firstRemainingAfterOrdinary = Math.max(0, cellNumber(amortizacao, "E6") - cellNumber(amortizacao, "F6"));
  const firstExpectedExtra = Math.min(firstRemainingAfterOrdinary, Math.max(0, amortizationCash - firstOrdinaryPayment));
  const maxExtraFormulaDiff = maxAbsRange(amortizacao, `Q6:Q${FIN_MONTHS + 5}`);
  const minCashSurplus = minRange(amortizacao, `P6:P${FIN_MONTHS + 5}`);
  const maxTotalDisbursement = maxRange(amortizacao, `K6:K${FIN_MONTHS + 5}`);
  const payoffIndex = amortizacao
    .getRange(`N6:N${FIN_MONTHS + 5}`)
    .values
    .flat()
    .findIndex((value) => value === "Quitado");
  const payoffMonth = payoffIndex >= 0 ? payoffIndex + 1 : null;
  const cashFormula = premissas.getRange("B35").formulas[0][0];
  const amortizationFormulaText = amortizacao
    .getRange(`J6:Q${FIN_MONTHS + 5}`)
    .formulas
    .flat()
    .join(" ");

  if (cashFormula !== "=B42") {
    throw new Error(`Premissas!B35 deve apontar para Premissas!B42, mas esta como ${cashFormula}`);
  }
  if (/Consorcio|B32|B36/i.test(amortizationFormulaText)) {
    throw new Error("Amortizacao extra voltou a depender de consorcio, B32 ou B36.");
  }
  if (minCashSurplus < -1) {
    throw new Error(`Sobra para amortizar ficou negativa: ${minCashSurplus}`);
  }

  return {
    cashFree,
    amortizationCash,
    firstOrdinaryPayment,
    firstExtraPayment,
    firstExpectedExtra,
    maxExtraFormulaDiff,
    minCashSurplus,
    maxTotalDisbursement,
    payoffMonth,
    results: [
      assertNear("Amortizacao: caixa livre do cenario 2 vem de B42", amortizationCash, cashFree),
      assertNear("Amortizacao: extra do primeiro mes usa sobra de caixa", firstExtraPayment, firstExpectedExtra),
      assertNear("Amortizacao: coluna J bate com menor entre sobra e saldo amortizavel", maxExtraFormulaDiff, 0),
    ],
  };
}

function runRentalComparabilityTests() {
  const original = {
    use: premissas.getRange("B47").values[0][0],
    financingAcquisition: premissas.getRange("B48").values[0][0],
    availability: premissas.getRange("B49").values[0][0],
    iqContemplation: premissas.getRange("B27").values[0][0],
    directContemplation: consorcioDireto.getRange("B29").values[0][0],
  };

  const paidFinIqCols = ["J", "Z", "AP"];
  const receivedFinIqCols = ["K", "AA", "AQ"];
  const paidAllCols = [...paidFinIqCols, "BF"];
  const receivedAllCols = [...receivedFinIqCols, "BG"];

  function setScenario(use) {
    premissas.getRange("B47").values = [[use]];
    premissas.getRange("B48").values = [[1]];
    premissas.getRange("B49").values = [[1]];
    premissas.getRange("B27").values = [[60]];
    consorcioDireto.getRange("B29").values = [[60]];
  }

  function restoreScenario() {
    premissas.getRange("B47").values = [[original.use]];
    premissas.getRange("B48").values = [[original.financingAcquisition]];
    premissas.getRange("B49").values = [[original.availability]];
    premissas.getRange("B27").values = [[original.iqContemplation]];
    consorcioDireto.getRange("B29").values = [[original.directContemplation]];
  }

  try {
    setScenario("Moradia propria");
    const moradiaHorizon = cellNumber(premissas, "B39");
    const moradiaFinAccess = Math.max(cellNumber(premissas, "B48"), cellNumber(premissas, "B49"));
    const moradiaDirectAccess = Math.max(cellNumber(consorcioDireto, "B29"), cellNumber(premissas, "B49"));
    const moradiaFinPaid = sumMotorCols(paidFinIqCols, 1, moradiaHorizon);
    const moradiaDirectPaidBefore = sumMotorCol("BF", 1, moradiaDirectAccess - 1);
    const moradiaDirectExpectedBefore = sumMotorCol("D", 1, moradiaDirectAccess - 1);
    const moradiaDirectPaidAfter = sumMotorCol("BF", moradiaDirectAccess, moradiaHorizon);
    const moradiaReceived = sumMotorCols(receivedAllCols, 1, moradiaHorizon);

    const moradiaResults = [
      assertNear("Moradia: estrategias 1, 2 e 3 sem aluguel pago", moradiaFinPaid, 0),
      assertNear("Moradia: consorcio direto paga aluguel na espera", moradiaDirectPaidBefore, moradiaDirectExpectedBefore),
      assertNear("Moradia: consorcio direto para aluguel no acesso", moradiaDirectPaidAfter, 0),
      assertNear("Moradia: nenhum aluguel recebido", moradiaReceived, 0),
    ];

    setScenario("Investimento para aluguel");
    const investimentoHorizon = cellNumber(premissas, "B39");
    const investimentoFinAccess = Math.max(cellNumber(premissas, "B48"), cellNumber(premissas, "B49"));
    const investimentoDirectAccess = Math.max(cellNumber(consorcioDireto, "B29"), cellNumber(premissas, "B49"));
    const investimentoPaid = sumMotorCols(paidAllCols, 1, investimentoHorizon);
    const investimentoFinExpected = sumMotorCol("E", investimentoFinAccess, investimentoHorizon);
    const investimentoDirectExpected = sumMotorCol("E", investimentoDirectAccess, investimentoHorizon);
    const investimentoDirectReceivedBefore = sumMotorCol("BG", 1, investimentoDirectAccess - 1);

    const investimentoResults = [
      assertNear("Investimento: nenhum aluguel pago", investimentoPaid, 0),
      assertNear("Investimento: estrategia 1 recebe aluguel do acesso ao horizonte", sumMotorCol("K", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: estrategia 2 recebe aluguel do acesso ao horizonte", sumMotorCol("AA", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: estrategia 3 recebe aluguel do acesso ao horizonte", sumMotorCol("AQ", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: consorcio direto sem aluguel antes do acesso", investimentoDirectReceivedBefore, 0),
      assertNear("Investimento: consorcio direto recebe aluguel do acesso ao horizonte", sumMotorCol("BG", investimentoDirectAccess, investimentoHorizon), investimentoDirectExpected),
    ];

    return {
      moradia: {
        accessFinancing: moradiaFinAccess,
        accessConsorcioDireto: moradiaDirectAccess,
        consorcioRentMonths: `${1}-${moradiaDirectAccess - 1}`,
        results: moradiaResults,
      },
      investimento: {
        accessFinancing: investimentoFinAccess,
        accessConsorcioDireto: investimentoDirectAccess,
        results: investimentoResults,
      },
    };
  } finally {
    restoreScenario();
  }
}

const rentalComparabilityTests = runRentalComparabilityTests();
const amortizationCashFreeTests = runAmortizationCashFreeTests();

await fs.writeFile(
  correctionReportPath,
  `# Relatorio de correcoes - amortizacao por caixa livre

Arquivo gerado: \`${path.basename(outputPath)}\`
Data: 2026-06-21

## Correcoes implementadas

- Corrigida a logica do cenario \`Financiamento com amortizacao\`.
- A amortizacao extra deixou de usar a parcela do consorcio como proxy.
- A nova regra usa apenas a diferenca mensal entre \`caixa livre mensal\` e \`prestacao ordinaria SAC do mes\`.
- Como a prestacao SAC diminui ao longo do tempo, a sobra para amortizacao extra e recalculada mes a mes na aba \`Amortizacao\`.
- A amortizacao extra e limitada ao saldo remanescente apos a amortizacao ordinaria, para evitar saldo devedor negativo.
- Foram adicionadas colunas auxiliares em \`Amortizacao!O:Q\`: caixa livre mensal, sobra para amortizar e check da amortizacao extra.
- \`Premissas!B35\` agora referencia \`Premissas!B42\`; \`Premissas!B36\` mostra apenas a sobra inicial para amortizacao extra.
- Corrigida e reforcada a logica de aluguel comparavel para a mesma pessoa escolhendo uma entre quatro alternativas.
- Mantidos dois modos em \`Premissas!B47\`: \`Moradia propria\` e \`Investimento para aluguel\`.
- Em \`Moradia propria\`, aluguel recebido e zero; aluguel pago entra apenas enquanto a estrategia ainda nao permite morar no imovel comprado.
- Em \`Investimento para aluguel\`, aluguel pago e zero; aluguel recebido liquido entra apenas apos a estrategia ter o imovel e ele estar disponivel.
- Estrategias 1, 2 e 3 usam acesso pelo maior entre aquisicao via financiamento e disponibilidade.
- Estrategia 4 usa acesso pelo maior entre contemplacao do consorcio direto e disponibilidade.
- O check \`Financiamento puro termina zerado\` agora valida o saldo no fim do prazo informado em \`Premissas\`, sem nota fixa de mes 360.
- Adicionados checks especificos de aluguel comparavel em \`Checks\`, mantendo \`Checks!F30\` como status geral.
- Preservada a separacao nominal/real: inflacao continua apenas como apresentacao, e nao altera Motor Caixa, aluguel, financiamento, consorcio, capital adicional, Gap VF ou VPL.

## Validacoes executadas

- Inspecao do \`Resumo\`, \`Premissas\`, \`Consorcio Direto\`, cards superiores e \`Motor Caixa\`.
- Varredura de erros de formula Excel: \`#REF!\`, \`#DIV/0!\`, \`#VALUE!\`, \`#NAME?\`, \`#N/A\` e \`#NUM!\`.
- Varredura de referencias legadas no workbook final.
- Varredura de formulas ou textos de retorno interno no workbook final.
- Confirmacao de que \`Motor Caixa\` nao referencia inflacao.
- Confirmacao de que \`VPL vs base\` nao referencia inflacao.
- Confirmacao de que calculos centrais de financiamento, consorcio, aluguel, capital adicional e motor nao usam inflacao.
- Teste amortizacao por caixa livre: caixa livre mensal de R$ ${Math.round(amortizationCashFreeTests.cashFree).toLocaleString("pt-BR")} em \`Premissas!B42\`; prestacao ordinaria inicial de R$ ${Math.round(amortizationCashFreeTests.firstOrdinaryPayment).toLocaleString("pt-BR")}; amortizacao extra inicial de R$ ${Math.round(amortizationCashFreeTests.firstExtraPayment).toLocaleString("pt-BR")}; quitacao no mes ${amortizationCashFreeTests.payoffMonth}.
- Confirmacao de regressao: formulas de \`Amortizacao!J:Q\` nao referenciam consorcio, \`B32\` nem \`B36\`.
- Teste Moradia propria: financiamento no mes ${rentalComparabilityTests.moradia.accessFinancing}, chaves no mes 1 e consorcio direto no mes ${rentalComparabilityTests.moradia.accessConsorcioDireto}. Estrategias 1, 2 e 3 ficaram sem aluguel pago; estrategia 4 pagou aluguel nos meses ${rentalComparabilityTests.moradia.consorcioRentMonths}; nenhum cenario recebeu aluguel.
- Teste Investimento para aluguel: estrategias 1, 2 e 3 receberam aluguel do mes ${rentalComparabilityTests.investimento.accessFinancing} ao horizonte; estrategia 4 recebeu aluguel do mes ${rentalComparabilityTests.investimento.accessConsorcioDireto} ao horizonte; nenhum cenario teve aluguel pago.

## Checks criados/alterados

- \`Financiamento puro termina zerado\`: usa \`INDEX(Financiamento!M6:M485, Premissas!B13)\`.
- \`Amortizacao extra = menor entre sobra de caixa e saldo amortizavel\`: valida \`Amortizacao!J\` contra a formula auxiliar em \`Amortizacao!Q\`.
- \`Sobra para amortizar nunca negativa\`: valida que \`Amortizacao!P\` usa piso zero.
- \`Caixa livre do cenario 2 vem de Premissas!B42\`: valida que \`Premissas!B35\` esta amarrada ao caixa livre mensal comum ao decisor.
- \`Aluguel pago somente em moradia\`: investimento deve ter soma zero de aluguel pago nas quatro estrategias.
- \`Aluguel recebido somente em investimento\`: moradia deve ter soma zero de aluguel recebido nas quatro estrategias.
- \`Consorcio direto sem aluguel recebido antes de contemplacao/chaves\`: valida que estrategia 4 nao recebe aluguel antes de MAX(contemplacao direta, disponibilidade).
- \`Financiamento/IQ recebem aluguel no mes de acesso\`: valida estrategias 1, 2 e 3 no modo investimento.
- \`Checks adicionais de aluguel comparavel\`: valida espera do consorcio direto em moradia, parada de aluguel pago em financiamento/IQ e ausencia de aluguel recebido antecipado.

## Resultado financeiro da correcao

A comparacao do financiamento com amortizacao passa a representar uma decisao mais coerente: a mesma pessoa usa sua capacidade mensal livre para pagar a prestacao SAC e direciona somente a sobra para amortizar o saldo devedor. Isso elimina a vinculacao indevida com o fluxo do consorcio e faz o beneficio do SAC aparecer corretamente: a queda da prestacao aumenta a sobra disponivel para amortizacao extra ao longo do tempo.

A comparacao deixa de privilegiar indevidamente o consorcio direto quando a finalidade e moradia propria: enquanto a carta nao e contemplada, a estrategia 4 carrega custo de aluguel de espera. No modo investimento, a comparacao tambem fica simetrica: nenhuma estrategia paga aluguel, e a receita de aluguel entra apenas quando aquela estrategia tem acesso efetivo ao imovel.

## Simplificacoes mantidas

- O modelo permanece nominal por padrao.
- Inflacao anual esperada fica em 0% a.a. por padrao, mas e editavel.
- Valorizacao do imovel fica em 0% a.a. por conservadorismo, mas e editavel.
- Nao foram incluidos ITBI, cartorio, INCC detalhado, imposto de venda, tributacao completa de aluguel ou risco de nao contemplacao no mes preset.
`,
  "utf8",
);

if (process.env.SKIP_RENDER !== "1") {
  for (const sheetName of finalSheetNames) {
    const preview = await workbook.render({
      sheetName,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    await fs.writeFile(
      path.join(outputDir, `preview_${safePreviewName(sheetName)}.png`),
      new Uint8Array(await preview.arrayBuffer()),
    );
  }
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`OUTPUT_XLSX ${outputPath}`);
