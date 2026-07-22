import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "outputs", "consorcio_financiamento");
const outputPath = path.join(outputDir, "Calculadora_Aquisicao_Imovel_Veiculo_SAC_PRICE_Carta_Contemplada.xlsx");
const correctionReportPath = path.join(outputDir, "relatorio_final_implementacao.md");
const manualUsoPlanilhaPath = path.join(outputDir, "manual_uso_planilha.txt");
const manualCalculosPath = path.join(outputDir, "manual_calculos_planilha.txt");
const manualWebPath = path.join(outputDir, "manual_uso_interface_web.txt");

const FIN_MONTHS = 480;
const OPP_MAX_MONTHS = 480;
const CARD_MAX_MONTHS = OPP_MAX_MONTHS * 2 - 1;
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
const cartaContemplada = workbook.worksheets.add("Carta Contemplada");
const motor = workbook.worksheets.add("Motor Caixa");
const consorcioDireto = workbook.worksheets.add("Consorcio Direto");
const financiamento = workbook.worksheets.add("Financiamento");
const consorcio = workbook.worksheets.add("Consorcio IQ");
const amortizacao = workbook.worksheets.add("Amortizacao");
const checks = workbook.worksheets.add("Checks");
const fontes = workbook.worksheets.add("Fontes");

const sheets = [resumo, premissas, cartaContemplada, motor, consorcioDireto, financiamento, consorcio, amortizacao, checks, fontes];
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
setWidths(premissas, [300, 170, 140, 360, 180], 110);
section(premissas, "A3:E3", "Premissas gerais");
section(premissas, "A12:E12", "Financiamento SAC/PRICE");
section(premissas, "A20:E20", "Consorcio como interveniente quitante");
section(premissas, "A34:E34", "Amortizacao com caixa livre");
section(premissas, "A38:E38", "Custo de oportunidade");
section(premissas, "A46:E46", "Uso, disponibilidade, renda e custos", colors.teal);
section(premissas, "A71:E71", "Ponte de caixa e rotulos", colors.teal);
section(premissas, "A80:E80", "Convencoes do modelo", colors.teal);
section(premissas, "A90:E90", "Premissas por tipo do bem", colors.teal);

premissas.getRange("A4:E45").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("A4:A45").format.fill = colors.paleGray;
premissas.getRange("D4:D45").format.fill = colors.paleGray;
premissas.getRange("D4:D45").format.wrapText = true;

const assumptionRows = [
  [4, "Data de inicio dos pagamentos", "value", new Date(2027, 5, 1), dateFmt, "Pagamento do imovel na planta e venda do imovel atual em junho/2027."],
  [5, "Valor de venda do bem atual", "value", 0, currencyFmt, "Entrada de caixa inicial considerada comum aos cenarios."],
  [6, "Valor do bem novo", "value", 200000, currencyFmt, "Preco do bem comparado na simulacao."],
  [7, "Entrada (% do bem novo)", "value", 0.2, pctFmt, "Percentual minimo de entrada do financiamento."],
  [8, "Entrada em R$", "formula", "=B79*B7", currencyFmt, "Calculada sobre o preco aplicavel no mes de aquisicao do financiamento."],
  [9, "Saldo da venda apos entrada", "formula", "=B5-B8", currencyFmt, "Liquidez remanescente da venda antes de outros custos de compra."],
  [10, "Valor financiado", "formula", "=MAX(0,B79-B8)", currencyFmt, "Preco aplicavel na aquisicao menos a entrada do financiamento."],
  [11, "Tipo do bem", "value", "Imovel", "@", "Escolha Imovel ou Veiculo. A selecao altera a valorizacao/depreciacao e os fluxos aplicaveis."],
  [13, "Prazo financiamento (meses)", "value", 420, monthFmt, "Prazo informado pelo usuario."],
  [14, "Juros a.a.", "value", 0.12, pctFmt, "Convertido para taxa mensal equivalente."],
  [15, "Juros mensal equivalente", "formula", "=(1+B14)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta."],
  [16, "Indexador/correcao do saldo a.a.", "value", 0, pctFmt, "Indice contratual generico do saldo. O preset PRICE usa zero; altere conforme a proposta real."],
  [17, "Indexador mensal equivalente", "formula", "=(1+B16)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta do indexador selecionado."],
  [18, "Seguro mensal", "value", 150, currencyFmt, "Incluido enquanto houver saldo devedor no financiamento."],
  [19, "Sistema de financiamento", "value", "SAC", "@", "Escolha SAC ou PRICE. O cronograma preserva as mesmas colunas logicas para alimentar IQ, amortizacao, motor e resumo."],
  [21, "Taxa de administracao", "value", 0.24, pctFmt, "Componente editavel da taxa total do consorcio."],
  [22, "Taxa do fundo reserva", "value", 0.02, pctFmt, "Componente editavel da taxa total do consorcio."],
  [23, "Seguro prestamista (% do credito)", "value", 0.00035, pctFmt, "Tratado como percentual total sobre o credito para simplificar a simulacao."],
  [24, "Taxas totais sobre credito", "formula", "=SUM(B21:B23)", pctFmt, "Soma de administracao, fundo reserva e seguro prestamista."],
  [25, "Prazo consorcio (meses)", "value", 240, monthFmt, "Prazo informado pelo usuario."],
  [26, "Reajuste anual credito/parcela", "value", 0.06, pctFmt, "Aplicado por degrau anual."],
  [27, "Mes de contemplacao preset", "value", 60, monthFmt, "Preset editavel para a contemplacao do consorcio."],
  [28, "Fator reajuste consorcio na contemplacao", "formula", "=(1+B26)^INT((B27-1)/12)", "0.0000x", "Mes 60 recebe quatro reajustes anuais no preset, pois o 5o aniversario ocorre no mes 61."],
  [29, "Saldo financiamento no mes da contemplacao", "formula", "=IF(AND(B27>=B48,B27-B48+1<=B13),INDEX(Financiamento!$M$6:$M$485,B27-B48+1),0)", currencyFmt, "Saldo ao fim do mes de contemplacao, considerando o mes de inicio do financiamento."],
  [30, "Carta inicial contratada", "formula", "=B29/B28", currencyFmt, "Calculada para que a carta atualizada cubra o saldo devedor no mes de contemplacao."],
  [31, "Credito atualizado na contemplacao", "formula", "=B30*B28", currencyFmt, "Valor usado como interveniente quitante."],
  [32, "Parcela inicial do consorcio", "formula", "=B30*(1+B24)/B25", currencyFmt, "Credito contratado acrescido das taxas totais, dividido pelo prazo."],
  [35, "Caixa livre mensal alocado ao financiamento", "formula", "=B42", currencyFmt, "Cenario 2 usa o caixa livre mensal informado em B42 como teto para prestacao ordinaria mais amortizacao extra enquanto houver saldo."],
  [36, "Sobra inicial para amortizacao extra", "formula", "=MAX(0,B35-INDEX(Financiamento!$J$6:$J$485,1))", currencyFmt, "Diferenca inicial entre caixa livre mensal e prestacao ordinaria do financiamento selecionado. No cronograma, essa sobra e recalculada mes a mes."],
  [39, "Horizonte de comparacao (meses)", "value", 420, monthFmt, "Prazo editavel para comparar patrimonio acumulado. O modelo mensal suporta ate 480 meses."],
  [40, "Rentabilidade liquida a.a.", "value", 0.12, pctFmt, "Preset de 12% liquido ao ano para investimento das sobras mensais."],
  [41, "Rentabilidade liquida mensal equivalente", "formula", "=(1+B40)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta usada na evolucao dos investimentos."],
  [42, "Caixa livre mensal antes dos fluxos do bem", "value", 7000, currencyFmt, "Caixa mensal disponivel antes de entrada, parcelas, custo de espera, renda, consorcio e lances."],
  [43, "Valor em reserva", "value", 90000, currencyFmt, "Reserva inicial disponivel. Deve ser pelo menos a maior entrada exigida entre as estrategias."],
  [44, "Maior entrada entre estrategias", "formula", "=MAX(B8,'Consorcio Direto'!$B$44,'Carta Contemplada'!$B$37)", currencyFmt, "Check de reserva minima para comparar estrategias com e sem entrada/custo inicial de forma justa."],
  [45, "Valor do bem no horizonte", "formula", "=MAX(0,B6*(1+B66)^((B39-1)/12))", currencyFmt, "Valor de mercado no horizonte, usando a taxa selecionada pelo Tipo do bem e piso zero."],
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
inputStyle(premissas, "B11");
inputStyle(premissas, "B13:B18");
inputStyle(premissas, "B19");
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
applyNumberFormat(premissas, "B11", "@");
applyNumberFormat(premissas, "B13", monthFmt);
applyNumberFormat(premissas, "B14", pctFmt);
applyNumberFormat(premissas, "B15", pctFmt2);
applyNumberFormat(premissas, "B16", pctFmt);
applyNumberFormat(premissas, "B17", pctFmt2);
applyNumberFormat(premissas, "B18", currencyFmt);
applyNumberFormat(premissas, "B19", "@");
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
  [47, "Finalidade do Imovel", "value", "Investimento", "@", "Moradia aplica aluguel pago ate o acesso; Investimento aplica receita liquida depois do acesso. Para Veiculo, o campo e ignorado."],
  [48, "Mes de aquisicao via financiamento", "formula", "=1", monthFmt, "Financiamento e IQ compram o bem no mes inicial pelo valor atual."],
  [49, "Mes de disponibilidade/uso", "value", 1, monthFmt, "Mes a partir do qual o bem pode ser ocupado, usado ou gerar renda."],
  [50, "Mes fim aluguel de espera fin./IQ", "formula", "=MAX(B48,B49)-1", monthFmt, "Em Moradia, financiamento e IQ carregam aluguel pago antes da aquisicao e disponibilidade."],
  [51, "Inicio renda fin./IQ", "formula", "=MAX(B48,B49)", monthFmt, "Em Investimento, financiamento e IQ reconhecem receita a partir do acesso ao imovel."],
  [52, "Inicio renda consorcio direto", "formula", "=MAX('Consorcio Direto'!$B$29,B49)", monthFmt, "Consorcio direto reconhece renda apenas apos contemplacao e disponibilidade."],
  [53, "Aluguel pago (% valor Imovel a.m.)", "value", 0.0035, pctFmt2, "Premissa opcional de Moradia. Para Veiculo, o campo e ignorado."],
  [54, "Aluguel pago inicial mensal", "formula", "=IF(B11=\"Imovel\",B6*B53,0)", currencyFmt, "Aluguel mensal inicial aplicado apenas em Moradia antes do acesso ao imovel."],
  [55, "Receita bruta (% valor a.m.)", "value", 0.004, pctFmt2, "Premissa opcional de aluguel ou renda operacional. Use zero quando o bem nao gerar receita."],
  [56, "Receita bruta inicial", "formula", "=B6*B55", currencyFmt, "Receita bruta mensal antes dos redutores operacionais."],
  [57, "Reajuste anual espera/renda", "value", 0.06, pctFmt, "Aplicado por degrau anual ao custo de espera e a receita."],
  [58, "Vacancia/ociosidade", "value", 0.05, pctFmt, "Percentual redutor da receita bruta do bem."],
  [59, "Inadimplencia", "value", 0.02, pctFmt, "Percentual redutor da receita bruta do bem."],
  [60, "Administracao da renda", "value", 0.08, pctFmt, "Percentual redutor da receita bruta do bem."],
  [61, "Manutencao/operacao", "value", 0.03, pctFmt, "Percentual redutor da receita bruta do bem."],
  [62, "Impostos/taxas nao repassados", "value", 0.02, pctFmt, "Percentual redutor da receita bruta do bem."],
  [63, "IR sobre renda", "value", 0, pctFmt, "Percentual redutor, se aplicavel."],
  [64, "Receita liquida inicial", "formula", '=IF(B11="Imovel",B56*(1-SUM(B58:B63)),0)', currencyFmt, "Valor liquido usado apenas em Investimento, apos aquisicao/contemplacao e disponibilidade."],
  [65, "Modelo nominal", "value", "Nominal", "@", "O motor permanece nominal. Valores em R$ de hoje sao apenas conversoes de apresentacao usando a inflacao esperada."],
  [66, "Taxa anual aplicada ao bem", "formula", '=IF(B11="Veiculo",B92,B91)', pctFmt, "Selecionada automaticamente: valorizacao de imovel ou depreciacao de veiculo."],
  [67, "Inflacao anual esperada", "value", 0.06, pctFmt, "Preset de 6% a.a.; usada apenas para converter valores nominais futuros em R$ de hoje. Nao altera o motor nominal."],
  [68, "Inflacao mensal equivalente", "formula", "=(1+B67)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta da inflacao esperada."],
  [69, "Fator inflacao acumulada no horizonte", "formula", "=(1+B67)^(B39/12)", "0.0000x", "Divide valores nominais futuros para apresentar equivalentes em R$ de hoje."],
  [72, "Caixa comum incluido", "formula", "=B43", currencyFmt, "Reserva efetivamente usada no motor mensal de caixa."],
  [73, "Venda do imovel atual", "formula", "=B5", currencyFmt, "Entrada de caixa informativa da operacao."],
  [74, "Entrada de financiamento", "formula", "=B8", currencyFmt, "Entrada que consome reserva nas estrategias com financiamento."],
  [75, "Saldo venda apos entrada", "formula", "=B9", currencyFmt, "Ponte informativa entre venda atual e entrada."],
  [76, "Caixa comum excluido do ranking", "formula", "=MAX(0,B5-B43)", currencyFmt, "Valor da venda que nao entra no motor caso exceda a reserva informada."],
  [77, "Reserva minima recomendada", "formula", "=B44", currencyFmt, "Maior entrada inicial entre as estrategias comparadas."],
  [78, "Regra do preco na aquisicao", "formula", '="Financiamento atual; consorcio futuro"', "@", "Financiamento usa o valor atual no mes 1; consorcios usam o preco futuro no mes de contemplacao/uso."],
  [79, "Preco do financiamento na aquisicao", "formula", "=B6", currencyFmt, "Preco atual usado para calcular entrada e valor financiado no mes 1."],
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

premissas.getRange("A47:E79").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("A47:A79").format.fill = colors.paleGray;
premissas.getRange("D47:D79").format.fill = colors.paleGray;
premissas.getRange("D47:D79").format.wrapText = true;
inputStyle(premissas, "B47");
inputStyle(premissas, "B49");
inputStyle(premissas, "B53");
inputStyle(premissas, "B55");
inputStyle(premissas, "B57:B63");
inputStyle(premissas, "B65");
inputStyle(premissas, "B67");
formulaStyle(premissas, "B50:B52");
formulaStyle(premissas, "B54");
formulaStyle(premissas, "B56");
formulaStyle(premissas, "B64");
formulaStyle(premissas, "B66:B69");
formulaStyle(premissas, "B72:B77");
formulaStyle(premissas, "B78");
formulaStyle(premissas, "B79");
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
applyNumberFormat(premissas, "B79", currencyFmt);
premissas.getRange("B11").dataValidation = { rule: { type: "list", values: ["Imovel", "Veiculo"] } };
premissas.getRange("B19").dataValidation = { rule: { type: "list", values: ["SAC", "PRICE"] } };
premissas.getRange("B47").dataValidation = { rule: { type: "list", values: ["Moradia", "Investimento"] } };
premissas.getRange("B65").dataValidation = { rule: { type: "list", values: ["Nominal"] } };

premissas.getRange("A81:E89").values = [
  ["Convencao", "Descricao", "", "", ""],
  ["SAC/PRICE com indexador", "O indexador mensal corrige o saldo no inicio do mes; juros incidem sobre o saldo corrigido. Em SAC, a amortizacao divide o saldo corrigido pelos meses restantes. Em PRICE, a prestacao financeira e recalculada pelo saldo corrigido e prazo remanescente.", "", "", ""],
  ["Consorcio", "Credito e parcela reajustam anualmente por degrau. A carta inicial e dimensionada para que o credito atualizado no mes da contemplacao quite o saldo do financiamento.", "", "", ""],
  ["Interveniente quitante", "No mes da contemplacao, o cliente paga a prestacao normal do financiamento e usa a carta para quitar o saldo restante daquele mes.", "", "", ""],
  ["Amortizacao de prazo", "A sobra entre o caixa livre mensal e a prestacao ordinaria do financiamento selecionado e usada como amortizacao extra direta no saldo devedor, reduzindo prazo.", "", "", ""],
  ["Motor Caixa", "O caixa livre mensal entra antes dos fluxos do imovel. Sobra vira investimento; deficit consome reserva/saldo investido; se a reserva/saldo acabar, o deficit nao coberto torna a estrategia inviavel.", "", "", ""],
  ["Patrimonio", "Mostra patrimonio apenas quando a estrategia nao tem deficit nao coberto. Estrategias inviaveis nao devem ser ranqueadas como recomendacao.", "", "", ""],
  ["Preco na aquisicao", "Financiamento compra no mes 1 pelo valor atual. Consorcio direto e carta contemplada usam o valor futuro do bem no mes de contemplacao ou uso do credito.", "", "", ""],
  ["Posicao dos consorcios", "Antes da contemplacao, o ativo inclui o direito economico descontado e a divida inclui o VP das parcelas futuras. Depois do uso do credito, o direito e substituido pelo bem, mas a obrigacao permanece ate o fim do grupo.", "", "", ""],
];
premissas.getRange("A81:E81").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
};
premissas.getRange("A82:E89").format.borders = { preset: "all", style: "thin", color: colors.line };
premissas.getRange("B82:E89").merge(true);
premissas.getRange("B82:E89").format.wrapText = true;

const assetTypeRows = [
  [91, "Valorizacao anual do Imovel", "value", 0.06, pctFmt, "Taxa aplicada automaticamente quando Tipo do bem = Imovel."],
  [92, "Depreciacao anual do Veiculo", "value", -0.15, pctFmt, "Taxa aplicada automaticamente quando Tipo do bem = Veiculo. Deve ser maior que -100%."],
];

for (const [row, label, kind, value, fmt, note] of assetTypeRows) {
  premissas.getRange(`A${row}`).values = [[label]];
  if (kind === "value") {
    setCellValue(premissas, `B${row}`, value, fmt);
  } else {
    setCellFormula(premissas, `B${row}`, value, fmt);
  }
  premissas.getRange(`D${row}`).values = [[note]];
}
frame(premissas, "A91:E92", colors.paleGray);
premissas.getRange("D91:D92").format.wrapText = true;
inputStyle(premissas, "B91:B92");
sourceLinkStyle(premissas, "B66");
applyNumberFormat(premissas, "B91:B92", pctFmt);

// Financiamento schedule
setTitle(financiamento, "A1:M1", "Cronograma do financiamento SAC/PRICE");
setWidths(financiamento, [58, 92, 118, 110, 118, 78, 112, 105, 88, 118, 110, 122, 118], FIN_MONTHS + 8);
financiamento.getRange("A5:M5").values = [[
  "Mes",
  "Data",
  "Saldo inicial",
  "Correcao/indexador",
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
  const pricePayment = `IF(Premissas!$B$15=0,E${row}/F${row},E${row}*Premissas!$B$15/(1-(1+Premissas!$B$15)^(-F${row})))`;
  finFormulas.push([
    `=EDATE(Premissas!$B$4,${month - 1})`,
    month === 1 ? "=Premissas!$B$10" : `=M${prev}`,
    `=IF(C${row}>0,C${row}*Premissas!$B$17,0)`,
    `=C${row}+D${row}`,
    `=MAX(0,Premissas!$B$13-A${row}+1)`,
    `=IF(AND(E${row}>0,F${row}>0),IF(Premissas!$B$19="PRICE",MAX(0,MIN(E${row},${pricePayment}-H${row})),MIN(E${row},E${row}/F${row})),0)`,
    `=IF(AND(E${row}>0,F${row}>0),E${row}*Premissas!$B$15,0)`,
    `=IF(AND(E${row}>0,F${row}>0),Premissas!$B$18,0)`,
    `=IF(AND(E${row}>0,F${row}>0),IF(Premissas!$B$19="PRICE",${pricePayment}+I${row},G${row}+H${row}+I${row}),0)`,
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
setTitle(consorcio, "A1:P1", "Cronograma - consorcio como interveniente quitante");
setWidths(consorcio, [58, 92, 90, 122, 116, 125, 122, 125, 122, 118, 124, 128, 180, 138, 138, 138], CONS_MONTHS + 8);
consorcio.getRange("A5:P5").values = [[
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
  "VP parcelas futuras",
  "Direito economico cota",
  "Divida total",
]];
header(consorcio, "A5:P5");
consorcio.freezePanes.freezeRows(5);

const consMonths = Array.from({ length: CONS_MONTHS }, (_, i) => [i + 1]);
const consFormulas = [];
for (let i = 0; i < CONS_MONTHS; i += 1) {
  const row = 6 + i;
  const next = row + 1;
  const month = i + 1;
  const isLast = i === CONS_MONTHS - 1;
  consFormulas.push([
    `=EDATE(Premissas!$B$4,${month - 1})`,
    `=(1+Premissas!$B$26)^INT((A${row}-1)/12)`,
    `=Premissas!$B$30*C${row}`,
    `=IF(A${row}<=Premissas!$B$25,Premissas!$B$32*C${row},0)`,
    `=IF(AND(A${row}>=Premissas!$B$48,A${row}<=Premissas!$B$27,A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Financiamento!$C$6:$C$485,A${row}-Premissas!$B$48+1),0)`,
    `=IF(AND(A${row}>=Premissas!$B$48,A${row}<=Premissas!$B$27,A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Financiamento!$J$6:$J$485,A${row}-Premissas!$B$48+1),0)`,
    `=IF(AND(A${row}>=Premissas!$B$48,A${row}<=Premissas!$B$27,A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Financiamento!$M$6:$M$485,A${row}-Premissas!$B$48+1),0)`,
    `=IF(A${row}=Premissas!$B$27,MIN(D${row},H${row}),0)`,
    `=IF(A${row}=Premissas!$B$27,D${row}-H${row},0)`,
    `=E${row}+G${row}`,
    `=SUM($K$6:K${row})`,
    `=IF(A${row}<Premissas!$B$27,"Financiamento + consorcio",IF(A${row}=Premissas!$B$27,"Quitacao via carta",IF(A${row}<=Premissas!$B$25,"Somente consorcio","Encerrado")))`,
    isLast ? "=0" : `=IF(A${row}>=Premissas!$B$25,0,(N${next}+E${next})/(1+Premissas!$B$41))`,
    `=IF(A${row}<Premissas!$B$27,Premissas!$B$31/(1+Premissas!$B$41)^(Premissas!$B$27-A${row}),0)`,
    `=MAX(0,H${row}-I${row})+N${row}`,
  ]);
}
consorcio.getRange(`A6:A${CONS_MONTHS + 5}`).values = consMonths;
consorcio.getRange(`B6:P${CONS_MONTHS + 5}`).formulas = consFormulas;
frame(consorcio, `A6:P${CONS_MONTHS + 5}`);
applyNumberFormat(consorcio, `A6:A${CONS_MONTHS + 5}`, monthFmt);
applyNumberFormat(consorcio, `B6:B${CONS_MONTHS + 5}`, dateFmt);
applyNumberFormat(consorcio, `C6:C${CONS_MONTHS + 5}`, "0.0000x");
applyNumberFormat(consorcio, `D6:P${CONS_MONTHS + 5}`, currencyFmt);
consorcio.getRange("A65:P65").format.fill = colors.paleGreen;

// Amortizacao alternativa
setTitle(amortizacao, "A1:Q1", "Cronograma - amortizacao de prazo com caixa livre");
setWidths(amortizacao, [58, 92, 118, 110, 118, 122, 105, 88, 122, 122, 124, 118, 128, 120, 124, 132, 118], FIN_MONTHS + 8);
amortizacao.getRange("A5:Q5").values = [[
  "Mes",
  "Data",
  "Saldo inicial",
  "Correcao/indexador",
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
setTitle(consorcioDireto, "A1:M1", "Consorcio Direto - parametros e cronograma");
setWidths(consorcioDireto, [58, 110, 145, 130, 138, 138, 138, 138, 138, 170, 145, 145, 145], OPP_MAX_MONTHS + 70);
consorcioDireto.freezePanes.freezeRows(54);

section(consorcioDireto, "A3:E3", "Inputs e parametros");
section(consorcioDireto, "A22:E22", "Consorcio e lances");
section(consorcioDireto, "G3:J3", "Checks auxiliares");

const cdInputs = [
  [4, "Preco futuro do bem na aquisicao", "formula", '=MAX(0,Premissas!$B$6*(1+Premissas!$B$66)^((B29-1)/12))', currencyFmt, "Valor do bem atualizado ate o mes de contemplacao pela taxa do tipo selecionado."],
  [5, "Data de inicio", "formula", "=Premissas!$B$4", dateFmt, "Inicio dos fluxos mensais."],
  [6, "Caixa livre mensal antes dos fluxos do bem", "formula", "=Premissas!$B$42", currencyFmt, "Caixa mensal antes de parcelas, renda, aluguel de espera e lances."],
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
  [32, "Lance embutido (% credito)", "value", 0.25, pctFmt, "Percentual sobre o credito bruto atualizado, limitado ao lance total."],
  [33, "Usar lance livre?", "value", "Nao", "@", "Se Sim, percentual compoe lance com recursos do cliente."],
  [34, "Lance livre (% credito)", "value", 0, pctFmt, "Percentual sobre o credito bruto atualizado."],
  [35, "Usar lance fixo?", "value", "Sim", "@", "Se Sim, percentual compoe lance com recursos do cliente."],
  [36, "Lance fixo (% credito)", "value", 0.25, pctFmt, "Percentual do lance total fixo sobre o credito bruto atualizado."],
  [37, "Lance embutido efetivo", "formula", '=IF(AND(B31="Sim",OR(B33="Sim",B35="Sim")),MIN(B32,B38),0)', pctFmt, "Percentual embutido efetivo, nunca maior que o lance total."],
  [38, "Lance total efetivo", "formula", '=IF(B35="Sim",B36,IF(B33="Sim",B34,0))', pctFmt, "Lance fixo ou livre; as duas modalidades sao mutuamente exclusivas."],
  [39, "Descapitalizacao do cliente", "formula", "=MAX(0,B38-B37)", pctFmt, "Lance total em R$ menos a parcela embutida."],
  [40, "Fator reajuste na contemplacao", "formula", "=(1+B28)^INT((B29-1)/12)", "0.0000x", "Fator acumulado ate o mes de contemplacao."],
  [41, "Carta inicial contratada", "formula", "=B4/MAX(0.01,B40*(1-B37))", currencyFmt, "Dimensionada para que o credito liquido na contemplacao cubra o preco futuro do bem."],
  [42, "Credito bruto na contemplacao", "formula", "=B41*B40", currencyFmt, "Credito antes do lance embutido."],
  [43, "Credito liquido para compra", "formula", "=B42*(1-B37)", currencyFmt, "Credito disponivel para pagar o bem depois do lance embutido."],
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

consorcioDireto.getRange("G4:J12").values = [
  ["Check", "Atual", "Esperado", "Status"],
  ["Credito liquido cobre bem futuro", null, null, null],
  ["Mes de aquisicao/contemplacao", null, null, null],
  ["Prazo dentro do limite", null, null, null],
  ["Lance cliente no mes correto", null, null, null],
  ["Lance fixo/livre exclusivo", null, null, null],
  ["Embutido exige lance total", null, null, null],
  ["Embutido nao supera total", null, null, null],
  ["Caixa cliente = total - embutido", null, null, null],
];
header(consorcioDireto, "G4:J4", colors.teal);
consorcioDireto.getRange("H5:J12").formulas = [
  ["=$B$43", "=$B$4", '=IF(H5>=I5,"OK","Revisar")'],
  ["=$B$29", "=$B$29", '=IF(H6=I6,"OK","Revisar")'],
  ["=$B$27", `=${OPP_MAX_MONTHS}`, '=IF(H7<=I7,"OK","Revisar")'],
  ["=INDEX($F$55:$F$534,$B$29)", "=$B$44", '=IF(ABS(H8-I8)<=1,"OK","Revisar")'],
  ['=IF(AND($B$33="Sim",$B$35="Sim"),1,0)', "=0", '=IF(H9=I9,"OK","Revisar")'],
  ['=IF(AND($B$31="Sim",$B$38=0),1,0)', "=0", '=IF(H10=I10,"OK","Revisar")'],
  ["=MAX(0,$B$37-$B$38)", "=0", '=IF(ABS(H11-I11)<=0.000001,"OK","Revisar")'],
  ["=$B$44", "=$B$42*MAX(0,$B$38-$B$37)", '=IF(ABS(H12-I12)<=1,"OK","Revisar")'],
];
frame(consorcioDireto, "G5:J12");
applyNumberFormat(consorcioDireto, "H5:I5", currencyFmt);
applyNumberFormat(consorcioDireto, "H6:I7", monthFmt);
applyNumberFormat(consorcioDireto, "H8:I8", currencyFmt);
applyNumberFormat(consorcioDireto, "H11:I11", pctFmt);
applyNumberFormat(consorcioDireto, "H12:I12", currencyFmt);

const cdStartRow = 55;
const cdEndRow = cdStartRow + OPP_MAX_MONTHS - 1;
consorcioDireto.getRange("A54:M54").values = [[
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
  "VP parcelas futuras",
  "Direito economico cota",
  "Posicao liquida cota",
]];
header(consorcioDireto, "A54:M54", colors.teal);

const cdMonths = Array.from({ length: OPP_MAX_MONTHS }, (_, i) => [i + 1]);
const cdFormulas = [];
for (let i = 0; i < OPP_MAX_MONTHS; i += 1) {
  const row = cdStartRow + i;
  const next = row + 1;
  const isLast = i === OPP_MAX_MONTHS - 1;
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
    isLast ? "=0" : `=IF(A${row}>=$B$27,0,(K${next}+E${next})/(1+$B$9))`,
    `=IF(A${row}<$B$29,$B$43/(1+$B$9)^($B$29-A${row}),0)`,
    `=L${row}-K${row}`,
  ]);
}
consorcioDireto.getRange(`A${cdStartRow}:A${cdEndRow}`).values = cdMonths;
consorcioDireto.getRange(`B${cdStartRow}:M${cdEndRow}`).formulas = cdFormulas;
frame(consorcioDireto, `A${cdStartRow}:M${cdEndRow}`);
applyNumberFormat(consorcioDireto, `A${cdStartRow}:A${cdEndRow}`, monthFmt);
applyNumberFormat(consorcioDireto, `B${cdStartRow}:B${cdEndRow}`, dateFmt);
applyNumberFormat(consorcioDireto, `C${cdStartRow}:C${cdEndRow}`, currencyFmt);
applyNumberFormat(consorcioDireto, `D${cdStartRow}:D${cdEndRow}`, "0.0000x");
applyNumberFormat(consorcioDireto, `E${cdStartRow}:I${cdEndRow}`, currencyFmt);
applyNumberFormat(consorcioDireto, `K${cdStartRow}:M${cdEndRow}`, currencyFmt);
consorcioDireto.getRange(`A${cdStartRow}:M${cdStartRow}`).format.fill = colors.paleGreen;

// Carta contemplada - estrategia 5
const ccStartRow = 55;
const ccEndRow = ccStartRow + CARD_MAX_MONTHS - 1;
setTitle(cartaContemplada, "A1:K1", "Carta Contemplada - parametros e cronograma");
setWidths(cartaContemplada, [58, 110, 145, 130, 138, 138, 138, 138, 145, 170, 138], CARD_MAX_MONTHS + 70);
cartaContemplada.freezePanes.freezeRows(54);

section(cartaContemplada, "A3:E3", "Inputs e parametros");
section(cartaContemplada, "A22:E22", "Compra da carta contemplada");
section(cartaContemplada, "G3:J3", "Checks auxiliares");

const cartaInputs = [
  [4, "Preco futuro do bem na aquisicao", "formula", '=MAX(0,Premissas!$B$6*(1+Premissas!$B$66)^((B44-1)/12))', currencyFmt, "Valor do bem atualizado ate o mes efetivo de aquisicao."],
  [5, "Data de inicio", "formula", "=Premissas!$B$4", dateFmt, "Inicio dos fluxos mensais."],
  [6, "Caixa livre mensal antes dos fluxos do bem", "formula", "=Premissas!$B$42", currencyFmt, "Caixa mensal antes de parcelas, custo da carta e demais fluxos."],
  [7, "Horizonte de comparacao (meses)", "formula", "=Premissas!$B$39", monthFmt, "Puxa o horizonte das Premissas."],
  [8, "Rentabilidade liquida a.a.", "formula", "=Premissas!$B$40", pctFmt, "Puxa a rentabilidade das Premissas."],
  [9, "Rentabilidade liquida mensal", "formula", "=(1+B8)^(1/12)-1", pctFmt2, "Taxa mensal equivalente composta."],
  [22, "Comprar carta contemplada?", "value", "Sim", "@", "Se Sim, a estrategia 5 entra no Motor Caixa e no Resumo."],
  [23, "Mes de compra/transferencia", "value", 1, monthFmt, "Mes em que o cliente paga a carta e assume as parcelas restantes."],
  [24, "Mes de uso do credito", "value", 1, monthFmt, "Mes em que o credito fica disponivel para compra do bem, sujeito a aprovacao da administradora."],
  [25, "Credito bruto da carta", "value", 200000, currencyFmt, "Credito bruto atualizado da carta contemplada comprada no mercado."],
  [26, "Restricoes do credito", "value", 0, currencyFmt, "Valor do credito que nao pode ser usado na compra, se houver restricao contratual."],
  [27, "Credito liquido utilizavel", "formula", "=MAX(0,B25-B26)", currencyFmt, "Credito efetivamente disponivel para aquisicao do bem."],
  [28, "Preco pago pela carta", "value", 45000, currencyFmt, "Preco efetivo pago ao vendedor da carta, incluindo eventual historico/agio negociado."],
  [29, "Agio/desagio inicial vs credito", "formula", "=B28-B27", currencyFmt, "Diagnostico simples: preco pago menos credito liquido utilizavel. Nao substitui o custo economico completo."],
  [30, "Taxa de transferencia", "value", 2000, currencyFmt, "Taxas de cessao, transferencia e custos administrativos."],
  [31, "Parcelas restantes", "value", 180, monthFmt, "Quantidade de parcelas que o comprador ainda assumira no grupo."],
  [32, "Parcela atual", "value", 1200, currencyFmt, "Parcela vigente da carta no mes de compra."],
  [33, "Reajuste anual da parcela", "value", 0.06, pctFmt, "Reajuste anual por degrau aplicado as parcelas restantes."],
  [34, "Credito excedente desconsiderado", "formula", "=MAX(0,B27-B4)", currencyFmt, "Tratamento conservador: o excedente nao vira caixa, investimento ou patrimonio financeiro."],
  [35, "Complemento necessario", "formula", "=MAX(0,B4-B27)", currencyFmt, "Recursos proprios necessarios se o credito liquido nao cobrir o valor do bem."],
  [36, "Credito excedente no caixa", "formula", "=0", currencyFmt, "O modelo nunca libera credito excedente como caixa investivel."],
  [37, "Custo inicial total da carta", "formula", "=MAX(0,B28+B30+B35)", currencyFmt, "Preco pago pela carta + transferencia + complemento necessario."],
  [38, "VP parcelas restantes", "formula", `=SUM($K$${ccStartRow}:$K$${ccEndRow})`, currencyFmt, "Valor presente das parcelas restantes pela taxa de oportunidade mensal."],
  [39, "Custo economico da carta", "formula", "=B28+B30+B38-B27", currencyFmt, "Preco + transferencia + VP das parcelas restantes - credito liquido utilizavel."],
  [40, "Obrigacao remanescente no horizonte", "formula", `=INDEX($I$${ccStartRow}:$I$${ccEndRow},MAX(1,MIN($B$7,${OPP_MAX_MONTHS})))`, currencyFmt, "Obrigacao de parcelas futuras considerada como divida no patrimonio da estrategia 5."],
  [41, "Parcela inicial da carta", "formula", "=B32", currencyFmt, "Primeira parcela assumida apos a transferencia."],
  [42, "Credito liquido cobre bem?", "formula", "=B27+B35-B4", currencyFmt, "Deve ser zero ou positivo quando o complemento foi calculado corretamente."],
  [43, "Caixa proprio usado na compra", "formula", "=B37", currencyFmt, "Desembolso inicial liquido usado pelo Motor Caixa."],
  [44, "Mes efetivo de aquisicao", "formula", "=MAX(B23,B24)", monthFmt, "Mes a partir do qual o bem pode entrar no patrimonio, antes da regra de disponibilidade/chaves."],
];

for (const [row, label, kind, value, fmt, note] of cartaInputs) {
  cartaContemplada.getRange(`A${row}`).values = [[label]];
  if (kind === "value") {
    setCellValue(cartaContemplada, `B${row}`, value, fmt);
  } else {
    setCellFormula(cartaContemplada, `B${row}`, value, fmt);
  }
  cartaContemplada.getRange(`D${row}:E${row}`).merge();
  cartaContemplada.getRange(`D${row}`).values = [[note]];
}

frame(cartaContemplada, "A4:E9", colors.paleGray);
frame(cartaContemplada, "A23:E44", colors.paleGray);
inputStyle(cartaContemplada, "B22:B26");
inputStyle(cartaContemplada, "B28");
inputStyle(cartaContemplada, "B30:B33");
formulaStyle(cartaContemplada, "B4:B9");
formulaStyle(cartaContemplada, "B27");
formulaStyle(cartaContemplada, "B29");
formulaStyle(cartaContemplada, "B34:B44");
cartaContemplada.getRange("D4:E44").format.wrapText = true;
cartaContemplada.getRange("B22").dataValidation = { rule: { type: "list", values: ["Sim", "Nao"] } };

cartaContemplada.getRange("G4:J9").values = [
  ["Check", "Atual", "Esperado", "Status"],
  ["Credito liquido + complemento cobre bem", null, null, null],
  ["Custo inicial no mes de compra", null, null, null],
  ["Quantidade de parcelas restantes", null, null, null],
  ["Obrigacao no horizonte bate cronograma", null, null, null],
  ["Credito excedente nunca vira caixa", null, null, null],
];
header(cartaContemplada, "G4:J4", colors.teal);
cartaContemplada.getRange("H5:J9").formulas = [
  ["=$B$27+$B$35", "=$B$4", '=IF(H5+1>=I5,"OK","Revisar")'],
  [`=INDEX($E$${ccStartRow}:$E$${ccEndRow},$B$23)`, "=$B$37", '=IF(ABS(H6-I6)<=1,"OK","Revisar")'],
  [`=COUNTIF($C$${ccStartRow}:$C$${ccEndRow},\">0\")`, "=$B$31", '=IF(H7=I7,"OK","Revisar")'],
  [`=INDEX($I$${ccStartRow}:$I$${ccEndRow},MAX(1,MIN($B$7,${OPP_MAX_MONTHS})))`, "=$B$40", '=IF(ABS(H8-I8)<=1,"OK","Revisar")'],
  ["=$B$36", "=0", '=IF(ABS(H9-I9)<=1,"OK","Revisar")'],
];
frame(cartaContemplada, "G5:J9");
applyNumberFormat(cartaContemplada, "H5:I6", currencyFmt);
applyNumberFormat(cartaContemplada, "H7:I7", monthFmt);
applyNumberFormat(cartaContemplada, "H8:I9", currencyFmt);

cartaContemplada.getRange("A54:K54").values = [[
  "Mes",
  "Data",
  "Parcela carta",
  "Fator reajuste",
  "Custo inicial",
  "Credito liquido",
  "Complemento",
  "Credito excedente no caixa",
  "Obrigacao remanescente",
  "Status",
  "VP parcela",
]];
header(cartaContemplada, "A54:K54", colors.teal);

const ccMonths = Array.from({ length: CARD_MAX_MONTHS }, (_, i) => [i + 1]);
const ccFormulas = [];
for (let i = 0; i < CARD_MAX_MONTHS; i += 1) {
  const row = ccStartRow + i;
  const next = row + 1;
  const isLast = i === CARD_MAX_MONTHS - 1;
  ccFormulas.push([
    `=EDATE($B$5,A${row}-1)`,
    `=IF(AND($B$22="Sim",A${row}>=$B$23,A${row}<$B$23+$B$31),$B$32*D${row},0)`,
    `=IF(A${row}<$B$23,1,(1+$B$33)^INT((A${row}-$B$23)/12))`,
    `=IF(AND($B$22="Sim",A${row}=$B$23),$B$37,0)`,
    `=IF(AND($B$22="Sim",A${row}=$B$24),$B$27,0)`,
    `=IF(AND($B$22="Sim",A${row}=$B$23),$B$35,0)`,
    `=IF(AND($B$22="Sim",A${row}=$B$23),$B$36,0)`,
    isLast ? "=0" : `=IF(OR($B$22<>"Sim",A${row}>=$B$23+$B$31-1),0,(I${next}+C${next})/(1+$B$9))`,
    `=IF($B$22<>"Sim","Desativada",IF(A${row}<$B$23,"Antes da compra",IF(A${row}<$B$23+$B$31,"Parcelas restantes","Encerrado")))`,
    `=IF(C${row}>0,C${row}/(1+$B$9)^A${row},0)`,
  ]);
}
cartaContemplada.getRange(`A${ccStartRow}:A${ccEndRow}`).values = ccMonths;
cartaContemplada.getRange(`B${ccStartRow}:K${ccEndRow}`).formulas = ccFormulas;
frame(cartaContemplada, `A${ccStartRow}:K${ccEndRow}`);
applyNumberFormat(cartaContemplada, `A${ccStartRow}:A${ccEndRow}`, monthFmt);
applyNumberFormat(cartaContemplada, `B${ccStartRow}:B${ccEndRow}`, dateFmt);
applyNumberFormat(cartaContemplada, `C${ccStartRow}:C${ccEndRow}`, currencyFmt);
applyNumberFormat(cartaContemplada, `D${ccStartRow}:D${ccEndRow}`, "0.0000x");
applyNumberFormat(cartaContemplada, `E${ccStartRow}:I${ccEndRow}`, currencyFmt);
applyNumberFormat(cartaContemplada, `K${ccStartRow}:K${ccEndRow}`, currencyFmt);
cartaContemplada.getRange(`A${ccStartRow}:K${ccStartRow}`).format.fill = colors.paleGreen;

// Motor mensal corrigido de caixa e patrimonio
const MOTOR_END_COL = "CI";
const MOTOR_CURRENCY_END_COL = "CH";
setTitle(motor, "A1:CI1", "Motor mensal de caixa e patrimonio liquido ajustado");
setWidths(
  motor,
  [
    58, 92, 118, 128, 142, 130, 138,
    ...Array.from({ length: 80 }, (_, i) => (i % 16 === 15 ? 118 : 122)),
  ],
  OPP_MAX_MONTHS + 8,
);
motor.freezePanes.freezeRows(5);
motor.getRange(`A3:${MOTOR_END_COL}3`).values = [[
  "Metodo",
  "A capacidade mensal e o caixa livre antes dos fluxos do bem. Aluguel de espera, renda, parcelas e entradas iniciais entram no mes. Consorcios carregam direito economico e VP das parcelas futuras; deficit nao coberto torna a estrategia inviavel.",
  ...Array.from({ length: 85 }, () => ""),
]];
motor.getRange("A3").format.font = { bold: true };
motor.getRange(`B3:${MOTOR_END_COL}3`).merge();
motor.getRange(`B3:${MOTOR_END_COL}3`).format.wrapText = true;

const motorSharedHeaders = [
  "Mes",
  "Data",
  "Caixa livre mensal",
  "Custo de espera base",
  "Renda liquida base",
  "Valor bem no mes",
  "FV base sem compra",
];
const motorMetricHeaders = [
  "Entrada/lance",
  "Parcela/desembolso",
  "Aluguel de espera",
  "Renda do bem",
  "Saidas totais",
  "Entradas totais",
  "Fluxo livre",
  "Saldo antes ajuste",
  "Deficit nao coberto mes",
  "Deficit nao coberto VF",
  "Saldo investido",
  "Saldo devedor",
  "Patrimonio bem",
  "Patrimonio bruto",
  "Patrimonio",
  "Status",
];
const motorStrategies = [
  {
    key: "fin",
    title: "1. Financiamento sem amortizacao",
    start: 7,
    activeExpr: "TRUE",
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: "Premissas!$B$48+Premissas!$B$13-1",
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF(AND($A${row}>=Premissas!$B$48,$A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Financiamento!$J$6:$J$485,$A${row}-Premissas!$B$48+1),0)`,
    debtFormula: (row) => `=IF(AND($A${row}>=Premissas!$B$48,$A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Financiamento!$M$6:$M$485,$A${row}-Premissas!$B$48+1),0)`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)`,
    note: "Compra no inicio via financiamento; uso ou renda so depois da disponibilidade.",
  },
  {
    key: "amort",
    title: "2. Financiamento com amortizacao",
    start: 23,
    activeExpr: "TRUE",
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: 'IFERROR(Premissas!$B$48+MATCH("Quitado",Amortizacao!$N$6:$N$485,0)-1,"N/A")',
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF(AND($A${row}>=Premissas!$B$48,$A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Amortizacao!$K$6:$K$485,$A${row}-Premissas!$B$48+1),0)`,
    debtFormula: (row) => `=IF(AND($A${row}>=Premissas!$B$48,$A${row}-Premissas!$B$48+1<=Premissas!$B$13),INDEX(Amortizacao!$L$6:$L$485,$A${row}-Premissas!$B$48+1),0)`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)`,
    note: "Usa a sobra entre caixa livre mensal e prestacao ordinaria do financiamento selecionado para reduzir prazo.",
  },
  {
    key: "iq",
    title: "3. Financiamento + consorcio IQ",
    start: 39,
    activeExpr: "TRUE",
    acquisitionExpr: "Premissas!$B$48",
    availabilityExpr: "MAX(Premissas!$B$48,Premissas!$B$49)",
    payoffExpr: "Premissas!$B$27",
    entryFormula: (row) => `=IF($A${row}=Premissas!$B$48,Premissas!$B$8,0)`,
    paymentFormula: (row) => `=IF($A${row}<=Premissas!$B$25,INDEX('Consorcio IQ'!$K$6:$K$${CONS_MONTHS + 5},$A${row}),0)`,
    debtFormula: (row) => `=INDEX('Consorcio IQ'!$P$6:$P$${CONS_MONTHS + 5},$A${row})`,
    propertyFormula: (row) => `=IF($A${row}>=Premissas!$B$48,$F${row},0)+IF($A${row}<Premissas!$B$27,INDEX('Consorcio IQ'!$O$6:$O$${CONS_MONTHS + 5},$A${row}),0)`,
    note: "Compra por financiamento; antes da contemplacao soma o direito da cota e, ate o fim do grupo, desconta o VP das parcelas futuras.",
  },
  {
    key: "cons",
    title: "4. Consorcio direto",
    start: 55,
    activeExpr: "TRUE",
    acquisitionExpr: "'Consorcio Direto'!$B$29",
    availabilityExpr: "MAX('Consorcio Direto'!$B$29,Premissas!$B$49)",
    payoffExpr: "'Consorcio Direto'!$B$29",
    entryFormula: (row) => `=IF($A${row}='Consorcio Direto'!$B$29,'Consorcio Direto'!$B$44,0)`,
    paymentFormula: (row) => `=IF($A${row}<='Consorcio Direto'!$B$27,INDEX('Consorcio Direto'!$E$55:$E$534,$A${row}),0)`,
    debtFormula: (row) => `=INDEX('Consorcio Direto'!$K$55:$K$534,$A${row})`,
    propertyFormula: (row) => `=IF($A${row}<'Consorcio Direto'!$B$29,INDEX('Consorcio Direto'!$L$55:$L$534,$A${row}),$F${row})`,
    note: "Antes da contemplacao reconhece direito economico conservador e obrigacao; depois substitui o direito pelo bem e mantem o VP das parcelas futuras.",
  },
  {
    key: "carta",
    title: "5. Carta contemplada",
    start: 71,
    activeExpr: "'Carta Contemplada'!$B$22=\"Sim\"",
    acquisitionExpr: "'Carta Contemplada'!$B$44",
    availabilityExpr: "MAX('Carta Contemplada'!$B$44,Premissas!$B$49)",
    payoffExpr: "'Carta Contemplada'!$B$23+'Carta Contemplada'!$B$31-1",
    entryFormula: (row) => `=IF(AND('Carta Contemplada'!$B$22="Sim",$A${row}='Carta Contemplada'!$B$23),'Carta Contemplada'!$B$37,0)`,
    paymentFormula: (row) => `=IF('Carta Contemplada'!$B$22="Sim",INDEX('Carta Contemplada'!$C$${ccStartRow}:$C$${ccEndRow},$A${row}),0)`,
    debtFormula: (row) => `=IF(AND('Carta Contemplada'!$B$22="Sim",$A${row}>='Carta Contemplada'!$B$23),INDEX('Carta Contemplada'!$I$${ccStartRow}:$I$${ccEndRow},$A${row}),0)`,
    propertyFormula: (row) => `=IF(AND('Carta Contemplada'!$B$22="Sim",$A${row}>='Carta Contemplada'!$B$44),$F${row},0)`,
    note: "Compra carta ja contemplada; considera preco aplicavel, custo inicial, credito liquido, complemento e VP das parcelas restantes.",
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
header(motor, `A4:${MOTOR_END_COL}4`, colors.teal);

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
    `=MAX(0,Premissas!$B$6*(1+Premissas!$B$66)^((A${row}-1)/12))`,
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
      `=IF(AND(${strategy.activeExpr},Premissas!$B$11="Imovel",Premissas!$B$47="Moradia",$A${row}<${strategy.availabilityExpr}),$D${row},0)`,
      `=IF(AND(${strategy.activeExpr},Premissas!$B$11="Imovel",Premissas!$B$47="Investimento",$A${row}>=${strategy.availabilityExpr}),$E${row},0)`,
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
      `=IF(${c[9]}${row}>1,"Inviavel",${c[13]}${row})`,
      `=IF(${c[9]}${row}>1,"Inviavel","OK")`,
    ]);
  }
  const startCol = colLetter(start);
  const endCol = colLetter(start + motorMetricHeaders.length - 1);
  motor.getRange(`${startCol}5:${endCol}${OPP_MAX_MONTHS + 4}`).formulas = formulas;
}

frame(motor, `A5:${MOTOR_END_COL}${OPP_MAX_MONTHS + 4}`);
applyNumberFormat(motor, `A5:A${OPP_MAX_MONTHS + 4}`, monthFmt);
applyNumberFormat(motor, `B5:B${OPP_MAX_MONTHS + 4}`, dateFmt);
applyNumberFormat(motor, `C5:${MOTOR_CURRENCY_END_COL}${OPP_MAX_MONTHS + 4}`, currencyFmt);
motor.getRange(`${MOTOR_END_COL}5:${MOTOR_END_COL}${OPP_MAX_MONTHS + 4}`).format.wrapText = true;
motor.getRange(`A5:${MOTOR_END_COL}5`).format.fill = colors.paleGreen;

// Checks
setTitle(checks, "A1:G1", "Checks do modelo");
setWidths(checks, [440, 135, 135, 120, 95, 85, 640], 85);
checks.getRange("A4:G4").values = [["Check", "Atual", "Esperado", "Diferenca", "Tolerancia", "Status", "Notas"]];
header(checks, "A4:G4", colors.teal);

const strategyCheckRefs = [
  { row: 10, paid: "J", received: "K", deficit: "Q", invested: "R", debt: "S", asset: "T", gross: "U", status: "W", access: "MAX(Premissas!$B$48,Premissas!$B$49)" },
  { row: 11, paid: "Z", received: "AA", deficit: "AG", invested: "AH", debt: "AI", asset: "AJ", gross: "AK", status: "AM", access: "MAX(Premissas!$B$48,Premissas!$B$49)" },
  { row: 12, paid: "AP", received: "AQ", deficit: "AW", invested: "AX", debt: "AY", asset: "AZ", gross: "BA", status: "BC", access: "MAX(Premissas!$B$48,Premissas!$B$49)" },
  { row: 13, paid: "BF", received: "BG", deficit: "BM", invested: "BN", debt: "BO", asset: "BP", gross: "BQ", status: "BS", access: "MAX('Consorcio Direto'!$B$29,Premissas!$B$49)" },
  { row: 14, paid: "BV", received: "BW", deficit: "CC", invested: "CD", debt: "CE", asset: "CF", gross: "CG", status: "CI", access: "MAX('Carta Contemplada'!$B$44,Premissas!$B$49)" },
];
const motorIndex = (col) => `INDEX('Motor Caixa'!$${col}$5:$${col}$484,Premissas!$B$39)`;
const motorRange = (col) => `'Motor Caixa'!$${col}$5:$${col}$484`;
const allPaidRanges = strategyCheckRefs.map((s) => motorRange(s.paid)).join(",");
const allReceivedRanges = strategyCheckRefs.map((s) => motorRange(s.received)).join(",");
const minInvestedFormula = `=MIN(${strategyCheckRefs.map((s) => motorRange(s.invested)).join(",")})`;
const patrimonioTieFormula = `=MAX(${strategyCheckRefs.map((s) => `IF(ISNUMBER(Resumo!$I$${s.row}),ABS(Resumo!$I$${s.row}-(${motorIndex(s.invested)}+${motorIndex(s.asset)}-${motorIndex(s.debt)})),0)`).join(",")})`;
const patrimonioStatusFormula = `=IF(AND(${strategyCheckRefs.map((s) => `IF(Resumo!$H$${s.row}>1,AND(Resumo!$I$${s.row}="Inviavel",Resumo!$J$${s.row}="Inviavel"),AND(ISNUMBER(Resumo!$I$${s.row}),ISNUMBER(Resumo!$J$${s.row})))`).join(",")}),0,1)`;
const deficitStatusFormula = `=MAX(${strategyCheckRefs.map((s) => `ABS(IF(Resumo!$O$${s.row}="Inviavel",1,0)-IF(Resumo!$H$${s.row}>1,1,0))`).join(",")})`;
const vplGapFormula = `=MAX(${strategyCheckRefs.map((s) => `IF(ISNUMBER(Resumo!$K$${s.row}),ABS(Resumo!$K$${s.row}*(1+Premissas!$B$41)^Premissas!$B$39-Resumo!$L$${s.row}),0)`).join(",")})`;
const patrimonioHojeFormula = `=MAX(${strategyCheckRefs.map((s) => `IF(ISNUMBER(Resumo!$J$${s.row}),ABS(Resumo!$J$${s.row}-Resumo!$I$${s.row}/Premissas!$B$69),0)`).join(",")})`;
const investmentAccessFormula = `=IF(AND(Premissas!$B$11="Imovel",Premissas!$B$47="Investimento"),MAX(${strategyCheckRefs.map((s) => `ABS(INDEX(${motorRange(s.received)},${s.access})-INDEX('Motor Caixa'!$E$5:$E$484,${s.access}))`).join(",")}),0)`;

checks.getRange("A5:G30").values = [
  ["Financiamento selecionado termina zerado", null, null, null, null, null, "Saldo final no fim do prazo deve ser zero em SAC e PRICE."],
  ["Carta IQ cobre saldo no mes de contemplacao", null, null, null, null, null, "Credito atualizado deve bater com o saldo devedor projetado no mes de contemplacao."],
  ["Consorcio direto: credito liquido cobre bem", null, null, null, null, null, "Credito liquido do consorcio direto deve cobrir o valor do bem ou exigir complemento fora deste bloco."],
  ["Consorcio direto: prazo dentro do limite", null, null, null, null, null, "Cronograma mensal do consorcio direto suporta ate 480 meses."],
  ["Reserva cobre maior entrada/lance/carta", null, null, null, null, null, "Reserva deve cobrir a maior necessidade inicial relevante entre as estrategias."],
  ["Motor: horizonte dentro do limite", null, null, null, null, null, "O motor mensal suporta ate 480 meses."],
  ["Motor: saldo investido nao fica negativo", null, null, null, null, null, "Saldo investido e truncado em zero; deficit nao coberto marca inviabilidade."],
  ["Motor referencia Consorcio Direto", null, null, null, null, null, "A estrategia 4 usa a aba Consorcio Direto para contemplacao, parcela e lance."],
  ["Tipo do bem valido", null, null, null, null, null, "Tipo do bem deve ser Imovel ou Veiculo."],
  ["Sistema de financiamento valido", null, null, null, null, null, "Sistema deve ser SAC ou PRICE."],
  ["Modelo permanece nominal", null, null, null, null, null, "Nao permitir Real sem logica real implementada."],
  ["Depreciacao anual dentro do limite", null, null, null, null, null, "A depreciacao anual do veiculo deve ser maior que -100%."],
  ["Patrimonio nominal bate formula", null, null, null, null, null, "Quando a estrategia e executavel, patrimonio nominal = saldo investido + bem - saldo devedor/obrigacao."],
  ["Patrimonio bate status", null, null, null, null, null, "Se houver deficit nao coberto, patrimonio deve mostrar Inviavel."],
  ["Deficit nao coberto marca inviabilidade", null, null, null, null, null, "Deficit nao coberto VF deve explicar o Status Inviavel."],
  ["Aluguel pago somente em Moradia de Imovel", null, null, null, null, null, "Fora de Moradia de Imovel, a coluna de aluguel pago deve ser zero."],
  ["Renda somente em Investimento de Imovel", null, null, null, null, null, "Fora de Investimento de Imovel, a receita do bem deve ser zero."],
  ["Consorcio direto sem renda antes de contemplacao/disponibilidade", null, null, null, null, null, "Renda da estrategia 4 so pode aparecer apos o maior entre contemplacao direta e disponibilidade."],
  ["Carta contemplada sem receita antes de aquisicao/disponibilidade", null, null, null, null, null, "Receita da estrategia 5 so pode aparecer apos o maior entre uso do credito e disponibilidade."],
  ["Estrategias recebem renda no mes de acesso", null, null, null, null, null, "Em Investimento de Imovel, as cinco estrategias recebem a receita no mes habilitado."],
  ["VPL e Gap usam mesma base temporal", null, null, null, null, null, "VPL deve ser o Gap VF descontado pela taxa de oportunidade."],
  ["Inflacao anual esperada nao negativa", null, null, null, null, null, "Inflacao usada apenas para apresentacao em R$ de hoje."],
  ["Fator inflacao acumulada >= 1", null, null, null, null, null, "Com inflacao esperada nao negativa, o fator deve ser no minimo 1."],
  ["Patrimonio R$ hoje bate formula", null, null, null, null, null, "Patrimonio em R$ de hoje = patrimonio nominal / fator de inflacao, apenas se executavel."],
  ["Varreduras de build sem bloqueio", null, null, null, null, null, "O script executa busca de erros Excel, referencias antigas, retorno interno e inflacao indevida antes de salvar."],
  ["Status geral", null, null, null, null, null, "Consolida os checks acima."],
];
checks.getRange("B5:F29").formulas = [
  ["=INDEX(Financiamento!$M$6:$M$485,Premissas!$B$13)", "=0", "=B5-C5", "=1", '=IF(ABS(D5)<=E5,"OK","Revisar")'],
  ["=Premissas!$B$31", "=Premissas!$B$29", "=B6-C6", "=1", '=IF(ABS(D6)<=E6,"OK","Revisar")'],
  ["='Consorcio Direto'!$B$43", "='Consorcio Direto'!$B$4", "=B7-C7", "=1", '=IF(B7+E7>=C7,"OK","Revisar")'],
  ["='Consorcio Direto'!$B$27", `=${OPP_MAX_MONTHS}`, "=C8-B8", "=0", '=IF(B8<=C8,"OK","Revisar")'],
  ["=Premissas!$B$43", "=Premissas!$B$44", "=B9-C9", "=0", '=IF(B9>=C9,"OK","Revisar")'],
  ["=Premissas!$B$39", `=${OPP_MAX_MONTHS}`, "=C10-B10", "=0", '=IF(B10<=C10,"OK","Revisar")'],
  [minInvestedFormula, "=0", "=B11-C11", "=0", '=IF(B11>=C11,"OK","Revisar")'],
  ["=Resumo!$C$13", "='Consorcio Direto'!$B$29", "=B12-C12", "=0", '=IF(B12=C12,"OK","Revisar")'],
  ["=Premissas!$B$11", '="Imovel/Veiculo"', '=IF(OR(B13="Imovel",B13="Veiculo"),0,1)', "=0", '=IF(D13=E13,"OK","Revisar")'],
  ["=Premissas!$B$19", '="SAC/PRICE"', '=IF(OR(B14="SAC",B14="PRICE"),0,1)', "=0", '=IF(D14=E14,"OK","Revisar")'],
  ["=Premissas!$B$65", '="Nominal"', '=IF(B15=C15,0,1)', "=0", '=IF(B15=C15,"OK","Revisar")'],
  ["=Premissas!$B$92", "=-100%", "=B16-C16", "=0", '=IF(B16>C16,"OK","Revisar")'],
  [patrimonioTieFormula, "=0", "=B17-C17", "=1", '=IF(ABS(D17)<=E17,"OK","Revisar")'],
  [patrimonioStatusFormula, "=0", "=B18-C18", "=0", '=IF(B18=C18,"OK","Revisar")'],
  [deficitStatusFormula, "=0", "=B19-C19", "=0", '=IF(B19=C19,"OK","Revisar")'],
  [`=IF(OR(Premissas!$B$11<>"Imovel",Premissas!$B$47<>"Moradia"),SUM(${allPaidRanges}),0)`, "=0", "=B20-C20", "=1", '=IF(ABS(D20)<=E20,"OK","Revisar")'],
  [`=IF(OR(Premissas!$B$11<>"Imovel",Premissas!$B$47<>"Investimento"),SUM(${allReceivedRanges}),0)`, "=0", "=B21-C21", "=1", '=IF(ABS(D21)<=E21,"OK","Revisar")'],
  ['=SUMIFS(\'Motor Caixa\'!$BG$5:$BG$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(\'Consorcio Direto\'!$B$29,Premissas!$B$49))', "=0", "=B22-C22", "=1", '=IF(ABS(D22)<=E22,"OK","Revisar")'],
  ['=SUMIFS(\'Motor Caixa\'!$BW$5:$BW$484,\'Motor Caixa\'!$A$5:$A$484,"<"&MAX(\'Carta Contemplada\'!$B$44,Premissas!$B$49))', "=0", "=B23-C23", "=1", '=IF(ABS(D23)<=E23,"OK","Revisar")'],
  [investmentAccessFormula, "=0", "=B24-C24", "=1", '=IF(ABS(D24)<=E24,"OK","Revisar")'],
  [vplGapFormula, "=0", "=B25-C25", "=1", '=IF(ABS(D25)<=E25,"OK","Revisar")'],
  ["=Premissas!$B$67", "=0", "=B26-C26", "=0", '=IF(B26>=C26,"OK","Revisar")'],
  ["=Premissas!$B$69", "=1", "=B27-C27", "=0", '=IF(B27>=C27,"OK","Revisar")'],
  [patrimonioHojeFormula, "=0", "=B28-C28", "=1", '=IF(ABS(D28)<=E28,"OK","Revisar")'],
  ['="Executado"', '="Sem erros bloqueantes"', '=IF(B29="Executado",0,1)', "=0", '=IF(D29=E29,"OK","Revisar")'],
];
checks.getRange("B30:F30").formulas = [["", "", "", "", '=IF(COUNTIF(F5:F29,"Revisar")+COUNTIF(F32:F44,"Revisar")+COUNTIF(F47:F60,"Revisar")=0,"OK","Revisar")']];

checks.getRange("A31:G31").values = [["Checks adicionais de amortizacao por caixa livre", "", "", "", "", "", ""]];
checks.getRange("A31:G31").merge();
checks.getRange("A31:G31").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.teal },
};
checks.getRange("A32:G34").values = [
  ["Amortizacao extra = menor entre sobra de caixa e saldo amortizavel", null, null, null, null, null, "Valida J = MIN(sobra mensal, saldo apos amortizacao ordinaria)."],
  ["Sobra para amortizar nunca negativa", null, null, null, null, null, "Sobra = MAX(0, caixa livre - prestacao ordinaria do mes)."],
  ["Caixa livre do cenario 2 vem de Premissas!B42", null, null, null, null, null, "Extra nao usa consorcio; usa o caixa livre mensal comum."],
];
checks.getRange("B32:F34").formulas = [
  ["=MAX(MAX(Amortizacao!$Q$6:$Q$485),-MIN(Amortizacao!$Q$6:$Q$485))", "=0", "=B32-C32", "=1", '=IF(ABS(D32)<=E32,"OK","Revisar")'],
  ["=MIN(Amortizacao!$P$6:$P$485)", "=0", "=B33-C33", "=0", '=IF(B33>=C33,"OK","Revisar")'],
  ["=Premissas!$B$35", "=Premissas!$B$42", "=B34-C34", "=1", '=IF(ABS(D34)<=E34,"OK","Revisar")'],
];

checks.getRange("A36:G36").values = [["Checks da carta contemplada", "", "", "", "", "", ""]];
checks.getRange("A36:G36").merge();
checks.getRange("A36:G36").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.teal },
};
checks.getRange("A37:G44").values = [
  ["Carta: credito liquido + complemento cobre bem", null, null, null, null, null, "Credito liquido utilizavel mais complemento deve cobrir o valor do bem."],
  ["Carta: custo inicial entra no mes correto", null, null, null, null, null, "Preco da carta, transferencia e complemento entram no mes de compra."],
  ["Carta: quantidade de parcelas restantes", null, null, null, null, null, "A soma de meses com parcela deve respeitar a quantidade informada."],
  ["Carta: credito excedente nunca vira caixa", null, null, null, null, null, "O excedente e desconsiderado e nao reduz desembolso nem vira investimento."],
  ["Carta: obrigacao remanescente bate resumo", null, null, null, null, null, "Saldo devedor da estrategia 5 no resumo deve bater com a obrigacao remanescente da carta."],
  ["Carta: VP parcelas restantes visivel", null, null, null, null, null, "VP das parcelas restantes deve ficar visivel para auditoria."],
  ["Carta: credito liquido nao negativo", null, null, null, null, null, "Restricoes nao podem tornar o credito liquido negativo."],
  ["Carta: compra ativada ou desativada por lista", null, null, null, null, null, "Campo deve ser Sim ou Nao."],
];
checks.getRange("B37:F44").formulas = [
  ["='Carta Contemplada'!$B$27+'Carta Contemplada'!$B$35", "='Carta Contemplada'!$B$4", "=B37-C37", "=1", '=IF(B37+E37>=C37,"OK","Revisar")'],
  [`=INDEX('Carta Contemplada'!$E$${ccStartRow}:$E$${ccEndRow},'Carta Contemplada'!$B$23)`, "='Carta Contemplada'!$B$37", "=B38-C38", "=1", '=IF(ABS(D38)<=E38,"OK","Revisar")'],
  [`=COUNTIF('Carta Contemplada'!$C$${ccStartRow}:$C$${ccEndRow},\">0\")`, "='Carta Contemplada'!$B$31", "=B39-C39", "=0", '=IF(B39=C39,"OK","Revisar")'],
  ["='Carta Contemplada'!$B$36", "=0", "=B40-C40", "=1", '=IF(ABS(D40)<=E40,"OK","Revisar")'],
  ["=Resumo!$N$14", "='Carta Contemplada'!$B$40", "=B41-C41", "=1", '=IF(ABS(D41)<=E41,"OK","Revisar")'],
  ["='Carta Contemplada'!$B$38", "=0", "=B42-C42", "=0", '=IF(B42>=C42,"OK","Revisar")'],
  ["='Carta Contemplada'!$B$27", "=0", "=B43-C43", "=0", '=IF(B43>=C43,"OK","Revisar")'],
  ["='Carta Contemplada'!$B$22", '="Sim/Nao"', '=IF(OR(B44="Sim",B44="Nao"),0,1)', "=0", '=IF(D44=E44,"OK","Revisar")'],
];
checks.getRange("A46:G46").values = [["Checks de posicao liquida dos consorcios e preco", "", "", "", "", "", ""]];
checks.getRange("A46:G46").merge();
checks.getRange("A46:G46").format = {
  fill: colors.teal,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.teal },
};
checks.getRange("A47:G60").values = [
  ["IQ antes da contemplacao: direito e obrigacao explicitos", null, null, null, null, null, "No mes anterior a contemplacao, o direito economico e o VP das parcelas futuras devem ser positivos."],
  ["IQ na contemplacao: divida total nao e zerada", null, null, null, null, null, "A divida total inclui saldo residual do financiamento e VP das parcelas futuras do grupo."],
  ["IQ apos contemplacao: divida bate obrigacao", null, null, null, null, null, "Depois do uso do credito e antes do fim do grupo, a divida deve continuar igual ao VP das parcelas futuras."],
  ["Direto antes da contemplacao: direito e obrigacao explicitos", null, null, null, null, null, "A cota direta reconhece direito economico descontado e obrigacao em VP antes da contemplacao."],
  ["Direto na contemplacao: obrigacao permanece", null, null, null, null, null, "No mes de contemplacao, o bem substitui o direito, mas as parcelas futuras permanecem como divida."],
  ["Consorcios antes do fim do grupo mantem obrigacao", null, null, null, null, null, "No penultimo mes do grupo, ainda existe a ultima parcela futura em valor presente."],
  ["Consorcios zeram obrigacao no fim do grupo", null, null, null, null, null, "Depois da ultima parcela do grupo, a obrigacao remanescente deve ser zero."],
  ["Precos seguem regra atual/futuro", null, null, null, null, null, "Financiamento usa o valor atual; consorcio direto e carta usam o preco futuro no mes de aquisicao."],
  ["Tipo do bem seleciona taxa correta", null, null, null, null, null, "A taxa anual aplicada deve vir do bloco correto do tipo do bem."],
  ["Veiculo nao gera aluguel ou renda", null, null, null, null, null, "Para Veiculo, aluguel pago e receita devem permanecer zerados em todas as estrategias."],
  ["Lance fixo e livre sao exclusivos", null, null, null, null, null, "As duas modalidades de lance total nao podem estar ativas simultaneamente."],
  ["Lance embutido exige lance total", null, null, null, null, null, "O embutido nao pode existir sem lance fixo ou livre."],
  ["Lance embutido nao supera total", null, null, null, null, null, "O percentual embutido efetivo deve ser menor ou igual ao lance total."],
  ["Carta nao libera credito excedente", null, null, null, null, null, "Credito excedente deve permanecer fora do caixa e do patrimonio financeiro."],
];
checks.getRange("B47:F60").formulas = [
  ['=MIN(INDEX(\'Consorcio IQ\'!$N$6:$N$485,MAX(1,Premissas!$B$27-1)),INDEX(\'Consorcio IQ\'!$O$6:$O$485,MAX(1,Premissas!$B$27-1)))', "=0", "=B47-C47", "=0", '=IF(B47>0,"OK","Revisar")'],
  ['=INDEX(\'Consorcio IQ\'!$P$6:$P$485,Premissas!$B$27)', '=INDEX(\'Consorcio IQ\'!$N$6:$N$485,Premissas!$B$27)+MAX(0,INDEX(\'Consorcio IQ\'!$H$6:$H$485,Premissas!$B$27)-INDEX(\'Consorcio IQ\'!$I$6:$I$485,Premissas!$B$27))', "=B48-C48", "=1", '=IF(AND(ABS(D48)<=E48,OR(Premissas!$B$27>=Premissas!$B$25,B48>0)),"OK","Revisar")'],
  ['=INDEX(\'Consorcio IQ\'!$P$6:$P$485,MIN(Premissas!$B$25-1,Premissas!$B$27+1))', '=INDEX(\'Consorcio IQ\'!$N$6:$N$485,MIN(Premissas!$B$25-1,Premissas!$B$27+1))', "=B49-C49", "=1", '=IF(ABS(D49)<=E49,"OK","Revisar")'],
  ['=MIN(INDEX(\'Consorcio Direto\'!$K$55:$K$534,MAX(1,\'Consorcio Direto\'!$B$29-1)),INDEX(\'Consorcio Direto\'!$L$55:$L$534,MAX(1,\'Consorcio Direto\'!$B$29-1)))', "=0", "=B50-C50", "=0", '=IF(B50>0,"OK","Revisar")'],
  ['=INDEX(\'Consorcio Direto\'!$K$55:$K$534,\'Consorcio Direto\'!$B$29)', "=0", "=B51-C51", "=0", '=IF(OR(\'Consorcio Direto\'!$B$29>=\'Consorcio Direto\'!$B$27,B51>0),\"OK\",\"Revisar\")'],
  ['=MIN(INDEX(\'Consorcio IQ\'!$N$6:$N$485,Premissas!$B$25-1),INDEX(\'Consorcio Direto\'!$K$55:$K$534,\'Consorcio Direto\'!$B$27-1))', "=0", "=B52-C52", "=0", '=IF(B52>0,"OK","Revisar")'],
  ['=MAX(INDEX(\'Consorcio IQ\'!$N$6:$N$485,Premissas!$B$25),INDEX(\'Consorcio Direto\'!$K$55:$K$534,\'Consorcio Direto\'!$B$27))', "=0", "=B53-C53", "=1", '=IF(ABS(D53)<=E53,"OK","Revisar")'],
  ['=MAX(ABS(Premissas!$B$79-Premissas!$B$6),ABS(\'Consorcio Direto\'!$B$4-MAX(0,Premissas!$B$6*(1+Premissas!$B$66)^((\'Consorcio Direto\'!$B$29-1)/12))),ABS(\'Carta Contemplada\'!$B$4-MAX(0,Premissas!$B$6*(1+Premissas!$B$66)^((\'Carta Contemplada\'!$B$44-1)/12))))', "=0", "=B54-C54", "=1", '=IF(ABS(D54)<=E54,"OK","Revisar")'],
  ['=ABS(Premissas!$B$66-IF(Premissas!$B$11="Veiculo",Premissas!$B$92,Premissas!$B$91))', "=0", "=B55-C55", "=1", '=IF(ABS(D55)<=E55,"OK","Revisar")'],
  [`=IF(Premissas!$B$11="Veiculo",SUM(${allPaidRanges})+SUM(${allReceivedRanges}),0)`, "=0", "=B56-C56", "=1", '=IF(ABS(D56)<=E56,"OK","Revisar")'],
  ['=IF(AND(\'Consorcio Direto\'!$B$33="Sim",\'Consorcio Direto\'!$B$35="Sim"),1,0)', "=0", "=B57-C57", "=0", '=IF(D57=E57,"OK","Revisar")'],
  ['=IF(AND(\'Consorcio Direto\'!$B$31="Sim",\'Consorcio Direto\'!$B$38=0),1,0)', "=0", "=B58-C58", "=0", '=IF(D58=E58,"OK","Revisar")'],
  ["=MAX(0,'Consorcio Direto'!$B$37-'Consorcio Direto'!$B$38)", "=0", "=B59-C59", "=0.000001", '=IF(ABS(D59)<=E59,"OK","Revisar")'],
  ["='Carta Contemplada'!$B$36", "=0", "=B60-C60", "=1", '=IF(ABS(D60)<=E60,"OK","Revisar")'],
];
frame(checks, "A5:G60");
applyNumberFormat(checks, "B5:E7", currencyFmt);
applyNumberFormat(checks, "B8:E8", monthFmt);
applyNumberFormat(checks, "B9:E9", currencyFmt);
applyNumberFormat(checks, "B10:E10", monthFmt);
applyNumberFormat(checks, "B11:E12", currencyFmt);
applyNumberFormat(checks, "B16:E16", pctFmt);
applyNumberFormat(checks, "B17:E25", currencyFmt);
applyNumberFormat(checks, "B26:E26", pctFmt);
applyNumberFormat(checks, "B27:E27", "0.0000x");
applyNumberFormat(checks, "B28:E28", currencyFmt);
applyNumberFormat(checks, "B32:E34", currencyFmt);
applyNumberFormat(checks, "B37:E43", currencyFmt);
applyNumberFormat(checks, "B47:E53", currencyFmt);
applyNumberFormat(checks, "B54:E56", currencyFmt);
applyNumberFormat(checks, "B59:E59", pctFmt);
applyNumberFormat(checks, "B60:E60", currencyFmt);
checks.getRange("A5:A60").format.wrapText = true;
checks.getRange("G5:G60").format.wrapText = true;

// Fontes / audit trail
setTitle(fontes, "A1:E1", "Fontes e trilha de auditoria");
setWidths(fontes, [270, 150, 110, 260, 430], 85);
fontes.getRange("A4:E4").values = [["Item", "Valor", "Unidade", "Fonte", "Observacoes"]];
header(fontes, "A4:E4", colors.teal);
const sourceRows = [
  ["Valor de venda do bem atual", "=Premissas!$B$5", "R$", "Premissa informada pelo usuario", "Usado apenas como liquidez inicial comum aos cenarios."],
  ["Valor do bem novo", "=Premissas!$B$6", "R$", "Premissa informada pelo usuario", "Base de entrada e financiamento."],
  ["Entrada", "=Premissas!$B$7", "%", "Premissa informada pelo usuario", "Percentual do valor do bem novo."],
  ["Tipo do bem", "=Premissas!$B$11", "texto", "Premissa informada pelo usuario", "Lista validada: Imovel ou Veiculo."],
  ["Sistema de financiamento", "=Premissas!$B$19", "texto", "Premissa informada pelo usuario", "Lista validada: SAC ou PRICE."],
  ["Juros financiamento", "=Premissas!$B$14", "% a.a.", "Premissa informada pelo usuario", "Convertido para taxa mensal equivalente."],
  ["Indexador/correcao do saldo", "=Premissas!$B$16", "% a.a.", "Premissa informada pelo usuario", "Indice contratual generico convertido para taxa mensal equivalente."],
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
  ["Caixa livre mensal antes dos fluxos do bem", "=Premissas!$B$42", "R$/mes", "Premissa informada pelo usuario", "Caixa livre mensal antes de entrada, parcelas, aluguel de espera, renda, consorcio, lances e carta."],
  ["Valor em reserva", "=Premissas!$B$43", "R$", "Premissa informada pelo usuario", "Deve ser no minimo a maior entrada entre as estrategias."],
  ["Maior entrada entre estrategias", "=Premissas!$B$44", "R$", "Calculo do modelo", "Usado para checar se a reserva inicial e suficiente."],
  ["Mes de contemplacao consorcio direto", "='Consorcio Direto'!$B$29", "mes", "Consorcio Direto", "Mes em que o credito fica disponivel na estrategia 4."],
  ["Carta inicial consorcio direto", "='Consorcio Direto'!$B$41", "R$", "Consorcio Direto", "Dimensionada para que o credito liquido cubra o preco aplicavel do bem."],
  ["Credito liquido consorcio direto", "='Consorcio Direto'!$B$43", "R$", "Consorcio Direto", "Credito recebido apos lance embutido."],
  ["Lance com recursos do cliente consorcio direto", "='Consorcio Direto'!$B$44", "R$", "Consorcio Direto", "Descapitalizacao efetiva do cliente quando ha lances combinados."],
  ["Parcela inicial consorcio direto", "='Consorcio Direto'!$B$45", "R$/mes", "Consorcio Direto", "Primeira parcela antes dos reajustes anuais."],
  ["Valor do bem no horizonte", "=Premissas!$B$45", "R$", "Calculo do modelo", "Considera valorizacao/depreciacao editavel e piso zero para valor residual."],
  ["Finalidade do Imovel", "=Premissas!$B$47", "texto", "Premissa editavel", "Moradia ou Investimento; ignorada quando Tipo do bem = Veiculo."],
  ["Mes aquisicao financiamento", "=Premissas!$B$48", "mes", "Convencao do modelo", "Financiamento e IQ compram no mes inicial."],
  ["Mes disponibilidade/uso", "=Premissas!$B$49", "mes", "Premissa editavel", "Controla quando uso proprio ou renda podem ocorrer."],
  ["Aluguel pago inicial", "=Premissas!$B$54", "R$/mes", "Calculo do modelo", "Moradia de Imovel: aluguel opcional ate a estrategia permitir usar o bem."],
  ["Receita liquida inicial", "=Premissas!$B$64", "R$/mes", "Calculo do modelo", "Investimento de Imovel: receita opcional apenas apos aquisicao e disponibilidade."],
  ["Modelo nominal", "=Premissas!$B$65", "texto", "Premissa editavel", "A validacao permite apenas Nominal enquanto nao houver logica real."],
  ["Taxa anual aplicada ao bem", "=Premissas!$B$66", "% a.a.", "Calculo do modelo", "Selecionada automaticamente conforme Imovel ou Veiculo."],
  ["Valorizacao anual do Imovel", "=Premissas!$B$91", "% a.a.", "Premissa editavel", "Aplicada quando Tipo do bem = Imovel."],
  ["Depreciacao anual do Veiculo", "=Premissas!$B$92", "% a.a.", "Premissa editavel", "Aplicada quando Tipo do bem = Veiculo; deve ser maior que -100%."],
  ["Regra do preco na aquisicao", "=Premissas!$B$78", "texto", "Convencao do modelo", "Financiamento atual; consorcios pelo preco futuro."],
  ["Preco do financiamento na aquisicao", "=Premissas!$B$79", "R$", "Calculo do modelo", "Base da entrada e do valor financiado."],
  ["Inflacao anual esperada", "=Premissas!$B$67", "% a.a.", "Premissa editavel", "Usada apenas para converter valores nominais futuros em R$ de hoje."],
  ["Inflacao mensal equivalente", "=Premissas!$B$68", "% a.m.", "Calculo do modelo", "Equivalente mensal composto da inflacao esperada."],
  ["Fator inflacao acumulada no horizonte", "=Premissas!$B$69", "x", "Calculo do modelo", "Fator usado apenas nas colunas de apresentacao do Resumo."],
  ["Observacao sobre inflacao", '=Premissas!$D$67', "texto", "Convencao do modelo", "A inflacao nao altera o Motor Caixa nem as logicas de financiamento, consorcio, renda/custos do bem, deficit nao coberto, VPL ou Gap."],
  ["FV base sem compra", "=INDEX('Motor Caixa'!$G$5:$G$484,Premissas!$B$39)", "R$", "Motor Caixa", "Reserva e capacidade mensal investidas a taxa de oportunidade."],
  ["IQ: obrigacao no horizonte", "=INDEX('Consorcio IQ'!$N$6:$N$485,Premissas!$B$39)", "R$", "Consorcio IQ", "VP das parcelas futuras depois da parcela do mes do horizonte."],
  ["IQ: direito economico no horizonte", "=INDEX('Consorcio IQ'!$O$6:$O$485,Premissas!$B$39)", "R$", "Consorcio IQ", "Direito conservador antes da contemplacao; zera quando o credito e utilizado."],
  ["IQ: divida total no horizonte", "=INDEX('Consorcio IQ'!$P$6:$P$485,Premissas!$B$39)", "R$", "Consorcio IQ", "Saldo residual do financiamento mais VP das parcelas futuras."],
  ["Direto: obrigacao no horizonte", "=INDEX('Consorcio Direto'!$K$55:$K$534,Premissas!$B$39)", "R$", "Consorcio Direto", "VP das parcelas futuras depois da parcela do mes do horizonte."],
  ["Direto: direito economico no horizonte", "=INDEX('Consorcio Direto'!$L$55:$L$534,Premissas!$B$39)", "R$", "Consorcio Direto", "Direito conservador antes da contemplacao; zera quando o bem e adquirido."],
  ["Carta contemplada ativa", "='Carta Contemplada'!$B$22", "texto", "Carta Contemplada", "Controla a estrategia 5."],
  ["Mes de compra da carta", "='Carta Contemplada'!$B$23", "mes", "Carta Contemplada", "Mes de transferencia e pagamento do custo inicial."],
  ["Mes de uso do credito da carta", "='Carta Contemplada'!$B$24", "mes", "Carta Contemplada", "Mes em que o credito fica disponivel, sujeito a aprovacao da administradora."],
  ["Credito bruto da carta", "='Carta Contemplada'!$B$25", "R$", "Carta Contemplada", "Credito atualizado informado para a carta contemplada."],
  ["Credito liquido utilizavel carta", "='Carta Contemplada'!$B$27", "R$", "Carta Contemplada", "Credito bruto menos restricoes do credito."],
  ["Preco pago pela carta", "='Carta Contemplada'!$B$28", "R$", "Carta Contemplada", "Preco efetivo pago pela cessao da carta."],
  ["Taxa de transferencia carta", "='Carta Contemplada'!$B$30", "R$", "Carta Contemplada", "Taxas de transferencia/cessao e custos administrativos."],
  ["Parcelas restantes carta", "='Carta Contemplada'!$B$31", "meses", "Carta Contemplada", "Quantidade de parcelas assumidas pelo comprador."],
  ["Parcela atual carta", "='Carta Contemplada'!$B$32", "R$/mes", "Carta Contemplada", "Parcela vigente no mes de compra."],
  ["Reajuste anual carta", "='Carta Contemplada'!$B$33", "% a.a.", "Carta Contemplada", "Reajuste anual por degrau das parcelas restantes."],
  ["Complemento necessario carta", "='Carta Contemplada'!$B$35", "R$", "Carta Contemplada", "Complemento de recursos proprios se credito liquido nao cobrir o bem."],
  ["Custo inicial total carta", "='Carta Contemplada'!$B$37", "R$", "Carta Contemplada", "Preco pago + transferencia + complemento; excedente nao reduz o custo."],
  ["VP parcelas restantes carta", "='Carta Contemplada'!$B$38", "R$", "Carta Contemplada", "Valor presente das parcelas restantes pela taxa de oportunidade."],
  ["Custo economico carta", "='Carta Contemplada'!$B$39", "R$", "Carta Contemplada", "Preco + transferencia + VP parcelas - credito liquido."],
  ["Obrigacao carta no horizonte", "='Carta Contemplada'!$B$40", "R$", "Carta Contemplada", "Obrigacao remanescente descontada no horizonte."],
  ["Patrimonio estrategia carta", "=INDEX('Motor Caixa'!$CG$5:$CG$484,Premissas!$B$39)", "R$", "Motor Caixa", "Patrimonio bruto da quinta estrategia antes do filtro de inviabilidade."],
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
setTitle(resumo, "A1:P1", "Resumo executivo - comparativo de estrategias");
setWidths(resumo, [250, 132, 95, 95, 138, 138, 138, 138, 148, 148, 148, 148, 138, 138, 132, 132, 118, 265, 24, 150, 150, 150, 150, 150, 150, 150], 90);

resumo.getRange("A3:N6").values = [
  ["Checks do modelo", null, "", "Tipo do bem", null, "", "Sistema financ.", null, "", "Caixa livre mensal", null, "", "Rentabilidade a.a.", null],
  ["Valor do bem", null, "", "Finalidade do bem", null, "", "Horizonte", null, "", "Valor em reserva", null, "", "FV base sem compra", null],
  ["Entrada financiamento", null, "", "Disponibilidade/uso", null, "", "Mes contemplacao IQ", null, "", "Maior entrada/carta", null, "", "Valoriz./deprec. a.a.", null],
  ["Modelo / preco", null, "", "Taxas consorcio", null, "", "Carta IQ atualizada", null, "", "Custo inicial carta", null, "", "Prazo consorcio", null],
];
resumo.getRange("B3:B6").formulas = [
  ["=Checks!$F$30"],
  ["=Premissas!$B$6"],
  ["=Premissas!$B$8"],
  ['=Premissas!$B$65&" | "&Premissas!$B$78'],
];
resumo.getRange("E3:E6").formulas = [
  ["=Premissas!$B$11"],
  ["=Premissas!$B$47"],
  ["=Premissas!$B$49"],
  ["=Premissas!$B$24"],
];
resumo.getRange("H3:H6").formulas = [
  ["=Premissas!$B$19"],
  ["=Premissas!$B$39"],
  ["=Premissas!$B$27"],
  ["=Premissas!$B$31"],
];
resumo.getRange("K3:K6").formulas = [
  ["=Premissas!$B$42"],
  ["=Premissas!$B$43"],
  ["=Premissas!$B$44"],
  ["='Carta Contemplada'!$B$37"],
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
applyNumberFormat(resumo, "E5", monthFmt);
applyNumberFormat(resumo, "E6", pctFmt);
applyNumberFormat(resumo, "H4:H5", monthFmt);
applyNumberFormat(resumo, "H6", currencyFmt);
applyNumberFormat(resumo, "K3:K6", currencyFmt);
applyNumberFormat(resumo, "N3", pctFmt);
applyNumberFormat(resumo, "N4", currencyFmt);
applyNumberFormat(resumo, "N5", pctFmt);
applyNumberFormat(resumo, "N6", monthFmt);

section(resumo, "A8:P8", "Comparativo principal");
resumo.getRange("A9:P9").values = [[
  "Estrategia",
  "Entrada/lance inicial",
  "Mes aquis./cont.",
  "Mes fim obrig./quit.",
  "Saidas nominais totais",
  "Entradas nominais totais",
  "Fluxo liquido nominal acum.",
  "Deficit nao coberto VF",
  "Patrimonio nominal",
  "Patrimonio em R$ de hoje",
  "VPL vs base",
  "Gap VF vs base",
  "Saldo investido",
  "Saldo devedor",
  "Status",
  "Observacao",
]];
header(resumo, "A9:P9");
resumo.getRange("A10:A14").values = [
  ["1. Financiamento sem amortizacao"],
  ["2. Financiamento com amortizacao"],
  ["3. Financiamento + consorcio IQ"],
  ["4. Consorcio direto"],
  ["5. Compra de carta contemplada"],
];
resumo.getRange("B10:P14").formulas = [
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    "=Premissas!$B$48+Premissas!$B$13-1",
    "=SUMIFS('Motor Caixa'!$L$5:$L$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$M$5:$M$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F10-E10",
    "=INDEX('Motor Caixa'!$Q$5:$Q$484,Premissas!$B$39)",
    '=IF(H10>1,"Inviavel",INDEX(\'Motor Caixa\'!$U$5:$U$484,Premissas!$B$39))',
    '=IF(H10>1,"Inviavel",INDEX(\'Motor Caixa\'!$U$5:$U$484,Premissas!$B$39)/Premissas!$B$69)',
    '=IF(H10>1,"Inviavel",(INDEX(\'Motor Caixa\'!$U$5:$U$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39)',
    '=IF(H10>1,"Inviavel",INDEX(\'Motor Caixa\'!$U$5:$U$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))',
    "=INDEX('Motor Caixa'!$R$5:$R$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$S$5:$S$484,Premissas!$B$39)",
    '=IF(H10>1,"Inviavel","OK")',
    '"Compra por financiamento no mes 1 pelo valor atual; aluguel ou renda existem apenas para Imovel conforme a finalidade."',
  ],
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    '=Premissas!$B$48+MATCH("Quitado",Amortizacao!$N$6:$N$485,0)-1',
    "=SUMIFS('Motor Caixa'!$AB$5:$AB$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$AC$5:$AC$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F11-E11",
    "=INDEX('Motor Caixa'!$AG$5:$AG$484,Premissas!$B$39)",
    '=IF(H11>1,"Inviavel",INDEX(\'Motor Caixa\'!$AK$5:$AK$484,Premissas!$B$39))',
    '=IF(H11>1,"Inviavel",INDEX(\'Motor Caixa\'!$AK$5:$AK$484,Premissas!$B$39)/Premissas!$B$69)',
    '=IF(H11>1,"Inviavel",(INDEX(\'Motor Caixa\'!$AK$5:$AK$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39)',
    '=IF(H11>1,"Inviavel",INDEX(\'Motor Caixa\'!$AK$5:$AK$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))',
    "=INDEX('Motor Caixa'!$AH$5:$AH$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AI$5:$AI$484,Premissas!$B$39)",
    '=IF(H11>1,"Inviavel","OK")',
    '"Pagamento extra usa a sobra entre caixa livre mensal e prestacao ordinaria do sistema selecionado."',
  ],
  [
    "=Premissas!$B$8",
    "=Premissas!$B$48",
    "=Premissas!$B$25",
    "=SUMIFS('Motor Caixa'!$AR$5:$AR$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$AS$5:$AS$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F12-E12",
    "=INDEX('Motor Caixa'!$AW$5:$AW$484,Premissas!$B$39)",
    '=IF(H12>1,"Inviavel",INDEX(\'Motor Caixa\'!$BA$5:$BA$484,Premissas!$B$39))',
    '=IF(H12>1,"Inviavel",INDEX(\'Motor Caixa\'!$BA$5:$BA$484,Premissas!$B$39)/Premissas!$B$69)',
    '=IF(H12>1,"Inviavel",(INDEX(\'Motor Caixa\'!$BA$5:$BA$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39)',
    '=IF(H12>1,"Inviavel",INDEX(\'Motor Caixa\'!$BA$5:$BA$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))',
    "=INDEX('Motor Caixa'!$AX$5:$AX$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$AY$5:$AY$484,Premissas!$B$39)",
    '=IF(H12>1,"Inviavel","OK")',
    '"Financiamento compra no inicio; a carta quita o saldo no mes de contemplacao, mas o VP das parcelas do consorcio permanece como obrigacao."',
  ],
  [
    "='Consorcio Direto'!$B$44",
    "='Consorcio Direto'!$B$29",
    "='Consorcio Direto'!$B$27",
    "=SUMIFS('Motor Caixa'!$BH$5:$BH$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$BI$5:$BI$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F13-E13",
    "=INDEX('Motor Caixa'!$BM$5:$BM$484,Premissas!$B$39)",
    '=IF(H13>1,"Inviavel",INDEX(\'Motor Caixa\'!$BQ$5:$BQ$484,Premissas!$B$39))',
    '=IF(H13>1,"Inviavel",INDEX(\'Motor Caixa\'!$BQ$5:$BQ$484,Premissas!$B$39)/Premissas!$B$69)',
    '=IF(H13>1,"Inviavel",(INDEX(\'Motor Caixa\'!$BQ$5:$BQ$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39)',
    '=IF(H13>1,"Inviavel",INDEX(\'Motor Caixa\'!$BQ$5:$BQ$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))',
    "=INDEX('Motor Caixa'!$BN$5:$BN$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$BO$5:$BO$484,Premissas!$B$39)",
    '=IF(H13>1,"Inviavel","OK")',
    '"Antes da contemplacao reconhece direito da cota e obrigacao; depois reconhece o bem pelo preco aplicavel e mantem o VP das parcelas futuras."',
  ],
  [
    "='Carta Contemplada'!$B$37",
    "='Carta Contemplada'!$B$44",
    "='Carta Contemplada'!$B$23+'Carta Contemplada'!$B$31-1",
    "=SUMIFS('Motor Caixa'!$BX$5:$BX$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=SUMIFS('Motor Caixa'!$BY$5:$BY$484,'Motor Caixa'!$A$5:$A$484,\"<=\"&Premissas!$B$39)",
    "=F14-E14",
    "=INDEX('Motor Caixa'!$CC$5:$CC$484,Premissas!$B$39)",
    '=IF(H14>1,"Inviavel",INDEX(\'Motor Caixa\'!$CG$5:$CG$484,Premissas!$B$39))',
    '=IF(H14>1,"Inviavel",INDEX(\'Motor Caixa\'!$CG$5:$CG$484,Premissas!$B$39)/Premissas!$B$69)',
    '=IF(H14>1,"Inviavel",(INDEX(\'Motor Caixa\'!$CG$5:$CG$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))/(1+Premissas!$B$41)^Premissas!$B$39)',
    '=IF(H14>1,"Inviavel",INDEX(\'Motor Caixa\'!$CG$5:$CG$484,Premissas!$B$39)-INDEX(\'Motor Caixa\'!$G$5:$G$484,Premissas!$B$39))',
    "=INDEX('Motor Caixa'!$CD$5:$CD$484,Premissas!$B$39)",
    "=INDEX('Motor Caixa'!$CE$5:$CE$484,Premissas!$B$39)",
    '=IF(H14>1,"Inviavel","OK")',
    '"Compra carta ja contemplada; considera preco aplicavel, custo inicial, credito liquido, complemento, parcelas e obrigacao remanescente."',
  ],
];
frame(resumo, "A10:P14");
applyNumberFormat(resumo, "B10:B14", currencyFmt);
applyNumberFormat(resumo, "C10:D14", monthFmt);
applyNumberFormat(resumo, "E10:N14", currencyFmt);
resumo.getRange("I10:J14").format.fill = colors.paleGreen;
resumo.getRange("K10:L14").format.fill = colors.paleBlue;
resumo.getRange("O10:P14").format.wrapText = true;

section(resumo, "A16:C16", "Grafico patrimonio corrigido", colors.teal);
resumo.getRange("A17:C22").values = [
  ["Estrategia", "Patrimonio corrigido", "Dif. vs melhor"],
  ["Fin. sem amort.", null],
  ["Fin. amort.", null],
  ["Fin. + IQ", null],
  ["Consorcio direto", null],
  ["Carta contemplada", null],
];
resumo.getRange("B18:C22").formulas = [
  ['=IF($O$10="Inviavel","",$J$10)', '=IF(B18="","",B18-MAX($B$18:$B$22))'],
  ['=IF($O$11="Inviavel","",$J$11)', '=IF(B19="","",B19-MAX($B$18:$B$22))'],
  ['=IF($O$12="Inviavel","",$J$12)', '=IF(B20="","",B20-MAX($B$18:$B$22))'],
  ['=IF($O$13="Inviavel","",$J$13)', '=IF(B21="","",B21-MAX($B$18:$B$22))'],
  ['=IF($O$14="Inviavel","",$J$14)', '=IF(B22="","",B22-MAX($B$18:$B$22))'],
];
header(resumo, "A17:C17", colors.teal);
frame(resumo, "A18:C22");
resumo.getRange("B18:B22").format.fill = colors.paleGreen;
resumo.getRange("C18:C22").format.fill = colors.paleBlue;
applyNumberFormat(resumo, "B18:C22", currencyFmt);

const correctedChart = resumo.charts.add("bar", resumo.getRange("A17:B22"));
correctedChart.title = "Patrimonio em R$ de hoje por estrategia";
correctedChart.hasLegend = false;
correctedChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
correctedChart.yAxis = { numberFormatCode: 'R$ #,##0' };
correctedChart.series.items[0].fill = colors.teal;
correctedChart.setPosition("T3", "Z24");

resumo.getRange("A24:P40").values = [
  ["Definicoes", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Saidas nominais totais", "Soma bruta das saidas do motor: entrada, parcelas, seguros, consorcio, lances e aluguel de espera quando aplicavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Entradas nominais totais", "Soma bruta das entradas do motor, principalmente renda liquida de Imovel em Investimento.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Fluxo liquido nominal acum.", "Entradas nominais totais menos saidas nominais totais. Nao substitui VPL nem Gap.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio nominal", "Mostra o patrimonio no horizonte apenas quando a estrategia nao tem deficit nao coberto. Se houver deficit, mostra Inviavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Patrimonio em R$ de hoje", "Patrimonio nominal convertido para R$ de hoje pelo fator de inflacao acumulada. Se a estrategia for inviavel, tambem mostra Inviavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Inflacao esperada", "Premissa editavel usada apenas para converter valores nominais futuros em R$ de hoje. Nao altera financiamento, consorcio, renda/custos do bem, deficit nao coberto ou motor.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Modelo nominal", "Financiamento, consorcio, renda/custos do bem, valorizacao/depreciacao e taxa de oportunidade permanecem em base nominal.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Gap VF vs base", "Para estrategias executaveis, patrimonio nominal menos o valor futuro da alternativa base de investir reserva e caixa livre mensal. Estrategias inviaveis mostram Inviavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["VPL vs base", "Para estrategias executaveis, Gap VF descontado pela taxa de oportunidade nominal. Estrategias inviaveis mostram Inviavel.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Metricas principais", "Usar Status de caixa primeiro. Se Status = Inviavel, patrimonio, VPL e Gap nao devem ser usados como recomendacao.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Tipo do bem", "Imovel usa valorizacao e pode ter Moradia ou Investimento. Veiculo usa depreciacao, nao gera aluguel/renda e tem valor residual com piso zero.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Sistema de financiamento", "SAC divide o saldo corrigido pelos meses restantes. PRICE recalcula a prestacao financeira pelo saldo corrigido e prazo remanescente, preservando amortizacao, juros, seguro e saldo final nas mesmas colunas.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Finalidade do Imovel", "Moradia aplica aluguel pago antes do acesso. Investimento aplica receita liquida depois do acesso. Para Veiculo, ambos os fluxos ficam zerados.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Preco na aquisicao", "Financiamento compra no mes 1 pelo valor atual. Consorcio direto e carta contemplada compram pelo preco futuro no mes de contemplacao ou uso.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Posicao dos consorcios", "IQ e consorcio direto reconhecem direito economico descontado e VP das parcelas futuras antes da contemplacao. Depois do uso do credito, o direito sai e a obrigacao permanece ate o fim do grupo.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Carta Contemplada", "Estrategia propria com preco aplicavel, custo inicial da carta, taxa de transferencia, credito liquido, complemento, parcelas restantes e obrigacao remanescente no patrimonio.", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];
resumo.getRange("A24:P24").merge();
resumo.getRange("A24:P24").format = {
  fill: colors.navy,
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "thin", color: colors.navy },
};
resumo.getRange("B25:P40").merge(true);
frame(resumo, "A25:P40", colors.paleGray);
resumo.getRange("B25:P40").format.wrapText = true;

// Final formatting touches
for (const sheet of sheets) {
  const used = sheet.getUsedRange();
  if (used) {
    used.format.font = { name: "Calibri", size: 10, color: colors.text };
  }
}
// Reapply title/header styles after workbook-wide font normalization.
setTitle(premissas, "A1:E1", "Simulador - Consorcio x Financiamento");
setTitle(financiamento, "A1:M1", "Cronograma do financiamento SAC/PRICE");
setTitle(consorcio, "A1:P1", "Cronograma - consorcio como interveniente quitante");
setTitle(amortizacao, "A1:Q1", "Cronograma - amortizacao de prazo com caixa livre");
setTitle(checks, "A1:G1", "Checks do modelo");
setTitle(fontes, "A1:E1", "Fontes e trilha de auditoria");
setTitle(motor, "A1:CI1", "Motor mensal de caixa e patrimonio");
setTitle(cartaContemplada, "A1:K1", "Carta Contemplada - parametros e cronograma");
setTitle(consorcioDireto, "A1:M1", "Consorcio Direto - parametros e cronograma");
setTitle(resumo, "A1:P1", "Resumo executivo - comparativo de estrategias");
header(motor, "A4:CI4", colors.teal);
header(cartaContemplada, "A54:K54", colors.teal);
header(cartaContemplada, "G4:J4", colors.teal);
header(consorcioDireto, "A54:M54", colors.teal);
header(consorcioDireto, "G4:J4", colors.teal);
header(financiamento, "A5:M5");
header(consorcio, "A5:P5");
header(amortizacao, "A5:Q5");
header(checks, "A4:G4", colors.teal);
header(fontes, "A4:E4", colors.teal);
header(resumo, "A9:P9");
header(resumo, "A17:C17", colors.teal);

// Restore input/formula styling after normalization.
inputStyle(premissas, "B4:B7");
inputStyle(premissas, "B11");
inputStyle(premissas, "B13:B18");
inputStyle(premissas, "B19");
inputStyle(premissas, "B21:B23");
inputStyle(premissas, "B25:B27");
inputStyle(premissas, "B39:B40");
inputStyle(premissas, "B42:B43");
inputStyle(premissas, "B47");
inputStyle(premissas, "B49");
inputStyle(premissas, "B53");
inputStyle(premissas, "B55");
inputStyle(premissas, "B57:B63");
inputStyle(premissas, "B65");
inputStyle(premissas, "B67");
inputStyle(premissas, "B91:B92");
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
formulaStyle(premissas, "B66:B69");
formulaStyle(premissas, "B72:B77");
formulaStyle(premissas, "B78");
formulaStyle(premissas, "B79");
sourceLinkStyle(premissas, "B29:B31");
sourceLinkStyle(premissas, "B35");
sourceLinkStyle(premissas, "B44:B45");
sourceLinkStyle(premissas, "B66");
inputStyle(cartaContemplada, "B22:B26");
inputStyle(cartaContemplada, "B28");
inputStyle(cartaContemplada, "B30:B33");
formulaStyle(cartaContemplada, "B4:B9");
formulaStyle(cartaContemplada, "B27");
formulaStyle(cartaContemplada, "B29");
formulaStyle(cartaContemplada, "B34:B44");
inputStyle(consorcioDireto, "B29");
inputStyle(consorcioDireto, "B31:B36");
formulaStyle(consorcioDireto, "B4:B9");
formulaStyle(consorcioDireto, "B23:B28");
formulaStyle(consorcioDireto, "B37:B45");

await fs.mkdir(outputDir, { recursive: true });

const finalSheetNames = [
  "Resumo",
  "Premissas",
  "Carta Contemplada",
  "Consorcio Direto",
  "Checks",
  "Fontes",
  "Motor Caixa",
  "Financiamento",
  "Consorcio IQ",
  "Amortizacao",
];
console.log(`FINAL_SHEETS ${finalSheetNames.join(" | ")}`);

const inspectSummary = await workbook.inspect({
  kind: "table",
  range: "Resumo!A8:P14",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 16,
  maxChars: 6000,
});
console.log("INSPECT_SUMMARY");
console.log(inspectSummary.ndjson);

const inspectAssumptions = await workbook.inspect({
  kind: "table",
  range: "Premissas!A4:B92",
  include: "values,formulas",
  tableMaxRows: 31,
  tableMaxCols: 2,
  maxChars: 5000,
});
console.log("INSPECT_ASSUMPTIONS");
console.log(inspectAssumptions.ndjson);

const inspectCartaContemplada = await workbook.inspect({
  kind: "table",
  range: "Carta Contemplada!A22:K60",
  include: "values,formulas",
  tableMaxRows: 39,
  tableMaxCols: 11,
  maxChars: 8000,
});
console.log("INSPECT_CARTA_CONTEMPLADA");
console.log(inspectCartaContemplada.ndjson);

const inspectConsorcioDireto = await workbook.inspect({
  kind: "table",
  range: "Consorcio Direto!A22:M60",
  include: "values,formulas",
  tableMaxRows: 39,
  tableMaxCols: 13,
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
  range: "Motor Caixa!A4:CI10",
  include: "values,formulas",
  tableMaxRows: 7,
  tableMaxCols: 87,
  maxChars: 12000,
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
  searchTerm: "\\bTIR\\b|IRR\\s*\\(|XIRR\\s*\\(|MIRR\\s*\\(|Custo efetivo|\\bCET\\b",
  options: { useRegex: true, maxResults: 300 },
  summary: "internal return metric scan",
  maxChars: 5000,
});
console.log("INTERNAL_RETURN_REFERENCES");
console.log(internalReturnReferences.ndjson);

const inflationSearchTerm = "Premissas!\\$B\\$67|Premissas!\\$B\\$68|Premissas!\\$B\\$69|Inflacao|inflacao";
const motorInflationReferences = await workbook.inspect({
  kind: "match",
  range: "Motor Caixa!A1:CI484",
  searchTerm: inflationSearchTerm,
  options: { useRegex: true, maxResults: 300 },
  summary: "motor inflation reference scan",
  maxChars: 5000,
});
console.log("MOTOR_INFLATION_REFERENCES");
console.log(motorInflationReferences.ndjson);

const vplInflationReferences = await workbook.inspect({
  kind: "match",
  range: "Resumo!K10:K14",
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

if (inspectHasMatch(errors)) {
  throw new Error("A varredura final encontrou erro de formula Excel bloqueante.");
}

const centralInflationReferenceScans = [];
for (const range of [
  "Financiamento!A1:M485",
  "Consorcio IQ!A1:P485",
  "Consorcio Direto!A1:M534",
  `Carta Contemplada!A1:K${ccEndRow}`,
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
    assetType: premissas.getRange("B11").values[0][0],
    use: premissas.getRange("B47").values[0][0],
    availability: premissas.getRange("B49").values[0][0],
    iqContemplation: premissas.getRange("B27").values[0][0],
    directContemplation: consorcioDireto.getRange("B29").values[0][0],
    cartaPurchase: cartaContemplada.getRange("B23").values[0][0],
    cartaUse: cartaContemplada.getRange("B24").values[0][0],
  };

  const paidFinIqCols = ["J", "Z", "AP"];
  const receivedFinIqCols = ["K", "AA", "AQ"];
  const paidAllCols = [...paidFinIqCols, "BF", "BV"];
  const receivedAllCols = [...receivedFinIqCols, "BG", "BW"];

  function setScenario(use) {
    premissas.getRange("B11").values = [["Imovel"]];
    premissas.getRange("B47").values = [[use]];
    premissas.getRange("B49").values = [[1]];
    premissas.getRange("B27").values = [[60]];
    consorcioDireto.getRange("B29").values = [[60]];
    cartaContemplada.getRange("B23").values = [[1]];
    cartaContemplada.getRange("B24").values = [[1]];
  }

  function restoreScenario() {
    premissas.getRange("B11").values = [[original.assetType]];
    premissas.getRange("B47").values = [[original.use]];
    premissas.getRange("B49").values = [[original.availability]];
    premissas.getRange("B27").values = [[original.iqContemplation]];
    consorcioDireto.getRange("B29").values = [[original.directContemplation]];
    cartaContemplada.getRange("B23").values = [[original.cartaPurchase]];
    cartaContemplada.getRange("B24").values = [[original.cartaUse]];
  }

  try {
    setScenario("Moradia");
    const moradiaHorizon = cellNumber(premissas, "B39");
    const moradiaFinAccess = Math.max(cellNumber(premissas, "B48"), cellNumber(premissas, "B49"));
    const moradiaDirectAccess = Math.max(cellNumber(consorcioDireto, "B29"), cellNumber(premissas, "B49"));
    const moradiaFinPaid = sumMotorCols(paidFinIqCols, 1, moradiaHorizon);
    const moradiaDirectPaidBefore = sumMotorCol("BF", 1, moradiaDirectAccess - 1);
    const moradiaDirectExpectedBefore = sumMotorCol("D", 1, moradiaDirectAccess - 1);
    const moradiaDirectPaidAfter = sumMotorCol("BF", moradiaDirectAccess, moradiaHorizon);
    const moradiaReceived = sumMotorCols(receivedAllCols, 1, moradiaHorizon);

    const moradiaResults = [
      assertNear("Moradia: estrategias 1, 2 e 3 sem aluguel de espera", moradiaFinPaid, 0),
      assertNear("Moradia: consorcio direto paga aluguel de espera", moradiaDirectPaidBefore, moradiaDirectExpectedBefore),
      assertNear("Moradia: consorcio direto encerra aluguel no acesso", moradiaDirectPaidAfter, 0),
      assertNear("Moradia: nenhuma renda recebida", moradiaReceived, 0),
    ];

    setScenario("Investimento");
    const investimentoHorizon = cellNumber(premissas, "B39");
    const investimentoFinAccess = Math.max(cellNumber(premissas, "B48"), cellNumber(premissas, "B49"));
    const investimentoDirectAccess = Math.max(cellNumber(consorcioDireto, "B29"), cellNumber(premissas, "B49"));
    const investimentoPaid = sumMotorCols(paidAllCols, 1, investimentoHorizon);
    const investimentoFinExpected = sumMotorCol("E", investimentoFinAccess, investimentoHorizon);
    const investimentoDirectExpected = sumMotorCol("E", investimentoDirectAccess, investimentoHorizon);
    const investimentoCartaAccess = Math.max(cellNumber(cartaContemplada, "B44"), cellNumber(premissas, "B49"));
    const investimentoCartaExpected = sumMotorCol("E", investimentoCartaAccess, investimentoHorizon);
    const investimentoDirectReceivedBefore = sumMotorCol("BG", 1, investimentoDirectAccess - 1);
    const investimentoCartaReceivedBefore = sumMotorCol("BW", 1, investimentoCartaAccess - 1);

    const investimentoResults = [
      assertNear("Investimento: nenhum aluguel de espera", investimentoPaid, 0),
      assertNear("Investimento: estrategia 1 recebe renda do acesso ao horizonte", sumMotorCol("K", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: estrategia 2 recebe renda do acesso ao horizonte", sumMotorCol("AA", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: estrategia 3 recebe renda do acesso ao horizonte", sumMotorCol("AQ", investimentoFinAccess, investimentoHorizon), investimentoFinExpected),
      assertNear("Investimento: consorcio direto sem renda antes do acesso", investimentoDirectReceivedBefore, 0),
      assertNear("Investimento: consorcio direto recebe renda do acesso ao horizonte", sumMotorCol("BG", investimentoDirectAccess, investimentoHorizon), investimentoDirectExpected),
      assertNear("Investimento: carta contemplada sem renda antes do acesso", investimentoCartaReceivedBefore, 0),
      assertNear("Investimento: carta contemplada recebe renda do acesso ao horizonte", sumMotorCol("BW", investimentoCartaAccess, investimentoHorizon), investimentoCartaExpected),
    ];

    premissas.getRange("B11").values = [["Veiculo"]];
    const vehiclePaid = sumMotorCols(paidAllCols, 1, investimentoHorizon);
    const vehicleReceived = sumMotorCols(receivedAllCols, 1, investimentoHorizon);
    const vehicleResults = [
      assertNear("Veiculo: nenhum aluguel pago", vehiclePaid, 0),
      assertNear("Veiculo: nenhuma renda recebida", vehicleReceived, 0),
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
        accessCartaContemplada: investimentoCartaAccess,
        results: investimentoResults,
      },
      veiculo: { results: vehicleResults },
    };
  } finally {
    restoreScenario();
  }
}

function runLowCashViabilityTests() {
  const original = {
    cash: premissas.getRange("B42").values[0][0],
    reserve: premissas.getRange("B43").values[0][0],
    use: premissas.getRange("B47").values[0][0],
  };

  try {
    premissas.getRange("B42").values = [[400]];
    premissas.getRange("B43").values = [[0]];
    premissas.getRange("B47").values = [["Investimento"]];

    const statuses = resumo.getRange("O10:O14").values.flat();
    const viabilityOutputs = resumo.getRange("I10:L14").values.flat();
    const deficits = resumo.getRange("H10:H14").values.flat().map(asNumber);

    if (!statuses.every((status) => status === "Inviavel")) {
      throw new Error(`Caixa baixo deveria tornar todas as estrategias inviaveis, mas status = ${statuses.join(", ")}`);
    }
    if (!viabilityOutputs.every((value) => value === "Inviavel")) {
      throw new Error("Caixa baixo deveria mostrar Inviavel em patrimonio, VPL e Gap.");
    }
    if (!deficits.every((value) => value > 1)) {
      throw new Error("Caixa baixo deveria gerar deficit nao coberto em todas as estrategias.");
    }

    return {
      cash: 400,
      reserve: 0,
      statuses,
      deficits,
    };
  } finally {
    premissas.getRange("B42").values = [[original.cash]];
    premissas.getRange("B43").values = [[original.reserve]];
    premissas.getRange("B47").values = [[original.use]];
  }
}

function runScenarioMatrixTests() {
  const original = {
    assetType: premissas.getRange("B11").values[0][0],
    system: premissas.getRange("B19").values[0][0],
    use: premissas.getRange("B47").values[0][0],
    horizon: premissas.getRange("B39").values[0][0],
  };
  const scenarios = [
    { assetType: "Imovel", system: "SAC", rate: 0.06 },
    { assetType: "Imovel", system: "PRICE", rate: 0.06 },
    { assetType: "Veiculo", system: "SAC", rate: -0.15 },
    { assetType: "Veiculo", system: "PRICE", rate: -0.15 },
  ];
  const horizons = [60, 120, 240, 420];
  const results = [];

  try {
    for (const scenario of scenarios) {
      for (const horizon of horizons) {
        premissas.getRange("B11").values = [[scenario.assetType]];
        premissas.getRange("B19").values = [[scenario.system]];
        premissas.getRange("B47").values = [["Investimento"]];
        premissas.getRange("B39").values = [[horizon]];

        const term = Math.max(1, Math.min(FIN_MONTHS, cellNumber(premissas, "B13")));
        const finalBalance = cellNumber(financiamento, `M${term + 5}`);
        const firstPayment = cellNumber(financiamento, "J6");
        const firstAmortization = cellNumber(financiamento, "G6");
        const minAmortization = minRange(financiamento, `G6:G${term + 5}`);
        const residualValue = cellNumber(motor, `F${horizon + 4}`);
        const selectedRate = cellNumber(premissas, "B66");
        const summaryStatuses = resumo.getRange("O10:O14").values.flat();

        assertNear(`${scenario.assetType}/${scenario.system}/${horizon}: taxa do bem selecionada`, selectedRate, scenario.rate, 0.000001);
        assertNear(`${scenario.assetType}/${scenario.system}/${horizon}: financiamento termina zerado`, finalBalance, 0, 1);
        if (firstPayment <= 0) {
          throw new Error(`${scenario.assetType}/${scenario.system}/${horizon}: prestacao inicial deve ser positiva.`);
        }
        if (firstAmortization < -1 || minAmortization < -1) {
          throw new Error(`${scenario.assetType}/${scenario.system}/${horizon}: amortizacao negativa detectada.`);
        }
        if (residualValue < -1) {
          throw new Error(`${scenario.assetType}/${scenario.system}/${horizon}: valor residual do bem ficou negativo.`);
        }
        if (summaryStatuses.some((status) => status !== "OK" && status !== "Inviavel")) {
          throw new Error(`${scenario.assetType}/${scenario.system}/${horizon}: status inesperado no Resumo.`);
        }

        results.push({
          name: `${scenario.assetType}/${scenario.system}`,
          horizon,
          finalBalance,
          firstPayment,
          firstAmortization,
          residualValue,
          summaryStatuses,
        });
      }
    }
    return results;
  } finally {
    premissas.getRange("B11").values = [[original.assetType]];
    premissas.getRange("B19").values = [[original.system]];
    premissas.getRange("B47").values = [[original.use]];
    premissas.getRange("B39").values = [[original.horizon]];
  }
}

function runCartaContempladaTests() {
  const original = {
    horizon: premissas.getRange("B39").values[0][0],
    use: premissas.getRange("B47").values[0][0],
    active: cartaContemplada.getRange("B22").values[0][0],
    purchase: cartaContemplada.getRange("B23").values[0][0],
    creditUse: cartaContemplada.getRange("B24").values[0][0],
    grossCredit: cartaContemplada.getRange("B25").values[0][0],
    remaining: cartaContemplada.getRange("B31").values[0][0],
    installment: cartaContemplada.getRange("B32").values[0][0],
  };

  try {
    premissas.getRange("B39").values = [[120]];
    premissas.getRange("B47").values = [["Investimento"]];
    cartaContemplada.getRange("B22").values = [["Sim"]];
    cartaContemplada.getRange("B23").values = [[1]];
    cartaContemplada.getRange("B24").values = [[1]];
    cartaContemplada.getRange("B31").values = [[180]];

    const initialCost = cellNumber(cartaContemplada, "E55");
    const expectedInitialCost = cellNumber(cartaContemplada, "B37");
    const parcelCount = cartaContemplada
      .getRange(`C${ccStartRow}:C${ccEndRow}`)
      .values
      .flat()
      .filter((value) => asNumber(value) > 0).length;
    const remainingParcels = cellNumber(cartaContemplada, "B31");
    const obligationAtHorizon = cellNumber(cartaContemplada, "B40");
    const summaryDebt = cellNumber(resumo, "N14");
    const economicCost = cellNumber(cartaContemplada, "B39");
    const creditCoverage = cellNumber(cartaContemplada, "B27") + cellNumber(cartaContemplada, "B35");
    const assetValueAtAcquisition = cellNumber(motor, "CF5");
    const sharedAssetValueAtAcquisition = cellNumber(motor, "F5");

    assertNear("Carta contemplada: custo inicial no mes de compra", initialCost, expectedInitialCost);
    assertNear("Carta contemplada: quantidade de parcelas restantes", parcelCount, remainingParcels, 0);
    assertNear("Carta contemplada: obrigacao remanescente no resumo", summaryDebt, obligationAtHorizon);
    assertNear("Carta contemplada: credito liquido + complemento cobre bem", creditCoverage, cellNumber(cartaContemplada, "B4"));
    assertNear("Carta contemplada: bem entra no patrimonio no mes de aquisicao", assetValueAtAcquisition, sharedAssetValueAtAcquisition);
    assertNear("Carta contemplada: credito excedente nao entra no caixa", cellNumber(cartaContemplada, "B36"), 0);

    if (obligationAtHorizon <= 0) {
      throw new Error("Carta contemplada: o cenario de teste deveria manter obrigacao remanescente positiva no horizonte.");
    }

    cartaContemplada.getRange("B25").values = [[500000]];
    if (cellNumber(cartaContemplada, "B34") <= 0) {
      throw new Error("Carta contemplada: o teste de credito excedente deveria produzir excedente positivo.");
    }
    assertNear("Carta contemplada: excedente alto continua fora do caixa", cellNumber(cartaContemplada, "B36"), 0);
    assertNear(
      "Carta contemplada: excedente alto nao reduz custo inicial",
      cellNumber(cartaContemplada, "B37"),
      cellNumber(cartaContemplada, "B28") + cellNumber(cartaContemplada, "B30"),
    );

    premissas.getRange("B39").values = [[OPP_MAX_MONTHS]];
    cartaContemplada.getRange("B23").values = [[OPP_MAX_MONTHS]];
    cartaContemplada.getRange("B24").values = [[OPP_MAX_MONTHS]];
    cartaContemplada.getRange("B25").values = [[original.grossCredit]];
    cartaContemplada.getRange("B31").values = [[OPP_MAX_MONTHS]];
    cartaContemplada.getRange("B32").values = [[1000]];

    const lateCardHorizonRow = ccStartRow + OPP_MAX_MONTHS - 1;
    const lateCardLastRow = ccEndRow;
    const lateCardParcelCount = cartaContemplada
      .getRange("C" + ccStartRow + ":C" + ccEndRow)
      .values
      .flat()
      .filter((value) => asNumber(value) > 0).length;
    const lateCardObligation = cellNumber(cartaContemplada, "B40");
    const lateCardSummaryDebt = cellNumber(resumo, "N14");
    const lateCardLastInstallment = cellNumber(cartaContemplada, "C" + lateCardLastRow);
    const monthlyOpportunityRate = cellNumber(cartaContemplada, "B9");
    const annualCardAdjustment = cellNumber(cartaContemplada, "B33");
    let expectedLateCardObligation = 0;

    for (let month = OPP_MAX_MONTHS + 1; month <= CARD_MAX_MONTHS; month += 1) {
      const adjustmentFactor = Math.pow(
        1 + annualCardAdjustment,
        Math.floor((month - OPP_MAX_MONTHS) / 12),
      );
      expectedLateCardObligation +=
        cellNumber(cartaContemplada, "B32") *
        adjustmentFactor /
        Math.pow(1 + monthlyOpportunityRate, month - OPP_MAX_MONTHS);
    }

    assertNear("Carta contemplada: cauda longa preserva 480 parcelas", lateCardParcelCount, OPP_MAX_MONTHS, 0);
    assertNear(
      "Carta contemplada: obrigacao da cauda longa no horizonte",
      lateCardObligation,
      expectedLateCardObligation,
    );
    assertNear("Carta contemplada: resumo inclui obrigacao da cauda longa", lateCardSummaryDebt, lateCardObligation);
    if (lateCardObligation <= 0 || lateCardLastInstallment <= 0) {
      throw new Error("Carta contemplada: parcelas posteriores ao horizonte foram truncadas.");
    }
    if (checks.getRange("F39").values[0][0] !== "OK" || checks.getRange("F41").values[0][0] !== "OK") {
      throw new Error("Carta contemplada: checks da cauda longa deveriam permanecer OK.");
    }

    return {
      initialCost,
      parcelCount,
      obligationAtHorizon,
      summaryDebt,
      economicCost,
      lateCardObligation,
      status: resumo.getRange("O14").values[0][0],
    };
  } finally {
    premissas.getRange("B39").values = [[original.horizon]];
    premissas.getRange("B47").values = [[original.use]];
    cartaContemplada.getRange("B22").values = [[original.active]];
    cartaContemplada.getRange("B23").values = [[original.purchase]];
    cartaContemplada.getRange("B24").values = [[original.creditUse]];
    cartaContemplada.getRange("B25").values = [[original.grossCredit]];
    cartaContemplada.getRange("B31").values = [[original.remaining]];
    cartaContemplada.getRange("B32").values = [[original.installment]];
  }
}

function runConsortiumPositionTests() {
  const original = {
    horizon: premissas.getRange("B39").values[0][0],
    use: premissas.getRange("B47").values[0][0],
  };
  const results = [];

  try {
    premissas.getRange("B47").values = [["Investimento"]];
    const iqContemplation = cellNumber(premissas, "B27");
    const directContemplation = cellNumber(consorcioDireto, "B29");
    const groupEnd = cellNumber(premissas, "B25");
    const iqPre = Math.max(1, iqContemplation - 1);
    const iqPost = Math.min(groupEnd - 1, iqContemplation + 1);
    const directPre = Math.max(1, directContemplation - 1);
    const directPost = Math.min(groupEnd - 1, directContemplation + 1);

    const iqRightPre = cellNumber(consorcio, `O${iqPre + 5}`);
    const iqObligationPre = cellNumber(consorcio, `N${iqPre + 5}`);
    const iqDebtAt = cellNumber(consorcio, `P${iqContemplation + 5}`);
    const iqObligationAt = cellNumber(consorcio, `N${iqContemplation + 5}`);
    const iqDebtPost = cellNumber(consorcio, `P${iqPost + 5}`);
    const iqObligationPost = cellNumber(consorcio, `N${iqPost + 5}`);
    const directRightPre = cellNumber(consorcioDireto, `L${directPre + 54}`);
    const directObligationPre = cellNumber(consorcioDireto, `K${directPre + 54}`);
    const directDebtAt = cellNumber(consorcioDireto, `K${directContemplation + 54}`);
    const directDebtPost = cellNumber(consorcioDireto, `K${directPost + 54}`);

    if (iqRightPre <= 0 || iqObligationPre <= 0) {
      throw new Error("IQ deve reconhecer direito economico e obrigacao antes da contemplacao.");
    }
    if (directRightPre <= 0 || directObligationPre <= 0) {
      throw new Error("Consorcio direto deve reconhecer direito economico e obrigacao antes da contemplacao.");
    }
    if (iqDebtAt <= 0 || directDebtAt <= 0 || directDebtPost <= 0) {
      throw new Error("As obrigacoes dos consorcios devem permanecer positivas apos a contemplacao e antes do fim do grupo.");
    }
    results.push(assertNear("IQ apos contemplacao: divida bate VP das parcelas", iqDebtPost, iqObligationPost));
    results.push(assertNear("IQ na contemplacao: divida inclui obrigacao", iqDebtAt, iqObligationAt));
    results.push(assertNear("IQ fim do grupo: obrigacao zerada", cellNumber(consorcio, `N${groupEnd + 5}`), 0));
    results.push(assertNear("Direto fim do grupo: obrigacao zerada", cellNumber(consorcioDireto, `K${groupEnd + 54}`), 0));

    const horizonResults = [];
    for (const horizon of [60, 120, 240, 420]) {
      premissas.getRange("B39").values = [[horizon]];
      const iqSummaryDebt = cellNumber(resumo, "N12");
      const directSummaryDebt = cellNumber(resumo, "N13");
      const iqScheduleDebt = cellNumber(consorcio, `P${horizon + 5}`);
      const directScheduleDebt = cellNumber(consorcioDireto, `K${horizon + 54}`);
      assertNear(`Horizonte ${horizon}: divida IQ bate cronograma`, iqSummaryDebt, iqScheduleDebt);
      assertNear(`Horizonte ${horizon}: divida direto bate cronograma`, directSummaryDebt, directScheduleDebt);
      horizonResults.push({ horizon, iqDebt: iqSummaryDebt, directDebt: directSummaryDebt });
    }

    return {
      iqPre: { month: iqPre, right: iqRightPre, obligation: iqObligationPre },
      iqAt: { month: iqContemplation, debt: iqDebtAt },
      directPre: { month: directPre, right: directRightPre, obligation: directObligationPre },
      directAt: { month: directContemplation, debt: directDebtAt },
      horizonResults,
      results,
    };
  } finally {
    premissas.getRange("B39").values = [[original.horizon]];
    premissas.getRange("B47").values = [[original.use]];
  }
}

function runFuturePriceTests() {
  const original = {
    assetType: premissas.getRange("B11").values[0][0],
    directMonth: consorcioDireto.getRange("B29").values[0][0],
    cartaPurchase: cartaContemplada.getRange("B23").values[0][0],
    cartaUse: cartaContemplada.getRange("B24").values[0][0],
  };

  try {
    premissas.getRange("B11").values = [["Imovel"]];
    consorcioDireto.getRange("B29").values = [[60]];
    cartaContemplada.getRange("B23").values = [[60]];
    cartaContemplada.getRange("B24").values = [[60]];
    const initialPrice = cellNumber(premissas, "B6");
    const rate = cellNumber(premissas, "B66");
    const prices = {
      financing: cellNumber(premissas, "B79"),
      direct: cellNumber(consorcioDireto, "B4"),
      carta: cellNumber(cartaContemplada, "B4"),
    };
    const expectedFuture = initialPrice * Math.pow(1 + rate, (60 - 1) / 12);
    assertNear("Preco atual: financiamento", prices.financing, initialPrice);
    assertNear("Preco futuro: consorcio direto", prices.direct, expectedFuture);
    assertNear("Preco futuro: carta contemplada", prices.carta, expectedFuture);
    if (prices.direct <= initialPrice || prices.carta <= initialPrice) {
      throw new Error("Consorcios com contemplacao futura devem usar preco futuro do Imovel.");
    }

    return { ...prices, expectedFuture };
  } finally {
    premissas.getRange("B11").values = [[original.assetType]];
    consorcioDireto.getRange("B29").values = [[original.directMonth]];
    cartaContemplada.getRange("B23").values = [[original.cartaPurchase]];
    cartaContemplada.getRange("B24").values = [[original.cartaUse]];
  }
}

function runBidLogicTests() {
  const original = {
    embeddedOn: consorcioDireto.getRange("B31").values[0][0],
    embeddedPct: consorcioDireto.getRange("B32").values[0][0],
    freeOn: consorcioDireto.getRange("B33").values[0][0],
    freePct: consorcioDireto.getRange("B34").values[0][0],
    fixedOn: consorcioDireto.getRange("B35").values[0][0],
    fixedPct: consorcioDireto.getRange("B36").values[0][0],
  };

  try {
    consorcioDireto.getRange("B31").values = [["Sim"]];
    consorcioDireto.getRange("B32").values = [[0.25]];
    consorcioDireto.getRange("B33").values = [["Nao"]];
    consorcioDireto.getRange("B34").values = [[0]];
    consorcioDireto.getRange("B35").values = [["Sim"]];
    consorcioDireto.getRange("B36").values = [[0.25]];
    assertNear("Lance fixo 25% + embutido 25%: caixa cliente", cellNumber(consorcioDireto, "B44"), 0);
    assertNear("Lance fixo 25% + embutido 25%: credito liquido", cellNumber(consorcioDireto, "B43"), cellNumber(consorcioDireto, "B4"));

    consorcioDireto.getRange("B32").values = [[0.1]];
    consorcioDireto.getRange("B33").values = [["Sim"]];
    consorcioDireto.getRange("B34").values = [[0.35]];
    consorcioDireto.getRange("B35").values = [["Nao"]];
    assertNear("Lance livre 35% + embutido 10%: descapitalizacao", cellNumber(consorcioDireto, "B39"), 0.25, 0.000001);
    assertNear("Lance livre: caixa cliente em R$", cellNumber(consorcioDireto, "B44"), cellNumber(consorcioDireto, "B42") * 0.25);

    consorcioDireto.getRange("B32").values = [[0.4]];
    consorcioDireto.getRange("B34").values = [[0.2]];
    assertNear("Embutido maior que total e corrigido", cellNumber(consorcioDireto, "B37"), 0.2, 0.000001);
    assertNear("Embutido corrigido elimina caixa externo", cellNumber(consorcioDireto, "B44"), 0);

    consorcioDireto.getRange("B32").values = [[0.1]];
    consorcioDireto.getRange("B33").values = [["Nao"]];
    consorcioDireto.getRange("B35").values = [["Sim"]];
    consorcioDireto.getRange("B36").values = [[0.25]];
    const output25 = cellNumber(resumo, "E13");
    consorcioDireto.getRange("B36").values = [[0.35]];
    const output35 = cellNumber(resumo, "E13");
    if (Math.abs(output25 - output35) <= 1) {
      throw new Error("Alterar o percentual de lance deve alterar as saidas da estrategia de consorcio direto.");
    }

    return {
      fixedEmbeddedClientCash: 0,
      freeEmbeddedClientPct: 0.25,
      clampedEmbeddedPct: 0.2,
      output25,
      output35,
    };
  } finally {
    consorcioDireto.getRange("B31").values = [[original.embeddedOn]];
    consorcioDireto.getRange("B32").values = [[original.embeddedPct]];
    consorcioDireto.getRange("B33").values = [[original.freeOn]];
    consorcioDireto.getRange("B34").values = [[original.freePct]];
    consorcioDireto.getRange("B35").values = [[original.fixedOn]];
    consorcioDireto.getRange("B36").values = [[original.fixedPct]];
  }
}

const scenarioMatrixTests = runScenarioMatrixTests();
const rentalComparabilityTests = runRentalComparabilityTests();
const amortizationCashFreeTests = runAmortizationCashFreeTests();
const cartaContempladaTests = runCartaContempladaTests();
const consortiumPositionTests = runConsortiumPositionTests();
const futurePriceTests = runFuturePriceTests();
const bidLogicTests = runBidLogicTests();
const lowCashViabilityTests = runLowCashViabilityTests();

const finalCheckStatus = checks.getRange("F30").values[0][0];
if (finalCheckStatus !== "OK") {
  throw new Error(`Checks!F30 deveria estar OK no preset final, mas esta como ${finalCheckStatus}.`);
}

const reportDate = new Date().toISOString().slice(0, 10);
const scenarioValidationLines = scenarioMatrixTests
  .map((scenario) => `- **${scenario.name}, horizonte ${scenario.horizon} meses**: saldo final R$ ${scenario.finalBalance.toFixed(2)}; prestacao inicial R$ ${scenario.firstPayment.toFixed(2)}; amortizacao inicial R$ ${scenario.firstAmortization.toFixed(2)}; valor residual R$ ${scenario.residualValue.toFixed(2)}.`)
  .join("\n");
const consortiumHorizonValidationLines = consortiumPositionTests.horizonResults
  .map((item) => `- **Horizonte ${item.horizon} meses**: divida IQ R$ ${item.iqDebt.toFixed(2)}; obrigacao do consorcio direto R$ ${item.directDebt.toFixed(2)}.`)
  .join("\n");

await fs.writeFile(
  correctionReportPath,
  `# Relatorio final de implementacao - Calculadora de Aquisicao

Arquivo gerado: \`${path.basename(outputPath)}\`
Data: ${reportDate}

## Escopo implementado

- \`Tipo do bem\` seleciona valorizacao do Imovel ou depreciacao do Veiculo; Veiculo nao usa finalidade, aluguel pago ou renda.
- \`Sistema de financiamento\` seleciona SAC ou PRICE, com indexador contratual generico.
- Financiamento compra no mes 1 pelo valor atual; consorcio direto e carta contemplada usam o preco futuro no mes de aquisicao.
- Generalizado o cronograma de \`Financiamento\` para manter as colunas de amortizacao, juros, seguro, prestacao e saldo final nos dois sistemas.
- Mantida a amortizacao extraordinaria como reducao de prazo, usando a sobra entre caixa livre mensal e prestacao ordinaria do sistema selecionado.
- A aba \`Carta Contemplada\` desconsidera conservadoramente credito excedente e nunca o transforma em caixa investivel.
- Adicionada a quinta estrategia em \`Motor Caixa!BT:CI\`.
- IQ e consorcio direto agora reconhecem direito economico conservador e VP das parcelas futuras antes da contemplacao.
- Depois do uso do credito, a obrigacao futura permanece no patrimonio ate o fim do grupo; a divida do IQ nao e zerada automaticamente.
- Para Imovel, Moradia usa aluguel pago antes do acesso e Investimento usa renda depois do acesso. Custo de posse foi removido.
- Expandido o \`Resumo\`, \`Checks\`, \`Fontes\` e graficos para as cinco estrategias.
- Preservado o modelo nominal: inflacao permanece apenas como conversao de apresentacao.

## Validacoes executadas

${scenarioValidationLines}

${consortiumHorizonValidationLines}

- **IQ antes da contemplacao (mes ${consortiumPositionTests.iqPre.month})**: direito R$ ${consortiumPositionTests.iqPre.right.toFixed(2)}; obrigacao R$ ${consortiumPositionTests.iqPre.obligation.toFixed(2)}.
- **IQ na contemplacao (mes ${consortiumPositionTests.iqAt.month})**: divida total R$ ${consortiumPositionTests.iqAt.debt.toFixed(2)}, incluindo parcelas futuras.
- **Consorcio direto antes da contemplacao (mes ${consortiumPositionTests.directPre.month})**: direito R$ ${consortiumPositionTests.directPre.right.toFixed(2)}; obrigacao R$ ${consortiumPositionTests.directPre.obligation.toFixed(2)}.
- **Consorcio direto na contemplacao (mes ${consortiumPositionTests.directAt.month})**: obrigacao R$ ${consortiumPositionTests.directAt.debt.toFixed(2)}.
- **Preco na aquisicao**: financiamento R$ ${futurePriceTests.financing.toFixed(2)}; consorcio direto R$ ${futurePriceTests.direct.toFixed(2)}; carta R$ ${futurePriceTests.carta.toFixed(2)}.
- **Lances**: preset fixo + embutido 25% com caixa externo R$ ${bidLogicTests.fixedEmbeddedClientCash.toFixed(2)}; livre + embutido validado; embutido excedente corrigido para ${(bidLogicTests.clampedEmbeddedPct * 100).toFixed(0)}%.
- **Carta Contemplada**: custo inicial R$ ${cartaContempladaTests.initialCost.toFixed(2)} no mes de compra; ${cartaContempladaTests.parcelCount} parcelas detectadas; obrigacao remanescente no horizonte de teste R$ ${cartaContempladaTests.obligationAtHorizon.toFixed(2)}; custo economico R$ ${cartaContempladaTests.economicCost.toFixed(2)}; no cenario com compra no mes 480 e 480 parcelas, obrigacao futura R$ ${cartaContempladaTests.lateCardObligation.toFixed(2)}.
- **Credito excedente**: testada carta de R$ 500.000; o excedente permaneceu fora do caixa e nao reduziu o custo inicial.
- **Caixa insuficiente**: com caixa livre mensal de R$ ${lowCashViabilityTests.cash.toFixed(2)} e reserva R$ ${lowCashViabilityTests.reserve.toFixed(2)}, as cinco estrategias ficaram \`Inviavel\`.
- **Amortizacao extraordinaria**: prestacao ordinaria inicial R$ ${amortizationCashFreeTests.firstOrdinaryPayment.toFixed(2)}, amortizacao extra inicial R$ ${amortizationCashFreeTests.firstExtraPayment.toFixed(2)} e quitacao no mes ${amortizationCashFreeTests.payoffMonth}.
- **Espera/renda comparavel**: financiamento, amortizacao, IQ, consorcio direto e carta contemplada respeitam o mes de acesso ao bem.
- Executadas varreduras para \`#REF!\`, \`#DIV/0!\`, \`#VALUE!\`, \`#NAME?\`, \`#N/A\` e \`#NUM!\`.
- Confirmada ausencia de TIR, IRR, XIRR, MIRR, CET e custo efetivo.
- \`Checks!F30\` terminou em \`${finalCheckStatus}\`.
- Confirmado que Motor Caixa, financiamento, consorcios, carta e VPL nominal nao usam inflacao como driver de calculo.

## Convencoes financeiras

- SAC: saldo corrigido dividido pelos meses restantes para a amortizacao ordinaria.
- PRICE: prestacao financeira recalculada sobre saldo corrigido e prazo remanescente; juros sobre saldo corrigido; amortizacao igual a prestacao financeira menos juros.
- Obrigacao de consorcio: VP das parcelas posteriores ao mes avaliado, descontadas pela taxa mensal de oportunidade.
- Direito economico antes da contemplacao: credito esperado na contemplacao descontado ao mes avaliado; o direito zera quando o credito e utilizado.
- Patrimonio dos consorcios: saldo investido + bem ou direito economico - financiamento residual - VP das parcelas futuras.
- Carta contemplada: patrimonio = saldo investido + valor do bem - obrigacao remanescente, inclusive para parcelas posteriores ao horizonte de comparacao.
- Credito excedente da carta e registrado para auditoria, mas nao reduz desembolso nem vira caixa ou patrimonio financeiro.
- Lance fixo e livre sao exclusivos; lance embutido exige lance total e e limitado a esse total. Caixa do cliente = credito bruto x (lance total - embutido).
- Estrategias com deficit nao coberto mostram \`Inviavel\` e ficam fora do grafico patrimonial.

## Arquivos auxiliares atualizados

- \`manual_uso_planilha.txt\`
- \`manual_calculos_planilha.txt\`
- \`manual_uso_interface_web.txt\`

## Limitacoes mantidas

- Contemplacao e transferencia de carta sao premissas de simulacao, nao garantias.
- O direito economico pre-contemplacao nao e valor de resgate nem cotacao de mercado da cota; e uma aproximacao conservadora baseada no credito esperado e no desconto de oportunidade.
- Custos tributarios, cartorarios e operacionais nao modelados devem ser revisados conforme o tipo do bem e a proposta real.
- O modelo nao calcula TIR nem custo efetivo; mantem VPL vs base, Gap VF, patrimonio corrigido e saidas nominais.
`,
  "utf8",
);

await fs.writeFile(
  manualUsoPlanilhaPath,
  `MANUAL DE USO DA PLANILHA
Calculadora de Aquisicao

1. ESTRATEGIAS COMPARADAS

1) Financiamento sem amortizacao extraordinaria.
2) Financiamento com amortizacao extraordinaria.
3) Financiamento + consorcio como interveniente quitante.
4) Consorcio direto.
5) Compra de carta contemplada.

2. CAMPOS NOVOS

- Tipo do bem: escolha Imovel ou Veiculo em Premissas!B11.
- Sistema de financiamento: escolha SAC ou PRICE em Premissas!B19.
- Regra do preco: financiamento usa o valor atual no mes 1; consorcios usam o preco futuro no mes de aquisicao.
- Imovel usa Premissas!B91; Veiculo usa Premissas!B92.
- O indexador em Premissas!B16 e generico e deve refletir o contrato real.
- Finalidade do Imovel: Moradia ou Investimento. Para Veiculo, aluguel e renda sao ignorados.

3. CARTA CONTEMPLADA

Preencha na aba Carta Contemplada:
- mes de compra/transferencia;
- mes de uso do credito;
- credito bruto e restricoes;
- preco pago pela carta;
- taxa de transferencia;
- parcelas restantes, parcela atual e reajuste;
- credito excedente e complemento necessario.

O modelo calcula complemento necessario, custo inicial total, VP das parcelas, custo economico e obrigacao remanescente. Parcelas que terminam depois do horizonte continuam compondo a obrigacao no horizonte.

4. POSICAO DOS CONSORCIOS

- Antes da contemplacao, IQ e consorcio direto mostram o direito economico descontado e o VP das parcelas futuras.
- Na contemplacao, o direito e substituido pelo bem ou usado na quitacao.
- As parcelas futuras continuam como obrigacao ate o fim do grupo.
- A coluna Saldo devedor do Resumo inclui financiamento residual e obrigacoes de consorcio.

5. ORDEM DE USO

1) Preencha Premissas.
2) Revise Consorcio Direto e Carta Contemplada.
3) Confira Checks!F30.
4) Compare as cinco linhas do Resumo.
5) Descarte estrategias Inviaveis antes de comparar patrimonio, VPL e gap.

6. LEITURA DO RESUMO

- Status: primeiro filtro da decisao.
- Patrimonio em R$ de hoje: metrica principal entre estrategias viaveis.
- VPL vs base e Gap VF vs base: comparacao com investir reserva e caixa livre.
- Saldo devedor: inclui financiamento e obrigacoes remanescentes de IQ, consorcio direto e carta contemplada.
- Saidas nominais: diagnostico de desembolso, nao ranking economico isolado.

7. ALERTAS

- PRICE nao tem a mesma trajetoria decrescente de parcela do SAC.
- Carta contemplada depende de transferencia, aprovacao cadastral, garantias e regras da administradora.
- Credito excedente nunca vira caixa, investimento ou patrimonio financeiro.
- Para veiculo, valor residual nunca fica abaixo de zero.
- Aluguel pago e renda sao opcionais apenas para Imovel; use zero quando nao se aplicarem.
- Nao interprete o direito economico da cota como valor garantido de resgate.
- O modelo nao usa TIR nem custo efetivo.
- Nao altere formulas do Motor Caixa, Financiamento, Amortizacao ou Resumo.
`,
  "utf8",
);

await fs.writeFile(
  manualCalculosPath,
  `MANUAL DE CALCULOS
Calculadora de Aquisicao

1. TAXAS MENSAIS

Taxa mensal = (1 + taxa anual)^(1/12) - 1.

2. FINANCIAMENTO SAC

Saldo corrigido = saldo inicial + indexador mensal.
Amortizacao = saldo corrigido / meses restantes.
Juros = saldo corrigido * juros mensais.
Prestacao total = amortizacao + juros + seguro.

3. FINANCIAMENTO PRICE

Se juros = 0:
Prestacao financeira = saldo corrigido / meses restantes.

Se juros > 0:
Prestacao financeira = saldo corrigido * juros / (1 - (1 + juros)^(-meses restantes)).

Juros = saldo corrigido * juros mensais.
Amortizacao = prestacao financeira - juros.
Prestacao total = prestacao financeira + seguro.

4. AMORTIZACAO EXTRAORDINARIA

Sobra = MAX(0, caixa livre mensal - prestacao ordinaria).
Extra = MIN(saldo apos amortizacao ordinaria, sobra).
Saldo final = MAX(0, saldo corrigido - amortizacao ordinaria - extra).

5. CONSORCIOS IQ E DIRETO

Obrigacao no mes = VP das parcelas posteriores ao mes, pela taxa mensal de oportunidade.
Direito antes da contemplacao = credito esperado / (1 + taxa oportunidade mensal)^(mes contemplacao - mes atual).

Antes da contemplacao:
Patrimonio = saldo investido + valor do bem ja adquirido + direito economico - saldo financiamento - obrigacao.

Depois da contemplacao:
Patrimonio = saldo investido + valor do bem - saldo residual do financiamento - obrigacao.

O IQ nao zera a divida na contemplacao: o saldo residual pode zerar, mas o VP das parcelas futuras permanece.

6. CARTA CONTEMPLADA

Credito liquido = MAX(0, credito bruto - restricoes).
Complemento = MAX(0, valor do bem - credito liquido).
Credito excedente = MAX(0, credito liquido - valor do bem), apenas para auditoria.
Credito excedente no caixa = 0.
Custo inicial = preco da carta + transferencia + complemento.
Parcela mensal = parcela atual reajustada anualmente durante as parcelas restantes.
Custo economico = preco + transferencia + VP das parcelas - credito liquido.
Se as parcelas terminarem depois do horizonte de comparacao, o VP e a obrigacao no horizonte continuam incluindo todas as parcelas posteriores.

7. PATRIMONIO

Patrimonio = saldo investido + valor do bem - saldo devedor/obrigacao.
Para carta contemplada, a divida e o valor presente das parcelas futuras remanescentes.
Se houver deficit nao coberto, o resultado decisorio e Inviavel.

8. TIPO DO BEM

Taxa aplicada = valorizacao do Imovel ou depreciacao do Veiculo, conforme o seletor.
Valor do bem no mes = MAX(0, valor inicial * (1 + taxa aplicada)^((mes-1)/12)).
Para Veiculo, aluguel pago e renda recebida permanecem zerados.

9. PRECO NA AQUISICAO

Financiamento = valor inicial do bem no mes 1.
Consorcios = valor inicial * (1 + taxa aplicada)^((mes aquisicao-1)/12).
A regra futura alimenta consorcio direto e complemento da carta contemplada.

10. MODELO NOMINAL

Inflacao nao altera os motores. Patrimonio em R$ de hoje = patrimonio nominal / fator de inflacao acumulada.
VPL vs base = Gap VF vs base / (1 + taxa oportunidade mensal)^horizonte.
O modelo nao calcula TIR nem custo efetivo.
`,
  "utf8",
);

await fs.writeFile(
  manualWebPath,
  `MANUAL DE USO DA INTERFACE WEB
Calculadora de Aquisicao

Esta interface replica a metodologia da planilha final e compara cinco estrategias:
1) Financiamento sem amortizacao.
2) Financiamento com amortizacao extraordinaria.
3) Financiamento + consorcio como interveniente quitante.
4) Consorcio direto.
5) Compra de carta contemplada.

1. CONTROLES PRINCIPAIS

- Tipo do bem: Imovel ou Veiculo. A escolha altera valorizacao/depreciacao e os campos aplicaveis.
- Finalidade do Imovel: Moradia ou Investimento. Para Veiculo, a finalidade fica oculta.
- Preco na aquisicao: financiamento usa o valor atual no mes 1; consorcios usam o preco futuro no mes de compra.
- Sistema de financiamento: SAC ou PRICE.
- Indexador/correcao: indice contratual generico do saldo, sem presumir TR.

2. USO, RENDA E CUSTOS

- Aluguel pago e opcional e entra apenas em Moradia de Imovel antes do acesso.
- Receita bruta e redutores entram apenas em Investimento de Imovel depois da aquisicao/contemplacao e disponibilidade.
- Para Veiculo, os campos de finalidade, aluguel pago e renda ficam ocultos e os fluxos permanecem zerados.
- Para Veiculo, a depreciacao pode ser negativa e o valor residual tem piso zero.

3. CONSORCIOS

- Antes da contemplacao, IQ e consorcio direto reconhecem direito economico conservador e VP das parcelas futuras.
- Depois do uso do credito, o direito e substituido pelo bem e a obrigacao das parcelas futuras permanece.
- A divida do IQ nao zera automaticamente na contemplacao; apenas o financiamento residual pode ser quitado.

4. CARTA CONTEMPLADA

Preencha credito bruto, restricoes, preco pago pela carta, taxa de transferencia, mes de compra, mes de uso, parcelas restantes, parcela atual e reajuste.
O patrimonio desconta o VP das parcelas futuras, mesmo se a ultima parcela ficar depois do horizonte de comparacao. O custo inicial inclui preco da carta, transferencia e complemento; credito excedente nunca vira caixa.

Fluxo recomendado:
1) Escolha tipo e finalidade do bem.
2) Informe valor, reserva, caixa livre e horizonte.
3) Escolha SAC ou PRICE e revise taxas, prazo, seguro e indexador.
4) Revise consorcio IQ, consorcio direto, lances e carta contemplada.
5) Verifique viabilidade de caixa.
6) Compare patrimonio, custo e status somente entre estrategias viaveis.

5. LEITURA DOS RESULTADOS

- Cards: mostram patrimonio corrigido, parcela inicial, entrada/lance, custo e status de caixa.
- Tabela comparativa: detalha entrada/lance, parcela inicial, aquisicao, custo, patrimonio e status.
- Graficos: usam as cinco estrategias e excluem resultados inviaveis quando a comparacao patrimonial ficaria enganosa.
- Nao ha TIR, IRR, CET ou custo efetivo na interface.

6. INTEGRACAO R2

Quando a calculadora e aberta a partir da R2, o botao Aplicar na R2 aparece no topo. Ele devolve a simulacao revisada, a melhor estrategia viavel e as premissas usadas. A integracao foi preservada nesta versao.

Atencao: a liberacao de carta contemplada depende da administradora, cadastro, garantias, contrato e regras de uso do credito.
`,
  "utf8",
);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`OUTPUT_XLSX ${outputPath}`);

if (process.env.SKIP_RENDER !== "1") {
  const previewRangeBySheet = {
    Resumo: "A1:Z40",
    Premissas: "A1:E92",
    "Carta Contemplada": "A1:K90",
    "Consorcio Direto": "A1:M100",
    Checks: "A1:G60",
    Fontes: "A1:E67",
    "Motor Caixa": "A1:CI12",
    Financiamento: "A1:M60",
    "Consorcio IQ": "A1:P60",
    Amortizacao: "A1:Q60",
  };

  for (const sheetName of finalSheetNames) {
    const range = previewRangeBySheet[sheetName];
    const preview = await workbook.render({
      sheetName,
      ...(range ? { range } : {}),
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
