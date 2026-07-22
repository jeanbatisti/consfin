# Calculadora de Aquisicao

Ferramenta para comparar estrategias de compra de imovel ou veiculo usando financiamento, amortizacao extraordinaria, consorcio como interveniente quitante, consorcio direto e carta contemplada.

## Como usar

- Planilha: abra `Calculadora de Aquisição.xlsx`.
- Interface web: abra `Calculadora de Aquisição Web.html`.
- Documentacao: consulte a pasta `Manuais`.
- Pacote final: use `dist/Calculadora de Aquisição.zip` para distribuir a calculadora completa.

A pagina inicial redireciona para a interface em `web_app/`. A interface tambem preserva o botao `Aplicar na R2` quando aberta a partir do fluxo de R2.

## Principais recursos

- Tipo do bem: `Imovel` ou `Veiculo`, com valorizacao/depreciacao separadas. Veiculo nao usa finalidade, aluguel ou renda.
- Sistema de financiamento: `SAC` ou `PRICE`.
- Financiamento pelo valor atual no mes 1; consorcios pelo preco futuro no mes de contemplacao ou uso do credito.
- Carta contemplada com preco da carta, transferencia, credito liquido, complemento e parcelas futuras — inclusive quando ultrapassam o horizonte de comparacao —, com credito excedente desconsiderado.
- Consorcios IQ e direto com direito economico antes da contemplacao e VP das parcelas futuras depois do uso do credito.
- Lances fixo/livre exclusivos, com embutido limitado ao lance total e descapitalizacao liquida do cliente.
- Imovel em `Moradia` usa aluguel pago antes do acesso; em `Investimento`, renda depois do acesso.
- A web compara status de caixa, patrimonio corrigido e saidas nominais sem exibir VPL ou obrigacao no horizonte.

## Estrutura

- `Calculadora de Aquisição.xlsx`: planilha final.
- `web_app/`: arquivos internos da Calculadora de Aquisicao Web.
- `Manuais/`: documentacao de uso e metodologia.
- `outputs/consorcio_financiamento/`: planilha gerada, previews e relatorio final de implementacao.
- `build_spreadsheet/`: gerador auditavel da planilha.
- `dist/Calculadora de Aquisição/`: pasta pronta para distribuicao.

## Observacao

Os resultados sao simulacoes indicativas. A analise nao substitui proposta bancaria, regulamento de consorcio, contrato de transferencia de carta contemplada, analise de credito, validacao tributaria ou parecer juridico. O modelo nao calcula TIR, CET ou custo efetivo. A planilha mantem VPL, Gap VF, patrimonio corrigido e saidas nominais; a web apresenta apenas as metricas executivas aplicaveis.
