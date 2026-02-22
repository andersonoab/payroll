# Validação de Folha por Verba – 6 Sigma  
Igarapé Digital | Anderson Marinho  

Aplicação web local (HTML + CSS + JavaScript) para validação estatística de verbas de folha de pagamento utilizando metodologia 6 Sigma (média ± 3σ), com leitura dinâmica de Excel, agregação por grupos, filtros avançados e comparador mensal completo.

---

# 1. Visão Geral

Este projeto foi desenvolvido para:

- Identificar variações anormais em verbas de folha
- Aplicar controle estatístico sobre métricas mensais
- Detectar outliers com base em Z-Score
- Permitir análise gerencial sem necessidade de Power BI
- Funcionar 100% local (sem backend ou servidor)

A aplicação é totalmente client-side e utiliza:
- HTML (estrutura)
- CSS (layout corporativo)
- JavaScript (lógica de leitura, agregação e estatística)
- XLSX.js para leitura de Excel

---

# 2. Arquitetura

## Arquivos principais

- index.html → Interface principal
- style.css → Estilo visual corporativo
- app.js → Motor estatístico e regras de negócio

Não há dependência de banco de dados.
Os dados são armazenados em localStorage.

---

# 3. Estrutura Esperada do Excel

O sistema detecta automaticamente:

- Código da verba
- Descrição
- Colunas mensais no padrão:

MMM/AA - Métrica

Exemplos válidos:

JUN/25 - Valor  
JUL/25 - Valor  
AGO/25 - Valor  
JAN/26 - Hora  
FEV/26 - Valor  

O nome do cabeçalho é o que define a leitura.
A posição da coluna não importa.

---

# 4. Funcionamento Estatístico

Para cada grupo agregado:

1. Define o mês de referência
2. Calcula o histórico anterior
3. Aplica:

μ = média histórica  
σ = desvio padrão amostral  
Z = (Valor_ref - μ) / σ  

Limites:

LCL = μ - 3σ  
UCL = μ + 3σ  

Classificação:

|Z| ≤ 2 → Aceitável  
2 < |Z| ≤ 3 → Alerta  
|Z| > 3 → Fora  
σ = 0 → Sem histórico  

---

# 5. Recursos do Sistema

## Comparador Mensal Completo

- Grid com todos os meses
- Destaque visual por status
- Faixa visual 6σ (±1σ, ±2σ, ±3σ)
- Indicador gráfico do mês de referência

## Totais por Mês

Exibe:

- Total geral
- Total apenas status “Fora”
- Total apenas status “Alerta”
- Classificação estatística dos totais

## Filtros Disponíveis

- Busca livre
- Verba (Código - Descrição)
- Métrica mensal (Valor, Hora etc.)
- Mês de referência
- Janela histórica
- Ignorar zeros
- Status
- Z mínimo
- Z máximo
- Filtros extras dinâmicos (ex: C.R., Clas.)

## Agrupamento Dinâmico

Pode agrupar por qualquer coluna disponível:

Exemplos:

- Empresa + CPF
- Centro de Custo
- Estabelecimento
- Função
- Matrícula
- Processo

O agrupamento define o nível do comparador.

---

# 6. Colunas Extras (Ex: Função)

O sistema é dinâmico.

Se você adicionar no Excel:

... | JAN/26 - Valor | Função

O sistema irá:

- Detectar automaticamente
- Permitir exibição
- Permitir filtro
- Permitir agrupamento
- Exportar junto

Colunas textuais não interferem nos cálculos estatísticos.

---

# 7. Armazenamento

Utiliza localStorage para guardar:

- Dados importados
- Configuração de colunas visíveis
- Agrupamento selecionado
- Métrica selecionada

Botão "Limpar storage" remove todos os dados locais.

---

# 8. Aplicações Estratégicas em RH

## Auditoria de Verbas

Detectar:
- Picos anormais
- Quedas atípicas
- Volatilidade excessiva

## Controle de Rescisões

Evitar distorções estatísticas:
- Usar salário base
- Não usar total bruto rescisório para banda salarial

## Governança

Aplicar controle estatístico formal sobre folha de pagamento.

## Gestão Orçamentária

Analisar variações por:

- Centro de custo
- Estabelecimento
- Empresa
- Função

---

# 9. Filosofia do Projeto

- Simples
- Estatístico
- Local
- Estrutura aberta
- Modelo compatível com ambientes corporativos
- Pensado para auditoria e melhoria contínua

---

# 10. Evoluções Futuras

- Modo executivo resumido
- Heatmap anual
- Ranking automático de risco
- Score de volatilidade por verba
- Exportação em PPT padrão corporativo
- Integração futura com BI

---

# 11. Como Executar

1. Abrir o arquivo index.html
2. Importar Excel
3. Ajustar filtros
4. Analisar resultados
5. Exportar se necessário

Não requer instalação.

---

# 12. Autor

Projeto desenvolvido dentro da linha Igarapé Digital.  
Foco em automação, estatística aplicada à folha e inteligência operacional em RH.

---

Validação de Folha por Verba – 6 Sigma  
Igarapé Digital
