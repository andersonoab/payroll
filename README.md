Validação de Folha por Verba – 6 Sigma Igarapé Digital | Anderson
Marinho

OBJETIVO Aplicação web local (HTML + CSS + JS) para validação
estatística de verbas de folha de pagamento utilizando metodologia 6
Sigma (média ± 3σ), com leitura dinâmica de Excel, agregação por grupos,
filtros avançados e comparador mensal completo.

FUNCIONAMENTO GERAL

1.  Importação

-   Importa arquivo Excel (.xlsx)
-   Detecta automaticamente:
    -   Código da verba
    -   Descrição
    -   Colunas mensais no formato: MMM/AA - Métrica
-   Armazena dados no localStorage

2.  Estrutura Detectada

-   Meses disponíveis
-   Métricas (Valor, Hora etc.)
-   Colunas base
-   Colunas extras dinâmicas
-   Agrupamentos possíveis

3.  Estatística Aplicada Para cada grupo: μ = média histórica σ = desvio
    padrão amostral Z = (Valor_ref - μ) / σ

Limites: LCL = μ - 3σ UCL = μ + 3σ

Classificação: |Z| ≤ 2 → Aceitável 2 < |Z| ≤ 3 → Alerta |Z| > 3 → Fora σ
= 0 → Sem histórico

4.  Recursos

-   Grid mensal completo
-   Totais por mês (Total / Fora / Alerta)
-   Filtros dinâmicos
-   Agrupamento configurável
-   Exportação TXT e Excel
-   Armazenamento local

5.  Estrutura do Excel Obrigatório:

-   Código
-   Descrição
-   Colunas no padrão: MMM/AA - Métrica

Exemplo: JUN/25 - Valor JUL/25 - Valor AGO/25 - Valor

Colunas adicionais (ex: Função) são detectadas automaticamente.

6.  Aplicações Estratégicas

-   Auditoria de verbas
-   Controle estatístico de variações
-   Governança de folha
-   Análise por Centro de Custo, Empresa ou Função

Projeto Igarapé Digital.
