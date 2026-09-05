# 📐 Dicionário de Fórmulas, Métricas e Engenharia Financeira
## Planilha: `Controle de Investimentos.xlsx`

Este documento apresenta a especificação matemática, técnica e funcional da planilha **`Controle de Investimentos.xlsx`**, desenvolvida para **Controle de Investimentos, Simulação Financeira Multiativos (Fundos Imobiliários, Ações de Dividendos e Renda Fixa com Aportes)** e planejamento patrimonial de longo prazo para a **Digital Innovation One (DIO)**.

---

## 1. Mapeamento de Parâmetros de Entrada (Inputs)

Os dados de entrada do investidor estão concentrados no intervalo `B5:D10`:

| Célula | Parâmetro | Valor Padrão | Descrição e Impacto no Modelo |
| :---: | :--- | :---: | :--- |
| `C6` | **Aporte Inicial** | `R$ 5.000,00` | Capital já disponível para início imediato da carteira de FIIs. |
| `C7` | **Aporte Mensal** | `R$ 1.000,00` | Capacidade de poupança mensal regular do investidor. |
| `C8` | **Prazo (Anos)** | `10 anos` | Horizonte temporal da estratégia ($10 \times 12 = 120\text{ meses}$). |
| `C9` | **Taxa Rendimento (% a.m.)** | `0,85%` | Retorno mensal total composto estimado (ganho de capital + proventos). |
| `C10` | **Dividend Yield (% a.m.)** | `0,60%` | Taxa de proventos mensais líquidos distribuídos em conta pelas cotas. |

---

## 2. Fórmulas dos Cartões de KPIs Executivos

Os cartões de resumo executivo sintetizam os resultados nos intervalos `F5:K10`:

### 2.1. Linha Superior de KPIs (`Linha 6`)

| Métrica | Célula | Fórmula em Inglês (EN-US) | Fórmula em Português (PT-BR) | Resultado Padrão |
| :--- | :---: | :--- | :--- | :---: |
| **Patrimônio Acumulado** | `F6` | `=FV(C9, C8*12, -C7, -C6)` | `=VF(C9; C8*12; -C7; -C6)` | **R$ 221.014,14** |
| **Total Aportado** | `H6` | `=C6+(C7*C8*12)` | `=C6+(C7*C8*12)` | **R$ 125.000,00** |
| **Lucro em Juros** | `I6` | `=F6-H6` | `=F6-H6` | **R$ 96.014,14** |
| **Dividendos Mensais** | `J6` | `=F6*C10` | `=F6*C10` | **R$ 1.326,08** |

### 2.2. Linha Inferior de KPIs (`Linha 10`)

| Métrica | Célula | Fórmula em Inglês (EN-US) | Fórmula em Português (PT-BR) | Resultado Padrão |
| :--- | :---: | :--- | :--- | :---: |
| **Efeito Bola de Neve** | `F10` | `=IF(J6>=C7, "ATIVADO! (Proventos > Aporte)", "EM FORMAÇÃO")` | `=SE(J6>=C7; "ATIVADO! (Proventos > Aporte)"; "EM FORMAÇÃO")` | **ATIVADO!** |
| **Renda Anual em Proventos** | `H10` | `=J6*12` | `=J6*12` | **R$ 15.913,02** |
| **Cobertura do Aporte** | `J10` | `=J6/C7` | `=J6/C7` | **132,61%** |

---

## 3. Mecânica da Tabela Mês a Mês (`Table_1` - Linhas 13 a 43)

A tabela progressiva detalha o comportamento do patrimônio mês a mês (período inicial de 30 meses):

| Coluna | Campo | Lógica Matemática | Fórmula Utilizada (Linha $i$) |
| :---: | :--- | :--- | :--- |
| **B** | **Mês** | Contador inteiro sequencial | `1, 2, 3, ..., 30` |
| **C** | **Saldo Inicial** | Saldo final do mês anterior (ou capital inicial no mês 1) | Mês 1: `=$C$6`<br>Mês 2 em diante: `=F14` |
| **D** | **Aporte** | Injeção constante de capital | `=$C$7` |
| **E** | **Rendimento (Juros)** | Juros aplicados sobre o montante no mês | `=(C14+D14)*$C$9` |
| **F** | **Saldo Final Acumulado** | Saldo Inicial + Aporte + Rendimento | `=C14+D14+E14` |
| **G** | **Total Aportado Acumulado** | Capital inicial + soma dos aportes realizados | `=$C$6+(B14*$C$7)` |
| **H** | **Dividendos Estimados** | Proventos mensais gerados pelas cotas acumuladas | `=F14*$C$10` |

---

## 4. Análise Crítica e Sugestões de Melhorias

Avaliando a planilha `Controle de Investimentos.xlsx` em relação aos objetivos do Desafio da DIO, destacamos as seguintes **correções e aprimoramentos recomendados**:

### 🎯 Sugestão 1: Harmonização do Prazo (10 Anos vs. 30 Meses)
- **Diagnóstico:** O input em `C8` indica `10 anos` (120 meses), e o card `F6` calcula o patrimônio em 120 meses (R$ 221.014,14). Porém, a tabela dinâmica para no mês **30** (R$ 40.742,87).
- **Correção Simples:**
  - Adicionar uma pequena tabela de **Marcos Temporais (Cenários)** ao lado da tabela:
    - 1 Ano (12 meses): `=VF($C$9; 12; -$C$7; -$C$6)`
    - 2 Anos (24 meses): `=VF($C$9; 24; -$C$7; -$C$6)`
    - 5 Anos (60 meses): `=VF($C$9; 60; -$C$7; -$C$6)`
    - 10 Anos (120 meses): `=VF($C$9; 120; -$C$7; -$C$6)`
  - Isso mantém a planilha leve (sem precisar criar 120 linhas) e oferece ao investidor a visão clara de curto, médio e longo prazo!

### 🎯 Sugestão 2: Adição do Módulo de Alocação por Tipo de FII (Alinhamento DIO)
- **Diagnóstico:** A planilha atual simula a taxa global de retorno, mas o desafio da DIO aborda a diversificação de carteira entre segmentos de Fundos Imobiliários.
- **Correção Intuitiva:**
  - Criar um bloco simples de 4 a 5 linhas abaixo dos inputs dividindo o **Aporte Mensal** (`$C$7`):
    - **FIIs de Tijolo (ex.: 40%):** `=$C$7 * 40%` (R$ 400,00)
    - **FIIs de Papel/Recebíveis (ex.: 35%):** `=$C$7 * 35%` (R$ 350,00)
    - **FIIs Híbridos / FOFs (ex.: 15%):** `=$C$7 * 15%` (R$ 150,00)
    - **Oportunidades / Desenvolvimento (ex.: 10%):** `=$C$7 * 10%` (R$ 100,00)
  - Essa adição é extremamente simples de fazer e atende 100% dos requisitos didáticos da DIO!

### 🎯 Sugestão 3: Adoção de Nomes Definidos (Name Manager)
- **Diagnóstico:** As fórmulas utilizam coordenadas diretas como `$C$6`, `$C$7`, `$C$9`.
- **Melhoria de Usabilidade:**
  - No Excel, selecionar `C6` e nomear como `Aporte_Inicial`.
  - Selecionar `C7` e nomear como `Aporte_Mensal`.
  - A fórmula de patrimônio se torna:
    ```excel
    =VF(Taxa_Mensal; Prazo_Anos*12; -Aporte_Mensal; -Aporte_Inicial)
    ```
  - Isso torna a planilha intuitiva até para quem nunca usou Excel avançado.

### 🎯 Sugestão 4: Compatibilidade de Idioma (PT-BR vs EN-US)
- No Excel em Português, as funções em inglês `FV` e `IF` devem ser inseridas como `VF` e `SE`. Caso um usuário digite em inglês numa versão em português, o Excel pode retornar o erro `#NOME?`.
- O `README.md` documenta ambas as versões para assegurar compatibilidade universal.
