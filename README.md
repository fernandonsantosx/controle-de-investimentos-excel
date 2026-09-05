# 📊 Controle de Investimentos & Simulador Financeiro em Excel
### Desafio de Projeto — Digital Innovation One (DIO)

<div align="center">

![Excel](https://img.shields.io/badge/Microsoft_Excel-Controle_de_Investimentos.xlsx-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![DIO](https://img.shields.io/badge/DIO-Desafio_de_Projeto-002B49?style=for-the-badge&logo=dio&logoColor=white)
![Finanças](https://img.shields.io/badge/Investimentos-Multi--Ativos-F7931A?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Concluído-success?style=for-the-badge)
![Licença](https://img.shields.io/badge/Licença-MIT-blue?style=for-the-badge)

</div>

---

## 📌 Sumário
- [📖 Sobre o Projeto](#-sobre-o-projeto)
- [🎯 Objetivos e Alinhamento com a DIO](#-objetivos-e-alinhamento-com-a-dio)
- [⚙️ Arquitetura da Planilha `Controle de Investimentos.xlsx`](#️-arquitetura-da-planilha-controle-de-investimentosxlsx)
- [🧮 Modelagem Matemática e Fórmulas Utilizadas](#-modelagem-matemática-e-fórmulas-utilizadas)
- [❄️ O Conceito do Efeito Bola de Neve (Magic Number)](#️-o-conceito-do-efeito-bola-de-neve-magic-number)
- [💡 Sugestões de Alterações e Correções para Alinhamento](#-sugestões-de-alterações-e-correções-para-alinhamento)
- [📁 Estrutura do Repositório](#-estrutura-do-repositório)
- [🚀 Como Utilizar e Testar a Planilha](#-como-utilizar-e-testar-a-planilha)
- [🌐 Como Publicar no GitHub e Entregar na Plataforma DIO](#-como-publicar-no-github-e-entregar-na-plataforma-dio)
- [📜 Licença](#-licença)

---

## 📖 Sobre o Projeto

Este repositório contém uma solução completa e versátil voltada para **Controle de Investimentos e Simulação Financeira Multiativos (Fundos Imobiliários, Ações de Dividendos e Renda Fixa com Aportes)**, atendendo e expandindo os objetivos do desafio prático da **Digital Innovation One (DIO)**. 

O projeto tem como núcleo o arquivo **[`Controle de Investimentos.xlsx`](Controle de Investimentos.xlsx)**, uma planilha interativa, elegante e intuitiva desenvolvida para automatizar cálculos de juros compostos, projeção de patrimônio, acúmulo de proventos e monitoramento do ponto de virada financeira conhecido como **Efeito Bola de Neve** (*Magic Number*).

A ferramenta foi projetada com arquitetura abrangente, permitindo que qualquer investidor consiga planejar tanto ativos imobiliários (FIIs) quanto qualquer carteira geradora de renda passiva:
1. Quanto seu patrimônio crescerá ao longo dos anos com aportes constantes e reinvestimento dos proventos.
2. Qual será sua renda passiva mensal estimada em dividendos e juros.
3. Em que momento os rendimentos gerados pelas cotas e títulos passarão a cobrir e superar o próprio aporte mensal.

---

## 🎯 Objetivos e Alinhamento com a DIO

- [x] **Construção de Ferramenta Financeira Prática:** Aplicação de funções financeiras essenciais do Excel (`VF` / `FV`, `SE` / `IF`, fórmulas dinâmicas).
- [x] **Automatização de Cálculos Complexos:** Total aportado, valor futuro projetado, lucro acumulado em juros e estimativa de dividendos mensais e anuais.
- [x] **Visualização com Tabela e Gráficos:** Tabela progressiva mês a mês e gráfico de linhas destacando o distanciamento exponencial entre o capital aportado e o patrimônio total.
- [x] **Interface Executiva e Intuitiva:** Organização em blocos visuais claros (Inputs $\rightarrow$ Cards de KPIs $\rightarrow$ Tabela de Projeção $\rightarrow$ Gráfico).
- [x] **Documentação Completa de Engenharia:** Guia técnico com equivalências de fórmulas em Português e Inglês e boas práticas de versionamento no GitHub.

---

## ⚙️ Arquitetura da Planilha `Controle de Investimentos.xlsx`

A planilha organiza o fluxo de análise de forma lógica e sequencial:

```mermaid
graph TD
    A[Inputs: Aporte Inicial, Aporte Mensal, Prazo, Taxa e DY] --> B[Cards de KPIs Executivos]
    A --> C[Tabela de Evolução Mês a Mês - 30 Meses]
    B --> D[Patrimônio Futuro - Função VF]
    B --> E[Dividendos Mensais e Anuais]
    B --> F[Efeito Bola de Neve - Função SE]
    C --> G[Gráfico de Evolução: Aportado vs Saldo Acumulado]
```

### 1. Painel de Parâmetros de Investimento (`Inputs: B5:D10`)
Área reservada para personalização pelo usuário:
- **Aporte Inicial (`C6`):** Capital inicial de largada (padrão: `R$ 5.000,00`).
- **Aporte Mensal (`C7`):** Capacidade de poupança mensal (padrão: `R$ 1.000,00`).
- **Prazo em Anos (`C8`):** Duração da estratégia (padrão: `10 anos = 120 meses`).
- **Taxa Rendimento (% a.m.) (`C9`):** Rendimento total composto mensal (padrão: `0,85% a.m.` $\approx 10,7\%$ a.a.).
- **Dividend Yield (% a.m.) (`C10`):** Proventos líquidos mensais (padrão: `0,60% a.m.` $\approx 7,4\%$ a.a.).

### 2. Painel de Resumo Executivo (Cards de KPIs - `F5:K10`)
- **Patrimônio Acumulado (`F6`):** `R$ 221.014,14` (calculado via função financeira `VF` para 120 meses).
- **Total Aportado (`H6`):** `R$ 125.000,00` (capital real desembolsado pelo investidor).
- **Lucro em Juros (`I6`):** `R$ 96.014,14` (ganho líquido proporcionado pelos juros compostos).
- **Dividendos Mensais (`J6`):** `R$ 1.326,08 / mês` (renda passiva mensal estimada ao fim do período).
- **Efeito Bola de Neve (`F10`):** `ATIVADO! (Proventos > Aporte)` (condição dinâmica via função `SE`).
- **Renda Anual em Dividendos (`H10`):** `R$ 15.913,02 / ano`.
- **Cobertura do Aporte Mensal (`J10`):** `132,61%` (proventos cobrem 1,32x o valor do aporte!).

### 3. Tabela de Projeção Mês a Mês (`Table_1` - Linhas 13 a 43)
Acompanhamento detalhado da evolução patrimonial ao longo dos primeiros 30 meses:
- **Mês:** Contador de 1 a 30.
- **Saldo Inicial:** Inicia com o aporte inicial e recebe o saldo final do mês anterior.
- **Aporte:** Depósito mensal recorrente de R$ 1.000,00.
- **Rendimento:** Ganho percentual sobre o saldo em conta.
- **Saldo Final Acumulado:** Soma de Saldo + Aporte + Rendimento.
- **Total Aportado:** Montante do bolso do investidor.
- **Dividendos Estimados:** Geração de caixa mensal correspondente ao saldo atingido.

### 4. Gráfico Dinâmico de Evolução Patrimonial
Gráfico de linhas intitulado *"Evolução Patrimonial: Total Aportado vs. Saldo Acumulado"*, demonstrando visualmente o efeito da aceleração dos juros compostos.

---

## 🧮 Modelagem Matemática e Fórmulas Utilizadas

A planilha foi construída com fórmulas compatíveis com o Excel internacional e brasileiro:

| Métrica / Campo | Fórmula em Inglês (EN-US) | Fórmula em Português (PT-BR) | Finalidade |
| :--- | :--- | :--- | :--- |
| **Patrimônio Final** | `=FV(C9, C8*12, -C7, -C6)` | `=VF(C9; C8*12; -C7; -C6)` | Projeta o valor futuro da série periódica com depósito inicial. |
| **Total Aportado** | `=C6+(C7*C8*12)` | `=C6+(C7*C8*12)` | Soma do capital inicial com os aportes dos 120 meses. |
| **Lucro em Juros** | `=F6-H6` | `=F6-H6` | Subtrai o total aportado do patrimônio acumulado. |
| **Dividendo Mensal** | `=F6*C10` | `=F6*C10` | Aplica a taxa de Dividend Yield sobre o patrimônio final. |
| **Status Bola de Neve** | `=IF(J6>=C7, "ATIVADO! (Proventos > Aporte)", "EM FORMAÇÃO")` | `=SE(J6>=C7; "ATIVADO! (Proventos > Aporte)"; "EM FORMAÇÃO")` | Avalia dinamicamente se a renda passiva já superou o aporte. |
| **Renda Anual** | `=J6*12` | `=J6*12` | Anualiza a renda mensal estimada em proventos. |
| **Cobertura de Aporte** | `=J6/C7` | `=J6/C7` | Calcula a razão entre dividendos recebidos e o valor aportado. |
| **Rendimento no Mês** | `=(C14+D14)*$C$9` | `=(C14+D14)*$C$9` | Calcula o rendimento mensal da tabela dinâmica. |

> 📘 **Para aprofundamento na matemática financeira:** Consulte [`FORMULAS_E_METRICAS.md`](FORMULAS_E_METRICAS.md).

---

## ❄️ O Conceito do Efeito Bola de Neve (Magic Number)

No mercado de Fundos Imobiliários, o **Efeito Bola de Neve** (ou *Número Mágico*) representa o divisor de águas na vida do investidor:

$$\text{Dividendos Mensais Recebidos} \ge \text{Aporte Mensal Regular}$$

Quando essa condição é atingida:
1. O investidor não precisa mais tirar recursos exclusivamente do próprio salário para continuar investindo.
2. A carteira torna-se **autossustentável**: o próprio fluxo de proventos compra novas cotas todos os meses.
3. No modelo da planilha, com aportes de R$ 1.000/mês e DY de 0,60% a.m., o efeito bola de neve é atingido plenamente antes dos 10 anos, gerando **R$ 1.326,08/mês** (cobertura de **132,6%**)!

---

## 💡 Sugestões de Alterações e Correções para Alinhamento

Para tornar a planilha ainda mais simples, intuitiva e 100% alinhada à ementa do Desafio da DIO, sugerem-se as seguintes melhorias:

### 1. Harmonização do Prazo (10 Anos vs Tabela de 30 Meses)
- **Cenário Atual:** O input define 10 anos (120 meses), mas a tabela apresenta apenas os primeiros 30 meses.
- **Sugestão:** Adicionar uma pequena tabela lateral com **Marcos de Cenários** para 1 ano (12m), 2 anos (24m), 5 anos (60m) e 10 anos (120m). Dessa forma, a planilha continua leve e intuitiva, mas o usuário enxerga a correspondência exata com o prazo final!

### 2. Módulo de Alocação por Tipo de FII (Requisito Temático da DIO)
- **Cenário Atual:** O modelo calcula o retorno geral, mas não especifica em quais fundos imobiliários o aporte é distribuído.
- **Sugestão:** Adicionar um pequeno bloco abaixo dos inputs com a distribuição recomendada do aporte mensal (`$C$7` = R$ 1.000,00):
  - **FIIs de Tijolo (ex.: 40%):** R$ 400,00 (galpões logísticos, shoppings e lajes corporativas)
  - **FIIs de Papel/Recebíveis (ex.: 35%):** R$ 350,00 (CRIs e títulos de crédito com juros atrativos)
  - **FIIs Híbridos / FOFs (ex.: 15%):** R$ 150,00 (diversificação de gestão)
  - **Desenvolvimento / Oportunidades (ex.: 10%):** R$ 100,00 (maior potencial de valorização)

### 3. Nomes Definidos no Excel (Name Manager)
- **Sugestão:** Em vez de fórmulas com `$C$6` e `$C$7`, nomear as células para `Aporte_Inicial` e `Aporte_Mensal`. Isso facilita a compreensão por qualquer recrutador ou avaliador que abrir a planilha.

### 4. Inclusão de Fórmulas Bilíngues
- Como alguns computadores utilizam a versão em português do Excel (onde `FV` e `IF` geram erro `#NOME?`), manter o guia com `VF` e `SE` documentado no README garante total acessibilidade.

---

## 📁 Estrutura do Repositório

```text
controle-de-investimentos/
│
├── .gitignore                    # Regras de exclusão de arquivos temporários do Office
├── Controle de Investimentos.xlsx # Planilha principal desenvolvida para simulação
├── FORMULAS_E_METRICAS.md        # Documentação matemática e técnica detalhada
└── README.md                     # Documentação oficial do projeto
```

---

## 🚀 Como Utilizar e Testar a Planilha

1. **Abra o arquivo:**
   - Abra [`Controle de Investimentos.xlsx`](Controle de Investimentos.xlsx) no Microsoft Excel, Excel Online ou Google Planilhas.
2. **Experimente novos cenários alterando as células em amarelo/verde:**
   - Mude o **Aporte Inicial** (`C6`) para R$ 1.000,00 ou R$ 10.000,00.
   - Ajuste o **Aporte Mensal** (`C7`) de acordo com sua realidade.
   - Modifique o **Prazo** (`C8`) para ver o impacto no patrimônio final e nos dividendos.
   - Veja o indicador **Efeito Bola de Neve** mudar de status automaticamente conforme a renda supera o aporte!

---

## 🌐 Como Publicar no GitHub e Entregar na Plataforma DIO

### Passo 1: Criar o Repositório no GitHub
1. Acesse o [GitHub](https://github.com/) e crie um novo repositório público com o nome: `controle-de-investimentos-excel`.
2. Não marque para criar README automático (pois já temos um arquivo completo).

### Passo 2: Fazer o Upload dos Arquivos
1. Na página do repositório no GitHub, clique em **`uploading an existing file`**.
2. Selecione ou arraste os seguintes arquivos desta pasta:
   - `Controle de Investimentos.xlsx`
   - `README.md`
   - `FORMULAS_E_METRICAS.md`
   - `.gitignore`
3. Mensagem do commit: `feat: entrega do projeto de controle de investimentos e simulador financeiro`
4. Clique em **Commit changes**.

### Passo 3: Enviar a Entrega na DIO
1. Copie o link do repositório no GitHub.
2. Acesse a plataforma da DIO, clique em **Entregar Projeto** e envie o link acompanhado de uma descrição destacando:
   - Planilha executiva com cálculo de Valor Futuro (`VF`/`FV`) e juros compostos;
   - Indicador dinâmico do **Efeito Bola de Neve** e cobertura de aporte mensal;
   - Tabela de projeção progressiva mês a mês e gráfico visual;
   - Documentação de engenharia e sugestões de aprimoramento.

---

## 📜 Licença

Distribuído sob a licença [MIT](https://opensource.org/licenses/MIT). Sinta-se livre para usar, aprimorar e compartilhar!

---

<div align="center">

Desenvolvido para o Desafio de Projeto da **DIO** 🚀  
*Construindo patrimônio com inteligência e disciplina financeira.*

</div>
