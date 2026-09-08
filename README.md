# 🚚 Sistema de Gestão de Abastecimentos (Excel / VBA)

> Sistema integrado em planilha para controle completo de abastecimentos, gestão de frota, motoristas, fornecedores e análise de indicadores operacionais (KM/L, CPK e custos).

---

## 📌 Sobre o Projeto

O **Sistema de Gestão de Abastecimentos** foi desenvolvido para centralizar, automatizar e otimizar o controle de consumo de combustível e custos operacionais de frotas, uma vez que o gasto com combustíveis representa grande parcela dos custos da operação das transportadoras. Com uma interface moderna, o sistema oferece navegação fluida por menu interativo e dashboards dinâmicos para tomada de decisão estratégica.

---

## 🚀 Funcionalidades Principais

### 📋 Módulo de Cadastros Base
* **Cadastro de Veículos:** Cadastro detalhado com placa, número de frota, fabricante de carroceria/chassi, modelo da corroceria/chassi, ano de fabricação, ano modelo, idade, situação (Ativo/Desativado), RENAVAM, CRV, espécie, tipo, Nº chassi, cor, categoria, potência, PBT, capacidade, Nº do motor, CMT, eixos, lotação, CNPJ/CPF, empresa, combustível, capacidade do tanque, arla (Sim/Não) e capacidade tanque arla.
* **Cadastro de Motoristas:** Controle de condutores com registro do nome completo, número da CNH, categoria, data de vencimento, contatos telefônicos, e-mail, endereço completo e status de atividade.
* **Cadastro de Fornecedores:** Registro de postos de combustível parceiros, CNPJ, localização e contatos comerciais/financeiros.
* **Configurações Gerais:** Parametrizações flexíveis de espécies de veículos, categorias, tipos de combustíveis (Gasolina, Diesel S10, Diesel S500, Etanol, Arla) e tipos de transporte.

---

### ⛽ Módulo de Lançamentos
* **Registro de Abastecimentos:** Lançamento simples e rápido associando veículo, número de frota, fornecedor, tipo de combustível, litragem, valor unitário, valor total, medição do hodômetro (KM), motorista responsável e número da NFC-e/Chave da nota fiscal.

---

### 📊 Relatórios e Indicadores Operacionais
* **Relatório Automático de Consumo:**
  * Cálculo dinâmico do **KM Rodado** entre abastecimentos.
  * Desempenho do veículo medido em **KM/L** com indicação visual de eficiência.
  * Custo por quilômetro rodado (**R$/KM**).
  * Filtros por período, veículo, motorista, tipo de combustível e número de frota.

---

### 📈 Dashboard Gerencial Interativo
Painel executivo com atualização de dados e filtros dinâmicos por Ano, Mês e Veículo:

* **KPIs Principais:**
  * **Gasto Total com Combustível** (R$)
  * **Preço Médio do Litro** (R$)
  * **Consumo Médio Geral** (KM/L)
  * **Custo Por KM (CPK)** (R$)
  * Total de Veículos na Frota e Motoristas Ativos.

* **Visualizações Gráficas:**
  * **Evolução de Gastos:** Acompanhamento mensal dos custos com combustível.
  * **Proporção de Consumo:** Distribuição do volume consumido por tipo de combustível.
  * **Ranking por Motorista:** Comparativo de desempenho médio (KM/L) entre condutores.
  * **Consumo por Veículo:** Tabela detalhada de médias com sinalizadores visuais de desempenho.
  * **Análise de Fornecedores:** Preço médio do litro, valor total abastecido e número de abastecimentos realizados por posto.

---

## 📸 Interface do Sistema

## 📸 Interface do Sistema

### 📌 Menu Principal
![Menu Inicial](assets/01_Menu.PNG)

### 🧭 Painel de Navegação
![Painel de Navegação](assets/02_Painel_de_Selecao.PNG)

### 🚍 Cadastro de Veículos
![Cadastro de Veículos](assets/03_Cadastro_de_Veiculos.PNG)

### 👨‍✈️ Cadastro de Motoristas
![Cadastro de Motoristas](assets/04_Cadastro_de_Motoristas.PNG)

### 🏪 Cadastro de Fornecedores
![Cadastro de Fornecedores](assets/05_Cadastro_de_Fornecedores.PNG)

### ⚙️ Cadastro Geral
![Cadastro Geral](assets/06_Cadastro_de_Geral.PNG)

### ⛽ Lançamento de Abastecimentos
![Abastecimentos](assets/07_Abastecimentos.PNG)

### 📄 Relatório de Consumo
![Relatório de Consumo](assets/08_Relatorio_de_Abastecimentos.PNG)

### 📊 Dashboard Gerencial
![Dashboard Gerencial](assets/09_Dashboard.PNG)

---

## 🛠️ Tecnologias e Recursos Técnicos Utilizados

### 🎛️ Customização de Interface (Custom UI)
* **Office RibbonX Editor (Custom XML):** Customização avançada da faixa de opções do Excel via código XML (`customUI.xml`). Criação de uma barra de ferramentas/menu totalmente personalizada no topo da aplicação, com grupos dedicados (*Cadastros, Lançamentos, Relatórios, Dashboards, Ações*), ícones customizados e callbacks integrados.

---

### ⚙️ VBA (Visual Basic for Applications)
* **Callbacks da Ribbon:** Integração do menu XML com rotinas em VBA para disparo de ações ao clicar nos botões da barra superior.
* **Automação de Interface:** Macros direcionadas para navegação rápida entre abas, atualização de base de dados e controle de tela.
* **Gerenciamento de Dados:** Rotinas para inserção e exclusão dinâmica de linhas mantendo a integridade de fórmulas, bordas e validações.

---

### 📊 Microsoft Excel & Modelagem de Dados
* **Fórmulas e Funções Utilizadas:**
  * `PROCV` / `INDIRETO`: Busca e manipulação dinâmica de referências entre as abas de cadastro e lançamentos.
  * `FILTRO` / `ESCOLHERCOLS`: Matrizes dinâmicas para filtragem e seleção específica de colunas nos relatórios gerenciais.
  * `AGREGAR`: Cálculos avançados com capacidade de ignorar erros e linhas ocultas.
  * `SE` / `SEERRO`: Lógica condicional e prevenção de erros de divisão ou busca (ex: `#DIV/0!`, `#N/D`).
  * `SOMA` / `MÉDIA`: Consolidação de volumes, totais financeiros e médias operacionais de consumo (KM/L).
  * `HOJE` / `ANO` / `TEXTO`: Manipulação, extração e formatação dinâmica de datas.
  * `MAIÚSCULA` / `LIN`: Padronização de textos e controle posicional/sequencial de linhas.
* **Tabelas Dinâmicas & Modelagem:** Estruturação dos dados brutos para alimentar os indicadores do Dashboard.
* **Segmentação de Dados (Slicers):** Filtros interativos aplicados ao Dashboard por **Ano**, **Mês** e **Veículo**.
* **Validação de Dados:** Listas suspensas dinâmicas (*Data Validation*) para padronização de preenchimento.
* **Formatação Condicional:** Alertas e sinalizadores visuais para acompanhamento de eficiência em KM/L e status de cadastros (Ativo/Desativado).

---

### 🎨 Design & UX/UI (User Experience)
* **Interface Dark Mode:** Visual estilizado em tons escuros focado na experiência do usuário e usabilidade.
* **Proporção e Layout de Sistema:** Estruturação visual projetada para parecer uma aplicação desktop dedicada, ocultando elementos padrão da planilha quando necessário.

---
