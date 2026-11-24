# 📦 Sistema de Controle de Estoque | Excel e Power Query

Projeto de portfólio voltado ao desenvolvimento de um sistema de controle de estoque para um minimercado fictício, utilizando Microsoft Excel e Power Query.

O objetivo é aplicar recursos avançados do Excel na organização de produtos, no acompanhamento de movimentações, na atualização de dados e na construção de relatórios gerenciais.

> **Status:** Em desenvolvimento. A geração dos dados sintéticos está concluída. As funcionalidades do sistema Excel serão implementadas progressivamente.

---

## 🎯 Objetivo

Desenvolver uma solução que permita:

- Organizar o cadastro de produtos e fornecedores;
- Importar e tratar dados utilizando Power Query;
- Registrar entradas e saídas de mercadorias;
- Calcular o estoque disponível por produto;
- Identificar produtos abaixo do estoque mínimo;
- Verificar inconsistências nos registros;
- Atualizar os dados sem perder lançamentos manuais;
- Acompanhar indicadores em um dashboard no Excel.

---

## 🛠️ Tecnologias do Projeto

- **Microsoft Excel:** controle operacional, fórmulas, validação de dados, formatação condicional e relatórios.
- **Power Query:** importação, limpeza, transformação e consolidação dos dados.
- **Tabelas Dinâmicas e gráficos:** análises e visualizações interativas.
- **Python:** geração dos dados sintéticos e validação da consistência do histórico.
- **Git e GitHub:** versionamento e documentação.

O Python é utilizado na preparação da base histórica. O Excel e o Power Query serão as tecnologias principais do sistema.

---

## 🗃️ Base de Dados

A base é **sintética**, gerada pelo script `python/gerar_dados.py`, com semente definida para permitir a reprodução dos registros.

O conjunto de dados inclui:

- **150 produtos** distribuídos em diferentes categorias;
- **12 fornecedores fictícios**, localizados em Pernambuco;
- **6.000 movimentações históricas** de entrada e saída;
- Período de janeiro a outubro de 2026;
- Estoques iniciais, estoques mínimos e custos unitários;
- Diferentes frequências de movimentação entre produtos;
- Inconsistências textuais controladas para tratamento com Power Query.

### Estrutura dos Dados

| Arquivo | Descrição |
|---|---|
| `produtos.csv` | Cadastro de produtos, categorias, fornecedores, custos e informações de estoque |
| `fornecedores.csv` | Cadastro dos fornecedores fictícios e suas localizações |
| `movimentacoes_historicas.csv` | Histórico de entradas e saídas de mercadorias |

Os arquivos CSV serão utilizados como fontes de dados no Power Query.

### Premissas de Negócio

O projeto representa um minimercado fictício abastecido por fornecedores regionais de Pernambuco.

O estoque disponível será calculado pela fórmula:

**Estoque Atual = Estoque Inicial + Entradas − Saídas**

As movimentações históricas foram geradas respeitando a disponibilidade de mercadorias, evitando saldos negativos durante a simulação.

A base também contempla produtos com diferentes níveis de estoque, permitindo desenvolver indicadores e alertas de reposição.

Os dados são fictícios e não representam operações de uma empresa real.

---

## 🐍 Geração dos Dados com Python

O arquivo `python/gerar_dados.py` é responsável pela geração da base histórica utilizada no projeto.

O gerador permite:

- Criar os cadastros fictícios de produtos e fornecedores;
- Simular entradas e saídas de mercadorias;
- Produzir um histórico com datas e identificadores consistentes;
- Associar os produtos aos fornecedores cadastrados;
- Respeitar as quantidades disponíveis durante a geração;
- Reproduzir a base utilizando uma semente fixa;
- Exportar os resultados para arquivos CSV.

O script utiliza bibliotecas padrão do Python, sem necessidade de dependências externas.


### Execução do Gerador

**Pré-requisito:** Python 3 instalado.

No terminal, a partir da pasta principal do projeto, execute:

```powershell
python python/gerar_dados.py --sobrescrever
```

O comando gera os três arquivos CSV na pasta `dados/`.

**Atenção:** a opção `--sobrescrever` autoriza substituir os CSVs existentes. Os arquivos gerados já estão disponíveis no repositório, portanto não é necessário executar o gerador para utilizar a base histórica.

Não execute novamente o gerador após iniciar registros operacionais sem antes fazer uma cópia de segurança e avaliar os efeitos sobre o histórico.
  

**Observação:** o gerador é utilizado para preparar os dados históricos. A operação diária do sistema não dependerá da execução do Python.

---

## 🔄 Power Query e Tratamento de Dados

O Power Query será responsável pela preparação das informações utilizadas no sistema.

### Transformações Planejadas

- Importação dos arquivos CSV;
- Definição dos tipos de dados;
- Padronização de nomes e categorias;
- Tratamento de espaços e inconsistências textuais;
- Verificação de valores nulos;
- Identificação de duplicatas;
- Verificação de relacionamentos entre os cadastros;
- Consolidação das movimentações históricas e manuais;
- Atualização das consultas sem refazer as transformações.

As inconsistências textuais foram introduzidas de maneira controlada para permitir a aplicação prática de recursos de limpeza e transformação.

Registros com problemas que comprometam sua interpretação deverão ser identificados antes de serem considerados nos cálculos de estoque.

---

## 📗 Sistema de Controle de Estoque no Excel

A aplicação será desenvolvida no Microsoft Excel Desktop.

### Funcionalidades Planejadas

**Controle operacional:**

- Consulta de produtos e quantidades disponíveis;
- Registro de entradas e saídas;
- Validação dos campos preenchidos;
- Atualização dos saldos;
- Alertas de estoque mínimo.

**Automatização e cálculos:**

- Tabelas estruturadas;
- Fórmulas de busca e agregação;
- Validação de Dados;
- Formatação Condicional;
- Integração com as consultas do Power Query.

Os lançamentos realizados pelo usuário deverão permanecer em uma tabela editável, separada dos resultados carregados pelo Power Query.

Essa organização tem como objetivo preservar os registros manuais durante as atualizações.

---

## 📊 Relatórios e Dashboard

Está previsto o desenvolvimento de um dashboard interativo dentro do Microsoft Excel.

### Indicadores Planejados

- Quantidade de produtos cadastrados;
- Produtos abaixo do estoque mínimo;
- Valor estimado do estoque;
- Quantidade de movimentações por período.

### Análises Visuais

- Produtos com maior volume de saídas;
- Distribuição do estoque por categoria;
- Evolução mensal de entradas e saídas;
- Situação dos produtos em relação ao estoque mínimo.

O painel utilizará gráficos, Tabelas Dinâmicas e filtros interativos.

### Visualização do Sistema

As capturas reais do dashboard, do controle de estoque, dos lançamentos e das transformações no Power Query serão adicionadas após a implementação.

---

## ⚙️ Funcionamento Previsto

O sistema utilizará arquivos CSV com os dados históricos gerados pelo Python.

O Power Query importará, tratará e consolidará essas informações. O Excel será utilizado para registrar novas movimentações, calcular os estoques e apresentar os relatórios.

O fluxo será organizado para que:

1. Os dados históricos sejam importados dos arquivos CSV;
2. O Power Query aplique as transformações configuradas;
3. O usuário registre novas entradas e saídas no Excel;
4. As movimentações manuais sejam incorporadas à consolidação;
5. Os saldos e indicadores sejam atualizados.

A versão inicial será destinada ao uso local por um único operador.

As instruções detalhadas de utilização serão adicionadas após a implementação e validação do sistema.

---

## 📁 Estrutura do Projeto

```text
controle-estoque-excel/
│
├── docs/
│
├── dados/
│   ├── fornecedores.csv
│   ├── movimentacoes_historicas.csv
│   └── produtos.csv
│
├── excel/
│
├── python/
│   └── gerar_dados.py
│
├── .gitignore
└── README.md
```

A pasta `excel/` armazenará o arquivo Controle_de_Estoque.xlsx, que será desenvolvido com as consultas do Power Query, o controle operacional, as fórmulas e o dashboard.

As capturas de tela serão organizadas em `docs/` quando estiverem disponíveis.

---

## 🚧 Desenvolvimento

O projeto está em desenvolvimento.

### Etapas Concluídas

- [x] Definição da arquitetura inicial;
- [x] Criação do repositório GitHub;
- [x] Desenvolvimento do gerador Python;
- [x] Geração dos arquivos CSV;
- [x] Validação da consistência da base histórica.

### Próximas Etapas

- [ ] Importação e tratamento definitivo com Power Query;
- [ ] Construção da estrutura operacional no Excel;
- [ ] Implementação dos lançamentos manuais;
- [ ] Cálculo dos estoques e alertas;
- [ ] Desenvolvimento dos relatórios e dashboard;
- [ ] Testes de atualização e integridade;
- [ ] Documentação de execução e capturas do sistema.

---

## ℹ️ Limitações do Projeto

- Os dados utilizados são sintéticos.
- Os fornecedores foram definidos dentro de um recorte regional de Pernambuco.
- A solução será desenvolvida para utilização local no Excel Desktop.
- A versão inicial não será destinada ao registro simultâneo de movimentações por vários operadores.
- A consolidação das novas movimentações dependerá da atualização das consultas.
- O valor estimado do estoque utilizará os custos cadastrados, sem representar uma avaliação contábil completa.

---

## ✅ Resultado Esperado

Construir uma solução funcional de controle de estoque, integrando geração de dados, tratamento automatizado, cálculos operacionais e visualizações gerenciais.

O objetivo principal é demonstrar a aplicação prática de Microsoft Excel avançado e Power Query, complementados por Python e Git/GitHub.

O projeto busca desenvolver e apresentar competências relevantes para oportunidades de entrada nas áreas de Dados/BI, TI e funções administrativas.