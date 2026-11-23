# 📦 Sistema de Controle de Estoque | Excel e Power Query

Projeto de portfólio voltado ao desenvolvimento de um sistema de controle de estoque para um minimercado fictício, utilizando Microsoft Excel e Power Query.

O objetivo é aplicar recursos avançados do Excel na organização de produtos, no acompanhamento de movimentações, na atualização de dados e na construção de relatórios gerenciais.

> **Status:** Em desenvolvimento. As funcionalidades descritas abaixo fazem parte do planejamento e serão implementadas progressivamente.

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

## 🛠️ Tecnologias Previstas

- **Microsoft Excel:** controle operacional, fórmulas, validação de dados, formatação condicional e relatórios.
- **Power Query:** importação, limpeza, transformação e consolidação dos dados.
- **Tabelas Dinâmicas e gráficos:** análises e visualizações interativas.
- **Python:** geração inicial dos dados sintéticos.
- **Git e GitHub:** versionamento e documentação.

---

## 🗃️ Base de Dados

O projeto utilizará dados sintéticos que representam as operações de um minimercado.

### Estrutura planejada

| Conjunto de dados | Descrição |
|---|---|
| Produtos | Cadastro, categorias, estoque inicial e estoque mínimo |
| Fornecedores | Informações dos fornecedores fictícios |
| Movimentações históricas | Entradas e saídas realizadas no período simulado |
| Lançamentos manuais | Novas movimentações registradas diretamente no Excel |

Os arquivos históricos serão armazenados em formato CSV e importados por meio do Power Query.

**Premissa de negócio:** o estoque disponível será calculado a partir do estoque inicial, das entradas e das saídas de mercadorias.

---

## 📁 Estrutura do Projeto

```text
controle-estoque-excel/
│
├── dados/
├── excel/
├── python/
│
├── .gitignore
└── README.md
```

As pastas serão preenchidas conforme o desenvolvimento.

---

## 🚧 Desenvolvimento

O projeto está em fase inicial.

As próximas etapas incluem:

1. Definição do modelo de dados;
2. Criação de uma base de testes;
3. Importação e tratamento com Power Query;
4. Geração dos dados sintéticos;
5. Implementação dos lançamentos operacionais;
6. Construção do controle de estoque;
7. Desenvolvimento dos relatórios e dashboard;
8. Testes, documentação e entrega final.

As instruções de instalação e utilização serão adicionadas quando as funcionalidades estiverem implementadas.

---

## ℹ️ Limitações Previstas

- Os dados serão fictícios e não representarão uma empresa real.
- A solução será desenvolvida para utilização local no Excel Desktop.
- A versão inicial não será destinada a múltiplos usuários registrando operações simultaneamente.
- A atualização das consultas será necessária para consolidar novas movimentações nos relatórios.

---

## ✅ Resultado Esperado

Construir uma solução funcional de controle de estoque, demonstrando conhecimentos em Excel avançado, Power Query, automação de processos e análise de dados.

O projeto tem como finalidade desenvolver habilidades práticas relevantes para oportunidades de entrada nas áreas de Dados/BI, TI e funções administrativas.
