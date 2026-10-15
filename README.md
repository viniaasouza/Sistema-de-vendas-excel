# Sistema de Vendas Profissional no Excel VBA 📊🛒

Sistema completo de gestão comercial e ponto de venda (PDV) desenvolvido em **Excel VBA**, baseado na série de aulas do canal **[Excel com Edivam Cabral](https://www.youtube.com/playlist?list=PLk88JSh0Zk6pKxvmUjdFbi7fzB8F2o_Er)**.

---

## 🚀 Funcionalidades do Sistema

### 📦 1. Módulo de Produtos (`FormProdutos`)
- **CRUD Completo:** Inclusão, alteração, exclusão e pesquisa dinâmica de produtos.
- **Prevenção de Duplicidade:** Verificação automática para impedir cadastro de produtos com nomes duplicados.
- **Busca Rápida em Tempo Real:** Filtragem inteligente por aproximação (`Like`) com suporte a caracteres acentuados via função customizada `AcSQL`.
- **Formatação de Moeda Automática:** Conversão e formatação monetária nos campos e na listagem.

### 👥 2. Módulo de Clientes (`FormClientes`)
- Cadastro completo de clientes (Pessoa Física e Pessoa Jurídica): Nome/Razão Social, CPF/CNPJ, Telefone/WhatsApp, E-mail e Endereço.
- Seleção de registros com 1 clique para preenchimento dos campos e alternância entre modos de inclusão e edição.
- Exclusão segura com diálogo de confirmação e prevenção de duplicidade por CPF/CNPJ.

### 💳 3. Ponto de Venda / Emissão de Vendas (`FormVendas`)
- **Numeração Automática:** Geração incremental do número de venda mestre (`ID_VENDA`).
- **Busca Inteligente por Digitação:** Autocomplete dinâmico tanto de clientes (por nome ou CPF) quanto de produtos diretamente no dropdown sem necessidade de rolagem.
- **Carrinho de Compras Interativo (`LstItens`):**
  - Preenchimento instantâneo de preço unitário e cálculo em tempo real do subtotal por item.
  - Duplo clique no carrinho para carregar e editar quantidade/preço do item.
  - Exclusão individual de itens com recálculo automático.
- **Fechamento Financeiro com Inteligência de Cartão:**
  - Soma automática do Total Bruto e aplicação de desconto em valor (R$).
  - Seleção da forma de pagamento (Dinheiro, PIX, Cartão de Crédito, Cartão de Débito, Boleto).
  - **Parcelamento de 1x a 12x** para vendas no cartão de crédito.
  - **Taxa Estimada da Maquininha (%):** Campo ajustável para estimativa do desconto da operadora.
  - **Previsão Líquida a Receber:** Destaque visual do valor líquido real que entrará em caixa.
- **Gravação Relacional (Mestre-Detalhe):**
  - Registro mestre na aba `VENDAS` com total bruto, líquido, parcelas, taxa e previsão a receber.
  - Registro dos itens na aba `ITENS_VENDA`.

### 📑 4. Relatórios Gerenciais e Emissão de PDF (`FormRelatorios` & `R_Vendas`)
- **Painel de Consultas Avançadas:**
  - Filtro por parte do nome do cliente (`LIKE`).
  - Filtro por período de datas (Data Inicial e Data Final).
  - Filtro por Status da Venda (`Concluída` ou `Cancelada`).
  - Filtro por Forma de Pagamento (`PIX`, `Cartão de Crédito`, `Dinheiro`, etc.).
- **Cards de Resumo (KPIs em Tempo Real):**
  - Quantidade de vendas concluídas no período.
  - Total bruto e líquido vendido.
  - Total de taxas retidas pela operadora de cartão.
  - Total líquido real a receber pelo lojista.
- **Gestão de Status de Vendas:**
  - Cancelamento e reativação de vendas com confirmação segura e estorno de totais.
- **Exportação de Relatório Gerencial para PDF:**
  - Geração automática na aba `R_Vendas` com cabeçalho repetido em todas as páginas e bordas automáticas.
  - Salvamento automático na pasta `Relatorios/`.
- **Emissão de Comprovante de Venda (Recibo A4 / Cupom PDF):**
  - Emissão individual com 1 clique (ou duplo clique na venda) contendo cabeçalho da empresa, dados do cliente, lista completa de itens, totais e forma de pagamento.
- **Dados da Empresa Configuráveis (`CONFIG`):**
  - Gerenciamento fácil de Razão Social, Nome Fantasia, CNPJ, Telefone, E-mail, Endereço e Mensagem de Rodapé refletidos automaticamente nos relatórios e recibos.

---

## 🗂️ Estrutura do Repositório

```text
├── sistema de vendas.xlsm       # Planilha oficial habilitada para macros (.xlsm)
├── .gitignore                   # Arquivos ignorados (bloqueios do Excel e backups)
├── README.md                    # Documentação do projeto
└── src/                         # Código fonte dos módulos e formulários VBA
    ├── FormProdutos.frm/.frx    # Formulário e layout do Cadastro de Produtos
    ├── FormClientes.frm/.frx    # Formulário e layout do Cadastro de Clientes
    ├── FormVendas.frm/.frx      # Formulário e layout do Ponto de Venda (PDV)
    ├── FormRelatorios.frm/.frx  # Formulário de Consultas e Relatórios de Vendas
    ├── ModProdutos.bas          # Regras de negócio de produtos
    ├── ModClientes.bas          # Regras de negócio de clientes
    ├── ModVendas.bas            # Motor de processamento de vendas, carrinho e cartões
    ├── ModRelatorios.bas        # Motor de filtros, cancelamento e exportação PDF
    └── ModFuncoes.bas           # Funções reutilizáveis e motor ADODB/SQL
```

---

## 🛠️ Requisitos e Como Utilizar

1. **Microsoft Excel** (versão 2016, 2019, 2021 ou Office 365 para Windows).
2. Ao abrir o arquivo `sistema de vendas.xlsm`, clique no botão **Habilitar Conteúdo / Habilitar Macros** na barra amarela superior.
3. Para acessar o Editor do VBA e visualizar os códigos, pressione <kbd>Alt</kbd> + <kbd>F11</kbd>.

---

## 👨‍💻 Autor e Créditos

- Desenvolvido por: **Vinícius Souza** ([@viniaasouza](https://github.com/viniaasouza))
- Referência didática: Curso gratuito de Sistema de Vendas em Excel VBA por **Edivam Cabral**.
