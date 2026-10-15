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
- Exclusão segura com diálogo de confirmação.

### 💳 3. Ponto de Venda / Emissão de Vendas (`FormVendas`)
- **Numeração Automática:** Geração incremental do número de venda mestre (`ID_VENDA`).
- **Carrinho de Compras Interativo (`LstItens`):**
  - Seleção dinâmica de clientes e produtos a partir das bases cadastradas.
  - Preenchimento instantâneo de preço unitário e cálculo em tempo real do subtotal por item.
  - Duplo clique no carrinho para carregar e editar quantidade/preço do item.
  - Exclusão individual de itens com recálculo automático.
- **Fechamento Financeiro:**
  - Soma automática do Total Bruto.
  - Aplicação de Desconto em valor monetário com atualização instantânea do Total a Pagar.
  - Seleção da forma de pagamento (Dinheiro, PIX, Cartão de Crédito, Cartão de Débito, Boleto).
- **Gravação Relacional (Mestre-Detalhe):**
  - Gravação da venda mestre na tabela `VENDAS`.
  - Gravação de todos os itens vinculados na tabela `ITENS_VENDA`.
- **Motor ADODB / SQL:** Integração com biblioteca ADODB para consultas de alta performance via SQL direto nas abas do Excel.

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
    ├── ModProdutos.bas          # Regras de negócio de produtos
    ├── ModClientes.bas          # Regras de negócio de clientes
    ├── ModVendas.bas            # Motor de processamento de vendas e carrinho
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
