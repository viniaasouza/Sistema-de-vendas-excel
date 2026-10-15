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

### 🔐 5. Multi-Usuário, Controle de Acesso e Auditoria (`FormLogin` & `ModSeguranca`)
- **Perfis de Usuário com PIN Rápido:**
  - `ADMIN`: Acesso irrestrito a configurações, catálogo de preços, relatórios gerenciais e cancelamento de vendas.
  - `RECEPCAO`: Acesso exclusivo para operação de vendas e reemissão de comprovantes de clientes.
- **Matriz de Credenciais Pré-Cadastradas (Aba `USUARIOS`):**
  - **Administrador:** Usuário `admin` | PIN `1234` | Perfil `ADMIN`
  - **Recepcionista 1:** Operadora `Ana Silva` (`recepcao1`) | PIN `1111` | Perfil `RECEPCAO`
  - **Recepcionista 2:** Operadora `Bruna Costa` (`recepcao2`) | PIN `2222` | Perfil `RECEPCAO`
- **Auditoria de Operador em Vendas:**
  - Registro automático do nome da recepcionista ativa na coluna `M` (`OPERADOR`) da aba `VENDAS`.
  - Exibição de badge do operador ativo no topo do PDV (`FormVendas`) com botão de **Troca Rápida de Operador**.
- **Trava de Permissões Críticas:**
  - **Catálogo de Produtos:** Bloqueio de edição de preços para recepcionistas (`TxtPreco.Locked = True`).
  - **Relatórios Gerenciais:** Mascaramento de faturamento total (`R$ ***`) e bloqueio de cancelamento de vendas sem PIN de Administrador.
- **Identidade Visual & Logotipo da Empresa:**
  - Suporte a logotipo institucional configurável na aba `CONFIG` (`LOGO_PATH`).
  - Inserção automática da logo no topo dos formulários e no cabeçalho do PDF do comprovante de venda.
  - Comprovante de venda com assinatura formal: *"Atendido por: [Nome da Recepcionista]"*.

### 🖥️ 6. Super Dashboard Gerencial Integrado & KPIs (`INÍCIO` - Aulas 22 a 25)
- **Tela Inicial Executiva Integrada:**
  - Abertura automática ao iniciar a pasta de trabalho com gridlines desativadas e visual moderno de aplicação comercial.
- **Barra Superior de Acesso Rápido:**
  - 🛒 **PDV (Vendas)** | 👥 **Clientes** | 📦 **Produtos** | 📑 **Relatórios** | 👤 **Trocar Operador**
- **Cards de Indicadores de Desempenho em Tempo Real (KPIs - Aulas 22 e 23):**
  - 💰 **Faturamento Total Líquido:** Soma das vendas concluídas.
  - 🛒 **Vendas Concluídas:** Total de pedidos e volume de itens vendidos.
  - 🎯 **Ticket Médio por Venda:** Valor médio gasto por cliente.
  - ⏳ **Previsão Líquida a Receber:** Faturamento deduzido das taxas de cartão.
- **Rankings de Inteligência Comercial (Aulas 24 e 25):**
  - 🏆 **TOP 5 Melhores Clientes:** Ranking automático dos 5 clientes com maior volume financeiro e total de compras.
  - 🥇 **TOP 5 Produtos Mais Vendidos:** Ranking automático dos 5 produtos líderes em unidades vendidas e faturamento gerado.
  - 💳 **Faturamento por Forma de Pagamento:** Resumo consolidado de entradas por Dinheiro, PIX, Cartão de Crédito, Débito e Boleto.
  - 🔄 **Botão de Atualização Instantânea:** Recálculo em tempo real via macro `ModDashboard.AtualizarDashboard`.
- **Blindagem de Segurança da Aba `USUARIOS`:**
  - Configuração permanente em modo **`xlSheetVeryHidden` (2)**: a aba é completamente invisível na barra de planilhas e **impossível de reexibir** pelo menu comum do Excel com o botão direito.
  - Acesso exclusivo aos logins e senhas através do botão **`🔐 GERENCIAR USUÁRIOS & PINS`**, que exige a confirmação com o PIN do Administrador (`1234`).
  - Auto-blindagem nos eventos `Workbook_Open` e `Workbook_BeforeClose`.

---

## 🗂️ Estrutura do Repositório

```text
├── sistema de vendas.xlsm       # Planilha oficial habilitada para macros (.xlsm)
├── .gitignore                   # Arquivos ignorados (bloqueios do Excel e backups)
├── README.md                    # Documentação do projeto
└── src/                         # Código fonte dos módulos e formulários VBA
    ├── FormLogin.frm/.frx       # Formulário de Autenticação e Troca de Operador
    ├── FormProdutos.frm/.frx    # Formulário e layout do Cadastro de Produtos
    ├── FormClientes.frm/.frx    # Formulário e layout do Cadastro de Clientes
    ├── FormVendas.frm/.frx      # Formulário e layout do Ponto de Venda (PDV)
    ├── FormRelatorios.frm/.frx  # Formulário de Consultas e Relatórios de Vendas
    ├── ModDashboard.bas         # Motor de KPIs, rankings TOP 5 e consolidação do painel
    ├── ModSeguranca.bas         # Controle de acesso, perfis, PINs e blindagem VeryHidden
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
