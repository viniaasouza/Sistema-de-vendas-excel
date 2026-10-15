# Sistema de Vendas Comercial no Excel VBA

Sistema completo e profissional de **Gestão Comercial, Ponto de Venda (PDV), Controle de Produtos e Serviços, Clientes, Relatórios Gerenciais com Exportação em PDF, Segurança Multi-Usuário com Blindagem de Acessos e Dashboard Interativo em Tempo Real**, desenvolvido no Microsoft Excel com programação VBA.

Inspirado e expandido a partir das aulas de Excel VBA do professor **Edivam Cabral** (incluindo as Aulas 22 a 27).

---

## 🚀 Funcionalidades do Sistema

### 📦 1. Gestão de Produtos e Serviços (`FormProdutos` & `ModProdutosServicos`)
- **Cadastro Completo:** Código ID automático, descrição do produto ou serviço e preço de venda unitário.
- **Prevenção de Duplicidade:** Validação inteligente que impede o cadastramento de produtos ou serviços com a mesma descrição.
- **Consulta em Tempo Real com SQL:** Campo de busca dinâmico com filtro incremental ao digitar (`LIKE %texto%`) e suporte a acentuação.
- **Edição e Exclusão Segura:** Seleção do item na lista para alteração de dados ou exclusão com diálogo de confirmação.
- **Proteção de Catálogo:** Perfis de Recepção operam em modo consulta com preço bloqueado para edição (`TxtPreco.Locked = True`), exigindo autorização de Administrador para alterações.

### 👥 2. Cadastro e Gestão de Clientes (`FormClientes` & `ModClientes`)
- **Ficha Cadastral:** Registro de ID automático, Nome Completo, CPF/CNPJ, Telefone/WhatsApp, E-mail e Endereço.
- **Prevenção de Duplicidade de CPF/CNPJ:** Verificação automática que bloqueia cadastro duplicado do mesmo documento.
- **Busca Incremental:** Filtragem instantânea por nome ou documento diretamente no `ListBox`.
- **Edição e Exclusão:** Carregamento nos campos para alteração ou exclusão com confirmação segura.

### 🛒 3. Ponto de Venda - PDV / Frente de Caixa (`FormVendas` & `ModVendas`)
- **Seleção Dinâmica de Clientes e Produtos/Serviços:**
  - Carregamento de clientes e itens cadastrados no catálogo.
  - **Busca preditiva inteligente ao digitar** na lista suspensa do produto/serviço.
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

### 📄 4. Relatórios Gerenciais e Emissão de PDF (`FormRelatorios` & `R_Vendas`)
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
  - Emissão individual com 1 clique (ou duplo clique na venda) contendo cabeçalho da empresa, dados do cliente, lista completa de produtos e serviços, totais e forma de pagamento.
- **Dados da Empresa Configuráveis (`CONFIG`):**
  - Gerenciamento fácil de Razão Social, Nome Fantasia, CNPJ, Telefone, E-mail, Endereço e Mensagem de Rodapé refletidos automaticamente nos relatórios e recibos.

### 🔐 5. Multi-Usuário, Controle de Acesso e Auditoria (`FormLogin` & `ModSeguranca`)
- **Perfis de Usuário com PIN Rápido:**
  - `ADMIN`: Acesso irrestrito a configurações, catálogo de preços, relatórios gerenciais e gerenciamento de usuários.
  - `RECEPCAO`: Acesso exclusivo para operação de vendas e reemissão de comprovantes de clientes.
- **Auditoria de Operador em Vendas:**
  - Registro automático do nome da recepcionista ativa na coluna `M` (`OPERADOR`) da aba `VENDAS`.
  - Exibição de badge do operador ativo no topo do PDV (`FormVendas`) com botão de **Troca Rápida de Operador**.
- **Trava de Permissões Críticas:**
  - **Catálogo de Produtos e Serviços:** Bloqueio de edição de preços para recepcionistas (`TxtPreco.Locked = True`).
  - **Relatórios Gerenciais:** Mascaramento de faturamento total (`R$ ***`) e bloqueio de cancelamento de vendas sem PIN de Administrador.
- **Blindagem de Segurança com Auto-Ocultação da Aba `USUARIOS`:**
  - Configuração permanente em modo **`xlSheetVeryHidden` (2)**: a aba é completamente invisível na barra de planilhas e **impossível de reexibir** pelo menu comum do Excel com o botão direito.
  - **Auto-Ocultação Imediata:**
    1. Ao trocar de operador ou abrir a tela de login (`TrocarOperador`).
    2. Ao autenticar qualquer usuário (`AutenticarUsuario`).
    3. Ao clicar fora ou desativar a aba de usuários (`Worksheet_Deactivate`).
    4. Ao ativar qualquer outra aba no Excel (`Workbook_SheetActivate`).
    5. Botão dedicado no topo da aba `USUARIOS`: `🔒 BLOQUEAR E OCULTAR PLANILHA AGORA`.
  - Acesso exclusivo aos logins e senhas através do botão **`🔐 GERENCIAR USUÁRIOS & PINS`**, exigindo confirmação com PIN Master do Administrador (`1234`).

### 📊 6. Dashboard Interativo em Tempo Real (`INÍCIO` - Aulas 22 a 27)
- **Barra Superior de Acesso Rápido:**
  - 🛒 **PDV (Vendas)** | 👥 **Clientes** | 📦 **Produtos & Serviços** | 📈 **Relatórios** | 👤 **Trocar Operador**
- **Filtros Interativos Dinâmicos (Aulas 26 e 27):**
  - `⭐ TODAS`: Visualiza todos os períodos consolidados.
  - `📅 HOJE`: Filtra vendas e faturamento do dia atual.
  - `📅 ESTE MÊS`: Filtra os números do mês e ano vigentes.
  - `👤 OPERADOR`: Filtra vendas exclusivamente pelo operador selecionado (ex: Vitória, Geovana, Administrador).
  - **Banner de Status do Filtro:** Mostra em destaque o período ativo, operador selecionado e o total de pedidos filtrados.
- **4 Cards de KPIs em Tempo Real (Aulas 22 e 23):**
  - 💰 **Faturamento Total Líquido:** Soma das vendas concluídas no filtro ativo.
  - 🛒 **Vendas Concluídas:** Total de pedidos e volume de itens vendidos.
  - 📈 **Ticket Médio:** Valor médio gasto por cliente.
  - 💳 **Líquido a Receber:** Faturamento deduzido das taxas de cartão.
- **Rankings Comerciais Dinâmicos (Aulas 24 e 25):**
  - 🏆 **TOP 5 Melhores Clientes:** Ranking dos 5 clientes com maior volume financeiro e quantidade de compras.
  - 🏆 **TOP 5 Produtos & Serviços Mais Vendidos:** Ranking dos 5 itens líderes em unidades e receita gerada.
  - 💳 **Faturamento por Forma de Pagamento:** Resumo consolidado de entradas por Dinheiro, PIX, Cartão de Crédito, Débito e Boleto.
- **Gráficos Visuais Nativos Interativos (Aulas 26 e 27):**
  - 🍩 **Gráfico de Rosca (Meios de Pagamento):** Proporção visual percentual e em R$ das formas de pagamento.
  - 📊 **Gráfico de Barras Horizontais (TOP 5 Produtos e Serviços):** Comparativo visual do faturamento gerado pelos produtos e serviços líderes.
  - **Atualização Instantânea:** Conexão direta com os filtros, atualizando métricas e gráficos sem recarregar a planilha.

---

## 📁 Estrutura do Repositório

```text
├── sistema de vendas.xlsm       # Planilha oficial habilitada para macros (.xlsm)
├── .gitignore                   # Arquivos ignorados (bloqueios do Excel e backups)
├── README.md                    # Documentação do projeto
└── src/                         # Código fonte dos módulos e formulários VBA
    ├── EstaPastaDeTrabalho.cls  # Eventos de abertura, fechamento e auto-blindagem
    ├── FormLogin.frm/.frx       # Formulário de Autenticação e Troca de Operador
    ├── FormProdutos.frm/.frx    # Formulário do Cadastro de Produtos e Serviços
    ├── FormClientes.frm/.frx    # Formulário do Cadastro de Clientes
    ├── FormVendas.frm/.frx      # Formulário do Ponto de Venda (PDV)
    ├── FormRelatorios.frm/.frx  # Formulário de Consultas e Relatórios de Vendas
    ├── ModDashboard.bas         # Motor do Dashboard interativo, filtros, KPIs e gráficos
    ├── ModSeguranca.bas         # Controle de acesso, perfis, PINs e blindagem VeryHidden
    ├── ModProdutosServicos.bas  # Regras de negócio de produtos e serviços
    ├── ModClientes.bas          # Regras de negócio de clientes
    ├── ModVendas.bas            # Motor de processamento de vendas, carrinho e cartões
    ├── ModRelatorios.bas        # Motor de filtros, cancelamento e exportação PDF
    └── ModFuncoes.bas           # Funções reutilizáveis e motor ADODB/SQL
```

---

## 💻 Requisitos e Como Utilizar

1. **Microsoft Excel** (versão 2016, 2019, 2021 ou Office 365 para Windows).
2. Ao abrir o arquivo `sistema de vendas.xlsm`, clique no botão **Habilitar Conteúdo / Habilitar Macros** na barra amarela superior.
3. Para acessar o Editor do VBA e visualizar os códigos, pressione <kbd>Alt</kbd> + <kbd>F11</kbd>.

---

## 👨‍💻 Autor e Créditos

- Desenvolvido por: **Vinícius Souza** ([@viniaasouza](https://github.com/viniaasouza))
- Referência didática: Curso gratuito de Sistema de Vendas em Excel VBA por **Edivam Cabral**.
