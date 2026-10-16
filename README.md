# Sistema de Vendas Comercial no Excel VBA

Sistema completo e profissional de **Gestão Comercial, Ponto de Venda (PDV), Controle de Produtos e Serviços, Clientes, Relatórios Gerenciais com Exportação em PDF, Segurança Multi-Usuário com Blindagem de Acessos e Dashboard Interativo em Tempo Real**, desenvolvido no Microsoft Excel com programação VBA.

Inspirado e expandido a partir das aulas de Excel VBA do professor **Edivam Cabral** (incluindo as Aulas 22 a 27).

---

## 🚀 Funcionalidades do Sistema

### 📦 1. Gestão de Produtos, Serviços & Formação de Preço (`FormProdutos` & `ModProdutosServicos`)
- **Layout Profissional Ampliado (`476 × 462 pt`):** Campos distribuídos em linhas dedicadas (Tipo e Descrição; Preço Custo, Margem % e Preço Venda; Estoque Atual e Estoque Mínimo; Pesquisa e ListBox com 7 colunas perfeitamente dimensionadas).
- **Cadastro Completo:** Código ID automático, Tipo (Produto ou Serviço), descrição detalhada, Preço de Custo, Margem de Lucro Desejada (%), Preço Final de Venda, Estoque Atual e Estoque Mínimo.
- **Formação Inteligente de Preço (Custo + Margem):**
  - Cálculo automático do Preço Final de Venda ao preencher Custo e Margem desejada (%): `Preço Final = Custo × (1 + Margem/100)`.
  - Cálculo reverso da Margem (%) caso o usuário informe diretamente o Preço Final desejado.
- **Inteligência de Estoque por Tipo:** Para itens do tipo "Serviço", os campos de estoque são automaticamente desativados e preenchidos com "N/A".
- **Prevenção de Duplicidade:** Validação inteligente que impede cadastramento de produtos ou serviços com a mesma descrição.
- **Consulta em Tempo Real:** Campo de busca dinâmico com filtro incremental ao digitar (`LIKE %texto%`) e suporte a acentuação.
- **Restrição Total para Recepcionistas:** Perfis de Recepção operam exclusivamente em modo leitura (`Locked = True` em todos os campos de produto e estoque, botões de edição desabilitados e proteção por PIN de Administrador).

### 📋 2. Painel Executivo de Inventário & Estoque (`INVENTARIO`)
- **Acesso Rápido na Barra Superior:** Botão `📦 INVENTÁRIO` posicionado na barra de navegação superior da aba `INÍCIO` (`Top=128`), sem qualquer sobreposição com a tabela ou gráficos de Top 5.
- **Blindagem contra Edição por Recepcionistas:** Planilha protegida contra edição manual de células (`ProtectContents = True`, `UserInterfaceOnly = True`) e botão de gerenciamento protegido por validação de perfil/PIN de Administrador.
- **5 Cards de KPIs em Tempo Real:**
  - 📦 **Total de Itens em Estoque:** Volume físico consolidado de mercadorias.
  - 💵 **Custo Total do Estoque:** Capital total imobilizado a preço de aquisição.
  - 💰 **Valor de Venda do Estoque:** Receita bruta projetada de venda total.
  - 📈 **Lucro Bruto Projetado:** Margem nominal total projetada (`Valor Venda - Custo Total`).
  - ⚠️ **Alertas de Estoque:** Quantidade de itens zerados ou abaixo do estoque mínimo.
- **Tabela Gerencial Completa:** Relatório detalhado com ID, Tipo, Descrição, Custo, Margem, Preço de Venda, Estoque Atual, Mínimo, Status visual (`✅ NORMAL`, `⚠️ ESTOQUE BAIXO`, `❌ ZERADO`, `🛠️ SERVIÇO`), Custo Total, Venda Total e Lucro Projetado por item.
- **Baixa Automática no PDV:** O fechamento de qualquer venda de itens do tipo "Produto" no caixa (`ModVendas.GravarVenda`) abate automaticamente a quantidade vendida do estoque em tempo real.

### 👥 3. Cadastro e Gestão de Clientes com Bairro e Idade em Tempo Real (`FormClientes` & `ModClientes`)
- **Ficha Cadastral Expandida e Organizada:** ID automático, Nome Completo, CPF/CNPJ, Telefone/WhatsApp, E-mail, Bairro, Endereço Completo, Data de Nascimento e Idade em linhas independentes e sem sobreposição.
- **Cálculo Automático de Idade ao Digitar (`TxtDataNasc_Change`):** Ao digitar a Data de Nascimento, aplica máscara automática (`DD/MM/AAAA`) e calcula instantaneamente a idade em anos em tempo real, sem necessidade de sair do campo.
- **Prevenção de Duplicidade de Nome/CPF:** Verificação automática que bloqueia cadastro duplicado.
- **Busca Incremental:** Filtragem instantânea por nome, CPF/CNPJ ou bairro diretamente no `ListBox`.

### 🛒 4. Ponto de Venda - PDV / Frente de Caixa (`FormVendas` & `ModVendas`)
- **Seleção Dinâmica de Clientes e Produtos/Serviços:**
  - Carregamento de clientes e itens cadastrados no catálogo.
  - **Busca preditiva inteligente ao digitar** na lista suspensa do produto/serviço.
- **Carrinho de Compras Interativo (`LstItens`):**
  - Preenchimento instantâneo de preço unitário e cálculo em tempo real do subtotal por item.
  - Duplo clique no carrinho para carregar e editar quantidade/preço do item.
  - Exclusão individual de itens com recálculo automático.
- **Fechamento Financeiro com Inteligência de Cartão & Baixa de Estoque:**
  - Soma automática do Total Bruto e aplicação de desconto em valor (R$).
  - Seleção da forma de pagamento (Dinheiro, PIX, Cartão de Crédito, Cartão de Débito, Boleto).
  - **Parcelamento de 1x a 12x** para vendas no cartão de crédito com taxa configurável e previsão líquida real a receber.
  - **Baixa em tempo real de produtos vendidos** no catálogo e inventário.
- **Gravação Relacional (Mestre-Detalhe):**
  - Registro mestre na aba `VENDAS` com total bruto, líquido, parcelas, taxa e previsão a receber.
  - Registro dos itens na aba `ITENS_VENDA`.

### 📄 5. Relatórios Gerenciais e Emissão de PDF (`FormRelatorios` & `R_Vendas`)
- **Painel de Consultas Avançadas:**
  - Filtro por cliente, intervalo de datas, status (`Concluída` ou `Cancelada`) e forma de pagamento.
- **Cards de Resumo (KPIs em Tempo Real):**
  - Total de pedidos, faturamento bruto/líquido, retenção de taxas e previsão líquida a receber.
- **Exportação de Relatório Gerencial e Recibo Individual para PDF:**
  - Emissão com 1 clique de comprovantes A4 com dados da empresa, cliente, itens detalhados e forma de pagamento.

### 🔐 6. Multi-Usuário, Controle de Acesso e Blindagem (`FormLogin` & `ModSeguranca`)
- **Perfis de Usuário com PIN Rápido:** `ADMIN` (acesso irrestrito) e `RECEPCAO` (operações de caixa e comprovantes).
- **Auditoria de Operador em Vendas:** Registro automático da colaboradora ativa na coluna `M` da aba `VENDAS` e badge visual com botão de **Troca Rápida de Operador**.
- **Layout Desobstruído da Aba `USUARIOS`:**
  - Cabeçalho reposicionado para a linha 4 e dados a partir da linha 5.
  - Botão `🔒 BLOQUEAR E OCULTAR PLANILHA AGORA` fixado no canto superior direito (`Left=460, Top=8`), eliminando 100% da sobreposição com o cadastro das colaboradoras.
- **Blindagem de Segurança com Auto-Ocultação da Aba `USUARIOS`:**
  - Modo permanente **`xlSheetVeryHidden` (2)**: 100% invisível na barra de planilhas.
  - Auto-ocultação imediata ao trocar de operador, autenticar ou sair da aba.

### 📊 7. Dashboard Interativo com Filtro Inteligente de Operador (`INÍCIO`)
- **Barra Superior de Acesso Rápido:**
  - 🛒 **PDV (Vendas)** | 👥 **Clientes** | 📦 **Produtos & Serviços** | 📦 **Inventário** | 📈 **Relatórios** | 👤 **Trocar Operador**
- **Filtros Interativos Dinâmicos:**
  - `⭐ TODAS`: Consolida todo o histórico.
  - `📅 HOJE`: Métricas do dia corrente.
  - `📅 ESTE MÊS`: Métricas do mês vigente.
  - `👤 OPERADOR (Inteligente com Autocomplete)`:
    - Abre o novo formulário `FormFiltroOperador` com lista com autocomplete (`fmMatchEntryComplete`).
    - Botões rápidos com 1 clique para `⭐ TODOS`, `Vitória`, `Geovana` e `Admin`.
    - **Reconhecimento Inteligente:** Aceita nomes em maiúsculas, minúsculas e **sem acentos** (ex: digitar "vitoria" filtra perfeitamente as vendas de "Vitória").
  - **Banner de Status do Filtro:** Exibe em tempo real o período ativo, operador selecionado e total de pedidos correspondentes.
- **4 Cards de KPIs Executivos & Rankings Comerciais:**
  - Faturamento Total Líquido, Volume de Vendas e Itens, Ticket Médio, Previsão Líquida a Receber, TOP 5 Clientes, TOP 5 Produtos/Serviços e Faturamento por Meio de Pagamento.
- **Gráficos Visuais Nativos Interativos:**
  - 🍩 Gráfico de Rosca (Meios de Pagamento) e 📊 Gráfico de Barras Horizontais (TOP 5 Itens Mais Vendidos), atualizados instantaneamente pelos filtros.

---

## 📁 Estrutura do Repositório

```text
├── sistema de vendas.xlsm         # Planilha oficial habilitada para macros (.xlsm)
├── .gitignore                     # Arquivos ignorados (bloqueios do Excel e backups)
├── README.md                      # Documentação detalhada do projeto
└── src/                           # Código fonte dos módulos e formulários VBA
    ├── EstaPastaDeTrabalho.cls    # Eventos de abertura, fechamento e auto-blindagem
    ├── FormLogin.frm/.frx         # Formulário de Autenticação e Troca de Operador
    ├── FormProdutos.frm/.frx      # Formulário de Produtos, Serviços, Custos e Margem
    ├── FormClientes.frm/.frx      # Formulário de Clientes (com Bairro, Nasc. e Idade)
    ├── FormVendas.frm/.frx        # Formulário do Ponto de Venda (PDV com baixa de estoque)
    ├── FormRelatorios.frm/.frx    # Formulário de Consultas e Relatórios de Vendas
    ├── FormFiltroOperador.frm/.frx# Formulário de Filtro de Operador com Autocomplete
    ├── ModDashboard.bas           # Motor do Dashboard, filtros inteligentes e gráficos
    ├── ModSeguranca.bas           # Controle de acesso, perfis, PINs e blindagem VeryHidden
    ├── ModProdutosServicos.bas    # Regras de produtos, serviços, markup e inventário
    ├── ModClientes.bas            # Regras de clientes e cálculo de idade
    ├── ModVendas.bas              # Motor de processamento de vendas e baixa de estoque
    ├── ModRelatorios.bas          # Motor de filtros, cancelamento e exportação PDF
    └── ModFuncoes.bas             # Funções reutilizáveis e navegação de planilhas
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
