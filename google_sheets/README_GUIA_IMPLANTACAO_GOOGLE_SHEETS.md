# 🚀 Guia de Implantação: Sistema de Vendas, PDV e Finanças no Google Sheets

Este pacote contém todos os arquivos necessários para rodar o sistema completo no **Google Sheets**, 100% online, sem custo de licença Microsoft e com acesso simultâneo para **Recepção** e **Administração**.

---

## 📂 Arquivos Gerados na Pasta `google_sheets/`

1. **`SISTEMA_GESTAO_FINANCEIRA.xlsx`**: Planilha Master da Administração (DRE, Fluxo de Caixa, Despesas, Estoque Real com Custo, Vendas e Usuários).
2. **`SISTEMA_VENDAS_PDV.xlsx`**: Planilha da Recepção (Terminal de Caixa PDV, Tabela de Preços sem custo, Cadastro de Clientes).
3. **Pasta `apps_script/`**:
   * `CodigoPDV.js` e `PdvSidebar.html`: Motor e interface do PDV da Recepção.
   * `CodigoAdmin.js` e `DespesasModal.html`: Motor e interface financeira da Administração.

---

## 🛠️ Passo a Passo de Implantação no Google Drive

### Passo 1: Subir as Planilhas no Google Drive
1. Acesse o seu [Google Drive](https://drive.google.com).
2. Crie uma pasta chamada **`SISTEMA DE VENDAS`**.
3. Arraste os dois arquivos para dentro da pasta:
   * `SISTEMA_GESTAO_FINANCEIRA.xlsx`
   * `SISTEMA_VENDAS_PDV.xlsx`
4. Dê dois cliques em cada um e clique no topo em **Abrir com Planilhas Google**.
5. No menu superior da planilha, clique em **Arquivo $\rightarrow$ Salvar como Planilha Google**. *(Isso converte o arquivo para o formato nativo do Google Sheets)*.

---

### Passo 2: Configurar o PDV na Planilha da Recepção (`SISTEMA_VENDAS_PDV`)
1. Abra a planilha **`SISTEMA_VENDAS_PDV`** no Google Sheets.
2. No menu superior, clique em **Extensões $\rightarrow$ Apps Script**.
3. No arquivo `Código.gs`, apague o conteúdo existente e cole o código do arquivo:
   * [`google_sheets/apps_script/CodigoPDV.js`](file:///C:/Users/vinia/OneDrive/Área%20de%20Trabalho/SISTEMA%20DE%20VENDAS/google_sheets/apps_script/CodigoPDV.js)
4. No menu esquerdo do Apps Script, clique no botão **`+`** (Adicionar arquivo) $\rightarrow$ **HTML**.
5. Nomeie o arquivo exatamente como **`PdvSidebar`** (sem a extensão `.html`).
6. Cole todo o conteúdo do arquivo:
   * [`google_sheets/apps_script/PdvSidebar.html`](file:///C:/Users/vinia/OneDrive/Área%20de%20Trabalho/SISTEMA%20DE%20VENDAS/google_sheets/apps_script/PdvSidebar.html)
7. Clique no ícone de **Salvar** (disquete) no topo.
8. Volte para a planilha e recarregue a página (F5). Surgirá no menu superior: **`🚀 SISTEMA PDV`**!

---

### Passo 3: Configurar o Painel Financeiro na Planilha Master (`SISTEMA_GESTAO_FINANCEIRA`)
1. Abra a planilha **`SISTEMA_GESTAO_FINANCEIRA`** no Google Sheets.
2. No menu superior, clique em **Extensões $\rightarrow$ Apps Script**.
3. No arquivo `Código.gs`, apague o conteúdo e cole o código do arquivo:
   * [`google_sheets/apps_script/CodigoAdmin.js`](file:///C:/Users/vinia/OneDrive/Área%20de%20Trabalho/SISTEMA%20DE%20VENDAS/google_sheets/apps_script/CodigoAdmin.js)
4. Clique no **`+`** (Adicionar arquivo) $\rightarrow$ **HTML**.
5. Nomeie como **`DespesasModal`**.
6. Cole o conteúdo de:
   * [`google_sheets/apps_script/DespesasModal.html`](file:///C:/Users/vinia/OneDrive/Área%20de%20Trabalho/SISTEMA%20DE%20VENDAS/google_sheets/apps_script/DespesasModal.html)
7. Salve e recarregue a planilha. Surgirá no menu superior: **`📊 GESTÃO FINANCEIRA`**!

---

## 🔒 Como Funciona a Segurança

* **Recepção**: Compartilhe a planilha `SISTEMA_VENDAS_PDV` com o e-mail da recepcionista. Ela terá acesso apenas ao PDV, produtos com preço de venda e clientes.
* **Administração**: Mantenha a planilha `SISTEMA_GESTAO_FINANCEIRA` compartilhada **apenas com o seu e-mail**. A recepção nunca terá acesso a custos, margens, lucros nem à DRE.
* **Sincronização**: Na planilha Master, basta clicar em **`📊 GESTÃO FINANCEIRA $\rightarrow$ 📥 Importar Vendas da Planilha da Recepção`** para puxar todas as vendas concluídas instantaneamente.
