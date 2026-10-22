/**
 * SISTEMA DE PONTO DE VENDA (PDV) - GOOGLE APPS SCRIPT
 * Planilha: SISTEMA_VENDAS_PDV
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 SISTEMA PDV')
    .addItem('🛒 Abrir Caixa (Barra Lateral)', 'abrirSidebarPDV')
    .addItem('🖥️ Abrir Caixa (Janela Modal)', 'abrirModalPDV')
    .addSeparator()
    .addItem('👥 Cadastrar Novo Cliente', 'abrirCadastroCliente')
    .addItem('🔄 Atualizar Estoque / Preços', 'atualizarTabelaPrecos')
    .addToUi();
}

function abrirSidebarPDV() {
  var html = HtmlService.createHtmlOutputFromFile('PdvSidebar')
    .setTitle('Caixa PDV Online')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirModalPDV() {
  var html = HtmlService.createHtmlOutputFromFile('PdvSidebar')
    .setWidth(850)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '🛒 Ponto de Venda (PDV) - Terminal de Caixa');
}

function abrirCadastroCliente() {
  var ui = SpreadsheetApp.getUi();
  var respNome = ui.prompt('Novo Cliente', 'Nome Completo do Cliente:', ui.ButtonSet.OK_CANCEL);
  if (respNome.getSelectedButton() !== ui.Button.OK || !respNome.getResponseText().trim()) return;
  
  var respTel = ui.prompt('Novo Cliente', 'Telefone / WhatsApp:', ui.ButtonSet.OK_CANCEL);
  var tel = respTel.getSelectedButton() === ui.Button.OK ? respTel.getResponseText().trim() : '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('CAD_CLIENTES');
  var ultLinha = ws.getLastRow() + 1;
  var proxID = ultLinha - 1;
  
  ws.appendRow([proxID, respNome.getResponseText().trim(), '', tel, '', '', '', '', '']);
  ui.alert('Sucesso', 'Cliente cadastrado com sucesso!', ui.ButtonSet.OK);
}

function obterDadosIniciais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsProd = ss.getSheetByName('TABELA_PRECOS');
  var wsCli = ss.getSheetByName('CAD_CLIENTES');
  
  // Obter Produtos (ID, Tipo, Descricao, Preco, Estoque)
  var prods = [];
  if (wsProd && wsProd.getLastRow() > 1) {
    var dadosP = wsProd.getRange(2, 1, wsProd.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < dadosP.length; i++) {
      if (dadosP[i][0]) {
        prods.push({
          id: dadosP[i][0],
          tipo: dadosP[i][1] || 'Produto',
          descricao: dadosP[i][2],
          preco: Number(dadosP[i][3]) || 0,
          estoque: Number(dadosP[i][4]) || 0
        });
      }
    }
  }
  
  // Obter Clientes
  var clientes = [];
  if (wsCli && wsCli.getLastRow() > 1) {
    var dadosC = wsCli.getRange(2, 1, wsCli.getLastRow() - 1, 4).getValues();
    for (var j = 0; j < dadosC.length; j++) {
      if (dadosC[j][0]) {
        clientes.push({
          id: dadosC[j][0],
          nome: dadosC[j][1],
          telefone: dadosC[j][3] || ''
        });
      }
    }
  }
  
  return {
    produtos: prods,
    clientes: clientes,
    operador: Session.getActiveUser().getEmail() || 'Operador Recepção'
  };
}

function finalizarVenda(venda) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var wsVendas = ss.getSheetByName('REGISTRO_VENDAS');
    var wsProd = ss.getSheetByName('TABELA_PRECOS');
    
    var ultID = 1;
    if (wsVendas.getLastRow() > 1) {
      var ultVal = wsVendas.getRange(wsVendas.getLastRow(), 1).getValue();
      ultID = (Number(ultVal) || 0) + 1;
    }
    
    var hoje = new Date();
    var taxaPct = Number(venda.taxaCartaoPct) || 0;
    var totLiq = Number(venda.totalLiquido) || 0;
    var taxaRS = Number((totLiq * taxaPct).toFixed(2));
    var liqReceber = Number((totLiq - taxaRS).toFixed(2));
    
    // 1. Gravar Venda
    wsVendas.appendRow([
      ultID,
      hoje,
      venda.idCliente || 0,
      venda.nomeCliente || 'Consumidor Final',
      Number(venda.totalBruto) || 0,
      Number(venda.desconto) || 0,
      totLiq,
      venda.formaPagto,
      'Concluida',
      venda.parcelas || '1x',
      taxaPct,
      liqReceber,
      venda.operador || 'Recepção'
    ]);
    
    // 2. Dar baixa no Estoque de cada item
    if (wsProd && wsProd.getLastRow() > 1) {
      var dadosP = wsProd.getRange(2, 1, wsProd.getLastRow() - 1, 5).getValues();
      for (var k = 0; k < venda.itens.length; k++) {
        var it = venda.itens[k];
        for (var p = 0; p < dadosP.length; p++) {
          if (dadosP[p][0] == it.id) {
            var estAtual = Number(dadosP[p][4]) || 0;
            var novoEst = estAtual - (Number(it.qtd) || 0);
            wsProd.getRange(p + 2, 5).setValue(novoEst);
            break;
          }
        }
      }
    }
    
    return {
      sucesso: true,
      idVenda: ultID,
      mensagem: 'Venda Nº ' + ultID + ' registrada com sucesso!'
    };
  } catch (err) {
    return {
      sucesso: false,
      mensagem: 'Erro ao registrar venda: ' + err.message
    };
  }
}
