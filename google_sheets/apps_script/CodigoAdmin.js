/**
 * MÓDULO ADMINISTRATIVO & FINANCEIRO (DRE & FLUXO DE CAIXA) - GOOGLE APPS SCRIPT
 * Planilha: SISTEMA_GESTAO_FINANCEIRA
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 GESTÃO FINANCEIRA')
    .addItem('💸 Abrir Painel de Despesas & Contas a Pagar', 'abrirPainelDespesas')
    .addItem('🔄 Recalcular DRE e Fluxo de Caixa', 'atualizarDemonstrativos')
    .addSeparator()
    .addItem('📥 Importar Vendas da Planilha da Recepção', 'solicitarSincronizacaoVendas')
    .addToUi();
}

function abrirPainelDespesas() {
  var html = HtmlService.createHtmlOutputFromFile('DespesasModal')
    .setWidth(950)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '💸 Gestão de Despesas & Contas a Pagar');
}

function obterDadosDespesas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('DESPESAS');
  var lista = [];
  
  if (ws && ws.getLastRow() > 1) {
    var dados = ws.getRange(2, 1, ws.getLastRow() - 1, 15).getValues();
    for (var i = 0; i < dados.length; i++) {
      if (dados[i][0]) {
        var dtComp = dados[i][1] instanceof Date ? Utilities.formatDate(dados[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        var dtVenc = dados[i][2] instanceof Date ? Utilities.formatDate(dados[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        var dtPagto = dados[i][3] instanceof Date ? Utilities.formatDate(dados[i][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        
        lista.push({
          id: dados[i][0],
          competencia: dtComp,
          vencimento: dtVenc,
          pagamento: dtPagto,
          grupoDRE: dados[i][4],
          categoriaFluxo: dados[i][5],
          descricao: dados[i][6],
          valor: Number(dados[i][7]) || 0,
          formaPagto: dados[i][8],
          status: dados[i][9] || 'Pendente',
          operador: dados[i][10] || ''
        });
      }
    }
  }
  
  return {
    despesas: lista,
    operador: Session.getActiveUser().getEmail() || 'Administrador'
  };
}

function salvarDespesa(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName('DESPESAS');
    var ultID = 1;
    var linEditar = 0;
    
    if (ws.getLastRow() > 1) {
      var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (d.id && ids[i][0] == d.id) {
          linEditar = i + 2;
          break;
        }
        if (Number(ids[i][0]) >= ultID) ultID = Number(ids[i][0]) + 1;
      }
    }
    
    var idFinal = linEditar > 0 ? d.id : ultID;
    var dtComp = d.competencia ? new Date(d.competencia + 'T00:00:00') : new Date();
    var dtVenc = d.vencimento ? new Date(d.vencimento + 'T00:00:00') : dtComp;
    var dtPagto = (d.status === 'Paga' && d.pagamento) ? new Date(d.pagamento + 'T00:00:00') : (d.status === 'Paga' ? new Date() : '');
    
    var mComp = dtComp.getMonth() + 1;
    var aComp = dtComp.getFullYear();
    var mPagto = dtPagto instanceof Date ? (dtPagto.getMonth() + 1) : '';
    var aPagto = dtPagto instanceof Date ? dtPagto.getFullYear() : '';
    
    var linha = [
      idFinal,
      dtComp,
      dtVenc,
      dtPagto,
      d.grupoDRE,
      d.categoriaFluxo,
      d.descricao || d.categoriaFluxo,
      Number(d.valor) || 0,
      d.formaPagto || 'PIX',
      d.status,
      d.operador || 'Admin',
      mComp,
      aComp,
      mPagto,
      aPagto
    ];
    
    if (linEditar > 0) {
      ws.getRange(linEditar, 1, 1, 15).setValues([linha]);
    } else {
      ws.appendRow(linha);
    }
    
    atualizarDemonstrativos();
    return { sucesso: true, mensagem: 'Despesa salva e integrada aos relatórios!' };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro ao salvar despesa: ' + err.message };
  }
}

function darBaixaDespesa(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName('DESPESAS');
    var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
    var hoje = new Date();
    
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] == id) {
        var r = i + 2;
        ws.getRange(r, 4).setValue(hoje); // DATA_PAGAMENTO
        ws.getRange(r, 10).setValue('Paga'); // STATUS
        ws.getRange(r, 14).setValue(hoje.getMonth() + 1); // MES_PAGTO
        ws.getRange(r, 15).setValue(hoje.getFullYear()); // ANO_PAGTO
        break;
      }
    }
    
    atualizarDemonstrativos();
    return { sucesso: true, mensagem: 'Baixa realizada! Despesa registrada no Fluxo de Caixa.' };
  } catch (err) {
    return { sucesso: false, mensagem: err.message };
  }
}

function excluirDespesa(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName('DESPESAS');
    var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] == id) {
        ws.deleteRow(i + 2);
        break;
      }
    }
    atualizarDemonstrativos();
    return { sucesso: true, mensagem: 'Despesa excluída com sucesso!' };
  } catch (err) {
    return { sucesso: false, mensagem: err.message };
  }
}

function atualizarDemonstrativos() {
  SpreadsheetApp.flush();
}

function solicitarSincronizacaoVendas() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Sincronização em Nuvem',
    'Cole o ID ou Link da Planilha SISTEMA_VENDAS_PDV (Recepção):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var entrada = resp.getResponseText().trim();
  var idMatch = entrada.match(/\/d\/([a-zA-Z0-9-_]+)/);
  var planId = idMatch ? idMatch[1] : entrada;
  
  try {
    var ssRemota = SpreadsheetApp.openById(planId);
    var wsRemotaVendas = ssRemota.getSheetByName('REGISTRO_VENDAS');
    if (!wsRemotaVendas || wsRemotaVendas.getLastRow() <= 1) {
      ui.alert('Aviso', 'Nenhuma nova venda encontrada na planilha remota.', ui.ButtonSet.OK);
      return;
    }
    
    var dadosRemotos = wsRemotaVendas.getRange(2, 1, wsRemotaVendas.getLastRow() - 1, 13).getValues();
    var ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    var wsLocalVendas = ssLocal.getSheetByName('VENDAS');
    
    // Obter IDs já existentes
    var idsLocais = {};
    if (wsLocalVendas.getLastRow() > 1) {
      var arrIds = wsLocalVendas.getRange(2, 1, wsLocalVendas.getLastRow() - 1, 1).getValues();
      arrIds.forEach(function(r) { idsLocais[r[0]] = true; });
    }
    
    var novasVendas = 0;
    dadosRemotos.forEach(function(row) {
      if (row[0] && !idsLocais[row[0]]) {
        wsLocalVendas.appendRow(row);
        novasVendas++;
      }
    });
    
    atualizarDemonstrativos();
    ui.alert('Sincronização Concluída', novasVendas + ' novas vendas importadas para a base Master com sucesso!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erro de Conexão', 'Não foi possível acessar a planilha da recepção: ' + e.message, ui.ButtonSet.OK);
  }
}
