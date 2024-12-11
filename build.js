function importarEstoque() {
  const urlBase = 'https://api.sigecloud.com.br/request/Estoque/BuscarQuantidades';  
  const depositos = ['MATRIZ', 'ATACADO', 'DIRCEU', 'PARNAIBA', 'PICOS', 'AUTOATENDIMENTO'];  

  const token = PropertiesService.getScriptProperties().getProperty('Authorization-Token'); 
  const user = PropertiesService.getScriptProperties().getProperty('User');
  const app = PropertiesService.getScriptProperties().getProperty('App');

  const options = {
    'method': 'GET',
    'headers': {
      'Accept': 'application/json',
      'Authorization-Token': token,
      'User': user,
      'App': app,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  try {
  Logger.log("Importando dados dos estoques ...");
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoques");
  planilha.clearContents();

  const headers = ['Código', 'Estoque Matriz', 'Estoque Atacado', 'Estoque Dirceu', 'Estoque Parnaiba', 'Estoque Picos', 'Estoque Auto-Atendimento'];
  planilha.appendRow(headers);

  function buscarDadosDeposito(deposito) {
    const url = `${urlBase}?deposito=${deposito}`;
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    return data.EstoqueItens.map(item => {
      const linha = [item.ProdutoCodigo, '', '', '', '', '', ''];
      const index = depositos.indexOf(deposito) + 1;
      linha[index] = item.EstoqueAtual;
      return linha;
    });
  }

  const dadosEstoque = depositos.map(buscarDadosDeposito);

  const dadosCombinados = (() => {
    const dadosCombinados = {};
    dadosEstoque.flat().forEach(item => {
      const codigo = item[0];
      if (!dadosCombinados[codigo]) {
        dadosCombinados[codigo] = [codigo, '', '', '', '', '', ''];
      }
      for (let i = 1; i < item.length; i++) {
        if (item[i] !== '') {
          dadosCombinados[codigo][i] = item[i];
        }
      }
    });
    return Object.values(dadosCombinados);
  })();

  planilha.getRange(2, 1, dadosCombinados.length, headers.length).setValues(dadosCombinados);
  Logger.log("Importação concluida com sucesso :)");

  } catch (error) {
    Logger.log("Erro ao formatar preços: " + error);
  }
}
