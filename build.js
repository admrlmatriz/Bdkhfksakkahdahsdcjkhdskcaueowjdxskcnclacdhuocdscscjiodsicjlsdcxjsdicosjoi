// Função para importar os dados de vendas
function importarVendas() {
  const urlBase = 'https://api.sigecloud.com.br/request/Pedidos/Pesquisar';
  const dataInicial = '2024-11-01T00:00:00-03:00'; // Data inicial do intervalo
  const dataFinal = '2024-12-01T00:00:00-03:00'; // Data final do intervalo
  const filtrarPor = '3'; // Filtro por "Data de faturamento do pedido"

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
    Logger.log("Importando dados de vendas ...");

    const planilha = SpreadsheetApp.getActiveSpreadsheet();

    // Obtém ou cria a aba "Vendas"
    const abaVendas = planilha.getSheetByName("Vendas") || planilha.insertSheet("Vendas");
    abaVendas.clearContents();

    const headers = ['Vendedor', 'Número do Pedido', 'Forma de Pagamento', 'Valor Total', 'Loja'];
    abaVendas.appendRow(headers);

    const url = `${urlBase}?dataInicial=${encodeURIComponent(dataInicial)}&dataFinal=${encodeURIComponent(dataFinal)}&filtrarPor=${filtrarPor}`;
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(`Erro na requisição: ${response.getContentText()}`);
    }

    const vendas = JSON.parse(response.getContentText());
    const linhas = vendas.map(venda => [
      venda.Vendedor || '',
      venda.Codigo || '',
      venda.FormaPagamento || '',
      venda.ValorFinal || '',
      venda.Empresa || ''
    ]);

    if (linhas.length > 0) {
      abaVendas.getRange(2, 1, linhas.length, headers.length).setValues(linhas);
    }

    // Adicionar data e hora atual na planilha
    const dataHora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'HH:mm - dd/MM/yy');
    abaVendas.getRange('A1').setNote(`Tabela atualizada em ${dataHora}`);

    Logger.log("Importação de vendas concluída com sucesso :)");

  } catch (error) {
    Logger.log("Erro ao importar vendas: " + error);
  }
}
