// Função para importar dados dos produtos como preço, fornecedor, CFOP, etc..
function importarProdutos() {
  const urlBase = 'https://api.sigecloud.com.br/request/Produtos/GetAll';
  const pageSize = 500; 
  let skip = 0; 
  let hasMore = true;

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
  
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Produtos");
  planilha.clearContents(); // Limpa dados anteriores
  
  // Se for a primeira execução, adiciona os cabeçalhos
  if (planilha.getLastRow() === 0) {
    const headers = ['Codigo', 'PrecoVenda', 'Ecommerce', 'Parnaiba', 'Picos', 'Categorias', 'Fornecedor', 'CFOP'];
    planilha.appendRow(headers);
  }

  // Define toda a primeira coluna como texto (isso evita que zeros à esquerda sejam removidos)
  planilha.getRange(1, 1, planilha.getMaxRows()).setNumberFormat('@');
  
  while (hasMore) {
    
    const url = `${urlBase}?pageSize=${pageSize}&skip=${skip}`;
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const dados = JSON.parse(response.getContentText());
        
        if (Array.isArray(dados) && dados.length > 0) {
          const rows = [];
          
          dados.forEach(produto => {
            // Filtra apenas produtos com o gênero "00 – Mercadoria para Revenda"
            if (produto.Genero === "00 – Mercadoria para Revenda") {
              const precoEcommerce = produto.PrecosTabelas.find(t => t.Tabela === 'ECOMMERCE')?.PrecoVenda || '';
              const precoParnaibaPadrao = produto.PrecosTabelas.find(t => t.Tabela === 'PARNAÍBA PADRÃO')?.PrecoVenda || '';
              const precoPicosPadrao = produto.PrecosTabelas.find(t => t.Tabela === 'PICOS PADRÃO')?.PrecoVenda || '';

              // Junta as categorias em uma única string separada por vírgulas
              const categorias = produto.Categorias ? produto.Categorias.join(', ') : 'N/A';

              
              rows.push([
                produto.Codigo || 'N/A',
                produto.PrecoVenda || 0,
                precoEcommerce,
                precoParnaibaPadrao,
                precoPicosPadrao,
                categorias,
                produto.Fornecedor,
                produto.CFOPPadrao
              ]);
            }
          });
          
          // Insere as linhas de dados na planilha
          if (rows.length > 0) {
            planilha.getRange(planilha.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
          }
          
          // Atualiza o valor de skip para a próxima requisição
          skip += pageSize;
          
          // Verifica se o número de produtos recebidos é menor que o tamanho da página (fim dos dados)
          if (dados.length < pageSize) {
            hasMore = false;
          }
          
        } else {
          // Não há mais dados a serem processados
          hasMore = false;
          Logger.log('Todos os dados foram importados com sucesso!');
        }
        
      } else {
        hasMore = false;
        Logger.log('Erro na requisição: Código de resposta ' + responseCode);
      }
      
    } catch (error) {
      hasMore = false;
      Logger.log('Ocorreu um erro: ' + error.message);
    }
  }

  // Adicionar data e hora atual na planilha
  const dataHora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'HH:mm - dd/MM/yy');
  planilha.getRange('A1').setNote(`Tabela atualizada em ${dataHora}`); // Adiciona comentário em 'A1'


  Logger.log('Script finalizado com Sucesso :)');
}


//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

// Função para importar os estoques de todas as lojas
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

  const headers = ['Código', 'Matriz', 'Atacado', 'Dirceu', 'Parnaiba', 'Picos', 'AutoAtendimento'];
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

  // Adicionar data e hora atual na planilha
  const dataHora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'HH:mm - dd/MM/yy');
  planilha.getRange('A1').setNote(`Tabela atualizada em ${dataHora}`); // Adiciona comentário em 'A1'

  Logger.log("Importação concluida com sucesso :)");

  } catch (error) {
    Logger.log("Erro ao importar estoques: " + error);
  }
}
