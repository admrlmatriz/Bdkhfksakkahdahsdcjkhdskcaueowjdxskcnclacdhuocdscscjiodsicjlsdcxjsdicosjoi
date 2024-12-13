// Função para formatação e ajustes de preços
function formatarPrecos() {
  try {
  Logger.log("Iniciando a Formatação de preços ...");
  const agora = new Date();
  const diaSemana = agora.getDay(), horaAtual = agora.getHours();
  const dataHora = Utilities.formatDate(new Date(), 'America/Fortaleza', 'HH:mm - dd/MM/yy');   // Obtém a data e hora atuais
  const emailDestino = PropertiesService.getScriptProperties().getProperty('emailDestino');

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  // Atualiza o título da planilha com data e hora atual
  var novoNome = 'LAB - Controle de Tabelas - ' + dataHora;
  planilha.setName(novoNome);

  // Verificar se a aba 'Principal' existe
  const abaPrincipal = planilha.getSheetByName("Principal");
  var lastRow = abaPrincipal.getLastRow();
  if (!abaPrincipal) {
    throw new Error("Aba 'Principal' não encontrada. Verifique se a aba existe e se o nome está correto.");
  }

  // Obtém ou cria a aba "Atualizados"
  const abaAtualizados = planilha.getSheetByName("Atualizados") || planilha.insertSheet("Atualizados");
  abaAtualizados.clearContents();

  // Define o cabeçalho na aba "Atualizados"
  var rangeCabecalho = abaAtualizados.getRange(1, 1, 1, 2);
  rangeCabecalho.setValues([["Código", "Preço de Venda"]]);
  var rangeA = abaAtualizados.getRange("A:A").setNumberFormat("@");

  // Obtém todos os dados da aba 'Principal' de uma vez em um array 2D
  const data = abaPrincipal.getRange(2, 1, lastRow - 1, 6).getValues();

  // Palavras-chave dos produtos que não aplicam o +2% em Parnaíba
  const palavrasChaves = ["COBRES", "CABOS", "GÁS REFRIGERANTE", "ELETRODOMÉSTICOS", "SPLITS"];
  const tolerancia = 0.05;

  // Códigos para filtrar
  const codigosFiltrados = ["92219321", "1780014587", "MLB3593567469"];

  // Processa os dados no array
  data.forEach((row, index) => {
    const [codigo, precoVenda, precoEcommerce, precoParnaiba, precoPicos, categoria] = row;

    // Faixa de tolerância para Ecommerce entre -4,5% e -5,5% do preço de venda
    var precoEcommerceInferior = precoVenda * 0.955;
    var precoEcommerceSuperior = precoVenda * 0.945;

    // Faixa de tolerância para Parnaíba entre 1,5% e 2,5% do preço de venda
    var precoParnaibaInferior = precoVenda * 1.015;
    var precoParnaibaSuperior = precoVenda * 1.025;

    // Faixa de tolerância para Picos entre 1,5% e 2,5% do preço de venda
    var precoPicosInferior = precoVenda * 1.015;
    var precoPicosSuperior = precoVenda * 1.025;

    // Formatação condicional para Ecommerce (-5%)
    if (precoEcommerce >= precoEcommerceSuperior - tolerancia && precoEcommerce <= precoEcommerceInferior + tolerancia) {
      abaPrincipal.getRange(index + 2, 3).setBackground("#57BB8A"); // Verde
    } else {
      abaPrincipal.getRange(index + 2, 3).setBackground("#E06666"); // Vermelho
      // Adiciona o produto à aba "Atualizados" se a condição do Ecommerce estiver fora da faixa (Indica que foi atualizado)
      if (!codigosFiltrados.includes(codigo.toString().trim())) {
        // Força a coluna código como texto usando aspas
        abaAtualizados.appendRow([`="${codigo}"`, precoVenda]);
      }
    }

    // Formatação condicional para Parnaíba
    if (palavrasChaves.some(palavra => categoria.includes(palavra))) {
      // Caso em que a categoria contém palavras-chave e não aplica o ajuste de +2%
      if (precoParnaiba != precoVenda) { 
       abaPrincipal.getRange(index + 2, 4).setBackground("#E06666"); // Vermelho
      } else {
        abaPrincipal.getRange(index + 2, 4).setBackground("#57BB8A"); // Verde
      }

    } else {
      // Caso sem palavras-chave - verifica tolerância entre 1,5% e 2,5%
      if (precoParnaiba >= precoParnaibaInferior && precoParnaiba <= precoParnaibaSuperior) {
        abaPrincipal.getRange(index + 2, 4).setBackground("#57BB8A"); // Verde
      } else {
        abaPrincipal.getRange(index + 2, 4).setBackground("#E06666"); // Vermelho
      }
    }

    // Formatação condicional para Picos
    if (palavrasChaves.some(palavra => categoria.includes(palavra))) {
      // Caso em que a categoria contém palavras-chave e não aplica o ajuste de +2%
      if (precoPicos != precoVenda) { 
       abaPrincipal.getRange(index + 2, 5).setBackground("#E06666"); // Vermelho
      } else {
        abaPrincipal.getRange(index + 2, 5).setBackground("#57BB8A"); // Verde
      }

    } else {
      // Caso sem palavras-chave - verifica tolerância entre 1,5% e 2,5%
      if (precoPicos >= precoPicosInferior && precoPicos <= precoPicosSuperior) {
        abaPrincipal.getRange(index + 2, 5).setBackground("#57BB8A"); // Verde
      } else {
        abaPrincipal.getRange(index + 2, 5).setBackground("#E06666"); // Vermelho
      }
    }
    
  });

  /* ------------------------------------------------------------------------------------------------ */
  // Enviar Tabela de Preços Atualizados por Email:
  // Verifica se está no horário permitido
  const permitido = (
    (diaSemana >= 1 && diaSemana <= 5 && horaAtual >= 7 && horaAtual < 17) || 
    (diaSemana === 6 && horaAtual >= 7 && horaAtual < 12)
  );
  if (!permitido) return Logger.log("Envio fora do horário permitido.");

  const dados = planilha.getSheetByName("Atualizados").getDataRange().getValues();
  if (dados.length <= 1) return Logger.log("Não há produtos para atualizar.");

  const doc = DocumentApp.create(`Produtos Atualizados - ${dataHora}`);
  const body = doc.getBody();

  // Adiciona cabeçalho ao documento
  body.appendParagraph("Produtos Atualizados")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`Data: ${dataHora}`)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("");

  // Formata tabela
  const tabela = [["Código", "Preço de Venda"], ...dados.slice(1).map(row => [row[0], row[1].toFixed(2).replace(".", ",")])];
  const table = body.appendTable(tabela);

  // Estiliza cabeçalho
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i)
      .setBackgroundColor("#f3f3f3")
      .getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  // Estiliza células da tabela
  for (let i = 1; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    for (let j = 0; j < row.getNumCells(); j++) {
      row.getCell(j)
        .getChild(0).asParagraph().setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  }

  doc.saveAndClose();

  // Converte para PDF, envia email e exclui o documento
  const pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');

  MailApp.sendEmail({
    to: emailDestino,
    subject: `Produtos atualizados - ${dataHora}`,
    body: `Segue em anexo a lista com produtos e preços atualizados hoje (${dataHora}).`,
    attachments: [pdf]
  });
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  Logger.log("Email enviado com sucesso!");


  /* ------------------------------------------------------------------------------------------------ */
  // Final do Script, se sucesso ou erro:
  Logger.log("Formatação de preços realizada com sucesso :)");
  } catch (error) {
    Logger.log("Erro ao formatar preços: " + error);
  }
}

