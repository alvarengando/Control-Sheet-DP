function ReformularFaltaNova(spreadsheet) {
  spreadsheet
    .getRangeList(["G5", "H6:H7", "K6:K7", "G10"])
    .clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet
    .getRange("D4")
    .setFormula('=IF(G5="";"";COUNTA(\'Faltas Dados\'!A:A))');
  spreadsheet
    .getRange("D5")
    .setFormula(
      "=IF(G5=\"\";\"\";INDEX('Cadastro Dados'!A2:A;MATCH(G5;'Cadastro Dados'!B2:B)))"
    );
  spreadsheet.getRange("D6").setFormula('=IF(G5="";"";Today())');
  spreadsheet.getRange("K4").setValue("False");
  // spreadsheet.getRange("H6").setFormula('=IF(G5="";"";Today())');
}
/** ********* **************************************************** */

function ModoFaltaNova() {
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange("AL3").setValue(1);
  spreadsheet.getRange("D1").setValue("Novo");

  ReformularFaltaNova(spreadsheet);

  spreadsheet.getRange("G5").activate();
}

// **************************************************************************************************

function SalvarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltas = spreadsheet.getSheetByName("Faltas");
  var faltaDados = spreadsheet.getSheetByName("Faltas Dados");

  if (
    faltas.getRange("AI3").getValue() > 0 ||
    faltas.getRange("AH3").getValue() > 0
  ) {
    if (faltas.getRange("AI3").getValue() > 0) {
      Browser.msgBox(
        "Erro",
        "Necessário preencher os campos Funcionário, Data de Falta e Dedução de horas!",
        Browser.Buttons.OK
      );
    } else {
      Browser.msgBox(
        "Erro",
        "Necessário informar a data e quantidade de dias do Atestado!",
        Browser.Buttons.OK
      );
    }
  } else {
    // Salvar na Página Faltas Dados

    var values = [
      [
        faltas.getRange("D4").getValue(), // ID
        faltas.getRange("D5").getValue(), // ID Funcionário
        faltas.getRange("D6").getValue(), // Data Lançamento
        faltas.getRange("G5").getValue(), // Nome Funcionário
        faltas.getRange("H6").getValue(), // Data de Falta
        faltas.getRange("H7").getValue(), // Dedução
        faltas.getRange("G10").getValue(), // Observação
        faltas.getRange("K4").getValue(), // Atestado
        faltas.getRange("K6").getValue(), // Data do Atestado
        faltas.getRange("K7").getValue(), // Dias de Atestado
      ],
    ];

    faltaDados
      .getRange(faltaDados.getLastRow() + 1, 1, 1, 10)
      .setValues(values);

    spreadsheet
      .getRangeList(["G5", "H6:H7", "K6:K7", "G10"])
      .clear({ contentsOnly: true, skipFilteredRows: true });
    
    spreadsheet.getRange("K4").setValue("False");

    Browser.msgBox(
      "Informativo",
      "Registro salvo com sucesso!",
      Browser.Buttons.OK
    );

    faltas.getRange("G5").activate();
  }
}

//Modo Deletar Registro

function modoDeletarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltas = spreadsheet.getSheetByName("Faltas");

  faltas.getRange("AL3").setValue(3);

  spreadsheet
    .getRangeList(["G5", "H6:H7", "K5"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  faltas.getRange("D1").setValue("Deletar");
  //ID Pedido

  faltas
    .getRange("D4")
    .setBackground("#ffffff")
    .setFontColor("#000000")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange("'Faltas'!$BH$4:$BH"), true)
        .build()
    );

  faltas.getRange("C16").activate();
}

//Deletar Registro

function deletarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltas = spreadsheet.getSheetByName("Faltas");
  var faltasDados = spreadsheet.getSheetByName("Faltas Dados");
  var linhaPedido = faltas.getRange("AJ3").getValue(); //linha correspondente em Faltas Dados

  if (faltas.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário preencher todos os campos essenciais!",
      Browser.Buttons.OK
    );
  } else {
    faltasDados.deleteRow(linhaPedido);
    faltas
      .getRangeList(["D4", "C13", "C16", "J11", "M10"])
      .clear({ contentsOnly: true, skipFilteredRows: true });
    Browser.msgBox(
      "Informativo",
      "Registro Deletado com sucesso!",
      Browser.Buttons.OK
    );

    faltas.getRange("C16").activate();
  }
}

function FinalizadorFalta() {
  let spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange("AL3").getValue() == 1) {
    SalvarFalta();
  } else if (spreadsheet.getRange("AL3").getValue() == 2) {
    EditarFalta();
  } else {
    DeletarFalta();
  }
}

function RelatoriosDespesasDialog() {
  var url =
    "https://datastudio.google.com/reporting/2d18cae6-c82e-4379-8e36-fa8af16f5a96/page/feyJB";
  var name = "Despesas Consolidado";

  var url2 =
    "https://datastudio.google.com/reporting/8bc163d7-8140-4abb-b6d3-f8d708d64d7c/page/feyJB";
  var name2 = "Vendas Analítico";

  var html =
    '<html><body><a href="' +
    url +
    '" target="blank" onclick="google.script.host.close()">' +
    name +
    '</a> <br><br/><a href="' +
    url2 +
    '" target="blank" onclick="google.script.host.close()">' +
    name2 +
    "</a></body></html>";
  var ui = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(ui, "Relatórios de Vendas");
}
