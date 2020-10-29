function FormularFaltaNova(spreadsheet) {
  spreadsheet
    .getRangeList(["G5", "H6:H7", "K6:K7", "G10"])
    .clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet
    .getRange("D4")
    .clearDataValidations()
    .setFontColor("#ffffff")
    .setBackground("#134f5c")
    .setFormula('=IF(G5="";"";COUNTA(\'Faltas Dados\'!A:A))');
  spreadsheet
    .getRange("D5")
    .setFormula(
      "=IF(G5=\"\";\"\";INDEX('Cadastro Dados'!A2:A;MATCH(G5;'Cadastro Dados'!B2:B)))"
    );

  spreadsheet
    .getRange("G5")
    .clearContent()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(
          spreadsheet.getRange("'Cadastro Dados'!$B$2:$B"),
          true
        )
        .build()
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

  FormularFaltaNova(spreadsheet);

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
        "Oshi... 😑",
        "Necessário preencher os campos Funcionário, Data de Falta e Dedução de horas! Me Ajuda!",
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
/** ******************************************************************************** */

//Formular editar Falta
function FormularEditarFalta(spreadsheet) {
  spreadsheet
    .getRangeList(["D4", "G5", "H6:H7", "K6:K7", "G10"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet.getRange("D5").setFormula('=IF(D4="";"";AR4)');
  spreadsheet.getRange("D6").setFormula('=IF(D4="";"";AS4)');
  spreadsheet.getRange("H6").setFormula('=IF(D4="";"";AU4)');
  spreadsheet.getRange("H7").setFormula('=IF(D4="";"";AV4)');
  spreadsheet.getRange("G10").setFormula('=IF(D4="";"";AW4)');
  spreadsheet.getRange("K4").setFormula('=IF(D4="";False;AX4)');
  spreadsheet.getRange("K6").setFormula('=IF(D4="";"";AY4)');
  spreadsheet.getRange("K7").setFormula('=IF(D4="";"";AZ4)');
}

//Modo Editar Falta
function ModoEditarFalta() {
  let spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange("AL3").setValue(2);
  spreadsheet.getRange("D1").setValue("Editar");

  spreadsheet
    .getRange("D4")
    .setBackground("#ffffff")
    .setFontColor("#0c343d")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange("'Faltas'!$AN$4:$AN"), true)
        .build()
    );

  spreadsheet
    .getRange("G5")
    .clearContent()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(
          spreadsheet.getRange("'Faltas Dados'!$D$2:$D"),
          true
        )
        .build()
    );

  FormularEditarFalta(spreadsheet);

  spreadsheet.getRange("G5").activate();
}

//Salvar alteração
function EditarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltasDados = spreadsheet.getSheetByName("Faltas Dados");
  var linhaFalta = spreadsheet.getRange("Ak3").getValue(); //linha correspondente em Faltas dados

  if (spreadsheet.getRange("AI3").getValue() > 0) {
    Browser.msgBox(
      " Oshi... 😑",
      "Sério? Como deseja alterar um registro sem seleciona-lo? Me ajuda!",
      Browser.Buttons.OK
    );
  } else {
    // Salvar na Página Faltas Dados

    var values = [
      [
        spreadsheet.getRange("D4").getValue(), // ID
        spreadsheet.getRange("D5").getValue(), // ID Funcionário
        spreadsheet.getRange("D6").getValue(), // Data Lançamento
        spreadsheet.getRange("G5").getValue(), // Nome Funcionário
        spreadsheet.getRange("H6").getValue(), // Data de Falta
        spreadsheet.getRange("H7").getValue(), // Dedução
        spreadsheet.getRange("G10").getValue(), // Observação
        spreadsheet.getRange("K4").getValue(), // Atestado
        spreadsheet.getRange("K6").getValue(), // Data do Atestado
        spreadsheet.getRange("K7").getValue(), // Dias de Atestado
      ],
    ];

    faltasDados.getRange(linhaFalta, 1, 1, 10).setValues(values);

    Browser.msgBox(
      "Uhuu!!",
      "Registro alterado com sucesso! Também atirando com o meu Fúsil fica Fácil! 😏",
      Browser.Buttons.OK
    );

    FormularEditarFalta(spreadsheet);
    spreadsheet.getRange("G5").activate();
  }
}

/**  ************************************************************************** */
//Modo Deletar Registro

//Formular editar Falta
function FormularDeletarFalta(spreadsheet) {
  spreadsheet
    .getRangeList(["D4", "G5", "H6:H7", "K6:K7", "G10"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  spreadsheet.getRange("D5").setFormula('=IF(D4="";"";AR4)');
  spreadsheet.getRange("D6").setFormula('=IF(D4="";"";AS4)');
  spreadsheet.getRange("H6").setFormula('=IF(D4="";"";AU4)');
  spreadsheet.getRange("H7").setFormula('=IF(D4="";"";AV4)');
  spreadsheet.getRange("G10").setFormula('=IF(D4="";"";AW4)');
  spreadsheet.getRange("K4").setFormula('=IF(D4="";False;AX4)');
  spreadsheet.getRange("K6").setFormula('=IF(D4="";"";AY4)');
  spreadsheet.getRange("K7").setFormula('=IF(D4="";"";AZ4)');
}

//Modo Deletar Falta
function ModoDeletarFalta() {
  let spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange("AL3").setValue(3);
  spreadsheet.getRange("D1").setValue("Deletar");

  spreadsheet
    .getRange("D4")
    .setBackground("#ffffff")
    .setFontColor("#0c343d")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(spreadsheet.getRange("'Faltas'!$AN$4:$AN"), true)
        .build()
    );

  spreadsheet
    .getRange("G5")
    .clearContent()
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(
          spreadsheet.getRange("'Faltas Dados'!$D$2:$D"),
          true
        )
        .build()
    );

  FormularDeletarFalta(spreadsheet);

  spreadsheet.getRange("G5").activate();
}

//Salvar Exclusão
function DeletarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltasDados = spreadsheet.getSheetByName("Faltas Dados");
  var linhaFalta = spreadsheet.getRange("Ak3").getValue(); //linha correspondente em Faltas dados

  if (spreadsheet.getRange("AI3").getValue() > 0) {
    Browser.msgBox(
      " Oshi... 😑",
      "Sério? Como deseja deletar um registro sem seleciona-lo? Me ajuda!",
      Browser.Buttons.OK
    );
  } else {
    // Exluir na Página Faltas Dados

    faltasDados.deleteRow(linhaFalta);

    Browser.msgBox(
      "Uhuu!!",
      "Registro Deletado com sucesso! Também atirando com o meu Fúsil fica Fácil! 😏",
      Browser.Buttons.OK
    );

    FormularDeletarFalta(spreadsheet);
    spreadsheet.getRange("G5").activate();
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
