function modoNovo() {
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange("AL3").setValue(1);
  spreadsheet.getRange("D1").setValue("Novo");
  //Limpar

  spreadsheet
    .getRangeList(["G5", "H6:H7", "K5"])
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
  // spreadsheet.getRange("H6").setFormula('=IF(G5="";"";Today())');

  spreadsheet.getRange("G5").activate();
}

// **************************************************************************************************

function salvarFalta() {
  var spreadsheet = SpreadsheetApp.getActive();
  var faltas = spreadsheet.getSheetByName("Faltas");
  var faltaDados = spreadsheet.getSheetByName("Faltas Dados");

  if (faltas.getRange("AK3").getValue() > 0) {
    Browser.msgBox(
      "Erro",
      "Necessário preencher todos os campos essenciais!",
      Browser.Buttons.OK
    );
  } else {
    // Salvar na Página Faltas Dados

    var values = [
      [
        faltas.getRange("D4").getValue(), // ID
        faltas.getRange("D5").getValue(), // ID Funcionário
        faltas.getRange("D6").getValue(), // Data Lançamento
        faltas.getRange("G5").getValue(), // Nome
        faltas.getRange("H6").getValue(), // Data de Falta
        faltas.getRange("H7").getValue(), // Dedução
        faltas.getRange("K5").getValue(), // Justificativa
      ],
    ];

    faltaDados.getRange(faltaDados.getLastRow() + 1, 1, 1, 7).setValues(values);
    spreadsheet
      .getRangeList(["G5", "H6:H7", "K5"])
      .clear({ contentsOnly: true, skipFilteredRows: true });

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
  var faltas = spreadsheet.getSheetByName('Faltas');

  faltas.getRange('AL3').setValue(3);
  
  spreadsheet
    .getRangeList(["G5", "H6:H7", "K5"])
    .clear({ contentsOnly: true, skipFilteredRows: true });

  faltas.getRange('D1').setValue("Deletar");
  //ID Pedido

  faltas.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Faltas\'!$BH$4:$BH'), true).build());

  faltas.getRange('C16').activate();

};

//Deletar Registro

function deletarFalta() {

  var spreadsheet = SpreadsheetApp.getActive();
  var faltas = spreadsheet.getSheetByName('Faltas');
  var faltasDados = spreadsheet.getSheetByName('Faltas Dados');
  var linhaPedido = faltas.getRange('AJ3').getValue(); //linha correspondente em Faltas Dados


  if (faltas.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    faltasDados.deleteRow(linhaPedido);
    faltas.getRangeList(['D4', 'C13', 'C16', 'J11', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);

    faltas.getRange('C16').activate();

  }

};
