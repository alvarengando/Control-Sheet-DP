function absc() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet
    .getRange("C10")
    .activate()
    .setFormula('=IF(H4="";"";COUNTA(Dados!A2:A))');
}

function COLOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange("C6").activate();
  spreadsheet.getActiveRangeList().setFontColor("#0c343d");
}

function indicecorresp() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet
    .getRange("D4")
    .activate()
    .setFormula(
      "=IF(G5=\"\";\"\";INDEX('Cadastro Dados'!A2:A;MATCH(G5;'Cadastro Dados'!B2:B)))"
    );
}
