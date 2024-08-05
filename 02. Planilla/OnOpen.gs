function onOpen() {

  //Obtiene la hoja actual
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName("Entrada");
  
  //Crea los formatos para la hoja
  var reglasFormatoCondicionales = sheet.getConditionalFormatRules();

  console.log(reglasFormatoCondicionales.length)
  //Si no tiene formatos, los crea
  if (reglasFormatoCondicionales.length === 0){

    //Establece las condiciones para "sin contestar" en la columna A
    var reglaSinContestar = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Sin contestar")
      .setBackground("#F5C4C4")
      .setFontColor("#8F1313")
      .setRanges([sheet.getRange("A2:A5000")])
      .build();
    
    // Define la regla para "Contestado" en la columna A
    var reglaContestado = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Contestado")
      .setBackground("#A9D4A9")
      .setFontColor("#156B15")
      .setRanges([sheet.getRange("A2:A5000")])
      .build();

    // Define la regla para "No aplica" en la columna A
    var reglaNoAplica = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("No aplica")
      .setBackground("#FCE8B2")
      .setFontColor("#7f6000")
      .setRanges([sheet.getRange("A2:A5000")])
      .build();

    //Define la regla de color para la columna L cuando el valor del plazo de días está entre 5 y 4
    var plazoHolgado = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sheet.getRange("L2:L5000")])
      .whenNumberBetween(4, 100)
      .setFontColor("#3E9408")
      .build();
    
    //Define la regla de color para la columna L cuando el valor del plazo de días está entre 3 y 2
    var plazoMedio = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sheet.getRange("L2:L5000")])
      .whenNumberBetween(3, 2)
      .setFontColor("#FBBC04")
      .build();
    
    //Define la regla de color para la columna L cuando el valor del plazo de días está entre 1 y 0
    var plazoLimite = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sheet.getRange("L2:L5000")])
      .whenNumberBetween(0, 1)
      .setFontColor("#FF0000")
      .build();

    // Aplica las reglas al rango
    sheet.setConditionalFormatRules([reglaSinContestar, reglaContestado,reglaNoAplica,plazoHolgado,plazoLimite,plazoMedio]);
    sheet.getRange("B2:B10000").setNumberFormat('yyyy/mm/dd');
    sheet.getRange("K2:K10000").setNumberFormat('yyyy/mm/dd');

  }

  //Crea el botón para desplegar el aplicativo web
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Servidor SRC')
      .addItem('OK', 'popUpWindow')
      .addToUi();

  
}

function popUpWindow() {
  Browser.msgBox("SRC OK")
}