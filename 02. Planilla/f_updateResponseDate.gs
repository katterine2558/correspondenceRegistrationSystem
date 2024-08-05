function updateResponseDate() {

  //Extrae los parámetros que necesita
  var scriptProperties = PropertiesService.getScriptProperties();
  var numDiasRespuesta = scriptProperties.getProperty('numDiasRespuesta');  // Número de días para la respuesta
  var numDiasRespuesta = parseFloat(numDiasRespuesta);
  var incluye_sabado = scriptProperties.getProperty('incluye_sabado');  // Booleano para tener en cuenta los sábados. Sólo para funciona para días habiles
  var incluye_sabado = JSON.parse(incluye_sabado)
  var festivos = scriptProperties.getProperty('festivos'); // Festivos 
  var festivos = festivos.split(',');
  var tipoPlazoRespuesta = scriptProperties.getProperty('tipoPlazoRespuesta'); //Tipo de plazo de respuesta: días habiles o calendario

  //Obtiene la hoja
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName("Entrada");

  //Obtiene el rango de filas que tienen texto en la columna K
  var rangoColumnaK = sheet.getRange("L2:L");

  // Obtener los valores de la columna K
  var valoresColumnaK = rangoColumnaK.getDisplayValues();
  
  var fila = 2;
  for (var valor of valoresColumnaK) {
    if (valor[0] !== ""){
      var fechaRespuesta = sheet.getRange(fila,11).getDisplayValue();
      console.log(fechaRespuesta)
      var plazoDias = estimateResponseTime(fechaRespuesta,incluye_sabado,festivos,tipoPlazoRespuesta,numDiasRespuesta)
      sheet.getRange(fila, 12).setValue(plazoDias);
      console.log("Nuevo plazo: "+plazoDias)
    }
    fila++;
  }
}