function sendAlertEmail() {

  //Lideres proyecto
  var scriptProperties = PropertiesService.getScriptProperties();
  var lideres_proyecto = scriptProperties.getProperty('lideres_proyecto'); // lideres del proyecto 

  //Obtiene la hoja
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName("Entrada");

  //Obtiene el rango de filas que tienen texto en la columna L
  var rangoColumnaL = sheet.getRange("L2:L");

  // Obtener los valores de la columna L
  var valoresColumnaL = rangoColumnaL.getDisplayValues();
  
  //Revisa si el plazo para responder el correo es menor o igual a 1 día.
  
  var fila = 2;
  for (var valor of valoresColumnaL) {
    if (valor[0] !== "" && valor[0] <= 1){

      //Obtiene el asunto del correo
      var asunto = sheet.getRange(fila,7).getDisplayValue();
      //Obtiene el remitente 
      var remitente = sheet.getRange(fila,5).getDisplayValue();
      //Obtiene la etiqueta
      var etiqueta = sheet.getRange(fila,4).getDisplayValue();
      //Obtiene fecha de recibido del correo
      var fechaRecibido = sheet.getRange(fila,2).getDisplayValue();
      //Obtiene destinatarios
      var destinatarios = sheet.getRange(fila,6).getDisplayValue();

      //Escribe el correo
      writeEmail(lideres_proyecto,asunto,remitente,etiqueta,fechaRecibido,destinatarios,valor[0])

    }
    fila++;
  }
  
}

//Envia el correo
function writeEmail(lideres_proyecto,asunto,remitente,etiqueta,fechaRecibido,destinatarios,plazoDia)
{

  var mensaje = "Quedan "+ plazoDia +" dias para contestar el correo enviado el "+ fechaRecibido +" por: "+remitente+ " con asunto "+ asunto + " al grupo de trabajo: "+ destinatarios + "\n\n\nEste es un correo automático, no responder.";
  
  // Enviar el correo electrónico con el remitente especificado
  GmailApp.sendEmail(lideres_proyecto, "ALERTA: Vencimiento plazo de respuesta "+etiqueta, mensaje, {from: "no-reply-notifications@empresa.com"});

}
