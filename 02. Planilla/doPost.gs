function doPost(e) {

  // Parsear el contenido del cuerpo de la solicitud
  var jsonString = e.postData.contents;
  var data = JSON.parse(jsonString);

  // Obtener el nombre de la hoja
  var sheetName = data.sheetName;
  
  //Verifica el nombre de la hoja para hacer acciones
  if (sheetName === "Entrada")
  {
    //Obtener hoja
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName(sheetName);

    //Fecha de envío del correo 
    var sentDate = data.date;

    //Hora de envío del correo 
    var sentHour = data.hour;

    //Etiqueta a revisar
    var labelEmail = data.etiquetaGmail;

    //Remitente
    var fromEmail = data.fromuser;

    //Destinatario
    var toEmail = data.touser;

    //Asunto
    var subjectEmail = data.subject;
    
    //Nombres de los archivos adjuntos
    var attachedFiles = data.nombreArchivosAdjuntos;

    //Número de archivos adjuntos
    var numAttachedFiles = data.numAttachedFiles;

    //Nombre de la carpeta donde se almacenan los archivos adjuntos
    var folderName = data.folderName;

    //Fecha máxima de respuesta
    var fechaRespuesta = data.fechaRespuesta;

    //Plazo en días para responder
    var plazoDias = data.plazoDias;

    //ID del email
    var idEmail = data.idEmail;

    //URL de la carpeta donde se almacenan los archivos adjuntos
    var idFolder = data.carpetaTemporal_ID;

    //Opciones desplegables
    var opcionesDesplegable = ["Sin contestar","Contestado", "No aplica"]

    // Obtener la última fila en la hoja de cálculo
    var ultimaFila = sheet.getLastRow() + 1;

    // Escribir en la hoja de cálculo
    sheet.getRange(ultimaFila, 2).setValue(sentDate);
    sheet.getRange(ultimaFila, 2).setNumberFormat('yyyy/mm/dd')
    sheet.getRange(ultimaFila, 3).setValue(sentHour);
    sheet.getRange(ultimaFila, 4).setValue(labelEmail);
    sheet.getRange(ultimaFila, 5).setValue(fromEmail);
    sheet.getRange(ultimaFila, 6).setValue(toEmail);
    sheet.getRange(ultimaFila, 7).setValue(subjectEmail);
    sheet.getRange(ultimaFila, 8).setValue(numAttachedFiles);
    sheet.getRange(ultimaFila, 9).setValue(attachedFiles);
    sheet.getRange(ultimaFila, 10).setValue(folderName + "/" + sentDate);
    sheet.getRange(ultimaFila, 11).setValue(fechaRespuesta);
    sheet.getRange(ultimaFila, 11).setNumberFormat('yyyy/mm/dd')
    sheet.getRange(ultimaFila, 12).setValue(plazoDias); sheet.getRange(ultimaFila, 12).setFontWeight("bold"); sheet.getRange(ultimaFila, 12).setBackground("#EFEFEF");
    sheet.getRange(ultimaFila, 13).setValue(idEmail);

    // Crear la regla de validación para la lista desplegable en la celda de la primera columna
    var reglaValidacion = SpreadsheetApp.newDataValidation().requireValueInList(opcionesDesplegable).build();
    
    // Aplicar la regla de validación a la celda de la primera columna
    sheet.getRange(ultimaFila, 1).setDataValidation(reglaValidacion);

    //Por defecto asigna el valor "Sin contestar"
    sheet.getRange(ultimaFila, 1).setValue("Sin contestar")

    // Insertar enlace en la celda en la columna 10 (J)
    var textoEnriquecido = SpreadsheetApp.newRichTextValue()
      .setText(folderName + "/" + sentDate)
      .setLinkUrl(0, (folderName + "/" + sentDate).length, "https://drive.google.com/drive/folders/"+ idFolder)
      .build();

    sheet.getRange(ultimaFila, 10).setRichTextValue(textoEnriquecido);

    // Devuelve una respuesta de éxito
    return ContentService.createTextOutput("Datos escritos correctamente: "+ idEmail );

  }
  else if(sheetName === "Salida")
  {

    //Obtener hoja
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName(sheetName);
    
    //Identificador del correo enviado por el cliente
    var identificador = data.identificador;

    //Fecha de envío del correo 
    var sentDate = data.sentDate;

    //Hora de envío del correo 
    var sentHour = data.sentHour;

    //Etiqueta a revisar
    var labelEmail = data.labelEmail;

    //Remitente
    var fromEmail = data.fromEmail;

    //Destinatario
    var toEmail = data.toEmail;

    //Asunto
    var subjectEmail = data.subjectEmail;

    //Número de archivos adjuntos
    var numAttachedFiles = data.numAttachedFiles;
    
    //Nombres de los archivos adjuntos
    var attachedFiles = data.attachedFiles;

    //Nombre de la carpeta donde se almacenan los archivos adjuntos
    var folderName = data.folderName;

    //URL de la carpeta donde se almacenan los archivos adjuntos
    var idFolder = data.carpetaTemporal_ID;

    // Obtener la última fila en la hoja de cálculo
    var ultimaFila = sheet.getLastRow() + 1;

    // Escribir en la hoja de cálculo
    sheet.getRange(ultimaFila, 1).setValue(identificador);
    sheet.getRange(ultimaFila, 2).setValue(sentDate);
    sheet.getRange(ultimaFila, 3).setValue(sentHour);
    sheet.getRange(ultimaFila, 4).setValue(labelEmail);
    sheet.getRange(ultimaFila, 5).setValue(fromEmail);
    sheet.getRange(ultimaFila, 6).setValue(toEmail);
    sheet.getRange(ultimaFila, 7).setValue(subjectEmail);
    sheet.getRange(ultimaFila, 8).setValue(numAttachedFiles);
    sheet.getRange(ultimaFila, 9).setValue(attachedFiles);
    sheet.getRange(ultimaFila, 10).setValue(folderName + "/" + sentDate);
    
    // Insertar enlace en la celda en la columna 10 (J)
    var textoEnriquecido = SpreadsheetApp.newRichTextValue()
      .setText(folderName + "/" + sentDate)
      .setLinkUrl(0, (folderName + "/" + sentDate).length, "https://drive.google.com/drive/folders/"+ idFolder)
      .build();

    sheet.getRange(ultimaFila, 10).setRichTextValue(textoEnriquecido);
    
    //Obtiene la hoja de entrada
    var sheet_incoming = book.getSheetByName("Entrada");

    //Busca el identificador del correo en la hoja de entrada y cambia el estatus
    var ultimaFila = sheet_incoming.getLastRow();
    var valoresColumnaM = sheet_incoming.getRange("M2:M" + ultimaFila).getValues();
    var arrayValores = valoresColumnaM.map(function(row) {
      return row[0];
    });
    var indexEmail = arrayValores.indexOf(identificador)
    sheet_incoming.getRange(indexEmail+2, 1).setValue("Contestado")

    //Elimina el plazo de días y la fecha máxima de respuesta
    sheet_incoming.getRange(indexEmail+2, 12).setValue("")
    sheet_incoming.getRange(indexEmail+2, 11).setValue("")

    // Devuelve una respuesta de éxito
    return ContentService.createTextOutput("Se cambio estado a Contestado el siguiente ID: " + identificador );
  }
  //Este condicional sirve para hacer la conexión entre las dos hojas
  else if(sheetName === "Conexión")
  {
    //Asigna las variables
    var numDiasRespuesta = data.numDiasRespuesta;
    var incluye_sabado = data.incluye_sabado;
    var festivos = data.festivos;
    var tipoPlazoRespuesta = data.tipoPlazoRespuesta;
    var lideres_proyectos = data.lideres_proyectos;

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('numDiasRespuesta', numDiasRespuesta);
    scriptProperties.setProperty('incluye_sabado', incluye_sabado);
    scriptProperties.setProperty('festivos', festivos.join(','));
    scriptProperties.setProperty('tipoPlazoRespuesta', tipoPlazoRespuesta);
    scriptProperties.setProperty('lideres_proyecto', lideres_proyectos);

    //Inicializa trigger updateResponseDate
    var existeTrigger = false;
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      if (trigger.getHandlerFunction() === "updateResponseDate") {
        existeTrigger = true; 
      }
    }
    if (!existeTrigger) {

      // Crea el trigger
      var trigger = ScriptApp.newTrigger("updateResponseDate")
        .timeBased()
        .atHour(6)  // Ejecutar a las 6 AM
        .everyDays(1) // Ejecutar todos los días
        .create();
    }

    //Inicializa trigger sendAlertEmail
    var existeTrigger = false;
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      if (trigger.getHandlerFunction() === "sendAlertEmail") {
        existeTrigger = true; 
      }
    }
    if (!existeTrigger) {
      // Crea el trigger
      var trigger = ScriptApp.newTrigger("sendAlertEmail")
        .timeBased()
        .atHour(7)  // Ejecutar a las 7 AM
        .everyDays(1) // Ejecutar todos los días
        .create();
  
    }

    return ContentService.createTextOutput("Se establecieron las constantes y se inicializaron los triggers");

  }
  else if(sheetName === "Autorizaciones")
  {

    //Obtener hoja
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName(sheetName);

    //Usuario 
    var userEmail = data.userEmail;

    sheet.appendRow([userEmail]);
    return ContentService.createTextOutput("Se añadió el usuario "+ userEmail + " a la hoja de Autorizaciones.");
  }
  
  
}
