function sendPostRequestSendEmail(sheetName,identificador,sentDate,sentHour,labelEmail,fromEmail,toEmail,subjectEmail,attachedFiles,folderName,carpetaTemporal_ID) {
  
  // URL servidor
  var scriptProperties = PropertiesService.getScriptProperties();
  var servidor =  scriptProperties.getProperty('servidor');

  //Cuenta los archivos adjuntos
  if (attachedFiles === "")
  {
    var numAttachedFiles = 0;
  }
  else{
    var numAttachedFiles = attachedFiles.split(",").length;
  }

  // Datos a enviar
  var data = {
    sheetName: sheetName,
    identificador: identificador,
    sentDate: sentDate,
    sentHour: sentHour,
    labelEmail: labelEmail,
    fromEmail: fromEmail,
    toEmail: toEmail,
    subjectEmail: subjectEmail,
    numAttachedFiles: numAttachedFiles,
    attachedFiles: attachedFiles,
    folderName: folderName,
    carpetaTemporal_ID: carpetaTemporal_ID
  };

  // Configuración de la solicitud POST
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data)
  };

  // Realiza la solicitud POST
  var response = UrlFetchApp.fetch(servidor, options);

  // Muestra la respuesta (opcional)
  Logger.log(response.getContentText());

}
