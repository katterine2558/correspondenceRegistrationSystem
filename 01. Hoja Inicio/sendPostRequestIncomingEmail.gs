function sendPostRequestIncomingEmail(sheetName,date,hour,etiquetaGmail,fromuser,touser,subject,nombreArchivosAdjuntos,folderName,fechaRespuesta,plazoDias,idEmail,carpetaTemporal_ID) {
  
  // URL servidor
  var scriptProperties = PropertiesService.getScriptProperties();
  var servidor =  scriptProperties.getProperty('servidor');

  //Cuenta los archivos adjuntos
  if (nombreArchivosAdjuntos === "")
  {
    var numAttachedFiles = 0;
  }
  else{
    var numAttachedFiles = nombreArchivosAdjuntos.split(",").length;
  }

  // Datos a enviar
  var data = {
    sheetName: sheetName,
    date: date,
    hour: hour,
    etiquetaGmail: etiquetaGmail,
    fromuser: fromuser,
    touser: touser,
    subject: subject,
    numAttachedFiles: numAttachedFiles,
    nombreArchivosAdjuntos: nombreArchivosAdjuntos,
    folderName: folderName,
    fechaRespuesta: fechaRespuesta,
    plazoDias: plazoDias,
    idEmail: idEmail,
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
