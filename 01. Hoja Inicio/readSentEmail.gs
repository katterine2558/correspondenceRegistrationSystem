/*
LEE LA BANDEJA DE SALIDA DEL CORREO
*/
function readSentEmail() {

  // Variables del proyecto
  var scriptProperties = PropertiesService.getScriptProperties();
  var etiquetaGmail = scriptProperties.getProperty('etiquetaGmail'); // etiqueta 
  var idFolder = scriptProperties.getProperty('idFolder');  //ID de la unidad compartida para almacenar los archivos adjuntos
  var id_sheetMirror =  scriptProperties.getProperty('id_sheetMirror'); //ID de la hoja espejo
  var servidor =  scriptProperties.getProperty('servidor'); //Servidor web
  var clientes = scriptProperties.getProperty('clientes').split(','); // Clientes del proyecto 
  var clientes = clientes.map(function(cadena) {
    return cadena.trim();
  });

  //Obtiene el nombre de la carpeta. Para mover los archivos adjuntos
  var folder = DriveApp.getFolderById(idFolder)
  var folderName = folder.getName()

  // Obtener los correos no leídos de la etiqueta especificada
  var labelFolder = GmailApp.getUserLabelByName(etiquetaGmail)

  // Obtiene todos los hilos (threads) dentro de la carpeta de la etiqueta
  var emails = labelFolder.getThreads();

  // Itera por los correos
  for (var i = 0; i < emails.length ; i ++){

    //Mensajes asociados a cada correo
    var mensajes = emails[i].getMessages();
    
    //Itera por los mensajes del correo. Revisa que sea alguno de los clientes quien envió el correo. Empieza a recorrerlo de atrás para adelante.
    for(var j = mensajes.length - 1 ; j >= 0; j--){

      //Obtiene los metadatos del correo de respuesta
      var metadataEmail = getEmailMetadata(mensajes[j])

      //Obtiene la lista de los destinatarios
      var destinatariosArray = metadataEmail.To.split(',').map(function(item) {
        return item.trim();
      });

      //Revisa que el remitente sea el usuario en sesión y que el destinatorio del correo sea algún cliente
      if(metadataEmail.From === Session.getActiveUser().getEmail() && destinatariosArray.some(value => clientes.includes(value)))
      {

        //Obtiene los metadatos del correo del cliente el cual fue contestado anteriormente. Es importante destacar que escoge el primero correo el cual el remitente coincida con el destinatario del correo actual de la iteración
        var clientEmail = getClientMessage(mensajes,j,clientes,destinatariosArray)

        //Si no encuentra el correo, sigue la iteración 
        if (clientEmail === undefined)
        {
          continue;
        }

        //Obtiene el ID del correo previamente enviado por el cliente
        var clientEmailID = generateEmailID(clientEmail)

        //Verifica si el correo el status del correo, es decir, si no ha sido contestado
        var repliedStatus = checkMailStatus(clientEmailID,id_sheetMirror)

        //Continua el proceso siempre y cuando el correo no se haya cambiado el estatus en la hoja de entrada
        if (repliedStatus === false){
          
          //Obtiene el nombre de los archivos adjuntos
          if (metadataEmail.AttachedFiles.length >= 1) {
            //Obtiene el nombre de los archivos adjuntos
            var nombreArchivosAdjuntos = obtenerNombresArchivosAdjuntos(metadataEmail.AttachedFiles)
          } else {
            var nombreArchivosAdjuntos = ""
          }

          // Crear una carpeta con la fecha del correo para almacenar los archivos adjuntos
          createFolder(metadataEmail.Date,idFolder)

          //Se mueven los archivos a la carpeta indicada
          moveFilesToDrive(metadataEmail.AttachedFiles,metadataEmail.Date,idFolder,metadataEmail.Body,clientEmailID,"SALIDA");

          //Obtiene el ID de la carpeta donde se almacenan los archivos
          var carpetaTemporal = findFolder(folder,metadataEmail.Date)
          var carpetaTemporal_ID = carpetaTemporal.getId()

          //Escribe el registro en la hoja espejo
          sendPostRequestSendEmail("Salida",clientEmailID,metadataEmail.Date,metadataEmail.Hour,etiquetaGmail,metadataEmail.From,metadataEmail.To,metadataEmail.Subject,nombreArchivosAdjuntos,folderName,carpetaTemporal_ID)

        }

      }
    }
  }
}

/*
Obtiene el correo del cliente el cual fue respondido. Siempre va a iterar de j hacia arriba
*/
function getClientMessage(mensajes,index_mensaje,clientes,destinatarios_respuesta)
{
  for (var i = index_mensaje - 1 ; i >= 0 ; i--)
  {
    //Metadata del correo 
    var mail = getEmailMetadata(mensajes[i])
    //Respondió el correo
    if(clientes.indexOf(mail.From) !== -1 && destinatarios_respuesta.indexOf(mail.From) !== -1){
      return mail
    }
    
  }
}

/*
OBTIENE EL ID DEL CORREO DEL CLIENTE QUE SE RESPONDIÓ. LO BUSCA EN LA HOJA DE SALIDAS
*/
function getClientMailID(metadata,clientEmailMetadata,id_sheetMirror){

  //tiene el asunto del correo
  var asunto = metadata.Subject

  //Fromatea la hora
  var hora = clientEmailMetadata.Hour.split(":")

  //Lee la hoja de los correos entrantes
  var book = SpreadsheetApp.openById(id_sheetMirror);
  var sheet = book.getSheetByName("Entrada");
  var ID_array = sheet.getRange("L2:L3000").getValues();
  for (var i = 0; i < ID_array.length; i++) {
    var valorCelda = ID_array[i][0];
    if (valorCelda !== '' && valorCelda.includes(asunto)) {
      if (valorCelda.includes(clientEmailMetadata.Date+"-"+hora[0]+"-"+hora[1])){
        return valorCelda
      }  
    }
  }
  return undefined

}


/*
CHEQUEA SI EL CORREO DEL CLIENTE YA SE CONTESTÓ
*/
function checkMailStatus(idEmail,id_sheetMirror){

  if (idEmail === undefined){
    return undefined
  }
  else{

    //Lee la hoja de los correos entrantes
    var book = SpreadsheetApp.openById(id_sheetMirror);
    var sheet = book.getSheetByName("Entrada");
    var ID_array = sheet.getRange("M2:M3000").getValues();
    for (var i = 0; i < ID_array.length; i++) {
      var valorCelda = ID_array[i][0];
      if (valorCelda === idEmail) {
        if (sheet.getRange(i+2,1).getDisplayValue() === "Sin contestar"){
          return false
        }
        else{
          return true
        }
      }
    }


  }

}

