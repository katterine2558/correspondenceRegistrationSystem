/*
LEE LA BANDEJA DE ENTRADA DEL CORREO
*/
function readInboxEmail() {

  // Variables del proyecto
  var scriptProperties = PropertiesService.getScriptProperties();
  var etiquetaGmail = scriptProperties.getProperty('etiquetaGmail'); // etiqueta 
  var idFolder = scriptProperties.getProperty('idFolder');  //ID de la unidad compartida para almacenar los archivos adjuntos
  var numDiasRespuesta = scriptProperties.getProperty('numDiasRespuesta');  // Número de días para la respuesta
  var numDiasRespuesta = parseFloat(numDiasRespuesta);
  var incluye_sabado = scriptProperties.getProperty('incluye_sabado');  // Booleano para tener en cuenta los sábados. Sólo para funciona para días habiles
  var incluye_sabado = JSON.parse(incluye_sabado)
  var festivos = scriptProperties.getProperty('festivos').split(','); // Festivos 
  var tipoPlazoRespuesta = scriptProperties.getProperty('tipoPlazoRespuesta'); //Tipo de plazo de respuesta: días habiles o calendario
  var id_sheetMirror =  scriptProperties.getProperty('id_sheetMirror'); //ID de la hoja espejo
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
    var mensajes = emails[i].getMessages()
    
    //Itera por los hilos de ese correo. Revisa que sea alguno de los clientes quien envió el correo. Empieza a recorrerlo de atrás para adelante.
    for(var j = 0 ; j < mensajes.length; j++){

      //Obtiene los metadatos del correo
      var metadata = getEmailMetadata(mensajes[j])

      //Genera el ID del correo
      var idEmail = generateEmailID(metadata)

      //Verifica si el correo ya existe en el registro
      var existingID = checkInboxEmailID(idEmail,id_sheetMirror)

      //Sigue el proceso si y solo si el ID no existe en la hoja espejo y el remitente esté incluído en la lista de clientes
      if (existingID === false && clientes.indexOf(metadata.From) !== -1){
        
        //Obtiene el nombre de los archivos adjuntos
        if (metadata.AttachedFiles.length >= 1) {
          //Obtiene el nombre de los archivos adjuntos
          var nombreArchivosAdjuntos = obtenerNombresArchivosAdjuntos(metadata.AttachedFiles)
        } else {
          var nombreArchivosAdjuntos = ""
        }

        // Crear una carpeta con la fecha del correo para almacenar los archivos adjuntos
        var carpeta = createFolder(metadata.Date,idFolder)

        //Se mueven los archivos a la carpeta indicada
        moveFilesToDrive(metadata.AttachedFiles,metadata.Date,idFolder,metadata.Body,idEmail,"ENTRADA");

        // Obtiene la fecha máxima de respuesta teniendo en cuenta el tipo de respuesta y la cantidad de dias.
        var fechaRespuesta = getResponseDate(metadata.Date,numDiasRespuesta,incluye_sabado,festivos,tipoPlazoRespuesta)

        //Estima la fecha de respuesta
        var plazoDias = estimateResponseTime(fechaRespuesta)

        //Obtiene el ID de la carpeta donde se almacenan los archivos
        var carpetaTemporal = findFolder(folder,metadata.Date)
        var carpetaTemporal_ID = carpetaTemporal.getId()

        // Escribe el registro en la hoja espejo
        sendPostRequestIncomingEmail("Entrada",metadata.Date,metadata.Hour,etiquetaGmail,metadata.From,metadata.To,metadata.Subject,nombreArchivosAdjuntos,folderName,fechaRespuesta,plazoDias,idEmail,carpetaTemporal_ID)

      }
      else if(existingID !== false)
      {
        console.log("Ya existe el correo "+ idEmail + " en la PLANILLA.")
      }
      else if(clientes.indexOf(metadata.From) === -1)
      {
        console.log("El remitente de "+ idEmail + " no está en la lista de Clientes.")
      }
    }
  }
}

/*
GENERA EL ID DEL CORREO
*/
function generateEmailID(metadata) {

  //Reemplaza los "/" de la fecha por "-"
  var id1 = metadata.Date.replace(/\//g, "-");
  
  //Reemplaza los dos puntos de la hora por "-"
  var id2 = metadata.Hour.replace(/:/g, "-")

  //Genera el id concatenando también el asunto del correo
  var id = id1 + "-" + id2 + "-" + metadata.Subject;

  return id
  
}


/*
CHEQUEA SI YA EXISTE EL ID DEL EMAIL EN LA HOJA ESPEJO DE ENTRADA
*/
function checkInboxEmailID(idEmail,id_sheetMirror){

  //Obtiene el libro
  var book = SpreadsheetApp.openById(id_sheetMirror); 
  var sheet = book.getSheetByName("Entrada");

  //obtiene la última fila de la hoja
  var ultimaFila = sheet.getLastRow();

  if (ultimaFila !== 1){

    // Obtener los valores en la columna L a partir de la fila 2
    var valoresColumnaL = sheet.getRange("M2:M" + ultimaFila).getValues();

    // Convertir los valores en un array
    var arrayValores = valoresColumnaL.map(function(row) {
      return row[0];
    });

    // Verifica si el ID existe
    var estaPresente = arrayValores.includes(idEmail)

    return estaPresente

  }
  else{
    return false
  }

}
