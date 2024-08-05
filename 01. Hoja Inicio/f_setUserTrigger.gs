/*
  CREA EL TRIGGER DEL USUARIO
*/


function setUserTrigger(){

  //Declaración variables
  var scriptProperties = PropertiesService.getScriptProperties();
  var hoursInterval = scriptProperties.getProperty('hoursInterval'); // Ajustar frecuencia (horas)
  var hoursInterval = parseFloat(hoursInterval);
  var labelName =  scriptProperties.getProperty('etiquetaGmail'); //Nombre de la etiqueta
  var clientes =  scriptProperties.getProperty('clientes'); //Correos de los clientes
  var clientes = clientes.split(',');
  var keywords =  scriptProperties.getProperty('keywords'); //Palabras claves
  var keywords = keywords.split(',');
  var id_sheetMirror =  scriptProperties.getProperty('id_sheetMirror'); //ID de la hoja espejo
  var servidor =  scriptProperties.getProperty('servidor'); //Servidor web

  //Obtiene el usuario activo
  var userEmail =   Session.getActiveUser().getEmail();

  //Verifica si ya existe la etiqueta en el usuario. 
  var labels = Gmail.Users.Labels.list('me')
  var labelExists = false;
  for (var i = 0; i < labels.labels.length; i++){
    var label = labels.labels[i];
    if (label.name === labelName) {
      labelExists = true;
      break;
    }
  }
  //Si no existe, la crea
  if (!labelExists) {
    GmailApp.createLabel(labelName);
    //Obtiene la etiqueta creada
    var label = Gmail.Users.Labels.list('me').labels.find(function(l) { return l.name === labelName; });
    //Concatena el correo de los clientes
    var fromEmail = ""
    for (var i = 0; i < clientes.length; i++) {
      var word = clientes[i].trim();
      fromEmail =  fromEmail + word;
      if (i !== clientes.length - 1 ) {
        fromEmail = fromEmail + " OR "
      }
    };
    //Genera los filtros
    for (var i = 0; i < keywords.length; i++) {
      var word = keywords[i].trim(); // Eliminar espacios en blanco alrededor de la palabra
      var query = "(subject:" + word + " OR " + word + " in:body)";
      var newFilter = {
      criteria: {
        from: fromEmail,
        query: word
      },
      action: {
        addLabelIds: [label.id]
              }
      };
      Gmail.Users.Settings.Filters.create(newFilter, "me");
    };

    // Crea el trigger de lectura de correos entrantes para el usuario
    var trigger = ScriptApp.newTrigger("readInboxEmail")
        .timeBased()
        .everyHours(hoursInterval) 
        //.everyMinutes(hoursInterval) 
        .create();
    console.log('Trigger de correos entrantes creado para: ', userEmail);

    // Crea el trigger de lectura de correos salientes para el usuario
    var trigger = ScriptApp.newTrigger("readSentEmail")
        .timeBased()
        .atHour(20)  // Ejecutar a las 8 PM
        .everyDays(1) // Ejecutar todos los días
        .create();
    console.log('Trigger de correos salientes creado para: ', userEmail);

    //Cuerpo de data para post (constantes y triggers)
    data = {
      sheetName : "Autorizaciones",
      userEmail: userEmail,
    }

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

    Browser.msgBox("Se ha creado la etiqueta "+labelName+ " en su correo.")

  }
}