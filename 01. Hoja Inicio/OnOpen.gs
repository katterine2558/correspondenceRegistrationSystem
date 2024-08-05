/*
  SE EJECUTA AL ABRIR LA HOJA DE CÁLCULO
*/
function onOpen() {

  //Obtiene el usuario activo
  var userEmail =   Session.getActiveUser().getEmail();

  // Objeto menú
  var ui = SpreadsheetApp.getUi();

  //Extrae la propiedad que establece si la URL del servidor ya está asignada
  var scriptProperties = PropertiesService.getScriptProperties();
  var bool_servidor = Boolean(scriptProperties.getProperty('bool_servidor'));
  //Extrae el booleano para verificar si la configuración de la hoja fue guardada
  var save_configurations = Boolean(scriptProperties.getProperty('save_configurations'));

  //Aquí se asigna la URL del servidor
  if (userEmail === "administrador@empresa.com")
  {
    //Menú de deploy de servidor
    var menu = ui.createMenu("Asignar implementación")
      .addItem("Insertar URL","saveURLDeploy")
      .addToUi();
  }
  
  //Menú de configuración
  var menu = ui.createMenu("Configurar Hoja Inicio")
    .addItem("Guardar configuración","projectConstants")
    .addToUi();

  //Menú de conexión
  if(bool_servidor ===true)
  {
    var menu = ui.createMenu("Conectar")
      .addItem("Conexión a Planilla","sheetConnection")
      .addToUi();
  }
  
  //Menú para autorizar registro. Lo hacen las personas que están trabajando en el proyecto
  if (save_configurations === true && bool_servidor ===true)
  {
    var menu = ui.createMenu("Autorizar SRC")
    .addItem("Aceptar","AuthProcess")
    .addToUi();
  }
  
}

// GUARDA LA URL CON EL SERVIDOR
function saveURLDeploy()
{
  var respuesta = SpreadsheetApp.getUi().prompt("Iniciar conexión", "Por favor, ingrese la URL del servidor", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

  // Verificar si el usuario ingresó un valor y presionó "OK"
  if (respuesta.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {

    //Extrae la respuesta
    var urlServidor = respuesta.getResponseText();
    if (urlServidor!=="")
    {
      var urlServidor = urlServidor.trim();
      // Llama al script como un objeto.
      var scriptProperties = PropertiesService.getScriptProperties();

      //Asigna las variables
      scriptProperties.setProperty('servidor', urlServidor);

      //Genera variable booleana de configuración de servidor
      scriptProperties.setProperty('bool_servidor', true);

      Browser.msgBox("Servidor asignado.");
      
    }

  } 

}

// PROCESO DE AUTORIZACIÓN DE USUARIO
function AuthProcess() {

  //Obtiene el usuario activo
  var userEmail =   Session.getActiveUser().getEmail();

  //Revisa si el usuario está autenticado
  var userAuthenticated = checkUserAuthentication(userEmail);
  console.log("el valor de userAuthenticated es: "+ userAuthenticated)

  if (userAuthenticated === undefined){
    console.log("El usuario "+userEmail+" no tiene permiso como lector");
    return;
  }
  else if(userAuthenticated === true){
    Browser.msgBox("El usuario: "+userEmail+ " ya ha autorizado al SRC.")
    console.log(userEmail + " ya existe y autorizó script");
  }
  else{
    setUserTrigger()       
  }
}

// HACE LLAMADO AL SERVIDOR PARA ESTABLECER LAS VARIABLES DE LA HOJA DE REGISTRO EN LA HOJA ESPEJO
function sheetConnection(){

  //Obtiene el usuario activo
  var userEmail =   Session.getActiveUser().getEmail();

  //Intenta obtener la variable para verificar si se guardo la configuración
  var scriptProperties = PropertiesService.getScriptProperties();
  var save_configurations = Boolean(scriptProperties.getProperty('save_configurations'));
  var servidor = scriptProperties.getProperty('servidor'); //url del servidor

  if (save_configurations !== true)
  {
    Browser.msgBox("No se ha guardado la configuración del proyecto.");
    return;
  }

  if(servidor === undefined || servidor === null)
  {
    Browser.msgBox("No se ha establecido conexión con el servidor.");
    return;
  }
  
  var owner = scriptProperties.getProperty('owner'); //administrador de la hoja espejo
  if(owner === userEmail)
  {
    //Obtiene las propieades
    var numDiasRespuesta = scriptProperties.getProperty('numDiasRespuesta');  // Número de días para la respuesta
    var numDiasRespuesta = parseFloat(numDiasRespuesta);
    var incluye_sabado = scriptProperties.getProperty('incluye_sabado');  // Booleano para tener en cuenta los sábados. Sólo para funciona para días habiles
    var incluye_sabado = JSON.parse(incluye_sabado);
    var festivos = scriptProperties.getProperty('festivos'); // Festivos 
    var festivos = festivos.split(',');
    var tipoPlazoRespuesta = scriptProperties.getProperty('tipoPlazoRespuesta'); //Tipo de plazo de respuesta: días habiles o calendario
    var lideres_proyectos = scriptProperties.getProperty('lideres_proyectos'); //Lideres del proyecto
    
    //Cuerpo de data para post (constantes y triggers)
    data = {
      sheetName : "Conexión",
      numDiasRespuesta: numDiasRespuesta,
      incluye_sabado: incluye_sabado,
      festivos: festivos,
      tipoPlazoRespuesta: tipoPlazoRespuesta,
      lideres_proyectos: lideres_proyectos
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

    Browser.msgBox("Sistema de Registro de Correspondencia inicializado.");
  }
  else
  {
    Browser.msgBox("No tiene permiso para inicializar el Sistema de Registro de Correspondencia.");
    return;
  }

}
  

