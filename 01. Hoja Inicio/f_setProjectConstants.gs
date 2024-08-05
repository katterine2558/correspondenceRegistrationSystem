/*
  CONSTANTES DEL PROYECTO
*/
function projectConstants() {

  // Obtiene la hoja de calculo para sacar variables
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja inicio");

  //Clientes del proyecto
  try{
    var rango = sheet.getRange("B6:G8");
    var valores = rango.getValues();
    var clientes = [];
    for (var i = 0; i < valores.length; i++) {
      for (var j = 0; j < valores[i].length; j++) {
        if(valores[i][j] !== ""){
          clientes.push(valores[i][j].trim())
        }
      }
    }
    var clientes = clientes.join(',')

    //Verifica el formato del correo
    var boolFormato = verificacionMailFormato(clientes)
    if(!boolFormato)
    {
      Browser.msgBox("Correo clientes: Formato de correo no válido.");
      return;
    }
    
  }
  catch
  {
    Browser.msgBox("Correo clientes: Los correos deben ser cadenas de caracteres.");
    return;
  }
  

  //Líderes del proyecto para enviar correo de notificación
  try
  {
    var rango = sheet.getRange("B9:G10");
    var valores = rango.getValues();
    var lideres_proyectos = [];
    for (var i = 0; i < valores.length; i++) {
      for (var j = 0; j < valores[i].length; j++) {
        if(valores[i][j] !== ""){
          lideres_proyectos.push(valores[i][j].trim())
        }
      }
    }
    var lideres_proyectos = lideres_proyectos.join(',')
    //Verifica el formato del correo
    var boolFormato = verificacionMailFormato(lideres_proyectos)
    if(!boolFormato)
    {
      Browser.msgBox("Correo líderes: Formato de correo no válido.");
      return;
    }
  }
  catch
  {
    Browser.msgBox("Correo líderes: Los correos deben ser cadenas de caracteres.");
    return;
  }
  
  //Palabras claves del proyecto
  try
  {
    var rango = sheet.getRange("B11:G12");
    var valores = rango.getValues();
    var keywords = [];
    for (var i = 0; i < valores.length; i++) {
      for (var j = 0; j < valores[i].length; j++) {
        if(valores[i][j] !== ""){
          keywords.push(valores[i][j].trim())
        }
      }
    }
    var keywords = keywords.join(',')
  }
  catch
  {
    Browser.msgBox("Las palabras claves deben ser cadenas de texto.");
    return;
  }
  

  // Nombre etiqueta
  try
  {
    var etiquetaGmail = sheet.getRange("B13").getValue();
    var etiquetaGmail = etiquetaGmail.trim();
  }
  catch
  {
    Browser.msgBox("La etiqueta debe ser una cadena de caracteres.");
    return;
  }
  
  //Tipo de plazo de respuesta: días habiles o calendario
  var tipoPlazoRespuesta = sheet.getRange("C14").getValue(); 

  // Número de días para respuesta
  var numDiasRespuesta = sheet.getRange("E14").getValue(); 
  if (!Number.isInteger(numDiasRespuesta) && numDiasRespuesta !== "")
  {
    Browser.msgBox("Plazo de días de respuesta debe ser un número entero.");
    return;
  }

  // Booleano para tener en cuenta los sábados. Sólo para funciona para días habiles
  var incluye_sabado = sheet.getRange("G14").getValue(); 
  if (incluye_sabado === "Sí"){
    var incluye_sabado = true;
  }
  else{
    var incluye_sabado = false;
  }

  //ID de la unidad compartida para almacenar los archivos adjuntos
  try
  {
    var idFolder = sheet.getRange("B15").getValue();
    var idFolder = idFolder.trim();
  }
  catch
  {
    Browser.msgBox("ID de la carpeta debe ser una cadena de caracteres.");
    return;
  }
  

  // Festivos en Colombia que cubren el proyecto
  try
  {
    var festivos = sheet.getRange("D15").getDisplayValue()
    var festivos = festivos.split(',');
    var festivos = festivos.map(function(cadena) {
      return cadena.trim();
    });
    var festivos = festivos.join(',');

    //Verifica si cumple el formato deseado
    if (festivos !== "")
    {
      var festivos_temp = festivos.split(",")
      var regex = /^\d{4}-\d{2}-\d{2}$/
      for(var i = 0 ; i < festivos_temp.length ; i++)
      {
        if(!regex.test(festivos_temp[i]))
        {
          Browser.msgBox("Los festivos no cumplen con el formato aaaa-mm-dd.");
          return;
        }
      }
    }
  }
  catch
  {
    Browser.msgBox("Los festivos deben ser cadenas de caracteres.");
    return;
  }
  

  // Ajustar frecuencia para que trigger se active(horas)
  var hoursInterval =  sheet.getRange("F15").getValue(); 
  if (!Number.isInteger(hoursInterval) && hoursInterval !== "")
  {
    Browser.msgBox("El intervalo de horas debe ser un número entero.");
    return;
  }

  //ID de la hoja espejo
  try
  {
    var id_sheetMirror = sheet.getRange("B16").getValue();
    var id_sheetMirror = id_sheetMirror.trim();
  }
  catch
  {
    Browser.msgBox("ID de la planilla debe ser una cadena de caracteres.");
    return;
  }

  //Email de la persona que inicializa la hoja (debe ser la dueña)
  var owner =  sheet.getRange("B20").getValue(); 
  var owner = owner.trim();

  //Verifica si el formulario está diligenciado
  var boolDiligenciado = verificacionFormulario(lideres_proyectos,clientes,keywords,etiquetaGmail,tipoPlazoRespuesta,numDiasRespuesta,incluye_sabado,idFolder,festivos, hoursInterval,id_sheetMirror,owner);

  if(boolDiligenciado)
  { 
    //Verifica si el usuario activo es el mismo administrador de correspondencia
    var boolAdmin = verificacionAdmin(owner);

    if(boolAdmin)
    {

      // Llama al script como un objeto.
      var scriptProperties = PropertiesService.getScriptProperties();

      //Asigna las variables
      scriptProperties.setProperty('lideres_proyectos', lideres_proyectos);
      scriptProperties.setProperty('clientes', clientes);
      scriptProperties.setProperty('keywords', keywords);
      scriptProperties.setProperty('etiquetaGmail', etiquetaGmail);
      scriptProperties.setProperty('tipoPlazoRespuesta', tipoPlazoRespuesta);
      scriptProperties.setProperty('numDiasRespuesta', numDiasRespuesta);
      scriptProperties.setProperty('incluye_sabado', incluye_sabado);
      scriptProperties.setProperty('idFolder', idFolder);
      scriptProperties.setProperty('festivos', festivos);
      scriptProperties.setProperty('hoursInterval', hoursInterval);
      scriptProperties.setProperty('id_sheetMirror', id_sheetMirror);
      scriptProperties.setProperty('owner', owner);

      //Almacena una propiedad donde verifica que ya se almaceno la configuración
      scriptProperties.setProperty('save_configurations',true);

      //Escribe el correo para solicitar deploy del bot.
      if (Boolean(scriptProperties.getProperty('bool_servidor')) !== true)
      {
        var cod_proyecto =  sheet.getRange("G3").getDisplayValue(); 
        requestAuthServer(etiquetaGmail,cod_proyecto)
      }
      
      Browser.msgBox("Configuración almacenada con éxito.");
      return;

    }
    else
    {
      Browser.msgBox("No tiene permiso para guardar la configuración el Sistema de Registro de Correspondencia.");
      return;
    }
  }
  else
  {
    Browser.msgBox("Diligenciar completamente la Hoja de Inicio.")
    return;
  }

}


// Verifica si el formulario no está vacío.
function verificacionFormulario(lideres_proyectos,clientes,keywords,etiquetaGmail,tipoPlazoRespuesta,numDiasRespuesta,incluye_sabado,idFolder,festivos, hoursInterval,id_sheetMirror,owner)
{
  if(lideres_proyectos!=="" && clientes!==""&& keywords!==""&& etiquetaGmail!==""&& tipoPlazoRespuesta!==""&& numDiasRespuesta!==""&& incluye_sabado!==""&& idFolder!==""&&festivos!==""&& hoursInterval!=="" && id_sheetMirror!==""&& owner!=="")
  { 
    return true;
  }
  else{
    return false;
  }

}

//Verifica si el usuario activo es el administrador de correspondencia
function verificacionAdmin(owner)
{
  //Obtiene el usuario activo
  var userEmail =   Session.getActiveUser().getEmail();

  if (owner === userEmail)
  {
    return true;
  }
  else
  {
    return false;
  }

}

//Verifica formato de los correos
function verificacionMailFormato(correos)
{
  if (correos !== "")
  { 
    var correos_temp = correos.split(',')
    var regex = /^[\w-]+(?:\.[\w-]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7}$/;
    for (var i = 0 ; i < correos_temp.length; i++)
    {
      if(!regex.test(correos_temp[i]))
      {
        return false
      }
      
    }
    
  }
  return true;
  
}

//Envia correo para solicitar deploy
function requestAuthServer(etiqueta,cod_proyecto){

  // Enviar correo electrónico para solicitar deploy
  MailApp.sendEmail("administracion@empresa.com", "DEPLOY: "+etiqueta +" / "+ cod_proyecto, "");
  return;
    
}