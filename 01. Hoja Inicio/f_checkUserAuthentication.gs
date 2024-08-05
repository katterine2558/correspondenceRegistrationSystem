/*
  CHEQUEA SI EL USUARIO ESTÁ AUTENTICADO Y HA AUTORIZADO EL SCRIPT
*/

function checkUserAuthentication(userEmail) {

  // Variables del proyecto
  var scriptProperties = PropertiesService.getScriptProperties();
  var id_sheetMirror = scriptProperties.getProperty('id_sheetMirror'); // id de la hoja espejo
  var owner = scriptProperties.getProperty('owner'); // propietario de la hoja espejo  

  try{

    //Lee el libro
    var book = SpreadsheetApp.openById(id_sheetMirror);
    var sheet = book.getSheetByName("Autorizaciones");

    //Lee los usuarios autorizados
    var authenticatedUsers = sheet.getRange("A1:A500").getValues();
    console.log("Se leyeron los usuarios autenticados")

    //Busca el usuario en la lista
    for (var i = 0; i < authenticatedUsers.length; i++) {
      var valorCelda = authenticatedUsers[i][0];
      if (valorCelda === userEmail) {
        return true;
      }
    }
    return false;
    
  }catch(error){
    console.log(error)
    Browser.msgBox('Error', 'Solicitar permiso como LECTOR en la planilla de registro a ' + owner, Browser.Buttons.OK);
    return;
  }


  
  

}

  