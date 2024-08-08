// Esta función verifica que los correos con estado "Sin contestar", estén marcados de esta forma únicamente para aquellos cuyos destinatarios hayan aprobado el SRC. Ejemplo, correos con temas de polizas (operaciones) y/o comercial
function checkToUsersEmail() {

  //Obtiene los correos que han autorizado el SRC
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName("Autorizaciones");
  var autorizados = sheet.getRange("A1:A");
  var autorizados = autorizados.getDisplayValues();
  var autorizados = autorizados.filter(function(subArray) {
    return subArray[0] !== "";
  });
  var lista_autorizados = []
  for (var i=0;i<autorizados.length;i++){
    lista_autorizados.push(autorizados[i][0]);
  }

  //Obtiene la hoja
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetByName("Entrada");

  //Obtiene el estado de los correos "Sin Contestar"
  var columnaA = sheet.getRange("A2:A");
  var columnaA = columnaA.getDisplayValues();
  var columnaA = columnaA.filter(function(subArray) {
    return subArray[0] !== "";
  });

  //Itera por cada registro
  for (var i =0; i<columnaA.length;i++){
    if (columnaA[i][0] === "Sin contestar"){

      // Encuentra los destinatarios
      var destinatarios = sheet.getRange("F"+(i+2)).getValue();
      //Vuelve un array los destinatarios
      var destinatarios = destinatarios.split(", ");

      //Verifica que dentro de los destinatarios NO estén en la lista de los autorizados
      var hayCoincidencia = destinatarios.some(function(elemento) {
        return lista_autorizados.includes(elemento);
      });

      // Cambia estado a "No Aplica"
      if(!hayCoincidencia){
        sheet.getRange("A"+(i+2)).setValue("No aplica")
      }

      
    }
  }

  
  
}
