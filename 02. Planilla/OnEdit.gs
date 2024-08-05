function onEdit(e) {

  if (e.source.getSheetName() === "Entrada") {

    //Revisa si la edición se hizo en la primera columna
    if (e.range.getColumn() === 1){
      
      //Obtiene la hoja
      var sheet = e.source.getSheetByName("Entrada");

      //Extrae los parámetros que necesita
      var scriptProperties = PropertiesService.getScriptProperties();
      var numDiasRespuesta = scriptProperties.getProperty('numDiasRespuesta');  // Número de días para la respuesta
      var numDiasRespuesta = parseFloat(numDiasRespuesta);
      var incluye_sabado = scriptProperties.getProperty('incluye_sabado');  // Booleano para tener en cuenta los sábados. Sólo para funciona para días habiles
      var incluye_sabado = JSON.parse(incluye_sabado)
      var festivos = scriptProperties.getProperty('festivos'); // Festivos 
      var festivos = festivos.split(',');
      var tipoPlazoRespuesta = scriptProperties.getProperty('tipoPlazoRespuesta'); //Tipo de plazo de respuesta: días habiles o calendario

      //Llama la función que elimina la fecha de respuesta si la columna A es "No aplica", de lo contrario la vuelve a calcular
      modifyResponseDate(e,sheet,numDiasRespuesta,incluye_sabado,festivos,tipoPlazoRespuesta)

    }
    else{
      console.log("Se hizo en otra columna")
    }
    
    
  }
}

function modifyResponseDate(event,hoja,numDiasHabiles,diaHabilNormal,festivos,tipoPlazoRespuesta){

  console.log("entró a la función")

  var fila = event.range.getRow();
  var valorColumnaA = event.value;

  // Si el valor cambia a "No aplica", borra la fila correspondiente en la columna J
  if (valorColumnaA === 'No aplica' || valorColumnaA === 'Contestado') {
    
    hoja.getRange(fila, 11).setValue('');
    hoja.getRange(fila, 12).setValue('');
  }
  else if(valorColumnaA === 'Sin contestar'){

    var fechaEnvio = hoja.getRange(fila, 2).getDisplayValue();
    var fechaRespuesta = getResponseDate(fechaEnvio,numDiasHabiles,diaHabilNormal,festivos,tipoPlazoRespuesta)
    var plazoDias = estimateResponseTime(fechaRespuesta,diaHabilNormal,festivos,tipoPlazoRespuesta,numDiasHabiles)
    hoja.getRange(fila, 11).setValue(fechaRespuesta);
    hoja.getRange(fila, 12).setValue(plazoDias);

  }

}
