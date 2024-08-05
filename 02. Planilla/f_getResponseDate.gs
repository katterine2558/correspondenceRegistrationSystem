
// Función para calcular la nueva fecha después de ciertos días hábiles incluyendo sábados y excluyendo festivos
function getResponseDate(emailDate,numDiasRespuesta,incluye_sabado,festivos,tipoPlazoRespuesta) {

  // Divide la cadena de texto en día, mes y año
  var partesFecha = emailDate.split("/");
  
  // Crea un objeto de fecha utilizando las partes extraídas
  var fecha = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]);


  //Verifica el tipo de plazo de respuesta
  if (tipoPlazoRespuesta === "Calendario"){

    while (numDiasRespuesta > 0){

      fecha.setDate(fecha.getDate() + 1);
      numDiasRespuesta--;

    }

  }
  else{

    // Itera hasta alcanzar el número de días hábiles deseados
    while (numDiasRespuesta > 0) {
      
      fecha.setDate(fecha.getDate() + 1);
      
      // Verifica si el día actual no es un festivo (día no hábil)
      var esFestivo = festivos.some(function(festivo) {
      return (
          Utilities.formatDate(fecha, 'GMT', 'yyyy-MM-dd') === festivo
        );
      });

      if (incluye_sabado===true){
        if (!esFestivo && fecha.getUTCDay() !== 0) {
          numDiasRespuesta--;
        }
      }
      else{
        if (!esFestivo && fecha.getUTCDay() !== 0 && fecha.getUTCDay() !== 6 ) {
          numDiasRespuesta--;
        }
      }
    }

  }
  
  // Formatea la fecha a un formato legible
  var formatoFecha = Utilities.formatDate(fecha, 'GMT', 'yyyy/MM/dd');
  console.log("Se calculó la fecha máxima de respuesta.")
  
  return formatoFecha;
  
}
