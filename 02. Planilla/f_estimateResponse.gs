function estimateResponseTime(fechaRespuesta,incluye_sabado,festivos,tipoPlazoRespuesta,numDiasRespuesta) {
    
  //Fecha actual
  //var fechaHoy = new Date("2024","07"-1,"19");
  var fechaHoy = new Date();
  fechaHoy.setHours(0,0,0,0);

  //Fecha máxima de respuesta
  var partesFecha = fechaRespuesta.split('/'); // Dividir el string en día, mes y año
  var fechaMaxima = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]); 

  //Genera una lista de fechas entre la fecha de hoy y la fecha máxima de respuesta
  var listaFechas = [];
  var fechaActual = new Date(fechaHoy);
  while (fechaActual <= fechaMaxima) {
    listaFechas.push(new Date(fechaActual)); // Agregar la fecha a la lista
    fechaActual.setDate(fechaActual.getDate() + 1); // Avanzar al siguiente día
  }

  //Filtra la lista de fechas según el tipo de plazo de respuesta.
  var listaFiltrada = [];

  if (tipoPlazoRespuesta === "Hábiles"){
  
    for (var i = 0; i < listaFechas.length; i++) {

      //Fecha en iteración
      var fechaActual = listaFechas[i];
      // Verifica si el día actual no es un festivo (día no hábil)
      var esFestivo = festivos.some(function(festivo) {
      return (
        Utilities.formatDate(fechaActual, 'GMT', 'yyyy-MM-dd') === festivo
        );
      });

      if (incluye_sabado===false){

        if (fechaActual.getUTCDay() !== 0 && fechaActual.getUTCDay() !== 6 && !esFestivo) {
          listaFiltrada.push(fechaActual); // Agregar la fecha a la lista filtrada
        }
      }
      else{

        // Verificar si el día de la semana no es domingo (0) ni festivo
        if (fechaActual.getUTCDay() !== 0 && !esFestivo) {
          listaFiltrada.push(fechaActual); // Agregar la fecha a la lista filtrada
          }
        }

    }

    //Retorna el plazo en día que hay para responder correspondencia
    if (listaFiltrada.length > numDiasRespuesta){
      return numDiasRespuesta;
    }
    else{
      return listaFiltrada.length;
    }
  }
  else{

    //Retorna el plazo en día que hay para responder correspondencia
    if(listaFechas.length > numDiasRespuesta){
      return numDiasRespuesta;
    }
    else{
      return listaFechas.length;
    }

    
  }
}
