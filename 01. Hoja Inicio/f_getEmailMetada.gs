/*
OBTIENE LOS METADATOS DEL CORREO
*/
function getEmailMetadata(mensaje) {

  // Formatear fecha
  var fecha = mensaje.getDate();
  var fechaFormateada = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  var horaFormateada = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm:ss');

  //Obtener remitente
  var remitente = extraerRemitente(mensaje.getFrom()); 

  //Obtener destinatario
  var destinatarios = extraerDestinarios(mensaje.getTo());

  //Obtener asunto
  var asunto =  eliminarPrefijoRe(mensaje.getSubject());

  // Obtener archivos adjuntos
  var archivosAdjuntos = mensaje.getAttachments();

  //Obtener cuerpo de mensaje
  var cuerpoMensaje = mensaje.getPlainBody();

  return {
    Date: fechaFormateada,
    Hour: horaFormateada,
    From: remitente,
    To: destinatarios,
    Subject: asunto,
    Body: cuerpoMensaje,
    AttachedFiles: archivosAdjuntos
  };

}

// Función para obtener todas las direcciones de correo electrónico de los destinatarios entre <>
function extraerDestinarios(destinatarios) {
  var regex = /<([^>]+)>/g;
  var destinatariosExtraidos = [];
  var match;
  
  while ((match = regex.exec(destinatarios)) !== null) {
    destinatariosExtraidos.push(match[1]);
  }
  
  return destinatariosExtraidos.length > 0 ? destinatariosExtraidos.join(', ') : destinatarios;
}

// Función para obtener la dirección de correo electrónico de remitente
function extraerRemitente(remitente) {
  var regex = /<([^>]+)>/g;
  var remitenteExtraido = [];
  var match;
  
  while ((match = regex.exec(remitente)) !== null) {
    remitenteExtraido.push(match[1]);
  }
  
  return remitenteExtraido.length > 0 ? remitenteExtraido.join(', ') : remitente;
}

// Función para eliminar el prefijo "Re:" del asunto
function eliminarPrefijoRe(asunto) {
  return asunto.replace(/^Re:\s*/, ''); // Elimina "Re:" al principio del asunto, seguido de cualquier espacio en blanco
}

// Función para obtener los nombres de los archivos adjuntos
function obtenerNombresArchivosAdjuntos(archivosAdjuntos) {
  var nombresArchivos = [];
  
  archivosAdjuntos.forEach(function(archivoAdjunto) {
    nombresArchivos.push(archivoAdjunto.getName());
  });

  return nombresArchivos.join(', ');
}
