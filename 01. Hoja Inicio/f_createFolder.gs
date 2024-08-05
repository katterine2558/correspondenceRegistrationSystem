/*
CREA LA CARPETA PARA ALMACENAR LOS ARCHIVOS ADJUNTOS
*/
function createFolder(nameFolder,idUnidadCompartida) {

  // Obtiene la unidad compartida por su ID
  var unidadCompartida = DriveApp.getFolderById(idUnidadCompartida);

  //Chequea si existe una carpeta con el mismo nombre y devuelve la carpeta
  var folder = findFolder(unidadCompartida,nameFolder)

  if (folder){
    console.log("Ya existe la carpeta con nombre "+ nameFolder)
    return folder;
  }
  else{
    // Crea la carpeta en la unidad compartida
    var nuevaCarpeta = unidadCompartida.createFolder(nameFolder);
    console.log("Se creó la carpeta "+ nameFolder)
    return nuevaCarpeta;
  }
}

function findFolder(unidadCompartida,nameFolder)
{

  // Obtiene todas las carpetas dentro del padre especificado
  var carpetas = unidadCompartida.getFolders();

  // Itera a través de las carpetas para encontrar la que coincida con el nombre
  while (carpetas.hasNext()) {
    var carpeta = carpetas.next();
    if (carpeta.getName() === nameFolder) {
      return carpeta; // Devuelve la carpeta si se encuentra
    }
  }

  return null; // Devuelve null si la carpeta no se encuentra
}