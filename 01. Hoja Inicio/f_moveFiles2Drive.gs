/*
MUEVE LOS ARCHIVOS A DRIVE
*/
function moveFilesToDrive(attachedFiles, nameFolder, idFolder, emailBody, idEmail,tipo) {
  
  // Almacena los archivos
  if (attachedFiles.length >= 1){

    for (var k = 0; k < attachedFiles.length; k++) {

      // Nombre del archivo
      fileName = attachedFiles[k].getName();
      
      //Verifica si el archivo ya existe en la carpeta
      checkFile = findFile(fileName,idFolder,nameFolder)

      //Intenta obtener la carpeta
      var folder = checkExistingFolder(idFolder,nameFolder)

      if (checkFile === false){
        // Mueve el archivo adjunto a la carpeta destino
        folder.createFile(attachedFiles[k]);
        console.log("El archivo " + fileName+ " fue almacenado en la carpeta.")
      }
      else{
        console.log("El archivo " + fileName+ " ya existe")
      }
      
    }

  }
  //Se añade el cuerpo del mensaje en un .txt
  writeEmailBody(emailBody,idEmail,tipo,idFolder,nameFolder)

}

// Verifica si el archivo existe

function findFile(fileName,idFolder,nameFolder) {

  //Intenta obtener la carpeta
  var folder = checkExistingFolder(idFolder,nameFolder)

  if (folder === undefined)
  {
    console.log("la carpeta es undefined")
    var switch_bool = false
    while (switch_bool === false)
    {
      //Intenta obtener la carpeta
      var folder = checkExistingFolder(idFolder,nameFolder)
    
      if(folder !== undefined)
      {
        console.log("la carpeta NO es undefined")
        switch_bool = true
      }
      else
      {
        console.log("la carpeta sigue siendo undefined")
      }
      
    }
    
  }
  
  // Obtener la carpeta por su ID
  var carpeta = DriveApp.getFolderById(folder.getId());

  // Buscar archivos por nombre en la carpeta
  var archivos = carpeta.getFilesByName(fileName);
  
  // Iterar sobre los archivos para verificar si el nombre del archivo existe
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    if (archivo.getName() === fileName) {
      return true; // El archivo existe
    }
  }

  return false; // El archivo no existe
}

//Escribe el cuerpo del email
function writeEmailBody(emailBody,idEmail,tipo,idFolder,nameFolder) {

  //Nombre del archivo del body
  var bodyFileName =  tipo + "_Body-"+idEmail+".txt"
  //Encuentra si el body ya está escrito
  var checkFile = findFile(bodyFileName,idFolder,nameFolder)

  //Folder
  var folder = checkExistingFolder(idFolder,nameFolder);
  if (checkFile === false){

    // Crear un nuevo archivo de texto
    var archivoTxt = folder.createFile(bodyFileName, emailBody);

  }
  
}

//Verifica si la carpeta existe y la retorna
function checkExistingFolder(idFolder,nameFolder)
{
  //Obtiene la carpeta para almacenar
  var unidadCompartida = DriveApp.getFolderById(idFolder);

  // Obtiene todas las carpetas dentro del padre especificado
  var carpetas = unidadCompartida.getFolders();

  // Itera a través de las carpetas para encontrar la que coincida con el nombre
  while (carpetas.hasNext()) {
    var carpeta = carpetas.next();
    if (carpeta.getName() === nameFolder) {
      var folder = carpeta; 
    }
  }

  return folder
}