/**
 * Recebe o id de um arquivo salvo no gogle drive e retorna seu conteúdo como string (base64)
 * 
 * @param {String} id_File - O id de um arquivo salvo no google drive
 */
function Files_DataAsString(id_File){
  return Utilities.base64Encode(DriveApp.getFileById(id_File).getBlob().getBytes())
}
