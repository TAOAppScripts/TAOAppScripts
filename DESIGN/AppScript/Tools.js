function filter(rangeval,key_field,pass_val){
  var vals=rangeval
  var dest=[] ;
  dest.push(vals[0]);
  for (var i = 0; i < vals.length; i++) if (pass_val.indexOf(vals[i][vals[0].indexOf(key_field)])!=-1) dest.push(vals[i]);
  return dest;
}

/*************************************************************************************************
 * Función comparaNumeros                                                                        *
 * Utilidad para la eliminación de columnas (al eliminar una columna se deben reubicar el resto).*
 * Parámetros:                                                                                   *
 *      - a y b: números de columnas a recolocar.                                                *
 * Retorna:                                                                                      *
 *      - La nueva posición                                                                      *
 *************************************************************************************************/
function comparaNumeros(a, b) {
  return a - b;
}

/****************************************************************************************************************
* Función eliminaColumnaSheet                                                                                   *  
* Elimina las columnas definidas según la expresión regular pRegExp y contenidas en el Sheet pSheet según       *
* la definición de la cabecera definida en headers                                                              *
* Parámetros:                                                                                                   *
*      - pRegExp: Expresión regular con los nombres de las columnas a eliminar                                  *
*      - pHeaders: Nombre de las columnas del Sheet                                                             *
*      - pSheet: Hoja de cálculo sobre la que se eliminan las columnas.                                         *
 ****************************************************************************************************************/
function eliminaColumnaSheet(pRegExp,pHeaders,pSheet)
{
  var colA = [];   
  for(var j = 0; j<pHeaders.length;j++)
  {
    if (pHeaders[j].match(pRegExp) && colA.indexOf(j+1)==-1)
    {
      colA.push(j+1);
    }
  }
  colA.sort(comparaNumeros);
  colA.reverse();
  for(var i =0;i < colA.length; i++)
  {
    pSheet.deleteColumn(colA[i]);
  }
}

/***********************************************************************
 * Función exportarDatosTsv                                            *  
 * Crea un fitxero de texto con los datos del Sheet pSheet.            *
 * Parámetros:                                                         *
 *     - pSheet: Hoja de cálculo de la que se exportarán los datos.    *
 *     - pNomFichero: Nombre del fichero de salida.                    *     
 *     - pNomCarpeta: Nombre de la carpeta donde se grabará el fichero *
 * *********************************************************************/

function exportarDatosTsv (pSheet,pNomFichero, pNomCarpeta)
{
//Export data to tsv and store
  Logger.log(pNomCarpeta);
  var range=pSheet.getRange(1,1,pSheet.getLastRow(),pSheet.getLastColumn());
  var data = range.getValues();
    // Loop through the data in the range and build a string with the data
    if (data.length > 1) {
      var txtFile = "";
      for (var row = 0; row < data.length; row++) {
        // Join each row's columns and add a carriage return to end of each row
        txtFile += "\""+data[row].join("\"\t\"").replace(/\n/g,"<br>") + "\"\r\n";
      }
    };

      // Create a file in the Docs List with the given name and the data
    var tech_generator_folder=getParentFolder(pNomCarpeta);
    //var tech_generator_folder=getParentFolderPorNombre(pNomCarpeta);
    //var tech_generator_folder= getDriveFolder(main_path)

     if (tech_generator_folder == null) 
    {
         Logger.log("La carpeta " + pNomCarpeta + " no existe. El fichero no se ha generado.")
         return 0;
    }
    else {
      DriveApp.getFolderById(tech_generator_folder).createFile(pNomFichero,txtFile);
      return 1;  
    }
    
} 
/**************************************************************************************
 * Función getParentFolderPorNombre                                                   *  
 * Obtención del ID de la carpeta ubicada en drive, a partir del nombre de la misma.  *
 * Parámetros:                                                                        *
 *     - pNomCarpeta: Nombre de la carpeta a buscar.                                  *
 * Retorna: Id de la carpeta obtenida                                                 *  
 * ************************************************************************************/
function getParentFolderPorNombre(pNomCarpeta)
{
    try
    {
          var idCarpeta = null;
          var folder = DriveApp.getFoldersByName(pNomCarpeta);
          if(folder.hasNext()) {
              idCarpeta = folder.next().getId();
          } else {
              Logger.log('La carpeta ' + pNomCarpeta + ' no existe en el Drive!');
          }
          return idCarpeta;
    }
    catch(e)
    {
      Logger.log(e);  
      return null;
    }
}

/**********************************************************************************************
 * Función getParentFolder                                                                    *  
 * Obtención del ID de la carpeta ubicada en drive, a partir de la ubicación del contenedor.  *
 * Parámetros:                                                                                *
 *     - pNomCarpeta: Nombre de la carpeta a buscar.                                          *
 * Retorna: Id de la carpeta obtenida                                                         *  
 * ********************************************************************************************/
function getParentFolder(pNomCarpeta)
{
  try
  { 
    var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var folders = file.getParents().next().getParents().next().getFolders();
    while (folders.hasNext()){
      folder=folders.next()
      Logger.log(folder.getName())
      if (folder.getName()==pNomCarpeta){
        return folder.getId()
      }
    }
  }
  catch(e)
  {
      Logger.log(e);  
      return null;
  }
  return null;
}

/**********************************************************************************************
 * Función getDriveFolder                                                                    *  
 * Obtención del ID de la carpeta ubicada en drive, a partir del path indicado por parámetros *
 * Si las carpetas no existen, se crean en el Drive.
 * Parámetros:                                                                                *
 *     - pPath: Ubicación de la carpeta.                                                *
 * Retorna: Id de la carpeta obtenida                                                         *  
 * ********************************************************************************************/
function getDriveFolder(pPath) {
    ////Proves
//var pPath ="Be-Wiki:Formació:NOVA:"
var name, folder, search, fullpath;
pPath = pPath.endsWith(":") ? pPath = pPath.substring(0,pPath.length-1) : pPath
  // Elimina els / de más y crea un array con cada una de las carpetas que forman el path
  fullpath = pPath.replace(/^\/*|\/*$/g, '').replace(/^\s*|\s*$/g, '').split(":");

  // Iniciamos el proceso en la raiz del Drive
  folder = DriveApp.getRootFolder();

  for (var subfolder in fullpath) {
  name = fullpath[subfolder];
    search = folder.getFoldersByName(name);
    // Si la carpeta no existeis, es crea al mateix nivell
    folder = search.hasNext() ? search.next() :folder.createFolder(name);
  }
  return folder.getId();
}


/*********************************************************************  
* Función UserException                                              *
* Crea un objeto de tipo UserException.                              * 
* Parámetros:                                                        *
*  - message: Mensaje asociado a la excepción.                       *                                                  **********************************************************************/

function UserException(message) {
  this.message = message;
  this.name = 'UserException';
}

/*********************************************************************  
* Hace que la excepción se muestre como una cadena con formato       *
**********************************************************************/

UserException.prototype.toString = function() {
  return `${this.name}: "${this.message}"`;
}