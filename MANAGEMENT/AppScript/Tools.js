function filter(rangeval,key_field,pass_val){
  var vals=rangeval
  var dest=[] ;
  dest.push(vals[0]);
  for (var i = 0; i < vals.length; i++) if (pass_val.indexOf(vals[i][vals[0].indexOf(key_field)])!=-1) dest.push(vals[i]);
  return dest;
}

/***********************************************************************
 * Función exportarDatosTsv                                            *  
 * Crea un fitxero de texto con los datos del Sheet pSheet.            *
 * Parámetros:                                                         *
 *     - pSheet: Hoja de cálculo de la que se exportarán los datos.    *
 *     - pNomFichero: Nombre del fichero de salida.                    *     
 *     - pNomCarpeta: Nombre de la carpeta donde se grabará el fichero *
 * *********************************************************************/

function exportarDatosTsv(pSheet,pNomFichero,pNomCarpeta)
{
//Export data to tsv and store
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
    if (tech_generator_folder == null) 
         return -2;
    else if (txtFile == null)
       return -1;
    else {
      DriveApp.getFolderById(tech_generator_folder).createFile(pNomFichero,txtFile);
      return 0;  
    }
}


function getParentFolder(pNomCarpeta)
{
  try
  { 
    var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var folders = file.getParents().next().getParents().next().getFolders();
    while (folders.hasNext()){
      folder=folders.next()
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