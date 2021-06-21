/*************************************************************
 * Definición de variables globales                            *
 * ***********************************************************/

regex=/^([A-Z0-9]+)\w/g 
const re = /(.*Style|.*Name|.*Img|.*Subfamily|.*Size Group|.*Create)/
var collection=SpreadsheetApp.getActiveSpreadsheet().getName().match(regex)[0] // Només funciona si el nom del fitxer comença per la col·lecció.
var nomCarpetaTSV="Tech_sheets_generators";
var main_path="Volumes:GoogleDrive:Unidades compartidas:Product Design:Product Design "+collection+":" + nomCarpetaTSV + ":"; //Tech_sheets_generators:";
var comments_path=main_path+"comments_img:";
var product_path=main_path+"product_img:";
var color_path=main_path+"color_img:";
var type_background_path=main_path+"type_background_img:";

var ss = SpreadsheetApp.getActiveSpreadsheet();
var filename=collection+"_structure_indesign.tsv";

var key_filter="Create";
var key1="Reference";
var global_key1="Reference";
var key2="Pattern Ref.";
var key_filter_designer="Designer";

/*****************************************************************
 * Función onOpen                                                *  
 * Crea una entrada de menú en el Google sheet que lo llama.     *
 * En este caso: Designer Menu -> Export All                     *  
 * ***************************************************************/
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Designer Menu')
      .addItem('Export All', 'exportaTodos')
      .addToUi();
}


/***********************************************************************
 * Función exportDesigner()                                           *  
 * Abre una ventana de diálogo para seleccionar los valores del filtro *
 * y ejecutar la función Export_All                                    *  
 * *********************************************************************/
function exportDesigner()
{
  openDialog(getdesigners());
}

/****************************************************************************************************************************
 * Función exportaTodos                                                                                                       *  
 * Seleccionamos los valores a exportar a partir de los datos de la hoja principal. Después añadimos los campos necesarios  *
 * par la generación del fichero tsc. Una vez creado el fichero, eliminamos la nueva hoja.                                  *
 * Parámetros:                                                                                                              *
 *     - designers: Selección de valores para el filtro de datos.                                                           *
 ****************************************************************************************************************************/

function exportaTodos(designers) 
{
  // DECLARACIÓN DE VARIABLES
  // Definimos el nombre del sheet principal
  //ar tech_generator_folder=getParentFolder();
  var sheet_main=ss.getSheetByName("Structure");
  //Definimos las columnas extra para la cabecera (columnas a crear que no existen en la hoja)
  var h_extra=["@type_background_img","@comments_img1","@comments_img10","@comments_img11","@comments_img12","@comments_img13"]
  //Definimos los nombres de las columnas a crear que no existen en la hoja.
  var h_extra_keys=["Reference","comments_img1","comments_img10","comments_img11","comments_img12","comments_img13"]
   // Definimos el path de las columnas correspondientes a comments_img y type_background_img  
  var comments_path=main_path+"comments_img:";
  var type_background_path=main_path+"type_background_img:";

  //Creamos una nueva hoja y le damos nombre
  var newsheet= ss.insertSheet();
  Logger.log(main_path);
  newsheet.hideSheet();
  newsheet.setName('Export_'+Utilities.formatDate(new Date(), "GMT", "yy-MM-dd-HH:mm"));
  
  //Obtenemos el rango total de valores de la hoja principal
  var range=sheet_main.getRange(1,1,sheet_main.getLastRow(),sheet_main.getLastColumn());
  var data = range.getValues();
  
  //Filtramos los datos de acuerdo a las reglas de filtro (Create = YES y designers).
  data= filter(data,key_filter,["YES"])
  if(designers!=null){
    data= filter(data,key_filter_designer,designers)
  }

  // Creamos la cabecera de la nueva hoja de cálculo con los datos de la primera fila del sheet principal. 
  var headers=data[0];
  
  // Añadimos los campos de cabecera que no encontramos en la pestaña principal y que hemos construido previamente.
  newsheet.getRange(1,1,1,headers.concat(h_extra).length).setValues([headers.concat(h_extra)])
  
  // Recorremos todo el rango de datos
  for(var i = 1; i < data.length; i++) 
  {
    // obtenemos el valor del campo referencia (el array headers tiene todas las cabeceras) dentro de la fila i (con los datos en data)
    var reference=data[i][headers.indexOf(key1)] 
    l_extra=extra(sheet_main,h_extra_keys,reference) // añade los valores de las columnas extra
    newsheet.getRange(i+1,1,1,headers.concat(l_extra).length).setNumberFormat("@")
    newsheet.getRange(i+1,1,1,headers.concat(l_extra).length).setValues([data[i].concat(l_extra)])
  }
  // Eliminamos las columnas de la nueva hoja que no son necesarias para crear el fichero tsv 
  // (en un principio hemos copiado todas las existentes en la hoja principal)
  eliminaColumnaSheet(re,headers,newsheet);
  // Creamos un fichero tsv con los datos contenidos en el nuevo sheet
 
  var FitxerExportat = exportarDatosTsv (newsheet,filename,nomCarpetaTSV); //"Tech_sheets_generators");
  // Una vez creado el fichero, eliminamos la nueva hoja creada
  newsheet.activate()
  ss.deleteActiveSheet() 
  // Crea una instancia del tipo de objeto y tírala
  if (FitxerExportat==0) {
    Logger.log("La carpeta " + nomCarpetaTSV + " no existe. El fichero no se ha generado.");
    throw new UserException("La carpeta " + nomCarpetaTSV + " no existe. El fichero no se ha generado.");  
  }

}









