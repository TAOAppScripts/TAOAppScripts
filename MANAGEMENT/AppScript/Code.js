function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Designer Menu')
      .addItem('Export All', 'Export_All')
      .addToUi();
   SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Production Menu')
      .addItem('Create_BOM', 'Create_BOMext')
      // Linea agregada por Daniel Prado - 4 dic 2020. Num001
      .addItem('Refresh_Pricing','Refrescar')
      // Fin de linea Num001
      .addToUi();
}

//Global parameters and sheet pointers
regex=/^([A-Z0-9]+)\w/g 

var key1="Reference";
var global_key1="Reference";
var key2="Pattern Ref.";
var key_filter="Create";
var key_filter_designer="Designer";

// Definimos los nombres del fichero a generar y de la carpeta donde será ubicado
var collection=SpreadsheetApp.getActiveSpreadsheet().getName().match(regex)[0]
var nomCarpetaTSV="Tech_sheets_generators";
var filename=collection+"_structure_indesign.tsv";

// Definimos las rutas de acceso que se utilizarán para indicar ubicación imágenes y archivos en el fichero TSV.
var main_path="Volumes:GoogleDrive:Unidades compartidas:Product Design:Product Design "+collection+ ":" + nomCarpetaTSV + ":";
var graphics_path=main_path+"graphics_img:";
var product_path=main_path+"product_img:";
var color_path=main_path+"color_img:";
var materials_path=main_path+"materials_img:";
var graphics_pos_path=main_path+"graphics_pos_img:";
var comments_path=main_path+"comments_img:";
var labels_pos_path=main_path+"labels_pos_img:";
var type_background_path=main_path+"type_background_img:";

/***********************************************************************
 * Función exportDesigner()                                           *  
 * Abre una ventana de diálogo para seleccionar los valores del filtro *
 * y ejecutar la función Export_All                                    *  
 * *********************************************************************/
function Export_Designer()
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
//Export function
function Export_All(designers) {

  const vColumnasBom = ["","@"]
  const vColumnasBol = ["","@","@pos_"]
  const vColumnasBog = ["","@","@pos_","sizes_","sizes_print_"]

  // Obtenemos las pestañas de la hoja de cálculo activa (ss).
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_main=ss.getSheetByName("Main");
  var sheet_bom=ss.getSheetByName("BOM");
  var sheet_bol=ss.getSheetByName("BOL");
  var sheet_bog=ss.getSheetByName("BOG");

  //var tech_generator_folder=getParentFolder()
   
  //Create new sheet
  var newsheet= ss.insertSheet();
  var date = Utilities.formatDate(new Date(), "GMT", "yy-MM-dd-HH:mm"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
  newsheet.setName('Export_'+date)
  
  //Definimos cabeceras y claves extras a añadir al tsv
  var h_extra=["@color","@colorprint","@product_img","@comments_img1","@comments_img2"]
  var h_extra_keys=["Color Num.","Color Ref.","Reference","Comments1","Comments2"]
 
  //var h_bom=bhead(20,"item")
  //var h_bol=bhead_l(10,"label")
  //var h_bog=bhead_g(10,"graph")

  var h_bom=auxCabeceras(20,"item",vColumnasBom);
  var h_bol=auxCabeceras(10,"label",vColumnasBol);
  var h_bog=auxCabeceras(10,"graph",vColumnasBog);
  
  //Get Main Range
  var range=sheet_main.getRange(1,1,sheet_main.getLastRow(),sheet_main.getLastColumn());
  var data = range.getValues();
  
  //Filter Data according to filter rules (Create and designer).
  data= filter(data,key_filter,["YES"])
  if(designers!=null){
    data= filter(data,key_filter_designer,designers)
  }
  
  //Paste data on newsheet along with extra data from different bills/lists
  var counter=1;
  var headers=data[0];
  Logger.log(headers)
  newsheet.getRange(counter,1,1,headers.concat("@type_background_img").concat(h_extra).concat(h_bog).concat(h_bom).concat(h_bol).length).setValues([headers.concat("@type_background_img").concat(h_extra).concat(h_bog).concat(h_bom).concat(h_bol)])
  counter+=1
  
  for(var i = 1; i < data.length; i++) 
  {
    var type_bg=get_type_bg(data[i][headers.indexOf('Type')])
    var reference=data[i][headers.indexOf(key1)]
    var pattern=data[i][headers.indexOf(key2)]

    l_extra=extra(sheet_main,h_extra_keys,reference)
    l_bom=Bom(sheet_bom,pattern,reference)
    l_bom=auxTrf(l_bom,h_bom.length,"")
    l_bol=Bol(sheet_bol,pattern,reference)
    l_bol=auxTrf(l_bol,h_bol.length,"")
    l_bog=Bog(sheet_bog,pattern,reference)
    l_bog=auxTrf(l_bog,h_bog.length,"")
    newsheet.getRange(counter,1,1,headers.concat(type_bg).concat(l_extra).concat(h_bog).concat(h_bom).concat(h_bol).  length).setNumberFormat("@")
    newsheet.getRange(counter,1,1,headers.concat(type_bg).concat(l_extra).concat(h_bog).concat(h_bom).concat(h_bol).length).setValues([data[i].concat(type_bg).concat(l_extra).concat(l_bog).concat(l_bom).concat(l_bol)])
    counter+=1; 
  } 
  
  // Creamos un fichero tsv con los datos contenidos en el nuevo sheet
  var FitxerExportat = exportarDatosTsv(newsheet,filename,nomCarpetaTSV); 
  // Eliminamos la pestaña con los datos a exportar.
  newsheet.activate()
  ss.deleteActiveSheet() 
  
  var missatge = "";
  switch (FitxerExportat) {
  case -1:
    missatge = "No existen registros coincidentes. El fichero " + filename + " no se ha generado."
    console.warn(missatge);
    throw new UserException(missatge);  
  case -2:
    missatge = "La carpeta " + nomCarpetaTSV + " no existe dentro de la estructura. El fichero " + filename + "no se ha podido generar." 
    console.warn(missatge);
    throw new UserException(missatge);  
  case 0:
    missatge = "El fichero " + filename + " se ha creado en la carpeta " + nomCarpetaTSV;
    console.info(missatge);
   }
   
  }