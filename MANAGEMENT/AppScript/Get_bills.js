function test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_main=ss.getSheetByName("Main");
  var sheet_bom=ss.getSheetByName("BOM");
  var sheet_bol=ss.getSheetByName("BOL");
  var sheet_bog=ss.getSheetByName("BOG");
  
  h_bom=bhead(20,"item")
  l_bom=Bom(sheet_bom,"P19001","000972_181_XX")
  l_bom=bom_trf(l_bom,h_bom.length,"")
  
  h_bol=bhead_l(10,"label")
  l_bol=Bol(sheet_bol,"P19001","000972_181_XX")
  l_bol=bol_trf(l_bol,h_bol.length,"")
  


  h_bog=bhead_l(10,"graph")
  l_bog=Bog(sheet_bog,"P19001","000972_181_XX")
  l_bog=bol_trf(l_bog,h_bog.length,"")
  
  
  //Logger.log(Bol(sheet_bol,"P19001","000972_181_XX"))
  //Logger.log(Bog(sheet_bog,"P19001","000972_181_XX"))
  //Logger.log(bhead(20,"item"))
}
//new product(uniques[i])

function get_type_bg(type_key) {
  return type_background_path+type_key+".jpg";
}

function extra(sheet_b,headers,key1) {
  var result=[];
  var range=sheet_b.getRange(1,1,sheet_b.getLastRow(),sheet_b.getLastColumn());
  var data = range.getValues();
  var index = data[0].indexOf("Reference");
  for(var i = 0; i < data.length; i++)   
  {
    if (data[i][index]==key1){
      for(var ii=0;ii<headers.length;ii++){
        
        switch(headers[ii]) {
          case "Color Num.":
            dtemp=color_path+data[i][data[0].indexOf(headers[ii])]+".jpg"
            break;
          case "Color Ref.":
            dtemp=color_path+data[i][data[0].indexOf(headers[ii])]+".jpg"
            break;
          case "Reference":
            dtemp=product_path+data[i][data[0].indexOf(headers[ii])]+".jpg"
            break;
          case "Comments1":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-1.jpg"
            break;
          case "Comments2":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-2.jpg"
            break;
          default:
            dtemp=product_path+data[i][data[0].indexOf(headers[ii])]+".jpg"
        }
        result.push(dtemp)
      }
     break;
    }
  }

  return result;
}


function Bom(sheet_b,key1,key2) {
  var result=[];
  var range=sheet_b.getRange(1,1,sheet_b.getLastRow(),sheet_b.getLastColumn());
  var data = range.getValues();
  var index = data[0].indexOf("Pattern|Code");
  var regex=/^([A-Z0-9a-z]+)/g
  for(var i = 0; i < data.length; i++) 
  {
    if (data[i][index]==key1||data[i][index]==key2){
      dtemp=[data[i][5]+data[i][2]+": "+data[i][1]+" "+data[i][3]+data[i][4],materials_path+data[i][1].match(regex)[0]+".jpg"]
    result.push(dtemp);
  } 
  }

  return result;
}

function Bol(sheet_b,key1,key2) {
  var result=[];
  var range=sheet_b.getRange(1,1,sheet_b.getLastRow(),sheet_b.getLastColumn());
  var data = range.getValues();
  var index = data[0].indexOf("Pattern|Code");
  var regex=/^([A-Z0-9a-z]+)/g
  for(var i = 0; i < data.length; i++) 
  {
    if (data[i][index]==key1||data[i][index]==key2){
      dtemp=[data[i][6]+data[i][2]+": "+data[i][1]+" "+data[i][3]+data[i][4]+" Position: "+data[i][5],materials_path+data[i][1].match(regex)[0]+".jpg",labels_pos_path+data[i][5]+".jpg"]
      result.push(dtemp);
    } 
  }
    return result;
}

function Bog(sheet_b,key1,key2) {
  var result=[];
  var range=sheet_b.getRange(1,1,sheet_b.getLastRow(),sheet_b.getLastColumn());
  var data = range.getValues();
  var index = data[0].indexOf("Pattern|Code");
  var regex=/^([A-Z0-9a-z_]+)/g
  
  for(var i = 0; i < data.length; i++) 
  {
    if (data[i][index]==key1||data[i][index]==key2){
      dtemp=[data[i][4]+data[i][2]+": "+data[i][1]+" "+data[i][3],graphics_path+data[i][1].match(regex)[0]+".jpg",graphics_pos_path+data[i][0]+"_print_position.jpg",data[i][15], data[i][16]] // AFegim la nova columna "ALL_SIZE_PRINT"
    result.push(dtemp);
       
    } 
  }
   return result;

}

function setValoresSheet(sheet_b,key1,key2) {
  var result=[];
  var range=sheet_b.getRange(1,1,sheet_b.getLastRow(),sheet_b.getLastColumn());
  var data = range.getValues();
  var index = data[0].indexOf("Pattern|Code");
  var regex=/^([A-Z0-9a-z]+)/g
  Logger.log("Nom sheet des de Bol: " + sheet_b.getName());
  var nomSheet = sheet_b.getName();
  for(var i = 0; i < data.length; i++) 
  {
    if (data[i][index]==key1||data[i][index]==key2){
      switch (nomSheet) {
        case "BOM":
          dtemp=[data[i][5]+data[i][2]+": "+data[i][1]+" "+data[i][3]+data[i][4],materials_path+data[i][1].match(regex)[0]+".jpg"]  
        case "BOL":
          dtemp=[data[i][6]+data[i][2]+": "+data[i][1]+" "+data[i][3]+data[i][4]+" Position: "+data[i][5],materials_path+data[i][1].match(regex)[0]+".jpg", labels_pos_path+data[i][5]+".jpg"]  
        case "BOG":
          dtemp=[data[i][4]+data[i][2]+": "+data[i][1]+" "+data[i][3],graphics_path+data[i][1].match(regex)[0]+".jpg",graphics_pos_path+data[i][0]+"_print_position.jpg",data[i][15], data[i][16]] // AFegim la nova columna "ALL_SIZE_PRINT"
    }
      result.push(dtemp);
    } 
  }
  return result;
}

/********************************************************************************************************************************
 * Funci??n auxCabeceras
 *   A??ade las columnas necesarias a la cabecera en funci??n de los valores pasados por par??metro
 * Parametros: 
 *    - pNumElements: N??mero de columnas  @item, @pos_ y sizes_ para cada item pasado por par??metro
 *    - pItem: Valor a a??adir en el nombre de la columna despu??s de cada prefijo.
 *    - pEntrades: Nombre de las columnas a a??adir.
 * Retorna: Un array con los nombres de las cabeceras a a??adir (originalmente se corresponden con las pesta??as BOG, BOM y BOL.
 ********************************************************************************************************************************/
function auxCabeceras(pNumElements,pItem,pEntrades){
  var result = [];
  for(var i=0;i<pNumElements;i++){
    for(var ii = 0; ii < pEntrades.length;ii++){
          dtemp=pEntrades[ii]+pItem+i
          result.push(dtemp)
        }
  }
return result;
}

/******************************************************************************************************************
 * Funci??n auxTrf
 *   A??ade las columnas necesarias de cada pesta??a auxiliar (BOL, BOL, BOM)
 * Par??metres: 
 *    - data : Contenido ya existente (columnas incialmente con valor) al que se a??adir??n el resto de columnas.
 *    - len: N??mero de columnas a a??adir al fichero (elementos del vector)
 *    - populator: caracter a a??adir cuando no hay valores.
 * Retorna: Un array con los valores a??adidos.
 ******************************************************************************************************************/

function auxTrf(data,len,populator){
    var result=[]
    var index=0
    for(var i = 0; i < len;){
      if(data[index]!=null){
        for(var ii = 0; ii < data[index].length & i<len ;ii++){
          dtemp=data[index][ii]
          result.push(dtemp)
        }
        index++
        i = i + ii
      }      
      else{
          dtemp=populator
          result.push(dtemp)
          i++
          index++ 
          }
    }
    return result
}

function product(style) {
  this.style=style;
  this.row ={};
  this.colors = [];
  this.images = [];
  this.prices= [];
  this.collection=[];
}