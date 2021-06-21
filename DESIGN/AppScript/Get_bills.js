/****************************************************************************************
* Function extra                                                                        *
* Construye el contenido de las columnas que no están contenidas en la hoja de cálculo  *
* Parámetros:                                                                           *
*  - sheet_b: Sheet origen de los datos.                                                * 
*  - headers: Array con las cabeceras de la nueva hoja de cálculo.                      *
*  - key: Columna principal a partir de la que se obtienen los datos.                   *
* Retorna: un array con todas las nuevas columnas a añadir                              * 
*****************************************************************************************/
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
           case "comments_img1":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-1.jpg"
            break;  
          case "comments_img10":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-10.jpg"
            break;  
           case "comments_img11":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-11.jpg"
            break;  
            case "comments_img12":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-12.jpg"
            break;  
           case "comments_img13":
            dtemp=comments_path+data[i][data[0].indexOf(global_key1)]+"-13.jpg"
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