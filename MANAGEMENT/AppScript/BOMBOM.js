keyfields=["Reference","Name","Description","Pattern Ref.","Supplier"]
//bomfields=["Pattern|Code","Item_description","Order","Cons.","Unit","Type","Style","Factory"]
bomfields=["Pattern|Code","CONCAT","Item_description","Order","Cons.","Unit","Real Cons.","Type","Style","Factory", "Price", "Total Cons."]

function Create_BOMext() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newsheet= ss.insertSheet();
  newsheet.getRange(1,1,1000,200).setNumberFormat('@STRING@')
  var date = Utilities.formatDate(new Date(), "GMT", "yy-MM-dd"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
  newsheet.setName('exBOM_'+date)
  
  var sheet = ss.getSheetByName("Main");
  var range=sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var data = range.getValues();
  var counter=1;
  var headers=data[0];
  var Style_index = keyfields.indexOf("Reference");
  var Pattern_index=keyfields.indexOf("Pattern Ref.");
  var bom =[];
  
  newsheet.getRange(counter,1,1,keyfields.concat(bomfields).length).setValues([keyfields.concat(bomfields)])
  
  for(var i = 0; i < data.length; i++) 
  {
    data_row=filter_data(data[i],headers);
    bom=Get_BOMdata(data_row[Style_index]);
    for(var ii = 0; ii < bom.length; ii++){ 
      row=data_row.concat(bom[ii])
      newsheet.getRange(counter+1,1,1,row.length).setValues([row])
      counter+=1;
    }
    bom=Get_BOMdata(data_row[Pattern_index]);
    for(var ii = 0; ii < bom.length; ii++){ 
      row=data_row.concat(bom[ii])
      newsheet.getRange(counter+1,1,1,row.length).setValues([row])
      counter+=1;
    }
  } 
}

function Get_BOMdata(pattern){
  var result=[];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("BOM");
  var range=sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var data = range.getValues();
  
  var index = data[0].indexOf("Pattern|Code");
  
  for(var i = 0; i < data.length; i++) 
  {
    if (data[i][index]==pattern){
      var d = data[i];
      
      // Add new fields
      // 1 - CONCAT
      d.splice(1,0,'=INDIRECT(CONCAT("A",ROW()))&INDIRECT(CONCAT("M",ROW()))');
      // 4 - Real Cons.
      // 10 - Price
      d[10] = d[11]; // 
      // 11 - Total Cons.
      d[11] = '=INDIRECT(CONCAT("L",ROW()))*INDIRECT(CONCAT("P",ROW()))';
      // Cut extra columns
      d = d.slice(0, 12);
      
      result.push(d);
    } 
  }

  return result;
}

function filter_data(data,headers){
  var dataf=[]
  for(var i=0;i<keyfields.length;i++){
  dataf.push(data[headers.indexOf(keyfields[i])]);
  }
  return dataf;
}

function Create_BOMextForni() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newsheet= ss.insertSheet();
  newsheet.getRange(1,1,1000,200).setNumberFormat('@STRING@')
  var date = Utilities.formatDate(new Date(), "GMT", "yy-MM-dd"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
  newsheet.setName('exBOMF_'+date)
  
  var sheet = ss.getSheetByName("Creatiu");
  var range=sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var data = range.getValues();
  var counter=1;
  var headers=data[0];
  var Fabric_index = keyfields.indexOf("Trimmings_Group");
  var bom =[];
  
  newsheet.getRange(counter,1,1,keyfields.concat(bomfields).length).setValues([keyfields.concat(bomfields)])
  
  for(var i = 0; i < data.length; i++) 
  {
    data_row=filter_data(data[i],headers);
    bom=Get_BOMdata(data_row[Fabric_index]);
    for(var ii = 0; ii < bom.length; ii++){ 
      row=data_row.concat(bom[ii])
      newsheet.getRange(counter+1,1,1,row.length).setValues([row])
      counter+=1;
    }
  } 
}

function generatePdf() {

var originalSpreadsheet = SpreadsheetApp.getActive();

var sourcesheet = originalSpreadsheet.getActiveSheet()
var sourcerange = sourcesheet.getActiveRange();  // range to get - here I get all of columns which i want
var sourcevalues = sourcerange.getValues();
var data = sourcesheet.getDataRange().getValues();

var newSpreadsheet = SpreadsheetApp.create(sourcesheet.getName()); // can give any name.
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var projectname = SpreadsheetApp.getActiveSpreadsheet();
var sheet = sourcesheet.copyTo(newSpreadsheet);
var destrange = sheet.getRange('A1:G7');
destrange.setValues(sourcevalues);
newSpreadsheet.getSheetByName('Sheet1').activate();
newSpreadsheet.deleteActiveSheet();

var pdf = DriveApp.getFileById(newSpreadsheet.getId());
var theBlob = pdf.getBlob().getAs('application/pdf').setName("name");

var folderID = "Folder Id"; // Folder id to save in a folder.
var folder = DriveApp.getFolderById(folderID);
var newFile = folder.createFile(theBlob);

DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}