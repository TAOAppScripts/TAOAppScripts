function openDialog(owners) {
  var t = HtmlService.createTemplateFromFile('index');
  //var userProperties = PropertiesService.getUserProperties();
  //var newProperties = {type: 'Kids',export:'Export',columns:22,image:'V35'};
  //userProperties.setProperties(newProperties);
  t.data = owners
  html=t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Select designers to export tsv file:');
  
}

function getValuesFromForm(form){
  //var firstName = form.firstName,
  //    lastName = form.lastName,
  //    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //sheet.appendRow([firstName, lastName]);
  //email=form.email;
  owners=[];
  for (item in form) {
    if (item!="Designer") owners.push(item);
  }
 //var userProperties = PropertiesService.getUserProperties();
 //var ss = SpreadsheetApp.getActiveSpreadsheet();
 //var sheet = ss.getSheetByName(userProperties.getProperty('type'))
 //var newsheet=ss.getSheetByName(userProperties.getProperty('export'));
 Export_All(owners)
}

function getdesigners(){
  var data=sheet_main.getRange(1,1,sheet_main.getLastRow(),sheet_main.getLastColumn()).getValues();
  var owners_ix=data[0].indexOf('Designer')
  var owners=[] 
  for (var i=1;i<data.length;i++){
    var ix=owners.indexOf(data[i][owners_ix])
    if (ix==-1){
      owners.push(data[i][owners_ix])  
    }
  }
  
  return owners;
}

