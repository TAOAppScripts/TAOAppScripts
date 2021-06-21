// Este es un codigo para refrescar la pestaña de Pricing 

function Refrescar() {

     // Columna letra D = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna D

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Main!B:Z,25,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('D2:D999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('D2:D999').activate();  
  
  // Columna letra E = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna E

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,Main!B:X,23,FALSE)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('E2:E999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('E2:E999').activate(); 
  
  // Columna letra F = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna F

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F2').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,BOM!A$2:H$999,2,FALSE)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('F2:F999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('F2:F999').activate(); 
  
  // Columna letra G = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna G

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G2').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(VLOOKUP(A2,BOG!A:E,5,FALSE)="allover","Si"," ")');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('G2:G999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('G2:G999').activate(); 
  
 // Columna letra H = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna H

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H2').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,Main!B:O,14,FALSE)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('H2:H999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('H2:H999').activate(); 
   
  
 
  // Columna letra I = Nueva Reparacion - mar 16 feb 2021 - se agrego programacion a la columna I
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,BOM!A$2:J$1100,7,FALSE)');

// Este es otro proceso que autorellena las celdas siguientes

         var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I2:I999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('I2:I999').activate();
  
  // Columna letra J = Nueva Reparacion - Vie 18 dic 2020/ 16 FEB - se agrego programacion a la columna j

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,10,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('J2:J999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('J2:J999').activate(); 
  
  // Columna letra K = Nueva Reparacion - Vie 18 dic 2020 / Dom 31 ene 2021 - se agrego programacion a la columna k

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,9,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('K2:K999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('K2:K999').activate(); 

// ***** PESTAÑA CONS. FABRIC *****
  
// Columna letra L = Nueva Reparacion - Vie 18 dic 2020 / Dom 31 ene 2021/ 16 FEB 21 - se agrego programacion a la columna L

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOM!L:L,BOM!M:M,A2&"fabric"))');
  


// Este es otro proceso que autorellena las celdas siguientes

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('L2:L999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('L2:L999').activate(); 
  
   // Columna letra M = Nueva Reparacion - Vie 18 dic 2020 - se agrego programacion a la columna M

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOM!L:L,BOM!M:M,A2&"trim"))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('M2:M999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('M2:M999').activate(); 
  
  // Columna letra N = Nueva Reparacion - Lun 21 dic 2020 - se agrego programacion a la columna N

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOM!L:L,BOM!M:M,A2&"Yarn"))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('N2:N999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('N2:N999').activate(); 
  
  
// nota : REPARAR ESTO COLUMNAS LETRA O, LETRA P, LETRA Q - DE PESTAÑA Pricing
         
       // ***** PESTAÑA FABRIC *****

// Columna letra O

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(IF(H2="pigmentos",J2*5.5,IF(H2="directo/reactivo",J2*3,0)))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('O2:O999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('O2:O999').activate(); 

// Columna letra P

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,12,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('P2:P999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('P2:P999').activate(); 

  // Columna letra Q

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Q2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOG!R:R,BOG!Q:Q,A2&"allover"))*I2');
  
   
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Q2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('Q2:Q999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('Q2:Q999').activate(); 

// Columna letra R = Nueva Reparacion - jue 17 dic 2020 - se agrego programacion a la columna R

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('R2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOG!R:R,BOG!Q:Q,A2&"positional"))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('R2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('R2:R999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('R2:R999').activate(); 
  
  // Columna letra S = Nueva Reparacion - Lun 21 dic 2020 - se agrego programacion a la columna S

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOG!R:R,BOG!Q:Q,A2&"badge"))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('S2:S999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('S2:S999').activate(); 
  
  
 // Columna letra T = Nueva Reparacion - Lun 21 dic 2020 - se agrego programacion a la columna T

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOG!R:R,BOG!Q:Q,A2&"embroidery"))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('T2:T999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('T2:T999').activate(); 
  
  // Columna letra U = Nueva Reparacion - Martes 22 dic 2020 - se agrego programacion a la columna U

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('U2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,14,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('U2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('U2:U999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('U2:U999').activate(); 
  
  // Columna letra V = Nueva Reparacion - Martes 22 dic 2020 - se agrego programacion a la columna V

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('V2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,17,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('V2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('V2:V999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('V2:V999').activate(); 
  
  // Columna letra W = Nueva Reparacion - jue 17 dic 2020 - se agrego programacion a la columna W

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('W2').activate();
  spreadsheet.getCurrentCell().setFormula('=IF(H2=0,IFERROR((SUM(L2:U2)*(V2-1)))*1.03,IFERROR((SUM(L2:U2)*(V2-1)))*1.05)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('W2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('W2:W999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('W2:W999').activate(); 
  
  // Columna letra X= Nueva Reparacion - Jueves 21 Ene 2021 - se agrego programacion a la columna X

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('X2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(SUMIFS(BOL!K:K,BOL!A:A,A2))');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('X2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('X2:X999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('X2:X999').activate();

   // Columna letra Y= Nueva Reparacion - Lunes 25 Ene 2021 - se agrego programacion a la columna Y

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Y2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,15,FALSE))');
  
  // formula anterior: =iferror((SUM(O2:V2)*(AH2-1))) - Deprecated
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Y2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('Y2:Y999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('Y2:Y999').activate();
  
 
  

// Columna letra Z = Nueva Reparacion - Martes 22 dic 2020 / Jueves 21 Ene 2021- se agrego programacion a la columna Z

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Z2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,16,FALSE))');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Z2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('Z2:Z999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('Z2:Z999').activate(); 

// Columna letra AA = Nueva Reparacion - Martes 22 dic 2020 / 9 Marzo 2021- se agrego programacion a la columna AA

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AA2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR((SUM(K2:Z2))-V2)*2%');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AA2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AA2:AA999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AA2:AA999').activate(); 

  // Columna letra AB = Nueva Reparacion - Martes 22 dic 2020 - se agrego programacion a la columna AB

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AB2').activate();
  spreadsheet.getCurrentCell().setFormula('=SUM(K2:AA2)-V2');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AB2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AB2:AB999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AB2:AB999').activate(); 

  // Columna letra AC = Nueva Reparacion - Martes 22 dic 2020 - se agrego programacion a la columna AC

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AC2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(AE2/AB2)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AC2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AC2:AC999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AC2:AC999').activate(); 

  // Columna letra AD= Nueva Reparacion - Martes 22 dic 2020 - se agrego programacion a la columna AD

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AD2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(AE2/2.5)');
  
  // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AD2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AD2:AD999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AD2:AD999').activate(); 
  
  // Columna letra AE= Nueva Reparacion - Martes 12 Ene 2021 / Dom 31 Ene 2021 - se agrego programacion a la columna AE

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AE2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,7,FALSE))');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AE2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AE2:AE999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AE2:AE999').activate();
  
// Columna letra AF= Nueva Reparacion - Jueves 21 Ene 2021/ Dom 31 Ene 2021 - se agrego programacion a la columna AF

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AF2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,8,FALSE))');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AF2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AF2:AF999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AF2:AF999').activate();

  // Columna letra AG= Nueva Reparacion - Jueves 21 Ene 2021 - se agrego programacion a la columna AG
//=(AB2-(SUM(X2:AA2)))-AL2
var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AG2').activate();
  spreadsheet.getCurrentCell().setFormula('=(AB2-(SUM(X2:AA2)))-AL2');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AG2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AG2:AG999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AG2:AG999').activate();
  
 
  // Columna letra AH= Nueva Reparacion - Jueves 21 Ene 2021 - se agrego programacion a la columna AH

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AH2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(A2,Production!A:Q,5,FALSE))');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AH2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AH2:AH999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AH2:AH999').activate();

   // Columna letra AI= Nueva Reparacion - Jueves 21 Ene 2021 / 16 FEB 21 - se agrego programacion a la columna AI

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AI2').activate();
  spreadsheet.getCurrentCell().setFormula('=((IFERROR(VLOOKUP(A2,Production!A:Q,13,FALSE)))+X2+Y2+Z2+AA2)+AH2+AL2');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AI2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AI2:AI999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AI2:AI999').activate();

   // Columna letra AJ= Nueva Reparacion - Jueves 21 Ene 2021/ 16 FEB 21 - se agrego programacion a la columna AJ

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AJ2').activate();
  spreadsheet.getCurrentCell().setFormula('=IFERROR(AE2/AI2)');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AJ2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AJ2:AJ999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AJ2:AJ999').activate();

   // Columna letra Ak= ***Contiene Imagenes*** - 17 FEB 21 - se agrego programacion a la columna AJ

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AK2').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,Img!B:D,3,FALSE)');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AK2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AK2:AK999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AK2:AJ999').activate();
  
  // Columna letra AL=  22 FEB 21 - se agrego programacion a la columna AL

var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AL2').activate();
  spreadsheet.getCurrentCell().setFormula('=iferror(VLOOKUP(A2,Production!A:Q,13,FALSE))');
  
 // Este es otro proceso que autorellena las celdas siguientes
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('AL2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AL2:AL999'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AL2:AL999').activate();
}








