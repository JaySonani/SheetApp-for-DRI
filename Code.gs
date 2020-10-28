var ui = SpreadsheetApp.getUi();
var ash = SpreadsheetApp.getActiveSpreadsheet();
var choki = ash.getActiveSheet();

var master = ash.getSheetByName("MASTER SHEET");
var home = ash.getSheetByName("HOME");

let LOADING = false;

var master_weight = 0;
var master_polish_weight = 0;
var master_piece = 0;
var cuts = [];

function amj(){
  SpreadsheetApp.getActiveSpreadsheet().toast("Loading...", "Please wait",1);
  SpreadsheetApp.getUi().show
  
  
}


function showLoading()
{
// Display a modeless dialog box with custom HtmlService content.
  var width = 500;
  var height = 200;

  var title = "Please wait";
  var html = `
  
  <!DOCTYPE html>
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
          
        <script>
          var myVar = setInterval(myTimer ,5000);
          
          function myTimer() {
            google.script.host.close();
          } 
        </script>  
      
        <style>
          .loader {
            margin-top: 40px;
            margin-left: 180px;
//            transform: translate(-50%, -50%);
            border: 16px solid #f3f3f3;
            border-radius: 50%;
            border-top: 16px solid #3498db;
            width: 80px;
            height: 80px;
            -webkit-animation: spin 2s linear infinite; /* Safari */
            animation: spin 2s linear infinite;
          }
          
          /* Safari */
          @-webkit-keyframes spin {
            0% { -webkit-transform: rotate(0deg); }
            100% { -webkit-transform: rotate(360deg); }
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <ceter>
        <div class="loader"></div>
        </center>
      </body>
    </html>


  `;
  
  var htmlOutput = HtmlService
     .createHtmlOutput(html)
     .setWidth(width)
     .setHeight(height);
  
 SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);

  
  
} 
  



/*function onEdit() {
  
  var rough_amount = choki.getRange("A9").getValue();
  var weight_total = choki.getRange("C26").getValue();

  var weightsEmpty = false;
  var weights = choki.getRange("weights");
  var weightValues = weights.getValues();
  for(var i=0;i<weightValues.length;i++)
    if(weightValues[i].toString().length == 0){
      weightsEmpty = true;
      break;
    }  
  if(!weightsEmpty && rough_amount!=weight_total){
    ui.alert("❌ Weights don't match !","Please check weights of all the types of diamonds.",ui.ButtonSet.OK);
    choki.setActiveRange(weights);
  }
  
//  UpdateMaster();
  
}*/


function openSheet() {
  var cell = SpreadsheetApp.getCurrentCell().getValue();
  try{
    ash.setActiveSheet(ash.getSheetByName(cell));
  }
  catch(e){
    ui.alert("  ❌ Sheet not available", "Please select from cut number column only.", ui.ButtonSet.OK);
  }
}

function openFromHome(){
 
  var cell = ash.getRange("E18").getValue();  
  ash.getRange("E18").setValue("");
   if(cell == ""){
      ui.alert("❌  Cut Number required!","Please enter the cut number in the box.",ui.ButtonSet.OK);
      choki.setActiveSelection("E18");
   }else if(ash.getSheetByName(cell) == null){ 
     ui.alert("  ❌ Sheet not available", "Please enter another cut number.", ui.ButtonSet.OK);
     choki.setActiveSelection("E18");
   }
  else
    
  ash.setActiveSheet(ash.getSheetByName(cell));
  
  
  
  
}


function sortSheets () {
  var mm = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = mm.getSheets();
  var names = [];
  for(var i=0;i<sheets.length;i++){
    if(sheets[i].getSheetName() == "CHOKI" || sheets[i].getSheetName() == "MASTER SHEET" || sheets[i].getSheetName() == "HOME")
      continue;
    names.push(sheets[i].getSheetName());
  }
  Logger.log(names);
//  names = names.sort();
  Logger.log(names);

  // 
  mm.setActiveSheet(ash.getSheetByName("HOME"));
  mm.moveActiveSheet(1);
  mm.setActiveSheet(ash.getSheetByName("MASTER SHEET"));
  mm.moveActiveSheet(2);
  mm.setActiveSheet(ash.getSheetByName("CHOKI"));
  mm.moveActiveSheet(3);
  
  for(var i=0;i<names.length;i++){  
      var mma = SpreadsheetApp.getActiveSpreadsheet();

    mma.setActiveSheet(mma.getSheetByName(names[i]));
    mma.moveActiveSheet(i+4);
   
  }
//  mm.setActiveSheet(mm.getSheetByName("MASTER SHEET"));

  
}



function NewSheet(){
  
  var newName = ash.getActiveSheet().getRange("E9").getValue();
  if(newName == ""){
    ui.alert("❌  Cut Number required!","Please enter the cut number in the box.",ui.ButtonSet.OK);
    choki.setActiveSelection("E9");
  }
  else if(ash.getSheetByName(newName) != null){
    ui.alert("❌ "+ newName + " already exists!","Please enter another cut number.",ui.ButtonSet.OK);
    choki.setActiveSelection("E9");
    choki.getActiveCell().setValue("");
    
  }
    
  else{

  ash.getActiveSheet().getRange("E9").setValue("");
  SpreadsheetApp.setActiveSheet(ash.getSheetByName("CHOKI"));
  ash.duplicateActiveSheet();
  ash.getActiveSheet().setName(newName);
  ash.getActiveSheet().getRange("B1").setValue(newName);
//  sortSheets();
  ash.setActiveSheet(ash.getSheetByName(newName));

    
    
  }
}



function DeleteSheet(){
  var newName = ash.getActiveSheet().getRange("L9").getValue();
  ash.getActiveSheet().getRange("L9").setValue("");

  if(newName == ""){
    ui.alert("❌  Cut Number required!","Please enter the cut number in the box.",ui.ButtonSet.OK);
    choki.setActiveSelection("L9");
  }
  else if(ash.getSheetByName(newName) == null){
    ui.alert("❌ "+ newName + " not found!","Please enter another cut number.",ui.ButtonSet.OK);
    choki.setActiveSelection("L9");
    choki.getActiveCell().setValue("");
    
  }
    
  else{
    var response = ui.alert("Are you sure ❓", "Sheet " + newName + " will be deleted.", ui.ButtonSet.YES_NO_CANCEL);
    if(response == ui.Button.YES){
      ash.setActiveSheet(ash.getSheetByName(newName));
      ash.deleteActiveSheet();
      choki.setActiveSelection("L9");

    } 
  }
  
}

function updateNew(){
   
  var all_sheets = ash.getSheets();
  cuts = [];
  
  for(var i=0;i<all_sheets.length;i++){
    if(all_sheets[i].getName().equals("CHOKI") || all_sheets[i].getName().equals("MASTER SHEET") || all_sheets[i].getName().equals("HOME")){
      continue;
    }
    cuts.push(all_sheets[i]);
  }
  

  cut_source = [];
  for(var i=0;i<cuts.length;i++){
    cut_source.push(cuts[i].getRangeList(['B3','B1','B6','A9','B9','D9','I18','K15','D12']).getRanges());
  }

  for(var i=15;i<16;i++){
    
    for(var i in cut_source[i]){
      ui.alert(i.getValue());
    }
    
    
    
  }
  
  
  
  
  
}


function getTotal(){
  
//  LOADING = true;
//  showLoading();
//     
//  var all_sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
//  master_weight = 0;
//  master_polish_weight = 0;
//  master_piece = 0;
//  cuts = [];
//  
//  for(var i=0;i<all_sheets.length;i++){
//    if(all_sheets[i].getSheetName() == "CHOKI" || all_sheets[i].getSheetName() == "MASTER SHEET" || all_sheets[i].getSheetName() == "HOME")
//      continue;
//
//    cuts.push(all_sheets[i]);
//  }
//  
//  for(var i=0;i<cuts.length;i++){
//    master_weight += cuts[i].getRange("A9").getValue();
////    master_polish_weight += cuts[i].getRange("I18").getValue();
//    master_piece += cuts[i].getRange("D12").getValue();
//  }

//  master.getRange(cuts.length+6,2,1,10).setFontWeight("bold");
//  master.getRange(cuts.length+6,2,1,10).setBackground("#5b0f00");
//  master.getRange(cuts.length+6,2,1,10).setFontColor("white");
  
  
//  master.getRange(cuts.length+6,2).setValue("Total...");
  
  
  master.setActiveSelection("B155");
  master.getRange(155,6).setValue("=SUM(F5:F153)");
  master.getRange(155,9).setValue("=SUM(I5:I153)");
  master.getRange(155,11).setValue("=SUM(K5:K153)");
  
}



function UpdateMaster(){
   
  var all_sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  cuts = [];
  
  for(var i=0;i<all_sheets.length;i++){
    if(all_sheets[i].getSheetName() == "CHOKI" || all_sheets[i].getSheetName() == "MASTER SHEET" || all_sheets[i].getSheetName() == "HOME")
      continue;
    cuts.push(all_sheets[i]);
  }
//  master.getRange("C5:K300").setValue(""); 
  for(var i=0;i<cuts.length;i++){
//    Logger.log(cuts[i].getName());

    master.getRange(i+5,2).setValue(i+1);
    master.getRange(i+5,3).setValue(cuts[i].getRange("B3").getValue());
    master.getRange(i+5,4).setValue(cuts[i].getName());
    master.getRange(i+5,5).setValue(cuts[i].getRange("B6").getValue());
    master.getRange(i+5,6).setValue(cuts[i].getRange("A9").getValue());
    
    master.getRange(i+5,7).setValue(cuts[i].getRange("B9").getValue());
    master.getRange(i+5,8).setValue(cuts[i].getRange("D9").getValue());
    master.getRange(i+5,9).setValue(cuts[i].getRange("I18").getValue());
    
    master.getRange(i+5,10).setValue(cuts[i].getRange("K15").getValue());
    master.getRange(i+5,11).setValue(cuts[i].getRange("D12").getValue())
  
  }
}




function mergeSheets(){
  var cell1 = ash.getRange("J18").getValue();
  var cell2 = ash.getRange("L18").getValue();
  
  ash.getRange("J18").setValue("");
  ash.getRange("L18").setValue("");
  
  var sheet1 = ash.getSheetByName(cell1);    
  var sheet2 = ash.getSheetByName(cell2);
  
  var mergeName = cell1 + "-" + cell2;
  
  if(cell1 == "" || cell2 == "")
     ui.alert("❌  Cut Number required!","Please enter the both the cut numbers.",ui.ButtonSet.OK);
  else if(sheet1 == null || sheet2 == null)
     ui.alert("❌  Sheet not available!","Make sure both the sheets exists in workbook.",ui.ButtonSet.OK);
  else if(cell1 == cell2)
     ui.alert("❌  Same sheets!","Merging of the same sheet is not possible.",ui.ButtonSet.OK);    
  else if(ash.getSheetByName(mergeName) != null)
     ui.alert("❌  Already exists!","Sheet " + mergeName + " already exists in the workbook.",ui.ButtonSet.OK);    
    
  else
  {
    var sheet1 = ash.getSheetByName(cell1);    
    var sheet2 = ash.getSheetByName(cell2);
    
    var mergeName = cell1 + "-" + cell2;
    
    
    ash.setActiveSheet(ash.getSheetByName("CHOKI"));
    ash.duplicateActiveSheet();
    ash.getActiveSheet().setName(mergeName);
    
    var mergeSheet = ash.getSheetByName(mergeName);
    
    //do formatting of cell B1 here
    ash.getActiveSheet().getRange("B1").setValue(mergeName);
    sortSheets();
    
    ash.setActiveSheet(ash.getSheetByName(mergeName));
    
    
    //filling values...
    
    //cut no.
    
    mergeSheet.getRange("B1").setValue(mergeName);
    //cut
    mergeSheet.getRange("B6:C6").merge();
    mergeSheet.getRange("B6").setHorizontalAlignment("right");
    mergeSheet.getRange("B6").setValue(cell1 + "(" + sheet1.getRange("B6").getValue() + ") " + cell2 + "(" + sheet2.getRange("B6").getValue() + ")");
    
    //weight
    mergeSheet.getRange("A9").setValue(sheet1.getRange("A9").getValue() + sheet2.getRange("A9").getValue());

    //amount
    mergeSheet.getRange("C9").setValue(sheet1.getRange("C9").getValue() + sheet2.getRange("C9").getValue())
    
    //rate
    mergeSheet.getRange("B9").setValue(Math.round(mergeSheet.getRange("C9").getValue()/mergeSheet.getRange("A9").getValue()));
    
    //party
    mergeSheet.getRange("D9").setValue(sheet1.getRange("D9").getValue() + " | " + sheet2.getRange("D9").getValue());
    
    //broker
    mergeSheet.getRange("F9").setValue(sheet1.getRange("F9").getValue() + " | " + sheet2.getRange("F9").getValue());
    
    //weight, pieces, prices, amounts
      mergeSheet.getRange("C12").setValue(sheet1.getRange("C12").getValue() + sheet2.getRange("C12").getValue()); 
      mergeSheet.getRange("D12").setValue(sheet1.getRange("D12").getValue() + sheet2.getRange("D12").getValue());

    for(var i=13;i<=25;i++){
      mergeSheet.getRange("C"+i).setValue(sheet1.getRange("C"+i).getValue() + sheet2.getRange("C"+i).getValue()); 
      mergeSheet.getRange("D"+i).setValue(sheet1.getRange("D"+i).getValue() + sheet2.getRange("D"+i).getValue());
      mergeSheet.getRange("F"+i).setValue(sheet1.getRange("F"+i).getValue() + sheet2.getRange("F"+i).getValue());   
      if(mergeSheet.getRange("C"+i).getValue() == 0)
        continue;
      mergeSheet.getRange("E"+i).setValue(Math.round(mergeSheet.getRange("F"+i).getValue() / mergeSheet.getRange("C"+i).getValue()));      
    }
    
    mergeSheet.getRange("K10").setValue(Math.round((sheet1.getRange("K10").getValue()+sheet2.getRange("K10").getValue())/2,2));
    
    var lpp = ((sheet1.getRange("D12").getValue() * sheet1.getRange("K10").getValue()) + (sheet2.getRange("D12").getValue() * sheet2.getRange("K10").getValue())) / (sheet1.getRange("D12").getValue() + sheet2.getRange("D12").getValue());
    mergeSheet.getRange("K10").setValue(Math.round(lpp));  
    
    
    mergeSheet.getRange("K12").setValue(sheet1.getRange("K12").getValue() + sheet2.getRange("K12").getValue());
    mergeSheet.getRange("I18").setValue(sheet1.getRange("I18").getValue() + sheet2.getRange("I18").getValue());

    
  }  
}

