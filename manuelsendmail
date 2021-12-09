function onOpen() {
  SpreadsheetApp.getUi() // This is to create a menu
      .createMenu('Special')
      .addItem('Skicka Mail', 'Mail')
      .addItem('Clear', 'clearRange')
      .addToUi();
}

function Mail() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Email',
     'Are you done?',
      ui.ButtonSet.YES_NO);

  // What happens if you click yes
  if (result == ui.Button.YES) {


 var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output sheet'); 
 var data = sh.getRange('A1:C5').getValues(); //Here is the range that will use for the mail table
  
  var emailAddress = 'test1@mail.com' + "," + 'testmail2'; //here you put who getting the mail
  var Subject = "weekly numb"; //here is the line of subject
  
  

var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + ' '  + '</td>';} 
  else
    if (row === 0)  {
      htmltable += '<th>' + data[row][col] + '</th>';
    }

  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     Logger.log(data);
     Logger.log(htmltable);
MailApp.sendEmail(emailAddress, Subject,'' ,{htmlBody: htmltable});
  




  } else {
    
  }
}

function clearRange() {
  //here it points what document should be clear
  var sheet = SpreadsheetApp.getActive().getSheetByName('data');
  sheet.getRange('A1:Z99999').clearContent();
}

