function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var OrderBook = ss.getSheetByName('Order Book');
  var Calculation = ss.getSheetByName('Calculation');
  var Responses = ss.getSheetByName('Form Responses 1');
  var EditedCol = e.range.getColumn(); //Gets the column # where the edit took place
  var EditedRow = e.range.getRow(); //Gets the row # where the edit took place
  var NewData = e.value;
  var OldData = e.oldValue;
  if (OldData == undefined) {
   OldData = ''; 
  }
  if (NewData == "[object Object]") {
   NewData = ''; 
  }
  if (e.source.getActiveSheet().getName() != "Order Book") {
   return //Stops excecution if the edited sheet isn't the Order Book sheet. 
  }
  if (EditedCol == 10) { //The yes/no column
    if (NewData == 'Yes') {
      OrderBook.getRange(EditedRow, 1,1,10).setBackground('#F4c7c3'); //Highlights red when removing
    }
    else if (NewData =='No') {
      OrderBook.getRange(EditedRow, 1,1,10).setBackground('#b7e1cd'); //Highlights green when keeping
    }
  }
  if (EditedCol == 8) { //Writes the data from the "Ordered" column in the Order Book and archives it in the Form Responses 1 sheet
    Responses.getRange(OrderBook.getRange(EditedRow,1).getNote(),9).setValue(NewData); 
  }
  if (EditedCol == 9) { //Writes the data from the "Called" column in the Order Book and archives it in the Form Responses 1 sheet
    Responses.getRange(OrderBook.getRange(EditedRow,1).getNote(),10).setValue(NewData); 
  }
  //Warning messages
  var redact = 0
  if (EditedCol < 8 || EditedCol == 11) { // || is this "or" that. Editing any of the non static columns in the Order Book
    if (NewData != '' && OrderBook.getRange(EditedRow, 11).getValue() == '') { //Only if the edited data isn't blank, and the edited line **IS** blank
      e.range.setValue(OldData); //Sets the data back correctly
      Browser.msgBox("Do not manually add orders on this page! Use the 'Submit Order' link in the bookmarks bar! Or use this link: https://docs.google.com/forms/d/e/1FAIpQLSequSqohWXj3XaVv02YyUafbH1wvfk7IbTdkmGRIpk77GlVJQ/viewform")
    }
    else if (Browser.msgBox("Warning!", "You have edited core data in the ["+ OrderBook.getRange(1, EditedCol).getValue().split(":")[0] + "] column from [" + OldData + "] to [" + NewData + "]! This will only reflect here and not in any archives, past backups, or printouts. Please make sure this is what you intend before proceeding.",Browser.Buttons.OK_CANCEL) =="ok") {
      if (Browser.msgBox("Warning!", "Press OK to confirm.",Browser.Buttons.OK_CANCEL) =="ok") {}
      else { redact++; }
    }     
    else { redact++; }
    if (redact > 0) {
      Browser.msgBox("Changes redacted", "Your edits have been automatically reverted to [" + OldData + "].",Browser.Buttons.OK);
      e.range.setValue(OldData);
    }
  }
}