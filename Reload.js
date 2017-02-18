function Reload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FormResponses = ss.getSheetByName("Form Responses 1");
  var OrderBook = ss.getSheetByName("Order Book");
  var Calculation = ss.getSheetByName("Calculation");
  var Import = Calculation.getRange('J1').getValue(); //Number of Orders to Import from Form Responses to Order Book
  var Blank = Calculation.getRange('N1').getValue(); //The first blank line in the Order Book
  var Delete = Calculation.getRange('O1').getValue(); //Total number of orders to delete
  OrderBook.getRange(1, 1, Blank, 11).setBackground("#FFFFFF"); //Clears out all highlighting from freshly imported orders
  while (Delete > 0) { //This loop clears orders marked "yes" on the order book
    var ToDelete = Calculation.getRange('M1').getValue(); //First line to delete, updates with each loop
    OrderBook.getRange(ToDelete,1,1,11).clear({contentsOnly:true}); //Deletes the line in question. Contents only to avoid notes from the line in question
    OrderBook.getRange(ToDelete,1).clearNote(); //Removes the note from the column A
    OrderBook.getRange(ToDelete,10,1,1).setValue('No') //Sets the line to not delete for next loop
    OrderBook.getRange(ToDelete+1,1,1000,11).copyTo(OrderBook.getRange(ToDelete,1,1,1)); //Copies all content below to fit the now empty space
    Delete--; //"Checks" off a row for the while loop
  }
  while (Import > 0) { //This loop imports each order into the order book
    var ToImport = Calculation.getRange('K1').getValue(); //Gets the order # of the order to import. Updates with each loop
    var Blank = Calculation.getRange('N1').getValue(); //The first blank line in the Order Book. Updates each loop
    FormResponses.getRange(ToImport,1,1,7).copyTo(OrderBook.getRange(Blank,1),{contentsOnly:true}); //Copies accross the order. Contents only to avoid accidental notes
    FormResponses.getRange(ToImport,8).copyTo(OrderBook.getRange(Blank,11));//Copies accross link
    OrderBook.getRange(Blank,1).setNote(ToImport); //Adds the order number to the A column for future scripts
    Calculation.getRange(ToImport,10).setValue('Y'); //Sets this order as imported in Calculation
    OrderBook.getRange(Blank,1,1,11).setBackground("#C9DAF8"); //Highlights light blue to show it is a new order
    Import--; //"Checks" off a row for the while loop
  }
}
//v1 = 2.744s, 5.14s
//v2 = 1.01s, 3.838s
//v2 = 63%, 25.3% (Improvement percent)