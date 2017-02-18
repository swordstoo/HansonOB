//Runs each time a form is submitted
function FormSubmission(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FormResponses = ss.getSheetByName('Form Responses 1');
  var Calculation = ss.getSheetByName('Calculation');
  var OrderBook = ss.getSheetByName('Order Book');
  var OrderCount = Calculation.getRange('L2').getValue(); //Gets the total number of orders in Form Responses list
  var date_time = e.values[0]; var Cust_N = e.values[1]; var Cust_P = e.values[2]; var SM = e.values[3]; var SKU = e.values[4]; var Quote = e.values[5]; var Category = e.values[6]; //Gets values from the Form and declares them as variable names
  var DestinationFolder = DriveApp.getFolderById('0BwZDnexHoR3dUXpDOTJjdDNkZ1k'); //Gets the ID of the folder to send a generated document to
  var newDocId = DriveApp.getFileById('1tK7l2e3B6Q7rCf9zLSfLEZ3idJMSrjUP05qaZhfF5Qk').makeCopy('Order for: '+ Cust_N+" ("+date_time+")",DestinationFolder).getId(); //Makes a copy of a template and gets the ID
  var link = 'https://docs.google.com/document/d/'+newDocId+'/edit'; //Gets the link of the generated document
  var newDoc = DocumentApp.openById(newDocId); //Gets the new document as a file
  var newDocBody = newDoc.getActiveSection(); //Gets the new document's Body
  var newDocFooter = newDoc.getFooter(); //Gets the new document's footer
  //This section replaces the text (See: ReplacementText array) in the new document with the information from the form (See: e.values array)
  var ReplacementText = ['<<Cust_N>>', '<<Cust_P>>', '<<SM>>', '<<SKU>>', '<<Quote>>'];
  for (i=0;i<ReplacementText.length();i++) {
   newDocBody.replaceText(ReplacementText[i],e.values[i]);
  }
  FormResponses.getRange(OrderCount, 8).setValue(link); //Adds the link to the orders row in Form Responses
  newDoc.saveAndClose(); //Save and close the document for good measure
  pdf = newDoc.getBlob().getAs('application/pdf'); //Creates a PDF copy of the document
  var subject = "An order has been added to the Stoughton online order book."; //Declaring subject
  var body = "A "+SKU+" has been ordered for "+Cust_N+" by "+SM+"."; //Declaring body
  if (Category=="Phone Repair") { //Sends an email to Phone Repair techs
    MailApp.sendEmail("sparkycbass@gmail.com", subject, body, {htmlBody: body, attachments: pdf}); 
    MailApp.sendEmail("chadyshack@gmail.com", subject, body, {htmlBody: body, attachments: pdf});
  }
  else { //Sends an email to all others otherwise
    MailAPp.sendEmail("fentonshack@gmail.com", subject, body, {htmlBody: body, attachments: pdf});
  }
}
//v1 Inconsistant runtime
//v2 .104s runtime