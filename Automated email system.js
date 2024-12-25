function formatDate(date) {
  if (!date) return "";
  if (Object.prototype.toString.call(date) === "[object Date]") {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy"); 
  }
  return date;
}
function generateAndSendDocs() {
  const sheetName = "Sheet1"; // enter sheet name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet with name "${sheetName}" not found.`);
  }

  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < startRow) {
    throw new Error("No data found in the sheet.");
  }

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn).getValues();

  data.forEach((row, index) => {
    const [Name, RollNo, Branch, College, Startdate, Enddate, ID, Hours, Email] = row;

    Logger.log(`Processing row ${index + startRow}: Name=${Name}, Roll No=${RollNo}, Branch=${Branch}, College=${College}, Startdate=${Startdate}, Enddate=${Enddate}, ID=${ID}, Hours=${Hours}, Email=${Email}`);

    if (!Email) {
      Logger.log(`Skipping row ${index + startRow} due to missing email.`);
      return;
    }

    const templateDocId = "1mZLEqWbavW6AbV82i-4o3FbwPivKDMyk68EzSQlgkZc"; // enter template id before /edit after /d/

    const docCopy = DriveApp.getFileById(templateDocId).makeCopy(`Document for ${Name}`);
    const doc = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();

    body.replaceText("{{Name}}", Name || "");
    body.replaceText("{{RollNo}}", RollNo || "");
    body.replaceText("{{Branch}}", Branch || "");
    body.replaceText("{{College}}", College || "");
    body.replaceText("{{Startdate}}", formatDate(Startdate)); 
    body.replaceText("{{Enddate}}", formatDate(Enddate));  
    body.replaceText("{{ID}}", ID || "");
    body.replaceText("{{Hours}}", Hours || "");

    Logger.log(`Replaced placeholders for ${Name}`);

    doc.saveAndClose();

    const pdf = DriveApp.getFileById(docCopy.getId()).getAs("application/pdf");

    GmailApp.sendEmail(Email, `Document for ${Name}`, 
      `Hello ${Name},\n\nPlease find attached your document.\n\nBest regards,`, {
        attachments: [pdf],
        name: "Dhara Global Solutions"
      });

    Logger.log(`Email sent to ${Email}`);
  });
}
