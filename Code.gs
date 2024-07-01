var spreadsheetId = "YOUR_SPREADSHEET_ID";
var spreadsheetUrl = "YOUR_SPREADSHEET_URL";

var sheetName = "Sheet1";

function main() {
 try {
    var sheet = SpreadsheetApp.openByUrl(spreadsheetUrl); 
    if (!sheet) {
      throw new Error("Sheet not found or cannot be accessed.");
    }
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var name = fullName(data[i][0]);
    var email = emailClean(data[i][1]);
        console.log('processing', i, name, email);

    var title = "YOUR_EMAIL_TITLE";
    var subtitle = "YOUR_EMAIL_SUBTITLE";
    var line1 = "YOUR_FIRST_BODY_LINE";
    var line2 = "YOUR_SECOND_BODY_LINE";
    var link = "YOUR_LINK_HERE";

    sendEmail(email, name, title, subtitle, line1, line2, link);
        console.log('sent', i, name, email);

  }
  } catch (error) {
    console.error("Error accessing Google Sheet:", error.message);
  }
}

function fullName(name) {
   const names = name.toLowerCase().trim().split(/\s+/);

  for (let i = 0; i < names.length; i++) {
    names[i] = names[i].charAt(0).toUpperCase() + names[i].slice(1);
  }
  var s = names[0] + " " + names[names.length - 1];
  return s;
}

function emailClean(text) {
  const email = text.toLowerCase().trim();
  return email;
}

function sendEmail(to, name, title, subtitle, line1, line2, link) {
  var emailFooter = DriveApp.getFileById("YOUR_FOOTER_IMAGE_ID").getAs("image/png");
  var emailHeader = DriveApp.getFileById("YOUR_HEADER_IMAGE_ID").getAs("image/png");
  var background_image = DriveApp.getFileById("YOUR_BACKGROUND_IMAGE_ID").getAs("image/png");

  var emailHeaderId = "YOUR_HEADER_IMAGE_ID";
  var emailFooterId = "YOUR_FOOTER_IMAGE_ID";
  var background_imageId = "YOUR_BACKGROUND_IMAGE_ID";

  var emailHeaderUrl = DriveApp.getFileById(emailHeaderId).getDownloadUrl();
  var emailFooterUrl = DriveApp.getFileById(emailFooterId).getDownloadUrl();
  var background_imageUrl = DriveApp.getFileById(background_imageId).getDownloadUrl();

  var emailImages = {
    "email_header": emailHeader,
    "email_footer": emailFooter
  };

  let emailBody = HtmlService.createTemplateFromFile("template");
  emailBody.title = subtitle; 
  emailBody.subtitle = "";
  emailBody.body_line1 = line1;
  emailBody.body_line2 = line2;
  emailBody.name = name;
  emailBody.link = link;
  emailBody.email_header = emailHeaderUrl;
  emailBody.email_footer = emailFooterUrl; 
  emailBody.background_image = background_imageUrl;

  const htmlBody = emailBody.evaluate().getContent();

  MailApp.sendEmail({
    to: to,
    subject: title,
    htmlBody: htmlBody,
  });
}
