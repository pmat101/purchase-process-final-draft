function onFormSubmit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let currentRow = sheet.getLastRow();
  let formData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  sendMail (formData[4], currentRow, formData[2], formData[6], formData[7], formData[8]);
}

function response (uid, say, comment) {
  console.log(uid);
  console.log(typeof(uid));
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i=0; i<formData.length; i++) {
    if (formData[i] === "") {
      sheet.getRange(uid, i + 1).setValue(say);
      sheet.getRange(uid, i + 2).setValue(comment);
      formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
      break;
    }
  }
  if (formData[11] == "" && say == "Accepted") {    // ** formData length is constant **
    sendMail (formData[5], uid, formData[2], formData[6], formData[7], formData[8]);
  }
  else if (formData[11] != "" && say == "Accepted") {
    let recipient = "pranav.mathur@perfactgroup.in";
    let subject = "New Office Essentials Purchase Request";
    let name = "Purchase Request";
    let body = `
    <head>
      <style>
        a {
          text-decoration: none;
          font-size: 1.1em;
          padding: 0.5em 1em;
          border: 1px solid #888;
          border-radius: 1em;
        }
      </style>
    </head>
    <body>
      <p>Dear Team,</p>
      <p>A new purchase request has been submitted.</p>
      <p>Employee Name: ${formData[2]}</p>
      <p>Item Name: ${formData[6]}</p>
      <p>Amount: ${formData[7]}</p>
      <p>Reason: ${formData[8]}</p>
      <p>Team Head Response: ${formData[10]}</p>
      <p>Business Head Response: ${comment}</p>
      <p>Please process this request.</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(recipient, subject, body, {
      htmlBody: body,
      name: name
    });
  }
  else if (say == "Rejected") {
    let employee = formData[3];
    let subject = "Purchase Request Rejected";
    let name = "Purchase Request";
    let body = `
    <head>
      <style>
        a {
          text-decoration: none;
          font-size: 1.1em;
          padding: 0.5em 1em;
          border: 1px solid #888;
          border-radius: 1em;
        }
      </style>
    </head>
    <body>
      <p>Dear Sir/Madam,</p>
      <p>Your purchase request for ${formData[6]} has not been accepted by your manager, for below given reason.</p>
      <p>${comment}</p>
      <p>For further clarifications please consult your manager</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(employee, subject, body, {
      htmlBody: body,
      name: name
    });
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("response");
}

function sendMail (reciever, uid, ename, item, amount, reason) {
  let recipient = reciever;
  let subject = "New Office Essentials Purchase Request";
  let name = "Purchase Request";
  let body = `
  <head>
    <style>
      a {
        text-decoration: none;
        font-size: 1.1em;
        padding: 0.5em 1em;
        border: 1px solid #888;
        border-radius: 1em;
      }
    </style>
  </head>
  <body>
    <p>Dear Sir/Madam,</p>
    <p>A new purchase request has been submitted.</p>
    <p>Employee Name: ${ename}</p>
    <p>Item Name: ${item}</p>
    <p>Amount: ${amount}</p>
    <p>Reason: ${reason}</p>
    <span>Do you approve this request?</span>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbza3fxKOcuFPL0FXCIvZ9giPvaggcmA9iGOOrYrcr9eGfMzM7CgZ8NBVw1g0_kt-SF_/exec?uid=${uid}">Respond</a>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name
  });
}
