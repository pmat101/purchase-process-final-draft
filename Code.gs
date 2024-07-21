function onFormSubmit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let currentRow = sheet.getLastRow();
  let formData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  mailtoTH (formData[2], formData[4], formData[6], formData[7], formData[8], formData[9], formData[10], formData[12], formData[13], currentRow);
}

function thResponse (uid, say, comment) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(uid, 16+1).setValue(say);    // ** formData length is constant **
  sheet.getRange(uid, 17+1).setValue(comment);
  if (say == "Accepted") {
    mailtoWG (formData[2], formData[4], formData[6], formData[7], formData[8], formData[9], formData[10], formData[12], formData[13], formData[15], formData[17], uid);
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

function wgResponse (uid, say, comment) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(uid, 18+1).setValue(say);    // ** formData length is constant **
  sheet.getRange(uid, 19+1).setValue(comment);
  if (say == "Accepted") {
    mailtoAcc (formData[2], formData[4], formData[6], formData[7], formData[8], formData[9], formData[10], formData[12], formData[13], formData[15], formData[17], formData[19]);
  }
  else if (say == "Rejected") {
    let employee = formData[3];
    let subject = "Purchase Request Rejected";
    let name = "Purchase Request";
    let cc = formData[4];
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
      <p>Your purchase request for ${formData[6]} has not been accepted by the Admin Council - Purchase WG, for below given reason.</p>
      <p>${comment}</p>
      <p>For further clarifications please consult the Naresh sir from Arctic team.</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(employee, subject, body, {
      htmlBody: body,
      name: name,
      cc: cc
    });
  }
}

function doGet(e) {
  if(e.parameter.res == "th") {
    return HtmlService.createHtmlOutputFromFile("thResponse");
  }
  else if(e.parameter.res == "wg") {
    return HtmlService.createHtmlOutputFromFile("wgResponse");
  }
}

function mailtoTH (ename, thID, item, amount, reason, ecode, team, company, reamrk, uid) {
  let recipient = thID;
  let subject = `Office Essential Purchase Request for ${item}`;
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
    <p><b>Name of the requestor: </b> &nbsp; ${ename}</p>
    <p><b>ECODE: </b> &nbsp; ${ecode}</p>
    <p><b>Team: </b> &nbsp; ${team}</p>
    <p><b>Concerned Company: </b> &nbsp; ${company}</p>
    <p><b>Item Name: </b> &nbsp; ${item}</p>
    <p><b>Amount: </b> &nbsp; ₹ ${amount}</p>
    <p><b>Reason: </b> &nbsp; ${reason}</p>
    <p><b>Remark: </b> &nbsp; ${reamrk}</p>
    <span>Do you approve this request?</span>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbww-E2OtoHRntY1KCtkS72Y8aQcXkzimDbAgbV8ZwkuQG7WZB3yp_YKu7LHOOfA4l34/exec?res=th&uid=${uid}">Respond</a>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name
  });
}

function mailtoWG (ename, thID, item, amount, reason, ecode, team, company, reamrk, thName, thComment, uid) {
  let recipient = "pranav.mathur@perfactgroup.in";   // change to Naresh ji
  let subject = `Office Essential Purchase Request for ${item}`;
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
    <p><b>Name of the requestor: </b> &nbsp; ${ename}</p>
    <p><b>ECODE: </b> &nbsp; ${ecode}</p>
    <p><b>Team: </b> &nbsp; ${team}</p>
    <p><b>TH Name: </b> &nbsp; ${thName}</p>
    <p><b>TH Email: </b> &nbsp; ${thID}</p>
    <p><b>Concerned Company: </b> &nbsp; ${company}</p>
    <p><b>Item Name: </b> &nbsp; ${item}</p>
    <p><b>Amount: </b> &nbsp; ₹ ${amount}</p>
    <p><b>Reason: </b> &nbsp; ${reason}</p>
    <p><b>Remark: </b> &nbsp; ${reamrk}</p>
    <p><b>This has been approved by TH, TH remark: </b> &nbsp; ${thComment}</p>
    <span>Do you approve this request?</span>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbww-E2OtoHRntY1KCtkS72Y8aQcXkzimDbAgbV8ZwkuQG7WZB3yp_YKu7LHOOfA4l34/exec?res=wg&uid=${uid}">Respond</a>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name
  });
}

function mailtoAcc (ename, thID, item, amount, reason, ecode, team, company, reamrk, thName, thComment, wgComment) {
  let recipient = "pranav.mathur@perfactgroup.in";   // change to Accounts team
  let subject = `Office Essential Purchase Request for ${item}`;
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
    <p>Dear team,</p>
    <p>A new purchase request has been submitted.</p>
    <p><b>Name of the requestor: </b> &nbsp; ${ename}</p>
    <p><b>ECODE: </b> &nbsp; ${ecode}</p>
    <p><b>Team: </b> &nbsp; ${team}</p>
    <p><b>TH Name: </b> &nbsp; ${thName}</p>
    <p><b>TH Email: </b> &nbsp; ${thID}</p>
    <p><b>Concerned Company: </b> &nbsp; ${company}</p>
    <p><b>Item Name: </b> &nbsp; ${item}</p>
    <p><b>Estimated Amount: </b> &nbsp; ₹ ${amount}</p>
    <p><b>Reason: </b> &nbsp; ${reason}</p>
    <p><b>Remark: </b> &nbsp; ${reamrk}</p>
    <p><b>This has been approved by TH, TH remark: </b> &nbsp; ${thComment}</p>
    <p><b>This has been approved by Purchase Working Group, WG remark: </b> &nbsp; ${wgComment}</p>
    <br>
    <p>Kindly process this request.</p>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name
  });
}
