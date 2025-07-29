// Record Offset is the row to start with, starting at 1
function sendMailMerge(
  subjectQuery,
  recipientsColumn,
  filterValue = false,
  uniqueColumn = false,
  recordOffset = 0
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const filterColumnName = "Status"; // <-- customize this if needed
  const uniqueColumnName = "GL Email"; // <-- only one email per value in this column

  const filterColumnIndex = headers.indexOf(filterColumnName);
  const uniqueColumnIndex = headers.indexOf(uniqueColumnName);

  if (filterValue && filterColumnIndex === -1) {
    Logger.log(`Filter column "${filterColumnName}" not found.`);
    return;
  }

  if (uniqueColumn && uniqueColumnIndex === -1) {
    Logger.log(`Unique column "${uniqueColumnName}" not found.`);
    return;
  }

  // Find the first draft with matching subject (or partial match)
  const drafts = GmailApp.getDrafts();
  let selectedDraft = null;

  for (let i = 0; i < drafts.length; i++) {
    const msg = drafts[i].getMessage();
    if (msg.getSubject().includes(subjectQuery)) {
      selectedDraft = msg;
      break;
    }
  }

  if (!selectedDraft) {
    Logger.log("No draft found with subject containing: " + subjectQuery);
    return;
  }

  const subjectTemplate = selectedDraft.getSubject();
  const bodyTemplate = selectedDraft.getBody();

  const seen = new Set();

  for (let i = recordOffset + 1; i < data.length; i++) {
    const isHidden =
      sheet.isRowHiddenByFilter(i + 1) || sheet.isRowHiddenByUser(i + 1);
    if (isHidden) {
      Logger.log("Skipping hidden " + data[i]);
      continue;
    }

    const row = data[i];

    if (filterValue && row[filterColumnIndex] !== filterValue) {
      Logger.log(
        `Skipping filtered for ${filterValue}. Value: ` + row[filterColumnIndex]
      );
      continue;
    }

    const uniqueValue = row[uniqueColumnIndex];
    // Logger.log("seen " + Array.from(seen).join(", "))
    // Logger.log("uniqueValue " + uniqueValue)
    // Logger.log("seen.has(uniqueValue) " +  seen.has(uniqueValue))

    if (uniqueColumn && seen.has(uniqueValue)) {
      // Logger.log(`Skipping duplicate for ${uniqueColumnName}: ${uniqueValue}`);
      continue;
    }
    seen.add(uniqueValue);

    let emailBody = bodyTemplate;
    let emailSubject = subjectTemplate;
    let recipient = "";

    for (let j = 0; j < headers.length; j++) {
      const header = headers[j];
      const value = row[j];
      if (header.toLowerCase() === recipientsColumn.toLowerCase())
        recipient = value;

      const placeholder = new RegExp(`{{\\s*${header}\\s*}}`, "g");
      emailBody = emailBody.replace(placeholder, value);
      emailSubject = emailSubject.replace(placeholder, value);
    }

    if (recipient) {
      Logger.log(i + ": pretentding to create draft for " + recipient);
      // Logger.log("Subject: " + emailSubject);
      // Logger.log(i + ": Creating draft for: " + recipient);
      // GmailApp.createDraft(recipient, emailSubject, "", {
      //   htmlBody: emailBody
      // });
    } else {
      Logger.log(i + ": skipping, no recipient: " + row[i]);
    }
  }
}

function testSendMailMerge() {
  // sendMailMerge("Thank you for mentoring a New", "Mentor Email", "New");
  // sendMailMerge("Thank you for leading a CCP Group", "GL Email", false, true)
  // sendMailMerge("Thank you for mentoring an Ongoing", "Mentor Email", "Ongoing");
  // sendMailMerge("Welcome back", "Email", "Ongoing");
  // sendMailMerge("Welcome to the Fellowship", "Email", "New");
}
