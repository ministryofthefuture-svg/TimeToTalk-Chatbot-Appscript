/****************************************************
 *  TIME TO TALK — MATCHMAKING + CALENDAR + EMAIL
 ****************************************************/

function onChange(e) {
  Logger.log("🟡 onChange fired. Event object: " + JSON.stringify(e));
  if (!e || !e.changeType) return;

  const allowedEvents = ["INSERT_ROW", "OTHER"];
  if (!allowedEvents.includes(e.changeType)) return;

  Logger.log("🟢 Valid Jotform update: " + e.changeType);
  runMatchmaking();
}

function runMatchmaking() {
  Logger.log("🚀 runMatchmaking started");

  var SS_ID = "1W_Bsy2ZqulgCIT2wWbBoym5Nc4FaaU8bE3G2xrkdytE";
  var SHEET_NAME = "Form responses";

  var COL_PAIRED = 22;
  var COL_PARTNER_NAME = 23;
  var COL_PARTNER_EMAIL = 24;

  var COL_MATCH_METHOD = 7;
  var COL_NAME = 8;
  var COL_INTEREST = 9;
  var COL_EMAIL = 10;

  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var unpaired = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    var rawMethod = row[COL_MATCH_METHOD - 1];
    var normalizedMethod = (rawMethod || "").toString().trim().toLowerCase();
    var alreadyPaired = (sheet.getRange(r + 1, COL_PAIRED).getValue() || "").toString().trim();

    var name = row[COL_NAME - 1];
    var interest = row[COL_INTEREST - 1];
    var email = row[COL_EMAIL - 1];

    if (normalizedMethod.indexOf("matchmaking") === -1) continue;
    if (alreadyPaired !== "") continue;

    unpaired.push({
      row: r + 1,
      name: name,
      interest: interest,
      email: email
    });
  }

  if (unpaired.length < 2) {
    Logger.log("⛔ Not enough people to match.");
    return;
  }

  unpaired.sort((a, b) => a.row - b.row);

  var used = {};
  var pairs = [];

  for (var i = 0; i < unpaired.length; i++) {
    var A = unpaired[i];
    if (used[A.row]) continue;

    for (var j = i + 1; j < unpaired.length; j++) {
      var B = unpaired[j];
      if (used[B.row]) continue;

      if (A.interest !== B.interest) {
        pairs.push([A, B]);
        used[A.row] = true;
        used[B.row] = true;
        break;
      }
    }
  }

  if (pairs.length === 0) {
    Logger.log("⛔ No compatible pairs created.");
    return;
  }

  pairs.forEach(function(pair) {
    var A = pair[0];
    var B = pair[1];

    sheet.getRange(A.row, COL_PAIRED).setValue("Yes");
    sheet.getRange(A.row, COL_PARTNER_NAME).setValue(B.name);
    sheet.getRange(A.row, COL_PARTNER_EMAIL).setValue(B.email);

    sheet.getRange(B.row, COL_PAIRED).setValue("Yes");
    sheet.getRange(B.row, COL_PARTNER_NAME).setValue(A.name);
    sheet.getRange(B.row, COL_PARTNER_EMAIL).setValue(A.email);

    var meetLink = createMeetEvent(A, B);

    sendPartnerEmail({
      recipientEmail: A.email,
      recipientName: A.name,
      partnerName: B.name,
      partnerEmail: B.email,
      partnerInterest: B.interest,
      meetLink: meetLink
    });

    Utilities.sleep(800);

    sendPartnerEmail({
      recipientEmail: B.email,
      recipientName: B.name,
      partnerName: A.name,
      partnerEmail: A.email,
      partnerInterest: A.interest,
      meetLink: meetLink
    });

    Utilities.sleep(800);
  });

  Logger.log("🎉 Matchmaking + Calendar scheduling complete!");
}


/*************** GOOGLE CALENDAR EVENT ***************/
function createMeetEvent(A, B) {
  const calendarId = "primary";
  const timezone = "Asia/Jakarta";

  let date = new Date();
  let day = date.getDay();
  let diff = (5 - day + 7) % 7;
  if (diff === 0) diff = 7;
  date.setDate(date.getDate() + diff);
  date.setHours(20, 0, 0, 0);

  let start = date.toISOString();
  let endDate = new Date(date.getTime() + (60 * 60 * 1000));
  let end = endDate.toISOString();

  let event = {
    summary: `🧠 Time To Talk: ${A.name} & ${B.name}`,
    description: `Automatic Time To Talk match.\n\n${A.name} ↔ ${B.name}`,
    start: { dateTime: start, timeZone: timezone },
    end: { dateTime: end, timeZone: timezone },
    attendees: [
      { email: A.email },
      { email: B.email }
    ],
    conferenceData: {
      createRequest: {
        requestId: "ttt-" + new Date().getTime(),
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    }
  };

  var created = Calendar.Events.insert(event, calendarId, { conferenceDataVersion: 1 });
  return created.conferenceData.entryPoints[0].uri;
}


/*************** EMAIL FUNCTION — USING HTML TEMPLATE ***************/
function sendPartnerEmail(obj) {

  var subject = "💬 Your Time To Talk Match & Session Details";

  // Load HTML template
  var template = HtmlService.createTemplateFromFile("match_email_template");

  // Inject variables
  template.recipientName = obj.recipientName;
  template.partnerName = obj.partnerName;
  template.partnerEmail = obj.partnerEmail;
  template.partnerInterest = obj.partnerInterest;
  template.meetLink = obj.meetLink;

  // Generate final HTML
  var htmlBody = template.evaluate().getContent();

  MailApp.sendEmail({
    to: obj.recipientEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}
