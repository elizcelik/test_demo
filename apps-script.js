var SHEET_NAME = "Waitlist";
var RAW_SHEET_NAME = "Raw Submissions";


/* =========================================================
   PUBLIC WAITLIST SUBMISSION

   Writes straight to both Raw Submissions (the backup log)
   and Waitlist in the same request. No trigger, no separate
   sync step — the signup lands in Waitlist immediately.
   ========================================================= */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Missing request body");
    }

    var data = JSON.parse(e.postData.contents);

    // Admin requests are handled separately.
    if (data.action === "adminRead") {
      return handleAdminRead(data);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var submissionId =
      data.submissionId ||
      Utilities.getUuid();

    var receivedAt =
      new Date().toISOString();

    var timestamp =
      data.ts ||
      receivedAt;

    var waitlistStatus = "SAVED";
    var waitlistError = "";


    /* ---------------------------------------------------------
       WRITE TO WAITLIST FIRST

       If this fails for any reason, we still want the signup
       captured in Raw Submissions below rather than losing it.
       --------------------------------------------------------- */

    try {

      var waitlist =
        ss.getSheetByName(SHEET_NAME);

      if (!waitlist) {
        throw new Error("Waitlist sheet not found");
      }

      waitlist.appendRow([
        timestamp,
        data.first || "",
        data.last || "",
        data.email || "",
        data.phone || "",
        data.device || "",
        data.concern || "",
        data.source || "",
        data.agreePrerelease ? "Yes" : "No",
        data.agreeConsent ? "Yes" : "No"
      ]);

    } catch (waitlistErr) {

      waitlistStatus = "WAITLIST WRITE FAILED";

      waitlistError =
        waitlistErr && waitlistErr.message
          ? waitlistErr.message
          : String(waitlistErr);

      Logger.log(
        "WAITLIST WRITE FAILED: " + waitlistError
      );
    }


    /* ---------------------------------------------------------
       CAPTURE THE FULL SIGNUP IN RAW SUBMISSIONS

       This is our source of truth / recovery backup, and
       records whether the Waitlist write above succeeded.
       --------------------------------------------------------- */

    var rawSheet =
      ss.getSheetByName(RAW_SHEET_NAME) ||
      getOrCreateRawSheet(ss);

    rawSheet.appendRow([
      submissionId,        // A - submission id
      receivedAt,          // B - received at
      timestamp,           // C - timestamp
      data.first || "",    // D - first name
      data.last || "",     // E - last name
      data.email || "",    // F - email
      data.phone || "",    // G - phone
      data.device || "",   // H - device
      data.concern || "",  // I - health concern
      data.source || "",   // J - how they heard
      data.agreePrerelease ? "Yes" : "No", // K - pre-release ack
      data.agreeConsent ? "Yes" : "No",    // L - 18+ consent
      waitlistStatus,      // M - waitlist status
      "SKIPPED",           // N - substack status
      waitlistError        // O - error
    ]);

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      captured: true,
      submissionId: submissionId
    });


  } catch (err) {

    Logger.log(
      "WAITLIST CAPTURE FAILED: " +
      (err && err.message ? err.message : err)
    );

    return jsonResponse({
      success: false,
      captured: false,
      error:
        err && err.message
          ? err.message
          : "Unable to capture signup"
    });
  }
}



/* =========================================================
   CREATE RAW SUBMISSIONS TAB IF IT DOESN'T EXIST
   ========================================================= */

function getOrCreateRawSheet(ss) {

  var sheet =
    ss.getSheetByName(
      RAW_SHEET_NAME
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        RAW_SHEET_NAME
      );

    sheet.appendRow([
      "submission id",
      "received at",
      "timestamp",
      "first name",
      "last name",
      "email",
      "phone",
      "device",
      "health concern",
      "how they heard",
      "pre-release ack",
      "18+ consent",
      "waitlist status",
      "substack status",
      "error"
    ]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}



/* =========================================================
   ADMIN READ
   Used by /api/admin/waitlist
   ========================================================= */

function handleAdminRead(data) {

  var expectedSecret =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "ADMIN_READ_SECRET"
      );


  if (!expectedSecret) {

    return jsonResponse({
      success: false,
      error:
        "ADMIN_READ_SECRET is not configured"
    });
  }


  if (
    !data.secret ||
    data.secret !== expectedSecret
  ) {

    return jsonResponse({
      success: false,
      error: "Unauthorized"
    });
  }


  try {

    var ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    var sheet =
      ss.getSheetByName(
        SHEET_NAME
      );


    if (!sheet) {

      return jsonResponse({
        success: false,
        error:
          "Waitlist sheet not found"
      });
    }


    var values =
      sheet
        .getDataRange()
        .getValues();


    if (!values.length) {

      return jsonResponse({
        success: true,
        headers: [],
        rows: []
      });
    }


    var headers =
      values[0];

    var rows =
      values.slice(1);


    return jsonResponse({
      success: true,
      headers: headers,
      rows: rows
    });


  } catch (err) {

    return jsonResponse({
      success: false,
      error:
        err && err.message
          ? err.message
          : "Admin read failed"
    });
  }
}



/* =========================================================
   SIMPLE WEB APP HEALTH CHECK

   Opening the /exec URL in a browser should show:
   {"status":"ok"}
   ========================================================= */

function doGet() {

  return jsonResponse({
    status: "ok"
  });
}



/* =========================================================
   JSON RESPONSE HELPER
   ========================================================= */

function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
