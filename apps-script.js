var SHEET_NAME = "Waitlist";
var RAW_SHEET_NAME = "Raw Submissions";


/* =========================================================
   PUBLIC WAITLIST SUBMISSION
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


    /* ---------------------------------------------------------
       STEP 1
       CAPTURE THE FULL SIGNUP IN RAW SUBMISSIONS FIRST

       This is our source of truth / recovery backup.
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
      "CAPTURED",          // M - waitlist status
      "SKIPPED",           // N - substack status
      ""                   // O - error
    ]);

    // Force Google Sheets to commit the backup before
    // telling the website that signup succeeded.
    SpreadsheetApp.flush();


    /* ---------------------------------------------------------
       IMPORTANT

       Once Raw Submissions contains the signup,
       the person is safely captured.

       Do NOT wait for Waitlist sync.
       Do NOT wait for Substack.
       --------------------------------------------------------- */

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
   BACKGROUND SYNC
   RAW SUBMISSIONS → WAITLIST

   Set this function to run automatically every minute.
   ========================================================= */

function syncCapturedToWaitlist() {

  // Prevent two trigger runs from syncing at the same time.
  var lock =
    LockService.getScriptLock();

  try {

    lock.tryLock(25000);

    if (!lock.hasLock()) {
      return;
    }

    var ss =
      SpreadsheetApp.getActiveSpreadsheet();

    var raw =
      ss.getSheetByName(RAW_SHEET_NAME);

    var waitlist =
      ss.getSheetByName(SHEET_NAME);

    if (!raw || !waitlist) {
      return;
    }


    var rawValues =
      raw.getDataRange().getValues();

    if (rawValues.length <= 1) {
      return;
    }


    /* ---------------------------------------------------------
       BUILD A LIST OF PEOPLE ALREADY IN WAITLIST

       This prevents duplicate rows if a previous sync
       succeeded but failed before changing Raw status to SAVED.
       --------------------------------------------------------- */

    var existing = {};

    var waitlistLastRow =
      waitlist.getLastRow();

    if (waitlistLastRow > 1) {

      var waitlistValues =
        waitlist
          .getRange(
            2,
            1,
            waitlistLastRow - 1,
            Math.max(waitlist.getLastColumn(), 10)
          )
          .getValues();

      for (
        var w = 0;
        w < waitlistValues.length;
        w++
      ) {

        var existingTimestamp =
          String(
            waitlistValues[w][0] || ""
          ).trim();

        var existingEmail =
          String(
            waitlistValues[w][3] || ""
          )
            .trim()
            .toLowerCase();

        if (
          existingTimestamp ||
          existingEmail
        ) {

          existing[
            existingTimestamp +
            "|" +
            existingEmail
          ] = true;
        }
      }
    }


    /* ---------------------------------------------------------
       PROCESS CAPTURED / RETRY ROWS
       --------------------------------------------------------- */

    for (
      var i = 1;
      i < rawValues.length;
      i++
    ) {

      var row =
        rawValues[i];

      var status =
        String(row[12] || "").trim();

      if (
        status !== "CAPTURED" &&
        status !== "NEEDS RETRY"
      ) {
        continue;
      }


      var timestamp =
        row[2];

      var first =
        row[3];

      var last =
        row[4];

      var email =
        row[5];

      var phone =
        row[6];

      var device =
        row[7];

      var concern =
        row[8];

      var source =
        row[9];

      var agreePrerelease =
        row[10];

      var agreeConsent =
        row[11];


      var duplicateKey =
        String(timestamp || "").trim() +
        "|" +
        String(email || "")
          .trim()
          .toLowerCase();


      try {

        /* ---------------------------------------------
           If this exact signup already exists,
           don't add it twice.
           --------------------------------------------- */

        if (existing[duplicateKey]) {

          raw
            .getRange(i + 1, 13)
            .setValue("SAVED");

          raw
            .getRange(i + 1, 15)
            .setValue("");

          continue;
        }


        /* ---------------------------------------------
           WRITE TO NORMAL WAITLIST
           --------------------------------------------- */

        waitlist.appendRow([
          timestamp,
          first,
          last,
          email,
          phone,
          device,
          concern,
          source,
          agreePrerelease,
          agreeConsent
        ]);


        existing[duplicateKey] =
          true;


        /* ---------------------------------------------
           MARK RAW BACKUP AS SUCCESSFULLY SYNCED
           --------------------------------------------- */

        raw
          .getRange(i + 1, 13)
          .setValue("SAVED");

        raw
          .getRange(i + 1, 15)
          .setValue("");


      } catch (err) {

        /* ---------------------------------------------
           NEVER DELETE THE RAW SIGNUP.

           Mark it for another retry instead.
           --------------------------------------------- */

        raw
          .getRange(i + 1, 13)
          .setValue("NEEDS RETRY");

        raw
          .getRange(i + 1, 15)
          .setValue(
            err && err.message
              ? err.message
              : String(err)
          );
      }
    }


    SpreadsheetApp.flush();


  } finally {

    try {
      lock.releaseLock();
    } catch (e) {
      // Nothing to do.
    }
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
