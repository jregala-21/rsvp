/**
 * EMPOWERED RSVP
 * Google Apps Script Backend
 *
 * Features:
 * - Saves RSVP submissions to Google Sheets
 * - Provides RSVP data to dashboard
 * - Returns totals for:
 *      - Total responses
 *      - In Person
 *      - Online
 *
 * Spreadsheet tab:
 * RSVP Responses
 */

const SHEET_NAME = "RSVP Responses";


/**
 * Handles GET requests.
 *
 * Example:
 * /exec?action=list
 *
 * Used by dashboard.html.
 */
function doGet(e) {

  try {

    const action = String(
      (e && e.parameter && e.parameter.action) || ""
    ).toLowerCase();


    if (action === "list") {

      return getRsvpData_();

    }


    return jsonResponse_({

      ok: true,

      message: "EMPOWERED RSVP API is running."

    });


  } catch (error) {

    console.error(error);


    return jsonResponse_({

      ok: false,

      message: "Server error while reading RSVP data."

    });

  }

}


/**
 * Handles RSVP submissions.
 *
 * Used by index.html.
 */
function doPost(e) {

  try {

    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();


    let sheet =
      spreadsheet.getSheetByName(SHEET_NAME);


    if (!sheet) {

      sheet =
        spreadsheet.insertSheet(SHEET_NAME);

    }


    ensureHeader_(sheet);


    const parameters =
      e && e.parameter
        ? e.parameter
        : {};


    const name =
      cleanValue_(parameters.name);


    const contact =
      cleanValue_(parameters.contact);


    const attendance =
      cleanValue_(parameters.attendance);


    const service =
      cleanValue_(parameters.service);


    const submittedAt =
      cleanValue_(parameters.submittedAt);


    const source =
      cleanValue_(parameters.source);


    /**
     * Validate required fields.
     */
    if (
      !name ||
      !contact ||
      !attendance
    ) {

      return jsonResponse_({

        ok: false,

        message:
          "Missing required RSVP fields."

      });

    }


    /**
     * Validate attendance value.
     */
    const allowedAttendance = [
      "In Person",
      "Online"
    ];


    if (
      !allowedAttendance.includes(attendance)
    ) {

      return jsonResponse_({

        ok: false,

        message:
          "Invalid attendance option."

      });

    }


    /**
     * Add RSVP to Google Sheet.
     */
    sheet.appendRow([

      new Date(),

      safeCell_(name),

      safeCell_(contact),

      attendance,

      service,

      submittedAt,

      source

    ]);


    return jsonResponse_({

      ok: true,

      message:
        "RSVP saved successfully."

    });


  } catch (error) {

    console.error(error);


    return jsonResponse_({

      ok: false,

      message:
        "Server error while saving RSVP."

    });

  }

}


/**
 * Reads RSVP data for dashboard.
 */
function getRsvpData_() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    spreadsheet.getSheetByName(SHEET_NAME);


  /**
   * No RSVP sheet yet.
   */
  if (!sheet) {

    return emptyDashboardResponse_();

  }


  const lastRow =
    sheet.getLastRow();


  /**
   * Header only or empty.
   */
  if (lastRow < 2) {

    return emptyDashboardResponse_();

  }


  /**
   * Get all RSVP rows.
   *
   * Columns:
   *
   * A = Received At
   * B = Full Name
   * C = Contact
   * D = Attendance
   * E = Service
   * F = Client Submitted At
   * G = Source
   */
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getValues();


  const records =
    values

      .filter(function (row) {

        return row.some(function (cell) {

          return (
            String(cell).trim() !== ""
          );

        });

      })

      .map(function (row) {

        return {

          receivedAt:
            formatDate_(row[0]),

          name:
            String(row[1] || ""),

          contact:
            String(row[2] || ""),

          attendance:
            String(row[3] || ""),

          service:
            String(row[4] || ""),

          submittedAt:
            String(row[5] || ""),

          source:
            String(row[6] || "")

        };

      })

      /**
       * Latest RSVP appears first.
       */
      .reverse();


  /**
   * Calculate totals.
   */
  const total =
    records.length;


  const inPerson =
    records.filter(function (record) {

      return normalize_(
        record.attendance
      ) === "in person";

    }).length;


  const online =
    records.filter(function (record) {

      return normalize_(
        record.attendance
      ) === "online";

    }).length;


  return jsonResponse_({

    ok: true,

    summary: {

      total: total,

      inPerson: inPerson,

      online: online

    },

    records: records,

    updatedAt:
      new Date().toISOString()

  });

}


/**
 * Returns empty dashboard data.
 */
function emptyDashboardResponse_() {

  return jsonResponse_({

    ok: true,

    summary: {

      total: 0,

      inPerson: 0,

      online: 0

    },

    records: [],

    updatedAt:
      new Date().toISOString()

  });

}


/**
 * Creates the spreadsheet header
 * if the sheet is empty.
 */
function ensureHeader_(sheet) {

  if (
    sheet.getLastRow() !== 0
  ) {

    return;

  }


  const headers = [

    "Received At",

    "Full Name",

    "Contact",

    "Attendance",

    "Service",

    "Client Submitted At",

    "Source"

  ];


  sheet.appendRow(headers);


  /**
   * Freeze header row.
   */
  sheet.setFrozenRows(1);


  /**
   * Header styling.
   */
  const headerRange =
    sheet.getRange(
      1,
      1,
      1,
      headers.length
    );


  headerRange
    .setFontWeight("bold")
    .setBackground("#8F2BB5")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");


  /**
   * Column widths.
   */
  sheet.setColumnWidth(
    1,
    170
  );


  sheet.setColumnWidth(
    2,
    200
  );


  sheet.setColumnWidth(
    3,
    220
  );


  sheet.setColumnWidth(
    4,
    130
  );


  sheet.setColumnWidth(
    5,
    180
  );


  sheet.setColumnWidth(
    6,
    190
  );


  sheet.setColumnWidth(
    7,
    220
  );

}


/**
 * Removes unnecessary spaces
 * and limits input size.
 */
function cleanValue_(value) {

  return String(
    value == null
      ? ""
      : value
  )
    .trim()
    .slice(
      0,
      500
    );

}


/**
 * Prevent spreadsheet formula injection.
 *
 * Example:
 *
 * =SUM(...)
 * +123
 * -123
 * @something
 */
function safeCell_(value) {

  const text =
    String(value || "");


  if (
    /^[=+\-@]/.test(text)
  ) {

    return "'" + text;

  }


  return text;

}


/**
 * Normalize text for comparison.
 */
function normalize_(value) {

  return String(
    value || ""
  )
    .trim()
    .toLowerCase();

}


/**
 * Convert spreadsheet dates
 * to ISO format.
 */
function formatDate_(value) {

  if (
    value instanceof Date
  ) {

    return value.toISOString();

  }


  return String(
    value || ""
  );

}


/**
 * Return JSON response.
 */
function jsonResponse_(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}