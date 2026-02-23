/**
 * Google Apps Script Backend for Gamified Leaderboard
 * Author: Jules (AI Assistant)
 * Description: Handles data fetching and updating for the real-time leaderboard.
 * Sheets Structure: Column A: ID (Num), B: Name (String), C: Score (Num), D: PhotoURL (String)
 */

/**
 * Handle CORS preflight requests
 */
function doOptions(e) {
  return createJsonResponse({ "status": "success" });
}

/**
 * Handle GET requests - Fetches all students
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0]; // First sheet (Sheet1)
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return createJsonResponse({ "status": "success", "data": [] });
    }

    var dataRows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var students = dataRows.map(function(row) {
      return {
        "id": parseInt(row[0]),
        "name": String(row[1]),
        "score": parseInt(row[2]) || 0,
        "photoUrl": String(row[3])
      };
    });

    return createJsonResponse({ "status": "success", "data": students });
  } catch (error) {
    return createJsonResponse({ "status": "error", "message": error.message });
  }
}

/**
 * Handle POST requests - Updates score
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === "updateScore") {
      var id = parseInt(data.id);
      var newTotalScore = parseInt(data.newTotalScore);

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheets()[0];
      var lastRow = sheet.getLastRow();

      if (lastRow < 2) throw new Error("Sheet is empty");

      var idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var foundRowIndex = -1;

      for (var i = 0; i < idColumn.length; i++) {
        if (parseInt(idColumn[i][0]) === id) {
          foundRowIndex = i + 2; // +2 because header is 1 and array index starts at 0
          break;
        }
      }

      if (foundRowIndex !== -1) {
        sheet.getRange(foundRowIndex, 3).setValue(newTotalScore); // Column C is index 3
        return createJsonResponse({
          "status": "success",
          "updatedId": id,
          "newScore": newTotalScore
        });
      } else {
        throw new Error("Student ID not found: " + id);
      }
    } else {
      throw new Error("Invalid action: " + data.action);
    }
  } catch (error) {
    return createJsonResponse({ "status": "error", "message": error.message });
  }
}

/**
 * Helper to create consistent JSON responses.
 * Note: Google Apps Script Web Apps handle CORS automatically for GET/POST.
 * TextOutput does not support .addHeader().
 */
function createJsonResponse(response) {
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
