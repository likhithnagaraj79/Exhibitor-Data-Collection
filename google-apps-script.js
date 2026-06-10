// ─────────────────────────────────────────────────────────────────────────────
// Exhibitor Registration — Google Apps Script Backend
//
// HOW TO DEPLOY:
//  1. Open https://script.google.com  →  New Project
//  2. Paste this entire file into the editor (replace the default code)
//  3. Save (Ctrl+S / Cmd+S)
//  4. Click Deploy → New deployment
//  5. Type: Web app
//  6. Execute as: Me  |  Who has access: Anyone
//  7. Click Deploy → Authorise (follow the OAuth prompts)
//  8. Copy the Web App URL shown — paste it into index.html as SCRIPT_URL
//
// GOOGLE SHEET STRUCTURE (auto-created on first submission):
//  Col A  : Timestamp
//  Col B  : Company Name
//  Col C  : No. of Co-Exhibitors
//  Col D  : Co-Ex 1 Name
//  Col E  : Co-Ex 1 Designation
//  Col F  : Co-Ex 2 Name
//  Col G  : Co-Ex 2 Designation
//  …      : (up to 9 co-exhibitors = columns D–U)
// ─────────────────────────────────────────────────────────────────────────────

const MAX_CO_EX = 9;

// Called automatically by Google on every POST request from the form.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Auto-create header row the very first time.
    if (sheet.getLastRow() === 0) {
      const headers = ['Timestamp', 'Company Name', 'No. of Co-Exhibitors'];
      for (let i = 1; i <= MAX_CO_EX; i++) {
        headers.push('Co-Ex ' + i + ' Name', 'Co-Ex ' + i + ' Designation');
      }
      sheet.appendRow(headers);

      // Basic header formatting
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      sheet.setFrozenRows(1);
    }

    // Build data row: fixed columns + up to MAX_CO_EX co-exhibitor pairs.
    const row = [
      new Date(),                  // A: Timestamp
      data.companyName,            // B: Company Name
      data.numCoExhibitors         // C: No. of Co-Exhibitors
    ];

    for (let i = 0; i < MAX_CO_EX; i++) {
      if (i < data.coExhibitors.length) {
        row.push(data.coExhibitors[i].name, data.coExhibitors[i].designation);
      } else {
        row.push('', '');           // Empty cells for unused co-exhibitor slots
      }
    }

    sheet.appendRow(row);

    // Auto-resize columns for readability (runs fast, no noticeable delay).
    sheet.autoResizeColumns(1, row.length);

    return respond({ status: 'success' });

  } catch (err) {
    return respond({ status: 'error', message: err.message });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Optional: test this function manually in the Apps Script editor.
// It will write a sample row to your sheet without needing an HTTP request.
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        companyName:     'Acme Corp',
        numCoExhibitors: 3,
        coExhibitors: [
          { name: 'Alice Johnson',  designation: 'CEO'           },
          { name: 'Bob Smith',      designation: 'Sales Manager' },
          { name: 'Carol Williams', designation: 'Technician'    }
        ]
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
