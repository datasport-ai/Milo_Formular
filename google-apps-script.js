// ============================================================
// Google Apps Script — Milo Feedback Forms
// ============================================================
// INSTRUCTIONS DE DÉPLOIEMENT :
// 1. Extensions → Apps Script
// 2. Coller ce code (remplacer tout le contenu existant)
// 3. Déployer → Nouveau déploiement → Web App
//    - Exécuter en tant que : Moi
//    - Accès : Tout le monde (anonyme)
// 4. Copier l'URL du déploiement (se termine par /exec)
// 5. Remplacer APPS_SCRIPT_URL dans les 3 fichiers HTML
//
// NOTE CORS : les requêtes utilisent Content-Type: text/plain
// pour éviter le preflight OPTIONS — Google Apps Script
// ajoute automatiquement Access-Control-Allow-Origin: *
// sur les réponses des web apps déployées publiquement.
// ============================================================

var SPREADSHEET_ID = '1O-dJv9CNcgidMwbd23dw5fErdzZuRW1MWrIDFtMFBsA';
var RESPONSES_SHEET_NAME = 'Responses';

function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(RESPONSES_SHEET_NAME);
      sheet.appendRow(['timestamp', 'formType', 'email', 'data']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }

    var timestamp = new Date().toISOString();
    var formType = data.formType || '';
    var email = data.email || '';

    var payload = {};
    for (var key in data) {
      if (key !== 'formType' && key !== 'email') {
        payload[key] = data[key];
      }
    }

    sheet.appendRow([timestamp, formType, email, JSON.stringify(payload)]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
