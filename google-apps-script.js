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
// STRUCTURE DES ONGLETS :
//   Responses_active    — soumissions form-active.html
//   Responses_inactive  — soumissions form-inactive.html
//   Responses_dropout   — soumissions form-onboarding-dropout.html
// ============================================================

var SPREADSHEET_ID = '1O-dJv9CNcgidMwbd23dw5fErdzZuRW1MWrIDFtMFBsA';

// Colonnes par formulaire : { field: clé dans data, header: libellé colonne }
var SHEET_CONFIG = {
  'active': {
    name: 'Responses_active',
    columns: [
      { field: 'timestamp',  header: 'Timestamp' },
      { field: 'email',      header: 'Email' },
      { field: 'nps',        header: 'NPS (0-10)' },
      { field: 'q1',         header: 'Coach aide (Q1)' },
      { field: 'q1_follow',  header: 'Q1 – Détail' },
      { field: 'q2',         header: 'À améliorer (Q2)' },
      { field: 'q4',         header: 'Description Milo (Q4)' }
    ]
  },
  'inactive': {
    name: 'Responses_inactive',
    columns: [
      { field: 'timestamp',  header: 'Timestamp' },
      { field: 'email',      header: 'Email' },
      { field: 'q1',         header: 'Raison inactivité (Q1)' },
      { field: 'q1_follow',  header: 'Q1 – Détail' },
      { field: 'q2',         header: 'Motivation retour (Q2)' },
      { field: 'q3',         header: 'Commentaire libre (Q3)' }
    ]
  },
  'dropout': {
    name: 'Responses_dropout',
    columns: [
      { field: 'timestamp',     header: 'Timestamp' },
      { field: 'email',         header: 'Email' },
      { field: 'stepRaw',       header: 'Étape (clé)' },
      { field: 'stepLabel',     header: 'Étape (libellé)' },
      { field: 'reason',        header: 'Raison abandon' },
      { field: 'reasonDetail',  header: 'Raison – Détail' },
      { field: 'completeLater', header: 'Finira le profil ?' },
      { field: 'freeComment',   header: 'Commentaire libre' }
    ]
  }
};

function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);
    var formType = data.formType || '';

    var config = SHEET_CONFIG[formType];
    if (!config) {
      return jsonResponse({ success: false, error: 'formType inconnu : ' + formType });
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(config.name);

    if (!sheet) {
      sheet = ss.insertSheet(config.name);
      var headers = config.columns.map(function(c) { return c.header; });
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }

    data.timestamp = new Date().toISOString();

    var row = config.columns.map(function(c) {
      var val = data[c.field];
      return (val === undefined || val === null) ? '' : val;
    });

    sheet.appendRow(row);

    return jsonResponse({ success: true });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
