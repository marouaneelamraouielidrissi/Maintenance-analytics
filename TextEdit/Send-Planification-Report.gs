// OCP_EMAIL_EWS, EWS_URL_EWS, getOcpPasswordEWS() et sendEmailEWS() sont définis dans Envoi-AVIS-MEC-INST.gs

// ═══════════════════════════════════════════════════════════════
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function Envoyer_Rapport_Planification_PDF() {
  const SPREADSHEET_ID  = "1EBACM8ou8B_9fmExToUKsMCvHL27hiwU2D0yZ_gQGOA";
  const SHEET_NAME      = "Vesrion imprimable";
  const ROWS_FOR_2_PAGES = 112;

  const feuille = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = feuille.getSheetByName(SHEET_NAME);

  const dateActuelle = new Date();
  const mois    = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "MMMM");
  const annee   = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "yyyy");
  const semaine = String(getWeekNumber(dateActuelle)).padStart(2, "0");

  const lastRow = sheet.getLastRow();
  if (lastRow > ROWS_FOR_2_PAGES) {
    sheet.hideRows(ROWS_FOR_2_PAGES + 1, lastRow - ROWS_FOR_2_PAGES);
  }

  const options = {
    exportFormat: "pdf", format: "pdf", size: "A4", landscape: false,
    sheetnames: false, printtitle: false, pagenumbers: false, gridlines: false,
    fzr: false, gid: sheet.getSheetId()
  };

  const url = 'https://docs.google.com/spreadsheets/d/' + feuille.getId() + '/export?' +
    Object.entries(options).map(function(kv){ return kv[0] + '=' + kv[1]; }).join('&');

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob     = response.getBlob().setName(
    'Rapport Mensuel de planification - ' + mois + ' ' + annee + '.pdf'
  );

  const corpsMessage = `
    <div style="font-family:Arial,sans-serif;color:#0A1E3F;font-size:14px;">
      <p>Bonjour,</p>
      <p>Veuillez trouver en pièce jointe <strong>le rapport mensuel de planification des travaux</strong>
      du mois de <strong>${mois} ${annee}</strong>.</p>
      <p>Cordialement,</p>
      <div style="font-family:'Times New Roman',serif;font-size:14px;color:#002060;line-height:1.5;">
        <span style="font-weight:bold;">Marouane ELAMRAOUI</span><br>
        <span style="color:#c55a11;">Méthode de Maintenance</span><br>
        <span style="font-weight:bold;">OCP SA - Khouribga</span><br>
        <span style="color:green;">Tél. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:green;">Cisco :</span> 8103388<br>
        <a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>
      </div>
    </div>`;

  sendEmailEWS(
    OCP_EMAIL_EWS,
    OCP_EMAIL_EWS,
    'Rapport Hebdomadaire de planification - S' + semaine + ' - ' + mois + ' ' + annee,
    corpsMessage,
    blob,
    "Bureau de Méthode Daoui - Section Planification"
  );

  if (lastRow > ROWS_FOR_2_PAGES) {
    sheet.showRows(ROWS_FOR_2_PAGES + 1, lastRow - ROWS_FOR_2_PAGES);
  }
}

function tester_Rapport_Planification() {
  const SPREADSHEET_ID  = "1EBACM8ou8B_9fmExToUKsMCvHL27hiwU2D0yZ_gQGOA";
  const SHEET_NAME      = "Vesrion imprimable";
  const ROWS_FOR_2_PAGES = 112;

  const feuille = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = feuille.getSheetByName(SHEET_NAME);

  const dateActuelle = new Date();
  const mois  = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "MMMM");
  const annee = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "yyyy");

  const lastRow = sheet.getLastRow();
  if (lastRow > ROWS_FOR_2_PAGES) sheet.hideRows(ROWS_FOR_2_PAGES + 1, lastRow - ROWS_FOR_2_PAGES);

  const options = {
    exportFormat: "pdf", format: "pdf", size: "A4", landscape: false,
    sheetnames: false, printtitle: false, pagenumbers: false, gridlines: false,
    fzr: false, gid: sheet.getSheetId()
  };
  const url      = 'https://docs.google.com/spreadsheets/d/' + feuille.getId() + '/export?' +
    Object.entries(options).map(function(kv){ return kv[0] + '=' + kv[1]; }).join('&');
  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob     = response.getBlob().setName('[TEST] Rapport planification.pdf');

  if (lastRow > ROWS_FOR_2_PAGES) sheet.showRows(ROWS_FOR_2_PAGES + 1, lastRow - ROWS_FOR_2_PAGES);

  sendEmailEWS(
    OCP_EMAIL_EWS, null,
    '[TEST] Rapport Hebdomadaire de planification',
    '<p>[TEST] Email de vérification — rapport PDF ci-dessous.</p>',
    blob,
    "Bureau de Méthode Daoui - Section Planification"
  );
  Logger.log('✅ Test envoyé à : ' + OCP_EMAIL_EWS);
}
