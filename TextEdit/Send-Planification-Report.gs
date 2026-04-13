// ── Configuration OCP Exchange (EWS) ─────────────────────────
const OCP_EMAIL_EWS = 'm.elamraoui@ocpgroup.ma';
const EWS_URL_EWS   = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';

function getOcpPasswordEWS() {
  return PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD') || '';
}

function sendEmailEWS(to, cc, subject, htmlBody, attachments, senderName) {
  const attachList = attachments ? (Array.isArray(attachments) ? attachments : [attachments]) : [];
  const boundary   = 'boundary_mainana_' + Date.now();
  const htmlB64    = Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8);
  const subjB64    = Utilities.base64Encode(subject,  Utilities.Charset.UTF_8);

  var mimeParts = [
    'From: "' + senderName + '" <' + OCP_EMAIL_EWS + '>',
    'To: '   + to,
    cc ? 'Cc: ' + cc : null,
    'Subject: =?UTF-8?B?' + subjB64 + '?=',
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    htmlB64,
    ''
  ];

  attachList.forEach(function(blob) {
    var attachB64  = Utilities.base64Encode(blob.getBytes());
    var attachName = blob.getName();
    mimeParts = mimeParts.concat([
      '--' + boundary,
      'Content-Type: ' + (blob.getContentType() || 'application/pdf') + '; name="' + attachName + '"',
      'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="' + attachName + '"',
      '',
      attachB64,
      ''
    ]);
  });

  mimeParts.push('--' + boundary + '--');

  var mime    = mimeParts.filter(function(l){ return l !== null; }).join('\r\n');
  var mimeB64 = Utilities.base64Encode(mime, Utilities.Charset.UTF_8);

  var soap = '<?xml version="1.0" encoding="utf-8"?>'
    + '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'
    + ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
    + ' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">'
    + '<soap:Header><t:RequestServerVersion Version="Exchange2010_SP2"/></soap:Header>'
    + '<soap:Body><m:CreateItem MessageDisposition="SendAndSaveCopy">'
    + '<m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>'
    + '<m:Items><t:Message>'
    + '<t:MimeContent CharacterSet="UTF-8">' + mimeB64 + '</t:MimeContent>'
    + '</t:Message></m:Items>'
    + '</m:CreateItem></soap:Body></soap:Envelope>';

  var credentials = Utilities.base64Encode(OCP_EMAIL_EWS + ':' + getOcpPasswordEWS());
  var response = UrlFetchApp.fetch(EWS_URL_EWS, {
    method: 'post',
    contentType: 'text/xml; charset=utf-8',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'SOAPAction': 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
    },
    payload: soap,
    muteHttpExceptions: true
  });

  var respCode = response.getResponseCode();
  var respText = response.getContentText();
  if (respCode !== 200 || respText.indexOf('NoError') === -1) {
    throw new Error('EWS send failed (' + respCode + '): ' + respText.substring(0, 500));
  }
}

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
