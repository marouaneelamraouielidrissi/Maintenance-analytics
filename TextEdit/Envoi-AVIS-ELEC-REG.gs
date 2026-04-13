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
function envoyer_AVIS_ELEC_REG() {
  const feuille   = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleID = feuille.getId();
  const token     = ScriptApp.getOAuthToken();

  const feuilles = [
    { nom: "Avis à envoyer électrique",  nomFichier: "AVIS_OUVERTS_ELECTRIQUE - " },
    { nom: "Avis à envoyer régulation",  nomFichier: "AVIS_OUVERTS_REGULATION - " }
  ];

  const piecesJointes = [];
  feuilles.forEach(function(feuilleInfo) {
    const feuilleOT = feuille.getSheetByName(feuilleInfo.nom);
    if (!feuilleOT) return;
    const url = "https://docs.google.com/spreadsheets/d/" + feuilleID + "/export?" +
      "format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false" +
      "&pagenumbers=false&gridlines=false&fzr=false&gid=" + feuilleOT.getSheetId();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    const blob = response.getBlob().setName(
      feuilleInfo.nomFichier +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyy") + ".pdf"
    );
    piecesJointes.push(blob);
  });

  const sheetPlan = feuille.getSheetByName("Indicateur planification");
  const values    = sheetPlan.getRange("A1:E9").getDisplayValues();
  let tableauHTML = '<table border="1" style="border-collapse:collapse;font-family:Arial;font-size:13px;">';
  values.forEach(function(row, i) {
    tableauHTML += '<tr>';
    row.forEach(function(cell) {
      if (i === 0) {
        tableauHTML += '<th style="background:#002060;color:white;padding:6px;text-align:center;">' + cell + '</th>';
      } else {
        tableauHTML += '<td style="padding:6px;text-align:center;">' + cell + '</td>';
      }
    });
    tableauHTML += '</tr>';
  });
  tableauHTML += '</table>';

  const to  = "boumazlag@ocpgroup.ma, h.bouelghellat@ocpgroup.ma, m.assouggane@ocpgroup.ma, redouane.zouaoui@ocpgroup.ma, e.fouzir@ocpgroup.ma, chinig@ocpgroup.ma, abdeljaouad.achlih@ocpgroup.ma, a.jebli2@ocpgroup.ma, aasli@ocpgroup.ma";
  const cc  = "m.mamouni@ocpgroup.ma, elkhyari@ocpgroup.ma, ibtissame.elkhloufi@ocpgroup.ma, wajid@ocpgroup.ma, a.dahmou@ocpgroup.ma, o.sirri@ocpgroup.ma, m.elamraoui@ocpgroup.ma";
  const sujet = "Liste des Avis ouverts (Électrique & Régulation) - " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const corpsMessage = `
    <div style="font-family:Arial,sans-serif;color:#0A1E3F;font-size:14px;">
      <p>Bonjour,</p>
      <p>Veuillez trouver en pièce jointe <strong>la liste des Avis actuellement ouverts</strong>.<br>
      Nous vous invitons à en prendre connaissance afin de procéder à leur <strong>approbation</strong> ou <strong>rejet</strong> selon le cas.</p>
      <p><strong>Tableau de suivi des avis par poste de travail :</strong></p>
      ${tableauHTML}
      <p>Cordialement,</p>
      <div style="font-family:'Times New Roman',serif;font-size:14px;color:#002060;line-height:1.5;">
        <span style="font-weight:bold;">Marouane ELAMRAOUI</span><br>
        <span style="color:#c55a11;">Méthode de Maintenance</span><br>
        <span style="font-weight:bold;">OCP SA - Khouribga</span><br>
        <span style="color:green;">Tél. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:green;">Cisco :</span> 8103388<br>
        <a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>
      </div>
    </div>`;

  sendEmailEWS(to, cc, sujet, corpsMessage, piecesJointes, "Bureau de méthode Daoui - Section Planification");
}

function tester_AVIS_ELEC_REG() {
  const feuille   = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleID = feuille.getId();
  const token     = ScriptApp.getOAuthToken();

  const feuilles = [
    { nom: "Avis à envoyer électrique",  nomFichier: "AVIS_OUVERTS_ELECTRIQUE - " },
    { nom: "Avis à envoyer régulation",  nomFichier: "AVIS_OUVERTS_REGULATION - " }
  ];

  const piecesJointes = [];
  feuilles.forEach(function(feuilleInfo) {
    const feuilleOT = feuille.getSheetByName(feuilleInfo.nom);
    if (!feuilleOT) return;
    const url = "https://docs.google.com/spreadsheets/d/" + feuilleID + "/export?" +
      "format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false" +
      "&pagenumbers=false&gridlines=false&fzr=false&gid=" + feuilleOT.getSheetId();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    piecesJointes.push(response.getBlob().setName("[TEST] " + feuilleInfo.nomFichier + "test.pdf"));
  });

  sendEmailEWS(
    OCP_EMAIL_EWS, null,
    '[TEST] Avis ouverts Électrique & Régulation',
    '<p>[TEST] Email de vérification — pièces jointes PDF ci-dessous.</p>',
    piecesJointes,
    "Bureau de méthode Daoui - Section Planification"
  );
  Logger.log('✅ Test envoyé à : ' + OCP_EMAIL_EWS);
}
