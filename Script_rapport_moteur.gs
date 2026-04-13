// ── Configuration OCP Exchange (EWS) ─────────────────────────
// OCP_EMAIL, EWS_URL, getOcpPassword() et sendEmailOCP() sont définis
// dans google_apps_script.js (script principal de l'application).
// Ce script étant indépendant, on redéfinit ici getOcpPassword()
// qui lit le mot de passe depuis les Propriétés du script.
// → Dans l'éditeur GAS de CE projet : Paramètres du projet → Propriétés du script
//   Ajouter la propriété : OCP_PASSWORD = <ton mot de passe OCP>

const OCP_EMAIL_RPT = 'm.elamraoui@ocpgroup.ma';
const EWS_URL_RPT   = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';

function getOcpPasswordRpt() {
  return PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD') || '';
}

/**
 * Envoie un email via EWS avec pièce jointe.
 * Utilise le format MIME encodé en base64 (MimeContent) pour gérer les pièces jointes.
 */
function sendEmailEWS(to, cc, subject, htmlBody, attachment, senderName) {
  const boundary = 'boundary_mainana_' + Date.now();

  const htmlB64  = Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8);
  const subjB64  = Utilities.base64Encode(subject,  Utilities.Charset.UTF_8);

  var mimeParts = [
    'From: "' + senderName + '" <' + OCP_EMAIL_RPT + '>',
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

  if (attachment) {
    const attachB64  = Utilities.base64Encode(attachment.getBytes());
    const attachName = attachment.getName();
    mimeParts = mimeParts.concat([
      '--' + boundary,
      'Content-Type: ' + attachment.getContentType() + '; name="' + attachName + '"',
      'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="' + attachName + '"',
      '',
      attachB64,
      ''
    ]);
  }

  mimeParts.push('--' + boundary + '--');

  const mime     = mimeParts.filter(function(l){ return l !== null; }).join('\r\n');
  const mimeB64  = Utilities.base64Encode(mime, Utilities.Charset.UTF_8);

  const soap = '<?xml version="1.0" encoding="utf-8"?>'
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

  const credentials = Utilities.base64Encode(OCP_EMAIL_RPT + ':' + getOcpPasswordRpt());
  const response = UrlFetchApp.fetch(EWS_URL_RPT, {
    method: 'post',
    contentType: 'text/xml; charset=utf-8',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'SOAPAction': 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
    },
    payload: soap,
    muteHttpExceptions: true
  });

  const respCode = response.getResponseCode();
  const respText = response.getContentText();
  if (respCode !== 200 || respText.indexOf('NoError') === -1) {
    throw new Error('EWS send failed (' + respCode + '): ' + respText.substring(0, 500));
  }
}

// ═══════════════════════════════════════════════════════════════
// FONCTION PRINCIPALE
// ═══════════════════════════════════════════════════════════════
function Envoyer_Rapport_Moteur_PDF() {
  const feuille    = SpreadsheetApp.openById("1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE");
  const feuilleNom = "Rapport moteur";
  const sujet      = "Rapport Mensuel de gestion des moteurs électriques";

  const dateActuelle = new Date();
  const mois  = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "MMMM");
  const annee = Utilities.formatDate(dateActuelle, Session.getScriptTimeZone(), "yyyy");

  const messageHtml = `
    <div style="color: #003366; font-family: Arial, sans-serif;">
      <p>Bonjour,</p>
      <p>Veuillez trouver ci-joint le <strong>rapport de gestion des moteurs électriques</strong> pour le mois de <strong>${mois} ${annee}</strong>.</p>
      <p>Cordialement.</p>
      <br>
      <div style="font-family: 'Times New Roman', serif; font-size: 14px; color: #002060; line-height: 1.5;">
        <span style="font-weight: bold;">Marouane ELAMRAOUI</span><br>
        <span style="color: #c55a11;">Méthode de Maintenance</span><br>
        <span style="font-weight: bold;">OCP SA - <span style="color:#002060;">Khouribga</span></span><br>
        <span style="color: green;">Tél.  :</span> 0661323784<br>
        <span style="color: green;">Cisco :</span>  8103388<br>
        <span style="font-style: italic; color: green;">E-mail: </span>
        <a href="mailto:m.elamraoui@ocpgroup.ma" style="font-style: italic; color: #002060; text-decoration: underline;">m.elamraoui@ocpgroup.ma</a>
      </div>
    </div>
  `;

  const options = {
    exportFormat: "pdf",
    format: "pdf",
    size: "A4",
    landscape: false,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    fzr: false,
    gid: feuille.getSheetByName(feuilleNom).getSheetId()
  };

  const url = 'https://docs.google.com/spreadsheets/d/' + feuille.getId() + '/export?' +
    Object.entries(options).map(function(kv){ return kv[0] + '=' + kv[1]; }).join('&');

  const token    = ScriptApp.getOAuthToken();
  const pdfResp  = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob     = pdfResp.getBlob().setName("Rapport mensuel de gestion des moteurs électriques.pdf");

  const to = "chinig@ocpgroup.ma, abdeljaouad.achlih@ocpgroup.ma, a.jebli2@ocpgroup.ma, aasli@ocpgroup.ma, hamid.kaabouchi@ocpgroup.ma, laghchioua@ocpgroup.ma, amhid@ocpgroup.ma, naoui@ocpgroup.ma, lhoussaine.kadiri@ocpgroup.ma, e.touhamy@ocpgroup.ma";
  const cc = "ad.benbaouali@ocpgroup.ma, m.mamouni@ocpgroup.ma, elkhyari@ocpgroup.ma, ibtissame.elkhloufi@ocpgroup.ma, wajid@ocpgroup.ma, a.dahmou@ocpgroup.ma, m.mabdoui@ocpgroup.ma, lazrag@ocpgroup.ma, m.elamraoui@ocpgroup.ma";

  sendEmailEWS(to, cc, sujet, messageHtml, blob, 'Bureau de méthode Daoui - Gestion des interchangeables');
}

function doGet(e) {
  try {
    Envoyer_Rapport_Moteur_PDF();
    return ContentService
      .createTextOutput("OK - Rapport envoyé avec succès")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput("ERREUR : " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
