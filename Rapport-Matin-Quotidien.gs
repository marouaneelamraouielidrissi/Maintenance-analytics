/**
 * Rapport-Matin-Quotidien.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Envoie chaque jour à 8h un email avec le rapport en PIÈCE JOINTE PDF natif.
 * Le PDF est généré via DocumentApp (Google Docs) — pas de conversion HTML.
 *
 * Filtres :
 *   PDR confirmés  : Dispo = "OUI"
 *                    ET Col V ∉ {TCLO, CLOT, LANC}
 *                    ET Col K ne commence pas par "SOPL"
 *   OT réalisés    : Réalisation = "Fait"
 *                    ET Col V ∉ {TCLO, CLOT, LANC}
 *
 * Scopes requis dans appsscript.json :
 *   "https://www.googleapis.com/auth/drive"
 *   "https://www.googleapis.com/auth/documents"
 *   "https://www.googleapis.com/auth/spreadsheets"
 *   "https://www.googleapis.com/auth/script.external_request"
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Configuration ─────────────────────────────────────────────────────────────

const RM_SHEET_ID   = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
const RM_SHEET_NAME = 'Travaux hebdomadaire';
const DESTINATAIRE_RAPPORT = 'm.elamraoui@ocpgroup.ma';

// Colonnes (index 0-basé)
const COL_ORDRE       = 0;   // A  – Numéro OT
const COL_DESC        = 3;   // D  – Description
const COL_OBJET       = 5;   // F  – Objet technique
const COL_POSTE       = 8;   // I  – Poste de travail
const COL_STATUT_UTIL = 10;  // K  – Statut utilisateur
const COL_REALISATION = 14;  // O  – "Fait" | "NFait"
const COL_PDR         = 18;  // S  – Désignation PDR
const COL_DISPO       = 19;  // T  – "OUI" | "NON" | vide
const COL_OBS         = 20;  // U  – Observation
const COL_STATUT_SYS  = 21;  // V  – Statut système ABR

// Couleurs
const VERT_FONCE  = '#166534';
const VERT_CLAIR  = '#d1fae5';
const VERT_ZEBRE  = '#f0fdf4';
const BLEU_FONCE  = '#1e3a5f';
const BLEU_CLAIR  = '#dbeafe';
const BLEU_ZEBRE  = '#eff6ff';
const GRIS_TEXTE  = '#6b7280';
const GRIS_PIED   = '#9ca3af';

// ── Helpers de filtrage ───────────────────────────────────────────────────────

function statutSysExclu(row) {
  const s = String(row[COL_STATUT_SYS] || '').toUpperCase();
  return s.includes('TCLO') || s.includes('CLOT') || s.includes('LANC');
}

function statutUtilSOPL(row) {
  return String(row[COL_STATUT_UTIL] || '').trim().toUpperCase().startsWith('SOPL');
}

// ── Fonction principale ───────────────────────────────────────────────────────

function envoyerRapportMatin() {
  try {
    const ss    = SpreadsheetApp.openById(RM_SHEET_ID);
    const sheet = ss.getSheetByName(RM_SHEET_NAME);
    if (!sheet) { Logger.log('[Rapport Matin] Feuille introuvable : ' + RM_SHEET_NAME); return; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) { Logger.log('[Rapport Matin] Aucune donnée.'); return; }

    const rows = data.slice(1);

    // ── PDR confirmés : Dispo=OUI  ET  Col V ∉ {TCLO,CLOT,LANC}  ET  Col K ne commence pas par SOPL
    const pdrConfirmes = rows
      .filter(r => String(r[COL_PDR] || '').trim()
               && String(r[COL_DISPO] || '').trim().toUpperCase() === 'OUI'
               && !statutSysExclu(r)
               && !statutUtilSOPL(r))
      .map(r => ({
        ordre : String(r[COL_ORDRE]       || '').trim(),
        desc  : String(r[COL_DESC]        || '').trim(),
        objet : String(r[COL_OBJET]       || '').trim(),
        poste : String(r[COL_POSTE]       || '').trim(),
        pdr   : String(r[COL_PDR]         || '').trim(),
        obs   : String(r[COL_OBS]         || '').trim(),
      }));

    // ── OT réalisés : Réalisation=Fait  ET  Col V ∉ {TCLO,CLOT,LANC}
    const otRealises = rows
      .filter(r => String(r[COL_REALISATION] || '').trim() === 'Fait' && !statutSysExclu(r))
      .map(r => ({
        ordre : String(r[COL_ORDRE]  || '').trim(),
        desc  : String(r[COL_DESC]   || '').trim(),
        objet : String(r[COL_OBJET]  || '').trim(),
        poste : String(r[COL_POSTE]  || '').trim(),
        obs   : String(r[COL_OBS]    || '').trim(),
      }));

    const today    = new Date();
    const dateStr  = Utilities.formatDate(today, Session.getScriptTimeZone(), "EEEE dd MMMM yyyy");
    const nomFich  = 'Rapport-Matin-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
    const subject  = 'Rapport Matin — ' + cap(dateStr);

    // ── Génération PDF natif (DocumentApp)
    const pdfBlob = genererPdfDocApp(dateStr, pdrConfirmes, otRealises, nomFich);

    // ── Sauvegarde dans Drive + lien de téléchargement
    const lienPdf = sauvegarderPdfDrive(pdfBlob);

    // ── Corps du mail avec lien Drive (envoyé via sendEmailOCP existant)
    const corps = '<p>Bonjour,</p>'
      + '<p>Le rapport matin du <strong>' + cap(dateStr) + '</strong> est disponible :</p>'
      + '<ul>'
      + '<li><strong>' + pdrConfirmes.length + '</strong> PDR confirmé(s)</li>'
      + '<li><strong>' + otRealises.length   + '</strong> OT réalisé(s)</li>'
      + '</ul>'
      + '<p style="margin:20px 0;">'
      + '<a href="' + lienPdf + '" style="background:#1e3a5f;color:#fff;padding:10px 20px;'
      + 'border-radius:6px;text-decoration:none;font-weight:bold;">Télécharger le rapport PDF</a>'
      + '</p>'
      + '<p style="color:#9ca3af;font-size:11px;">Maintenance Analytics · OCP Daoui</p>';

    sendEmailOCP(DESTINATAIRE_RAPPORT, subject, '', { htmlBody: corps });

    Logger.log('[Rapport Matin] Envoyé | PDR=' + pdrConfirmes.length + ' | OT=' + otRealises.length);

  } catch (err) {
    Logger.log('[Rapport Matin] ERREUR : ' + err.toString() + '\n' + (err.stack || ''));
  }
}

// ── Génération PDF natif via DocumentApp ──────────────────────────────────────

function genererPdfDocApp(dateStr, pdrConfirmes, otRealises, nomFichier) {

  const doc  = DocumentApp.create('__tmp_rapport_matin__');
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  // ── Titre ──────────────────────────────────────────────────────────────────
  const titre = body.appendParagraph('RAPPORT MATIN — MAINTENANCE DAOUI');
  titre.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titre.editAsText().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(BLEU_FONCE);

  const datePara = body.appendParagraph(cap(dateStr));
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  datePara.editAsText().setFontFamily('Arial').setFontSize(11).setItalic(true).setForegroundColor(GRIS_TEXTE);

  body.appendParagraph('');

  // ── Résumé (tableau 2 colonnes) ────────────────────────────────────────────
  const resumeData = [
    [String(pdrConfirmes.length), String(otRealises.length)],
    ['PDR Confirmés',             'OT Réalisés'            ],
  ];
  const resumeTable = body.appendTable(resumeData);
  resumeTable.setBorderWidth(0);

  [[VERT_CLAIR, BLEU_CLAIR], [VERT_CLAIR, BLEU_CLAIR]].forEach((cols, ri) => {
    const tRow = resumeTable.getRow(ri);
    cols.forEach((bg, ci) => {
      const cell = tRow.getCell(ci);
      cell.setBackgroundColor(bg);
      const para = cell.getChild(0).asParagraph();
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const txt = para.editAsText().setFontFamily('Arial');
      if (ri === 0) {
        txt.setFontSize(28).setBold(true).setForegroundColor(ci === 0 ? VERT_FONCE : BLEU_FONCE);
      } else {
        txt.setFontSize(10).setBold(false).setForegroundColor(ci === 0 ? VERT_FONCE : BLEU_FONCE);
      }
    });
  });

  body.appendParagraph('');

  // ── Section PDR confirmés ──────────────────────────────────────────────────
  const pdrTitre = body.appendParagraph('PDR CONFIRMÉS');
  pdrTitre.editAsText().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(VERT_FONCE);

  if (pdrConfirmes.length === 0) {
    const vide = body.appendParagraph('Aucun PDR confirmé pour le moment.');
    vide.editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor(GRIS_TEXTE);
  } else {
    const entetesPDR = ['Ordre OT', 'Description', 'Objet technique', 'Poste', 'PDR', 'Observation'];
    const pdrTable   = body.appendTable();
    pdrTable.setBorderWidth(1);

    // En-tête
    const hRow = pdrTable.appendTableRow();
    entetesPDR.forEach(label => {
      const c = hRow.appendTableCell(label);
      c.setBackgroundColor(VERT_FONCE);
      c.editAsText().setFontFamily('Arial').setFontSize(8).setBold(true).setForegroundColor('#ffffff');
    });

    // Données
    pdrConfirmes.forEach((r, i) => {
      const bg   = i % 2 === 0 ? VERT_ZEBRE : '#ffffff';
      const dRow = pdrTable.appendTableRow();
      [r.ordre, r.desc, r.objet, r.poste, r.pdr, r.obs || '—'].forEach((val, j) => {
        const c = dRow.appendTableCell(val);
        c.setBackgroundColor(bg);
        const t = c.editAsText().setFontFamily('Arial').setFontSize(8).setForegroundColor('#111827');
        if (j === 0) t.setBold(true).setForegroundColor(VERT_FONCE);
      });
    });
  }

  body.appendParagraph('');

  // ── Section OT réalisés ────────────────────────────────────────────────────
  const otTitre = body.appendParagraph('OT RÉALISÉS');
  otTitre.editAsText().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(BLEU_FONCE);

  if (otRealises.length === 0) {
    const vide = body.appendParagraph('Aucun OT réalisé pour le moment.');
    vide.editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor(GRIS_TEXTE);
  } else {
    const entetesOT = ['Ordre OT', 'Description', 'Objet technique', 'Poste', 'Observation'];
    const otTable   = body.appendTable();
    otTable.setBorderWidth(1);

    // En-tête
    const hRow = otTable.appendTableRow();
    entetesOT.forEach(label => {
      const c = hRow.appendTableCell(label);
      c.setBackgroundColor(BLEU_FONCE);
      c.editAsText().setFontFamily('Arial').setFontSize(8).setBold(true).setForegroundColor('#ffffff');
    });

    // Données
    otRealises.forEach((r, i) => {
      const bg   = i % 2 === 0 ? BLEU_ZEBRE : '#ffffff';
      const dRow = otTable.appendTableRow();
      [r.ordre, r.desc, r.objet, r.poste, r.obs || '—'].forEach((val, j) => {
        const c = dRow.appendTableCell(val);
        c.setBackgroundColor(bg);
        const t = c.editAsText().setFontFamily('Arial').setFontSize(8).setForegroundColor('#111827');
        if (j === 0) t.setBold(true).setForegroundColor(BLEU_FONCE);
      });
    });
  }

  // ── Pied de page ──────────────────────────────────────────────────────────
  body.appendParagraph('');
  const pied = body.appendParagraph('Rapport généré automatiquement chaque jour à 8h00 — Maintenance Analytics · OCP Daoui');
  pied.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pied.editAsText().setFontFamily('Arial').setFontSize(8).setItalic(true).setForegroundColor(GRIS_PIED);

  doc.saveAndClose();

  // ── Export en PDF via Drive API ────────────────────────────────────────────
  const fileId = doc.getId();
  const token  = ScriptApp.getOAuthToken();

  const pdfResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId + '/export?mimeType=application/pdf',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );

  // Suppression du document temporaire
  UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId,
    { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );

  if (pdfResp.getResponseCode() !== 200) {
    throw new Error('Export PDF erreur ' + pdfResp.getResponseCode() + ' : ' + pdfResp.getContentText().substring(0, 300));
  }

  const pdfBlob = pdfResp.getBlob();
  pdfBlob.setName(nomFichier);
  return pdfBlob;
}

// ── Sauvegarde PDF dans Google Drive ─────────────────────────────────────────

/**
 * Sauvegarde le PDF dans un dossier Drive "Rapports Matin" et retourne
 * l'URL de prévisualisation (accessible à quiconque possède le lien).
 */
function sauvegarderPdfDrive(pdfBlob) {
  const NOM_DOSSIER = 'Rapports Matin — Maintenance Daoui';

  // Cherche ou crée le dossier
  const folders = DriveApp.getFoldersByName(NOM_DOSSIER);
  const dossier = folders.hasNext() ? folders.next() : DriveApp.createFolder(NOM_DOSSIER);

  // Sauvegarde le PDF
  const fichier = dossier.createFile(pdfBlob);
  fichier.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return fichier.getUrl();
}

// ── [OBSOLÈTE — conservé pour référence] ─────────────────────────────────────
// L'envoi EWS avec pièce jointe est bloqué par le pare-feu OCP.
// Le PDF est désormais sauvegardé dans Drive et le lien est inclus dans le mail.

function envoyerEmailAvecPDF(to, subject, htmlBody, pdfBlob) {
  const props     = PropertiesService.getScriptProperties();
  const OCP_EMAIL = props.getProperty('OCP_EMAIL')   || 'm.elamraoui@ocpgroup.ma';
  const OCP_PASS  = props.getProperty('OCP_PASSWORD') || '';
  const EWS_URL   = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';

  if (!OCP_PASS) throw new Error('OCP_PASSWORD non défini dans les propriétés du script.');

  const auth = { Authorization: 'Basic ' + Utilities.base64Encode(OCP_EMAIL + ':' + OCP_PASS) };

  // ── Étape 1 : créer le brouillon (SaveOnly) ────────────────────────────────
  const xmlTo      = xmlEsc(to);
  const xmlSubject = xmlEsc(subject);
  const xmlBody    = xmlEsc(htmlBody);

  const soap1 = ewsEnvelope(
    '<m:CreateItem MessageDisposition="SaveOnly">'
    + '<m:Items><t:Message>'
    + '<t:Subject>'  + xmlSubject + '</t:Subject>'
    + '<t:Body BodyType="HTML">' + xmlBody + '</t:Body>'
    + '<t:ToRecipients>'
    + '<t:Mailbox><t:EmailAddress>' + xmlTo + '</t:EmailAddress></t:Mailbox>'
    + '</t:ToRecipients>'
    + '</t:Message></m:Items>'
    + '</m:CreateItem>'
  );

  const resp1 = ewsFetch(EWS_URL, auth, 'CreateItem', soap1);
  const idMatch  = resp1.match(/ItemId Id="([^"]+)"/);
  const ckMatch  = resp1.match(/ItemId Id="[^"]+" ChangeKey="([^"]+)"/);
  if (!idMatch) throw new Error('EWS CreateItem : ItemId introuvable.\n' + resp1.substring(0, 400));

  const itemId    = idMatch[1];
  const changeKey = ckMatch ? ckMatch[1] : '';

  // ── Étape 2 : attacher le PDF ──────────────────────────────────────────────
  const pdfB64 = Utilities.base64Encode(pdfBlob.getBytes());
  const pdfName = xmlEsc(pdfBlob.getName());

  const soap2 = ewsEnvelope(
    '<m:CreateAttachment>'
    + '<m:ParentItemId Id="' + itemId + '" ChangeKey="' + changeKey + '"/>'
    + '<m:Attachments>'
    + '<t:FileAttachment>'
    + '<t:Name>' + pdfName + '</t:Name>'
    + '<t:ContentType>application/pdf</t:ContentType>'
    + '<t:IsInline>false</t:IsInline>'
    + '<t:Content>' + pdfB64 + '</t:Content>'
    + '</t:FileAttachment>'
    + '</m:Attachments>'
    + '</m:CreateAttachment>'
  );

  const resp2 = ewsFetch(EWS_URL, auth, 'CreateAttachment', soap2);
  const rootIdMatch  = resp2.match(/RootItemId Id="([^"]+)"/);
  const rootCkMatch  = resp2.match(/RootItemId Id="[^"]+" ChangeKey="([^"]+)"/);
  if (!rootIdMatch) throw new Error('EWS CreateAttachment : RootItemId introuvable.\n' + resp2.substring(0, 400));

  const rootId  = rootIdMatch[1];
  const rootCk  = rootCkMatch ? rootCkMatch[1] : '';

  // ── Étape 3 : envoyer le brouillon ────────────────────────────────────────
  const soap3 = ewsEnvelope(
    '<m:SendItem SaveItemToFolder="true">'
    + '<m:ItemIds>'
    + '<t:ItemId Id="' + rootId + '" ChangeKey="' + rootCk + '"/>'
    + '</m:ItemIds>'
    + '</m:SendItem>'
  );

  ewsFetch(EWS_URL, auth, 'SendItem', soap3);
}

// ── Helpers EWS ───────────────────────────────────────────────────────────────

function ewsEnvelope(body) {
  return '<?xml version="1.0" encoding="utf-8"?>'
    + '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'
    + ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
    + ' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">'
    + '<soap:Body>' + body + '</soap:Body>'
    + '</soap:Envelope>';
}

function ewsFetch(url, authHeaders, action, soap) {
  const ns   = 'http://schemas.microsoft.com/exchange/services/2006/messages/';
  const resp = UrlFetchApp.fetch(url, {
    method            : 'POST',
    contentType       : 'text/xml; charset=utf-8',
    headers           : Object.assign({ SOAPAction: ns + action }, authHeaders),
    payload           : soap,
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code !== 200) throw new Error('EWS ' + action + ' HTTP ' + code + ' : ' + text.substring(0, 400));
  if (text.indexOf('ResponseClass="Error"') !== -1) {
    const m = text.match(/<m:MessageText>(.*?)<\/m:MessageText>/);
    throw new Error('EWS ' + action + ' erreur : ' + (m ? m[1] : text.substring(0, 300)));
  }
  return text;
}

function xmlEsc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function chunk76(b64) { return (b64.match(/.{1,76}/g) || []).join('\r\n'); }
function cap(str)     { return str ? str.charAt(0).toUpperCase() + str.slice(1) : str; }

// ── Autorisation (à exécuter une seule fois) ──────────────────────────────────

/**
 * Déclenche la demande d'autorisation pour Drive et Documents.
 * Exécutez cette fonction UNE SEULE FOIS et acceptez toutes les permissions.
 */
function autoriserAcces() {
  const doc = DocumentApp.create('__test_autorisation__');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  Logger.log('Autorisation Drive + Documents accordée.');
}

// ── Envoi instantané ──────────────────────────────────────────────────────────

/**
 * Envoie immédiatement le rapport (test ou envoi manuel).
 * → Sélectionnez cette fonction dans le menu et cliquez "Exécuter".
 */
function envoyerRapportMaintenant() {
  Logger.log('[Rapport Matin] Envoi instantané déclenché manuellement.');
  envoyerRapportMatin();
}

// ── Trigger quotidien à 8h ────────────────────────────────────────────────────

/**
 * Crée le trigger quotidien à 8h.
 * ⚠️ À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script.
 */
function configurerDeclencheurMatin() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'envoyerRapportMatin') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('envoyerRapportMatin')
    .timeBased().everyDays(1).atHour(8).inTimezone('Africa/Casablanca').create();
  Logger.log('Trigger quotidien configure : 8h (Africa/Casablanca)');
}
