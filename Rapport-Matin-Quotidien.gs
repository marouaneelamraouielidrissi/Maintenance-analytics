/**
 * Rapport-Matin-Quotidien.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Envoie chaque jour à 8h un email avec le rapport en PIÈCE JOINTE PDF.
 * Approche : les données sont écrites dans un onglet Google Sheets dédié,
 * puis exportées en PDF via l'URL d'export natif de Sheets.
 *
 * Filtres :
 *   PDR confirmés : Dispo="OUI"  ET  Col V ∉ {TCLO,CLOT,LANC}  ET  Col K ne commence pas par "SOPL"
 *   OT réalisés   : Réalisation="Fait"  ET  Col V ∉ {TCLO,CLOT,LANC}
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Configuration ─────────────────────────────────────────────────────────────

const OCP_EMAIL_RM  = 'm.elamraoui@ocpgroup.ma';
const EWS_URL_RM    = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';
const DEST_RAPPORT  = 'm.elamraoui@ocpgroup.ma';

const RM_SHEET_ID   = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
const RM_SHEET_NAME = 'Travaux hebdomadaire';
const RPT_TAB_NAME  = 'Rapport Matin';          // Onglet dédié au rapport PDF

// Colonnes source (index 0-basé)
const COL_ORDRE       = 0;   // A – Numéro OT
const COL_DESC        = 3;   // D – Description
const COL_OBJET       = 5;   // F – Objet technique
const COL_POSTE       = 8;   // I – Poste de travail
const COL_STATUT_UTIL = 10;  // K – Statut utilisateur
const COL_REALISATION = 14;  // O – "Fait" | "NFait"
const COL_PDR         = 18;  // S – Désignation PDR
const COL_DISPO       = 19;  // T – "OUI" | "NON" | vide
const COL_OBS         = 20;  // U – Observation
const COL_STATUT_SYS  = 21;  // V – Statut système ABR

// ── Helpers de filtrage ───────────────────────────────────────────────────────

function rm_statutSysExclu(row) {
  const s = String(row[COL_STATUT_SYS] || '').toUpperCase();
  return s.includes('TCLO') || s.includes('CLOT') || s.includes('LANC');
}

function rm_statutUtilSOPL(row) {
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

    // ── PDR confirmés
    // Un OT peut avoir plusieurs lignes PDR : seule la 1ère ligne a Ordre/Desc/Objet/Poste.
    // On reporte ces valeurs sur les lignes suivantes du même OT quand elles sont vides.
    const pdrConfirmes = [];
    let dernierOrdre = '', dernierDesc = '', dernierObjet = '', dernierPoste = '';
    rows.forEach(function(r) {
      const pdr   = String(r[COL_PDR]   || '').trim();
      const dispo = String(r[COL_DISPO] || '').trim().toUpperCase();
      if (!pdr || dispo !== 'OUI' || rm_statutSysExclu(r) || rm_statutUtilSOPL(r)) return;

      const ordre = String(r[COL_ORDRE] || '').trim();
      const desc  = String(r[COL_DESC]  || '').trim();
      const objet = String(r[COL_OBJET] || '').trim();
      const poste = String(r[COL_POSTE] || '').trim();

      // Si la ligne a un Ordre renseigné, on met à jour les valeurs de référence
      if (ordre) { dernierOrdre = ordre; dernierDesc = desc; dernierObjet = objet; dernierPoste = poste; }

      pdrConfirmes.push([
        dernierOrdre,
        dernierDesc,
        dernierObjet,
        dernierPoste,
        pdr,
        String(r[COL_OBS] || '').trim() || '—',
      ]);
    });

    // ── OT réalisés
    // On dédoublonne par Ordre OT : une seule ligne par OT même si plusieurs lignes existent.
    const vusOT = {};
    const otRealises = [];
    rows.forEach(function(r) {
      const real  = String(r[COL_REALISATION] || '').trim();
      const ordre = String(r[COL_ORDRE]       || '').trim();
      if (real !== 'Fait' || rm_statutSysExclu(r)) return;
      if (!ordre || vusOT[ordre]) return;
      vusOT[ordre] = true;
      otRealises.push([
        ordre,
        String(r[COL_DESC]  || '').trim(),
        String(r[COL_OBJET] || '').trim(),
        String(r[COL_POSTE] || '').trim(),
        String(r[COL_OBS]   || '').trim() || '—',
      ]);
    });

    const today   = new Date();
    const tz      = Session.getScriptTimeZone();
    const dateStr = Utilities.formatDate(today, tz, "EEEE dd MMMM yyyy");
    const nomFich = 'Rapport-Matin-' + Utilities.formatDate(today, tz, 'yyyy-MM-dd') + '.pdf';
    const subject = 'Rapport Matin — ' + rm_cap(dateStr);

    // ── Remplir l'onglet rapport
    const rptSheet = remplirOngletRapport(ss, dateStr, pdrConfirmes, otRealises);

    // ── Exporter en PDF via URL Sheets native
    const pdfBlob = exporterOngletEnPdf(ss, rptSheet, nomFich);

    // ── Corps de l'email (bref résumé HTML)
    const corps = '<div style="font-family:Arial,sans-serif;color:#1e3a5f;">'
      + '<p>Bonjour,</p>'
      + '<p>Veuillez trouver ci-joint le <strong>Rapport Matin</strong> du <strong>' + rm_cap(dateStr) + '</strong>.</p>'
      + '<table style="border-collapse:collapse;margin:12px 0;">'
      + '<tr>'
      + '<td style="padding:10px 20px;background:#d1fae5;text-align:center;border-radius:6px 0 0 6px;">'
      + '<div style="font-size:26px;font-weight:800;color:#166534;">' + pdrConfirmes.length + '</div>'
      + '<div style="font-size:11px;color:#166534;">PDR confirmés</div></td>'
      + '<td style="padding:10px 20px;background:#dbeafe;text-align:center;border-radius:0 6px 6px 0;">'
      + '<div style="font-size:26px;font-weight:800;color:#1e3a5f;">' + otRealises.length + '</div>'
      + '<div style="font-size:11px;color:#1e3a5f;">OT réalisés</div></td>'
      + '</tr></table>'
      + '<p style="color:#9ca3af;font-size:11px;">Maintenance Analytics · OCP Daoui</p>'
      + '</div>';

    // ── Envoi via EWS avec PDF en pièce jointe
    sendEmailEWS_RM(DEST_RAPPORT, null, subject, corps, pdfBlob, 'Maintenance Analytics — OCP Daoui');

    Logger.log('[Rapport Matin] Envoyé | PDR=' + pdrConfirmes.length + ' | OT=' + otRealises.length);

  } catch (err) {
    Logger.log('[Rapport Matin] ERREUR : ' + err.toString() + '\n' + (err.stack || ''));
  }
}

// ── Remplissage de l'onglet rapport ──────────────────────────────────────────

function remplirOngletRapport(ss, dateStr, pdrConfirmes, otRealises) {
  // Crée ou réinitialise l'onglet
  let rpt = ss.getSheetByName(RPT_TAB_NAME);
  if (rpt) { rpt.clear(); } else { rpt = ss.insertSheet(RPT_TAB_NAME); }

  const VERT  = '#166534';
  const BLEU  = '#1e3a5f';
  const VCLR  = '#d1fae5';
  const BCLR  = '#dbeafe';
  const VZEBR = '#f0fdf4';
  const BZEBR = '#eff6ff';
  const BLANC = '#ffffff';
  const GRIS  = '#f9fafb';

  let ligne = 1;

  // ── Titre ──────────────────────────────────────────────────────────────────
  rpt.getRange(ligne, 1, 1, 6).merge()
    .setValue('RAPPORT MATIN — MAINTENANCE DAOUI')
    .setBackground(BLEU).setFontColor(BLANC)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  rpt.setRowHeight(ligne, 36);
  ligne++;

  rpt.getRange(ligne, 1, 1, 6).merge()
    .setValue(rm_cap(dateStr))
    .setBackground(GRIS).setFontColor('#6b7280')
    .setFontSize(11).setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  rpt.setRowHeight(ligne, 24);
  ligne++;

  // ── Résumé ─────────────────────────────────────────────────────────────────
  ligne++; // ligne vide

  rpt.getRange(ligne, 1, 1, 3).merge()
    .setValue('PDR Confirmés : ' + pdrConfirmes.length)
    .setBackground(VCLR).setFontColor(VERT)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  rpt.getRange(ligne, 4, 1, 3).merge()
    .setValue('OT Réalisés : ' + otRealises.length)
    .setBackground(BCLR).setFontColor(BLEU)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  rpt.setRowHeight(ligne, 30);
  ligne++;

  // ── Section PDR confirmés ──────────────────────────────────────────────────
  ligne++; // ligne vide

  rpt.getRange(ligne, 1, 1, 6).merge()
    .setValue('PDR CONFIRMÉS')
    .setBackground(VERT).setFontColor(BLANC)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  rpt.setRowHeight(ligne, 26);
  ligne++;

  const enPDR = ['Ordre OT', 'Description', 'Objet technique', 'Poste', 'PDR', 'Observation'];
  const rngEnPDR = rpt.getRange(ligne, 1, 1, 6);
  rngEnPDR.setValues([enPDR])
    .setBackground('#bbf7d0').setFontColor(VERT)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  rpt.setRowHeight(ligne, 20);
  ligne++;

  if (pdrConfirmes.length === 0) {
    rpt.getRange(ligne, 1, 1, 6).merge()
      .setValue('Aucun PDR confirmé pour le moment.')
      .setFontStyle('italic').setFontColor('#9ca3af').setFontSize(9)
      .setHorizontalAlignment('center');
    ligne++;
  } else {
    pdrConfirmes.forEach(function(row, i) {
      const bg = i % 2 === 0 ? VZEBR : BLANC;
      rpt.getRange(ligne, 1, 1, 6).setValues([row]).setBackground(bg).setFontSize(9);
      rpt.getRange(ligne, 1).setFontWeight('bold').setFontColor(VERT);
      rpt.setRowHeight(ligne, 18);
      ligne++;
    });
  }

  // ── Section OT réalisés ────────────────────────────────────────────────────
  ligne++; // ligne vide

  rpt.getRange(ligne, 1, 1, 5).merge()
    .setValue('OT RÉALISÉS — Liste de mise à profit & Plan de charge')
    .setBackground(BLEU).setFontColor(BLANC)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  rpt.setRowHeight(ligne, 26);
  ligne++;

  const enOT = ['Ordre OT', 'Description', 'Objet technique', 'Poste', 'Observation'];
  rpt.getRange(ligne, 1, 1, 5).setValues([enOT])
    .setBackground('#bfdbfe').setFontColor(BLEU)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  rpt.setRowHeight(ligne, 20);
  ligne++;

  if (otRealises.length === 0) {
    rpt.getRange(ligne, 1, 1, 5).merge()
      .setValue('Aucun OT réalisé pour le moment.')
      .setFontStyle('italic').setFontColor('#9ca3af').setFontSize(9)
      .setHorizontalAlignment('center');
    ligne++;
  } else {
    otRealises.forEach(function(row, i) {
      const bg = i % 2 === 0 ? BZEBR : BLANC;
      rpt.getRange(ligne, 1, 1, 5).setValues([row]).setBackground(bg).setFontSize(9);
      rpt.getRange(ligne, 1).setFontWeight('bold').setFontColor(BLEU);
      rpt.setRowHeight(ligne, 18);
      ligne++;
    });
  }

  // ── Pied de page ──────────────────────────────────────────────────────────
  ligne++;
  rpt.getRange(ligne, 1, 1, 6).merge()
    .setValue('Rapport généré automatiquement chaque jour à 8h00 — Maintenance Analytics · OCP Daoui')
    .setFontSize(8).setFontStyle('italic').setFontColor('#9ca3af')
    .setHorizontalAlignment('center').setBackground(GRIS);
  rpt.setRowHeight(ligne, 18);

  // ── Largeurs de colonnes ──────────────────────────────────────────────────
  rpt.setColumnWidth(1, 80);   // Ordre OT
  rpt.setColumnWidth(2, 180);  // Description
  rpt.setColumnWidth(3, 150);  // Objet technique
  rpt.setColumnWidth(4, 75);   // Poste
  rpt.setColumnWidth(5, 160);  // PDR / Observation
  rpt.setColumnWidth(6, 120);  // Observation (PDR uniquement)

  // ── Bordures sur les tableaux ─────────────────────────────────────────────
  const nbLignesPDR = Math.max(pdrConfirmes.length, 1) + 1;
  rpt.getRange(/* header PDR */ 7, 1, nbLignesPDR, 6)
    .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();
  return rpt;
}

// ── Export PDF via URL native Sheets ─────────────────────────────────────────

function exporterOngletEnPdf(ss, sheet, nomFichier) {
  const options = {
    exportFormat : 'pdf',
    format       : 'pdf',
    size         : 'A4',
    landscape    : false,
    sheetnames   : false,
    printtitle   : false,
    pagenumbers  : false,
    gridlines    : false,
    fzr          : false,
    gid          : sheet.getSheetId(),
  };

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + Object.entries(options).map(function(kv) { return kv[0] + '=' + kv[1]; }).join('&');

  const token   = ScriptApp.getOAuthToken();
  const pdfResp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob    = pdfResp.getBlob();
  blob.setName(nomFichier);
  return blob;
}

// ── Envoi EWS avec pièce jointe (même pattern que le script rapport moteur) ───

function sendEmailEWS_RM(to, cc, subject, htmlBody, attachment, senderName) {
  const password = PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD') || '';
  if (!password) throw new Error('OCP_PASSWORD non défini dans les propriétés du script.');

  const boundary = 'boundary_rm_' + Date.now();
  const htmlB64  = Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8);
  const subjB64  = Utilities.base64Encode(subject,  Utilities.Charset.UTF_8);

  var parts = [
    'From: "' + senderName + '" <' + OCP_EMAIL_RM + '>',
    'To: '    + to,
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
    '',
  ];

  if (attachment) {
    const attB64 = Utilities.base64Encode(attachment.getBytes());
    const attName = attachment.getName();
    parts = parts.concat([
      '--' + boundary,
      'Content-Type: ' + attachment.getContentType() + '; name="' + attName + '"',
      'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="' + attName + '"',
      '',
      attB64,
      '',
    ]);
  }

  parts.push('--' + boundary + '--');

  const mime    = parts.filter(function(l) { return l !== null; }).join('\r\n');
  const mimeB64 = Utilities.base64Encode(mime, Utilities.Charset.UTF_8);

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

  const creds = Utilities.base64Encode(OCP_EMAIL_RM + ':' + password);
  const resp  = UrlFetchApp.fetch(EWS_URL_RM, {
    method            : 'post',
    contentType       : 'text/xml; charset=utf-8',
    headers           : {
      'Authorization' : 'Basic ' + creds,
      'SOAPAction'    : 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem',
    },
    payload           : soap,
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code !== 200 || text.indexOf('NoError') === -1) {
    throw new Error('EWS send failed (' + code + '): ' + text.substring(0, 500));
  }
}

// ── Helper ────────────────────────────────────────────────────────────────────

function rm_cap(str) { return str ? str.charAt(0).toUpperCase() + str.slice(1) : str; }

// ── Envoi instantané ──────────────────────────────────────────────────────────

/**
 * Envoie immédiatement le rapport.
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
  Logger.log('Trigger quotidien configuré : 8h (Africa/Casablanca)');
}
