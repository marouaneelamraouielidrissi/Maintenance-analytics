/**
 * Rapport-Matin-Quotidien.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Envoie chaque jour à 8h un email avec le rapport en PIÈCE JOINTE PDF :
 *   • PDR confirmés  : Dispo = "OUI"  AND  Col V ∉ {TCLO, CLOT}  AND  Col K ne commence pas par "SOPL"
 *   • OT réalisés    : Réalisation = "Fait"  AND  Col V ∉ {TCLO, CLOT}
 *
 * INSTALLATION :
 *   1. Copiez ce fichier dans le même projet Google Apps Script que google_apps_script.js
 *   2. Vérifiez que RM_SHEET_ID correspond bien à votre feuille "Travaux hebdomadaire"
 *   3. Modifiez DESTINATAIRE_RAPPORT si besoin
 *   4. Exécutez UNE SEULE FOIS configurerDeclencheurMatin() pour créer le trigger quotidien
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
const COL_STATUT_UTIL = 10;  // K  – Statut utilisateur (ex : CRPR, SOPL…)
const COL_REALISATION = 14;  // O  – Réalisation : "Fait" | "NFait"
const COL_PDR         = 18;  // S  – Désignation PDR
const COL_DISPO       = 19;  // T  – Disponibilité : "OUI" | "NON" | vide
const COL_OBS         = 20;  // U  – Observation
const COL_STATUT_SYS  = 21;  // V  – Statut système ABR (TCLO, CLOT, créé…)

// ── Helpers de filtrage ───────────────────────────────────────────────────────

/** Retourne true si le statut système ABR exclut la ligne (TCLO ou CLOT) */
function estStatutSysExclu(row) {
  const s = String(row[COL_STATUT_SYS] || '').toUpperCase();
  return s.includes('TCLO') || s.includes('CLOT');
}

/** Retourne true si le statut utilisateur commence par SOPL */
function estStatutUtilSOPL(row) {
  return String(row[COL_STATUT_UTIL] || '').trim().toUpperCase().startsWith('SOPL');
}

// ── Fonction principale ───────────────────────────────────────────────────────

function envoyerRapportMatin() {
  try {
    const ss    = SpreadsheetApp.openById(RM_SHEET_ID);
    const sheet = ss.getSheetByName(RM_SHEET_NAME);

    if (!sheet) {
      Logger.log('[Rapport Matin] Feuille introuvable : ' + RM_SHEET_NAME);
      return;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('[Rapport Matin] Aucune donnée dans la feuille.');
      return;
    }

    const rows = data.slice(1);

    // ── PDR confirmés ──────────────────────────────────────────────────────
    // Règle : Dispo = "OUI"  ET  Col V ∉ {TCLO, CLOT}  ET  Col K ne commence pas par "SOPL"
    const pdrConfirmes = rows
      .filter(row => {
        const pdr   = String(row[COL_PDR]   || '').trim();
        const dispo = String(row[COL_DISPO] || '').trim().toUpperCase();
        return pdr && dispo === 'OUI' && !estStatutSysExclu(row) && !estStatutUtilSOPL(row);
      })
      .map(row => ({
        ordre      : String(row[COL_ORDRE]       || '').trim(),
        desc       : String(row[COL_DESC]        || '').trim(),
        objet      : String(row[COL_OBJET]       || '').trim(),
        poste      : String(row[COL_POSTE]       || '').trim(),
        statutUtil : String(row[COL_STATUT_UTIL] || '').trim(),
        pdr        : String(row[COL_PDR]         || '').trim(),
        obs        : String(row[COL_OBS]         || '').trim(),
      }));

    // ── OT réalisés ────────────────────────────────────────────────────────
    // Règle : Réalisation = "Fait"  ET  Col V ∉ {TCLO, CLOT}
    const otRealises = rows
      .filter(row => {
        const real = String(row[COL_REALISATION] || '').trim();
        return real === 'Fait' && !estStatutSysExclu(row);
      })
      .map(row => ({
        ordre      : String(row[COL_ORDRE]       || '').trim(),
        desc       : String(row[COL_DESC]        || '').trim(),
        objet      : String(row[COL_OBJET]       || '').trim(),
        poste      : String(row[COL_POSTE]       || '').trim(),
        statutUtil : String(row[COL_STATUT_UTIL] || '').trim(),
        obs        : String(row[COL_OBS]         || '').trim(),
      }));

    const today   = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "EEEE dd MMMM yyyy");
    const subject = 'Rapport Matin — ' + capitaliserPremiere(dateStr);

    // ── Génération du PDF ──────────────────────────────────────────────────
    const htmlContent = construireHtmlRapport(dateStr, pdrConfirmes, otRealises);
    const nomFichier  = 'Rapport-Matin-' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
    const pdfBlob     = genererPdfDepuisHtml(htmlContent, nomFichier);

    // ── Envoi avec pièce jointe ────────────────────────────────────────────
    const corpsMail = '<p>Bonjour,</p>'
      + '<p>Veuillez trouver ci-joint le rapport matin du <strong>' + capitaliserPremiere(dateStr) + '</strong>.</p>'
      + '<ul>'
      + '<li><strong>' + pdrConfirmes.length + '</strong> PDR confirmé(s)</li>'
      + '<li><strong>' + otRealises.length + '</strong> OT réalisé(s)</li>'
      + '</ul>'
      + '<p style="color:#6b7280;font-size:12px;">Maintenance Analytics · OCP Daoui</p>';

    envoyerEmailAvecPDF(DESTINATAIRE_RAPPORT, subject, corpsMail, pdfBlob);

    Logger.log('[Rapport Matin] Email envoyé | PDR confirmés : ' + pdrConfirmes.length
               + ' | OT réalisés : ' + otRealises.length);

  } catch (err) {
    Logger.log('[Rapport Matin] ERREUR : ' + err.toString() + '\n' + err.stack);
  }
}

// ── Génération PDF via DriveApp ───────────────────────────────────────────────

/**
 * Convertit un contenu HTML en blob PDF en passant par Google Drive.
 * Le fichier temporaire est supprimé immédiatement après conversion.
 */
function genererPdfDepuisHtml(htmlContent, nomFichier) {
  const htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, 'tmp_rapport.html');
  const tempFile = DriveApp.createFile(htmlBlob);
  const pdfBlob  = tempFile.getAs(MimeType.PDF);
  pdfBlob.setName(nomFichier);
  tempFile.setTrashed(true);
  return pdfBlob;
}

// ── Envoi EWS avec pièce jointe PDF (MIME multipart) ─────────────────────────

/**
 * Envoie un email via OCP Exchange (EWS) avec un PDF en pièce jointe.
 * Utilise la méthode MimeContent de l'API EWS pour un support natif des attachements.
 */
function envoyerEmailAvecPDF(to, subject, htmlBody, pdfBlob) {
  const props       = PropertiesService.getScriptProperties();
  const OCP_EMAIL   = props.getProperty('OCP_EMAIL')   || 'm.elamraoui@ocpgroup.ma';
  const OCP_PASS    = props.getProperty('OCP_PASSWORD') || '';
  const EWS_URL     = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';

  if (!OCP_PASS) throw new Error('OCP_PASSWORD non défini dans les propriétés du script.');

  const boundary  = 'boundary_' + Utilities.getUuid().replace(/-/g, '');
  const subjectB64 = Utilities.base64Encode(subject, Utilities.Charset.UTF_8);
  const htmlB64   = chunkBase64(Utilities.base64Encode(htmlBody,        Utilities.Charset.UTF_8));
  const pdfB64    = chunkBase64(Utilities.base64Encode(pdfBlob.getBytes()));
  const pdfName   = pdfBlob.getName();

  // Construction du message MIME multipart/mixed
  const mimeLines = [
    'From: ' + OCP_EMAIL,
    'To: ' + to,
    'Subject: =?UTF-8?B?' + subjectB64 + '?=',
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    htmlB64,
    '',
    '--' + boundary,
    'Content-Type: application/pdf',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="' + pdfName + '"',
    '',
    pdfB64,
    '',
    '--' + boundary + '--',
  ];

  const mimeRaw   = mimeLines.join('\r\n');
  const mimeB64   = Utilities.base64Encode(mimeRaw, Utilities.Charset.UTF_8);

  const soap = '<?xml version="1.0" encoding="utf-8"?>'
    + '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'
    + ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
    + ' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">'
    + '<soap:Body>'
    + '<m:CreateItem MessageDisposition="SendAndSaveCopy">'
    + '<m:Items>'
    + '<t:Message>'
    + '<t:MimeContent CharacterSet="UTF-8">' + mimeB64 + '</t:MimeContent>'
    + '</t:Message>'
    + '</m:Items>'
    + '</m:CreateItem>'
    + '</soap:Body>'
    + '</soap:Envelope>';

  const resp = UrlFetchApp.fetch(EWS_URL, {
    method          : 'POST',
    contentType     : 'text/xml; charset=utf-8',
    headers         : {
      'Authorization' : 'Basic ' + Utilities.base64Encode(OCP_EMAIL + ':' + OCP_PASS),
      'SOAPAction'    : 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem',
    },
    payload         : soap,
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('EWS HTTP ' + code + ' : ' + resp.getContentText().substring(0, 600));
  }

  const xml = resp.getContentText();
  if (xml.indexOf('ResponseClass="Error"') !== -1) {
    const msg = xml.match(/<m:MessageText>(.*?)<\/m:MessageText>/);
    throw new Error('EWS erreur : ' + (msg ? msg[1] : xml.substring(0, 400)));
  }
}

/** Découpe une chaîne base64 en lignes de 76 caractères (standard MIME) */
function chunkBase64(b64) {
  return (b64.match(/.{1,76}/g) || []).join('\r\n');
}

// ── Construction du HTML (utilisé pour le PDF) ────────────────────────────────

function construireHtmlRapport(dateStr, pdrConfirmes, otRealises) {
  const totalPDR = pdrConfirmes.length;
  const totalOT  = otRealises.length;

  // ── Tableau PDR confirmés
  let tablePDR = '';
  if (totalPDR === 0) {
    tablePDR = '<p style="color:#6b7280;font-style:italic;margin:8px 0 0;">Aucun PDR confirmé.</p>';
  } else {
    const lignesPDR = pdrConfirmes.map((r, i) =>
      '<tr style="background:' + (i % 2 === 0 ? '#f0fdf4' : '#ffffff') + ';">'
      + '<td style="' + tdStyle + 'font-weight:600;color:#166534;">' + esc(r.ordre)      + '</td>'
      + '<td style="' + tdStyle + '">'                                + esc(r.desc)       + '</td>'
      + '<td style="' + tdStyle + '">'                                + esc(r.objet)      + '</td>'
      + '<td style="' + tdStyle + '">'                                + badgePoste(r.poste) + '</td>'
      + '<td style="' + tdStyle + 'font-weight:600;">'               + esc(r.pdr)        + '</td>'
      + '<td style="' + tdStyle + 'color:#6b7280;">'                 + (esc(r.obs) || '—') + '</td>'
      + '</tr>'
    ).join('');

    tablePDR = '<table style="width:100%;border-collapse:collapse;font-size:12px;">'
      + '<thead><tr style="background:#166534;color:#fff;">'
      + thStyle('Ordre OT') + thStyle('Description') + thStyle('Objet technique')
      + thStyle('Poste') + thStyle('PDR') + thStyle('Observation')
      + '</tr></thead><tbody>' + lignesPDR + '</tbody></table>';
  }

  // ── Tableau OT réalisés
  let tableOT = '';
  if (totalOT === 0) {
    tableOT = '<p style="color:#6b7280;font-style:italic;margin:8px 0 0;">Aucun OT réalisé.</p>';
  } else {
    const lignesOT = otRealises.map((r, i) =>
      '<tr style="background:' + (i % 2 === 0 ? '#eff6ff' : '#ffffff') + ';">'
      + '<td style="' + tdStyle + 'font-weight:600;color:#1e3a5f;">' + esc(r.ordre)        + '</td>'
      + '<td style="' + tdStyle + '">'                                + esc(r.desc)         + '</td>'
      + '<td style="' + tdStyle + '">'                                + esc(r.objet)        + '</td>'
      + '<td style="' + tdStyle + '">'                                + badgePoste(r.poste)  + '</td>'
      + '<td style="' + tdStyle + 'color:#6b7280;">'                 + (esc(r.obs) || '—')  + '</td>'
      + '</tr>'
    ).join('');

    tableOT = '<table style="width:100%;border-collapse:collapse;font-size:12px;">'
      + '<thead><tr style="background:#1e3a5f;color:#fff;">'
      + thStyleBlue('Ordre OT') + thStyleBlue('Description') + thStyleBlue('Objet technique')
      + thStyleBlue('Poste') + thStyleBlue('Observation')
      + '</tr></thead><tbody>' + lignesOT + '</tbody></table>';
  }

  return '<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">'
    + '<style>body{margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;}'
    + 'h2{margin:0 0 14px;font-size:15px;} .section{padding:22px 28px;}'
    + '</style></head><body>'
    + '<div style="max-width:820px;margin:20px auto;background:#fff;border-radius:8px;'
    + 'box-shadow:0 2px 10px rgba(0,0,0,.10);overflow:hidden;">'

    // En-tête
    + '<div style="background:linear-gradient(135deg,#1e3a5f 0%,#166534 100%);padding:26px 28px;color:#fff;">'
    + '<div style="font-size:20px;font-weight:700;">Rapport Matin — Maintenance Daoui</div>'
    + '<div style="font-size:13px;margin-top:5px;opacity:.85;">' + capitaliserPremiere(dateStr) + '</div>'
    + '</div>'

    // Compteurs
    + '<div style="display:flex;border-bottom:1px solid #e5e7eb;">'
    + '<div style="flex:1;padding:18px 28px;border-right:1px solid #e5e7eb;text-align:center;">'
    + '<div style="font-size:34px;font-weight:800;color:#166534;">' + totalPDR + '</div>'
    + '<div style="font-size:12px;color:#6b7280;margin-top:3px;">PDR confirmés</div></div>'
    + '<div style="flex:1;padding:18px 28px;text-align:center;">'
    + '<div style="font-size:34px;font-weight:800;color:#1e3a5f;">' + totalOT + '</div>'
    + '<div style="font-size:12px;color:#6b7280;margin-top:3px;">OT réalisés</div></div>'
    + '</div>'

    // Section PDR
    + '<div class="section">'
    + '<h2 style="color:#166534;border-left:4px solid #166534;padding-left:10px;">'
    + 'PDR Confirmés <span style="font-size:11px;font-weight:400;color:#6b7280;">— Dispo = OUI · Col V ∉ TCLO/CLOT · Col K ne commence pas par SOPL</span></h2>'
    + tablePDR + '</div>'

    + '<div style="height:1px;background:#e5e7eb;margin:0 28px;"></div>'

    // Section OT
    + '<div class="section">'
    + '<h2 style="color:#1e3a5f;border-left:4px solid #1e3a5f;padding-left:10px;">'
    + 'OT Réalisés <span style="font-size:11px;font-weight:400;color:#6b7280;">— Liste de mise à profit &amp; Plan de charge · Col V ∉ TCLO/CLOT</span></h2>'
    + tableOT + '</div>'

    // Pied de page
    + '<div style="background:#f9fafb;border-top:1px solid #e5e7eb;padding:14px 28px;'
    + 'font-size:10px;color:#9ca3af;text-align:center;">'
    + 'Rapport généré automatiquement chaque jour à 8h00 — Maintenance Analytics · OCP Daoui</div>'
    + '</div></body></html>';
}

// ── Micro-helpers HTML ────────────────────────────────────────────────────────

const tdStyle     = 'padding:6px 9px;border:1px solid #d1fae5;';
const tdStyleBlue = 'padding:6px 9px;border:1px solid #bfdbfe;';

function thStyle(label) {
  return '<th style="padding:8px 9px;text-align:left;border:1px solid #d1fae5;">' + label + '</th>';
}
function thStyleBlue(label) {
  return '<th style="padding:8px 9px;text-align:left;border:1px solid #bfdbfe;">' + label + '</th>';
}

function esc(str) {
  return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function badgePoste(poste) {
  const map = {
    '421-MEC'  : 'background:#fef3c7;color:#92400e;border:1px solid #fcd34d;',
    '421-CHAU' : 'background:#fef3c7;color:#92400e;border:1px solid #fcd34d;',
    '421-INST' : 'background:#e0f2fe;color:#075985;border:1px solid #7dd3fc;',
    '423-ELEC' : 'background:#f3e8ff;color:#6b21a8;border:1px solid #c084fc;',
    '423-REG'  : 'background:#fce7f3;color:#9d174d;border:1px solid #f9a8d4;',
  };
  const style = map[poste] || 'background:#f3f4f6;color:#374151;border:1px solid #d1d5db;';
  return '<span style="display:inline-block;padding:2px 7px;border-radius:999px;font-size:10px;font-weight:600;' + style + '">' + esc(poste) + '</span>';
}

function capitaliserPremiere(str) {
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : str;
}

// ── Envoi instantané ──────────────────────────────────────────────────────────

/**
 * Envoie immédiatement le rapport (test ou envoi manuel).
 * → Sélectionnez cette fonction dans le menu déroulant et cliquez "Exécuter".
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
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'envoyerRapportMatin') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('envoyerRapportMatin')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone('Africa/Casablanca')
    .create();

  Logger.log('Trigger quotidien configuré : envoyerRapportMatin() a 8h (Africa/Casablanca)');
}
