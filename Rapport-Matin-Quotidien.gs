/**
 * Rapport-Matin-Quotidien.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Envoie chaque jour à 8h un email récapitulatif contenant :
 *   • La liste des PDR confirmés (Disponibilité = "OUI")
 *   • La liste des OT réalisés (Réalisation = "Fait")
 *   … qu'ils apparaissent dans la liste de mise à profit ou dans le plan de charge
 *
 * INSTALLATION :
 *   1. Copiez ce fichier dans votre projet Google Apps Script
 *   2. Renseignez l'adresse email du destinataire dans DESTINATAIRE_RAPPORT
 *   3. Exécutez une seule fois la fonction  configurerDeclencheurMatin()
 *      → Elle crée le trigger quotidien à 8h automatiquement
 *
 * NOTE : La fonction sendEmailOCP() est définie dans google_apps_script.js
 *        et est accessible ici car tous les fichiers .gs partagent le même projet.
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Configuration ─────────────────────────────────────────────────────────────

const RM_SHEET_ID   = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
const RM_SHEET_NAME = 'Travaux hebdomadaire';

// Email du destinataire du rapport (modifiez si besoin)
const DESTINATAIRE_RAPPORT = 'm.elamraoui@ocpgroup.ma';

// Colonnes (index 0-basé, conformément au schéma de la feuille)
const COL_ORDRE       = 0;   // A  – Numéro OT
const COL_DESC        = 3;   // D  – Description
const COL_OBJET       = 5;   // F  – Objet technique
const COL_POSTE       = 8;   // I  – Poste de travail
const COL_STATUT_UTIL = 10;  // K  – Statut utilisateur (CRPR …)
const COL_REALISATION = 14;  // O  – Réalisation : "Fait" | "NFait"
const COL_PDR         = 18;  // S  – Désignation PDR
const COL_DISPO       = 19;  // T  – Disponibilité : "OUI" | "NON" | vide
const COL_OBS         = 20;  // U  – Observation
const COL_STATUT_SYS  = 21;  // V  – Statut système SAP (contient "créé" si actif)

// ── Fonction principale ───────────────────────────────────────────────────────

/**
 * Point d'entrée du rapport matinal.
 * Appelé automatiquement par le trigger quotidien à 8h.
 */
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

    const rows = data.slice(1); // Ignore la ligne d'en-tête

    // ── Filtrage PDR confirmés ─────────────────────────────────────────────
    // PDR confirmé = PDR renseigné ET Disponibilité = "OUI"
    const pdrConfirmes = rows
      .filter(row => {
        const pdr    = String(row[COL_PDR]   || '').trim();
        const dispo  = String(row[COL_DISPO] || '').trim().toUpperCase();
        return pdr && dispo === 'OUI';
      })
      .map(row => ({
        ordre : String(row[COL_ORDRE]  || '').trim(),
        desc  : String(row[COL_DESC]   || '').trim(),
        objet : String(row[COL_OBJET]  || '').trim(),
        poste : String(row[COL_POSTE]  || '').trim(),
        pdr   : String(row[COL_PDR]    || '').trim(),
        obs   : String(row[COL_OBS]    || '').trim(),
      }));

    // ── Filtrage OT réalisés ───────────────────────────────────────────────
    // OT réalisé = Réalisation = "Fait" (liste de mise à profit + plan de charge)
    const otRealises = rows
      .filter(row => {
        const real = String(row[COL_REALISATION] || '').trim();
        return real === 'Fait';
      })
      .map(row => ({
        ordre : String(row[COL_ORDRE]  || '').trim(),
        desc  : String(row[COL_DESC]   || '').trim(),
        objet : String(row[COL_OBJET]  || '').trim(),
        poste : String(row[COL_POSTE]  || '').trim(),
        obs   : String(row[COL_OBS]    || '').trim(),
      }));

    // ── Construction du mail ───────────────────────────────────────────────
    const today   = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "EEEE dd MMMM yyyy");
    const subject = '📋 Rapport Matin — ' + capitaliserPremiereMot(dateStr);
    const body    = construireHtmlRapport(dateStr, pdrConfirmes, otRealises);

    // ── Envoi ──────────────────────────────────────────────────────────────
    sendEmailOCP(DESTINATAIRE_RAPPORT, subject, body);
    Logger.log('[Rapport Matin] Email envoyé à ' + DESTINATAIRE_RAPPORT +
               ' | PDR confirmés : ' + pdrConfirmes.length +
               ' | OT réalisés : ' + otRealises.length);

  } catch (err) {
    Logger.log('[Rapport Matin] ERREUR : ' + err.toString());
  }
}

// ── Construction du HTML ──────────────────────────────────────────────────────

function construireHtmlRapport(dateStr, pdrConfirmes, otRealises) {
  const totalPDR = pdrConfirmes.length;
  const totalOT  = otRealises.length;

  // ── Tableau PDR confirmés
  let tablePDR = '';
  if (totalPDR === 0) {
    tablePDR = '<p style="color:#6b7280;font-style:italic;margin:8px 0;">Aucun PDR confirmé pour le moment.</p>';
  } else {
    tablePDR = `
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:#166534;color:#fff;">
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">Ordre OT</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">Description</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">Objet technique</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">Poste</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">PDR</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #d1fae5;">Observation</th>
        </tr>
      </thead>
      <tbody>
        ${pdrConfirmes.map((r, i) => `
        <tr style="background:${i % 2 === 0 ? '#f0fdf4' : '#ffffff'};">
          <td style="padding:7px 10px;border:1px solid #d1fae5;font-weight:600;color:#166534;">${esc(r.ordre)}</td>
          <td style="padding:7px 10px;border:1px solid #d1fae5;">${esc(r.desc)}</td>
          <td style="padding:7px 10px;border:1px solid #d1fae5;">${esc(r.objet)}</td>
          <td style="padding:7px 10px;border:1px solid #d1fae5;">${badgePoste(r.poste)}</td>
          <td style="padding:7px 10px;border:1px solid #d1fae5;font-weight:600;">${esc(r.pdr)}</td>
          <td style="padding:7px 10px;border:1px solid #d1fae5;color:#6b7280;">${esc(r.obs) || '—'}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
  }

  // ── Tableau OT réalisés
  let tableOT = '';
  if (totalOT === 0) {
    tableOT = '<p style="color:#6b7280;font-style:italic;margin:8px 0;">Aucun OT réalisé pour le moment.</p>';
  } else {
    tableOT = `
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead>
        <tr style="background:#1e3a5f;color:#fff;">
          <th style="padding:8px 10px;text-align:left;border:1px solid #bfdbfe;">Ordre OT</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #bfdbfe;">Description</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #bfdbfe;">Objet technique</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #bfdbfe;">Poste</th>
          <th style="padding:8px 10px;text-align:left;border:1px solid #bfdbfe;">Observation</th>
        </tr>
      </thead>
      <tbody>
        ${otRealises.map((r, i) => `
        <tr style="background:${i % 2 === 0 ? '#eff6ff' : '#ffffff'};">
          <td style="padding:7px 10px;border:1px solid #bfdbfe;font-weight:600;color:#1e3a5f;">${esc(r.ordre)}</td>
          <td style="padding:7px 10px;border:1px solid #bfdbfe;">${esc(r.desc)}</td>
          <td style="padding:7px 10px;border:1px solid #bfdbfe;">${esc(r.objet)}</td>
          <td style="padding:7px 10px;border:1px solid #bfdbfe;">${badgePoste(r.poste)}</td>
          <td style="padding:7px 10px;border:1px solid #bfdbfe;color:#6b7280;">${esc(r.obs) || '—'}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
  }

  return `
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:820px;margin:24px auto;background:#fff;border-radius:10px;
            box-shadow:0 2px 12px rgba(0,0,0,.10);overflow:hidden;">

  <!-- En-tête -->
  <div style="background:linear-gradient(135deg,#1e3a5f 0%,#166534 100%);
              padding:28px 32px;color:#fff;">
    <div style="font-size:22px;font-weight:700;letter-spacing:.5px;">
      📋 Rapport Matin — Maintenance Daoui
    </div>
    <div style="font-size:14px;margin-top:6px;opacity:.85;">
      ${capitaliserPremiereMot(dateStr)}
    </div>
  </div>

  <!-- Résumé -->
  <div style="display:flex;gap:0;border-bottom:1px solid #e5e7eb;">
    <div style="flex:1;padding:20px 28px;border-right:1px solid #e5e7eb;text-align:center;">
      <div style="font-size:36px;font-weight:800;color:#166534;">${totalPDR}</div>
      <div style="font-size:13px;color:#6b7280;margin-top:4px;">PDR confirmés</div>
    </div>
    <div style="flex:1;padding:20px 28px;text-align:center;">
      <div style="font-size:36px;font-weight:800;color:#1e3a5f;">${totalOT}</div>
      <div style="font-size:13px;color:#6b7280;margin-top:4px;">OT réalisés</div>
    </div>
  </div>

  <!-- Section PDR confirmés -->
  <div style="padding:24px 28px;">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
      <div style="width:4px;height:22px;background:#166534;border-radius:2px;"></div>
      <h2 style="margin:0;font-size:16px;font-weight:700;color:#166534;">
        PDR Confirmés
        <span style="font-size:12px;font-weight:400;color:#6b7280;margin-left:8px;">
          — Disponibilité = OUI
        </span>
      </h2>
    </div>
    ${tablePDR}
  </div>

  <!-- Séparateur -->
  <div style="height:1px;background:#e5e7eb;margin:0 28px;"></div>

  <!-- Section OT réalisés -->
  <div style="padding:24px 28px;">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
      <div style="width:4px;height:22px;background:#1e3a5f;border-radius:2px;"></div>
      <h2 style="margin:0;font-size:16px;font-weight:700;color:#1e3a5f;">
        OT Réalisés
        <span style="font-size:12px;font-weight:400;color:#6b7280;margin-left:8px;">
          — Liste de mise à profit &amp; Plan de charge
        </span>
      </h2>
    </div>
    ${tableOT}
  </div>

  <!-- Pied de page -->
  <div style="background:#f9fafb;border-top:1px solid #e5e7eb;
              padding:16px 28px;font-size:11px;color:#9ca3af;text-align:center;">
    Ce rapport est généré automatiquement chaque jour à 8h00 — Maintenance Analytics · OCP Daoui
  </div>
</div>
</body>
</html>`;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Échappe les caractères HTML pour éviter les injections dans le tableau */
function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Badge coloré par poste de travail */
function badgePoste(poste) {
  const colors = {
    '421-MEC'  : { bg: '#fef3c7', text: '#92400e', border: '#fcd34d' },
    '421-CHAU' : { bg: '#fef3c7', text: '#92400e', border: '#fcd34d' },
    '421-INST' : { bg: '#e0f2fe', text: '#075985', border: '#7dd3fc' },
    '423-ELEC' : { bg: '#f3e8ff', text: '#6b21a8', border: '#c084fc' },
    '423-REG'  : { bg: '#fce7f3', text: '#9d174d', border: '#f9a8d4' },
  };
  const c = colors[poste] || { bg: '#f3f4f6', text: '#374151', border: '#d1d5db' };
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;
                        font-size:11px;font-weight:600;
                        background:${c.bg};color:${c.text};border:1px solid ${c.border};">
            ${esc(poste)}
          </span>`;
}

/** Met la première lettre en majuscule */
function capitaliserPremiereMot(str) {
  if (!str) return str;
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ── Trigger (à exécuter une seule fois manuellement) ──────────────────────────

/**
 * Crée un trigger quotidien à 8h pour envoyerRapportMatin().
 * ⚠️ Exécutez cette fonction UNE SEULE FOIS depuis l'éditeur Apps Script.
 *    Elle supprime les anciens triggers du même nom avant d'en créer un nouveau.
 */
function configurerDeclencheurMatin() {
  // Supprime les triggers existants pour éviter les doublons
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'envoyerRapportMatin') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crée le trigger quotidien entre 8h et 9h (GAS arrondit à l'heure)
  ScriptApp.newTrigger('envoyerRapportMatin')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone('Africa/Casablanca')
    .create();

  Logger.log('✅ Trigger quotidien configuré : envoyerRapportMatin() à 8h (Africa/Casablanca)');
}
