// ═══════════════════════════════════════════════════════════════
//  RAPPORT HEBDOMADAIRE DE PLANIFICATION
//  Envoi chaque début de semaine avec :
//    - Calendrier des arrêts de la semaine précédente
//    - KPI mensuel (calculés depuis les données SAP)
// ═══════════════════════════════════════════════════════════════

// ── IDs des fichiers Google Sheets ───────────────────────────
const RH_OT_FILE_ID     = '1aQAvb1DUv6Vk1Y1C-WEYgQnYN1BxujEAg8lbMt1sP3s'; // Données SAP (OTs)
const RH_ARRETS_FILE_ID = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE'; // Planning arrêts
const RH_ARRETS_SHEET   = 'Travaux hebdomadaire';

// ── EWS (réutilise OCP_EMAIL + getOcpPassword() de google_apps_script.js) ──

// ── Utilitaires date ─────────────────────────────────────────
function rhGetWeekNumber(d) {
  var dc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  dc.setUTCDate(dc.getUTCDate() + 4 - (dc.getUTCDay() || 7));
  var yearStart = new Date(Date.UTC(dc.getUTCFullYear(), 0, 1));
  return Math.ceil((((dc - yearStart) / 86400000) + 1) / 7);
}

function rhFormatDate(dateStr) {
  // "2025-04-07" → "07/04/2025"
  if (!dateStr) return '';
  var p = dateStr.split('-');
  if (p.length !== 3) return dateStr;
  return p[2] + '/' + p[1] + '/' + p[0];
}

function rhGetPrevWeekRange() {
  var today = new Date();
  var dow = today.getDay() || 7; // 1=Lun … 7=Dim
  // Début semaine courante = Lundi
  var mondayCurrent = new Date(today);
  mondayCurrent.setDate(today.getDate() - (dow - 1));
  // Début semaine précédente
  var mondayPrev = new Date(mondayCurrent);
  mondayPrev.setDate(mondayCurrent.getDate() - 7);
  var sundayPrev = new Date(mondayPrev);
  sundayPrev.setDate(mondayPrev.getDate() + 6);
  return { start: mondayPrev, end: sundayPrev };
}

function rhDateToStr(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

// ── Lecture des arrêts de la semaine précédente ───────────────
function rhGetArretsSemainePrecedente() {
  var range   = rhGetPrevWeekRange();
  var startStr = rhDateToStr(range.start);
  var endStr   = rhDateToStr(range.end);

  var ss    = SpreadsheetApp.openById(RH_ARRETS_FILE_ID);
  var sheet = ss.getSheetByName(RH_ARRETS_SHEET);
  if (!sheet) return { rows: [], startStr: startStr, endStr: endStr, semaine: rhGetWeekNumber(range.start) };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim().toLowerCase(); });

  // Détecter les colonnes clés
  function col(names) {
    for (var i = 0; i < names.length; i++) {
      var idx = headers.indexOf(names[i].toLowerCase());
      if (idx >= 0) return idx;
    }
    return -1;
  }
  var cDate    = col(['start date', 'date début', 'date debut', 'date']);
  var cInstall = col(['installation', 'équipement', 'equipement']);
  var cSection = col(['section']);
  var cSemaine = col(['semaine']);
  var cAnnee   = col(['année', 'annee']);
  var cReal    = col(['réalisation', 'realisation']);

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    // Parser la date de la ligne
    var rawDate = cDate >= 0 ? r[cDate] : null;
    var dateStr = '';
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      dateStr = rhDateToStr(rawDate);
    } else if (typeof rawDate === 'number') {
      var d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
      dateStr = rhDateToStr(d);
    } else if (rawDate) {
      dateStr = rawDate.toString().trim();
      // Si format DD/MM/YYYY
      var parts = dateStr.split('/');
      if (parts.length === 3 && parts[2].length === 4) {
        dateStr = parts[2] + '-' + parts[1] + '-' + parts[0];
      }
    }
    if (!dateStr) continue;
    if (dateStr < startStr || dateStr > endStr) continue;

    var realRaw = (cReal >= 0 ? r[cReal] : '').toString().trim().toLowerCase();
    var statut  = 'Non réalisé';
    var couleur = '#ffc7ce'; // rouge clair
    if (realRaw === 'oui')                                            { statut = 'Réalisé';  couleur = '#c6efce'; }
    else if (realRaw.includes('imprévu') || realRaw.includes('imprevu')) { statut = 'Imprévu';  couleur = '#ffeb9c'; }

    rows.push({
      date:         dateStr,
      installation: cInstall >= 0 ? r[cInstall].toString().trim() : '',
      section:      cSection >= 0 ? r[cSection].toString().trim() : '',
      semaine:      cSemaine >= 0 ? r[cSemaine].toString().trim() : '',
      statut:       statut,
      couleur:      couleur
    });
  }

  // Trier par date
  rows.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return { rows: rows, startStr: startStr, endStr: endStr, semaine: rhGetWeekNumber(range.start) };
}

// ── Lecture et calcul des KPI depuis les OTs SAP ─────────────
function rhGetKpiMensuel() {
  var now   = new Date();
  var year  = now.getFullYear();
  var month = now.getMonth(); // 0-indexé

  var ss    = SpreadsheetApp.openById(RH_OT_FILE_ID);
  // Prendre le premier onglet ou chercher un onglet connu
  var sheet = ss.getSheets()[0];

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  function colIdx(names) {
    for (var i = 0; i < names.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (headers[j].toString().trim().toLowerCase() === names[i].toLowerCase()) return j;
      }
    }
    return -1;
  }
  var cDebut     = colIdx(['début au plus tôt', 'debut au plus tot', 'date début', 'date debut']);
  var cStatutSys = colIdx(['statut système', 'statut systeme', 'statut sys']);
  var cStatutUt  = colIdx(['statut utilis.', 'statut util', 'statut utilis']);
  var cType      = colIdx(["type d'ordre", 'type ordre', 'type']);

  var MONTHS_FR = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];

  var total = 0, realise = 0, lanc = 0, crpr = 0, sys = 0, cur = 0, sysReal = 0, curReal = 0;

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    // Parser la date
    var rawD = cDebut >= 0 ? r[cDebut] : null;
    var d;
    if (rawD instanceof Date) d = rawD;
    else if (typeof rawD === 'number') d = new Date(Math.round((rawD - 25569) * 86400 * 1000));
    else if (rawD) d = new Date(rawD);
    else continue;
    if (isNaN(d) || d.getFullYear() !== year || d.getMonth() !== month) continue;

    total++;
    var ss_ = cStatutSys >= 0 ? r[cStatutSys].toString() : '';
    var su  = cStatutUt  >= 0 ? r[cStatutUt ].toString() : '';
    var tp  = cType      >= 0 ? r[cType     ].toString() : '';

    var isReal = ss_.includes('CONF') || ss_.includes('TCLO') || ss_.includes('CLOT');
    if (isReal) realise++;
    if (ss_.includes('LANC') && !ss_.includes('CONF') && !ss_.includes('TCLO')) lanc++;
    if (su.includes('CRPR')) crpr++;

    var isSys = ['ZCON','ZEST','ZETL'].indexOf(tp) >= 0;
    var isCur = tp === 'ZCOR';
    if (isSys) { sys++; if (isReal) sysReal++; }
    if (isCur) { cur++; if (isReal) curReal++; }
  }

  function pct(n, t) { return t ? ((n / t) * 100).toFixed(1) + '%' : '—'; }

  return {
    mois:       MONTHS_FR[month] + ' ' + year,
    total:      total,
    realise:    realise,
    tauxReal:   pct(realise, total),
    lanc:       lanc,
    crpr:       crpr,
    sys:        sys,
    cur:        cur,
    tauxPrev:   pct(sysReal, sys),
    tauxCor:    pct(curReal, cur),
    ratioPrevCor: pct(sys, total) + ' / ' + pct(cur, total)
  };
}

// ── Construction du tableau HTML des arrêts ──────────────────
function rhBuildTableauArrets(arrets) {
  if (!arrets.rows || arrets.rows.length === 0) {
    return '<p style="color:#6b7280;font-style:italic;">Aucun arrêt enregistré pour la semaine S' + arrets.semaine + '.</p>';
  }

  var html = '<table border="1" style="border-collapse:collapse;font-family:Arial;font-size:12px;width:100%;">';
  html += '<tr>'
    + '<th style="background:#002060;color:white;padding:7px 10px;text-align:left;">Date</th>'
    + '<th style="background:#002060;color:white;padding:7px 10px;text-align:left;">Installation</th>'
    + '<th style="background:#002060;color:white;padding:7px 10px;text-align:left;">Section</th>'
    + '<th style="background:#002060;color:white;padding:7px 10px;text-align:center;">Statut</th>'
    + '</tr>';

  for (var i = 0; i < arrets.rows.length; i++) {
    var r   = arrets.rows[i];
    var bg  = i % 2 === 0 ? '#f9fafb' : '#ffffff';
    html += '<tr style="background:' + bg + ';">'
      + '<td style="padding:6px 10px;">' + rhFormatDate(r.date) + '</td>'
      + '<td style="padding:6px 10px;">' + r.installation + '</td>'
      + '<td style="padding:6px 10px;">' + r.section + '</td>'
      + '<td style="padding:6px 10px;text-align:center;background:' + r.couleur + ';font-weight:600;">' + r.statut + '</td>'
      + '</tr>';
  }
  html += '</table>';
  return html;
}

// ── Construction du tableau HTML des KPIs ────────────────────
function rhBuildTableauKPI(kpi) {
  function tagColor(v) {
    var n = parseFloat(v);
    if (isNaN(n)) return '#e5e7eb';
    if (n >= 80)  return '#c6efce';
    if (n >= 50)  return '#ffeb9c';
    return '#ffc7ce';
  }
  function tagText(v) {
    var n = parseFloat(v);
    if (isNaN(n)) return '#374151';
    if (n >= 80)  return '#006100';
    if (n >= 50)  return '#9c6500';
    return '#9c0006';
  }

  var rows = [
    ['Total OTs planifiés', kpi.total, null],
    ['OTs réalisés (CONF/TCLO/CLOT)', kpi.realise, kpi.tauxReal],
    ['OTs lancés (LANC)', kpi.lanc, null],
    ['OTs en retard (CRPR)', kpi.crpr, null],
    ['OTs préventifs (ZCON/ZEST/ZETL)', kpi.sys, null],
    ['OTs correctifs (ZCOR)', kpi.cur, null],
    ['Taux réalisation préventif', null, kpi.tauxPrev],
    ['Taux réalisation correctif', null, kpi.tauxCor],
    ['Ratio Préventif / Correctif', null, kpi.ratioPrevCor]
  ];

  var html = '<table border="1" style="border-collapse:collapse;font-family:Arial;font-size:12px;width:100%;">';
  html += '<tr>'
    + '<th style="background:#002060;color:white;padding:7px 14px;text-align:left;">Indicateur</th>'
    + '<th style="background:#002060;color:white;padding:7px 14px;text-align:center;">Valeur</th>'
    + '<th style="background:#002060;color:white;padding:7px 14px;text-align:center;">Taux</th>'
    + '</tr>';

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var bg  = i % 2 === 0 ? '#f9fafb' : '#ffffff';
    var tauxVal = row[2] || '—';
    var tauxBg  = row[2] ? tagColor(row[2]) : '#f9fafb';
    var tauxFg  = row[2] ? tagText(row[2])  : '#374151';
    html += '<tr style="background:' + bg + ';">'
      + '<td style="padding:6px 14px;">' + row[0] + '</td>'
      + '<td style="padding:6px 14px;text-align:center;font-weight:600;">' + (row[1] !== null ? row[1].toLocaleString('fr-FR') : '—') + '</td>'
      + '<td style="padding:6px 14px;text-align:center;background:' + tauxBg + ';color:' + tauxFg + ';font-weight:600;">' + tauxVal + '</td>'
      + '</tr>';
  }
  html += '</table>';
  return html;
}

// ── Fonction principale d'envoi ───────────────────────────────
function envoyerRapportHebdomadaire() {
  var arrets = rhGetArretsSemainePrecedente();
  var kpi    = rhGetKpiMensuel();

  var semaine    = arrets.semaine;
  var dateDebut  = rhFormatDate(arrets.startStr);
  var dateFin    = rhFormatDate(arrets.endStr);

  var tableauArrets = rhBuildTableauArrets(arrets);
  var tableauKPI    = rhBuildTableauKPI(kpi);

  var sujet = 'Rapport Hebdomadaire de Planification — S' + semaine + ' (' + dateDebut + ' → ' + dateFin + ')';

  var corps = `
<div style="font-family:Arial,sans-serif;color:#0A1E3F;font-size:14px;max-width:700px;">
  <p>Bonjour,</p>
  <p>Veuillez trouver ci-dessous le <strong>rapport hebdomadaire de planification</strong> pour la semaine
  <strong>S${semaine}</strong> (${dateDebut} → ${dateFin}).</p>

  <hr style="border:none;border-top:2px solid #002060;margin:20px 0;">

  <h3 style="color:#002060;margin-bottom:10px;">
    📅 Calendrier des arrêts — Semaine S${semaine}
  </h3>
  <p style="color:#6b7280;font-size:12px;margin-bottom:10px;">
    ${arrets.rows.length} arrêt(s) enregistré(s) du ${dateDebut} au ${dateFin}
  </p>
  ${tableauArrets}

  <hr style="border:none;border-top:2px solid #002060;margin:20px 0;">

  <h3 style="color:#002060;margin-bottom:10px;">
    📊 KPIs Mensuels — ${kpi.mois}
  </h3>
  ${tableauKPI}

  <br>
  <p>Cordialement,</p>
  <div style="font-family:'Times New Roman',serif;font-size:14px;color:#002060;line-height:1.5;">
    <span style="font-weight:bold;">Marouane ELAMRAOUI</span><br>
    <span style="color:#c55a11;">Méthode de Maintenance</span><br>
    <span style="font-weight:bold;">OCP SA - Khouribga</span><br>
    <span style="color:green;">Tél. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:green;">Cisco :</span> 8103388<br>
    <a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>
  </div>
</div>`;

  var to = OCP_EMAIL; // À remplacer par la liste des destinataires réels
  // var to = "dest1@ocpgroup.ma, dest2@ocpgroup.ma";
  // var cc = "m.elamraoui@ocpgroup.ma";

  sendEmailOCP(to, sujet, '', {
    htmlBody:   corps,
    name:       'Bureau de Méthode Daoui - Section Planification'
  });

  Logger.log('✅ Rapport hebdomadaire S' + semaine + ' envoyé à : ' + to);
}

// ── Fonction de test (envoi uniquement à soi-même) ───────────
function testerRapportHebdomadaire() {
  var arrets = rhGetArretsSemainePrecedente();
  var kpi    = rhGetKpiMensuel();

  var semaine   = arrets.semaine;
  var dateDebut = rhFormatDate(arrets.startStr);
  var dateFin   = rhFormatDate(arrets.endStr);

  Logger.log('Semaine précédente : S' + semaine + ' (' + arrets.startStr + ' → ' + arrets.endStr + ')');
  Logger.log('Arrêts trouvés : ' + arrets.rows.length);
  Logger.log('KPI mois : ' + JSON.stringify(kpi));

  var tableauArrets = rhBuildTableauArrets(arrets);
  var tableauKPI    = rhBuildTableauKPI(kpi);

  var corps = `
<div style="font-family:Arial,sans-serif;color:#0A1E3F;font-size:14px;max-width:700px;">
  <p><strong>[TEST]</strong> Rapport Hebdomadaire — S${semaine} (${dateDebut} → ${dateFin})</p>

  <h3 style="color:#002060;">📅 Arrêts — S${semaine}</h3>
  <p style="color:#6b7280;font-size:12px;">${arrets.rows.length} arrêt(s) du ${dateDebut} au ${dateFin}</p>
  ${tableauArrets}

  <h3 style="color:#002060;margin-top:20px;">📊 KPIs — ${kpi.mois}</h3>
  ${tableauKPI}
</div>`;

  sendEmailOCP(OCP_EMAIL, '[TEST] Rapport Hebdomadaire S' + semaine, '', {
    htmlBody: corps,
    name:     'Bureau de Méthode Daoui - Section Planification'
  });

  Logger.log('✅ Test envoyé à : ' + OCP_EMAIL);
}
