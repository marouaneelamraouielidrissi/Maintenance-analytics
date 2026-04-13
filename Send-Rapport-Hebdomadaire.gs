// ═══════════════════════════════════════════════════════════════
//  RAPPORT HEBDOMADAIRE DE PLANIFICATION
//  Calendrier des arrêts (semaine précédente) + KPIs mensuels
//  Style visuel identique à l'application Maintenance Analytics
// ═══════════════════════════════════════════════════════════════

const RH_OT_FILE_ID     = '1aQAvb1DUv6Vk1Y1C-WEYgQnYN1BxujEAg8lbMt1sP3s';
const RH_ARRETS_FILE_ID = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
const RH_ARRETS_SHEET   = 'Travaux hebdomadaire';

// ── Utilitaires ────────────────────────────────────────────────
function rhGetWeekNumber(d) {
  var dc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  dc.setUTCDate(dc.getUTCDate() + 4 - (dc.getUTCDay() || 7));
  var yearStart = new Date(Date.UTC(dc.getUTCFullYear(), 0, 1));
  return Math.ceil((((dc - yearStart) / 86400000) + 1) / 7);
}

function rhDateToStr(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function rhFormatDate(dateStr) {
  var p = dateStr.split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : dateStr;
}

function rhGetPrevWeekRange() {
  var today = new Date();
  var dow   = today.getDay() || 7;
  var mon   = new Date(today); mon.setDate(today.getDate() - (dow - 1));
  var monPrev = new Date(mon); monPrev.setDate(mon.getDate() - 7);
  var sunPrev = new Date(monPrev); sunPrev.setDate(monPrev.getDate() + 6);
  return { start: monPrev, end: sunPrev };
}

function rhGetMondayOfISOWeek(year, week) {
  var jan4    = new Date(year, 0, 4);
  var jan4Day = jan4.getDay() || 7;
  var monday  = new Date(jan4);
  monday.setDate(jan4.getDate() - (jan4Day - 1) + (week - 1) * 7);
  return monday;
}

// ── Lecture des arrêts ─────────────────────────────────────────
function rhGetArrets() {
  var range    = rhGetPrevWeekRange();
  var startStr = rhDateToStr(range.start);
  var endStr   = rhDateToStr(range.end);
  var semaine  = rhGetWeekNumber(range.start);

  var ss    = SpreadsheetApp.openById(RH_ARRETS_FILE_ID);
  var sheet = ss.getSheetByName(RH_ARRETS_SHEET);
  if (!sheet) return { rows: [], startStr: startStr, endStr: endStr, semaine: semaine, annee: range.start.getFullYear() };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return h.toString().trim().toLowerCase(); });

  function col(names) {
    for (var i = 0; i < names.length; i++) {
      var idx = headers.indexOf(names[i].toLowerCase());
      if (idx >= 0) return idx;
    }
    return -1;
  }
  var cDate    = col(['start date','date début','date debut','date']);
  var cInstall = col(['installation','équipement','equipement']);
  var cSection = col(['section']);
  var cSemaine = col(['semaine']);
  var cAnnee   = col(['année','annee']);
  var cReal    = col(['réalisation','realisation']);

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rawDate = cDate >= 0 ? r[cDate] : null;
    var dateStr = '';
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      dateStr = rhDateToStr(rawDate);
    } else if (typeof rawDate === 'number') {
      var d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
      dateStr = rhDateToStr(d);
    } else if (rawDate) {
      dateStr = rawDate.toString().trim();
      var parts = dateStr.split('/');
      if (parts.length === 3 && parts[2].length === 4) {
        dateStr = parts[2] + '-' + parts[1].padStart(2,'0') + '-' + parts[0].padStart(2,'0');
      }
    }
    if (!dateStr || dateStr < startStr || dateStr > endStr) continue;

    var realRaw = (cReal >= 0 ? r[cReal] : '').toString().trim().toLowerCase();
    var statut  = 'nonreal';
    if (realRaw === 'oui') statut = 'realise';
    else if (realRaw.includes('imprévu') || realRaw.includes('imprevu')) statut = 'imprevu';

    rows.push({
      date:         dateStr,
      installation: cInstall >= 0 ? r[cInstall].toString().trim() : '',
      section:      cSection >= 0 ? r[cSection].toString().trim() : '',
      semaine:      cSemaine >= 0 ? r[cSemaine].toString().trim() : String(semaine),
      annee:        cAnnee  >= 0 ? parseInt(r[cAnnee]) : range.start.getFullYear(),
      statut:       statut
    });
  }

  rows.sort(function(a, b){ return a.date.localeCompare(b.date); });
  return { rows: rows, startStr: startStr, endStr: endStr, semaine: semaine, annee: range.start.getFullYear() };
}

// ── KPIs mensuels ─────────────────────────────────────────────
function rhGetKpi() {
  var now   = new Date();
  var year  = now.getFullYear();
  var month = now.getMonth();

  var ss    = SpreadsheetApp.openById(RH_OT_FILE_ID);
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];

  function cIdx(names) {
    for (var i = 0; i < names.length; i++)
      for (var j = 0; j < hdrs.length; j++)
        if (hdrs[j].toString().trim().toLowerCase() === names[i].toLowerCase()) return j;
    return -1;
  }
  var cDebut = cIdx(['début au plus tôt','debut au plus tot','date début','date debut']);
  var cStat  = cIdx(['statut système','statut systeme','statut sys']);
  var cUtil  = cIdx(['statut utilis.','statut util','statut utilis']);
  var cType  = cIdx(["type d'ordre",'type ordre','type']);
  var cPoste = cIdx(['poste de travail','poste travail','poste']);

  var MONTHS_FR = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];

  var total=0, realise=0, lanc=0, crpr=0, sys=0, cur=0, sysReal=0, curReal=0, backlog=0;
  var posteMap = {};

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var rawD = cDebut >= 0 ? r[cDebut] : null;
    var d;
    if (rawD instanceof Date) d = rawD;
    else if (typeof rawD === 'number') d = new Date(Math.round((rawD - 25569) * 86400 * 1000));
    else if (rawD) d = new Date(rawD);
    else continue;
    if (isNaN(d) || d.getFullYear() !== year || d.getMonth() !== month) continue;

    total++;
    var ss_ = cStat  >= 0 ? r[cStat ].toString() : '';
    var su  = cUtil  >= 0 ? r[cUtil ].toString() : '';
    var tp  = cType  >= 0 ? r[cType ].toString() : '';
    var pt  = cPoste >= 0 ? r[cPoste].toString().trim() : '';

    var isReal = ss_.includes('CONF') || ss_.includes('TCLO') || ss_.includes('CLOT');
    if (isReal) realise++;
    if (ss_.includes('LANC') && !ss_.includes('CONF') && !ss_.includes('TCLO')) lanc++;
    if (su.includes('CRPR')) crpr++;
    if (su.includes('ATPL') && ss_.includes('LANC')) backlog++;

    var isSys = ['ZCON','ZEST','ZETL'].indexOf(tp) >= 0;
    var isCur = tp === 'ZCOR';
    if (isSys) { sys++; if (isReal) sysReal++; }
    if (isCur) { cur++; if (isReal) curReal++; }

    if (pt) {
      if (!posteMap[pt]) posteMap[pt] = { total: 0, real: 0 };
      posteMap[pt].total++;
      if (isReal) posteMap[pt].real++;
    }
  }

  function pct(n, t) { return t ? parseFloat(((n/t)*100).toFixed(1)) : 0; }
  function pctStr(n, t) { return t ? pct(n,t).toFixed(1)+'%' : '—'; }

  var postes = Object.keys(posteMap).map(function(p){
    return { nom: p, total: posteMap[p].total, real: posteMap[p].real, taux: pct(posteMap[p].real, posteMap[p].total) };
  }).sort(function(a,b){ return b.total - a.total; }).slice(0, 10);

  return {
    mois: MONTHS_FR[month] + ' ' + year,
    total: total, realise: realise, tauxReal: pct(realise, total), tauxRealStr: pctStr(realise, total),
    lanc: lanc, lancPct: pctStr(lanc, total),
    crpr: crpr, crprPct: pctStr(crpr, total),
    sys: sys, sysPct: pctStr(sys, total),
    cur: cur, curPct: pctStr(cur, total),
    tauxPrev: pct(sysReal, sys), tauxPrevStr: pctStr(sysReal, sys),
    tauxCor: pct(curReal, cur), tauxCorStr: pctStr(curReal, cur),
    backlog: backlog,
    postes: postes
  };
}

// ── Helpers HTML (inline CSS = compatible email) ──────────────
function rhTagColor(v) {
  if (v >= 80) return { bg: '#dcfce7', color: '#166534' };
  if (v >= 50) return { bg: '#fef9c3', color: '#854d0e' };
  return { bg: '#fee2e2', color: '#991b1b' };
}

function rhKpiCard(iconBg, iconStroke, svgPath, value, label, sub, tagVal) {
  var tagHtml = '';
  if (tagVal !== undefined && tagVal !== null) {
    var tc = rhTagColor(tagVal);
    tagHtml = '<span style="font-size:11px;font-weight:600;border-radius:4px;padding:2px 8px;background:' + tc.bg + ';color:' + tc.color + ';">' + tagVal.toFixed(1) + '%</span>';
  }
  return '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:18px 20px 16px;box-shadow:0 1px 4px rgba(0,0,0,0.06);min-width:160px;flex:1;">'
    + '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px;">'
    + '<div style="width:36px;height:36px;border-radius:8px;background:' + iconBg + ';display:flex;align-items:center;justify-content:center;flex-shrink:0;">'
    + '<svg viewBox="0 0 24 24" style="width:16px;height:16px;stroke:' + iconStroke + ';fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;">' + svgPath + '</svg>'
    + '</div>' + tagHtml + '</div>'
    + '<div style="font-size:28px;font-weight:700;letter-spacing:-1.5px;line-height:1;margin-bottom:4px;color:#0f172a;">' + value + '</div>'
    + '<div style="font-size:12px;font-weight:500;color:#475569;">' + label + '</div>'
    + '<div style="font-size:11px;color:#94a3b8;margin-top:10px;padding-top:10px;border-top:1px solid #f1f5f9;">' + sub + '</div>'
    + '</div>';
}

function rhSectionLabel(text) {
  return '<div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;margin:28px 0 12px;display:flex;align-items:center;gap:10px;">'
    + text + '<span style="flex:1;height:1px;background:#e2e8f0;display:inline-block;margin-left:8px;"></span></div>';
}

function rhSubSection(bg, borderColor, iconColor, svgPath, label) {
  return '<div style="font-size:10px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;padding:5px 12px;border-radius:5px;display:inline-flex;align-items:center;gap:7px;margin-bottom:10px;margin-top:16px;background:' + bg + ';color:' + iconColor + ';border:1px solid ' + borderColor + ';">'
    + '<svg viewBox="0 0 24 24" style="width:12px;height:12px;stroke:' + iconColor + ';fill:none;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round;">' + svgPath + '</svg>'
    + label + '</div>';
}

// ── Construction du calendrier HTML ───────────────────────────
function rhBuildCalendrier(arrets) {
  var DAYS_FR = ['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];

  if (!arrets.rows.length) {
    return '<p style="color:#94a3b8;font-style:italic;font-size:13px;padding:12px 0;">Aucun arrêt enregistré pour la semaine S' + arrets.semaine + '.</p>';
  }

  // Grouper par semaine → par jour
  var weekMap = {};
  arrets.rows.forEach(function(r) {
    var parts = r.date.split('-').map(Number);
    var d     = new Date(parts[0], parts[1]-1, parts[2]);
    var wKey  = r.annee + '-' + r.semaine;
    if (!weekMap[wKey]) weekMap[wKey] = { annee: r.annee, semaine: r.semaine, days: {} };
    var dow = (d.getDay() + 6) % 7; // 0=Lun … 6=Dim
    if (!weekMap[wKey].days[dow]) weekMap[wKey].days[dow] = [];
    weekMap[wKey].days[dow].push({ label: r.installation, statut: r.statut });
  });

  var weeks = Object.values(weekMap).sort(function(a,b){
    if (a.annee !== b.annee) return a.annee - b.annee;
    return parseInt(a.semaine) - parseInt(b.semaine);
  });

  var cellBase  = 'padding:6px 7px;border:1px solid #e2e8f0;vertical-align:top;min-width:100px;';
  var hdrBase   = 'font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.5px;background:#f8fafc;padding:7px 8px;border:1px solid #e2e8f0;text-align:center;';
  var weekHdr   = 'font-size:11px;font-weight:700;color:#0f172a;background:#f8fafc;padding:8px 12px;border:1px solid #e2e8f0;white-space:nowrap;';
  var dateStyle = 'font-size:9px;font-weight:500;color:#94a3b8;text-align:center;padding:3px 6px;background:#f8fafc;border:1px solid #e2e8f0;border-top:none;font-family:monospace;';

  var html = '<table style="width:100%;border-collapse:collapse;font-size:12px;">';

  // Légende
  html += '<tr><td colspan="8" style="padding:0 0 10px 0;border:none;">'
    + '<span style="font-size:11px;color:#64748b;margin-right:14px;">● <span style="color:#166534;font-weight:600;">Réalisé</span></span>'
    + '<span style="font-size:11px;color:#64748b;margin-right:14px;">● <span style="color:#991b1b;font-weight:600;">Non réalisé</span></span>'
    + '<span style="font-size:11px;color:#64748b;">● <span style="color:#9a3412;font-weight:600;">Imprévu</span></span>'
    + '</td></tr>';

  // En-tête jours
  html += '<tr><th style="' + weekHdr + '">Semaine</th>';
  DAYS_FR.forEach(function(d){ html += '<th style="' + hdrBase + '">' + d + '</th>'; });
  html += '</tr>';

  weeks.forEach(function(week, wi) {
    var wNum   = parseInt(week.semaine);
    var monday = rhGetMondayOfISOWeek(week.annee, wNum);
    var sep    = wi > 0 ? 'border-top:2px solid #cbd5e1;' : '';

    // Ligne dates
    html += '<tr>';
    html += '<td style="' + weekHdr + sep + 'font-weight:700;color:#1d4ed8;">' + week.semaine + '</td>';
    for (var di = 0; di < 7; di++) {
      var d  = new Date(monday); d.setDate(d.getDate() + di);
      var dd = String(d.getDate()).padStart(2,'0');
      var mm = String(d.getMonth()+1).padStart(2,'0');
      html += '<td style="' + dateStyle + sep + '">' + dd + '/' + mm + '</td>';
    }
    html += '</tr>';

    // Ligne données
    html += '<tr>';
    html += '<td style="font-family:monospace;font-size:9px;font-weight:700;color:#94a3b8;vertical-align:middle;text-align:center;' + cellBase + '">' + week.annee + '</td>';
    for (var di = 0; di < 7; di++) {
      var items = week.days[di] || [];
      html += '<td style="' + cellBase + '">';
      if (!items.length) {
        html += '<span style="color:#cbd5e1;font-size:12px;">·</span>';
      } else {
        items.forEach(function(it){
          var tagBg, tagColor;
          if (it.statut === 'realise')      { tagBg='#dcfce7'; tagColor='#166534'; }
          else if (it.statut === 'imprevu') { tagBg='#ffedd5'; tagColor='#9a3412'; }
          else                               { tagBg='#fee2e2'; tagColor='#991b1b'; }
          html += '<span style="display:inline-block;font-size:10px;font-weight:600;padding:3px 7px;border-radius:3px;margin:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px;background:' + tagBg + ';color:' + tagColor + ';">' + it.label + '</span>';
        });
      }
      html += '</td>';
    }
    html += '</tr>';
  });

  html += '</table>';
  return html;
}

// ── Construction des KPI par poste ────────────────────────────
function rhBuildPostes(postes) {
  if (!postes.length) return '<p style="color:#94a3b8;font-size:12px;font-style:italic;">Aucune donnée par poste.</p>';
  var html = '';
  postes.forEach(function(p) {
    var tc      = rhTagColor(p.taux);
    var barColor = p.taux >= 80 ? '#059669' : p.taux >= 50 ? '#d97706' : '#dc2626';
    var w       = Math.min(100, p.taux).toFixed(0);
    html += '<div style="display:flex;align-items:center;gap:12px;padding:9px 0;border-bottom:1px solid #f1f5f9;">'
      + '<div style="font-size:12px;font-weight:600;min-width:130px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + p.nom + '</div>'
      + '<div style="flex:1;height:10px;background:#f1f5f9;border-radius:5px;overflow:hidden;border:1px solid #e2e8f0;">'
      + '<div style="height:100%;border-radius:5px;width:' + w + '%;background:' + barColor + ';"></div>'
      + '</div>'
      + '<div style="font-family:monospace;font-size:12px;font-weight:600;min-width:44px;text-align:right;color:' + barColor + ';">' + p.taux.toFixed(1) + '%</div>'
      + '<div style="font-size:11px;color:#94a3b8;min-width:54px;text-align:right;">' + p.real + '/' + p.total + '</div>'
      + '</div>';
  });
  return html;
}

// ── Corps email principal ─────────────────────────────────────
function rhBuildEmail(arrets, kpi) {
  var semaine   = arrets.semaine;
  var dateDebut = rhFormatDate(arrets.startStr);
  var dateFin   = rhFormatDate(arrets.endStr);

  var calendrier = rhBuildCalendrier(arrets);
  var postesHtml = rhBuildPostes(kpi.postes);

  var cardStyle = 'background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:20px 22px 18px;box-shadow:0 1px 4px rgba(0,0,0,0.06);margin-bottom:16px;';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f8fafc;font-family:Inter,Arial,sans-serif;color:#0f172a;">'

  // ── HEADER ──
  + '<div style="background:#1d4ed8;padding:20px 32px;display:flex;align-items:center;gap:12px;">'
  + '<div style="width:36px;height:36px;background:rgba(255,255,255,0.2);border-radius:6px;display:inline-flex;align-items:center;justify-content:center;">'
  + '<svg viewBox="0 0 24 24" style="width:18px;height:18px;fill:none;stroke:white;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>'
  + '</div>'
  + '<div style="display:inline-block;margin-left:10px;">'
  + '<div style="color:white;font-weight:700;font-size:15px;letter-spacing:-0.3px;">Maintenance Analytics</div>'
  + '<div style="color:rgba(255,255,255,0.7);font-size:11px;">Bureau Méthode Daoui · OCP SA Khouribga</div>'
  + '</div>'
  + '<div style="margin-left:auto;background:rgba(255,255,255,0.15);border-radius:6px;padding:5px 14px;color:white;font-weight:700;font-size:13px;letter-spacing:-0.3px;display:inline-block;">'
  + 'S' + semaine + ' · ' + dateDebut + ' → ' + dateFin
  + '</div>'
  + '</div>'

  // ── CONTENU ──
  + '<div style="padding:28px 32px;max-width:900px;margin:0 auto;">'

  // ── TITRE ──
  + '<h2 style="font-size:20px;font-weight:700;letter-spacing:-0.4px;margin:0 0 4px;color:#0f172a;">Rapport Hebdomadaire de Planification</h2>'
  + '<p style="font-size:13px;color:#64748b;margin:0 0 24px;">Semaine <strong>S' + semaine + '</strong> · ' + dateDebut + ' au ' + dateFin + ' · Généré le ' + new Date().toLocaleDateString('fr-FR', {day:'2-digit',month:'long',year:'numeric'}) + '</p>'

  // ══ SECTION 1 : CALENDRIER ══
  + rhSectionLabel('Calendrier des arrêts préventifs')
  + '<div style="' + cardStyle + '">'
  + '<div style="margin-bottom:14px;">'
  + '<div style="font-size:13px;font-weight:600;color:#0f172a;">Calendrier des arrêts — Semaine S' + semaine + '</div>'
  + '<div style="font-size:11px;color:#94a3b8;margin-top:2px;">' + arrets.rows.length + ' arrêt(s) enregistré(s) · ' + dateDebut + ' → ' + dateFin + '</div>'
  + '</div>'
  + '<div style="overflow-x:auto;">' + calendrier + '</div>'
  + '</div>'

  // ══ SECTION 2 : KPIs ══
  + rhSectionLabel('Indicateurs clés du mois — ' + kpi.mois)

  // Global
  + rhSubSection('#eff6ff','#bfdbfe','#1d4ed8','<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>','Global')
  + '<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  + rhKpiCard('#eff6ff','#1d4ed8','<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
      kpi.total.toLocaleString('fr-FR'), 'Total OT planifiés', 'Ordres de travail du mois', null)
  + rhKpiCard('#ecfdf5','#059669','<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
      kpi.realise.toLocaleString('fr-FR'), 'OT Réalisés', 'Taux de réalisation : <b style="color:#0f172a;">' + kpi.tauxRealStr + '</b>', kpi.tauxReal)
  + rhKpiCard('#fffbeb','#d97706','<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
      kpi.lanc.toLocaleString('fr-FR'), 'OT Lancés', 'En cours : <b style="color:#0f172a;">' + kpi.lancPct + '</b>', null)
  + rhKpiCard('#fef2f2','#dc2626','<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>',
      kpi.crpr.toLocaleString('fr-FR'), 'Non lancés (CRPR)', 'Part du total : <b style="color:#0f172a;">' + kpi.crprPct + '</b>', null)
  + rhKpiCard('#f1f5f9','#475569','<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
      kpi.backlog.toLocaleString('fr-FR'), 'Backlog', 'ATPL + LANC', null)
  + '</div>'

  // Préventif
  + rhSubSection('#ecfdf5','#a7f3d0','#059669','<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>','Préventif')
  + '<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  + rhKpiCard('#f5f3ff','#7c3aed','<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
      kpi.sys.toLocaleString('fr-FR'), 'OT Préventif systématique', 'ZCON + ZEST + ZETL : <b style="color:#0f172a;">' + kpi.sysPct + '</b>', null)
  + rhKpiCard('#ecfdf5','#059669','<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
      kpi.tauxPrevStr, 'Taux réalisation préventif', 'CONF+TCLO+CLOT / total ZCON+ZEST+ZETL', kpi.tauxPrev)
  + '</div>'

  // Correctif
  + rhSubSection('#fef2f2','#fca5a5','#dc2626','<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>','Correctif')
  + '<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  + rhKpiCard('#fef2f2','#dc2626','<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
      kpi.cur.toLocaleString('fr-FR'), 'OT Correctif (ZCOR)', 'Part du total : <b style="color:#0f172a;">' + kpi.curPct + '</b>', null)
  + rhKpiCard('#fef2f2','#dc2626','<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>',
      kpi.tauxCorStr, 'Taux réalisation correctif', 'CONF+TCLO+CLOT / total ZCOR', kpi.tauxCor)
  + '</div>'

  // ══ SECTION 3 : RÉALISATION PAR POSTE ══
  + rhSectionLabel('Taux de réalisation par corps de métier')
  + '<div style="' + cardStyle + '">'
  + '<div style="margin-bottom:14px;">'
  + '<div style="font-size:13px;font-weight:600;color:#0f172a;">Taux de réalisation — ' + kpi.mois + '</div>'
  + '<div style="font-size:11px;color:#94a3b8;margin-top:2px;">OT confirmés ou techniquement clos (CONF + TCLO + CLOT) / total · '
  + '<span style="color:#059669;font-weight:600;">≥ 80% Bon</span> · '
  + '<span style="color:#d97706;font-weight:600;">50–79% Moyen</span> · '
  + '<span style="color:#dc2626;font-weight:600;">&lt; 50% Faible</span></div>'
  + '</div>'
  + postesHtml
  + '</div>'

  // ── SIGNATURE ──
  + '<div style="margin-top:32px;padding-top:20px;border-top:1px solid #e2e8f0;">'
  + '<p style="margin:0 0 10px;color:#64748b;font-size:13px;">Cordialement,</p>'
  + '<div style="font-family:Georgia,serif;font-size:14px;color:#002060;line-height:1.6;">'
  + '<strong>Marouane ELAMRAOUI</strong><br>'
  + '<span style="color:#c55a11;">Méthode de Maintenance</span><br>'
  + '<strong>OCP SA - Khouribga</strong><br>'
  + '<span style="color:#059669;">Tél. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:#059669;">Cisco :</span> 8103388<br>'
  + '<a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>'
  + '</div>'
  + '</div>'

  + '</div>' // fin contenu
  + '</body></html>';
}

// ── Envoi principal ───────────────────────────────────────────
function envoyerRapportHebdomadaire() {
  var arrets = rhGetArrets();
  var kpi    = rhGetKpi();
  var html   = rhBuildEmail(arrets, kpi);

  var semaine = arrets.semaine;
  var sujet   = 'Rapport Hebdomadaire de Planification — S' + semaine
    + ' (' + rhFormatDate(arrets.startStr) + ' → ' + rhFormatDate(arrets.endStr) + ')';

  var to = OCP_EMAIL; // Remplace par la liste réelle des destinataires
  // var to = "dest1@ocpgroup.ma, dest2@ocpgroup.ma";
  // var cc = "m.elamraoui@ocpgroup.ma";

  sendEmailOCP(to, sujet, '', { htmlBody: html, name: 'Bureau de Méthode Daoui - Section Planification' });
  Logger.log('✅ Rapport hebdomadaire S' + semaine + ' envoyé.');
}

// ── Test ──────────────────────────────────────────────────────
function testerRapportHebdomadaire() {
  var arrets = rhGetArrets();
  var kpi    = rhGetKpi();
  var html   = rhBuildEmail(arrets, kpi);

  Logger.log('Semaine : S' + arrets.semaine + ' (' + arrets.startStr + ' → ' + arrets.endStr + ')');
  Logger.log('Arrêts : ' + arrets.rows.length);
  Logger.log('KPI : total=' + kpi.total + ' réalisé=' + kpi.realise + ' taux=' + kpi.tauxRealStr);

  sendEmailOCP(OCP_EMAIL, '[TEST] Rapport Hebdomadaire S' + arrets.semaine, '', {
    htmlBody: html,
    name: 'Bureau de Méthode Daoui - Section Planification'
  });
  Logger.log('✅ Test envoyé à : ' + OCP_EMAIL);
}
