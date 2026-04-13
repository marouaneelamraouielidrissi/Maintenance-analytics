// ═══════════════════════════════════════════════════════════════
//  RAPPORT HEBDOMADAIRE — Backend GAS
//  Sert l'interface HTML + envoie les rapports via EWS (OCP)
// ═══════════════════════════════════════════════════════════════

// ── IDs des fichiers ──────────────────────────────────────────
const RH_OT_FILE_ID     = '1aQAvb1DUv6Vk1Y1C-WEYgQnYN1BxujEAg8lbMt1sP3s';
const RH_ARRETS_FILE_ID = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
const RH_ARRETS_SHEET   = 'Travaux hebdomadaire';

// ── Configuration OCP Exchange (EWS) ─────────────────────────
const RH_OCP_EMAIL = 'm.elamraoui@ocpgroup.ma';
const RH_EWS_URL   = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';

function getRhPassword() {
  return PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD') || '';
}

function sendEmailRH(to, subject, htmlBody, senderName) {
  var toList = Array.isArray(to) ? to : to.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
  var boundary = 'rh_boundary_' + Date.now();
  var subjB64  = Utilities.base64Encode(subject,  Utilities.Charset.UTF_8);
  var bodyB64  = Utilities.base64Encode(htmlBody,  Utilities.Charset.UTF_8);

  var mimeParts = [
    'From: "' + senderName + '" <' + RH_OCP_EMAIL + '>',
    'To: ' + toList.join(', '),
    'Subject: =?UTF-8?B?' + subjB64 + '?=',
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    bodyB64,
    '',
    '--' + boundary + '--'
  ];

  var mime    = mimeParts.join('\r\n');
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

  var credentials = Utilities.base64Encode(RH_OCP_EMAIL + ':' + getRhPassword());
  var response = UrlFetchApp.fetch(RH_EWS_URL, {
    method: 'post', contentType: 'text/xml; charset=utf-8',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'SOAPAction': 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
    },
    payload: soap, muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code !== 200 || text.indexOf('NoError') === -1) {
    throw new Error('EWS send failed (' + code + '): ' + text.substring(0, 400));
  }
}

// ── Servir l'interface HTML ───────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Interface')
    .setTitle('Rapport Hebdomadaire')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Config par défaut pour l'interface ───────────────────────
function getConfigInterfaceDefaut() {
  var now = new Date();
  var props = PropertiesService.getScriptProperties();
  return {
    mo:     now.getMonth(),
    yr:     now.getFullYear(),
    mode:   'month',
    emails: props.getProperty('RH_EMAILS') || RH_OCP_EMAIL
  };
}

// ── Utilitaires date ──────────────────────────────────────────
function rhWeekNum(d) {
  var dc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  dc.setUTCDate(dc.getUTCDate() + 4 - (dc.getUTCDay() || 7));
  var ys = new Date(Date.UTC(dc.getUTCFullYear(), 0, 1));
  return Math.ceil((((dc - ys) / 86400000) + 1) / 7);
}

function rhDateStr(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function rhFmtDate(s) {
  var p = s.split('-');
  return p.length === 3 ? p[2]+'/'+p[1]+'/'+p[0] : s;
}

function rhMondayOf(year, week) {
  var jan4    = new Date(year, 0, 4);
  var jan4day = jan4.getDay() || 7;
  var m       = new Date(jan4);
  m.setDate(jan4.getDate() - (jan4day - 1) + (week - 1) * 7);
  return m;
}

// ── Lecture des arrêts (semaine précédente) ───────────────────
function rhGetArrets() {
  var today = new Date();
  var dow   = today.getDay() || 7;
  var mon   = new Date(today); mon.setDate(today.getDate() - (dow - 1));
  var monP  = new Date(mon);   monP.setDate(mon.getDate() - 7);
  var sunP  = new Date(monP);  sunP.setDate(monP.getDate() + 6);
  var s0    = rhDateStr(monP), s1 = rhDateStr(sunP);
  var sem   = rhWeekNum(monP);

  var ss    = SpreadsheetApp.openById(RH_ARRETS_FILE_ID);
  var sheet = ss.getSheetByName(RH_ARRETS_SHEET);
  if (!sheet) return { rows:[], s0:s0, s1:s1, sem:sem, annee:monP.getFullYear() };

  var data = sheet.getDataRange().getValues();
  var hdr  = data[0].map(function(h){ return h.toString().trim().toLowerCase(); });

  function ci(names) {
    for (var i=0;i<names.length;i++){ var x=hdr.indexOf(names[i]); if(x>=0) return x; }
    return -1;
  }
  var cD=ci(['start date','date début','date debut','date']);
  var cI=ci(['installation','équipement','equipement']);
  var cS=ci(['section']);
  var cW=ci(['semaine']);
  var cA=ci(['année','annee']);
  var cR=ci(['réalisation','realisation']);

  var rows = [];
  for (var i=1;i<data.length;i++) {
    var r=data[i], raw=cD>=0?r[cD]:null, ds='';
    if (raw instanceof Date && !isNaN(raw))  ds = rhDateStr(raw);
    else if (typeof raw==='number')           ds = rhDateStr(new Date(Math.round((raw-25569)*86400000)));
    else if (raw) { ds=raw.toString().trim(); var p=ds.split('/'); if(p.length===3&&p[2].length===4) ds=p[2]+'-'+p[1].padStart(2,'0')+'-'+p[0].padStart(2,'0'); }
    if (!ds||ds<s0||ds>s1) continue;
    var rv=(cR>=0?r[cR]:'').toString().trim().toLowerCase();
    var st=rv==='oui'?'realise':rv.includes('imprévu')||rv.includes('imprevu')?'imprevu':'nonreal';
    rows.push({ date:ds, install:cI>=0?r[cI].toString().trim():'', section:cS>=0?r[cS].toString().trim():'', semaine:cW>=0?r[cW].toString().trim():String(sem), annee:cA>=0?parseInt(r[cA]):monP.getFullYear(), statut:st });
  }
  rows.sort(function(a,b){return a.date.localeCompare(b.date);});
  return { rows:rows, s0:s0, s1:s1, sem:sem, annee:monP.getFullYear() };
}

// ── Lecture des KPIs OT ───────────────────────────────────────
function rhGetKpi(mo, yr) {
  var ss    = SpreadsheetApp.openById(RH_OT_FILE_ID);
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];

  function ci(names) {
    for (var i=0;i<names.length;i++)
      for (var j=0;j<hdrs.length;j++)
        if (hdrs[j].toString().trim().toLowerCase()===names[i].toLowerCase()) return j;
    return -1;
  }
  var cDeb  = ci(['début au plus tôt','debut au plus tot','date début','date debut']);
  var cStat = ci(['statut système','statut systeme','statut sys']);
  var cUtil = ci(['statut utilis.','statut util','statut utilis']);
  var cType = ci(["type d'ordre",'type ordre','type']);
  var cPost = ci(['poste de travail','poste travail','poste']);

  var total=0,real=0,lanc=0,crpr=0,sys=0,cur=0,sysR=0,curR=0,backlog=0;
  var posteMap={};

  for (var i=1;i<data.length;i++) {
    var r=data[i], rawD=cDeb>=0?r[cDeb]:null, d;
    if (rawD instanceof Date) d=rawD;
    else if (typeof rawD==='number') d=new Date(Math.round((rawD-25569)*86400000));
    else if (rawD) d=new Date(rawD);
    else continue;
    if (isNaN(d)||d.getFullYear()!==yr||d.getMonth()!==mo) continue;
    total++;
    var ss_=cStat>=0?r[cStat].toString():'', su=cUtil>=0?r[cUtil].toString():'', tp=cType>=0?r[cType].toString():'', pt=cPost>=0?r[cPost].toString().trim():'';
    var isR=ss_.includes('CONF')||ss_.includes('TCLO')||ss_.includes('CLOT');
    if(isR) real++;
    if(ss_.includes('LANC')&&!ss_.includes('CONF')&&!ss_.includes('TCLO')) lanc++;
    if(su.includes('CRPR')) crpr++;
    if(su.includes('ATPL')&&ss_.includes('LANC')) backlog++;
    var isSys=['ZCON','ZEST','ZETL'].indexOf(tp)>=0, isCur=tp==='ZCOR';
    if(isSys){sys++;if(isR)sysR++;}
    if(isCur){cur++;if(isR)curR++;}
    if(pt){if(!posteMap[pt])posteMap[pt]={total:0,real:0};posteMap[pt].total++;if(isR)posteMap[pt].real++;}
  }

  function p(n,t){return t?parseFloat(((n/t)*100).toFixed(1)):0;}
  function ps(n,t){return t?p(n,t).toFixed(1)+'%':'—';}

  var postes=Object.keys(posteMap).map(function(k){return{nom:k,total:posteMap[k].total,real:posteMap[k].real,taux:p(posteMap[k].real,posteMap[k].total)};}).sort(function(a,b){return b.total-a.total;}).slice(0,10);

  var MOIS=['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  return {
    mois:MOIS[mo]+' '+yr, mo:mo, yr:yr,
    total:total, real:real, tauxReal:p(real,total), tauxRealStr:ps(real,total),
    lanc:lanc, lancPct:ps(lanc,total),
    crpr:crpr, crprPct:ps(crpr,total),
    backlog:backlog,
    sys:sys, sysPct:ps(sys,total),
    cur:cur, curPct:ps(cur,total),
    tauxPrev:p(sysR,sys), tauxPrevStr:ps(sysR,sys),
    tauxCor:p(curR,cur),  tauxCorStr:ps(curR,cur),
    postes:postes
  };
}

// ── Construction HTML du rapport ──────────────────────────────
function rhBuildHtml(arrets, kpi) {
  var s=arrets.sem, d0=rhFmtDate(arrets.s0), d1=rhFmtDate(arrets.s1);

  // ── Calendrier ──
  function buildCal() {
    if (!arrets.rows.length) return '<p style="color:#94a3b8;font-style:italic;font-size:13px;padding:8px 0;">Aucun arrêt enregistré pour la semaine S'+s+'.</p>';
    var DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];
    var wMap={};
    arrets.rows.forEach(function(r){
      var parts=r.date.split('-').map(Number), d=new Date(parts[0],parts[1]-1,parts[2]);
      var k=r.annee+'-'+r.semaine;
      if(!wMap[k]) wMap[k]={annee:r.annee,semaine:r.semaine,days:{}};
      var dow=(d.getDay()+6)%7;
      if(!wMap[k].days[dow]) wMap[k].days[dow]=[];
      wMap[k].days[dow].push({label:r.install,statut:r.statut});
    });
    var weeks=Object.values(wMap).sort(function(a,b){return a.annee!==b.annee?a.annee-b.annee:parseInt(a.semaine)-parseInt(b.semaine);});
    var th='padding:7px 8px;border:1px solid #e2e8f0;font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.5px;background:#f8fafc;text-align:center;min-width:90px;';
    var wh='padding:8px 10px;border:1px solid #e2e8f0;font-size:11px;font-weight:700;background:#f8fafc;white-space:nowrap;';
    var tc='padding:5px 6px;border:1px solid #e2e8f0;vertical-align:top;min-width:90px;';
    var html='<table style="width:100%;border-collapse:collapse;font-size:12px;">';
    html+='<tr><th style="'+wh+'">Semaine</th>'+DAYS.map(function(d){return'<th style="'+th+'">'+d+'</th>';}).join('')+'</tr>';
    weeks.forEach(function(w,wi){
      var wn=parseInt(w.semaine), mon=rhMondayOf(w.annee,wn);
      var sep=wi>0?'border-top:2px solid #cbd5e1;':'';
      html+='<tr><td style="'+wh+sep+'color:#1d4ed8;">'+w.semaine+'</td>';
      for(var di=0;di<7;di++){var d=new Date(mon);d.setDate(d.getDate()+di);html+='<td style="font-size:9px;color:#94a3b8;text-align:center;padding:3px 4px;background:#f8fafc;border:1px solid #e2e8f0;'+sep+'">'+ String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'</td>';}
      html+='</tr><tr>';
      html+='<td style="font-size:9px;color:#94a3b8;text-align:center;font-weight:700;'+tc+'font-family:monospace;">'+w.annee+'</td>';
      for(var di=0;di<7;di++){
        var items=w.days[di]||[];
        html+='<td style="'+tc+'">';
        if(!items.length) html+='<span style="color:#cbd5e1;">·</span>';
        else items.forEach(function(it){
          var bg=it.statut==='realise'?'#dcfce7':it.statut==='imprevu'?'#ffedd5':'#fee2e2';
          var fg=it.statut==='realise'?'#166534':it.statut==='imprevu'?'#9a3412':'#991b1b';
          html+='<span style="display:inline-block;font-size:10px;font-weight:600;padding:2px 6px;border-radius:3px;margin:1px;background:'+bg+';color:'+fg+';white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100px;">'+it.label+'</span>';
        });
        html+='</td>';
      }
      html+='</tr>';
    });
    html+='</table>';
    return html;
  }

  // ── KPI card ──
  function kpiCard(iconBg,iconFg,svg,val,label,sub,taux) {
    var tagHtml='';
    if(taux!==null&&taux!==undefined){
      var tc=taux>=80?{bg:'#dcfce7',fg:'#166534'}:taux>=50?{bg:'#fef9c3',fg:'#854d0e'}:{bg:'#fee2e2',fg:'#991b1b'};
      tagHtml='<span style="font-size:11px;font-weight:600;border-radius:4px;padding:2px 8px;background:'+tc.bg+';color:'+tc.fg+';">'+taux.toFixed(1)+'%</span>';
    }
    return '<td style="width:50%;padding:6px;vertical-align:top;"><div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:16px 18px;height:100%;">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:12px;">'
      +'<div style="width:34px;height:34px;border-radius:8px;background:'+iconBg+';display:flex;align-items:center;justify-content:center;">'
      +'<svg viewBox="0 0 24 24" style="width:15px;height:15px;stroke:'+iconFg+';fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;">'+svg+'</svg></div>'
      +tagHtml+'</div>'
      +'<div style="font-size:26px;font-weight:700;letter-spacing:-1.5px;line-height:1;margin-bottom:4px;color:#0f172a;">'+val+'</div>'
      +'<div style="font-size:12px;font-weight:500;color:#475569;">'+label+'</div>'
      +'<div style="font-size:11px;color:#94a3b8;margin-top:8px;padding-top:8px;border-top:1px solid #f1f5f9;">'+sub+'</div>'
      +'</div></td>';
  }

  // ── Sous-section label ──
  function subSection(bg,border,color,svg,label) {
    return '<div style="font-size:10px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;padding:4px 11px;border-radius:5px;display:inline-flex;align-items:center;gap:6px;margin:14px 0 10px;background:'+bg+';color:'+color+';border:1px solid '+border+';">'
      +'<svg viewBox="0 0 24 24" style="width:12px;height:12px;stroke:'+color+';fill:none;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round;">'+svg+'</svg>'+label+'</div>';
  }

  // ── Section label ──
  function secLabel(txt) {
    return '<div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;margin:24px 0 12px;border-bottom:1px solid #e2e8f0;padding-bottom:8px;">'+txt+'</div>';
  }

  // ── Postes ──
  function buildPostes() {
    if(!kpi.postes.length) return '<p style="color:#94a3b8;font-size:12px;">Aucune donnée.</p>';
    return kpi.postes.map(function(p){
      var c=p.taux>=80?'#059669':p.taux>=50?'#d97706':'#dc2626';
      return '<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #f1f5f9;">'
        +'<div style="font-size:12px;font-weight:600;min-width:120px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'+p.nom+'</div>'
        +'<div style="flex:1;height:8px;background:#f1f5f9;border-radius:4px;overflow:hidden;">'
        +'<div style="height:100%;border-radius:4px;width:'+Math.min(100,p.taux).toFixed(0)+'%;background:'+c+';"></div></div>'
        +'<div style="font-size:12px;font-weight:700;min-width:42px;text-align:right;color:'+c+';">'+p.taux.toFixed(1)+'%</div>'
        +'<div style="font-size:11px;color:#94a3b8;min-width:50px;text-align:right;">'+p.real+'/'+p.total+'</div>'
        +'</div>';
    }).join('');
  }

  var cal = buildCal();
  var cardStyle='background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:20px 22px;margin-bottom:14px;';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
  +'<body style="margin:0;padding:0;background:#f8fafc;font-family:Arial,sans-serif;color:#0f172a;">'

  // Header
  +'<div style="background:linear-gradient(135deg,#1d4ed8 0%,#1e40af 100%);padding:20px 32px;">'
  +'<div style="max-width:860px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">'
  +'<div><div style="color:#fff;font-weight:700;font-size:16px;letter-spacing:-.3px;">Rapport de Maintenance</div>'
  +'<div style="color:#bfdbfe;font-size:11px;margin-top:2px;">Bureau des Méthodes Daoui · OCP Group Khouribga</div></div>'
  +'<div style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:6px;padding:6px 14px;color:#fff;font-weight:600;font-size:12px;">S'+s+' · '+d0+' → '+d1+'</div>'
  +'</div></div>'

  // Contenu
  +'<div style="max-width:860px;margin:0 auto;padding:24px 20px 48px;">'
  +'<h2 style="font-size:18px;font-weight:700;letter-spacing:-.4px;margin:0 0 4px;">Rapport Hebdomadaire de Planification</h2>'
  +'<p style="font-size:13px;color:#64748b;margin:0 0 20px;">Semaine <strong>S'+s+'</strong> · '+d0+' au '+d1+' · Généré le '+new Date().toLocaleDateString('fr-FR',{day:'2-digit',month:'long',year:'numeric'})+'</p>'

  // Calendrier
  +secLabel('Calendrier des arrêts préventifs — Semaine S'+s)
  +'<div style="'+cardStyle+'">'
  +'<div style="margin-bottom:12px;">'
  +'<div style="font-size:13px;font-weight:600;color:#0f172a;">Arrêts S'+s+' · '+arrets.rows.length+' enregistré(s)</div>'
  +'<div style="font-size:11px;color:#94a3b8;margin-top:2px;">'
  +'<span style="margin-right:10px;color:#166534;font-weight:600;">● Réalisé</span>'
  +'<span style="margin-right:10px;color:#991b1b;font-weight:600;">● Non réalisé</span>'
  +'<span style="color:#9a3412;font-weight:600;">● Imprévu</span>'
  +'</div></div>'
  +'<div style="overflow-x:auto;">'+cal+'</div>'
  +'</div>'

  // KPIs
  +secLabel('Indicateurs clés du mois — '+kpi.mois)

  // Global
  +subSection('#eff6ff','#bfdbfe','#1d4ed8','<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>','Global')
  +'<table style="width:100%;border-collapse:collapse;margin-bottom:4px;"><tr>'
  +kpiCard('#eff6ff','#1d4ed8','<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',kpi.total.toLocaleString('fr-FR'),'Total OT planifiés','Ordres de travail du mois',null)
  +kpiCard('#ecfdf5','#059669','<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',kpi.real.toLocaleString('fr-FR'),'OT Réalisés','Taux : <b>'+kpi.tauxRealStr+'</b>',kpi.tauxReal)
  +'</tr><tr>'
  +kpiCard('#fffbeb','#d97706','<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',kpi.lanc.toLocaleString('fr-FR'),'OT Lancés','En cours : <b>'+kpi.lancPct+'</b>',null)
  +kpiCard('#fef2f2','#dc2626','<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>',kpi.crpr.toLocaleString('fr-FR'),'Non lancés (CRPR)','Part : <b>'+kpi.crprPct+'</b>',null)
  +'</tr><tr>'
  +kpiCard('#f1f5f9','#475569','<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',kpi.backlog.toLocaleString('fr-FR'),'Backlog','ATPL + LANC',null)
  +kpiCard('#fffbeb','#d97706','<line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',kpi.sys+'/'+ kpi.cur,'Préventif / Correctif',kpi.sysPct+' / '+kpi.curPct,null)
  +'</tr></table>'

  // Préventif
  +subSection('#ecfdf5','#a7f3d0','#059669','<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>','Préventif')
  +'<table style="width:100%;border-collapse:collapse;margin-bottom:4px;"><tr>'
  +kpiCard('#f5f3ff','#7c3aed','<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',kpi.sys.toLocaleString('fr-FR'),'OT Préventif systématique','ZCON + ZEST + ZETL : <b>'+kpi.sysPct+'</b>',null)
  +kpiCard('#ecfdf5','#059669','<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',kpi.tauxPrevStr,'Taux réalisation préventif','CONF+TCLO+CLOT / total ZCON+ZEST+ZETL',kpi.tauxPrev)
  +'</tr></table>'

  // Correctif
  +subSection('#fef2f2','#fca5a5','#dc2626','<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>','Correctif')
  +'<table style="width:100%;border-collapse:collapse;margin-bottom:4px;"><tr>'
  +kpiCard('#fef2f2','#dc2626','<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',kpi.cur.toLocaleString('fr-FR'),'OT Correctif (ZCOR)','Part : <b>'+kpi.curPct+'</b>',null)
  +kpiCard('#fef2f2','#dc2626','<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>',kpi.tauxCorStr,'Taux réalisation correctif','CONF+TCLO+CLOT / total ZCOR',kpi.tauxCor)
  +'</tr></table>'

  // Postes
  +secLabel('Taux de réalisation par corps de métier')
  +'<div style="'+cardStyle+'">'
  +'<div style="font-size:13px;font-weight:600;margin-bottom:12px;">'+kpi.mois+'</div>'
  +buildPostes()
  +'</div>'

  // Signature
  +'<div style="margin-top:28px;padding-top:18px;border-top:1px solid #e2e8f0;">'
  +'<p style="color:#64748b;font-size:13px;margin-bottom:8px;">Cordialement,</p>'
  +'<div style="font-family:Georgia,serif;font-size:14px;color:#002060;line-height:1.6;">'
  +'<strong>Marouane ELAMRAOUI</strong><br>'
  +'<span style="color:#c55a11;">Méthode de Maintenance</span><br>'
  +'<strong>OCP SA - Khouribga</strong><br>'
  +'<span style="color:#059669;">Tél. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:#059669;">Cisco :</span> 8103388<br>'
  +'<a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>'
  +'</div></div>'

  +'</div></body></html>';
}

// ── Envoi depuis l'interface ──────────────────────────────────
function envoyerRapportDepuisInterface(p) {
  try {
    if (!p.emails) return { ok: false, msg: 'Aucun destinataire.' };
    var mo = parseInt(p.mo), yr = parseInt(p.yr);
    var arrets = rhGetArrets();
    var kpi    = rhGetKpi(mo, yr);
    var html   = rhBuildHtml(arrets, kpi);
    var MOIS   = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
    var sujet  = 'Rapport Hebdomadaire de Planification — S' + arrets.sem + ' · ' + MOIS[mo] + ' ' + yr;
    sendEmailRH(p.emails, sujet, html, 'Bureau Méthode Daoui - Planification');
    PropertiesService.getScriptProperties().setProperty('RH_EMAILS', p.emails);
    return { ok: true, msg: 'Rapport envoyé avec succès à : ' + p.emails };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// ── Planification ─────────────────────────────────────────────
function planifierRapportInterface(p) {
  try {
    if (!p.emails) return { ok: false, msg: 'Aucun destinataire.' };
    var trigger;

    if (p.frequence === 'unique') {
      var dt = new Date(p.dateHeure);
      trigger = ScriptApp.newTrigger('executerRapportPlanifie')
        .timeBased().at(dt).create();

    } else if (p.frequence === 'mensuel') {
      trigger = ScriptApp.newTrigger('executerRapportPlanifie')
        .timeBased().onMonthDay(p.jourMois || 1).atHour(p.heure || 8).create();

    } else if (p.frequence === 'hebdomadaire') {
      var DAYS = { MONDAY: ScriptApp.WeekDay.MONDAY, TUESDAY: ScriptApp.WeekDay.TUESDAY,
        WEDNESDAY: ScriptApp.WeekDay.WEDNESDAY, THURSDAY: ScriptApp.WeekDay.THURSDAY,
        FRIDAY: ScriptApp.WeekDay.FRIDAY };
      trigger = ScriptApp.newTrigger('executerRapportPlanifie')
        .timeBased().onWeekDay(DAYS[p.jourSemaine] || ScriptApp.WeekDay.MONDAY)
        .atHour(p.heure || 8).create();
    } else {
      return { ok: false, msg: 'Fréquence invalide.' };
    }

    // Sauvegarder la config du trigger
    var cfg = JSON.parse(JSON.stringify(p));
    cfg.triggerId = trigger.getUniqueId();
    PropertiesService.getScriptProperties().setProperty('PLANIF_' + cfg.triggerId, JSON.stringify(cfg));
    PropertiesService.getScriptProperties().setProperty('RH_EMAILS', p.emails);
    return { ok: true, msg: 'Planification créée avec succès.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// ── Exécution d'un rapport planifié ──────────────────────────
function executerRapportPlanifie(e) {
  var triggerId = e && e.triggerUid;
  var props = PropertiesService.getScriptProperties();
  var cfg   = triggerId ? JSON.parse(props.getProperty('PLANIF_' + triggerId) || '{}') : {};

  var mo = cfg.mo !== undefined ? parseInt(cfg.mo) : new Date().getMonth();
  var yr = cfg.yr ? parseInt(cfg.yr) : new Date().getFullYear();

  // Mode relatif (mois courant ou précédent)
  if (cfg.moRelatif !== undefined) {
    var now = new Date();
    var rel = parseInt(cfg.moRelatif) || 0;
    var d   = new Date(now.getFullYear(), now.getMonth() + rel, 1);
    mo = d.getMonth(); yr = d.getFullYear();
  }

  var emails = cfg.emails || props.getProperty('RH_EMAILS') || RH_OCP_EMAIL;
  var arrets = rhGetArrets();
  var kpi    = rhGetKpi(mo, yr);
  var html   = rhBuildHtml(arrets, kpi);
  var MOIS   = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var sujet  = 'Rapport Hebdomadaire de Planification — S' + arrets.sem + ' · ' + MOIS[mo] + ' ' + yr;

  sendEmailRH(emails, sujet, html, 'Bureau Méthode Daoui - Planification');

  // Supprimer le trigger si "unique"
  if (cfg.frequence === 'unique' && triggerId) {
    ScriptApp.getProjectTriggers().forEach(function(t){
      if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
    });
    props.deleteProperty('PLANIF_' + triggerId);
  }
}

// ── Liste des planifications ──────────────────────────────────
function obtenirPlanificationsInterface() {
  var triggers = ScriptApp.getProjectTriggers();
  var props    = PropertiesService.getScriptProperties().getProperties();
  var list     = [];
  triggers.forEach(function(t) {
    var id  = t.getUniqueId();
    var key = 'PLANIF_' + id;
    if (props[key]) {
      try {
        var cfg = JSON.parse(props[key]);
        cfg.triggerId = id;
        list.push(cfg);
      } catch(e) {}
    }
  });
  return list;
}

// ── Suppression d'une planification ──────────────────────────
function supprimerPlanificationInterface(triggerId) {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t){
      if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
    });
    PropertiesService.getScriptProperties().deleteProperty('PLANIF_' + triggerId);
    return { ok: true, msg: 'Planification supprimée.' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// ── Fonction de test ──────────────────────────────────────────
function testerRapportHebdo() {
  var now    = new Date();
  var arrets = rhGetArrets();
  var kpi    = rhGetKpi(now.getMonth(), now.getFullYear());

  Logger.log('Semaine : S' + arrets.sem + ' (' + arrets.s0 + ' → ' + arrets.s1 + ')');
  Logger.log('Arrêts : ' + arrets.rows.length);
  Logger.log('KPI : total=' + kpi.total + ' réalisé=' + kpi.real + ' taux=' + kpi.tauxRealStr);

  var html  = rhBuildHtml(arrets, kpi);
  var sujet = '[TEST] Rapport Hebdomadaire S' + arrets.sem;

  sendEmailRH(RH_OCP_EMAIL, sujet, html, 'Bureau Méthode Daoui - Planification');
  Logger.log('✅ Test envoyé à : ' + RH_OCP_EMAIL);
}
