// ═══════════════════════════════════════════════════════════════
//  RAPPORT HEBDOMADAIRE — Backend GAS
//  Sert l'interface HTML + envoie les rapports via EWS (OCP)
// ═══════════════════════════════════════════════════════════════

// ── IDs des fichiers ──────────────────────────────────────────
const RH_OT_FILE_ID     = '1aQAvb1DUv6Vk1Y1C-WEYgQnYN1BxujEAg8lbMt1sP3s';
const RH_ARRETS_FILE_ID = '1EBACM8ou8B_9fmExToUKsMCvHL27hiwU2D0yZ_gQGOA';
const RH_ARRETS_SHEET   = 'Planning des arrets';
const RH_AVIS_FILE_ID   = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';

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
  var cD=ci(['start date','date début','date debut','début','date']);
  var cI=ci(['installation','équipement','equipement','arrêt','arret']);
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
  var posteMap={}, typeMap={};

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
    if(tp) typeMap[tp]=(typeMap[tp]||0)+1;
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
    postes:postes,
    typeData:Object.keys(typeMap).map(function(k){return{type:k,count:typeMap[k]};}).sort(function(a,b){return b.count-a.count;})
  };
}

// ── Lecture des Avis (mois courant) ──────────────────────────
function rhGetAvis(mo, yr) {
  try {
    var ss = SpreadsheetApp.openById(RH_AVIS_FILE_ID);
    var sheet = null;
    ss.getSheets().forEach(function(s){
      if(s.getName().toLowerCase().indexOf('avis')>=0) sheet=s;
    });
    if(!sheet) return null;
    var data=sheet.getDataRange().getValues(), hdr=data[0].map(function(h){return h.toString().trim();});
    function ci(n){for(var i=0;i<n.length;i++){var x=hdr.indexOf(n[i]);if(x>=0)return x;}return -1;}
    var cCree =ci(['Créé le','Crée le','Date création']);
    var cOrdre=ci(['Ordre','N° ordre']);
    var cStatA=ci(['Statut ABR','Statut Abr']);
    var cPoste=ci(['Poste trav.','Poste trav','Poste travail']);
    var cInst =ci(['Installation']);
    var cSect =ci(['Secteur']);
    var total=0,avecOT=0,ouverts=0;
    var bySecteur={},byPoste={},openByPoste={},byInstall={};
    for(var i=1;i<data.length;i++){
      var r=data[i], rawD=cCree>=0?r[cCree]:null, d;
      if(rawD instanceof Date) d=rawD;
      else if(typeof rawD==='number') d=new Date(Math.round((rawD-25569)*86400000));
      else if(rawD) d=new Date(rawD); else continue;
      if(isNaN(d)||d.getFullYear()!==yr||d.getMonth()!==mo) continue;
      total++;
      var ordre=cOrdre>=0?r[cOrdre].toString().trim():'';
      var statA=cStatA>=0?r[cStatA].toString().trim():'';
      var poste=cPoste>=0?r[cPoste].toString().trim():'';
      var inst =cInst >=0?r[cInst ].toString().trim():'';
      var sect =cSect >=0?r[cSect ].toString().trim():'';
      if(ordre) avecOT++;
      var isOpen=statA==='AOUV'||statA==='AENC';
      if(isOpen) ouverts++;
      if(sect) bySecteur[sect]=(bySecteur[sect]||0)+1;
      if(poste){byPoste[poste]=(byPoste[poste]||0)+1; if(isOpen) openByPoste[poste]=(openByPoste[poste]||0)+1;}
      if(inst) byInstall[inst]=(byInstall[inst]||0)+1;
    }
    function toArr(map,lim){return Object.keys(map).map(function(k){return{label:k,count:map[k]};}).sort(function(a,b){return b.count-a.count;}).slice(0,lim||10);}
    return {total:total,avecOT:avecOT,ouverts:ouverts,
      txConv:total?parseFloat(((avecOT/total)*100).toFixed(1)):0,
      bySecteur:toArr(bySecteur,8), byPoste:toArr(byPoste,6),
      openByPoste:toArr(openByPoste,6), byInstall:toArr(byInstall,8)};
  } catch(e){ Logger.log('rhGetAvis: '+e.message); return null; }
}

// ── Générateurs de graphiques (GAS Charts → base64 PNG) ──────
function rhMakePieImg(labels, values, title, w, h) {
  try {
    var dt=Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING,'Cat')
      .addColumn(Charts.ColumnType.NUMBER,'Val');
    for(var i=0;i<labels.length;i++) dt.addRow([labels[i],values[i]]);
    dt.build();
    var c=Charts.newPieChart().setDataTable(dt).setDimensions(w||380,h||210)
      .setOption('title',title||'')
      .setOption('backgroundColor','#ffffff')
      .setOption('chartArea',{left:10,top:28,width:'62%',height:'78%'})
      .setOption('legend',{position:'right',textStyle:{fontSize:9}})
      .setOption('pieSliceTextStyle',{fontSize:9})
      .build();
    return 'data:image/png;base64,'+Utilities.base64Encode(c.getAs('image/png').getBytes());
  } catch(e){ Logger.log('rhMakePieImg: '+e.message); return ''; }
}

function rhMakeBarImg(labels, values, color, title, w, h) {
  try {
    var dt=Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING,'Item')
      .addColumn(Charts.ColumnType.NUMBER,'Count');
    for(var i=0;i<labels.length;i++) dt.addRow([labels[i],values[i]]);
    dt.build();
    var c=Charts.newBarChart().setDataTable(dt).setDimensions(w||380,h||210)
      .setOption('title',title||'')
      .setOption('backgroundColor','#ffffff')
      .setOption('colors',[color||'#3b82f6'])
      .setOption('legend',{position:'none'})
      .setOption('chartArea',{left:90,top:28,width:'60%',height:'75%'})
      .setOption('hAxis',{textStyle:{fontSize:9}})
      .setOption('vAxis',{textStyle:{fontSize:9}})
      .build();
    return 'data:image/png;base64,'+Utilities.base64Encode(c.getAs('image/png').getBytes());
  } catch(e){ Logger.log('rhMakeBarImg: '+e.message); return ''; }
}

// ── Construction HTML du rapport (style Maintenance Analytics) ─
function rhBuildHtml(arrets, kpi, avis) {
  var s=arrets.sem, d0=rhFmtDate(arrets.s0), d1=rhFmtDate(arrets.s1);
  var CS = 'background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:20px 22px 18px;box-shadow:0 1px 4px rgba(0,0,0,0.06);margin-bottom:16px;';

  // ── Calendrier ──
  function buildCal() {
    if (!arrets.rows.length) return '<p style="color:#94a3b8;font-style:italic;font-size:13px;padding:8px 0;">Aucun arr&#234;t enregistr&#233; pour S'+s+'.</p>';
    var DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];
    var wMap={};
    arrets.rows.forEach(function(r){
      var parts=r.date.split('-').map(Number);
      var d=new Date(parts[0],parts[1]-1,parts[2]);
      var wn=parseInt(r.semaine.toString().replace(/\D/g,''));
      var k=r.annee+'-'+wn;
      if(!wMap[k]) wMap[k]={annee:r.annee,semaine:r.semaine,wn:wn,days:{}};
      var dow=(d.getDay()+6)%7;
      if(!wMap[k].days[dow]) wMap[k].days[dow]=[];
      wMap[k].days[dow].push({label:r.install,statut:r.statut});
    });
    var weeks=Object.values(wMap).sort(function(a,b){return a.annee!==b.annee?a.annee-b.annee:a.wn-b.wn;});
    var TH='font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.5px;background:#f8fafc;padding:7px 8px;border:1px solid #e2e8f0;text-align:center;';
    var WH='font-size:11px;font-weight:700;color:#0f172a;background:#f8fafc;padding:8px 12px;border:1px solid #e2e8f0;white-space:nowrap;';
    var TC='padding:5px 6px;border:1px solid #e2e8f0;vertical-align:top;';
    var DT='font-size:9px;font-weight:500;color:#94a3b8;text-align:center;padding:3px 6px;background:#f8fafc;border:1px solid #e2e8f0;border-top:none;font-family:monospace;';
    var html='<table style="width:100%;border-collapse:collapse;font-size:12px;">';
    html+='<tr><th style="'+WH+'">Sem.</th>'+DAYS.map(function(d){return'<th style="'+TH+'">'+d+'</th>';}).join('')+'</tr>';
    weeks.forEach(function(w,wi){
      var mon=rhMondayOf(w.annee,w.wn);
      var sep=wi>0?'border-top:2px solid #cbd5e1;':'';
      html+='<tr><td style="'+WH+sep+'font-weight:700;color:#1d4ed8;">'+w.semaine+'</td>';
      for(var di=0;di<7;di++){
        var dd=new Date(mon); dd.setDate(dd.getDate()+di);
        html+='<td style="'+DT+sep+'">'+String(dd.getDate()).padStart(2,'0')+'/'+String(dd.getMonth()+1).padStart(2,'0')+'</td>';
      }
      html+='</tr>';
      html+='<tr><td style="font-family:monospace;font-size:9px;font-weight:700;color:#94a3b8;text-align:center;vertical-align:middle;'+TC+'">'+w.annee+'</td>';
      for(var di=0;di<7;di++){
        var items=w.days[di]||[];
        html+='<td style="'+TC+'">';
        if(!items.length) html+='<span style="color:#cbd5e1;font-size:12px;">&#183;</span>';
        else items.forEach(function(it){
          var bg=it.statut==='realise'?'#dcfce7':it.statut==='imprevu'?'#ffedd5':'#fee2e2';
          var fg=it.statut==='realise'?'#166534':it.statut==='imprevu'?'#9a3412':'#991b1b';
          html+='<span style="display:inline-block;font-size:10px;font-weight:600;padding:3px 7px;border-radius:3px;margin:2px 2px 3px 0;white-space:nowrap;background:'+bg+';color:'+fg+';">'+it.label+'</span>';
        });
        html+='</td>';
      }
      html+='</tr>';
    });
    return html+'</table>';
  }

  // ── KPI card ──
  function kpiCard(iconBg,iconFg,svg,val,label,sub,taux) {
    var tag='';
    if(taux!==null&&taux!==undefined){
      var tc=taux>=80?{bg:'#dcfce7',fg:'#166534'}:taux>=50?{bg:'#fef9c3',fg:'#854d0e'}:{bg:'#fee2e2',fg:'#991b1b'};
      tag='<span style="font-size:11px;font-weight:600;border-radius:4px;padding:2px 8px;background:'+tc.bg+';color:'+tc.fg+';">'+taux.toFixed(1)+'%</span>';
    }
    return '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:18px 20px 16px;box-shadow:0 1px 4px rgba(0,0,0,0.06);min-width:140px;flex:1;">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px;">'
      +'<div style="width:36px;height:36px;border-radius:8px;background:'+iconBg+';display:flex;align-items:center;justify-content:center;flex-shrink:0;">'
      +'<svg viewBox="0 0 24 24" style="width:16px;height:16px;stroke:'+iconFg+';fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;">'+svg+'</svg>'
      +'</div>'+tag+'</div>'
      +'<div style="font-size:28px;font-weight:700;letter-spacing:-1.5px;line-height:1;margin-bottom:4px;color:#0f172a;">'+val+'</div>'
      +'<div style="font-size:12px;font-weight:500;color:#475569;">'+label+'</div>'
      +'<div style="font-size:11px;color:#94a3b8;margin-top:10px;padding-top:10px;border-top:1px solid #f1f5f9;">'+sub+'</div>'
      +'</div>';
  }

  // ── Section label ──
  function secLabel(txt) {
    return '<div style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;margin:28px 0 12px;display:flex;align-items:center;gap:10px;">'
      +txt+'<span style="flex:1;height:1px;background:#e2e8f0;display:inline-block;margin-left:8px;"></span></div>';
  }

  // ── Sub-section label ──
  function subSection(bg,border,color,svg,label) {
    return '<div style="font-size:10px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;padding:5px 12px;border-radius:5px;display:inline-flex;align-items:center;gap:7px;margin:16px 0 10px;background:'+bg+';color:'+color+';border:1px solid '+border+';">'
      +'<svg viewBox="0 0 24 24" style="width:12px;height:12px;stroke:'+color+';fill:none;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round;">'+svg+'</svg>'
      +label+'</div>';
  }

  // ── Chart card ──
  function chartCard(imgSrc,title) {
    if(!imgSrc) return '<div style="flex:1;min-width:200px;"></div>';
    return '<div style="flex:1;min-width:200px;background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,0.06);">'
      +'<div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#64748b;margin-bottom:8px;">'+title+'</div>'
      +'<img src="'+imgSrc+'" style="width:100%;display:block;border:0;" alt="'+title+'">'
      +'</div>';
  }

  // ── Postes ──
  function buildPostes() {
    if(!kpi.postes.length) return '<p style="color:#94a3b8;font-size:12px;font-style:italic;">Aucune donn&#233;e.</p>';
    var html='';
    kpi.postes.forEach(function(p){
      var c=p.taux>=80?'#059669':p.taux>=50?'#d97706':'#dc2626';
      var w=Math.min(100,p.taux).toFixed(0);
      html+='<div style="display:flex;align-items:center;gap:12px;padding:9px 0;border-bottom:1px solid #f1f5f9;">'
        +'<div style="font-size:12px;font-weight:600;min-width:130px;color:#0f172a;white-space:nowrap;">'+p.nom+'</div>'
        +'<div style="flex:1;height:10px;background:#f1f5f9;border-radius:5px;overflow:hidden;border:1px solid #e2e8f0;">'
        +'<div style="height:100%;border-radius:5px;width:'+w+'%;background:'+c+';"></div></div>'
        +'<div style="font-family:monospace;font-size:12px;font-weight:600;min-width:44px;text-align:right;color:'+c+';">'+p.taux.toFixed(1)+'%</div>'
        +'<div style="font-size:11px;color:#94a3b8;min-width:54px;text-align:right;">'+p.real+'/'+p.total+'</div>'
        +'</div>';
    });
    return html;
  }

  // ── Avis non-clôturés : tableau détail ──
  function buildAvisNonClos() {
    if(!avis||!avis.openByPoste.length) return '';
    var total=avis.openByPoste.reduce(function(s,r){return s+r.count;},0);
    var html='<div style="margin-top:16px;">'
      +'<div style="font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#dc2626;margin-bottom:8px;">D&#233;tail des avis non cl&#244;tur&#233;s &#8212; AOUV + AENC ('+avis.ouverts+' avis)</div>'
      +'<table style="width:100%;border-collapse:collapse;font-size:12px;">'
      +'<tr><th style="text-align:left;padding:7px 10px;border-bottom:2px solid #e2e8f0;color:#475569;font-size:11px;">Corps de M&#233;tier</th>'
      +'<th style="text-align:right;padding:7px 10px;border-bottom:2px solid #e2e8f0;color:#dc2626;font-size:11px;">Total</th>'
      +'<th style="text-align:right;padding:7px 10px;border-bottom:2px solid #e2e8f0;color:#475569;font-size:11px;">Part</th></tr>';
    avis.openByPoste.forEach(function(r){
      var part=total?((r.count/total)*100).toFixed(1)+'%':'—';
      html+='<tr><td style="padding:8px 10px;border-bottom:1px solid #f1f5f9;font-weight:600;color:#0f172a;">'+r.label+'</td>'
        +'<td style="padding:8px 10px;border-bottom:1px solid #f1f5f9;text-align:right;font-weight:700;color:#dc2626;">'+r.count+'</td>'
        +'<td style="padding:8px 10px;border-bottom:1px solid #f1f5f9;text-align:right;color:#94a3b8;">'+part+'</td></tr>';
    });
    html+='<tr><td style="padding:8px 10px;font-weight:700;color:#0f172a;">TOTAL</td>'
      +'<td style="padding:8px 10px;text-align:right;font-weight:700;color:#dc2626;">'+total+'</td>'
      +'<td style="padding:8px 10px;text-align:right;color:#94a3b8;">100%</td></tr>'
      +'</table></div>';
    return html;
  }

  // ── SVG paths ──
  var SVG = {
    copy:    '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
    check:   '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
    clock:   '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
    alert:   '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>',
    calendar:'<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
    shield:  '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
    wrench:  '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
    layers:  '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>',
    target:  '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
    users:   '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    arrow:   '<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>',
    grid:    '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/>'
  };

  // ── Compute synthèse ──
  var nbA=arrets.rows.length;
  var nbR=arrets.rows.filter(function(r){return r.statut==='realise';}).length;
  var tA=nbA?parseFloat(((nbR/nbA)*100).toFixed(1)):0;
  var tAStr=nbA?tA.toFixed(1)+'%':'—';
  var scTot=kpi.sys+kpi.cur;
  var sysPct2=scTot?parseFloat(((kpi.sys/scTot)*100).toFixed(1)):0;
  var curPct2=scTot?parseFloat(((kpi.cur/scTot)*100).toFixed(1)):0;

  // ── Compute charts ──
  var imgType =(kpi.typeData.length)?rhMakePieImg(kpi.typeData.slice(0,6).map(function(x){return x.type;}),kpi.typeData.slice(0,6).map(function(x){return x.count;}),'R\u00e9partition par type d\'OT'):'';
  var imgPoste=(kpi.postes.length)?rhMakeBarImg(kpi.postes.slice(0,7).map(function(x){return x.nom;}),kpi.postes.slice(0,7).map(function(x){return x.total;}),'#3b82f6','Volume OT par corps de m\u00e9tier'):'';
  var imgSect =(avis&&avis.bySecteur.length)?rhMakePieImg(avis.bySecteur.map(function(x){return x.label;}),avis.bySecteur.map(function(x){return x.count;}),'R\u00e9partition par secteur'):'';
  var imgAPoste=(avis&&avis.byPoste.length)?rhMakePieImg(avis.byPoste.map(function(x){return x.label;}),avis.byPoste.map(function(x){return x.count;}),'Avis par corps de m\u00e9tier'):'';
  var imgInst=(avis&&avis.byInstall.length)?rhMakeBarImg(avis.byInstall.slice(0,8).map(function(x){return x.label;}),avis.byInstall.slice(0,8).map(function(x){return x.count;}),'#0891b2','Avis par installation'):'';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
  +'<body style="margin:0;padding:0;background:#f8fafc;font-family:Inter,Arial,sans-serif;color:#0f172a;">'

  // ── HEADER ──
  +'<div style="background:#1d4ed8;padding:20px 32px;display:flex;align-items:center;gap:12px;">'
  +'<div style="width:36px;height:36px;background:rgba(255,255,255,0.2);border-radius:6px;display:inline-flex;align-items:center;justify-content:center;">'
  +'<svg viewBox="0 0 24 24" style="width:18px;height:18px;fill:none;stroke:white;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;">'+SVG.grid+'</svg>'
  +'</div>'
  +'<div style="display:inline-block;margin-left:10px;">'
  +'<div style="color:white;font-weight:700;font-size:15px;letter-spacing:-0.3px;">Maintenance Analytics</div>'
  +'<div style="color:rgba(255,255,255,0.7);font-size:11px;">Bureau M&#233;thode Daoui &middot; OCP SA Khouribga</div>'
  +'</div>'
  +'<div style="margin-left:auto;background:rgba(255,255,255,0.15);border-radius:6px;padding:5px 14px;color:white;font-weight:700;font-size:13px;display:inline-block;">'
  +'S'+s+' &middot; '+d0+' &#8594; '+d1
  +'</div>'
  +'</div>'

  // ── CONTENU ──
  +'<div style="padding:28px 32px;max-width:900px;margin:0 auto;">'
  +'<h2 style="font-size:20px;font-weight:700;letter-spacing:-0.4px;margin:0 0 4px;color:#0f172a;">Rapport Hebdomadaire de Planification</h2>'
  +'<p style="font-size:13px;color:#64748b;margin:0 0 24px;">Semaine <strong>S'+s+'</strong> &middot; '+d0+' au '+d1+' &middot; G&#233;n&#233;r&#233; le '+new Date().toLocaleDateString('fr-FR',{day:'2-digit',month:'long',year:'numeric'})+'</p>'

  // ══ CALENDRIER ══
  +secLabel('Calendrier des arr&#234;ts pr&#233;ventifs &#8212; Semaine S'+s)
  +'<div style="'+CS+'">'
  +'<div style="margin-bottom:14px;">'
  +'<div style="font-size:13px;font-weight:600;color:#0f172a;">Arr&#234;ts S'+s+' &middot; '+arrets.rows.length+' enregistr&#233;(s) &middot; '+d0+' &#8594; '+d1+'</div>'
  +'<div style="font-size:11px;color:#64748b;margin-top:4px;">'
  +'<span style="margin-right:14px;">&#9679; <span style="color:#166534;font-weight:600;">R&#233;alis&#233;</span></span>'
  +'<span style="margin-right:14px;">&#9679; <span style="color:#991b1b;font-weight:600;">Non r&#233;alis&#233;</span></span>'
  +'<span>&#9679; <span style="color:#9a3412;font-weight:600;">Impr&#233;vu</span></span>'
  +'</div></div>'
  +'<div style="overflow-x:auto;">'+buildCal()+'</div>'
  +'</div>'

  // ══ KPI SYNTHÈSE (3 cards) ══
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;">'
  +kpiCard('#ecfdf5','#059669',SVG.check,kpi.tauxRealStr,'Taux r&#233;alisation OT','Mois courant &middot; <b>'+kpi.real.toLocaleString('fr-FR')+'</b> / <b>'+kpi.total.toLocaleString('fr-FR')+'</b>',kpi.tauxReal)
  +kpiCard('#eff6ff','#1d4ed8',SVG.target,tAStr,'Taux r&#233;alisation arr&#234;ts','S'+s+' &middot; <b>'+nbR+'</b> r&#233;alis&#233;(s) / <b>'+nbA+'</b>',tA)
  +kpiCard('#f5f3ff','#7c3aed',SVG.layers,sysPct2.toFixed(1)+'% / '+curPct2.toFixed(1)+'%','Syst&#233;matique / Curatif','ZCON+ZEST+ZETL&nbsp;<b>'+kpi.sys+'</b> &middot; ZCOR&nbsp;<b>'+kpi.cur+'</b>',null)
  +'</div>'

  // ══ GRAPHIQUES OT ══
  +secLabel('R&#233;partitions OT')
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;">'
  +chartCard(imgType,'R&#233;partition par type d\'OT')
  +chartCard(imgPoste,'Volume OT par corps de m&#233;tier')
  +'</div>'

  // ══ KPIs DÉTAILLÉS ══
  +secLabel('Indicateurs cl&#233;s du mois &#8212; '+kpi.mois)

  // Global
  +subSection('#eff6ff','#bfdbfe','#1d4ed8',SVG.copy,'Global')
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  +kpiCard('#eff6ff','#1d4ed8',SVG.copy,kpi.total.toLocaleString('fr-FR'),'Total OT planifi&#233;s','Ordres de travail du mois',null)
  +kpiCard('#ecfdf5','#059669',SVG.check,kpi.real.toLocaleString('fr-FR'),'OT R&#233;alis&#233;s','Taux : <b>'+kpi.tauxRealStr+'</b>',kpi.tauxReal)
  +kpiCard('#fffbeb','#d97706',SVG.clock,kpi.lanc.toLocaleString('fr-FR'),'OT Lanc&#233;s','En cours : <b>'+kpi.lancPct+'</b>',null)
  +'</div>'
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  +kpiCard('#fef2f2','#dc2626',SVG.alert,kpi.crpr.toLocaleString('fr-FR'),'Non lanc&#233;s (CRPR)','Part du total : <b>'+kpi.crprPct+'</b>',null)
  +kpiCard('#f1f5f9','#475569',SVG.calendar,kpi.backlog.toLocaleString('fr-FR'),'Backlog','ATPL + LANC',null)
  +kpiCard('#fffbeb','#d97706',SVG.layers,kpi.sys+' / '+kpi.cur,'Pr&#233;ventif / Correctif',kpi.sysPct+' syst. &middot; '+kpi.curPct+' curatif',null)
  +'</div>'

  // Préventif
  +subSection('#ecfdf5','#a7f3d0','#059669',SVG.shield,'Pr&#233;ventif')
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  +kpiCard('#f5f3ff','#7c3aed',SVG.shield,kpi.sys.toLocaleString('fr-FR'),'OT Pr&#233;ventif syst&#233;matique','ZCON + ZEST + ZETL : <b>'+kpi.sysPct+'</b>',null)
  +kpiCard('#ecfdf5','#059669',SVG.check,kpi.tauxPrevStr,'Taux r&#233;alisation pr&#233;ventif','CONF+TCLO+CLOT / total ZCON+ZEST+ZETL',kpi.tauxPrev)
  +'</div>'

  // Correctif
  +subSection('#fef2f2','#fca5a5','#dc2626',SVG.wrench,'Correctif')
  +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:6px;">'
  +kpiCard('#fef2f2','#dc2626',SVG.wrench,kpi.cur.toLocaleString('fr-FR'),'OT Correctif (ZCOR)','Part du total : <b>'+kpi.curPct+'</b>',null)
  +kpiCard('#fef2f2','#dc2626',SVG.check,kpi.tauxCorStr,'Taux r&#233;alisation correctif','CONF+TCLO+CLOT / total ZCOR',kpi.tauxCor)
  +'</div>'

  // ══ POSTES ══
  +secLabel('Taux de r&#233;alisation par corps de m&#233;tier')
  +'<div style="'+CS+'">'
  +'<div style="font-size:13px;font-weight:600;color:#0f172a;margin-bottom:4px;">Taux de r&#233;alisation &#8212; '+kpi.mois+'</div>'
  +'<div style="font-size:11px;color:#94a3b8;margin-bottom:14px;">CONF + TCLO + CLOT / total &middot; '
  +'<span style="color:#059669;font-weight:600;">&#8805; 80% Bon</span> &middot; '
  +'<span style="color:#d97706;font-weight:600;">50&#8211;79% Moyen</span> &middot; '
  +'<span style="color:#dc2626;font-weight:600;">&lt;50% Faible</span></div>'
  +buildPostes()
  +'</div>'

  // ══ SECTION AVIS ══
  +(avis ? (
    secLabel('Analyse des Avis (ZC) &#8212; '+kpi.mois)
    +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:16px;">'
    +kpiCard('#fdf4ff','#7c3aed',SVG.users,avis.total.toLocaleString('fr-FR'),'Total Avis','Avis de type ZC',null)
    +kpiCard('#ecfdf5','#059669',SVG.arrow,avis.avecOT.toLocaleString('fr-FR'),'Convertis en OT','Taux : <b>'+avis.txConv.toFixed(1)+'%</b>',avis.txConv)
    +kpiCard('#fef2f2','#dc2626',SVG.alert,avis.ouverts.toLocaleString('fr-FR'),'Avis Ouverts','AOUV + AENC',null)
    +'</div>'
    +subSection('#fdf4ff','#e9d5ff','#7c3aed',SVG.grid,'R&#233;partitions')
    +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:12px;">'
    +chartCard(imgSect,'R&#233;partition par secteur')
    +chartCard(imgAPoste,'Avis par corps de m&#233;tier')
    +'</div>'
    +'<div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:12px;">'
    +chartCard(imgInst,'Avis par installation')
    +'</div>'
    +'<div style="'+CS+'margin-bottom:24px;">'
    +buildAvisNonClos()
    +'</div>'
  ) : '')

  // ── SIGNATURE ──
  +'<div style="margin-top:32px;padding-top:20px;border-top:1px solid #e2e8f0;">'
  +'<p style="margin:0 0 10px;color:#64748b;font-size:13px;">Cordialement,</p>'
  +'<div style="font-family:Georgia,serif;font-size:14px;color:#002060;line-height:1.6;">'
  +'<strong>Marouane ELAMRAOUI</strong><br>'
  +'<span style="color:#c55a11;">M&#233;thode de Maintenance</span><br>'
  +'<strong>OCP SA - Khouribga</strong><br>'
  +'<span style="color:#059669;">T&#233;l. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:#059669;">Cisco :</span> 8103388<br>'
  +'<a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>'
  +'</div></div>'

  +'</div>' // fin contenu
  +'</body></html>';
}

// ── Envoi depuis l'interface ──────────────────────────────────
function envoyerRapportDepuisInterface(p) {
  try {
    if (!p.emails) return { ok: false, msg: 'Aucun destinataire.' };
    var mo = parseInt(p.mo), yr = parseInt(p.yr);
    var arrets = rhGetArrets();
    var kpi    = rhGetKpi(mo, yr);
    var avis   = rhGetAvis(mo, yr);
    var html   = rhBuildHtml(arrets, kpi, avis);
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
  var avis   = rhGetAvis(mo, yr);
  var html   = rhBuildHtml(arrets, kpi, avis);
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
  var mo     = now.getMonth(), yr = now.getFullYear();
  var arrets = rhGetArrets();
  var kpi    = rhGetKpi(mo, yr);
  var avis   = rhGetAvis(mo, yr);

  Logger.log('Semaine : S' + arrets.sem + ' (' + arrets.s0 + ' → ' + arrets.s1 + ')');
  Logger.log('Arrêts : ' + arrets.rows.length);
  Logger.log('KPI : total=' + kpi.total + ' réalisé=' + kpi.real + ' taux=' + kpi.tauxRealStr);
  Logger.log('Avis : ' + (avis ? avis.total + ' total / ' + avis.ouverts + ' ouverts' : 'non disponibles'));

  var html  = rhBuildHtml(arrets, kpi, avis);
  var sujet = '[TEST] Rapport Hebdomadaire S' + arrets.sem;

  sendEmailRH(RH_OCP_EMAIL, sujet, html, 'Bureau Méthode Daoui - Planification');
  Logger.log('✅ Test envoyé à : ' + RH_OCP_EMAIL);
}
