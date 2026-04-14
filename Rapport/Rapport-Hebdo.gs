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

function sendEmailRH(to, subject, htmlBody, senderName, attachments) {
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
    bodyB64
  ];

  // Pièces jointes optionnelles
  if (attachments && attachments.length) {
    attachments.forEach(function(att) {
      if (!att) return;
      var attB64  = Utilities.base64Encode(att.getBytes());
      var nameB64 = Utilities.base64Encode(att.getName(), Utilities.Charset.UTF_8);
      mimeParts.push('');
      mimeParts.push('--' + boundary);
      mimeParts.push('Content-Type: ' + (att.getContentType() || 'application/octet-stream'));
      mimeParts.push('Content-Transfer-Encoding: base64');
      mimeParts.push('Content-Disposition: attachment; filename="=?UTF-8?B?' + nameB64 + '?="');
      mimeParts.push('');
      mimeParts.push(attB64);
    });
  }

  mimeParts.push('');
  mimeParts.push('--' + boundary + '--');

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

// ── Lecture des arrêts (toutes les semaines du mois courant) ──
function rhGetArrets() {
  var today = new Date();
  var dow   = today.getDay() || 7;
  var mon   = new Date(today); mon.setDate(today.getDate() - (dow - 1));
  var monP  = new Date(mon);   monP.setDate(mon.getDate() - 7);
  var sunP  = new Date(monP);  sunP.setDate(monP.getDate() + 6);
  var s1    = rhDateStr(sunP);
  var sem   = rhWeekNum(monP);
  // s0 = lundi de la 1re semaine qui contient le 1er du mois courant
  var firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var dowFirst = firstOfMonth.getDay() || 7;
  var monFirst = new Date(firstOfMonth); monFirst.setDate(firstOfMonth.getDate() - (dowFirst - 1));
  var s0 = rhDateStr(monFirst);

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
  var total=0,real=0,lanc=0,crpr=0,sys=0,cur=0,sysR=0,curR=0,backlog=0,caract=0,nonCaract=0;
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
    if(su!=='SOPL'){if(su.length===9)caract++;else if(su.length===4)nonCaract++;}
    var isSys=['ZCON','ZEST','ZETL'].indexOf(tp)>=0, isCur=tp==='ZCOR';
    if(isSys){sys++;if(isR)sysR++;}
    if(isCur){cur++;if(isR)curR++;}
    if(tp) typeMap[tp]=(typeMap[tp]||0)+1;
    if(pt){if(!posteMap[pt])posteMap[pt]={total:0,real:0};posteMap[pt].total++;if(isR)posteMap[pt].real++;}
  }

  function p(n,t){return t?parseFloat(((n/t)*100).toFixed(1)):0;}
  function ps(n,t){return t?p(n,t).toFixed(1)+'%':'—';}

  var EXCL_POSTES=['421-GRAI','425-INCD'];
  var postes=Object.keys(posteMap).filter(function(k){return EXCL_POSTES.indexOf(k)<0;}).map(function(k){return{nom:k,total:posteMap[k].total,real:posteMap[k].real,taux:p(posteMap[k].real,posteMap[k].total)};}).sort(function(a,b){return b.total-a.total;}).slice(0,10);

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
    caract:caract, nonCaract:nonCaract,
    tauxCaract:p(caract,caract+nonCaract), tauxCaractStr:ps(caract,caract+nonCaract),
    pdrTotal:0, pdrConf:0,
    tauxPdrConf:0, tauxPdrConfStr:'—',
    otAttente:0, tempsMoyen:null, tempsMoyenStr:'—',
    typeData:Object.keys(typeMap).map(function(k){return{type:k,count:typeMap[k]};}).sort(function(a,b){return b.count-a.count;})
  };
}

// ── Lecture Préparation PDR (feuille Travaux hebdomadaire) ───
function rhGetPreparation(mo, yr) {
  try {
    var ss    = SpreadsheetApp.openById(RH_AVIS_FILE_ID);
    var sheet = ss.getSheetByName('Travaux hebdomadaire');
    if (!sheet) { Logger.log('rhGetPreparation: feuille "Travaux hebdomadaire" introuvable'); return null; }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    var hdr = data[0];
    function ci(names, fallback) {
      for (var i=0;i<names.length;i++)
        for (var j=0;j<hdr.length;j++)
          if (hdr[j].toString().trim().toLowerCase()===names[i].toLowerCase()) return j;
      return fallback !== undefined ? fallback : -1;
    }
    function ciContains(names, fallback) {
      for (var i=0;i<names.length;i++)
        for (var j=0;j<hdr.length;j++)
          if (hdr[j].toString().trim().toLowerCase().indexOf(names[i].toLowerCase())>=0) return j;
      return fallback !== undefined ? fallback : -1;
    }
    var cDate  = ciContains(['début au plus tôt','debut au plus tot','date début','date debut','debut','date'], 6); // col G
    var cUtil  = ciContains(['statut utilis'], 10);   // col K
    var cPdr   = 18;                                  // col S
    var cCreat = ciContains(['créé le','cree le','date création','date creation','cree'], 12); // col M
    var cSys   = ciContains(['statut système','statut systeme','statut sys'], 7);       // col H (détecté via web app)

    var pdrTotal=0, pdrConf=0, otAttente=0, tempsSomme=0, tempsCount=0;
    var today = new Date();
    for (var i=1;i<data.length;i++) {
      var r=data[i];
      var su=(r[cUtil]||'').toString().trim();
      var sysStat=(r[cSys]||'').toString().trim().toLowerCase();

      // ── Filtre mois/année : pdrTotal / pdrConf ──
      var inMonth = false;
      if (cDate>=0) {
        var rawD=r[cDate], d;
        if (rawD instanceof Date) d=rawD;
        else if (typeof rawD==='number') d=new Date(Math.round((rawD-25569)*86400000));
        else if (rawD) d=new Date(rawD);
        else d=null;
        if (d && !isNaN(d) && d.getFullYear()===yr && d.getMonth()===mo) inMonth=true;
      }
      if (inMonth) {
        var pdrVal=r.length>cPdr?r[cPdr].toString().trim():'';
        if (pdrVal!=='') pdrTotal++;
        if (su==='CRPR ATPD'||su==='CRPR AVPD') pdrConf++;
      }

      // ── Sans filtre date : OT en attente ──
      if (su==='CRPR ATPD'||su==='CRPR AVPD') {
        otAttente++;
        // ── Temps moyen : filtre statut système = 'créé' ──
        if (sysStat.includes('créé')) {
          var rawC=r[cCreat], dc;
          if (rawC instanceof Date) dc=rawC;
          else if (typeof rawC==='number') dc=new Date(Math.round((rawC-25569)*86400000));
          else if (rawC) dc=new Date(rawC);
          else dc=null;
          if (dc && !isNaN(dc)) {
            var jours=Math.round((today-dc)/(1000*60*60*24));
            if (jours>=0) { tempsSomme+=jours; tempsCount++; }
          }
        }
      }
    }

    function p(n,t){return t?parseFloat(((n/t)*100).toFixed(1)):0;}
    function ps(n,t){return t?p(n,t).toFixed(1)+'%':'—';}
    var tempsMoyen = tempsCount>0 ? Math.round(tempsSomme/tempsCount) : null;
    var tempsMoyenStr = tempsMoyen!==null ? tempsMoyen+' j' : '—';
    Logger.log('rhGetPreparation: pdrTotal='+pdrTotal+' pdrConf='+pdrConf+' otAttente='+otAttente+' tempsMoyen='+tempsMoyenStr+' ('+tempsCount+' OTs)');
    return {
      pdrTotal:pdrTotal, pdrConf:pdrConf, tauxPdrConf:p(pdrConf,pdrTotal), tauxPdrConfStr:ps(pdrConf,pdrTotal),
      otAttente:otAttente, tempsMoyen:tempsMoyen, tempsMoyenStr:tempsMoyenStr, tempsCount:tempsCount
    };
  } catch(e) { Logger.log('rhGetPreparation error: '+e.message); return null; }
}

// ── Lecture des Avis (mois courant) ──────────────────────────
function rhGetAvis(mo, yr) {
  try {
    var ss = SpreadsheetApp.openById(RH_AVIS_FILE_ID);
    // Recherche exacte d'abord, puis fuzzy
    var sheet = ss.getSheetByName('Avis') || ss.getSheetByName('avis') || ss.getSheetByName('AVIS');
    if (!sheet) {
      ss.getSheets().forEach(function(s){
        if (s.getName().toLowerCase().indexOf('avis') >= 0) sheet = s;
      });
    }
    if (!sheet) { Logger.log('rhGetAvis: feuille Avis introuvable'); return null; }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) { Logger.log('rhGetAvis: feuille vide'); return null; }

    // Normalisation : minuscule + suppression des accents + trim
    function norm(s) {
      return (s||'').toString().toLowerCase()
        .replace(/[éèêë]/g,'e').replace(/[àâä]/g,'a').replace(/[ùûü]/g,'u')
        .replace(/[îï]/g,'i').replace(/[ôö]/g,'o').replace(/ç/g,'c')
        .replace(/°/g,'').replace(/\./g,'').replace(/\s+/g,' ').trim();
    }

    var hdr  = data[0].map(function(h){ return h ? h.toString().trim() : ''; });
    var hdrN = hdr.map(norm);
    Logger.log('rhGetAvis headers: ' + JSON.stringify(hdr));

    // Cherche la meilleure colonne : exact normalisé, puis contient
    function ci(names) {
      var nn = names.map(norm);
      for (var j = 0; j < hdrN.length; j++)
        for (var i = 0; i < nn.length; i++)
          if (hdrN[j] === nn[i]) return j;
      for (var j = 0; j < hdrN.length; j++)
        for (var i = 0; i < nn.length; i++)
          if (hdrN[j].indexOf(nn[i]) >= 0 || nn[i].indexOf(hdrN[j]) >= 0) return j;
      return -1;
    }

    var cCree  = ci(['cree le','cre le','date creation','date de creation','creee le','date creat','date']);
    var cOrdre = ci(['ordre de travail','n ordre','no ordre','numero ordre','order','ordre']);
    var cStatA = ci(['statut abr','stat abr','statut']);
    var cPoste = ci(['poste trav','poste de travail','corps de metier','poste']);
    var cInst  = ci(['installation']);
    var cSect  = ci(['secteur']);

    Logger.log('rhGetAvis cols → cree='+cCree+' ordre='+cOrdre+' statut='+cStatA+' poste='+cPoste+' install='+cInst+' sect='+cSect);

    var total=0, avecOT=0, ouverts=0;
    var bySecteur={}, byPoste={}, openByPoste={}, byInstall={};

    for (var i = 1; i < data.length; i++) {
      var r = data[i];

      // Filtre par date (si colonne trouvée)
      if (cCree >= 0) {
        var rawD = r[cCree], d;
        if (rawD instanceof Date)                     d = rawD;
        else if (typeof rawD === 'number' && rawD > 1000)
                                                      d = new Date(Math.round((rawD - 25569) * 86400000));
        else if (rawD && typeof rawD === 'string' && rawD.trim())
                                                      d = new Date(rawD);
        else continue;
        if (!d || isNaN(d.getTime())) continue;
        if (d.getFullYear() !== parseInt(yr) || d.getMonth() !== parseInt(mo)) continue;
      }

      // Ignorer les lignes vides
      if (!r.some(function(v){ return v !== null && v !== undefined && v !== ''; })) continue;

      total++;
      var ordre = cOrdre >= 0 ? r[cOrdre].toString().trim() : '';
      var statA = cStatA >= 0 ? r[cStatA].toString().trim().toUpperCase() : '';
      var poste = cPoste >= 0 ? r[cPoste].toString().trim() : '';
      var inst  = cInst  >= 0 ? r[cInst ].toString().trim() : '';
      var sect  = cSect  >= 0 ? r[cSect ].toString().trim() : '';

      if (ordre) avecOT++;
      var isOpen = statA === 'AOUV' || statA === 'AENC';
      if (isOpen) ouverts++;
      if (sect)  bySecteur[sect] = (bySecteur[sect] || 0) + 1;
      if (poste) { byPoste[poste] = (byPoste[poste] || 0) + 1; if (isOpen) openByPoste[poste] = (openByPoste[poste] || 0) + 1; }
      if (inst)  byInstall[inst]  = (byInstall[inst]  || 0) + 1;
    }

    Logger.log('rhGetAvis résultat: total='+total+' avecOT='+avecOT+' ouverts='+ouverts);

    function toArr(map, lim) {
      return Object.keys(map).map(function(k){ return {label:k, count:map[k]}; })
        .sort(function(a,b){ return b.count - a.count; }).slice(0, lim || 10);
    }
    return {
      total:total, avecOT:avecOT, ouverts:ouverts,
      txConv: total ? parseFloat(((avecOT/total)*100).toFixed(1)) : 0,
      bySecteur:toArr(bySecteur,8), byPoste:toArr(byPoste,6),
      openByPoste:toArr(openByPoste,6), byInstall:toArr(byInstall,8)
    };
  } catch(e) { Logger.log('rhGetAvis error: '+e.message); return null; }
}

// ── Générateurs d'images via QuickChart.io (URL externe — compatible Outlook/OWA) ─────
function rhChartFetch_(chartStr, w, h) {
  var payload = JSON.stringify({
    chart: chartStr, width: w, height: h,
    backgroundColor: 'white', format: 'png', devicePixelRatio: 3
  });
  // /chart/create retourne une URL permanente au lieu de base64 (meilleure compatibilité email)
  var resp = UrlFetchApp.fetch('https://quickchart.io/chart/create', {
    method: 'post', contentType: 'application/json',
    payload: payload, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200)
    throw new Error('QuickChart create ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 200));
  var result = JSON.parse(resp.getContentText());
  if (!result.success || !result.url) throw new Error('QuickChart: pas d\'URL dans la réponse');
  return result.url;
}

function rhMakePieImg(labels, values, title, w, h) {
  try {
    var colors = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#f97316','#84cc16','#ec4899','#14b8a6'];
    var chart = '{'
      + 'type:"doughnut",'
      + 'data:{labels:' + JSON.stringify(labels)
        + ',datasets:[{data:' + JSON.stringify(values)
          + ',backgroundColor:' + JSON.stringify(colors.slice(0, Math.min(labels.length, colors.length)))
          + ',borderWidth:3,borderColor:"#ffffff",hoverBorderWidth:4}]},'
      + 'options:{'
        + 'cutoutPercentage:50,'
        + 'title:{display:' + (title ? 'true' : 'false') + ',text:' + JSON.stringify(title || '') + ',fontStyle:"bold",fontSize:15,fontColor:"#0f172a",padding:16},'
        + 'legend:{position:"right",labels:{fontSize:13,fontColor:"#334155",padding:16,usePointStyle:true}},'
        + 'plugins:{datalabels:{'
          + 'formatter:function(v,ctx){'
            + 'var s=ctx.dataset.data.reduce(function(a,b){return a+b;},0);'
            + 'return s>0&&(v/s)>0.04?Math.round(v/s*100)+"%":"";'
          + '},'
          + 'color:"#fff",font:{size:13,weight:"bold"},'
          + 'textShadowBlur:4,textShadowColor:"rgba(0,0,0,0.35)"'
        + '}}'
      + '}'
    + '}';
    return rhChartFetch_(chart, w || 780, h || 370);
  } catch(e) { Logger.log('rhMakePieImg: ' + e.message); return ''; }
}

function rhMakeBarImg(labels, values, color, title, w, h) {
  try {
    var chart = '{'
      + 'type:"horizontalBar",'
      + 'data:{labels:' + JSON.stringify(labels)
        + ',datasets:[{data:' + JSON.stringify(values)
          + ',backgroundColor:' + JSON.stringify(color || '#3b82f6')
          + ',borderWidth:0}]},'
      + 'options:{'
        + 'title:{display:' + (title ? 'true' : 'false') + ',text:' + JSON.stringify(title || '') + ',fontStyle:"bold",fontSize:15,fontColor:"#0f172a",padding:16},'
        + 'legend:{display:false},'
        + 'layout:{padding:{right:50}},'
        + 'scales:{'
          + 'xAxes:[{ticks:{beginAtZero:true,fontColor:"#64748b",fontSize:12},gridLines:{color:"#f1f5f9",zeroLineColor:"#e2e8f0"}}],'
          + 'yAxes:[{ticks:{fontColor:"#1e293b",fontSize:13},gridLines:{display:false}}]'
        + '},'
        + 'plugins:{datalabels:{'
          + 'anchor:"end",align:"right",'
          + 'formatter:function(v){return v;},'
          + 'color:"#0f172a",font:{size:12,weight:"bold"}'
        + '}}'
      + '}'
    + '}';
    return rhChartFetch_(chart, w || 780, h || 370);
  } catch(e) { Logger.log('rhMakeBarImg: ' + e.message); return ''; }
}

// ── Graphique barres verticales QuickChart (Avis par installation) ──
function rhMakeVertBarImg(labels, values, color, title, w, h) {
  try {
    var chart = '{'
      + 'type:"bar",'
      + 'data:{labels:' + JSON.stringify(labels)
        + ',datasets:[{data:' + JSON.stringify(values)
          + ',backgroundColor:' + JSON.stringify(color || '#0891b2')
          + ',borderWidth:0,borderRadius:4}]},'
      + 'options:{'
        + 'title:{display:' + (title ? 'true' : 'false') + ',text:' + JSON.stringify(title || '') + ',fontStyle:"bold",fontSize:15,fontColor:"#0f172a",padding:16},'
        + 'legend:{display:false},'
        + 'layout:{padding:{top:20}},'
        + 'scales:{'
          + 'xAxes:[{ticks:{fontColor:"#1e293b",fontSize:11,maxRotation:45,minRotation:30},gridLines:{display:false}}],'
          + 'yAxes:[{ticks:{beginAtZero:true,fontColor:"#64748b",fontSize:11},gridLines:{color:"#f1f5f9"}}]'
        + '},'
        + 'plugins:{datalabels:{'
          + 'anchor:"end",align:"top",'
          + 'formatter:function(v){return v>0?v:"";},'
          + 'color:"#0f172a",font:{size:11,weight:"bold"}'
        + '}}'
      + '}'
    + '}';
    return rhChartFetch_(chart, w || 900, h || 380);
  } catch(e) { Logger.log('rhMakeVertBarImg: ' + e.message); return ''; }
}

// ── Graphique colonnes verticales HTML pur — compatible Outlook ──
function rhMakeColHtml(labels, values, color) {
  if (!labels.length) return '';
  var maxVal = Math.max.apply(null, values) || 1;
  var col = color || '#6366f1';
  var MAX_H = 140;
  var n = labels.length;
  var colW = Math.floor(100 / n);
  var html = '<table cellpadding="0" cellspacing="0" width="100%"><tr>';
  for (var i = 0; i < n; i++) {
    var barH = Math.max(4, Math.round((values[i] / maxVal) * MAX_H));
    var spacerH = MAX_H - barH;
    html += '<td width="' + colW + '%" align="center" valign="top">'
      + '<table cellpadding="0" cellspacing="0" align="center">'
      + '<tr><td align="center" style="font-size:11px;font-weight:700;color:' + col + ';padding-bottom:3px;white-space:nowrap;">' + values[i] + '</td></tr>'
      + (spacerH > 0 ? '<tr><td height="' + spacerH + '" style="font-size:1px;">&nbsp;</td></tr>' : '')
      + '<tr><td width="36" height="' + barH + '" bgcolor="' + col + '" style="font-size:1px;">&nbsp;</td></tr>'
      + '<tr><td align="center" style="font-size:10px;font-weight:600;color:#334155;padding-top:5px;white-space:nowrap;">' + labels[i] + '</td></tr>'
      + '</table>'
      + '</td>';
  }
  html += '</tr></table>';
  return html;
}

// Graphique barres en HTML pur — compatible Outlook, valeurs affichées
function rhMakeBarHtml(labels, values, color) {
  if (!labels.length) return '';
  var maxVal = Math.max.apply(null, values) || 1;
  var col = color || '#3b82f6';
  var html = '<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">';
  for (var i = 0; i < labels.length; i++) {
    var pct = Math.round((values[i] / maxVal) * 100);
    var remain = 100 - pct;
    html += '<tr>'
      + '<td width="110" style="font-size:11px;font-weight:600;color:#374151;padding:5px 8px 5px 0;white-space:nowrap;vertical-align:middle;">' + labels[i] + '</td>'
      + '<td style="padding:5px 6px 5px 0;vertical-align:middle;">'
      +   '<table cellpadding="0" cellspacing="0" width="100%"><tr>'
      +     '<td width="' + pct + '%" style="height:18px;background:' + col + ';font-size:1px;">&nbsp;</td>'
      +     (remain > 0 ? '<td style="height:18px;background:#f1f5f9;font-size:1px;">&nbsp;</td>' : '')
      +   '</tr></table>'
      + '</td>'
      + '<td width="32" align="right" style="font-size:11px;font-weight:700;color:' + col + ';padding:5px 0;white-space:nowrap;vertical-align:middle;">' + values[i] + '</td>'
      + '</tr>';
  }
  html += '</table>';
  return html;
}

// ── Construction HTML du rapport (compatible Outlook desktop) ─
function rhBuildHtml(arrets, kpi, avis) {
  var s=arrets.sem, d0=rhFmtDate(arrets.s0), d1=rhFmtDate(arrets.s1);

  // ── Calendrier ──
  function buildCal() {
    if (!arrets.rows.length) return '<p style="color:#94a3b8;font-style:italic;font-size:13px;padding:8px 0;">Aucun arr&#234;t enregistr&#233; pour la semaine S'+s+'.</p>';
    var DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];
    var wMap={};
    arrets.rows.forEach(function(r){
      var parts=r.date.split('-').map(Number);
      var d=new Date(parts[0],parts[1]-1,parts[2]);
      var wn=parseInt(r.semaine.toString().replace(/\D/g,'')); // strip "S" prefix
      var k=r.annee+'-'+wn;
      if(!wMap[k]) wMap[k]={annee:r.annee,semaine:r.semaine,wn:wn,days:{}};
      var dow=(d.getDay()+6)%7;
      if(!wMap[k].days[dow]) wMap[k].days[dow]=[];
      wMap[k].days[dow].push({label:r.install,statut:r.statut});
    });
    var weeks=Object.values(wMap).sort(function(a,b){return a.annee!==b.annee?a.annee-b.annee:a.wn-b.wn;});
    var TH='padding:6px 8px;border:1px solid #e2e8f0;font-size:10px;font-weight:700;color:#8a97ab;text-transform:uppercase;background:#f8fafc;text-align:center;';
    var WH='padding:7px 10px;border:1px solid #e2e8f0;font-size:11px;font-weight:700;background:#f8fafc;white-space:nowrap;color:#1d4ed8;';
    var TC='padding:5px 6px;border:1px solid #e2e8f0;vertical-align:top;';
    var DT='font-size:9px;color:#94a3b8;text-align:center;padding:3px 4px;background:#f8fafc;border:1px solid #e2e8f0;border-top:none;';
    var html='<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:12px;">';
    // En-tête
    html+='<tr><th style="'+WH+'">Semaine</th>'+DAYS.map(function(d){return'<th style="'+TH+'">'+d+'</th>';}).join('')+'</tr>';
    weeks.forEach(function(w,wi){
      var mon=rhMondayOf(w.annee,w.wn);
      var sepT=wi>0?'border-top:2px solid #cbd5e1;':'';
      // Ligne dates
      html+='<tr><td style="'+WH+sepT+'">'+w.semaine+'</td>';
      for(var di=0;di<7;di++){
        var dd=new Date(mon); dd.setDate(dd.getDate()+di);
        html+='<td style="'+DT+sepT+'">'+String(dd.getDate()).padStart(2,'0')+'/'+String(dd.getMonth()+1).padStart(2,'0')+'</td>';
      }
      html+='</tr>';
      // Ligne données
      html+='<tr><td style="font-size:9px;color:#94a3b8;text-align:center;font-weight:700;'+TC+'">'+w.annee+'</td>';
      for(var di=0;di<7;di++){
        var items=w.days[di]||[];
        html+='<td style="'+TC+'">';
        if(!items.length){html+='<span style="color:#cbd5e1;">&#183;</span>';}
        else items.forEach(function(it){
          var bg=it.statut==='realise'?'#dcfce7':it.statut==='imprevu'?'#ffedd5':'#fee2e2';
          var fg=it.statut==='realise'?'#166534':it.statut==='imprevu'?'#9a3412':'#991b1b';
          html+='<div style="font-size:10px;font-weight:600;padding:2px 6px;background:'+bg+';color:'+fg+';white-space:nowrap;margin-bottom:4px;">'+it.label+'</div>';
        });
        html+='</td>';
      }
      html+='</tr>';
    });
    html+='</table>';
    return html;
  }

  // ── KPI card (table-based, Outlook-safe) ──
  function kpiCard(iconBg,iconFg,val,label,sub,taux,w) {
    w = w || '50%';
    var tagHtml='';
    if(taux!==null&&taux!==undefined){
      var tc=taux>=80?{bg:'#dcfce7',fg:'#166534'}:taux>=50?{bg:'#fef9c3',fg:'#854d0e'}:{bg:'#fee2e2',fg:'#991b1b'};
      tagHtml='<div style="text-align:right;"><span style="font-size:11px;font-weight:700;padding:2px 8px;background:'+tc.bg+';color:'+tc.fg+';">'+taux.toFixed(1)+'%</span></div>';
    }
    return '<td width="'+w+'" valign="top" style="padding:6px;">'
      +'<table cellpadding="0" cellspacing="0" width="100%" style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;">'
      +'<tr><td style="padding:14px 16px;">'
      +'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
      +'<td><div style="width:32px;height:32px;background:'+iconBg+';text-align:center;padding-top:8px;font-size:14px;font-weight:700;color:'+iconFg+';">&#9632;</div></td>'
      +'<td align="right">'+tagHtml+'</td>'
      +'</tr></table>'
      +'<div style="font-size:24px;font-weight:700;color:#0f172a;margin:10px 0 3px;line-height:1;">'+val+'</div>'
      +'<div style="font-size:12px;font-weight:600;color:#475569;">'+label+'</div>'
      +'<div style="font-size:11px;color:#94a3b8;margin-top:8px;padding-top:8px;border-top:1px solid #f1f5f9;">'+sub+'</div>'
      +'</td></tr></table>'
      +'</td>';
  }

  // ── Chart card image (pie) — hauteur fixe 260px ──
  function chartCard(imgSrc,title,w) {
    w=w||'50%';
    if(!imgSrc) return '<td width="'+w+'" valign="top" style="padding:6px;"></td>';
    return '<td width="'+w+'" valign="top" style="padding:6px;">'
      +'<table cellpadding="0" cellspacing="0" width="100%" style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;">'
      +'<tr><td style="padding:12px 16px;">'
      +'<img src="'+imgSrc+'" width="100%" style="display:block;border:0;height:auto;max-width:100%;" alt="'+title+'">'
      +'</td></tr></table></td>';
  }

  // ── Chart card barres HTML — même hauteur fixe 260px ──
  function barChartCard(barHtml,title,w) {
    w=w||'50%';
    if(!barHtml) return '<td width="'+w+'" valign="top" style="padding:6px;"></td>';
    return '<td width="'+w+'" valign="top" style="padding:6px;">'
      +'<table cellpadding="0" cellspacing="0" width="100%" height="260" style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;">'
      +'<tr><td height="260" valign="middle" style="padding:12px 16px;">'
      +'<div style="font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:10px;">'+title+'</div>'
      +barHtml
      +'</td></tr></table></td>';
  }

  // ── Sous-section label ──
  function subSection(bg,border,color,label) {
    return '<table cellpadding="0" cellspacing="0" style="margin:16px 0 10px;">'
      +'<tr><td style="padding:5px 14px;background:'+bg+';border:1px solid '+border+';border-radius:20px;font-size:10px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:'+color+';">'+label+'</td></tr>'
      +'</table>';
  }

  // ── Section label ──
  function secLabel(txt) {
    return '<table cellpadding="0" cellspacing="0" width="100%" style="margin:24px 0 12px;">'
      +'<tr><td style="font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;padding-bottom:8px;border-bottom:1px solid #e2e8f0;">'+txt+'</td></tr>'
      +'</table>';
  }

  // ── Postes (table-based) ──
  function buildPostes() {
    if(!kpi.postes.length) return '<p style="color:#94a3b8;font-size:12px;">Aucune donn&#233;e.</p>';
    var html='<table cellpadding="0" cellspacing="0" width="100%">';
    kpi.postes.forEach(function(p){
      var c=p.taux>=80?'#059669':p.taux>=50?'#d97706':'#dc2626';
      var w=Math.min(100,Math.round(p.taux));
      html+='<tr style="border-bottom:1px solid #f1f5f9;">'
        +'<td width="130" style="font-size:12px;font-weight:600;color:#0f172a;padding:7px 8px 7px 0;white-space:nowrap;">'+p.nom+'</td>'
        +'<td style="padding:7px 8px;">'
        +'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
        +'<td width="'+w+'%" style="height:8px;background:'+c+';font-size:1px;">&nbsp;</td>'
        +(w<100?'<td style="height:8px;background:#f1f5f9;font-size:1px;">&nbsp;</td>':'')
        +'</tr></table></td>'
        +'<td width="48" align="right" style="font-size:12px;font-weight:700;color:'+c+';padding:7px 0 7px 8px;white-space:nowrap;">'+p.taux.toFixed(1)+'%</td>'
        +'<td width="54" align="right" style="font-size:11px;color:#94a3b8;padding:7px 0 7px 8px;white-space:nowrap;">'+p.real+'/'+p.total+'</td>'
        +'</tr>';
    });
    html+='</table>';
    return html;
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
  +'<body style="margin:0;padding:0;background:#f4f6f9;font-family:Arial,Helvetica,sans-serif;color:#0f172a;">'

  // ── Wrapper centré ──
  +'<table cellpadding="0" cellspacing="0" width="100%"><tr><td align="center" style="padding:12px 4px 32px;">'
  +'<table cellpadding="0" cellspacing="0" width="1300" style="max-width:1300px;">'

  // ── Header ──
  +'<tr><td style="background:#1d4ed8;border-radius:10px 10px 0 0;padding:18px 24px;">'
  +'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
  +'<td><div style="color:#ffffff;font-weight:700;font-size:16px;">Rapport de Maintenance</div>'
  +'<div style="color:#bfdbfe;font-size:11px;margin-top:3px;">Bureau des M&#233;thodes Daoui &middot; OCP Group Khouribga</div></td>'
  +'<td align="right"><span style="background:#ffffff;color:#1d4ed8;font-weight:700;font-size:12px;padding:5px 14px;border-radius:4px;">S'+s+' &middot; '+d0+' &#8594; '+d1+'</span></td>'
  +'</tr></table>'
  +'</td></tr>'

  // ── Contenu ──
  +'<tr><td style="background:#ffffff;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px;padding:20px 18px 32px;">'

  // Calendrier
  +secLabel('Calendrier des arr&#234;ts pr&#233;ventifs &#8212; Semaines du mois (jusqu\'&#224; S'+s+')')
  +'<table cellpadding="0" cellspacing="0" width="100%" style="background:#ffffff;border:1px solid #dde2ea;border-radius:12px;margin-bottom:12px;">'
  +'<tr><td style="padding:18px 20px;">'
  +'<div style="font-size:13px;font-weight:700;color:#0f172a;margin-bottom:6px;">Arr&#234;ts S'+s+' &middot; '+arrets.rows.length+' enregistr&#233;(s)</div>'
  +'<div style="font-size:11px;margin-bottom:12px;">'
  +'<span style="color:#166534;font-weight:700;margin-right:12px;">&#9679; R&#233;alis&#233;</span>'
  +'<span style="color:#991b1b;font-weight:700;margin-right:12px;">&#9679; Non r&#233;alis&#233;</span>'
  +'<span style="color:#9a3412;font-weight:700;">&#9679; Impr&#233;vu</span>'
  +'</div>'
  +buildCal()
  +'</td></tr></table>'

  // ── KPIs synthèse sous le calendrier (3 cards) ──
  +(function(){
    var nbA=arrets.rows.length;
    var nbR=arrets.rows.filter(function(r){return r.statut==='realise';}).length;
    var tA=nbA?parseFloat(((nbR/nbA)*100).toFixed(1)):0;
    var tAStr=nbA?tA.toFixed(1)+'%':'&#8212;';
    var scTot=kpi.sys+kpi.cur;
    var sysPct=scTot?parseFloat(((kpi.sys/scTot)*100).toFixed(1)):0;
    var curPct2=scTot?parseFloat(((kpi.cur/scTot)*100).toFixed(1)):0;
    var sysCurStr='Sys&nbsp;<b>'+sysPct.toFixed(1)+'%</b>&nbsp;/&nbsp;Cur&nbsp;<b>'+curPct2.toFixed(1)+'%</b>';
    return '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:16px;"><tr>'
      +kpiCard('#ecfdf5','#059669',kpi.tauxRealStr,'Taux r&#233;alisation OT','Mois courant &middot; <b>'+kpi.real.toLocaleString('fr-FR')+'</b> / <b>'+kpi.total.toLocaleString('fr-FR')+'</b>',kpi.tauxReal,'33%')
      +kpiCard('#eff6ff','#1d4ed8',tAStr,'Taux r&#233;alisation arr&#234;ts','S'+s+' &middot; <b>'+nbR+'</b> r&#233;alis&#233;(s) / <b>'+nbA+'</b>',tA,'33%')
      +kpiCard('#f5f3ff','#7c3aed',kpi.sys+'/'+kpi.cur,'Taux syst&#233;matique / curatif',sysCurStr,null,'33%')
      +'</tr></table>';
  })()

  // ── Graphiques OT (type d'ordre + corps de métier) ──
  +(function(){
    var typeD=kpi.typeData.slice(0,6);
    var postesD=kpi.postes.slice(0,7);
    var TYPE_LBL={'ZEST':'Syst\u00e9matique','ZCOR':'Curatif','ZCON':'Conditionnel','ZETL':'Etalonnage'};
    var imgType=typeD.length?rhMakePieImg(
      typeD.map(function(x){return TYPE_LBL[x.type]||x.type;}),
      typeD.map(function(x){return x.count;}),
      'R\u00e9partition par type d\'ordre'):'';
    var imgPoste=postesD.length?rhMakeBarImg(
      postesD.map(function(x){return x.nom;}),
      postesD.map(function(x){return x.total;}),
      '#3b82f6','Volume OT par corps de m\u00e9tier'):'';
    return '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:20px;"><tr>'
      +chartCard(imgType,'R&#233;partition par type d\'ordre')
      +chartCard(imgPoste,'Volume OT par corps de m&#233;tier')
      +'</tr></table>';
  })()

  // KPIs détaillés
  +secLabel('Indicateurs cl&#233;s du mois &#8212; '+kpi.mois)

  // Global – 6 KPIs en une seule ligne
  +subSection('#eff6ff','#bfdbfe','#1d4ed8','&#9632; Global')
  +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:8px;"><tr>'
  +kpiCard('#eff6ff','#1d4ed8',kpi.total.toLocaleString('fr-FR'),'Total OT planifi&#233;s','Ordres de travail du mois',null,'16%')
  +kpiCard('#ecfdf5','#059669',kpi.real.toLocaleString('fr-FR'),'OT R&#233;alis&#233;s','Taux : <b>'+kpi.tauxRealStr+'</b>',kpi.tauxReal,'16%')
  +kpiCard('#fffbeb','#d97706',kpi.lanc.toLocaleString('fr-FR'),'OT Lanc&#233;s','En cours : <b>'+kpi.lancPct+'</b>',null,'16%')
  +kpiCard('#fef2f2','#dc2626',kpi.crpr.toLocaleString('fr-FR'),'Non lanc&#233;s (CRPR)','En attente de planification',null,'16%')
  +kpiCard('#f1f5f9','#475569',kpi.backlog.toLocaleString('fr-FR'),'Backlog','En attente de planification',null,'16%')
  +kpiCard('#fffbeb','#d97706',kpi.sys+'/'+kpi.cur,'Pr&#233;ventif / Correctif',kpi.sysPct+' / '+kpi.curPct,null,'20%')
  +'</tr></table>'

  // Préventif + Correctif – 4 KPIs en une seule ligne
  +subSection('#e0e7ff','#a5b4fc','#3730a3','&#9632; Pr&#233;ventif &nbsp;&nbsp;&#9632; Correctif')
  +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:8px;"><tr>'
  +kpiCard('#f5f3ff','#7c3aed',kpi.sys.toLocaleString('fr-FR'),'OT Pr&#233;ventif','ZCON + ZEST + ZETL : <b>'+kpi.sysPct+'</b>',null,'25%')
  +kpiCard('#ecfdf5','#059669',kpi.tauxPrevStr,'Taux r&#233;alisation','CONF+TCLO+CLOT / ZCON+ZEST+ZETL',kpi.tauxPrev,'25%')
  +kpiCard('#fef2f2','#dc2626',kpi.cur.toLocaleString('fr-FR'),'OT Correctif (ZCOR)','Part : <b>'+kpi.curPct+'</b>',null,'25%')
  +kpiCard('#fef2f2','#dc2626',kpi.tauxCorStr,'Taux r&#233;alisation','CONF+TCLO+CLOT / total ZCOR',kpi.tauxCor,'25%')
  +'</tr></table>'

  // Préparation
  +subSection('#fff7ed','#fdba74','#c2410c','&#9632; Pr&#233;paration')
  +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:8px;"><tr>'
  +kpiCard('#f0fdf4','#15803d',kpi.tauxCaractStr,'Taux caract&#233;risation','<b>'+kpi.caract.toLocaleString('fr-FR')+'</b> / <b>'+(kpi.caract+kpi.nonCaract).toLocaleString('fr-FR')+'</b> &middot; excl. SOPL',kpi.tauxCaract,'25%')
  +kpiCard('#fff7ed','#c2410c',kpi.tauxPdrConfStr,'Taux confirmation PDR','<b>'+kpi.pdrConf.toLocaleString('fr-FR')+'</b> confirm&#233;s sur <b>'+kpi.pdrTotal.toLocaleString('fr-FR')+'</b> OTs avec PDR',kpi.tauxPdrConf,'25%')
  +kpiCard('#eff6ff','#1d4ed8',(kpi.otAttente||0).toLocaleString('fr-FR'),'OT en attente confirmation PDR','Statut utilis. CRPR ATPD / CRPR AVPD',null,'25%')
  +kpiCard('#fdf4ff','#7c3aed',kpi.tempsMoyenStr||'—','Temps moyen de pr&#233;paration','Moy. (aujourd\'hui &minus; date cr&#233;ation) &middot; statut syst. <b>cr&#233;&#233;</b>',null,'25%')
  +'</tr></table>'

  // Postes
  +secLabel('Taux de r&#233;alisation par corps de m&#233;tier')
  +'<table cellpadding="0" cellspacing="0" width="100%" style="background:#ffffff;border:1px solid #dde2ea;border-radius:12px;margin-bottom:14px;">'
  +'<tr><td style="padding:18px 20px;">'
  +'<div style="font-size:13px;font-weight:700;color:#0f172a;margin-bottom:12px;">'+kpi.mois+'</div>'
  +buildPostes()
  +'</td></tr></table>'

  // ── Section Avis ──
  +(avis ? (
    secLabel('Analyse des Avis (ZC) &#8212; '+kpi.mois)
    // KPIs Avis
    +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:12px;"><tr>'
    +kpiCard('#fdf4ff','#7c3aed',avis.total.toLocaleString('fr-FR'),'Total Avis','Avis de type ZC',null,'33%')
    +kpiCard('#ecfdf5','#059669',avis.avecOT.toLocaleString('fr-FR'),'Convertis en OT','Taux : <b>'+avis.txConv.toFixed(1)+'%</b>',avis.txConv,'33%')
    +kpiCard('#fef2f2','#dc2626',avis.ouverts.toLocaleString('fr-FR'),'Avis Ouverts','AOUV + AENC',null,'33%')
    +'</tr></table>'
    // Graphiques Avis ligne 1 : Répartition par secteur + Corps de Métier
    +(function(){
      var imgSect=avis.bySecteur.length?rhMakePieImg(
        avis.bySecteur.map(function(x){return x.label;}),
        avis.bySecteur.map(function(x){return x.count;}),
        'R\u00e9partition par secteur'):'';
      var imgPoste=avis.byPoste.length?rhMakeBarImg(
        avis.byPoste.map(function(x){return x.label;}),
        avis.byPoste.map(function(x){return x.count;}),
        '#7c3aed','Corps de M\u00e9tier (Poste trav.)'):'';
      return '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:8px;"><tr>'
        +chartCard(imgSect,'R&#233;partition par secteur')
        +chartCard(imgPoste,'Corps de M&#233;tier (Poste trav.)')
        +'</tr></table>';
    })()
    // Graphiques Avis ligne 2 : Non Clôturés + Par installation
    +(function(){
      var imgOpen=avis.openByPoste.length?rhMakeBarImg(
        avis.openByPoste.map(function(x){return x.label;}),
        avis.openByPoste.map(function(x){return x.count;}),
        '#dc2626','Avis non cl\u00f4tur\u00e9s par corps de m\u00e9tier'):'';
      var imgInst=avis.byInstall.length?rhMakeVertBarImg(
        avis.byInstall.map(function(x){return x.label;}),
        avis.byInstall.map(function(x){return x.count;}),
        '#0891b2','Avis par installation'):'';
      return '<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:20px;"><tr>'
        +chartCard(imgOpen,'Avis non cl&#244;tur&#233;s par corps de m&#233;tier')
        +'</tr></table>'
        +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:20px;"><tr>'
        +chartCard(imgInst,'Avis par installation','100%')
        +'</tr></table>';
    })()
  ) : '')

  // Signature
  +'<table cellpadding="0" cellspacing="0" width="100%" style="margin-top:28px;border-top:1px solid #e2e8f0;">'
  +'<tr><td style="padding-top:18px;">'
  +'<p style="color:#64748b;font-size:13px;margin:0 0 8px;">Cordialement,</p>'
  +'<div style="font-family:Georgia,serif;font-size:14px;color:#002060;line-height:1.7;">'
  +'<strong>Marouane ELAMRAOUI</strong><br>'
  +'<span style="color:#c55a11;">M&#233;thode de Maintenance</span><br>'
  +'<strong>OCP SA - Khouribga</strong><br>'
  +'<span style="color:#059669;">T&#233;l. :</span> 0661323784 &nbsp;|&nbsp; <span style="color:#059669;">Cisco :</span> 8103388<br>'
  +'<a href="mailto:m.elamraoui@ocpgroup.ma" style="color:#002060;">m.elamraoui@ocpgroup.ma</a>'
  +'</div>'
  +'</td></tr></table>'

  +'</td></tr>'        // fin contenu
  +'</table>'          // fin table 900px
  +'</td></tr></table>' // fin wrapper centré
  +'</body></html>';
}

// ── Envoi depuis l'interface ──────────────────────────────────
function envoyerRapportDepuisInterface(p) {
  try {
    if (!p.emails) return { ok: false, msg: 'Aucun destinataire.' };
    var mo = parseInt(p.mo), yr = parseInt(p.yr);
    var arrets = rhGetArrets();
    var kpi    = rhGetKpi(mo, yr);
    var prep   = rhGetPreparation(mo, yr);
    if (prep) { kpi.pdrTotal=prep.pdrTotal; kpi.pdrConf=prep.pdrConf; kpi.tauxPdrConf=prep.tauxPdrConf; kpi.tauxPdrConfStr=prep.tauxPdrConfStr; kpi.otAttente=prep.otAttente; kpi.tempsMoyenStr=prep.tempsMoyenStr; }
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
  var prep   = rhGetPreparation(mo, yr);
  if (prep) { kpi.pdrTotal=prep.pdrTotal; kpi.pdrConf=prep.pdrConf; kpi.tauxPdrConf=prep.tauxPdrConf; kpi.tauxPdrConfStr=prep.tauxPdrConfStr; kpi.otAttente=prep.otAttente; kpi.tempsMoyenStr=prep.tempsMoyenStr; }
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
  var prep   = rhGetPreparation(mo, yr);
  if (prep) { kpi.pdrTotal=prep.pdrTotal; kpi.pdrConf=prep.pdrConf; kpi.tauxPdrConf=prep.tauxPdrConf; kpi.tauxPdrConfStr=prep.tauxPdrConfStr; kpi.otAttente=prep.otAttente; kpi.tempsMoyenStr=prep.tempsMoyenStr; }
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
