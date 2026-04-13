// OCP_EMAIL_EWS, EWS_URL_EWS, getOcpPasswordEWS() et sendEmailEWS() sont définis dans Envoi-AVIS-MEC-INST.gs

// ============================================================
//  BUREAU MÉTHODE DAOUI — Priorisation Bobinage
//  Google Apps Script — À coller dans Extensions > Apps Script
// ============================================================

var CONFIG = {
  EMAIL_BOBINAGE      : "m.elamraoui@ocpgroup.ma",
  EMAIL_EXPEDITEUR    : "",
  EMAIL_REPLY_TO      : "m.elamraoui@ocpgroup.ma",
  SHEET_STOCK         : "Stock",
  SHEET_DEMANDE       : "Demande des intercheable",
  COL_STOCK_MATRICULE : "Matricule",
  COL_STOCK_PUISSANCE : "Puissance",
  COL_STOCK_STATUT    : "Statut",
  STATUT_EN_STOCK     : ["en stock","stock","disponible","reserve","réserve"],
  STATUT_EN_REPARE    : ["en réparation","en reparation","réparation","reparation","bobinage"],
  STATUT_INSTALLE     : ["installé","installe","en service","actif"],
  COL_DEM_MATRICULE   : "Matricule",
  COL_DEM_PUISSANCE   : "Puissance",
  COL_DEM_TENSION     : "Tension",
  COL_DEM_VITESSE     : "Vitesse",
  COL_DEM_DATE        : "Date d'opération",
  COL_DEM_STATUT      : "Statut de réparation",
  STATUT_DEM_ACTIF    : "en réparation",
  PUISSANCE_MIN       : 5.5,
  FLENDER_PUISSANCE   : 5.5,
  FLENDER_MARQUE      : "flender",
  FLENDER_STOCK_SEUIL : 3,
  PUISSANCES_BONUS    : [5.5,7.5,55,75,90,160,168,315,368,500],
  BONUS_SCORE         : 20,
  NB_PRIORITES        : 3,
};


// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Bobinage")
    .addItem("📊 Analyser les priorités",         "analyserEtAfficher")
    .addSeparator()
    .addItem("✉️  Envoyer le mail maintenant",     "envoyerMailManuel")
    .addItem("👁️  Prévisualiser le mail",          "previsualiserMail")
    .addSeparator()
    .addItem("⏰  Configurer l'envoi automatique", "configurerTrigger")
    .addToUi();
}


// ============================================================
//  ENVOI MANUEL
// ============================================================
function envoyerMailManuel() {
  var ui     = SpreadsheetApp.getUi();
  var result = calculerPriorites();
  if (result.erreur) { ui.alert("❌ Erreur\n\n" + result.erreur); return; }

  var apercu = result.top3.map(function(m, i) {
    return (i+1) + ". " + m.matricule + " — " + m.puissance + " kW (" + m.delai + " jours)";
  }).join("\n");

  var rep = ui.alert(
    "✉️ Confirmation d'envoi — Semaine " + result.semaine,
    "Destinataire : " + CONFIG.EMAIL_BOBINAGE + "\n\nTop 3 :\n" + apercu + "\n\nConfirmer ?",
    ui.ButtonSet.YES_NO
  );
  if (rep !== ui.Button.YES) return;

  try {
    sendEmailEWS(CONFIG.EMAIL_BOBINAGE, null, genererSujet(result), genererCorpsMailHTML(result), null,
      "Bureau de méthode Daoui - Section Interchangeable");
    loggerEnvoi(result);
    ui.alert("✅ Mail envoyé !", "Envoyé à " + CONFIG.EMAIL_BOBINAGE + "\nSemaine " + result.semaine, ui.ButtonSet.OK);
  } catch(e) { ui.alert("❌ Erreur\n\n" + e.message); }
}

function envoyerMailDepuisSidebar() {
  var result = calculerPriorites();
  if (result.erreur) return "❌ " + result.erreur;
  try {
    sendEmailEWS(CONFIG.EMAIL_BOBINAGE, null, genererSujet(result), genererCorpsMailHTML(result), null,
      "Bureau de méthode Daoui - Section Interchangeable");
    loggerEnvoi(result);
    return "✅ Mail envoyé — Semaine " + result.semaine;
  } catch(e) { return "❌ " + e.message; }
}

function envoyerMailPriorite() { envoyerMailManuel(); }


// ============================================================
//  TRIGGER
// ============================================================
function configurerTrigger() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:'Google Sans',sans-serif;background:#0a0e14;color:#e2e8f0;padding:20px}
      h2{color:#f97316;font-size:16px;margin-bottom:4px}
      .sub{color:#64748b;font-size:11px;font-family:monospace;margin-bottom:20px}
      label{display:block;font-size:11px;color:#94a3b8;margin-bottom:5px;text-transform:uppercase;letter-spacing:1px}
      select,input{width:100%;padding:9px 12px;background:#1a2332;border:1px solid #243044;border-radius:8px;color:#e2e8f0;font-size:13px;margin-bottom:14px;outline:none}
      select:focus,input:focus{border-color:#f97316}
      .section{background:#111820;border:1px solid #243044;border-radius:8px;padding:14px;margin-bottom:14px}
      .section-title{font-size:11px;color:#f97316;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px}
      .freq-opts{display:flex;flex-direction:column;gap:8px;margin-bottom:14px}
      .freq-opt{display:flex;align-items:center;gap:10px;padding:9px 12px;background:#1a2332;border:1px solid #243044;border-radius:8px;cursor:pointer}
      .freq-opt:hover{border-color:#f97316}
      .freq-opt input[type=radio]{accent-color:#f97316;width:15px;height:15px;flex-shrink:0}
      .freq-opt .label{font-size:13px;font-weight:600}
      .freq-opt .desc{font-size:11px;color:#64748b}
      .status-box{border:1px solid #243044;border-radius:8px;padding:10px 14px;font-family:monospace;font-size:12px;color:#64748b;margin-bottom:14px}
      .status-box.active{color:#10b981;border-color:rgba(16,185,129,0.4);background:rgba(16,185,129,0.06)}
      .status-box.inactive{color:#ef4444;border-color:rgba(239,68,68,0.4);background:rgba(239,68,68,0.06)}
      .btn{width:100%;padding:11px;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;margin-bottom:8px;font-family:'Google Sans',sans-serif}
      .btn-primary{background:#f97316;color:#000}.btn-primary:hover{background:#ea6c0a}
      .btn-danger{background:#1a2332;color:#ef4444;border:1px solid rgba(239,68,68,0.3)}
      .btn-danger:hover{background:rgba(239,68,68,0.08)}
      .msg{display:none;padding:9px 13px;border-radius:8px;font-size:12px;margin-bottom:10px}
      .msg.ok{background:rgba(16,185,129,0.1);color:#10b981;border:1px solid rgba(16,185,129,0.3);display:block}
      .msg.err{background:rgba(239,68,68,0.1);color:#ef4444;border:1px solid rgba(239,68,68,0.3);display:block}
      .note{font-size:10px;color:#475569;margin-top:-10px;margin-bottom:14px;font-family:monospace}
    </style>
    <h2>⏰ Envoi automatique</h2>
    <div class="sub">Configuration du mail — Service Bobinage</div>
    <label>Statut actuel</label>
    <div class="status-box" id="statusBox">Chargement...</div>
    <div class="section">
      <div class="section-title">Fréquence</div>
      <div class="freq-opts">
        <label class="freq-opt">
          <input type="radio" name="freq" value="hebdo" checked onchange="toggleFreq()">
          <div><div class="label">📅 Hebdomadaire</div><div class="desc">Une fois par semaine</div></div>
        </label>
        <label class="freq-opt">
          <input type="radio" name="freq" value="quotidien" onchange="toggleFreq()">
          <div><div class="label">🔁 Quotidien</div><div class="desc">Tous les jours</div></div>
        </label>
        <label class="freq-opt">
          <input type="radio" name="freq" value="mensuel" onchange="toggleFreq()">
          <div><div class="label">🗓️ Mensuel</div><div class="desc">Une fois par mois</div></div>
        </label>
      </div>
    </div>
    <div id="daySection">
      <label>Jour de la semaine</label>
      <select id="jour">
        <option value="MONDAY">Lundi</option><option value="TUESDAY">Mardi</option>
        <option value="WEDNESDAY">Mercredi</option><option value="THURSDAY">Jeudi</option>
        <option value="FRIDAY" selected>Vendredi</option>
        <option value="SATURDAY">Samedi</option><option value="SUNDAY">Dimanche</option>
      </select>
    </div>
    <div id="domSection" style="display:none">
      <label>Jour du mois (1 à 28)</label>
      <input type="number" id="jourMois" value="1" min="1" max="28"/>
      <div class="note">⚠ Maximum 28 pour éviter les erreurs en février</div>
    </div>
    <label>Heure d'envoi (0 – 23)</label>
    <input type="number" id="heure" value="9" min="0" max="23"/>
    <div id="msg" class="msg"></div>
    <button class="btn btn-primary" onclick="activer()">✅ Activer / Mettre à jour</button>
    <button class="btn btn-danger"  onclick="desactiver()">🔕 Désactiver</button>
    <script>
      google.script.run.withSuccessHandler(function(info) {
        var box = document.getElementById('statusBox');
        if (info.actif) {
          box.textContent = '✅ ' + info.description;
          box.className = 'status-box active';
          if (info.jourCode) document.getElementById('jour').value = info.jourCode;
          if (info.heure !== undefined) document.getElementById('heure').value = info.heure;
          if (info.freq) {
            document.querySelectorAll('input[name=freq]').forEach(function(r){ if(r.value===info.freq) r.checked=true; });
            toggleFreq();
          }
        } else {
          box.textContent = '🔕 Inactif — aucun envoi programmé';
          box.className = 'status-box inactive';
        }
      }).getTriggerInfo();
      function toggleFreq() {
        var freq = document.querySelector('input[name=freq]:checked').value;
        document.getElementById('daySection').style.display = freq==='hebdo' ? 'block' : 'none';
        document.getElementById('domSection').style.display = freq==='mensuel' ? 'block' : 'none';
      }
      function activer() {
        var freq=document.querySelector('input[name=freq]:checked').value;
        var heure=parseInt(document.getElementById('heure').value);
        var jour=document.getElementById('jour').value;
        var jourM=parseInt(document.getElementById('jourMois').value);
        if(isNaN(heure)||heure<0||heure>23){showMsg("Heure invalide (0-23).",false);return;}
        if(freq==='mensuel'&&(isNaN(jourM)||jourM<1||jourM>28)){showMsg("Jour invalide (1-28).",false);return;}
        google.script.run
          .withSuccessHandler(function(res){showMsg(res,true);rechargerStatut();})
          .withFailureHandler(function(e){showMsg("Erreur : "+e.message,false);})
          .installerTrigger(freq,jour,heure,jourM);
      }
      function desactiver() {
        google.script.run
          .withSuccessHandler(function(res){showMsg(res,true);rechargerStatut();})
          .withFailureHandler(function(e){showMsg("Erreur : "+e.message,false);})
          .desactiverTrigger();
      }
      function rechargerStatut() {
        google.script.run.withSuccessHandler(function(info){
          var box=document.getElementById('statusBox');
          box.textContent=info.actif?'✅ '+info.description:'🔕 Inactif — aucun envoi programmé';
          box.className='status-box '+(info.actif?'active':'inactive');
        }).getTriggerInfo();
      }
      function showMsg(text,ok){
        var el=document.getElementById('msg');
        el.textContent=text;
        el.className='msg '+(ok?'ok':'err');
      }
    <\/script>
  `).setTitle("Envoi automatique").setWidth(400).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, "⏰ Configuration de l'envoi automatique");
}

function getTriggerInfo() {
  var props = PropertiesService.getScriptProperties();
  var freq  = props.getProperty("TRIGGER_FREQ")  || "hebdo";
  var jour  = props.getProperty("TRIGGER_JOUR")  || "FRIDAY";
  var heure = props.getProperty("TRIGGER_HEURE") || "9";
  var jours = {"MONDAY":"Lundi","TUESDAY":"Mardi","WEDNESDAY":"Mercredi",
    "THURSDAY":"Jeudi","FRIDAY":"Vendredi","SATURDAY":"Samedi","SUNDAY":"Dimanche"};
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "envoyerMailAutomatique") {
      var h = parseInt(heure);
      var desc = "";
      if      (freq==="hebdo")     desc = (jours[jour]||jour) + " à " + h + "h00";
      else if (freq==="quotidien") desc = "Tous les jours à " + h + "h00";
      else if (freq==="mensuel")   desc = "Le " + (props.getProperty("TRIGGER_JOURM")||"1") + " du mois à " + h + "h00";
      return {actif:true, freq:freq, jourCode:jour, heure:h, description:desc};
    }
  }
  return {actif:false};
}

function installerTrigger(freq, jourCode, heure, jourMois) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="envoyerMailAutomatique") ScriptApp.deleteTrigger(t);
  });
  var joursMap = {
    "MONDAY":ScriptApp.WeekDay.MONDAY,"TUESDAY":ScriptApp.WeekDay.TUESDAY,
    "WEDNESDAY":ScriptApp.WeekDay.WEDNESDAY,"THURSDAY":ScriptApp.WeekDay.THURSDAY,
    "FRIDAY":ScriptApp.WeekDay.FRIDAY,"SATURDAY":ScriptApp.WeekDay.SATURDAY,"SUNDAY":ScriptApp.WeekDay.SUNDAY
  };
  var joursLabels = {"MONDAY":"Lundi","TUESDAY":"Mardi","WEDNESDAY":"Mercredi",
    "THURSDAY":"Jeudi","FRIDAY":"Vendredi","SATURDAY":"Samedi","SUNDAY":"Dimanche"};
  var builder = ScriptApp.newTrigger("envoyerMailAutomatique").timeBased();
  var desc = "";
  if (freq==="quotidien") {
    builder.everyDays(1).atHour(heure).create();
    desc = "Tous les jours à " + heure + "h00";
  } else if (freq==="mensuel") {
    builder.onMonthDay(jourMois).atHour(heure).create();
    desc = "Le " + jourMois + " du mois à " + heure + "h00";
  } else {
    builder.onWeekDay(joursMap[jourCode]||ScriptApp.WeekDay.FRIDAY).atHour(heure).create();
    desc = (joursLabels[jourCode]||jourCode) + " à " + heure + "h00";
  }
  PropertiesService.getScriptProperties().setProperties({
    TRIGGER_FREQ:freq, TRIGGER_JOUR:jourCode,
    TRIGGER_HEURE:String(heure), TRIGGER_JOURM:String(jourMois||1)
  });
  return "✅ Envoi programmé : " + desc;
}

function desactiverTrigger() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==="envoyerMailAutomatique") { ScriptApp.deleteTrigger(t); count++; }
  });
  return count > 0 ? "🔕 Envoi automatique désactivé." : "ℹ️ Aucun trigger actif trouvé.";
}

function envoyerMailAutomatique() {
  var result = calculerPriorites();
  if (result.erreur) {
    sendEmailEWS(CONFIG.EMAIL_REPLY_TO, null,
      "⚠️ Erreur — Priorisation Bobinage Semaine " + getNumeroSemaine(new Date()),
      "<pre style='font-family:monospace'>Erreur :\n\n" + result.erreur + "</pre>",
      null, "Bureau de méthode Daoui - Section Interchangeable");
    return;
  }
  sendEmailEWS(CONFIG.EMAIL_BOBINAGE, null, genererSujet(result), genererCorpsMailHTML(result), null,
    "Bureau de méthode Daoui - Section Interchangeable");
  loggerEnvoi(result);
}


// ============================================================
//  ANALYSER + SIDEBAR
// ============================================================
function analyserEtAfficher() {
  var result = calculerPriorites();
  if (result.erreur) { SpreadsheetApp.getUi().alert("❌ Erreur\n\n" + result.erreur); return; }
  afficherSidebar(result);
}

function previsualiserMail() {
  var result = calculerPriorites();
  if (result.erreur) { SpreadsheetApp.getUi().alert("❌ Erreur\n\n" + result.erreur); return; }
  var html = HtmlService.createHtmlOutput(
    "<pre style='font-family:monospace;font-size:12px;white-space:pre-wrap;padding:12px;line-height:1.6'>" +
    escapeHtml(genererCorpsMail(result)) + "</pre>"
  ).setTitle("Prévisualisation S" + result.semaine).setWidth(650).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "✉️ Prévisualisation — Semaine " + result.semaine);
}


// ============================================================
//  TEST
// ============================================================
function testerBobinage() {
  var result = calculerPriorites();
  if (result.erreur) {
    Logger.log('❌ Erreur calcul priorités : ' + result.erreur);
    return;
  }
  sendEmailEWS(OCP_EMAIL_EWS, null,
    '[TEST] ' + genererSujet(result),
    genererCorpsMailHTML(result),
    null,
    "Bureau de méthode Daoui - Section Interchangeable"
  );
  Logger.log('✅ Email de test envoyé à : ' + OCP_EMAIL_EWS + ' — Semaine ' + result.semaine);
}


// ============================================================
//  CALCUL DES PRIORITÉS
// ============================================================
function calculerPriorites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetStock = ss.getSheetByName(CONFIG.SHEET_STOCK);
  if (!sheetStock) return {erreur:"Onglet '"+CONFIG.SHEET_STOCK+"' introuvable."};
  var dataStock = sheetStock.getDataRange().getValues();
  if (dataStock.length < 2) return {erreur:"L'onglet Stock est vide."};

  var headersStock = dataStock[0].map(function(h){return h.toString().trim();});
  var iMatS  = trouverColonne(headersStock, CONFIG.COL_STOCK_MATRICULE);
  var iPwS   = trouverColonne(headersStock, CONFIG.COL_STOCK_PUISSANCE);
  var iStatS = trouverColonne(headersStock, CONFIG.COL_STOCK_STATUT);
  var iTensS = trouverColonne(headersStock, CONFIG.COL_DEM_TENSION);
  var iVitS  = trouverColonne(headersStock, CONFIG.COL_DEM_VITESSE);

  if (iMatS<0||iPwS<0||iStatS<0)
    return {erreur:"Colonnes introuvables dans Stock.\nTrouvé : "+headersStock.join(" | ")};

  var stockByPower = {};
  for (var r = 1; r < dataStock.length; r++) {
    var row    = dataStock[r];
    var pw     = parseFloat(row[iPwS].toString().replace(",","."));
    if (isNaN(pw)||pw<CONFIG.PUISSANCE_MIN) continue;
    var statut  = normaliser(row[iStatS].toString());
    var tension = iTensS>=0 ? row[iTensS].toString().trim() : "";
    var vitesse = iVitS>=0  ? row[iVitS].toString().trim()  : "";
    var key     = pw+"|"+tension+"|"+vitesse;
    if (!stockByPower[key]) stockByPower[key]={puissance:pw,tension:tension,vitesse:vitesse,stock:0,repare:0,installe:0,total:0};
    stockByPower[key].total++;
    if      (contient(statut,CONFIG.STATUT_EN_STOCK))  stockByPower[key].stock++;
    else if (contient(statut,CONFIG.STATUT_EN_REPARE)) stockByPower[key].repare++;
    else                                               stockByPower[key].installe++;
  }

  var sheetDem = ss.getSheetByName(CONFIG.SHEET_DEMANDE);
  if (!sheetDem) return {erreur:"Onglet '"+CONFIG.SHEET_DEMANDE+"' introuvable."};
  var dataDem = sheetDem.getDataRange().getValues();
  if (dataDem.length<2) return {erreur:"L'onglet Demande est vide."};

  var headersDem  = dataDem[0].map(function(h){return h.toString().trim();});
  var iMatD       = trouverColonne(headersDem, CONFIG.COL_DEM_MATRICULE);
  var iPwD        = trouverColonne(headersDem, CONFIG.COL_DEM_PUISSANCE);
  var iDateD      = trouverColonne(headersDem, CONFIG.COL_DEM_DATE);
  var iStatD      = trouverColonne(headersDem, CONFIG.COL_DEM_STATUT);
  var iTensD      = trouverColonne(headersDem, CONFIG.COL_DEM_TENSION);
  var iVitD       = trouverColonne(headersDem, CONFIG.COL_DEM_VITESSE);
  var iMarqueD    = trouverColonne(headersDem, "Marque");
  var iAnomalieD  = trouverColonne(headersDem, "Anomalie");
  var iAvisD      = trouverColonne(headersDem, "Avis/OT");
  var iDateEnvoiD = trouverColonne(headersDem, "Date d'envoi au réparation");
  var iDateOpD    = trouverColonne(headersDem, "Date d'opération");
  var iColAE      = 30;

  if (iMatD<0||iPwD<0||iDateD<0)
    return {erreur:"Colonnes introuvables dans Demande.\nTrouvé : "+headersDem.join(" | ")};
  if (iStatD<0)
    return {erreur:"Colonne '"+CONFIG.COL_DEM_STATUT+"' introuvable.\nTrouvé : "+headersDem.join(" | ")};

  var today = new Date(); today.setHours(0,0,0,0);
  var moteurs = [];

  for (var r = 1; r < dataDem.length; r++) {
    var row       = dataDem[r];
    var statutDem = normaliser(row[iStatD].toString());
    if (!contient(statutDem,[CONFIG.STATUT_DEM_ACTIF])) continue;

    var valColAE = row[iColAE];
    if (valColAE===false||valColAE.toString().trim().toUpperCase()==="FALSE") continue;

    var matricule = row[iMatD].toString().trim();
    var pw        = parseFloat(row[iPwD].toString().replace(",","."));
    if (!matricule||isNaN(pw)||pw<CONFIG.PUISSANCE_MIN) continue;

    var dateVal = row[iDateD];
    var date = null;
    if (dateVal instanceof Date&&!isNaN(dateVal)) { date=new Date(dateVal); }
    else if (typeof dateVal==="string"&&dateVal.trim()!=="") { date=parseDate(dateVal.trim()); }
    if (!date||isNaN(date)) continue;
    date.setHours(0,0,0,0);

    var delai    = Math.max(0,Math.floor((today-date)/86400000));
    var tension  = iTensD>=0   ? row[iTensD].toString().trim()    : "";
    var vitesse  = iVitD>=0    ? row[iVitD].toString().trim()     : "";
    var marque   = iMarqueD>=0 ? row[iMarqueD].toString().trim()  : "";
    var anomalie = iAnomalieD>=0? row[iAnomalieD].toString().trim(): "";
    var avis     = iAvisD>=0   ? row[iAvisD].toString().trim()    : "";

    var rawDE = iDateEnvoiD>=0 ? row[iDateEnvoiD] : "";
    var rawDO = iDateOpD>=0    ? row[iDateOpD]    : "";
    var dateEnvoi = "";
    if      (rawDE instanceof Date&&!isNaN(rawDE))              dateEnvoi = formatDate(rawDE);
    else if (typeof rawDE==="string"&&rawDE.trim()!=="")         dateEnvoi = rawDE.trim();
    else if (rawDO instanceof Date&&!isNaN(rawDO))              dateEnvoi = formatDate(rawDO);
    else if (typeof rawDO==="string"&&rawDO.trim()!=="")         dateEnvoi = rawDO.trim();

    var key        = pw+"|"+tension+"|"+vitesse;
    var stockInfo  = stockByPower[key]||{puissance:pw,tension:tension,vitesse:vitesse,stock:0,repare:0,installe:0,total:0};
    var stockDispo = stockInfo.stock;

    var estFlenderUrgent = (pw===CONFIG.FLENDER_PUISSANCE &&
      normaliser(marque).indexOf(normaliser(CONFIG.FLENDER_MARQUE))>=0 &&
      stockDispo<CONFIG.FLENDER_STOCK_SEUIL);

    moteurs.push({
      matricule:matricule, puissance:pw, tension:tension, vitesse:vitesse,
      marque:marque, anomalie:anomalie, avis:avis, dateEnvoi:dateEnvoi,
      dateStr:formatDate(date), delai:delai,
      stockDispo:stockDispo, stockInfo:stockInfo, score:0,
      prioriteAbsolue:(stockDispo===0||estFlenderUrgent),
      estFlenderUrgent:estFlenderUrgent,
      estPuissanceBonus:(CONFIG.PUISSANCES_BONUS.indexOf(pw)>=0)
    });
  }

  if (moteurs.length===0) return {erreur:"Aucun moteur valide trouvé."};

  var absolus = moteurs.filter(function(m){return  m.prioriteAbsolue;});
  var normaux = moteurs.filter(function(m){return !m.prioriteAbsolue;});
  var delaiMax = 0;
  normaux.forEach(function(m){if(m.delai>delaiMax)delaiMax=m.delai;});
  normaux.forEach(function(m){
    var sS=70*(1/m.stockDispo), sD=delaiMax>0?30*(m.delai/delaiMax):0, sB=m.estPuissanceBonus?CONFIG.BONUS_SCORE:0;
    m.score=Math.round((sS+sD+sB)*100)/100;
  });
  absolus.forEach(function(m){m.score=null;});
  absolus.sort(function(a,b){return b.delai-a.delai;});
  normaux.sort(function(a,b){return b.score-a.score;});
  moteurs=absolus.concat(normaux);

  var top3=moteurs.slice(0,CONFIG.NB_PRIORITES);
  var semaine=getNumeroSemaine(today);
  var delaiMoyen=Math.round(moteurs.reduce(function(s,m){return s+m.delai;},0)/moteurs.length);
  return {top3:top3,moteurs:moteurs,stockByPower:stockByPower,
          semaine:semaine,dateStr:formatDate(today),delaiMoyen:delaiMoyen,erreur:null};
}


// ============================================================
//  GÉNÉRATION SUJET
// ============================================================
function genererSujet(result) {
  return "Priorisation réparations moteurs électriques — Semaine " + result.semaine + " / " + new Date().getFullYear();
}


// ============================================================
//  GÉNÉRATION CORPS TEXTE (fallback)
// ============================================================
function genererCorpsMail(result) {
  var lignes=[], sep="─────────────────────────────────────────";
  lignes.push("Bonjour,");
  lignes.push("");
  lignes.push("Suite a l'analyse hebdomadaire de l'etat de reserve des moteurs electriques et de l'historique d'envoi au reparation, veuillez trouver ci-dessous les " + CONFIG.NB_PRIORITES + " moteurs a prioriser pour la semaine " + result.semaine + ".");
  lignes.push(""); lignes.push(sep);
  lignes.push("  MOTEURS A PRIORISER — SEMAINE " + result.semaine);
  lignes.push(sep); lignes.push("");
  var rangs=["1ere PRIORITE","2eme PRIORITE","3eme PRIORITE"];
  var icons=["[P1]","[P2]","[P3]"];
  result.top3.forEach(function(m,i){
    var specs=m.puissance+" kW"+(m.tension?" / "+m.tension+" V":"")+(m.vitesse?" / "+m.vitesse+" tr/min":"");
    lignes.push(icons[i]+" "+rangs[i]+" : Matricule : "+m.matricule+" — "+specs);
    if(m.avis) lignes.push("   ├ N Avis/OT         : "+m.avis);
    lignes.push("   ├ Delai en reparation : "+m.delai+" jours (depuis le "+m.dateStr+")");
    lignes.push("   └ Anomalie            : "+(m.anomalie||"Non renseignee"));
    lignes.push("");
  });
  lignes.push(sep); lignes.push("");
  lignes.push("Merci de bien vouloir tenir compte de ces priorites.");
  lignes.push(""); lignes.push("Cordialement,"); lignes.push("");
  lignes.push("Bureau de methode Daoui");
  lignes.push("Gestion des interchangeables");
  lignes.push("Marouane ELAMRAOUI");
  lignes.push("Tel : 06 61 32 37 84");
  lignes.push("Cisco : 8103388");
  lignes.push("E-mail : m.elamraoui@ocpgroup.ma");
  return lignes.join("\n");
}


// ============================================================
//  GÉNÉRATION CORPS HTML
// ============================================================
function genererCorpsMailHTML(result) {
  var puissancesTop3 = [];
  result.top3.forEach(function(m) {
    var pw = parseFloat(m.puissance);
    var found = false;
    for (var k=0; k<puissancesTop3.length; k++) { if (puissancesTop3[k]===pw) { found=true; break; } }
    if (!found) puissancesTop3.push(pw);
  });
  var moteursTableau = result.moteurs.filter(function(m) {
    var pw = parseFloat(m.puissance);
    for (var k=0; k<puissancesTop3.length; k++) { if (puissancesTop3[k]===pw) return true; }
    return false;
  });

  var rangs  = ["1\u00e8re PRIORIT\u00c9","2\u00e8me PRIORIT\u00c9","3\u00e8me PRIORIT\u00c9"];
  var icons  = ["[P1]","[P2]","[P3]"];
  var colors = ["#f97316","#3b82f6","#10b981"];
  var tdBorder = "border:1px solid #e2e8f0;";

  var blocTop3 = result.top3.map(function(m, i) {
    var specs = m.puissance + " kW" + (m.tension?" / "+m.tension+" V":"") + (m.vitesse?" / "+m.vitesse+" tr/min":"");
    return "<tr>" +
      "<td style='" + tdBorder + "padding:6px 10px;font-weight:700;color:"+colors[i]+";white-space:nowrap'>" + icons[i] + " " + rangs[i] + "</td>" +
      "<td style='" + tdBorder + "padding:6px 10px;font-weight:600;white-space:nowrap'>" + escapeHtml(m.matricule) + "</td>" +
      "<td style='" + tdBorder + "padding:6px 10px;white-space:nowrap'>" + escapeHtml(specs) + "</td>" +
      "<td style='" + tdBorder + "padding:6px 10px;white-space:nowrap'>" + (m.avis ? escapeHtml(m.avis) : "<span style='color:#64748b'>\u2014</span>") + "</td>" +
      "<td style='" + tdBorder + "padding:6px 10px;white-space:nowrap'>" + m.delai + " jours<br><span style='font-size:11px;color:#64748b'>depuis le " + m.dateStr + "</span></td>" +
      "<td style='" + tdBorder + "padding:6px 10px;white-space:nowrap'>" + escapeHtml(m.anomalie || "Non renseign\u00e9e") + "</td>" +
    "</tr>";
  }).join("");

  var lignesTableau = moteursTableau.filter(function(m) {
    return m.avis && m.avis.trim() !== "";
  }).map(function(m) {
    return "<tr>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap'>" + escapeHtml(m.dateEnvoi || "\u2014") + "</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap'>" + escapeHtml(m.avis) + "</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap;font-weight:600'>" + escapeHtml(m.matricule) + "</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap;font-weight:600'>" + m.puissance + " kW</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap'>" + escapeHtml(m.tension || "\u2014") + "</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap'>" + escapeHtml(m.vitesse || "\u2014") + "</td>" +
      "<td style='" + tdBorder + "padding:5px 8px;white-space:nowrap'>" + escapeHtml(m.anomalie || "Non renseign\u00e9e") + "</td>" +
    "</tr>";
  }).join("");

  var thStyle = "padding:6px 10px;text-align:left;font-size:10px;color:#fff;text-transform:uppercase;letter-spacing:1px;border:1px solid #e2e8f0;white-space:nowrap;";

  var html =
    "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>" +
    "<body style='font-family:Arial,sans-serif;font-size:14px;color:#1e293b;max-width:900px;margin:0 auto;padding:20px'>" +
    "<p style='margin-bottom:16px;line-height:1.7'>Bonjour,</p>" +
    "<p style='margin-bottom:24px;line-height:1.7'>Suite \u00e0 l\u2019analyse hebdomadaire de l\u2019\u00e9tat de r\u00e9serve des moteurs \u00e9lectriques et de l\u2019historique d\u2019envoi au r\u00e9paration, veuillez trouver ci-dessous les " + result.top3.length + " moteurs \u00e0 prioriser pour la semaine " + result.semaine + ".</p>" +
    "<div style='display:inline-block'>" +
    "<table cellspacing='0' cellpadding='0' style='border-collapse:collapse;margin-bottom:0;table-layout:auto'>" +
      "<thead><tr style='background:#1e293b'>" +
        "<th style='" + thStyle + "'>Priorit\u00e9</th>" +
        "<th style='" + thStyle + "'>Matricule</th>" +
        "<th style='" + thStyle + "'>Puissance / Tension / Vitesse</th>" +
        "<th style='" + thStyle + "'>N\u00b0 Avis/OT</th>" +
        "<th style='" + thStyle + "'>D\u00e9lai</th>" +
        "<th style='" + thStyle + "'>Anomalie</th>" +
      "</tr></thead>" +
      "<tbody>" + blocTop3 + "</tbody>" +
    "</table></div><br>" +
    (moteursTableau.length > 0 ?
      "<p style='margin-bottom:16px;line-height:1.7'>Par ailleurs, vous trouverez ci-dessous le tableau r\u00e9capitulatif des moteurs de ces puissances d\u00e9j\u00e0 envoy\u00e9s en r\u00e9paration :</p>" +
      "<div style='display:inline-block'>" +
      "<table cellspacing='0' cellpadding='0' style='border-collapse:collapse;margin-bottom:0;table-layout:auto'>" +
        "<thead><tr style='background:#1e293b'>" +
          "<th style='" + thStyle + "'>Date d\u2019envoi</th>" +
          "<th style='" + thStyle + "'>N\u00b0 Avis/OT</th>" +
          "<th style='" + thStyle + "'>Matricule</th>" +
          "<th style='" + thStyle + "'>Puissance</th>" +
          "<th style='" + thStyle + "'>Tension</th>" +
          "<th style='" + thStyle + "'>Vitesse</th>" +
          "<th style='" + thStyle + "'>Anomalie</th>" +
        "</tr></thead>" +
        "<tbody>" + lignesTableau + "</tbody>" +
      "</table></div><br>"
    : "") +
    "<p style='line-height:1.7;margin-bottom:24px'>Merci de bien vouloir tenir compte de ces priorit\u00e9s et de nous informer de tout avancement sur ces moteurs.</p>" +
    "<p style='margin-bottom:4px'>Cordialement,</p>" +
    "<div style='margin-top:16px;padding:16px;background:#f8fafc;border-left:4px solid #f97316;border-radius:0 6px 6px 0'>" +
      "<div style='font-weight:700;font-size:14px'>Bureau de m\u00e9thode Daoui</div>" +
      "<div style='color:#64748b;font-size:13px'>Gestion des interchangeables</div>" +
      "<div style='font-weight:600;margin-top:6px'>Marouane ELAMRAOUI</div>" +
      "<div style='font-size:13px;color:#475569;margin-top:4px'>" +
        "T\u00e9l : 06 61 32 37 84 &nbsp;&middot;&nbsp; Cisco : 8103388<br>" +
        "E-mail : <a href='mailto:m.elamraoui@ocpgroup.ma' style='color:#f97316'>m.elamraoui@ocpgroup.ma</a>" +
      "</div>" +
    "</div>" +
    "</body></html>";

  return html;
}


// ============================================================
//  SIDEBAR
// ============================================================
function afficherSidebar(result) {
  var rows = result.top3.map(function(m, i) {
    var rangs  = ["[P1] Priorit\u00e9 1","[P2] Priorit\u00e9 2","[P3] Priorit\u00e9 3"];
    var colors = ["#f97316","#3b82f6","#10b981"];
    return "<tr style='border-bottom:1px solid #1e293b'>" +
      "<td style='padding:10px 8px;font-weight:700;color:"+colors[i]+"'>" + rangs[i] + "</td>" +
      "<td style='padding:10px 8px;font-weight:600'>" + m.matricule + "</td>" +
      "<td style='padding:10px 8px'><span style='background:#1e3a5f;color:#93c5fd;padding:2px 8px;border-radius:4px;font-size:11px'>" +
        m.puissance+" kW"+(m.tension?" / "+m.tension+"V":"")+(m.vitesse?" / "+m.vitesse+"rpm":"") +
      "</span></td>" +
      "<td style='padding:10px 8px;color:"+(m.delai>60?"#ef4444":m.delai>30?"#f59e0b":"#e2e8f0")+";font-weight:600'>" + m.delai + " j</td>" +
      "<td style='padding:10px 8px;color:"+(m.prioriteAbsolue?"#ef4444":m.stockDispo===1?"#f59e0b":"#10b981")+"'>" +
        (m.prioriteAbsolue?"⚠ URGENT":m.stockDispo+" dispo") +
      "</td></tr>";
  }).join("");

  var html = HtmlService.createHtmlOutput(
    "<style>" +
    "body{font-family:'Google Sans',sans-serif;background:#0a0e14;color:#e2e8f0;padding:16px;margin:0}" +
    "h2{font-size:15px;margin:0 0 4px;color:#f97316}" +
    ".sub{font-size:11px;color:#64748b;margin-bottom:16px;font-family:monospace}" +
    "table{width:100%;border-collapse:collapse;font-size:12px}" +
    "th{font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:1px;padding:6px 8px;text-align:left;border-bottom:1px solid #1e293b}" +
    ".stat{display:flex;gap:8px;margin-bottom:16px}" +
    ".stat-box{flex:1;background:#111820;border:1px solid #1e293b;border-radius:8px;padding:10px;text-align:center}" +
    ".stat-num{font-size:22px;font-weight:800}" +
    ".stat-label{font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:1px}" +
    ".btn{width:100%;margin-top:10px;padding:11px;border:none;border-radius:8px;font-weight:700;font-size:13px;cursor:pointer;font-family:'Google Sans',sans-serif}" +
    ".btn-orange{background:#f97316;color:#000}.btn-orange:hover{background:#ea6c0a}" +
    ".btn-gray{background:#1e293b;color:#e2e8f0;border:1px solid #334155}.btn-gray:hover{background:#263548}" +
    ".msg{display:none;margin-top:10px;padding:9px 12px;border-radius:8px;font-size:12px;font-family:monospace}" +
    ".msg.ok{display:block;background:rgba(16,185,129,0.1);color:#10b981;border:1px solid rgba(16,185,129,0.3)}" +
    ".msg.err{display:block;background:rgba(239,68,68,0.1);color:#ef4444;border:1px solid rgba(239,68,68,0.3)}" +
    "</style>" +
    "<h2>⚡ Priorisation Bobinage</h2>" +
    "<div class='sub'>Semaine " + result.semaine + " — " + result.dateStr + "</div>" +
    "<div class='stat'>" +
      "<div class='stat-box'><div class='stat-num' style='color:#3b82f6'>" + result.moteurs.length + "</div><div class='stat-label'>En réparation</div></div>" +
      "<div class='stat-box'><div class='stat-num' style='color:#f97316'>" + (result.moteurs[0]?result.moteurs[0].delai:0) + "</div><div class='stat-label'>Délai max (j)</div></div>" +
      "<div class='stat-box'><div class='stat-num' style='color:#10b981'>" + result.delaiMoyen + "</div><div class='stat-label'>Délai moyen</div></div>" +
    "</div>" +
    "<table><thead><tr><th>Rang</th><th>Matricule</th><th>Puissance</th><th>Délai</th><th>Stock</th></tr></thead>" +
    "<tbody>" + rows + "</tbody></table>" +
    "<button class='btn btn-orange' onclick='envoyer()'>✉️ Envoyer le mail au bobinage</button>" +
    "<button class='btn btn-gray' onclick='google.script.run.previsualiserMail()'>👁 Prévisualiser</button>" +
    "<div class='msg' id='msg'></div>" +
    "<script>" +
    "function envoyer(){showMsg('⏳ Envoi en cours...',true);google.script.run.withSuccessHandler(function(r){showMsg(r,true);}).withFailureHandler(function(e){showMsg('❌ '+e.message,false);}).envoyerMailDepuisSidebar();}" +
    "function showMsg(t,ok){var el=document.getElementById('msg');el.textContent=t;el.className='msg '+(ok?'ok':'err');}" +
    "<\/script>"
  ).setTitle("Priorisation Bobinage").setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ============================================================
//  LOGGER
// ============================================================
function loggerEnvoi(result) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Historique envois");
  if (!sheet) {
    sheet = ss.insertSheet("Historique envois");
    sheet.appendRow(["Date envoi","Semaine","P1 Matricule","P1 Puissance","P1 Délai",
                     "P2 Matricule","P2 Puissance","P2 Délai","P3 Matricule","P3 Puissance","P3 Délai"]);
    sheet.getRange(1,1,1,11).setFontWeight("bold").setBackground("#f97316").setFontColor("#000000");
  }
  var t = result.top3;
  sheet.appendRow([new Date(), result.semaine,
    t[0]?t[0].matricule:"", t[0]?t[0].puissance:"", t[0]?t[0].delai:"",
    t[1]?t[1].matricule:"", t[1]?t[1].puissance:"", t[1]?t[1].delai:"",
    t[2]?t[2].matricule:"", t[2]?t[2].puissance:"", t[2]?t[2].delai:""]);
}


// ============================================================
//  UTILITAIRES
// ============================================================
function trouverColonne(headers, nom) {
  var n = nom.trim().toLowerCase();
  for (var i=0;i<headers.length;i++) { if(headers[i].toLowerCase().trim()===n) return i; }
  for (var i=0;i<headers.length;i++) {
    if(headers[i].toLowerCase().trim().indexOf(n)>=0||n.indexOf(headers[i].toLowerCase().trim())>=0) return i;
  }
  return -1;
}
function normaliser(str) {
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();
}
function contient(str, mots) {
  var s=normaliser(str);
  for(var i=0;i<mots.length;i++){if(s.indexOf(normaliser(mots[i]))>=0)return true;}
  return false;
}
function parseDate(str) {
  var m=str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if(m) return new Date(parseInt(m[3]),parseInt(m[2])-1,parseInt(m[1]));
  m=str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if(m) return new Date(parseInt(m[1]),parseInt(m[2])-1,parseInt(m[3]));
  return null;
}
function formatDate(date) {
  var d=date.getDate().toString().padStart(2,"0");
  var m=(date.getMonth()+1).toString().padStart(2,"0");
  return d+"/"+m+"/"+date.getFullYear();
}
function getNumeroSemaine(d) {
  var date=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  date.setUTCDate(date.getUTCDate()+4-(date.getUTCDay()||7));
  var yearStart=new Date(Date.UTC(date.getUTCFullYear(),0,1));
  return Math.ceil((((date-yearStart)/86400000)+1)/7);
}
function escapeHtml(str) {
  return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
