const SHEET_ID    = '1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4';
const SHEET_NAME  = 'Demandes';

// ── Configuration OCP Exchange (EWS) ─────────────────────────
const OCP_EMAIL = 'm.elamraoui@ocpgroup.ma';
const EWS_URL   = 'https://owa.ocpgroup.ma/EWS/Exchange.asmx';
function getOcpPassword() {
  return PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD') || '';
}

function sendEmailOCP(to, subject, body, options) {
  const toList = Array.isArray(to) ? to : [to];
  const cc = options && options.cc
    ? (Array.isArray(options.cc) ? options.cc : options.cc.split(',').map(function(e){ return e.trim(); }).filter(Boolean))
    : [];
  const toRecipients = toList.map(function(e){
    return '<t:Mailbox><t:EmailAddress>' + e + '</t:EmailAddress></t:Mailbox>';
  }).join('');
  const ccBlock = cc.length
    ? '<t:CcRecipients>' + cc.map(function(e){
        return '<t:Mailbox><t:EmailAddress>' + e + '</t:EmailAddress></t:Mailbox>';
      }).join('') + '</t:CcRecipients>'
    : '';
  const fromName = (options && options.name) ? options.name : 'Maintenance Analytics';
  const bodyType = (options && options.htmlBody) ? 'HTML' : 'Text';
  const bodyContent = (options && options.htmlBody) ? options.htmlBody : (body || '');
  const soap = '<?xml version="1.0" encoding="utf-8"?>'
    + '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'
    + ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"'
    + ' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">'
    + '<soap:Body><m:CreateItem MessageDisposition="SendAndSaveCopy">'
    + '<m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>'
    + '<m:Items><t:Message>'
    + '<t:Subject>' + subject + '</t:Subject>'
    + '<t:Body BodyType="' + bodyType + '">' + bodyContent + '</t:Body>'
    + '<t:From><t:Mailbox><t:Name>' + fromName + '</t:Name>'
    + '<t:EmailAddress>' + OCP_EMAIL + '</t:EmailAddress></t:Mailbox></t:From>'
    + '<t:ToRecipients>' + toRecipients + '</t:ToRecipients>'
    + ccBlock
    + '</t:Message></m:Items>'
    + '</m:CreateItem></m:Body></soap:Envelope>';
  const credentials = Utilities.base64Encode(OCP_EMAIL + ':' + getOcpPassword());
  const response = UrlFetchApp.fetch(EWS_URL, {
    method: 'post',
    contentType: 'text/xml; charset=utf-8',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'SOAPAction': 'http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem'
    },
    payload: soap,
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200 || response.getContentText().indexOf('NoError') === -1) {
    throw new Error('EWS send failed (' + response.getResponseCode() + '): ' + response.getContentText().substring(0, 300));
  }
}
const SHEET_INTERCH    = 'Demande des intercheable'; // feuille matricule rechange
const COL_MATRECHANGE  = 20;                         // Colonne T (1-based)
const SHEET_USERS      = 'Users';
const USERS_HEADERS    = ['id','nom','prenom','tel','email','fonction','dateAjout'];
const EMAIL_ADMIN = 'm.elamraoui@ocpgroup.ma';
const EMAIL_DEMANDEUR_FALLBACK = 'mar.elamraoui@gmail.com';
const SENDER_NAME = 'Bureau de méthode Daoui - Section Interchangeable';

const HEADERS = [
  'id','type','installation','objetTechnique','puissance','tension','vitesse',        // A-G
  'anomalie','matricule','demandeur','demandeurEmail','dateDemande','statut','etatReparation','justifRefus', // H-O
  '',                // P — réservé
  'remarqueAnomalie' // Q
];

// ══════════════════════════════════════════════
// LECTURE + ACTIONS GET
// ══════════════════════════════════════════════
function doGet(e) {
  try {

    // ── Installations + Demandeurs ────────────────────────────
    if (e && e.parameter && e.parameter.action === 'getInstallations') {
      var iSS    = SpreadsheetApp.openById(SHEET_ID);
      var iSheet = iSS.getSheetByName('Installation');
      if (!iSheet) return jsonResponse({ success: false, error: 'Feuille Installation introuvable' });
      var iVals  = iSheet.getDataRange().getValues();
      var installations = [];
      var demandeurs    = [];
      var instSet = {};
      var demSet  = {};
      for (var i = 1; i < iVals.length; i++) {
        var inst  = String(iVals[i][0] || '').trim();
        var nom   = String(iVals[i][1] || '').trim();
        var email = String(iVals[i][2] || '').trim();
        if (inst && !instSet[inst]) { installations.push(inst); instSet[inst] = true; }
        if (nom  && !demSet[nom])   { demandeurs.push({ nom: nom, email: email }); demSet[nom] = true; }
      }
      return jsonResponse({ success: true, installations: installations, demandeurs: demandeurs });
    }

    // ── Confirmation PDR ─────────────────────────────────────
    if (e && e.parameter && e.parameter.action === 'updatePDR') {
      var p        = e.parameter;
      var pdrSS    = SpreadsheetApp.openById(p.fileId);
      var pdrSheet = pdrSS.getSheetByName(p.sheetName);
      var values   = pdrSheet.getDataRange().getValues();

      var targetRow = -1;
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === String(p.ordre).trim()) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow === -1) {
        return jsonResponse({ success: false, error: 'Ordre ' + p.ordre + ' non trouvé' });
      }

      pdrSheet.getRange(targetRow, 20).setValue(p.disponibilite); // Colonne T
      pdrSheet.getRange(targetRow, 21).setValue(p.justification); // Colonne U
      pdrSheet.getRange(targetRow, 29).setValue(p.delai);         // Colonne AC

      return jsonResponse({ success: true, row: targetRow });
    }

    // ── Réalisation OT : écrit 'Fait' / 'NFait' en colonne O ─
    if (e && e.parameter && e.parameter.action === 'updateRealisation') {
      var rp     = e.parameter;
      var rSS    = SpreadsheetApp.openById(rp.fileId);
      var rSheet = rSS.getSheetByName(rp.sheetName);
      var rVals  = rSheet.getDataRange().getValues();
      var rRow   = -1;
      for (var ri = 1; ri < rVals.length; ri++) {
        if (String(rVals[ri][0]).trim() === String(rp.ordre).trim()) {
          rRow = ri + 1; break;
        }
      }
      if (rRow === -1) return jsonResponse({ success: false, error: 'Ordre non trouvé' });
      rSheet.getRange(rRow, 15).setValue(rp.valeur); // Colonne O

      return jsonResponse({ success: true, row: rRow });
    }

    // ── Sauvegarder abonnement Web Push ─────────────────────────
    if (e && e.parameter && e.parameter.action === 'saveSubscription') {
      var sp = e.parameter;
      var sSS = SpreadsheetApp.openById(SHEET_ID);
      var sSheet = sSS.getSheetByName('Subscriptions');
      if (!sSheet) {
        sSheet = sSS.insertSheet('Subscriptions');
        sSheet.appendRow(['endpoint', 'profile', 'subscription', 'updatedAt']);
      }
      var sVals = sSheet.getDataRange().getValues();
      var existingRow = -1;
      for (var si = 1; si < sVals.length; si++) {
        if (String(sVals[si][0]).trim() === String(sp.endpoint).trim()) {
          existingRow = si + 1; break;
        }
      }
      if (existingRow > 0) {
        sSheet.getRange(existingRow, 2).setValue(sp.profile || '');
        sSheet.getRange(existingRow, 3).setValue(sp.subscription || '');
        sSheet.getRange(existingRow, 4).setValue(new Date().toISOString());
      } else {
        sSheet.appendRow([sp.endpoint || '', sp.profile || '', sp.subscription || '', new Date().toISOString()]);
      }
      return jsonResponse({ success: true });
    }

    // ── Récupérer abonnements Web Push par profil ────────────────
    if (e && e.parameter && e.parameter.action === 'getSubscriptions') {
      var gp = e.parameter;
      var gsSS = SpreadsheetApp.openById(SHEET_ID);
      var gsSheet = gsSS.getSheetByName('Subscriptions');
      if (!gsSheet) return jsonResponse([]);
      var gsVals = gsSheet.getDataRange().getValues();
      var subResult = [];
      for (var gsi = 1; gsi < gsVals.length; gsi++) {
        if (String(gsVals[gsi][1]).trim() === String(gp.profile || 'admin').trim()) {
          subResult.push({ endpoint: gsVals[gsi][0], profile: gsVals[gsi][1], subscription: gsVals[gsi][2] });
        }
      }
      return jsonResponse(subResult);
    }

    // ── Log notification (broadcast à tous les devices) ─────────
    if (e && e.parameter && e.parameter.action === 'logNotif') {
      var lp = e.parameter;
      var lSS = SpreadsheetApp.openById(SHEET_ID);
      var lSheet = lSS.getSheetByName('Notifications');
      if (!lSheet) {
        lSheet = lSS.insertSheet('Notifications');
        lSheet.appendRow(['timestamp', 'id', 'title', 'body', 'targetProfile']);
      }
      var notifId = lp.notifId || (new Date().getTime() + '_' + Math.random().toString(36).slice(2,7));
      lSheet.appendRow([new Date().toISOString(), notifId, lp.title || '', lp.body || '', lp.targetProfile || 'admin']);
      return jsonResponse({ success: true });
    }

    // ── Récupérer les notifications récentes (dernières 24h) ────
    if (e && e.parameter && e.parameter.action === 'getNotifs') {
      var gSS = SpreadsheetApp.openById(SHEET_ID);
      var gSheet = gSS.getSheetByName('Notifications');
      if (!gSheet) return jsonResponse([]);
      var gVals = gSheet.getDataRange().getValues();
      var since = Date.now() - 24 * 60 * 60 * 1000;
      var notifResult = [];
      for (var gi = 1; gi < gVals.length; gi++) {
        var ts = new Date(gVals[gi][0]).getTime();
        if (ts > since) {
          notifResult.push({ timestamp: gVals[gi][0], id: String(gVals[gi][1]), title: gVals[gi][2], body: gVals[gi][3], targetProfile: gVals[gi][4] });
        }
      }
      return jsonResponse(notifResult);
    }

    // ── Marquer un arrêt comme reporté — colonne S du fichier Planning ──
    if (e && e.parameter && e.parameter.action === 'markArretReporte') {
      var mp   = e.parameter;
      var mSS  = SpreadsheetApp.openById('1EBACM8ou8B_9fmExToUKsMCvHL27hiwU2D0yZ_gQGOA');
      var mSh  = mSS.getSheetByName('Planning des arrets');
      if (!mSh) return jsonResponse({ success: false, error: 'Feuille introuvable' });

      var mVals   = mSh.getDataRange().getValues();
      var mHdrs   = mVals[0].map(function(h){ return String(h).trim().toLowerCase(); });

      // Détection flexible des colonnes date et installation
      var dateNames    = ['start date','date début','date de début','date','début'];
      var installNames = ['installation','equipement','équipement','arrêt','arret'];
      var iDc = -1, iIc = -1;
      for (var dn = 0; dn < dateNames.length; dn++)    { var f1 = mHdrs.indexOf(dateNames[dn]);    if (f1 >= 0){ iDc = f1; break; } }
      for (var dn2 = 0; dn2 < installNames.length; dn2++){ var f2 = mHdrs.indexOf(installNames[dn2]); if (f2 >= 0){ iIc = f2; break; } }
      if (iDc < 0 || iIc < 0) return jsonResponse({ success: false, error: 'Colonnes date/installation introuvables' });

      var targetDate    = String(mp.date || '').trim();   // YYYY-MM-DD
      var targetInstall = String(mp.installation || '').trim().toLowerCase();

      var matchRow = -1;
      for (var mi = 1; mi < mVals.length; mi++) {
        var cv = mVals[mi][iDc];
        var cellDate = '';
        if (cv instanceof Date) {
          var cy = cv.getFullYear(), cm2 = cv.getMonth()+1, cd = cv.getDate();
          cellDate = cy + '-' + (cm2<10?'0':'') + cm2 + '-' + (cd<10?'0':'') + cd;
        } else {
          var ps = String(cv).trim().split(/[\/\-\.]/);
          if (ps.length === 3) {
            if (parseInt(ps[2]) >= 2020) cellDate = ps[2]+'-'+(ps[1].length<2?'0':'')+ps[1]+'-'+(ps[0].length<2?'0':'')+ps[0];
            else if (parseInt(ps[0]) >= 2020) cellDate = ps[0]+'-'+(ps[1].length<2?'0':'')+ps[1]+'-'+(ps[2].length<2?'0':'')+ps[2];
          }
        }
        var cellInstall = String(mVals[mi][iIc]).trim().toLowerCase();
        if (cellDate === targetDate && cellInstall === targetInstall) { matchRow = mi + 1; break; }
      }

      if (matchRow < 0) return jsonResponse({ success: false, error: 'Arrêt non trouvé', date: targetDate, install: targetInstall });

      var now2 = new Date();
      var reporteVal = 'Reporté le ' + now2.getDate()+'/'+(now2.getMonth()+1)+'/'+now2.getFullYear();
      mSh.getRange(matchRow, 19).setValue(reporteVal); // Colonne S (1-based index 19)
      return jsonResponse({ success: true, row: matchRow, value: reporteVal });
    }

    // ── Mot de passe OCP Exchange ────────────────────────────
    if (e && e.parameter && e.parameter.action === 'setOcpPassword') {
      var newPwd = e.parameter.password || '';
      if (!newPwd) return jsonResponse({ success: false, error: 'Mot de passe vide.' });
      PropertiesService.getScriptProperties().setProperty('OCP_PASSWORD', newPwd);
      return jsonResponse({ success: true });
    }
    if (e && e.parameter && e.parameter.action === 'getOcpPasswordStatus') {
      var pwd = PropertiesService.getScriptProperties().getProperty('OCP_PASSWORD');
      return jsonResponse({ success: true, isSet: !!(pwd && pwd.length > 0) });
    }

    // ── Codes d'accès : admin, appro, exec, cm ───────────────
    if (e && e.parameter && e.parameter.action === 'updateCode') {
      var cp        = e.parameter;
      var codeSS    = SpreadsheetApp.openById(cp.fileId);
      var codeSheet = codeSS.getSheetByName('Code');
      if (!codeSheet) return jsonResponse({ success: false, error: 'Feuille Code introuvable' });
      var codeVals  = codeSheet.getDataRange().getValues();
      var updated   = [];

      var codeMap = {
        admin_code: cp.adminCode || '',
        appro_code: cp.approCode || '',
        exec_code:  cp.execCode  || '',
        cm_code:    cp.cmCode    || ''
      };

      for (var ci = 1; ci < codeVals.length; ci++) {
        var cKey = String(codeVals[ci][0]).trim();
        if (codeMap[cKey] !== undefined && codeMap[cKey] !== '') {
          codeSheet.getRange(ci + 1, 2).setValue(codeMap[cKey]);
          updated.push(cKey);
        }
      }
      return jsonResponse({ success: true, updated: updated });
    }

    // ── Matricule de rechange + email ─────────────────────────
    if (e && e.parameter && e.parameter.action === 'updateMatriculeRechange') {
      var mp        = e.parameter;
      var mSS       = SpreadsheetApp.openById(mp.fileId);
      var mRow      = -1;

      // 1. Chercher d'abord dans 'Demande des intercheable'
      var mSheet = mSS.getSheetByName(SHEET_INTERCH);
      if (mSheet) {
        var mVals = mSheet.getDataRange().getValues();
        for (var mi = 1; mi < mVals.length; mi++) {
          if (String(mVals[mi][0]).trim() === String(mp.demandeId).trim()) {
            mSheet.getRange(mi + 1, COL_MATRECHANGE).setValue(mp.matriculeRechange);
            mRow = mi + 1;
            break;
          }
        }
      }

      // 2. Si introuvable dans 'Demande des intercheable', écrire dans 'Demandes' (source de vérité)
      if (mRow === -1) {
        var dSheet = mSS.getSheetByName(SHEET_NAME);
        if (dSheet) {
          var dVals = dSheet.getDataRange().getValues();
          for (var di = 1; di < dVals.length; di++) {
            if (String(dVals[di][0]).trim() === String(mp.demandeId).trim()) {
              dSheet.getRange(di + 1, COL_MATRECHANGE).setValue(mp.matriculeRechange);
              mRow = di + 1;
              Logger.log('Matricule écrit dans Demandes, ligne ' + mRow);
              break;
            }
          }
        }
      }

      if (mRow === -1) {
        Logger.log('ERREUR: demandeId introuvable dans les deux feuilles : ' + mp.demandeId);
        // Créer une nouvelle ligne dans 'Demandes' pour ne pas perdre la valeur
        var dSheet2 = mSS.getSheetByName(SHEET_NAME);
        if (dSheet2) {
          dSheet2.appendRow([mp.demandeId, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', mp.matriculeRechange]);
          mRow = dSheet2.getLastRow();
          Logger.log('Nouvelle ligne créée en row ' + mRow);
        }
      }

      // Envoi email HTML
      if (mp.sendEmail === 'true' && mp.emailTo) {
        try {
          var typeLabel = mp.typeDemande === 'changement' ? '🔄 Changement'
                        : mp.typeDemande === 'pose'       ? '➕ Pose'
                        : mp.typeDemande;

          var sujet = '[Contrôle Matériel] Rechange saisi — ' + mp.demandeId;
          var corps =
            '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;">' +
              '<div style="background:#1d4ed8;color:white;padding:20px 24px;">' +
                '<h2 style="margin:0;font-size:18px;">Moteur de rechange renseigné</h2>' +
                '<p style="margin:6px 0 0;opacity:.85;font-size:13px;">' + mp.demandeId + ' — ' + typeLabel + '</p>' +
              '</div>' +
              '<div style="padding:20px 24px;">' +
                '<p style="font-size:14px;color:#374151;">Le Service <b>Contrôle Matériel</b> vient de renseigner le matricule du moteur de rechange :</p>' +
                '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:12px;">' +
                  '<tr style="background:#f9fafb;"><td style="padding:8px 12px;color:#6b7280;width:42%;">N° Demande</td>' +
                    '<td style="padding:8px 12px;font-weight:700;font-family:monospace;">' + mp.demandeId + '</td></tr>' +
                  '<tr><td style="padding:8px 12px;color:#6b7280;">Type</td>' +
                    '<td style="padding:8px 12px;font-weight:600;">' + typeLabel + '</td></tr>' +
                  '<tr style="background:#f9fafb;"><td style="padding:8px 12px;color:#6b7280;">Demandeur</td>' +
                    '<td style="padding:8px 12px;font-weight:600;">' + (mp.demandeur || '—') + '</td></tr>' +
                  '<tr><td style="padding:8px 12px;color:#6b7280;">Installation</td>' +
                    '<td style="padding:8px 12px;font-weight:600;">' + (mp.installation || '—') + '</td></tr>' +
                  '<tr style="background:#f9fafb;"><td style="padding:8px 12px;color:#6b7280;">Objet technique</td>' +
                    '<td style="padding:8px 12px;font-weight:600;">' + (mp.objetTechnique || '—') + '</td></tr>' +
                  '<tr><td style="padding:8px 12px;color:#6b7280;">Puissance</td>' +
                    '<td style="padding:8px 12px;font-weight:600;">' + (mp.puissance || '—') + ' kW</td></tr>' +
                  '<tr style="background:#ecfdf5;">' +
                    '<td style="padding:10px 12px;color:#065f46;font-weight:700;">✅ Matricule rechange</td>' +
                    '<td style="padding:10px 12px;font-weight:800;color:#059669;font-size:15px;font-family:monospace;">' + mp.matriculeRechange + '</td></tr>' +
                '</table>' +
              '</div>' +
              '<div style="background:#f1f5f9;padding:12px 24px;font-size:12px;color:#64748b;">' + SENDER_NAME + '</div>' +
            '</div>';

          var destCM = getUserEmails(['Interchangeable électrique']);
          if (destCM) {
            var destCMList = destCM.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
            var toCM = destCMList[0];
            var ccCMRest = destCMList.slice(1).join(',');
            var optsCM = { htmlBody: corps, name: SENDER_NAME };
            if (ccCMRest) optsCM.cc = ccCMRest;
            sendEmailOCP(toCM, sujet, '', optsCM);
          }
        } catch(mailErr) {
          Logger.log('Erreur envoi email CM : ' + mailErr.toString());
        }
      }

      return jsonResponse({ success: true, row: mRow, emailSent: mp.sendEmail === 'true' });
    }

    // ── Lire la liste des utilisateurs ───────────────────────
    if (e && e.parameter && e.parameter.action === 'getUsers') {
      var ss2 = SpreadsheetApp.openById(SHEET_ID);
      var us = ss2.getSheetByName(SHEET_USERS);
      if (!us || us.getLastRow() < 1) return jsonResponse([]);
      var uVals = us.getDataRange().getValues();
      // Détecter si la première ligne est un en-tête (contient 'id' ou 'nom')
      var firstCell = String(uVals[0][0]).trim().toLowerCase();
      var hasHeaders = (firstCell === 'id' || firstCell === 'nom' || firstCell === 'prenom');
      var dataRows = hasHeaders ? uVals.slice(1) : uVals;
      var uRows = dataRows.filter(function(row) {
        return row.some(function(cell) { return cell !== '' && cell !== null; });
      }).map(function(row) {
        var obj = {};
        // Mapping par position sur USERS_HEADERS (id,nom,prenom,tel,email,fonction,dateAjout)
        USERS_HEADERS.forEach(function(h, i) { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
        return obj;
      });
      return jsonResponse(uRows);
    }

    // ── Ajouter un utilisateur ────────────────────────────────
    if (e && e.parameter && e.parameter.action === 'addUser') {
      var ap = e.parameter;
      var ss3 = SpreadsheetApp.openById(SHEET_ID);
      var us3 = ss3.getSheetByName(SHEET_USERS);
      if (!us3) {
        us3 = ss3.insertSheet(SHEET_USERS);
        // Nouvelle feuille vide : pas d'en-tête, données dès ligne 1
      }
      var newRow = USERS_HEADERS.map(function(h) { return ap[h] || ''; });
      us3.appendRow(newRow);
      return jsonResponse({ success: true });
    }

    // ── Supprimer un utilisateur ──────────────────────────────
    if (e && e.parameter && e.parameter.action === 'deleteUser') {
      var dp = e.parameter;
      var ss4 = SpreadsheetApp.openById(SHEET_ID);
      var us4 = ss4.getSheetByName(SHEET_USERS);
      if (!us4) return jsonResponse({ success: false, error: 'Feuille Users introuvable' });
      var dVals4 = us4.getDataRange().getValues();
      for (var di4 = 1; di4 < dVals4.length; di4++) {
        if (String(dVals4[di4][0]).trim() === String(dp.id).trim()) {
          us4.deleteRow(di4 + 1);
          return jsonResponse({ success: true });
        }
      }
      return jsonResponse({ success: false, error: 'Utilisateur introuvable' });
    }

    // ── Données moteurs (comportement existant) ───────────────
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return jsonResponse([]);

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    // Charger matricules rechange :
    // 1) depuis 'Demande des intercheable' col T (source principale)
    // 2) fallback : col T de 'Demandes' (si écriture directe)
    var mRechangeMap = {};
    try {
      // Lecture depuis SHEET_INTERCH
      var shI = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INTERCH);
      if (shI) {
        var iVals = shI.getDataRange().getValues();
        for (var k = 1; k < iVals.length; k++) {
          var did = String(iVals[k][0]).trim();
          var mat = iVals[k][COL_MATRECHANGE - 1] ? String(iVals[k][COL_MATRECHANGE - 1]).trim() : '';
          if (did && mat) mRechangeMap[did] = mat;
        }
      }
      // Fallback : lire col T de la feuille Demandes (pour les valeurs écrites directement)
      for (var fd = 1; fd < data.length; fd++) {
        var fid = String(data[fd][0]).trim();
        if (fid && !mRechangeMap[fid]) {
          var fmat = data[fd][COL_MATRECHANGE - 1] ? String(data[fd][COL_MATRECHANGE - 1]).trim() : '';
          if (fmat) mRechangeMap[fid] = fmat;
        }
      }
    } catch(e2) { /* non bloquant */ }

    const rows = data.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      obj.matriculeRechange = mRechangeMap[obj.id] || '';
      return obj;
    });

    return jsonResponse(rows);

  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}


// ══════════════════════════════════════════════
// ÉCRITURE / MISE À JOUR
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'save') {
      saveDemande(payload.demande);
      sendEmailNouvelleDemande(payload.demande);
    } else if (payload.action === 'update') {
      updateDemande(payload.id, payload.updates);
      sendEmailChangementStatut(payload.id, payload.updates);
    }
    return jsonResponse({ success: true });
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// ══════════════════════════════════════════════
// HELPERS SHEET
// ══════════════════════════════════════════════
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const hRow = sheet.getRange(1, 1, 1, HEADERS.length);
    hRow.setValues([HEADERS]);
    hRow.setFontWeight('bold').setBackground('#1d4ed8').setFontColor('white');
  }
  return sheet;
}

function saveDemande(demande) {
  const sheet   = getOrCreateSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row     = headers.map(h => (demande[h] !== undefined && demande[h] !== null) ? demande[h] : '');
  sheet.appendRow(row);
}

function updateDemande(id, updates) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const hdrs  = data[0];
  const idCol = hdrs.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      Object.keys(updates).forEach(key => {
        const col = hdrs.indexOf(key);
        if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(updates[key]);
      });
      break;
    }
  }
}

// ══════════════════════════════════════════════
// EMAIL — NOUVELLE DEMANDE → ADMIN
// ══════════════════════════════════════════════
function sendEmailNouvelleDemande(d) {
  const sujet = `🔧 Nouvelle demande moteur — ${d.id}`;
  const corps = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;">
  <div style="background:#1d4ed8;color:white;padding:20px 24px;">
    <h2 style="margin:0;font-size:18px;">Nouvelle demande de moteur électrique</h2>
    <p style="margin:6px 0 0;opacity:.85;font-size:13px;">
      Demande de <b>${d.type === 'changement' ? 'Changement' : 'Pose'}</b>
      — ${d.puissance} kW · ${d.installation} · ${d.objetTechnique}
    </p>
  </div>
  <div style="padding:20px 24px;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr><td style="padding:7px 0;color:#666;width:40%;">N° Demande</td><td style="font-weight:700;font-family:monospace;">${d.id}</td></tr>
      <tr style="background:#f9fafb;"><td style="padding:7px 8px;color:#666;">Type</td><td style="padding:7px 8px;font-weight:600;">${d.type === 'changement' ? '🔄 Changement' : '➕ Pose'}</td></tr>
      <tr><td style="padding:7px 0;color:#666;">Demandeur</td><td style="font-weight:600;">${d.demandeur}${d.matricule ? ' · Matricule : <b>' + d.matricule + '</b>' : ''}</td></tr>
      <tr style="background:#f9fafb;"><td style="padding:7px 8px;color:#666;">Installation</td><td style="padding:7px 8px;font-weight:600;">${d.installation}</td></tr>
      <tr><td style="padding:7px 0;color:#666;">Objet technique</td><td style="font-weight:600;">${d.objetTechnique}</td></tr>
      <tr style="background:#f9fafb;"><td style="padding:7px 8px;color:#666;">Puissance</td><td style="padding:7px 8px;font-weight:600;">${d.puissance} kW</td></tr>
      <tr><td style="padding:7px 0;color:#666;">Tension</td><td style="font-weight:600;">${d.tension} V</td></tr>
      <tr style="background:#f9fafb;"><td style="padding:7px 8px;color:#666;">Vitesse</td><td style="padding:7px 8px;font-weight:600;">${d.vitesse} tr/min</td></tr>
      ${d.anomalie ? `<tr style="background:#fff7ed;"><td style="padding:7px 8px;color:#d97706;">Anomalie</td><td style="padding:7px 8px;font-weight:700;color:#d97706;">${d.anomalie}</td></tr>` : ''}
      <tr><td style="padding:7px 0;color:#666;">Date</td><td style="font-weight:600;">${new Date(d.dateDemande).toLocaleString('fr-FR')}</td></tr>
    </table>
  </div>
  <div style="background:#f1f5f9;padding:12px 24px;font-size:12px;color:#64748b;">${SENDER_NAME}</div>
</div>`;
  var destNvlle = getUserEmails(['Interchangeable électrique']);
  if (destNvlle) {
    var destNvlleList = destNvlle.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
    var toNvlle = destNvlleList[0];
    var ccNvlleRest = destNvlleList.slice(1).join(',');
    var optsNvlle = { htmlBody: corps, name: SENDER_NAME };
    if (ccNvlleRest) optsNvlle.cc = ccNvlleRest;
    sendEmailOCP(toNvlle, sujet, '', optsNvlle);
  }
}

// Récupère les emails des utilisateurs filtrés par liste de fonctions
// Si fonctionsFilter est vide/null → retourne tous les emails
function getUserEmails(fonctionsFilter) {
  try {
    var us = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_USERS);
    if (!us || us.getLastRow() < 1) return '';
    var vals = us.getDataRange().getValues();
    // Colonnes par position : [id,nom,prenom,tel,email,fonction,dateAjout]
    var emailCol   = 4; // USERS_HEADERS.indexOf('email')
    var fonctCol   = 5; // USERS_HEADERS.indexOf('fonction')
    // Détecter si ligne 1 est un en-tête
    var firstCell  = String(vals[0][0]).trim().toLowerCase();
    var hasHeaders = (firstCell === 'id' || firstCell === 'nom' || firstCell === 'prenom');
    var dataRows   = hasHeaders ? vals.slice(1) : vals;
    var emails = [];
    for (var i = 0; i < dataRows.length; i++) {
      var em   = String(dataRows[i][emailCol] || '').trim();
      var fonc = String(dataRows[i][fonctCol]  || '').trim();
      if (!em || em.indexOf('@') < 0) continue;
      if (fonctionsFilter && fonctionsFilter.length > 0) {
        var match = false;
        for (var f = 0; f < fonctionsFilter.length; f++) {
          if (fonc.toLowerCase() === fonctionsFilter[f].toLowerCase()) { match = true; break; }
        }
        if (!match) continue;
      }
      emails.push(em);
    }
    return emails.join(',');
  } catch(e) {
    Logger.log('getUserEmails error: ' + e.toString());
    return '';
  }
}

function getEmailByDemandeur(nomDemandeur) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Installation');
    if (!sheet) return EMAIL_ADMIN;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      const nom   = String(data[i][1] || '').trim();
      const email = String(data[i][2] || '').trim();
      if (nom.toLowerCase() === nomDemandeur.toLowerCase() && email) return email;
    }
    return EMAIL_ADMIN;
  } catch(e) {
    return EMAIL_ADMIN;
  }
}

// ══════════════════════════════════════════════
// EMAIL — CHANGEMENT STATUT OU ÉTAT MOTEUR → DEMANDEUR
// ══════════════════════════════════════════════
function sendEmailChangementStatut(id, updates) {
  if (!updates.statut && !updates.etatReparation) return;

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const hdrs  = data[0];
  const idCol = hdrs.indexOf('id');
  let demande = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      demande = {};
      hdrs.forEach((h, j) => demande[h] = data[i][j]);
      break;
    }
  }
  if (!demande) return;
  Object.assign(demande, updates);

  const emailDest = demande.demandeurEmail || getEmailByDemandeur(demande.demandeur) || EMAIL_ADMIN;
  const statutLabel = { envoyee: '● Envoyée', approuvee: '✅ Approuvée', refusee: '❌ Refusée' };
  const etatLabel   = { attente: 'En attente de réparation', en_reparation: 'En réparation', repare: '✅ Réparé' };

  let sujet, corps;

  if (updates.etatReparation) {
    const etat = etatLabel[demande.etatReparation] || demande.etatReparation;
    let couleurEtat = '#1d4ed8';
    if (demande.etatReparation === 'repare') couleurEtat = '#059669';
    if (demande.etatReparation === 'attente') couleurEtat = '#d97706';

    sujet = `État réparation — Moteur ${demande.puissance} kW (${demande.installation})`;
    corps = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;">
  <div style="background:#1d4ed8;color:white;padding:20px 24px;">
    <h2 style="margin:0;font-size:18px;">Mise à jour — État du moteur</h2>
    <p style="margin:4px 0 0;opacity:.85;font-size:13px;">${id}</p>
  </div>
  <div style="padding:24px;">
    <p style="font-size:14px;margin-bottom:16px;">Bonjour <b>${demande.demandeur}</b>,</p>
    <p style="font-size:14px;">
      Le moteur de <b>${demande.puissance} kW</b>${demande.matricule ? ', matricule : <b>' + demande.matricule + '</b>,' : ''}
      déposé de <b>${demande.objetTechnique} (${demande.installation})</b> est :
    </p>
    <div style="margin:24px 0;text-align:center;">
      <span style="display:inline-block;background:#f0f6ff;color:${couleurEtat};border:2px solid ${couleurEtat};border-radius:10px;padding:14px 36px;font-size:20px;font-weight:700;">
        ${etat}
      </span>
    </div>
  </div>
  <div style="background:#f1f5f9;padding:12px 24px;font-size:12px;color:#64748b;">${SENDER_NAME}</div>
</div>`;
    // CC pour tous les états : Responsable méthode + Interchangeable électrique
    // + Responsable électricien en plus si état = réparé
    var ccEtat = getUserEmails(['Responsable méthode', 'Interchangeable électrique']);
    if (demande.etatReparation === 'repare') {
      var ccExtra = getUserEmails(['Responsable électricien']);
      if (ccExtra) ccEtat = ccEtat ? ccEtat + ',' + ccExtra : ccExtra;
    }
    var opts = { htmlBody: corps, name: SENDER_NAME };
    if (ccEtat) opts.cc = ccEtat;
    sendEmailOCP(emailDest, sujet, '', opts);

  } else {
    const statut = statutLabel[demande.statut] || demande.statut;
    let couleur  = '#1d4ed8';
    if (demande.statut === 'approuvee') couleur = '#059669';
    if (demande.statut === 'refusee')   couleur = '#dc2626';

    sujet = `Demande ${id} — ${statut}`;
    corps = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;">
  <div style="background:${couleur};color:white;padding:20px 24px;">
    <h2 style="margin:0;font-size:18px;">Mise à jour de votre demande</h2>
    <p style="margin:4px 0 0;opacity:.85;font-size:13px;">${id}</p>
  </div>
  <div style="padding:24px;">
    <p style="font-size:14px;margin-bottom:16px;">Bonjour <b>${demande.demandeur}</b>,</p>
    <p style="font-size:14px;">
      Votre demande de <b>${demande.type === 'changement' ? 'Changement' : 'Pose'}</b>
      du moteur électrique <b>${demande.puissance} kW</b> —
      <b>${demande.installation}</b> — <b>${demande.objetTechnique}</b> a été mise à jour :
    </p>
    <div style="background:#f8faff;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin:16px 0;">
      <div style="font-size:14px;">
        Statut : <b style="color:${couleur};">${statut}</b>
      </div>
      ${demande.justifRefus ? `
      <div style="margin-top:12px;padding:12px;background:#fef2f2;border-radius:6px;font-size:13px;color:#dc2626;">
        <b>Motif de refus :</b> ${demande.justifRefus}
      </div>` : ''}
    </div>
  </div>
  <div style="background:#f1f5f9;padding:12px 24px;font-size:12px;color:#64748b;">${SENDER_NAME}</div>
</div>`;
    // CC approbation/refus : Resp. électricien, Resp. mécanicien, Resp. méthode, Interchangeable électrique
    var ccStatut = '';
    if (demande.statut === 'approuvee' || demande.statut === 'refusee') {
      ccStatut = getUserEmails(['Responsable électricien', 'Responsable mécanicien', 'Responsable méthode', 'Interchangeable électrique', 'Contrôle matériel']);
    }
    var optsStatut = { htmlBody: corps, name: SENDER_NAME };
    if (ccStatut) optsStatut.cc = ccStatut;
    sendEmailOCP(emailDest, sujet, '', optsStatut);
  }
}

// ══════════════════════════════════════════════
// UTILITAIRE
// ══════════════════════════════════════════════
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fonction de test
function testEmail() {
  const demande = {
    id: 'DMD-TEST',
    type: 'changement',
    installation: 'ATL-01',
    objetTechnique: 'Pompe P101',
    puissance: '15',
    tension: '380',
    vitesse: '1500',
    anomalie: 'Vibration excessive',
    matricule: 'M-1234',
    demandeur: 'Test User',
    demandeurEmail: 'mar.elamraoui@gmail.com',
    dateDemande: new Date().toISOString()
  };
  sendEmailNouvelleDemande(demande);
  Logger.log('✅ Email envoyé');
}

// Test envoi email CM
function testEmailCM() {
  var fakeParams = {
    fileId:            SHEET_ID,
    demandeId:         'DMD-TEST',
    matriculeRechange: 'MAT-2024-001',
    sendEmail:         'true',
    emailTo:           EMAIL_ADMIN,
    demandeur:         'Test User',
    installation:      'ATL-01',
    typeDemande:       'changement',
    objetTechnique:    'Pompe P101',
    puissance:         '15'
  };
  // Simuler l'action
  var mp = fakeParams;
  var typeLabel = mp.typeDemande === 'changement' ? '🔄 Changement' : '➕ Pose';
  var sujet = '[Contrôle Matériel] Rechange saisi — ' + mp.demandeId;
  Logger.log('Sujet : ' + sujet);
  Logger.log('✅ Test email CM — vérifiez les logs');
}

// ══════════════════════════════════════════════
// DEBUG — vérifier les feuilles + IDs
// ══════════════════════════════════════════════
function debugMatriculeRechange() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Lister tous les onglets
  var sheets = ss.getSheets();
  Logger.log('=== ONGLETS DU FICHIER ===');
  sheets.forEach(function(s) { Logger.log('  → ' + s.getName()); });

  // Afficher les 5 premiers IDs de 'Demandes'
  Logger.log('\n=== IDs dans "' + SHEET_NAME + '" (5 premiers) ===');
  var dSheet = ss.getSheetByName(SHEET_NAME);
  if (dSheet) {
    var dVals = dSheet.getDataRange().getValues();
    for (var i = 1; i <= Math.min(5, dVals.length - 1); i++) {
      Logger.log('  Ligne ' + (i+1) + ' → col A : "' + dVals[i][0] + '"  |  col T : "' + dVals[i][19] + '"');
    }
  } else {
    Logger.log('  ⚠ Feuille "' + SHEET_NAME + '" introuvable !');
  }

  // Afficher les 5 premiers IDs de 'Demande des intercheable'
  Logger.log('\n=== IDs dans "' + SHEET_INTERCH + '" (5 premiers) ===');
  var mSheet = ss.getSheetByName(SHEET_INTERCH);
  if (mSheet) {
    var mVals = mSheet.getDataRange().getValues();
    for (var j = 1; j <= Math.min(5, mVals.length - 1); j++) {
      Logger.log('  Ligne ' + (j+1) + ' → col A : "' + mVals[j][0] + '"  |  col T : "' + mVals[j][19] + '"');
    }
  } else {
    Logger.log('  ⚠ Feuille "' + SHEET_INTERCH + '" introuvable !');
  }
}
