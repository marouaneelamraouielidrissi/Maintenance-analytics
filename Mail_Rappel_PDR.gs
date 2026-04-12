// ============================================================
//  Mail_Rappel_PDR.gs
//  Rappel automatique PDR en attente de confirmation
//  Déclencheur : chaque mercredi à 08h00
// ============================================================

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

function envoyerRappelPDR() {

  // ── IDs des classeurs ────────────────────────────────────────
  var travSpreadsheetId  = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
  var usersSpreadsheetId = '1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4';

  // ── Chargement de la feuille Travaux hebdomadaire ────────────
  var travSS    = SpreadsheetApp.openById(travSpreadsheetId);
  var travSheet = travSS.getSheetByName('Travaux hebdomadaire');
  var travData  = travSheet.getDataRange().getValues();

  // Indices colonnes (base 0) dans la feuille Travaux :
  //  A=0:ordre  B=1  C=2  D=3:desc  E=4  F=5:objet
  //  G=6  H=7  I=8:poste  K=10:statutUtil  ...
  //  S=18:pdr   T=19:dispo   U=20:obs   V=21:statutSys
  var COL_ORDRE       = 0;
  var COL_DESC        = 3;
  var COL_OBJET       = 5;
  var COL_POSTE       = 8;
  var COL_STATUT_UTIL = 10; // Colonne K — statut utilisateur (CRPR)
  var COL_PDR         = 18;
  var COL_DISPO       = 19;
  var COL_OBS         = 20;
  var COL_STATUT_SYS  = 21; // Colonne V — statut système (LANC)

  // ── Chargement de la feuille Users ──────────────────────────
  var usersSS    = SpreadsheetApp.openById(usersSpreadsheetId);
  var usersSheet = usersSS.getSheetByName('Users');
  var usersData  = usersSheet.getDataRange().getValues();

  // Indices colonnes Users (base 0) :
  //  A=0:id  B=1:nom  C=2:prenom  D=3:tel  E=4:email  F=5:fonction
  var COL_U_NOM     = 1;
  var COL_U_PRENOM  = 2;
  var COL_U_EMAIL   = 4;
  var COL_U_FONCTION = 5;

  // ── Fonction : retourner la liste des destinataires par profil ─
  function getDestinataires(profil) {
    var result = [];
    for (var i = 1; i < usersData.length; i++) {
      var fonc  = String(usersData[i][COL_U_FONCTION] || '').trim();
      var email = String(usersData[i][COL_U_EMAIL]    || '').trim();
      if (fonc.toLowerCase() === profil.toLowerCase() && email) {
        result.push({
          email  : email,
          prenom : String(usersData[i][COL_U_PRENOM] || '').trim(),
          nom    : String(usersData[i][COL_U_NOM]    || '').trim()
        });
      }
    }
    return result;
  }

  // ── Liste CC : retourne les emails des profils demandés ────────
  // profils = tableau de noms de fonctions (ex: ['Responsable méthode', 'Interchangeable électrique'])
  function getCCList(profils) {
    var ccEmails = [];
    for (var i = 1; i < usersData.length; i++) {
      var fonc  = String(usersData[i][COL_U_FONCTION] || '').trim().toLowerCase()
                    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      var email = String(usersData[i][COL_U_EMAIL] || '').trim();
      if (!email) continue;
      for (var p = 0; p < profils.length; p++) {
        var profilNorm = profils[p].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        if (fonc === profilNorm) {
          ccEmails.push(email);
          break;
        }
      }
    }
    return ccEmails;
  }

  // Mots-clés PDR réservés au Bureau de méthode — exclus du service mécanique
  var PDR_BUREAU_KEYWORDS = ['reducteur', 'pompe'];

  function pdrContientMotBureau(pdr) {
    var pdrNorm = pdr.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    for (var k = 0; k < PDR_BUREAU_KEYWORDS.length; k++) {
      if (pdrNorm.indexOf(PDR_BUREAU_KEYWORDS[k]) >= 0) return true;
    }
    return false;
  }

  // ── Fonction de base : filtre statut + PDR + dispo + obs ───────
  function getBaseFiltered(poste) {
    var result = [];
    for (var i = 1; i < travData.length; i++) {
      var row = travData[i];

      var rowPoste = String(row[COL_POSTE] || '').trim().toUpperCase();
      if (rowPoste !== poste.toUpperCase()) continue;

      var statutSys  = String(row[COL_STATUT_SYS]  || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      var statutUtil = String(row[COL_STATUT_UTIL]  || '').trim().toUpperCase();
      if (statutSys.indexOf('cree') < 0)  continue;
      if (statutUtil.indexOf('CRPR') < 0) continue;

      var pdr  = String(row[COL_PDR]  || '').trim(); if (!pdr)  continue;
      var dispo = String(row[COL_DISPO] || '').trim(); if (dispo) continue;
      var obs   = String(row[COL_OBS]   || '').trim(); if (obs)   continue;

      result.push({
        ordre : String(row[COL_ORDRE] || '').trim(),
        desc  : String(row[COL_DESC]  || '').trim(),
        objet : String(row[COL_OBJET] || '').trim(),
        pdr   : pdr
      });
    }
    return result;
  }

  // OTs standards (sans réducteur/pompe)
  function getPending(poste) {
    return getBaseFiltered(poste).filter(function(r) { return !pdrContientMotBureau(r.pdr); });
  }

  // OTs réducteur/pompe uniquement (→ Interchangeable mécanique)
  function getPendingBureau(poste) {
    return getBaseFiltered(poste).filter(function(r) { return pdrContientMotBureau(r.pdr); });
  }

  // ── Fonction : construire le corps HTML de l'email ──────────
  function buildHtml(posteLabel, dest, pendingList) {
    var salutation = dest.prenom ? 'Bonjour ' + dest.prenom + ',' : 'Bonjour,';
    var rows = '';
    for (var k = 0; k < pendingList.length; k++) {
      var item = pendingList[k];
      rows += '<tr style="border-bottom:1px solid #e5e7eb;">'
            + '<td style="padding:8px 12px;font-weight:600;color:#1e40af;">' + item.ordre + '</td>'
            + '<td style="padding:8px 12px;">' + (item.objet || '-') + '</td>'
            + '<td style="padding:8px 12px;">' + (item.desc  || '-') + '</td>'
            + '<td style="padding:8px 12px;color:#b45309;">' + item.pdr + '</td>'
            + '</tr>';
    }

    return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#374151;background:#f9fafb;margin:0;padding:0;">'
         + '<div style="max-width:700px;margin:30px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);">'
         + '<div style="background:#1e3a5f;padding:24px 32px;">'
         + '<h2 style="margin:0;color:#fff;font-size:18px;">PDR en attente de confirmation</h2>'
         + '</div>'
         + '<div style="padding:28px 32px;">'
         + '<p style="margin:0 0 16px;">' + salutation + '</p>'
         + '<p style="margin:0 0 20px;">Les ordres de travail suivants ont des <strong>pieces de rechange (PDR)</strong> en attente de votre confirmation de disponibilite :</p>'
         + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
         + '<thead><tr style="background:#f1f5f9;">'
         + '<th style="padding:10px 12px;text-align:left;color:#475569;">N&deg; OT</th>'
         + '<th style="padding:10px 12px;text-align:left;color:#475569;">Objet technique</th>'
         + '<th style="padding:10px 12px;text-align:left;color:#475569;">Description</th>'
         + '<th style="padding:10px 12px;text-align:left;color:#475569;">PDR demandee</th>'
         + '</tr></thead>'
         + '<tbody>' + rows + '</tbody>'
         + '</table>'
         + '</div>'
         + '<div style="background:#f8fafc;padding:16px 32px;border-top:1px solid #e5e7eb;">'
         + '<p style="margin:0;font-size:12px;color:#9ca3af;">Message automatique - Bureau de methode Daoui | Ne pas repondre a cet email</p>'
         + '</div>'
         + '</div></body></html>';
  }

  // ── Envoi email générique ────────────────────────────────────
  // pending    : liste des OTs
  // profil     : fonction des destinataires TO
  // ccProfils  : tableau des fonctions à mettre en CC
  // poste      : label pour le log
  function sendEmail(pending, profil, ccProfils, poste) {
    if (pending.length === 0) {
      Logger.log(poste + ' [' + profil + '] — Aucun PDR en attente.');
      return;
    }

    var destinataires = getDestinataires(profil);
    if (destinataires.length === 0) {
      Logger.log(poste + ' — Email non envoye : aucun destinataire pour profil "' + profil + '"');
      return;
    }

    var ccString = getCCList(ccProfils).join(',');
    var subject  = 'Rappel - ' + pending.length + ' PDR en attente de confirmation';

    for (var d = 0; d < destinataires.length; d++) {
      var dest     = destinataires[d];
      var htmlBody = buildHtml(poste, dest, pending);
      var options  = { htmlBody: htmlBody, name: 'Bureau de methode Daoui - Section Planification' };
      if (ccString) options.cc = ccString;
      sendEmailOCP(dest.email, subject, '', options);
      Logger.log(poste + ' [' + profil + '] — Email envoye a : ' + dest.email + ' (' + pending.length + ' PDR)');
    }
  }

  // ── Profils CC communs ────────────────────────────────────────
  var CC_BASE          = ['Responsable méthode', 'Interchangeable électrique'];
  var CC_APPRO_MEC     = ['Responsable appro mécanique',    'Responsable méthode', 'Interchangeable électrique'];
  var CC_APPRO_INST    = ['Responsable appro installation', 'Responsable méthode', 'Interchangeable électrique'];

  // ── Envois ───────────────────────────────────────────────────

  // 421-MEC — OTs standards → Appro mécanique
  sendEmail(getPending('421-MEC'),       'Appro mécanique',      CC_APPRO_MEC,  '421-MEC');
  // 421-MEC — OTs réducteur/pompe → Interchangeable mécanique
  sendEmail(getPendingBureau('421-MEC'), 'Interchangeable mécanique', CC_BASE,  '421-MEC (reducteur/pompe)');

  // 421-CHAU → Appro mécanique (mêmes règles que MEC, sans filtre bureau)
  sendEmail(getPending('421-CHAU'),      'Appro mécanique',      CC_APPRO_MEC,  '421-CHAU');

  // 421-INST → Appro installation
  sendEmail(getPending('421-INST'),      'Appro installation',   CC_APPRO_INST, '421-INST');

  // 423-ELEC → Appro électrique
  sendEmail(getPending('423-ELEC'),      'Appro électrique',     CC_BASE,       '423-ELEC');

  // 423-REG → Appro Instrumentation
  sendEmail(getPending('423-REG'),       'Appro Instrumentation',CC_BASE,       '423-REG');
}

// ============================================================
//  DIAGNOSTIC — affiche dans les logs les valeurs réelles
//  des colonnes clés pour comprendre pourquoi aucun PDR n'est trouvé
// ============================================================
function diagnostiquerPDR() {

  var travSpreadsheetId = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';
  var travSS = SpreadsheetApp.openById(travSpreadsheetId);

  // Lister toutes les feuilles disponibles
  var allSheets = travSS.getSheets();
  Logger.log('=== FEUILLES DISPONIBLES ===');
  for (var s = 0; s < allSheets.length; s++) {
    Logger.log('  [' + s + '] "' + allSheets[s].getName() + '"');
  }

  var travSheet = travSS.getSheetByName('Travaux hebdomadaire');
  Logger.log('\nFeuille utilisée : "' + travSheet.getName() + '"');

  var travData = travSheet.getDataRange().getValues();
  Logger.log('Nombre total de lignes : ' + travData.length);

  // Afficher la ligne d'en-tête
  Logger.log('\n=== EN-TÊTES (ligne 1) ===');
  var headers = travData[0];
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim()) {
      Logger.log('  Col ' + h + ' (' + columnLetter(h) + ') = "' + headers[h] + '"');
    }
  }

  // Analyser les 10 premières lignes de données
  Logger.log('\n=== VALEURS COLONNES CLÉS (10 premières lignes) ===');
  Logger.log('Ligne | Col0(ordre) | Col8(poste) | Col18(PDR) | Col19(dispo) | Col20(obs) | Col21(statut)');
  for (var i = 1; i <= Math.min(10, travData.length - 1); i++) {
    var r = travData[i];
    Logger.log(
      'L' + (i+1) + ' | ' +
      '"' + String(r[0]||'').trim() + '" | ' +
      '"' + String(r[8]||'').trim() + '" | ' +
      '"' + String(r[18]||'').trim() + '" | ' +
      '"' + String(r[19]||'').trim() + '" | ' +
      '"' + String(r[20]||'').trim() + '" | ' +
      '"' + String(r[21]||'').trim() + '"'
    );
  }

  // Compter par filtre pour identifier où ça bloque
  Logger.log('\n=== ANALYSE FILTRES ===');
  var cntTotal   = 0;
  var cntPoste   = 0;
  var cntStatut  = 0;
  var cntPDR     = 0;
  var cntPending = 0;

  // Valeurs uniques des postes et statuts rencontrés
  var postesVus  = {};
  var statutsVus = {};

  for (var i = 1; i < travData.length; i++) {
    var row = travData[i];
    if (!String(row[0]||'').trim()) continue; // ligne vide
    cntTotal++;

    var rowPoste = String(row[8]||'').trim();
    var rowStatut = String(row[21]||'').trim();
    var rowPDR   = String(row[18]||'').trim();
    var rowDispo  = String(row[19]||'').trim();

    postesVus[rowPoste]   = (postesVus[rowPoste]   || 0) + 1;
    statutsVus[rowStatut] = (statutsVus[rowStatut] || 0) + 1;

    var posteMatch = ['421-MEC','423-ELEC','421-INST','423-REG'].indexOf(rowPoste.toUpperCase()) >= 0;
    if (!posteMatch) continue;
    cntPoste++;

    var statutUpper = rowStatut.toUpperCase();
    var statutMatch = statutUpper.indexOf('LANC') >= 0 || statutUpper.indexOf('CRPR') >= 0;
    if (!statutMatch) continue;
    cntStatut++;

    if (!rowPDR) continue;
    cntPDR++;

    if (rowDispo) continue;
    cntPending++;
  }

  Logger.log('Lignes non vides              : ' + cntTotal);
  Logger.log('Après filtre POSTE (4 postes) : ' + cntPoste);
  Logger.log('Après filtre STATUT (LANC/CRPR): ' + cntStatut);
  Logger.log('Après filtre PDR non vide     : ' + cntPDR);
  Logger.log('PDR EN ATTENTE (dispo vide)   : ' + cntPending);

  Logger.log('\n=== POSTES PRÉSENTS DANS LA FEUILLE ===');
  for (var p in postesVus) Logger.log('  "' + p + '" : ' + postesVus[p] + ' lignes');

  Logger.log('\n=== STATUTS PRÉSENTS DANS LA FEUILLE (col 21) ===');
  for (var st in statutsVus) Logger.log('  "' + st + '" : ' + statutsVus[st] + ' lignes');
}

function columnLetter(index) {
  var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  if (index < 26) return letters[index];
  return letters[Math.floor(index/26)-1] + letters[index%26];
}

// ============================================================
//  FONCTION DE TEST — envoie un email de test aux
//  utilisateurs ayant le profil 'Interchangeable électrique'
// ============================================================
function testerEnvoiInterchangeable() {

  var usersSpreadsheetId = '1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4';
  var usersSS    = SpreadsheetApp.openById(usersSpreadsheetId);
  var usersSheet = usersSS.getSheetByName('Users');
  var usersData  = usersSheet.getDataRange().getValues();

  var COL_U_NOM      = 1;
  var COL_U_PRENOM   = 2;
  var COL_U_EMAIL    = 4;
  var COL_U_FONCTION = 5;

  // Récupérer tous les Interchangeable électrique
  var destinataires = [];
  for (var i = 1; i < usersData.length; i++) {
    var fonc  = String(usersData[i][COL_U_FONCTION] || '').trim().toLowerCase();
    var email = String(usersData[i][COL_U_EMAIL]    || '').trim();
    if (!email) continue;
    if (fonc === 'interchangeable electrique' || fonc === 'interchangeable électrique') {
      destinataires.push({
        email  : email,
        prenom : String(usersData[i][COL_U_PRENOM] || '').trim(),
        nom    : String(usersData[i][COL_U_NOM]    || '').trim()
      });
    }
  }

  if (destinataires.length === 0) {
    Logger.log('TEST — Aucun utilisateur trouve avec le profil Interchangeable electrique.');
    Logger.log('Verifiez la colonne F de la feuille Users.');
    return;
  }

  Logger.log('TEST — ' + destinataires.length + ' destinataire(s) trouve(s) :');

  for (var d = 0; d < destinataires.length; d++) {
    var dest = destinataires[d];
    Logger.log('  -> ' + dest.prenom + ' ' + dest.nom + ' (' + dest.email + ')');

    var salutation = dest.prenom ? 'Bonjour ' + dest.prenom + ',' : 'Bonjour,';

    var htmlBody = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#374151;background:#f9fafb;margin:0;padding:0;">'
      + '<div style="max-width:600px;margin:30px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);">'
      + '<div style="background:#1e3a5f;padding:24px 32px;">'
      + '<h2 style="margin:0;color:#fff;font-size:18px;">[TEST] Email de verification</h2>'
      + '<p style="margin:4px 0 0;color:#93c5fd;font-size:13px;">Bureau de methode Daoui - Section Planification</p>'
      + '</div>'
      + '<div style="padding:28px 32px;">'
      + '<p style="margin:0 0 16px;">' + salutation + '</p>'
      + '<p style="margin:0 0 16px;">Ceci est un <strong>email de test</strong> pour verifier que vous recevez bien les rappels PDR automatiques.</p>'
      + '<p style="margin:0 0 16px;">Vous etes enregistre avec le profil : <strong>Interchangeable electrique</strong></p>'
      + '<p style="margin:0;color:#6b7280;font-size:13px;">Si vous recevez cet email, la configuration est correcte.</p>'
      + '</div>'
      + '<div style="background:#f8fafc;padding:16px 32px;border-top:1px solid #e5e7eb;">'
      + '<p style="margin:0;font-size:12px;color:#9ca3af;">Message automatique - Bureau de methode Daoui | Ne pas repondre</p>'
      + '</div>'
      + '</div></body></html>';

    sendEmailOCP(
      dest.email,
      '[TEST] Verification email - Maintenance Analytics',
      '',
      { htmlBody: htmlBody, name: 'Bureau de methode Daoui - Section Planification' }
    );

    Logger.log('  -> Email envoye a : ' + dest.email);
  }

  Logger.log('TEST termine — ' + destinataires.length + ' email(s) envoye(s).');
}

// ============================================================
//  Créer/réinitialiser le déclencheur hebdomadaire
//  Exécuter UNE SEULE FOIS manuellement pour l'activer
// ============================================================
function creerDeclencheurHebdomadaire() {
  // Supprimer les anciens déclencheurs du même nom pour éviter les doublons
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'envoyerRappelPDR') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Créer le nouveau déclencheur : chaque mercredi à 08h00
  ScriptApp.newTrigger('envoyerRappelPDR')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8)
    .nearMinute(0)
    .create();

  Logger.log('Declencheur cree : chaque mercredi a 08h00');
}

// ============================================================
//  TEST RAPIDE — envoie un email de test à m.elamraoui@ocpgroup.ma
//  avec un aperçu réel des PDR en attente dans la feuille
// ============================================================
function testerEnvoiMarouane() {

  var TEST_EMAIL = 'm.elamraoui@ocpgroup.ma';
  var travSpreadsheetId = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE';

  var travSS    = SpreadsheetApp.openById(travSpreadsheetId);
  var travSheet = travSS.getSheetByName('Travaux hebdomadaire');
  var travData  = travSheet.getDataRange().getValues();

  Logger.log('Feuille : "' + travSheet.getName() + '" — ' + (travData.length - 1) + ' lignes de données');

  // Récupérer les PDR en attente pour le poste 421-MEC uniquement
  var POSTE_TEST = '421-INST';
  var allPending = [];
  for (var i = 1; i < travData.length; i++) {
    var row    = travData[i];
    var ordre      = String(row[0]  || '').trim();
    var poste      = String(row[8]  || '').trim();
    var statutUtil = String(row[10] || '').trim().toUpperCase(); // col K
    var pdr        = String(row[18] || '').trim();
    var dispo      = String(row[19] || '').trim();
    var statutSys  = String(row[21] || '').trim().toUpperCase(); // col V

    if (!ordre) continue;
    if (poste.toUpperCase() !== POSTE_TEST) continue;
    if (!pdr)   continue;
    if (dispo)  continue;

    // Exclure si observation (col U) remplie
    var obs = String(row[20] || '').trim();
    if (obs) continue;

    // Statut sys col V = 'créé'  ET  statut util col K = CRPR
    var statutSysNorm = statutSys.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (statutSysNorm.indexOf('cree') < 0)  continue;
    if (statutUtil.indexOf('CRPR') < 0)     continue;

    // Exclure les PDR contenant réducteur ou pompe (→ Bureau de méthode)
    var pdrNorm = pdr.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (pdrNorm.indexOf('reducteur') >= 0 || pdrNorm.indexOf('pompe') >= 0) continue;

    allPending.push({
      ordre     : ordre,
      poste     : poste,
      pdr       : pdr,
      statutSys : statutSys,
      statutUtil: statutUtil,
      desc      : String(row[3] || '').trim(),
      objet     : String(row[5] || '').trim()
    });
  }

  Logger.log('PDR trouvées : ' + allPending.length);
  for (var k = 0; k < allPending.length; k++) {
    Logger.log('  ' + allPending[k].ordre + ' | ' + allPending[k].poste + ' | SYS="' + allPending[k].statutSys + '" | UTIL="' + allPending[k].statutUtil + '" | PDR=' + allPending[k].pdr);
  }

  // Construire le tableau HTML
  var rows = '';
  if (allPending.length === 0) {
    rows = '<tr><td colspan="5" style="padding:12px;text-align:center;color:#6b7280;">Aucune PDR en attente trouvée</td></tr>';
  } else {
    for (var j = 0; j < allPending.length; j++) {
      var item = allPending[j];
      var flagCouleur = '#16a34a';
      rows += '<tr style="border-bottom:1px solid #e5e7eb;">'
            + '<td style="padding:8px 10px;font-weight:600;color:#1e40af;">' + item.ordre + '</td>'
            + '<td style="padding:8px 10px;">' + (item.objet || '-') + '</td>'
            + '<td style="padding:8px 10px;">' + (item.desc  || '-') + '</td>'
            + '<td style="padding:8px 10px;color:#b45309;">' + item.pdr + '</td>'
            + '<td style="padding:8px 10px;font-weight:600;color:' + flagCouleur + ';">' + (item.poste || '-') + '</td>'
            + '</tr>';
    }
  }

  var htmlBody = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#374151;background:#f9fafb;margin:0;padding:0;">'
    + '<div style="max-width:750px;margin:30px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);">'
    + '<div style="background:#1e3a5f;padding:20px 28px;">'
    + '<h2 style="margin:0;color:#fff;font-size:17px;">[TEST] Rappel PDR en attente de confirmation</h2>'
    + '<p style="margin:4px 0 0;color:#93c5fd;font-size:12px;">Ceci est un email de test — Poste 421-MEC (' + allPending.length + ' PDR)</p>'
    + '</div>'
    + '<div style="padding:24px 28px;">'
    + '<p style="margin:0 0 16px;">Bonjour Marouane,</p>'
    + '<p style="margin:0 0 18px;">Voici un apercu des PDR en attente de confirmation dans la feuille <strong>Travaux hebdomadaire</strong> :</p>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<thead><tr style="background:#f1f5f9;">'
    + '<th style="padding:9px 10px;text-align:left;color:#475569;">N&deg; OT</th>'
    + '<th style="padding:9px 10px;text-align:left;color:#475569;">Objet technique</th>'
    + '<th style="padding:9px 10px;text-align:left;color:#475569;">Description</th>'
    + '<th style="padding:9px 10px;text-align:left;color:#475569;">PDR demandee</th>'
    + '<th style="padding:9px 10px;text-align:left;color:#475569;">Poste</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table>'
    + '<p style="margin:20px 0 0;font-size:12px;color:#6b7280;">Si vous recevez bien cet email, la configuration est correcte. En production, chaque service ne recoit que ses propres PDR.</p>'
    + '</div>'
    + '<div style="background:#f8fafc;padding:14px 28px;border-top:1px solid #e5e7eb;">'
    + '<p style="margin:0;font-size:11px;color:#9ca3af;">Email de test — Maintenance Analytics | Bureau de methode Daoui</p>'
    + '</div>'
    + '</div></body></html>';

  sendEmailOCP(
    TEST_EMAIL,
    '[TEST] Rappel PDR — ' + allPending.length + ' PDR en attente — Maintenance Analytics',
    '',
    { htmlBody: htmlBody, name: 'Bureau de methode Daoui - Section Planification' }
  );

  Logger.log('Email de test envoye a : ' + TEST_EMAIL);
}
