// ============================================================
//  WebApp.gs — Apps Script principal (Maintenance Analytics)
//  Déployer comme Web App : accès Tout le monde, Moi
// ============================================================

// ── IDs des classeurs ─────────────────────────────────────────
var MOTEURS_SS_ID = '1rOPljpAHYIs_uQ5-EUnL4yVwy2ciKAn20htExnU2vG4'; // Moteurs + Users + Codes + Notifs
var TRAVAUX_SS_ID = '1C9bYkPsoYg81ARgolVDlZRwsMZk4Seff6aC7vfxoVeE'; // Travaux hebdomadaire + Arrêts

// ── Nom des feuilles (classeur MOTEURS) ──────────────────────
var SH_DEMANDES      = 'Demandes';
var SH_USERS         = 'Users';
var SH_INSTALLATIONS = 'Installations';
var SH_CODES         = 'Codes';
var SH_NOTIFS        = 'Notifications';
var SH_SUBSCRIPTIONS = 'Subscriptions';

// ── Colonnes feuille Demandes (base 1) ────────────────────────
// A  B     C               D             E         F        G          H               I         J          K               L       M            N            O
// id|type|objetTechnique|installation|puissance|tension|demandeur|demandeurEmail|anomalie|matricule|etatReparation|statut|dateDemande|justifRefus|matriculeRechange

// ── Colonnes feuille Users (base 1) ───────────────────────────
// A  B    C       D    E      F
// id|nom|prenom|tel|email|fonction

// ── Colonnes feuille Installations (base 1) ───────────────────
// A             B             C
// installation|demandeurNom|demandeurEmail

// ── Colonnes feuille Codes (base 1) ───────────────────────────
// A    B
// key|value

// ── Colonnes feuille Notifications (base 1) ──────────────────
// A   B      C     D             E
// id|title|body|targetProfile|timestamp

// ── Colonnes feuille Subscriptions (base 1) ──────────────────
// A        B          C
// profile|endpoint|subscription


/* ════════════════════════════════════════════════════════════
   HELPERS GÉNÉRAUX
   ════════════════════════════════════════════════════════════ */

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(ssId, sheetName) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Feuille introuvable : ' + sheetName);
  return sh;
}

function sheetToObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}


/* ════════════════════════════════════════════════════════════
   doGet — Requêtes GET
   ════════════════════════════════════════════════════════════ */

function doGet(e) {
  var action = e.parameter.action || '';
  try {
    switch (action) {

      /* ── Données de base ─────────────────────────────────── */
      case 'getInstallations': return getInstallations();
      case 'getUsers':         return getUsers_();

      /* ── Gestion utilisateurs ────────────────────────────── */
      case 'addUser':    return addUser_(e);
      case 'deleteUser': return deleteUser_(e);

      /* ── Moteurs : codes accès ───────────────────────────── */
      case 'updateCode': return updateCode_(e);

      /* ── Moteurs : matricule de rechange ─────────────────── */
      case 'updateMatriculeRechange': return updateMatriculeRechange_(e);

      /* ── Moteurs : email changement état réparation ──────── */
      case 'sendEmailEtatMoteur': return sendEmailEtatMoteur_(e);

      /* ── Travaux hebdomadaire ────────────────────────────── */
      case 'updateRealisation': return updateRealisation_(e);
      case 'updatePDR':         return updatePDR_(e);
      case 'markArretReporte':  return markArretReporte_(e);

      /* ── Notifications broadcast ─────────────────────────── */
      case 'logNotif':  return logNotif_(e);
      case 'getNotifs': return getNotifs_(e);

      /* ── Web Push subscriptions ──────────────────────────── */
      case 'saveSubscription': return saveSubscription_(e);

      default:
        // Lecture globale (fallback initial : retourne les demandes)
        return getAllDemandes_();
    }
  } catch(err) {
    Logger.log('doGet error [' + action + '] : ' + err.message);
    return jsonError(err.message);
  }
}


/* ════════════════════════════════════════════════════════════
   doPost — Requêtes POST
   ════════════════════════════════════════════════════════════ */

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(ex) {}
  var action = body.action || '';
  try {
    switch (action) {
      case 'save':   return saveDemande_(body);
      case 'update': return updateDemande_(body);
      default:       return jsonError('Action inconnue : ' + action);
    }
  } catch(err) {
    Logger.log('doPost error [' + action + '] : ' + err.message);
    return jsonError(err.message);
  }
}


/* ════════════════════════════════════════════════════════════
   INSTALLATIONS & DEMANDEURS
   ════════════════════════════════════════════════════════════ */

function getInstallations() {
  var sh   = getSheet(MOTEURS_SS_ID, SH_INSTALLATIONS);
  var data = sh.getDataRange().getValues();
  var installations = [];
  var demandeursMap = {};

  for (var i = 1; i < data.length; i++) {
    var row  = data[i];
    var inst = String(row[0] || '').trim();
    var nom  = String(row[1] || '').trim();
    var mail = String(row[2] || '').trim();
    if (inst && installations.indexOf(inst) < 0) installations.push(inst);
    if (nom && !demandeursMap[nom]) demandeursMap[nom] = mail;
  }

  var demandeurs = Object.keys(demandeursMap).map(function(n) {
    return { nom: n, email: demandeursMap[n] };
  });

  return jsonOk({ success: true, installations: installations, demandeurs: demandeurs });
}


/* ════════════════════════════════════════════════════════════
   UTILISATEURS
   ════════════════════════════════════════════════════════════ */

function getUsers_() {
  var sh   = getSheet(MOTEURS_SS_ID, SH_USERS);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return jsonOk([]);
  var headers = data[0];
  var users = data.slice(1).map(function(row) {
    var u = {};
    headers.forEach(function(h, i) { u[h] = row[i]; });
    return u;
  });
  return jsonOk(users);
}

function addUser_(e) {
  var sh = getSheet(MOTEURS_SS_ID, SH_USERS);
  var id = e.parameter.id || ('USR_' + Date.now());
  sh.appendRow([
    id,
    e.parameter.nom     || '',
    e.parameter.prenom  || '',
    e.parameter.tel     || '',
    e.parameter.email   || '',
    e.parameter.fonction || ''
  ]);
  return jsonOk({ success: true, id: id });
}

function deleteUser_(e) {
  var id = e.parameter.id || '';
  if (!id) return jsonError('id manquant');
  var sh   = getSheet(MOTEURS_SS_ID, SH_USERS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      sh.deleteRow(i + 1);
      return jsonOk({ success: true });
    }
  }
  return jsonError('Utilisateur introuvable : ' + id);
}

/* Helper : récupérer les emails d'une liste de fonctions */
function getEmailsByFonctions(fonctions) {
  var sh   = getSheet(MOTEURS_SS_ID, SH_USERS);
  var data = sh.getDataRange().getValues();
  var emails = [];
  for (var i = 1; i < data.length; i++) {
    var fonc  = String(data[i][5] || '').trim();
    var email = String(data[i][4] || '').trim();
    if (!email) continue;
    for (var j = 0; j < fonctions.length; j++) {
      if (fonc === fonctions[j]) { emails.push(email); break; }
    }
  }
  // Dédoublonner
  return emails.filter(function(v, i, a) { return a.indexOf(v) === i; });
}


/* ════════════════════════════════════════════════════════════
   CODES D'ACCÈS
   ════════════════════════════════════════════════════════════ */

function updateCode_(e) {
  var sh   = getSheet(MOTEURS_SS_ID, SH_CODES);
  var data = sh.getDataRange().getValues();
  var map  = { adminCode: e.parameter.adminCode, approCode: e.parameter.approCode,
               execCode:  e.parameter.execCode,  cmCode:    e.parameter.cmCode };

  Object.keys(map).forEach(function(key) {
    var val = map[key];
    if (!val) return;
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        sh.getRange(i + 1, 2).setValue(val);
        data[i][1] = val;
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, val]);
  });

  return jsonOk({ success: true });
}


/* ════════════════════════════════════════════════════════════
   DEMANDES MOTEURS
   ════════════════════════════════════════════════════════════ */

function getAllDemandes_() {
  try {
    var sh   = getSheet(MOTEURS_SS_ID, SH_DEMANDES);
    var data = sheetToObjects(sh);
    return jsonOk(data);
  } catch(e) {
    return jsonOk([]);
  }
}

function saveDemande_(body) {
  var d  = body.demande || {};
  var sh = getSheet(MOTEURS_SS_ID, SH_DEMANDES);

  // Si feuille vide, écrire les en-têtes
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id','type','objetTechnique','installation','puissance','tension',
                  'demandeur','demandeurEmail','anomalie','matricule','etatReparation',
                  'statut','dateDemande','justifRefus','matriculeRechange']);
  }

  sh.appendRow([
    d.id               || '',
    d.type             || '',
    d.objetTechnique   || '',
    d.installation     || '',
    d.puissance        || '',
    d.tension          || '',
    d.demandeur        || '',
    d.demandeurEmail   || '',
    d.anomalie         || '',
    d.matricule        || '',
    d.etatReparation   || 'attente',
    d.statut           || 'envoyee',
    d.dateDemande      || new Date().toISOString(),
    d.justifRefus      || '',
    d.matriculeRechange|| ''
  ]);

  return jsonOk({ success: true, id: d.id });
}

function updateDemande_(body) {
  var id      = body.id      || '';
  var updates = body.updates || {};
  if (!id) return jsonError('id manquant');

  var sh   = getSheet(MOTEURS_SS_ID, SH_DEMANDES);
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      Object.keys(updates).forEach(function(key) {
        var col = headers.indexOf(key);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(updates[key]);
      });
      return jsonOk({ success: true });
    }
  }
  return jsonError('Demande introuvable : ' + id);
}


/* ════════════════════════════════════════════════════════════
   MATRICULE DE RECHANGE + EMAIL
   ════════════════════════════════════════════════════════════ */

function updateMatriculeRechange_(e) {
  var demandeId = e.parameter.demandeId || '';
  var mat       = e.parameter.matriculeRechange || '';

  // Mise à jour dans la feuille Demandes
  if (demandeId && mat) {
    var sh   = getSheet(MOTEURS_SS_ID, SH_DEMANDES);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colMat  = headers.indexOf('matriculeRechange');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === demandeId) {
        if (colMat >= 0) sh.getRange(i + 1, colMat + 1).setValue(mat);
        break;
      }
    }
  }

  // Envoi email si demandé
  if (e.parameter.sendEmail === 'true') {
    var emailTo      = e.parameter.emailTo      || '';
    var demandeur    = e.parameter.demandeur    || '';
    var installation = e.parameter.installation || '';
    var typeDemande  = e.parameter.typeDemande  || '';
    var objet        = e.parameter.objetTechnique || '';
    var puissance    = e.parameter.puissance    || '';

    if (emailTo) {
      var subject = '[Maintenance] Matricule de rechange enregistré — ' + installation;
      var htmlBody = buildEmailHtml(
        'Matricule de rechange enregistré',
        demandeur,
        [
          { label: 'Installation',      value: installation },
          { label: 'Type de demande',   value: typeDemande  },
          { label: 'Objet technique',   value: objet        },
          { label: 'Puissance',         value: puissance + ' kW' },
          { label: 'Matricule rechange',value: mat          }
        ],
        'Le matricule du moteur de rechange a été enregistré par le service Contrôle Matériel.'
      );
      GmailApp.sendEmail(emailTo, subject, '', {
        htmlBody: htmlBody,
        name: 'Maintenance Analytics — OCP'
      });
    }
  }

  return jsonOk({ success: true });
}


/* ════════════════════════════════════════════════════════════
   EMAIL CHANGEMENT D'ÉTAT RÉPARATION
   ════════════════════════════════════════════════════════════ */

function sendEmailEtatMoteur_(e) {
  var emailTo      = e.parameter.emailTo      || '';
  var cc           = e.parameter.cc           || '';
  var demandeur    = e.parameter.demandeur    || '';
  var installation = e.parameter.installation || '';
  var objet        = e.parameter.objetTechnique || '';
  var puissance    = e.parameter.puissance    || '';
  var matricule    = e.parameter.matricule    || '';
  var etatLabel    = e.parameter.etatLabel    || '';

  // Si cc vide → lire depuis la feuille Users les Responsable méthode + Interchangeable électrique
  if (!cc) {
    var ccEmails = getEmailsByFonctions(['Responsable méthode', 'Interchangeable électrique']);
    cc = ccEmails.join(',');
  }

  if (!emailTo) return jsonOk({ success: false, reason: 'emailTo manquant' });

  var subject = '[Maintenance] Mise à jour état moteur — ' + installation;
  var htmlBody = buildEmailHtml(
    'Mise à jour de l\'état du moteur',
    demandeur,
    [
      { label: 'Installation',    value: installation },
      { label: 'Objet technique', value: objet        },
      { label: 'Puissance',       value: puissance + ' kW' },
      { label: 'Matricule',       value: matricule    },
      { label: 'Nouvel état',     value: etatLabel, highlight: true }
    ],
    'L\'état de votre demande de moteur a été mis à jour par l\'administration.'
  );

  var options = { htmlBody: htmlBody, name: 'Maintenance Analytics — OCP' };
  if (cc) options.cc = cc;

  GmailApp.sendEmail(emailTo, subject, '', options);
  Logger.log('sendEmailEtatMoteur → TO:' + emailTo + ' | CC:' + cc + ' | état:' + etatLabel);

  return jsonOk({ success: true });
}


/* ════════════════════════════════════════════════════════════
   TRAVAUX HEBDOMADAIRE — Réalisation OT
   ════════════════════════════════════════════════════════════ */

function updateRealisation_(e) {
  var fileId    = e.parameter.fileId    || TRAVAUX_SS_ID;
  var sheetName = e.parameter.sheetName || 'Travaux hebdomadaire';
  var ordre     = e.parameter.ordre     || '';
  var valeur    = e.parameter.valeur    || ''; // 'Fait' | 'NFait' | ''
  var desc      = e.parameter.desc      || '';
  var poste     = e.parameter.poste     || '';
  var objet     = e.parameter.objet     || '';

  if (!ordre) return jsonError('ordre manquant');

  var sh   = getSheet(fileId, sheetName);
  var data = sh.getDataRange().getValues();

  // Colonne de réalisation : chercher en-tête 'realisation' ou 'Réalisation' (col W = index 22)
  // Par défaut on utilise la colonne W (index 22 = col 23 base 1)
  var headers    = data[0];
  var colReal    = headers.indexOf('realisation');
  if (colReal < 0) colReal = headers.indexOf('Réalisation');
  if (colReal < 0) colReal = 22; // colonne W par défaut

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === ordre) {
      sh.getRange(i + 1, colReal + 1).setValue(valeur);
      Logger.log('updateRealisation OT=' + ordre + ' → ' + valeur);
      return jsonOk({ success: true });
    }
  }

  return jsonError('OT introuvable : ' + ordre);
}


/* ════════════════════════════════════════════════════════════
   TRAVAUX HEBDOMADAIRE — Confirmation PDR
   ════════════════════════════════════════════════════════════ */

function updatePDR_(e) {
  var fileId       = e.parameter.fileId    || TRAVAUX_SS_ID;
  var sheetName    = e.parameter.sheetName || 'Travaux hebdomadaire';
  var ordre        = e.parameter.ordre     || '';
  var disponibilite= e.parameter.disponibilite || '';
  var justification= e.parameter.justification || '';
  var delai        = e.parameter.delai     || '';
  var observation  = e.parameter.observation || '';
  var dateConf     = e.parameter.dateConf  || new Date().toISOString();

  if (!ordre) return jsonError('ordre manquant');

  // Colonnes (base 0) dans Travaux hebdomadaire :
  // T=19:dispo  U=20:obs/justif  W=22:delai  X=23:dateConf
  var sh   = getSheet(fileId, sheetName);
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === ordre) {
      sh.getRange(i + 1, 20).setValue(disponibilite);  // col T
      sh.getRange(i + 1, 21).setValue(observation || justification); // col U
      sh.getRange(i + 1, 23).setValue(delai);          // col W
      sh.getRange(i + 1, 24).setValue(dateConf);       // col X
      Logger.log('updatePDR OT=' + ordre + ' dispo=' + disponibilite);
      return jsonOk({ success: true });
    }
  }

  return jsonError('OT introuvable pour PDR : ' + ordre);
}


/* ════════════════════════════════════════════════════════════
   ARRÊTS — Marquer comme reporté
   ════════════════════════════════════════════════════════════ */

function markArretReporte_(e) {
  var date         = e.parameter.date         || '';
  var installation = e.parameter.installation || '';
  if (!date || !installation) return jsonError('date et installation requis');

  var sh   = getSheet(TRAVAUX_SS_ID, 'Arrêts');
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowDate = String(data[i][0] || '').trim();
    var rowInst = String(data[i][1] || '').trim();
    if (rowDate === date && rowInst === installation) {
      // Colonne C = statut
      sh.getRange(i + 1, 3).setValue('Reporté');
      return jsonOk({ success: true });
    }
  }
  return jsonError('Arrêt introuvable');
}


/* ════════════════════════════════════════════════════════════
   NOTIFICATIONS BROADCAST
   ════════════════════════════════════════════════════════════ */

function logNotif_(e) {
  var sh     = getSheet(MOTEURS_SS_ID, SH_NOTIFS);
  var notifId       = e.parameter.notifId       || ('N_' + Date.now());
  var title         = e.parameter.title         || '';
  var body          = e.parameter.body          || '';
  var targetProfile = e.parameter.targetProfile || 'admin';
  var timestamp     = new Date().toISOString();

  // Purger les notifications de plus de 7 jours
  try {
    var data     = sh.getDataRange().getValues();
    var cutoff   = Date.now() - 7 * 24 * 3600 * 1000;
    var toDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var ts = new Date(data[i][4] || 0).getTime();
      if (ts < cutoff) toDelete.push(i + 1);
    }
    toDelete.forEach(function(r) { sh.deleteRow(r); });
  } catch(ex) {}

  sh.appendRow([notifId, title, body, targetProfile, timestamp]);
  return jsonOk({ success: true, id: notifId });
}

function getNotifs_(e) {
  var sh = getSheet(MOTEURS_SS_ID, SH_NOTIFS);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return jsonOk([]);
  var notifs = [];
  for (var i = 1; i < data.length; i++) {
    notifs.push({
      id:            String(data[i][0] || ''),
      title:         String(data[i][1] || ''),
      body:          String(data[i][2] || ''),
      targetProfile: String(data[i][3] || ''),
      timestamp:     String(data[i][4] || '')
    });
  }
  return jsonOk(notifs);
}


/* ════════════════════════════════════════════════════════════
   WEB PUSH SUBSCRIPTIONS
   ════════════════════════════════════════════════════════════ */

function saveSubscription_(e) {
  var profile      = e.parameter.profile      || 'autre';
  var endpoint     = e.parameter.endpoint     || '';
  var subscription = e.parameter.subscription || '';
  if (!endpoint) return jsonError('endpoint manquant');

  var sh   = getSheet(MOTEURS_SS_ID, SH_SUBSCRIPTIONS);
  var data = sh.getDataRange().getValues();

  // Mettre à jour si l'endpoint existe déjà
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === endpoint) {
      sh.getRange(i + 1, 1).setValue(profile);
      sh.getRange(i + 1, 3).setValue(subscription);
      return jsonOk({ success: true, updated: true });
    }
  }

  sh.appendRow([profile, endpoint, subscription]);
  return jsonOk({ success: true, created: true });
}


/* ════════════════════════════════════════════════════════════
   TEMPLATE EMAIL HTML
   ════════════════════════════════════════════════════════════ */

function buildEmailHtml(titre, destinataire, champs, intro) {
  var salut = destinataire ? 'Bonjour ' + destinataire + ',' : 'Bonjour,';
  var rows  = champs.map(function(c) {
    var valStyle = c.highlight
      ? 'font-weight:700;color:#16a34a;'
      : 'color:#1f2937;';
    return '<tr style="border-bottom:1px solid #e5e7eb;">'
         + '<td style="padding:9px 16px;color:#6b7280;font-size:13px;white-space:nowrap;">' + c.label + '</td>'
         + '<td style="padding:9px 16px;font-size:13px;' + valStyle + '">' + (c.value || '—') + '</td>'
         + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;">'
       + '<div style="max-width:600px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.08);">'

       // Header
       + '<div style="background:linear-gradient(135deg,#166534,#15803d);padding:28px 32px;">'
       + '<div style="font-size:11px;font-weight:700;color:#bbf7d0;letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">Maintenance Analytics — OCP</div>'
       + '<h2 style="margin:0;color:#fff;font-size:19px;font-weight:700;">' + titre + '</h2>'
       + '</div>'

       // Body
       + '<div style="padding:28px 32px;">'
       + '<p style="margin:0 0 18px;font-size:14px;color:#374151;">' + salut + '</p>'
       + '<p style="margin:0 0 20px;font-size:13px;color:#6b7280;">' + (intro || '') + '</p>'
       + '<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;font-size:13px;">'
       + '<thead><tr style="background:#f9fafb;">'
       + '<th style="padding:10px 16px;text-align:left;color:#374151;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.5px;">Champ</th>'
       + '<th style="padding:10px 16px;text-align:left;color:#374151;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.5px;">Valeur</th>'
       + '</tr></thead>'
       + '<tbody>' + rows + '</tbody>'
       + '</table>'
       + '</div>'

       // Footer
       + '<div style="background:#f9fafb;padding:16px 32px;border-top:1px solid #e5e7eb;">'
       + '<p style="margin:0;font-size:11px;color:#9ca3af;">Message automatique — Maintenance Analytics | Bureau de méthode Daoui · Ne pas répondre à cet email</p>'
       + '</div>'

       + '</div></body></html>';
}
