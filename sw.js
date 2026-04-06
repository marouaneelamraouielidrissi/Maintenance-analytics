// Service Worker — Maintenance Analytics OCP
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
// ── Web Push : notification reçue même app fermée ────────────
self.addEventListener('push', e => {
  const data = e.data ? e.data.json() : {};
  const title = data.title || 'Notification';
  const body  = data.body  || '';
  e.waitUntil((async () => {
    await self.registration.showNotification(title, { body, icon: '/icon-192.png', badge: '/favicon-32.png' });
    // Informer les pages ouvertes pour mettre à jour l'inbox
    const cls = await self.clients.matchAll({ type: 'window' });
    cls.forEach(c => c.postMessage({ type: 'PUSH_RECEIVED', title, body }));
  })());
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window' }).then(list => {
    if (list.length) return list[0].focus();
    return self.clients.openWindow('/');
  }));
});

// ── IndexedDB helpers ─────────────────────────────────────────
const IDB_NAME  = 'ma_sw_db';
const IDB_STORE = 'kv';

function idbOpen() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(IDB_NAME, 1);
    r.onupgradeneeded = e => e.target.result.createObjectStore(IDB_STORE);
    r.onsuccess = e => res(e.target.result);
    r.onerror   = () => rej('idb error');
  });
}
async function idbGet(key) {
  try {
    const db = await idbOpen();
    return new Promise(res => {
      const req = db.transaction(IDB_STORE, 'readonly').objectStore(IDB_STORE).get(key);
      req.onsuccess = () => res(req.result ?? null);
      req.onerror   = () => res(null);
    });
  } catch { return null; }
}
async function idbSet(key, val) {
  try {
    const db = await idbOpen();
    return new Promise(res => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(val, key);
      tx.oncomplete = () => res();
      tx.onerror    = () => res();
    });
  } catch {}
}

// ── Periodic Background Sync ──────────────────────────────────
// Déclenché par le navigateur même si la page est fermée (Android Chrome PWA)
self.addEventListener('periodicsync', e => {
  if (e.tag === 'check-ma-notif') {
    e.waitUntil(bgCheckNotifications());
  }
});

async function bgCheckNotifications() {
  try {
    // Si la page est ouverte, elle gère les notifications elle-même
    const openClients = await self.clients.matchAll({ type: 'window' });
    if (openClients.length > 0) return;

    const profile    = await idbGet('profile');
    const webappUrl  = await idbGet('webapp_url');
    if (!profile || profile === 'autre' || !webappUrl) return;

    let state = (await idbGet('notif_state')) || {};
    const isFirstRun = !state.lastCheck;
    const newState   = { seen: { ...(state.seen || {}) }, lastCheck: Date.now() };

    const resp = await fetch(webappUrl + '?t=' + Date.now());
    if (!resp.ok) return;
    const demandes = await resp.json();
    if (!Array.isArray(demandes)) return;

    const toShow = [];

    for (const d of demandes) {
      if (!d.id) continue;

      // ── Profil ÉLECTRIQUE ───────────────────────────────────
      if (profile === 'electrique') {
        if (d.statut === 'approuvee' || d.statut === 'refusee') {
          const key = d.id + '_' + d.statut;
          if (!isFirstRun && !state.seen?.[key]) {
            const label = d.statut === 'approuvee' ? 'approuvée ✅' : 'refusée ❌';
            toShow.push({ title: 'Demande ' + label, body: (d.installation || '') + ' — ' + (d.objetTechnique || '') });
          }
          newState.seen[key] = true;
        }
        if (d.etatReparation === 'repare') {
          const key = d.id + '_repare';
          if (!isFirstRun && !state.seen?.[key]) {
            toShow.push({ title: '🔧 Moteur réparé', body: (d.installation || '') + ' — ' + (d.objetTechnique || '') });
          }
          newState.seen[key] = true;
        }
      }

      // ── Profil ADMIN ────────────────────────────────────────
      if (profile === 'admin') {
        const keyNew = d.id + '_new';
        if (!isFirstRun && !state.seen?.[keyNew]) {
          toShow.push({ title: '🆕 Nouvelle demande moteur', body: (d.demandeur || '') + ' — ' + (d.installation || '') });
        }
        newState.seen[keyNew] = true;
      }
    }

    await idbSet('notif_state', newState);

    // ── Notifications broadcastées (OT réalisé, PDR confirmé…) ─
    try {
      const notifResp = await fetch(webappUrl + '?action=getNotifs&t=' + Date.now());
      if (notifResp.ok) {
        const notifEvents = await notifResp.json();
        if (Array.isArray(notifEvents)) {
          let st2 = (await idbGet('notif_state')) || {};
          const ns2 = { seen: { ...(st2.seen || {}) }, lastCheck: st2.lastCheck || 0 };
          const isFirst2 = !st2.lastCheck || st2.lastCheck === 0;

          for (const ev of notifEvents) {
            if (!ev.id) continue;
            if (ev.targetProfile !== profile) continue;
            const evKey = 'ev_' + ev.id;
            if (!isFirst2 && !st2.seen?.[evKey]) {
              toShow.push({ title: ev.title || '', body: ev.body || '' });
            }
            ns2.seen[evKey] = true;
          }
          await idbSet('notif_state', ns2);
        }
      }
    } catch(e) {}

    for (const n of toShow) {
      await self.registration.showNotification(n.title, { body: n.body, icon: '/icon-192.png' });
    }
  } catch(e) {}
}
