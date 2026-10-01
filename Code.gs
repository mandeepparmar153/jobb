/**
 * Mod Entreprenad AB – GPS-spårning (backend)
 * Google Apps Script, kopplat till ett Google Sheet.
 *
 * Första gången: kör funktionen setup() en gång från editorn.
 * Sedan: Distribuera > Ny distribution > Webbapp
 *        (Kör som: Jag, Åtkomst: Vem som helst).
 */

const TZ = 'Europe/Stockholm';

const HEAD = {
  Settings: ['key', 'value'],
  Users: ['username', 'name', 'role', 'salt', 'hash', 'deviceId', 'active', 'createdAt'],
  Tokens: ['token', 'username', 'deviceId', 'createdAt'],
  Sites: ['id', 'name', 'address', 'lat', 'lng', 'radius', 'polygon', 'group', 'active'],
  Shifts: ['shiftId', 'username', 'start', 'startTime', 'end', 'endTime', 'workType', 'companions', 'endReason', 'startLat', 'startLng'],
  Events: ['ts', 'time', 'username', 'shiftId', 'type', 'siteId', 'lat', 'lng', 'note'],
};
// Positioner sparas i en egen kalkylfil per månad, en flik per dag
// (så att huvudfilen aldrig når Googles storleksgräns och dagsrapporten går snabbt).
const POINT_HEAD = ['ts', 'username', 'shiftId', 'lat', 'lng', 'acc'];

const DEFAULT_SETTINGS = {
  autoStopTime: '19:00',      // tom = av. Gäller bara pass som startade före denna tid
  maxShiftHours: 14,          // pass stängs automatiskt efter så många timmar
  stopMinutes: 5,             // minst så många minuter för att räknas som "stannade"
  stopRadius: 50,             // inom så många meter
  doneMinutes: 10,            // så länge på en plats innan den räknas som "klar"
  awayMinutes: 30,            // stilla utanför arbetsplats så länge -> fråga "slut?"
  retentionMonths: 6,         // data raderas automatiskt efter så många månader
  workTypes: 'Handarbete,Traktor,Transport,Administration',
  deviceBinding: 'yes',       // ett konto = en telefon
};

/* ------------------------------------------------------------------ */
/* Webbapp                                                             */
/* ------------------------------------------------------------------ */

function doGet() {
  return json_({ ok: true, data: 'Mod Entreprenad GPS API' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = { ok: true, data: handle_(req) };
  } catch (x) {
    out = { ok: false, error: x.code || 'server', message: String(x.message || x) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function err_(code, msg) {
  const e = new Error(msg || code);
  e.code = code;
  return e;
}

const PUBLIC_ACTIONS = { login: true, ping: true };
const ADMIN_ACTIONS = {
  adminDay: true, live: true, listUsers: true, addUser: true, resetPassword: true,
  resetDevice: true, setActive: true, saveSite: true, deleteSite: true, saveSettings: true,
};
const WRITE_ACTIONS = {
  login: true, sync: true, logout: true, changePassword: true, addUser: true, resetPassword: true,
  resetDevice: true, setActive: true, saveSite: true, deleteSite: true, saveSettings: true,
};

function handle_(req) {
  const action = req.action;
  const fn = ACTIONS[action];
  if (!fn) throw err_('unknown', 'Okänd åtgärd: ' + action);
  let user = null;
  if (!PUBLIC_ACTIONS[action]) user = auth_(req.token, !!ADMIN_ACTIONS[action]);
  if (WRITE_ACTIONS[action]) {
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try { return fn(req, user); } finally { lock.releaseLock(); }
  }
  return fn(req, user);
}

/* ------------------------------------------------------------------ */
/* Sheet-hjälp                                                         */
/* ------------------------------------------------------------------ */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Strängar sparas med ' framför så att Sheets inte gör om dem till datum/tal/formler.
function cell_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && v !== '') return "'" + v;
  return v;
}

function rows_(name) {
  const sh = sheet_(name);
  const v = sh.getDataRange().getValues();
  const h = v.shift() || [];
  return v.map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  });
}

function append_(name, obj) {
  sheet_(name).appendRow(HEAD[name].map(function (k) { return cell_(obj[k]); }));
}

function update_(name, row, obj) {
  const sh = sheet_(name);
  const h = HEAD[name];
  const cur = sh.getRange(row, 1, 1, h.length).getValues()[0].map(cell_);
  h.forEach(function (k, j) { if (k in obj) cur[j] = cell_(obj[k]); });
  sh.getRange(row, 1, 1, h.length).setValues([cur]);
}

function rewrite_(name, keepFn) {
  const sh = sheet_(name);
  const v = sh.getDataRange().getValues();
  const h = v.shift();
  const keep = v.filter(keepFn).map(function (r) { return r.map(cell_); });
  if (v.length) sh.getRange(2, 1, v.length, h.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, h.length).setValues(keep);
  return v.length - keep.length;
}

function yes_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 'yes' || v === 1; }

function fmt_(ts) { return Utilities.formatDate(new Date(Number(ts)), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function monthKey_(ts) { return Utilities.formatDate(new Date(Number(ts)), TZ, 'yyyy_MM'); }
function dayKey_(ts) { return Utilities.formatDate(new Date(Number(ts)), TZ, 'yyyy-MM-dd'); }

const BOOKS_ = {};
function pointBook_(ts, create) {
  const key = 'pts_' + monthKey_(ts);
  if (BOOKS_[key]) return BOOKS_[key];
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(key);
  if (id) {
    try { return (BOOKS_[key] = SpreadsheetApp.openById(id)); } catch (e) { /* raderad fil – skapa ny */ }
  }
  if (!create) return null;
  const book = SpreadsheetApp.create('Mod GPS punkter ' + monthKey_(ts).replace('_', '-'));
  book.setSpreadsheetTimeZone(TZ);
  try {
    const parents = DriveApp.getFileById(ss_().getId()).getParents();
    if (parents.hasNext()) DriveApp.getFileById(book.getId()).moveTo(parents.next());
  } catch (e) { /* ligger kvar i Min enhet */ }
  props.setProperty(key, book.getId());
  return (BOOKS_[key] = book);
}

function pointSheet_(ts, create) {
  const book = pointBook_(ts, create);
  if (!book) return null;
  const name = dayKey_(ts);
  let sh = book.getSheetByName(name);
  if (!sh && create) {
    const all = book.getSheets();
    if (all.length === 1 && all[0].getLastRow() === 0 && !/^\d{4}-\d{2}-\d{2}$/.test(all[0].getName())) {
      sh = all[0];
      sh.setName(name);
    } else sh = book.insertSheet(name);
    sh.appendRow(POINT_HEAD);
    sh.setFrozenRows(1);
    if (sh.getMaxColumns() > POINT_HEAD.length) sh.deleteColumns(POINT_HEAD.length + 1, sh.getMaxColumns() - POINT_HEAD.length);
  }
  return sh;
}

function readPoints_(ts, maxRows) {
  const sh = pointSheet_(ts, false);
  if (!sh || sh.getLastRow() < 2) return [];
  let n = sh.getLastRow() - 1;
  if (maxRows && n > maxRows) n = maxRows;
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, POINT_HEAD.length).getValues();
}

/* ------------------------------------------------------------------ */
/* Inställningar                                                       */
/* ------------------------------------------------------------------ */

function settings_() {
  const s = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { s[k] = DEFAULT_SETTINGS[k]; });
  rows_('Settings').forEach(function (r) {
    if (r.key && r.key in DEFAULT_SETTINGS) {
      const d = DEFAULT_SETTINGS[r.key];
      s[r.key] = typeof d === 'number' ? Number(r.value) : String(r.value);
    }
  });
  return s;
}

function setSetting_(key, value) {
  const r = rows_('Settings').find(function (x) { return x.key === key; });
  if (r) update_('Settings', r._row, { value: value });
  else append_('Settings', { key: key, value: value });
}

/* ------------------------------------------------------------------ */
/* Inloggning                                                          */
/* ------------------------------------------------------------------ */

function hash_(salt, pw) {
  let d = salt + '|' + pw;
  for (let i = 0; i < 300; i++) {
    d = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d, Utilities.Charset.UTF_8));
  }
  return d;
}

function randomPassword_(n) {
  const a = 'abcdefghjkmnpqrstuvwxyz23456789';
  let p = '';
  for (let i = 0; i < (n || 8); i++) p += a.charAt(Math.floor(Math.random() * a.length));
  return p;
}

function newToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }

function auth_(token, needAdmin) {
  if (!token) throw err_('auth', 'Inte inloggad');
  const cache = CacheService.getScriptCache();
  let username = cache.get('tok:' + token);
  if (!username) {
    const t = rows_('Tokens').find(function (r) { return r.token === token; });
    if (!t) throw err_('auth', 'Inloggningen har gått ut');
    username = t.username;
    cache.put('tok:' + token, username, 21600);
  }
  const u = rows_('Users').find(function (r) { return r.username === username; });
  if (!u || !yes_(u.active)) throw err_('auth', 'Kontot är avstängt');
  if (needAdmin && u.role !== 'admin') throw err_('forbidden', 'Bara för administratör');
  u.token = token;
  return u;
}

function dropTokens_(username) {
  const cache = CacheService.getScriptCache();
  rows_('Tokens').forEach(function (t) { if (t.username === username) cache.remove('tok:' + t.token); });
  rewrite_('Tokens', function (r) { return r[1] !== username; });
}

function profile_(u) {
  const s = settings_();
  return {
    username: u.username,
    name: u.name,
    role: u.role,
    settings: s,
    workTypes: String(s.workTypes).split(',').map(function (x) { return x.trim(); }).filter(String),
    sites: sitesList_(),
    coworkers: rows_('Users')
      .filter(function (r) { return r.role === 'worker' && yes_(r.active) && r.username !== u.username; })
      .map(function (r) { return { username: r.username, name: r.name }; }),
  };
}

function sitesList_() {
  return rows_('Sites').filter(function (r) { return yes_(r.active); }).map(function (r) {
    let poly = null;
    try { poly = r.polygon ? JSON.parse(r.polygon) : null; } catch (e) { poly = null; }
    return {
      id: String(r.id), name: r.name, address: r.address, group: r.group,
      lat: Number(r.lat), lng: Number(r.lng), radius: Number(r.radius) || 50, polygon: poly,
    };
  });
}

/* ------------------------------------------------------------------ */
/* Åtgärder                                                            */
/* ------------------------------------------------------------------ */

const ACTIONS = {
  ping: function () { return { time: Date.now() }; },

  login: function (req) {
    const username = String(req.username || '').trim().toLowerCase();
    const cache = CacheService.getScriptCache();
    const failKey = 'fail:' + username;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= 10) throw err_('locked', 'För många försök. Vänta 15 minuter.');
    const u = rows_('Users').find(function (r) { return r.username === username; });
    if (!u || !yes_(u.active) || hash_(u.salt, String(req.password || '')) !== u.hash) {
      cache.put(failKey, String(fails + 1), 900);
      throw err_('login', 'Fel användarnamn eller lösenord');
    }
    const s = settings_();
    const deviceId = String(req.deviceId || '');
    if (u.role === 'worker' && yes_(s.deviceBinding)) {
      if (!u.deviceId) update_('Users', u._row, { deviceId: deviceId });
      else if (u.deviceId !== deviceId) throw err_('device', 'Kontot är kopplat till en annan telefon');
    }
    const token = newToken_();
    append_('Tokens', { token: token, username: u.username, deviceId: deviceId, createdAt: Date.now() });
    cache.remove(failKey);
    return { token: token, profile: profile_(u) };
  },

  me: function (req, u) { return profile_(u); },

  logout: function (req, u) {
    CacheService.getScriptCache().remove('tok:' + u.token);
    rewrite_('Tokens', function (r) { return r[0] !== u.token; });
    return true;
  },

  changePassword: function (req, u) {
    if (hash_(u.salt, String(req.oldPassword || '')) !== u.hash) throw err_('login', 'Fel nuvarande lösenord');
    const pw = String(req.newPassword || '');
    if (pw.length < 6) throw err_('weak', 'Lösenordet måste ha minst 6 tecken');
    const salt = Utilities.getUuid();
    update_('Users', u._row, { salt: salt, hash: hash_(salt, pw) });
    if (u.role === 'admin') setSetting_('adminInitialPassword', '(ändrat)');
    return true;
  },

  // Arbetarens telefon skickar en kö av händelser: start, points, event, end.
  sync: function (req, u) {
    const ops = Array.isArray(req.ops) ? req.ops : [];
    const cache = CacheService.getScriptCache();
    const done = [];
    let shifts = null;
    const shiftMap = function () {
      if (!shifts) {
        shifts = {};
        rows_('Shifts').forEach(function (r) { shifts[r.shiftId] = r; });
      }
      return shifts;
    };
    const pointBuckets = {};

    ops.forEach(function (op) {
      if (!op || !op.opId) return;
      if (cache.get('op:' + op.opId)) { done.push(op.opId); return; }

      if (op.type === 'start') {
        if (!shiftMap()[op.shiftId]) {
          const row = {
            shiftId: op.shiftId, username: u.username, start: Number(op.ts), startTime: fmt_(op.ts),
            workType: String(op.workType || ''), companions: (op.companions || []).join(','),
            startLat: op.lat || '', startLng: op.lng || '',
          };
          append_('Shifts', row);
          shifts[op.shiftId] = Object.assign({ _row: sheet_('Shifts').getLastRow() }, row);
        }
      } else if (op.type === 'points') {
        (op.pts || []).forEach(function (p) {
          const ts = Number(p[0]);
          if (!ts || isNaN(Number(p[1])) || isNaN(Number(p[2]))) return;
          const name = dayKey_(ts);
          (pointBuckets[name] = pointBuckets[name] || { ts: ts, rows: [] }).rows.push([
            ts, cell_(u.username), cell_(op.shiftId), Number(p[1]), Number(p[2]), Math.round(Number(p[3]) || 0),
          ]);
        });
      } else if (op.type === 'event') {
        append_('Events', {
          ts: Number(op.ts), time: fmt_(op.ts), username: u.username, shiftId: op.shiftId || '',
          type: String(op.kind || ''), siteId: op.siteId || '', lat: op.lat || '', lng: op.lng || '',
          note: op.note ? String(op.note) : '',
        });
      } else if (op.type === 'end') {
        const s = shiftMap()[op.shiftId];
        if (s && s.username === u.username && !s.end) {
          update_('Shifts', s._row, { end: Number(op.ts), endTime: fmt_(op.ts), endReason: String(op.reason || 'manual') });
          s.end = Number(op.ts);
        }
      }
      cache.put('op:' + op.opId, '1', 21600);
      done.push(op.opId);
    });

    Object.keys(pointBuckets).forEach(function (name) {
      const b = pointBuckets[name];
      const sh = pointSheet_(b.ts, true);
      sh.getRange(sh.getLastRow() + 1, 1, b.rows.length, POINT_HEAD.length).setValues(b.rows);
    });

    return { done: done, serverTime: Date.now() };
  },

  /* ---------------- Admin ---------------- */

  listUsers: function () { return usersPublic_(); },

  adminDay: function (req) {
    const date = String(req.date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
    const from = Utilities.parseDate(date + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss').getTime();
    const next = new Date(from + 36 * 3600000);
    const to = Utilities.parseDate(Utilities.formatDate(next, TZ, 'yyyy-MM-dd') + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss').getTime();

    const shifts = rows_('Shifts').filter(function (r) {
      const s = Number(r.start), e = Number(r.end) || Date.now();
      return s < to && e >= from;
    }).map(function (r) {
      return {
        shiftId: r.shiftId, username: r.username, start: Number(r.start), end: Number(r.end) || null,
        workType: r.workType, companions: r.companions ? String(r.companions).split(',') : [], endReason: r.endReason,
      };
    });

    const events = rows_('Events').filter(function (r) { return Number(r.ts) >= from && Number(r.ts) < to; })
      .map(function (r) {
        return { ts: Number(r.ts), username: r.username, shiftId: r.shiftId, type: r.type, siteId: String(r.siteId), lat: Number(r.lat), lng: Number(r.lng), note: r.note };
      });

    const points = {};
    readPoints_(from).forEach(function (r) {
      const ts = Number(r[0]);
      if (ts < from || ts >= to) return;
      (points[r[1]] = points[r[1]] || []).push([ts, Number(r[3]), Number(r[4]), Number(r[5]) || 0, r[2]]);
    });
    Object.keys(points).forEach(function (k) { points[k].sort(function (a, b) { return a[0] - b[0]; }); });

    return { date: date, from: from, to: to, users: usersPublic_(), sites: sitesList_(), settings: settings_(), shifts: shifts, events: events, points: points, serverTime: Date.now() };
  },

  live: function () {
    const last = {};
    readPoints_(Date.now(), 4000).forEach(function (r) {
      const ts = Number(r[0]);
      if (!last[r[1]] || last[r[1]][0] < ts) last[r[1]] = [ts, Number(r[3]), Number(r[4]), Number(r[5]) || 0];
    });
    const open = rows_('Shifts').filter(function (r) { return !r.end; }).map(function (r) { return { username: r.username, start: Number(r.start), workType: r.workType }; });
    return { last: last, open: open, serverTime: Date.now() };
  },

  addUser: function (req) {
    const username = String(req.username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{2,30}$/.test(username)) throw err_('bad', 'Användarnamn: bara a-z, 0-9, punkt, bindestreck (2–30 tecken)');
    if (rows_('Users').some(function (r) { return r.username === username; })) throw err_('exists', 'Användarnamnet finns redan');
    const pw = randomPassword_(8);
    const salt = Utilities.getUuid();
    append_('Users', {
      username: username, name: String(req.name || username).trim(), role: req.role === 'admin' ? 'admin' : 'worker',
      salt: salt, hash: hash_(salt, pw), deviceId: '', active: true, createdAt: Date.now(),
    });
    return { username: username, password: pw };
  },

  resetPassword: function (req) {
    const u = findUser_(req.username);
    const pw = randomPassword_(8);
    const salt = Utilities.getUuid();
    update_('Users', u._row, { salt: salt, hash: hash_(salt, pw) });
    dropTokens_(u.username);
    return { username: u.username, password: pw };
  },

  resetDevice: function (req) {
    const u = findUser_(req.username);
    update_('Users', u._row, { deviceId: '' });
    dropTokens_(u.username);
    return true;
  },

  setActive: function (req, me) {
    const u = findUser_(req.username);
    if (u.username === me.username) throw err_('bad', 'Du kan inte stänga av ditt eget konto');
    update_('Users', u._row, { active: !!req.active });
    if (!req.active) dropTokens_(u.username);
    return true;
  },

  saveSite: function (req) {
    const s = req.site || {};
    if (!s.name || isNaN(Number(s.lat)) || isNaN(Number(s.lng))) throw err_('bad', 'Namn och position krävs');
    const row = {
      name: String(s.name), address: String(s.address || ''), lat: Number(s.lat), lng: Number(s.lng),
      radius: Math.max(10, Math.min(2000, Number(s.radius) || 50)),
      polygon: s.polygon && s.polygon.length >= 3 ? JSON.stringify(s.polygon) : '',
      group: String(s.group || ''), active: true,
    };
    const existing = s.id ? rows_('Sites').find(function (r) { return String(r.id) === String(s.id); }) : null;
    if (existing) { update_('Sites', existing._row, row); row.id = String(existing.id); }
    else { row.id = 'S' + Date.now().toString(36); append_('Sites', row); }
    return sitesList_();
  },

  deleteSite: function (req) {
    const r = rows_('Sites').find(function (x) { return String(x.id) === String(req.id); });
    if (r) update_('Sites', r._row, { active: false });
    return sitesList_();
  },

  saveSettings: function (req) {
    const s = req.settings || {};
    Object.keys(s).forEach(function (k) { if (k in DEFAULT_SETTINGS) setSetting_(k, s[k]); });
    return settings_();
  },
};

function findUser_(username) {
  const u = rows_('Users').find(function (r) { return r.username === String(username || '').toLowerCase(); });
  if (!u) throw err_('notfound', 'Användaren finns inte');
  return u;
}

function usersPublic_() {
  return rows_('Users').map(function (r) {
    return { username: r.username, name: r.name, role: r.role, active: yes_(r.active), hasDevice: !!r.deviceId };
  });
}

/* ------------------------------------------------------------------ */
/* Automatik (triggers)                                                */
/* ------------------------------------------------------------------ */

// Varje timme: stäng pass som glömts öppna (telefonen stängd, sidan stängd osv).
function autoCloseShifts() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const s = settings_();
    const now = Date.now();
    const open = rows_('Shifts').filter(function (r) { return !r.end; });
    if (!open.length) return;
    const lastPt = {};
    const startDays = {};
    open.forEach(function (r) { startDays[dayKey_(r.start)] = Number(r.start); });
    startDays[dayKey_(now)] = now;
    startDays[dayKey_(now - 86400000)] = now - 86400000;
    Object.keys(startDays).forEach(function (k) {
      readPoints_(startDays[k]).forEach(function (r) {
        const id = r[2], ts = Number(r[0]);
        if (!lastPt[id] || lastPt[id] < ts) lastPt[id] = ts;
      });
    });
    open.forEach(function (r) {
      const start = Number(r.start);
      const last = lastPt[r.shiftId] || start;
      let reason = '';
      if (now - start > Number(s.maxShiftHours) * 3600000) reason = 'auto_maxhours';
      if (s.autoStopTime && /^\d{1,2}:\d{2}$/.test(s.autoStopTime)) {
        const day = Utilities.formatDate(new Date(start), TZ, 'yyyy-MM-dd');
        const stopAt = Utilities.parseDate(day + ' ' + s.autoStopTime + ':00', TZ, 'yyyy-MM-dd HH:mm:ss').getTime();
        if (start < stopAt && now > stopAt + 30 * 60000) reason = 'auto_time';
      }
      if (!reason && now - last > 3 * 3600000) reason = 'auto_nosignal';
      if (reason) {
        const end = Math.max(start, Math.min(last, now));
        update_('Shifts', r._row, { end: end, endTime: fmt_(end), endReason: reason });
      }
    });
  } finally { lock.releaseLock(); }
}

// Varje natt: radera gammal data enligt retentionMonths.
function dailyCleanup() {
  const s = settings_();
  const months = Math.max(1, Number(s.retentionMonths) || 6);
  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - months);
  const cutoff = cutoffDate.getTime();
  const cutoffKey = 'pts_' + monthKey_(cutoff);
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(function (k) {
    if (/^pts_\d{4}_\d{2}$/.test(k) && k < cutoffKey) {
      try { DriveApp.getFileById(props.getProperty(k)).setTrashed(true); } catch (e) { /* redan borta */ }
      props.deleteProperty(k);
    }
  });
  rewrite_('Shifts', function (r) { return Number(r[2]) >= cutoff; });
  rewrite_('Events', function (r) { return Number(r[0]) >= cutoff; });
  rewrite_('Tokens', function (r) { return Number(r[3]) >= Date.now() - 365 * 86400000; });
}

/* ------------------------------------------------------------------ */
/* Installation                                                        */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  ['Settings', 'Users', 'Tokens', 'Sites', 'Shifts', 'Events'].forEach(sheet_);
  const s0 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Blad1');
  if (s0 && ss.getSheets().length > 1 && s0.getLastRow() === 0) ss.deleteSheet(s0);

  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    if (!rows_('Settings').some(function (r) { return r.key === k; })) append_('Settings', { key: k, value: DEFAULT_SETTINGS[k] });
  });

  let msg = 'Klart.';
  if (!rows_('Users').some(function (r) { return r.role === 'admin'; })) {
    const pw = randomPassword_(10);
    const salt = Utilities.getUuid();
    append_('Users', { username: 'admin', name: 'Admin', role: 'admin', salt: salt, hash: hash_(salt, pw), deviceId: '', active: true, createdAt: Date.now() });
    setSetting_('adminInitialPassword', pw);
    msg = 'Klart. Admin-användare: admin  Lösenord: ' + pw + '  (finns också i fliken Settings)';
  }

  ScriptApp.getProjectTriggers().forEach(function (t) {
    const f = t.getHandlerFunction();
    if (f === 'autoCloseShifts' || f === 'dailyCleanup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoCloseShifts').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyCleanup').timeBased().atHour(3).everyDays(1).inTimezone(TZ).create();

  Logger.log(msg);
  return msg;
}
