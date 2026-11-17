// ห้องเรียนออนไลน์ v3 — ล็อกอินด้วยรหัสผ่าน + ตรวจสิทธิ์ฝั่งเซิร์ฟเวอร์
const TEACHER_CODE = 'CHANGE-ME'; // รหัสลับสำหรับสมัครบัญชีครู — เปลี่ยนก่อนใช้งานจริง!
const ALLOW_STUDENT_SIGNUP = false; // false = นักเรียนสมัครเองไม่ได้ ครูเป็นผู้เพิ่มรายชื่อ
const SS_ID = ''; // ว่าง = ใช้ Spreadsheet ที่ผูกกับสคริปต์
const T = {
  classes: ['id', 'name', 'code', 'teacher'],
  members: ['id', 'cls', 'name'],
  asgs:    ['id', 'cls', 'title', 'desc', 'due', 'max'],
  subs:    ['id', 'asg', 'name', 'text', 'at', 'score', 'fb', 'fname', 'furl'],
  users:   ['name', 'role', 'salt', 'hash', 'sid'] // sid = รหัสนักเรียน (ใช้ล็อกอินแทนชื่อได้)
};

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const deny = () => ({ ok: false, error: 'ไม่มีสิทธิ์ทำรายการนี้' });
const ss = () => SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActive();

function sh(t) {
  let s = ss().getSheetByName(t);
  if (!s) { s = ss().insertSheet(t); s.appendRow(T[t]); s.setFrozenRows(1); }
  else if (s.getLastColumn() < T[t].length) s.getRange(1, 1, 1, T[t].length).setValues([T[t]]); // เพิ่มหัวคอลัมน์ใหม่ให้ชีตเดิมอัตโนมัติ
  return s;
}

function rows(t) {
  const v = sh(t).getDataRange().getValues(), h = v.shift();
  return v.filter(x => x[0] !== '').map(x => {
    const o = {};
    h.forEach((k, i) => {
      const c = x[i];
      o[k] = c instanceof Date ? Utilities.formatDate(c, 'Asia/Bangkok', k === 'due' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm") : c;
    });
    ['id', 'cls', 'asg'].forEach(k => { if (k in o) o[k] = String(o[k]); });
    return o;
  });
}

// ---------- รหัสผ่านและโทเคน ----------
const b64 = x => Utilities.base64EncodeWebSafe(x);

function secret() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
const sign = s => b64(Utilities.computeHmacSha256Signature(s, secret()));

function hash(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < 300; i++) h = b64(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}

function mkToken(u) {
  const p = b64(JSON.stringify({ n: u.name, r: u.role, e: Date.now() + 30 * 864e5 })); // อายุ 30 วัน
  return p + '.' + sign(p);
}

function who(tok) {
  try {
    const a = String(tok).split('.');
    if (sign(a[0]) !== a[1]) return null;
    const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(a[0])).getDataAsString());
    return o.e > Date.now() ? { name: o.n, role: o.r } : null;
  } catch (e) { return null; }
}

function register(b) {
  const n = String(b.name || '').trim(), pw = String(b.password || '');
  if (!n || pw.length < 4) return { ok: false, error: 'ชื่อว่าง หรือรหัสผ่านสั้นกว่า 4 ตัว' };
  const role = b.role === 'teacher' ? 'teacher' : 'student';
  if (role === 'student' && !ALLOW_STUDENT_SIGNUP) return { ok: false, error: 'นักเรียนไม่ต้องสมัครเอง ให้ใช้ชื่อและรหัสผ่านที่ครูแจกให้' };
  if (role === 'teacher' && b.code !== TEACHER_CODE) return { ok: false, error: 'รหัสสมัครครูไม่ถูกต้อง' };
  if (rows('users').some(u => String(u.name).toLowerCase() === n.toLowerCase())) return { ok: false, error: 'ชื่อนี้ถูกใช้แล้ว' };
  const salt = Utilities.getUuid();
  sh('users').appendRow([n, role, salt, hash(pw, salt), '']);
  return { ok: true, token: mkToken({ name: n, role: role }), user: { name: n, role: role } };
}

function login(b) {
  const n = String(b.name || '').trim().toLowerCase();
  const u = rows('users').filter(x => n && (String(x.name).toLowerCase() === n || (x.sid !== '' && String(x.sid).toLowerCase() === n)))[0]; // ล็อกอินด้วยชื่อหรือรหัสนักเรียน
  if (!u || hash(String(b.password || ''), u.salt) !== u.hash) return { ok: false, error: 'ชื่อหรือรหัสผ่านไม่ถูกต้อง' };
  return { ok: true, token: mkToken(u), user: { name: u.name, role: u.role } };
}

// ---------- อ่านข้อมูล (กรองตามสิทธิ์) ----------
function state(u) {
  const C = rows('classes'), M = rows('members'), A = rows('asgs'), S = rows('subs');
  if (u.role === 'teacher') {
    const cs = C.filter(c => c.teacher === u.name), ids = cs.map(c => c.id);
    const as = A.filter(a => ids.indexOf(a.cls) >= 0), aid = as.map(a => a.id);
    return { classes: cs, members: M.filter(m => ids.indexOf(m.cls) >= 0), asgs: as, subs: S.filter(s => aid.indexOf(s.asg) >= 0) };
  }
  const mine = M.filter(m => m.name === u.name), ids = mine.map(m => m.cls);
  return { classes: C.filter(c => ids.indexOf(c.id) >= 0), members: mine, asgs: A.filter(a => ids.indexOf(a.cls) >= 0), subs: S.filter(s => s.name === u.name) };
}

function doGet(e) {
  const u = who(e.parameter.token);
  return out(u ? state(u) : { error: 'auth' });
}

// ---------- เขียนข้อมูล (ตรวจสิทธิ์ทุกครั้ง) ----------
function join(u, b) {
  if (u.role !== 'student') return { ok: false, error: 'เฉพาะนักเรียน' };
  const code = String(b.code || '').trim().toUpperCase();
  const c = rows('classes').filter(x => String(x.code).toUpperCase() === code)[0];
  if (!c) return { ok: false, error: 'ไม่พบรหัสห้องนี้' };
  if (!rows('members').some(m => m.cls === c.id && m.name === u.name)) sh('members').appendRow([Utilities.getUuid().slice(0, 8), c.id, u.name]);
  return { ok: true, cls: c.id };
}

function upload(b) {
  const it = DriveApp.getFoldersByName('ห้องเรียน-ไฟล์งาน');
  const dir = it.hasNext() ? it.next() : DriveApp.createFolder('ห้องเรียน-ไฟล์งาน');
  const file = dir.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), b.type || 'application/octet-stream', b.name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: file.getUrl() };
}

function write(u, b) {
  const t = b.table, r = b.row || {}, del = b.op === 'del', isT = u.role === 'teacher';
  if (['classes', 'asgs', 'subs'].indexOf(t) < 0) return deny();
  const ex = rows(t).filter(x => x.id === String(r.id))[0];
  const mine = rows('classes').filter(c => c.teacher === u.name).map(c => c.id);
  let row = r;
  if (t === 'classes') {
    if (!isT || (ex && ex.teacher !== u.name)) return deny();
    row = Object.assign({}, r, { teacher: u.name });
  } else if (t === 'asgs') {
    if (!isT || mine.indexOf(r.cls) < 0 || (ex && mine.indexOf(ex.cls) < 0)) return deny();
  } else {
    const a = rows('asgs').filter(x => x.id === String(ex ? ex.asg : r.asg))[0];
    if (!a) return deny();
    if (isT) { // ครู: ให้คะแนน/ลบได้เฉพาะงานในห้องตัวเอง
      if (!ex || mine.indexOf(a.cls) < 0) return deny();
      row = del ? ex : Object.assign({}, ex, { score: r.score, fb: r.fb });
    } else {   // นักเรียน: ส่งงานของตัวเองได้ และห้ามแก้หลังครูให้คะแนนแล้ว
      const ok = rows('members').some(m => m.cls === a.cls && m.name === u.name);
      if (del || !ok || (ex && (ex.name !== u.name || ex.score !== ''))) return deny();
      row = Object.assign({}, r, { name: u.name, score: '', fb: '' });
    }
  }
  const s = sh(t), h = T[t];
  const i = s.getRange(1, 1, s.getLastRow(), 1).getValues().flat().map(String).indexOf(String(row.id));
  if (del) { if (i > 0) s.deleteRow(i + 1); }
  else {
    const a = h.map(k => row[k] == null ? '' : row[k]);
    if (i > 0) s.getRange(i + 1, 1, 1, h.length).setValues([a]); else s.appendRow(a);
  }
  return { ok: true };
}

// ---------- จัดการรายชื่อและรหัสผ่าน ----------
const pin = () => String(Math.floor(100000 + Math.random() * 900000)); // รหัสสุ่ม 6 หลัก
const ownClass = (u, id) => rows('classes').some(c => c.id === String(id) && c.teacher === u.name);

function setPw(name, pw) {
  const s = sh('users'), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === String(name).toLowerCase()) {
      const salt = Utilities.getUuid();
      s.getRange(i + 1, 3, 1, 2).setValues([[salt, hash(pw, salt)]]);
      return true;
    }
  }
  return false;
}

// ครูเพิ่มรายชื่อเข้าห้อง: สร้างบัญชีนักเรียนให้ (รหัสสุ่ม 4 หลักถ้าไม่ระบุ) และใส่เข้าห้อง
function setSid(name, sid) {
  const s = sh('users'), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === String(name).toLowerCase()) { s.getRange(i + 1, 5).setNumberFormat('@').setValue(sid); return; }
  }
}

function roster(u, b) {
  if (u.role !== 'teacher' || !ownClass(u, b.cls)) return deny();
  const us = rows('users'), mem = rows('members'), res = { created: [], existing: [], skipped: [] };
  (b.list || []).forEach(it => {
    const n = String(it.name || '').trim(), sid = String(it.sid || '').trim();
    // หาบัญชีเดิมจากรหัสนักเรียนก่อน แล้วค่อยหาจากชื่อ → เรียนหลายวิชาใช้บัญชีเดียวกัน
    let x = sid ? us.filter(y => String(y.sid) === sid)[0] : null;
    if (!x && n) {
      x = us.filter(y => String(y.name).toLowerCase() === n.toLowerCase())[0];
      if (x && sid && x.sid !== '' && String(x.sid) !== sid) { res.skipped.push(n + ' (ชื่อซ้ำกับรหัสอื่น)'); return; }
    }
    if (x && x.role !== 'student') { res.skipped.push(n || sid); return; }
    if (!x && !n) { res.skipped.push(sid + ' (ไม่มีชื่อ)'); return; }
    const nm = x ? String(x.name) : n;
    if (!x) {
      const pw = String(it.pw || '').trim() || pin(), salt = Utilities.getUuid(), s = sh('users');
      s.appendRow([nm, 'student', salt, hash(pw, salt), '']);
      if (sid) s.getRange(s.getLastRow(), 5).setNumberFormat('@').setValue(sid);
      us.push({ name: nm, role: 'student', sid: sid });
      res.created.push({ sid: sid, name: nm, pw: pw });
    } else {
      if (sid && x.sid === '') { setSid(nm, sid); x.sid = sid; }
      res.existing.push(nm);
    }
    if (!mem.some(m => m.cls === String(b.cls) && String(m.name) === nm)) {
      sh('members').appendRow([Utilities.getUuid().slice(0, 8), b.cls, nm]);
      mem.push({ cls: String(b.cls), name: nm });
    }
  });
  return Object.assign({ ok: true }, res);
}

// ครูรีเซ็ตรหัสผ่านนักเรียนในห้องของตัวเอง
function resetpw(u, b) {
  const n = String(b.name || '');
  const ok = rows('members').some(m => String(m.name) === n && ownClass(u, m.cls));
  if (u.role !== 'teacher' || !ok) return deny();
  const pw = String(b.pw || '').trim() || pin();
  setPw(n, pw);
  return { ok: true, pw: pw };
}

// ผู้ใช้เปลี่ยนรหัสผ่านของตัวเอง
function chpw(u, b) {
  const x = rows('users').filter(y => String(y.name).toLowerCase() === u.name.toLowerCase())[0];
  const np = String(b.new || '');
  if (!x || hash(String(b.old || ''), x.salt) !== x.hash || np.length < 4)
    return { ok: false, error: 'รหัสผ่านเดิมไม่ถูกต้อง หรือรหัสใหม่สั้นกว่า 4 ตัว' };
  setPw(u.name, np);
  return { ok: true };
}

// ครูลบนักเรียนออกจากห้อง (งานที่ส่งแล้วยังอยู่)
function unenroll(u, b) {
  if (u.role !== 'teacher' || !ownClass(u, b.cls)) return deny();
  const s = sh('members'), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i > 0; i--) {
    if (String(v[i][1]) === String(b.cls) && String(v[i][2]) === String(b.name)) s.deleteRow(i + 1);
  }
  return { ok: true };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.op === 'register') return out(register(b));
    if (b.op === 'login') return out(login(b));
    const u = who(b.token);
    if (!u) return out({ ok: false, error: 'auth' });
    if (b.op === 'upload') return out(upload(b));
    if (b.op === 'join') return out(join(u, b));
    if (b.op === 'roster') return out(roster(u, b));
    if (b.op === 'resetpw') return out(resetpw(u, b));
    if (b.op === 'chpw') return out(chpw(u, b));
    if (b.op === 'unenroll') return out(unenroll(u, b));
    return out(write(u, b));
  } catch (err) {
    return out({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}
