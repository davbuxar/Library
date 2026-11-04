// School Library – backend. Paste in Sheet > Extensions > Apps Script (file name: Code.gs)
const SS = SpreadsheetApp.getActive();
const TZ = Session.getScriptTimeZone();
const STAFF = ['Admin', 'Librarian'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('School Library')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl(S_().LogoURL).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
const sh_ = n => SS.getSheetByName(n);
const hdr_ = s => s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
const AL_ = { 'Entry Date': 'EntryDate', 'Accession No': 'BookID', 'Place & Publisher': 'Publisher', 'BOOK NO': 'BookNo', 'CLASS No.': 'ClassNo', 'Bill No': 'BillNo', 'Bill Date': 'BillDate', 'Last Verified': 'LastVerified' };
const kc_ = (n, h) => n === 'Books' ? 'BookID' : h[0];
const CN_ = ['Students', 'Teachers', 'Rules', 'Settings', 'Users'];
function rows_(n) {
  const cs = CacheService.getScriptCache(), k = 'r_' + n, hot = CN_.indexOf(n) >= 0;
  if (hot) { const x = cs.get(k); if (x) return JSON.parse(x); }
  const v = sh_(n).getDataRange().getDisplayValues(), h = v.shift();
  const out = v.map((r, i) => { const o = { _r: i + 2 }; h.forEach((q, j) => o[AL_[q] || q] = r[j]); return o; }).filter(o => o[kc_(n, h)] !== '');
  if (hot) try { cs.put(k, JSON.stringify(out), 120); } catch (e) {}
  return out;
}
function add_(n, o) {
  CacheService.getScriptCache().remove('r_' + n);
  const s = sh_(n), h = hdr_(s), r = s.getRange(s.getLastRow() + 1, 1, 1, h.length);
  r.setNumberFormat('@').setValues([h.map(k => { const v = o[AL_[k] || k]; return v === undefined ? '' : v; })]);
}
function set_(n, row, o) {
  CacheService.getScriptCache().remove('r_' + n);
  const s = sh_(n), h = hdr_(s);
  for (const k in o) { const i = h.findIndex(x => (AL_[x] || x) === k); if (i >= 0) s.getRange(row, i + 1).setNumberFormat('@').setValue(o[k]); }
}
const S_ = () => { const o = {}; rows_('Settings').forEach(r => o[r.Key] = r.Value); return o; };
const today_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const addDays_ = (d, n) => Utilities.formatDate(new Date(new Date(d + 'T12:00:00').getTime() + n * 864e5), TZ, 'yyyy-MM-dd');
const diff_ = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 864e5);
const rule_ = t => rows_('Rules').find(r => r.Role === t);
const C_ = {}, cache_ = n => C_[n] || (C_[n] = rows_(n));
function member_(id) {
  id = String(id).trim();
  let m = cache_('Students').find(r => r.AdmNo === id);
  if (m) return { ID: id, Type: 'Student', Name: m.Name, Info: 'Class ' + m.Class + '-' + m.Section + ' Roll ' + m.Roll, Class: m.Class + '-' + m.Section, Email: m.Email || '', Mobile: m.Mobile || '', PhotoLink: m.PhotoLink || '' };
  m = cache_('Teachers').find(r => r.EmpID === id);
  if (m) return { ID: id, Type: 'Teacher', Name: m.Name, Info: m.Department, Class: 'Teachers', Email: m.Email || '', Mobile: m.Mobile || '', PhotoLink: m.PhotoLink || '' };
  return null;
}
function barred_(id) {
  if (!sh_('Barred')) return null;
  return rows_('Barred').find(r => r.MemberID === String(id).trim()) || null;
}
function prof_() {
  let s = sh_('Profiles'); if (!s) s = SS.insertSheet('Profiles');
  const h = s.getLastColumn() ? hdr_(s) : [];
  ['Username', 'Image', 'Mobile', 'Email', 'WhatsApp'].forEach(k => { if (h.indexOf(k) < 0) { s.getRange(1, h.length + 1).setValue(k); h.push(k); } });
  return s;
}
function pmap_() {
  const cs = CacheService.getScriptCache(), c = cs.get('pm'); if (c) return JSON.parse(c);
  const m = {}, it = DriveApp.getFolderById(S_().PhotoFolderId).getFiles();
  while (it.hasNext()) { const f = it.next(); m[f.getName().replace(/\.[^.]+$/, '')] = f.getId(); }
  try { cs.put('pm', JSON.stringify(m), 600); } catch (e) {}
  return m;
}
function photo_(id, link) {
  try {
    const mm = link && /\/d\/([\w-]+)/.exec(link), fid = mm ? mm[1] : pmap_()[id]; if (!fid) return '';
    const cs = CacheService.getScriptCache(), k = 'ph' + id, c = cs.get(k); if (c) return c;
    const f = DriveApp.getFileById(fid), d = 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
    if (d.length < 95000) cs.put(k, d, 3600);
    return d;
  } catch (e) { return ''; }
}

const SESS_DAYS = 7;
function putSess_(t, U) {
  U.exp = Date.now() + SESS_DAYS * 864e5; const j = JSON.stringify(U);
  CacheService.getScriptCache().put(t, j, 21600); PropertiesService.getScriptProperties().setProperty('s_' + t, j);
}
function getSess_(t) {
  if (!t) return null;
  let j = CacheService.getScriptCache().get(t);
  if (!j) { j = PropertiesService.getScriptProperties().getProperty('s_' + t); if (j) CacheService.getScriptCache().put(t, j, 21600); }
  if (!j) return null; const U = JSON.parse(j); return U.exp > Date.now() ? U : null;
}
function cleanSess_() {
  const P = PropertiesService.getScriptProperties(), a = P.getProperties();
  for (const k in a) if (k.indexOf('s_') === 0) { try { if (JSON.parse(a[k]).exp < Date.now()) P.deleteProperty(k); } catch (e) { P.deleteProperty(k); } }
}
function payload_(t, U) {
  const s = S_(); delete s.SmsApiKey; delete s.AnthropicKey;
  const pr = sh_('Profiles') ? rows_('Profiles').find(r => r.Username === U.user) : null;
  return { t, role: U.role, ref: U.ref, name: U.name, s, pic: pr ? pr.Image : '', prof: pr ? { Mobile: pr.Mobile || '', Email: pr.Email || '', WhatsApp: pr.WhatsApp || '' } : {} };
}
function login(u, p) {
  const x = rows_('Users').find(r => r.Username.toLowerCase() === String(u).trim().toLowerCase() && r.Password === p);
  if (!x) throw new Error('Invalid username or password');
  const m = x.RefID ? member_(x.RefID) : null, name = m ? m.Name : x.Username, t = Utilities.getUuid();
  cleanSess_(); const U = { role: x.Role, ref: x.RefID, name, user: x.Username }; putSess_(t, U);
  return payload_(t, U);
}
function resume(t) { const U = getSess_(t); if (!U) throw new Error('Session expired, please login again'); return payload_(t, U); }
function logout(t) { CacheService.getScriptCache().remove(t); PropertiesService.getScriptProperties().deleteProperty('s_' + t); return 1; }
function doPost(e) {   // used when the page is hosted outside Apps Script (e.g. GitHub Pages)
  let out;
  try { const q = JSON.parse(e.postData.contents), F = { login, api, resume, logout }; if (!F[q.fn]) throw new Error('Bad request'); out = { data: F[q.fn].apply(null, q.args || []) }; }
  catch (err) { out = { error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function api(t, a, p) {
  const U = getSess_(t); if (!U) throw new Error('Session expired, please login again');
  const staff = STAFF.indexOf(U.role) >= 0;
  if (!(a in A)) throw new Error('Unknown action');
  if (A[a].staff && !staff) throw new Error('Not allowed');
  if (A[a].admin && U.role !== 'Admin') throw new Error('Admin only');
  return A[a].f(p || {}, U, staff);
}

function issue_(bid, mid) {
  const m = member_(mid); if (!m) throw new Error('Member not found: ' + mid);
  const bar = barred_(m.ID); if (bar) throw new Error('Member is BARRED from issue' + (bar.Reason ? ': ' + bar.Reason : ''));
  const r = rule_(m.Type); if (!r || r.CanIssue !== 'Yes') throw new Error('Issue not allowed for ' + m.Type);
  const b = rows_('Books').find(x => x.BookID === bid); if (!b) throw new Error('Book not found: ' + bid);
  if (b.Type === 'Reference') throw new Error('Reference books cannot be issued');
  if (b.Status !== 'Available') throw new Error('Book is ' + b.Status);
  const open = rows_('Issues').filter(i => i.MemberID === m.ID && i.Status === 'Issued'), t = today_();
  if (open.length >= +r.MaxBooks) throw new Error(m.Type + ' limit reached (' + r.MaxBooks + ' books)');
  if (open.some(i => i.DueDate < t)) throw new Error('Member has overdue books – return them first');
  const due = addDays_(t, +(b.Type === 'Magazine' ? r.MagazineDays : r.Days));
  add_('Issues', { IssueID: 'I' + Date.now(), BookID: bid, MemberID: m.ID, MemberName: m.Name, Title: b.Title, IssueDate: t, DueDate: due, Status: 'Issued' });
  set_('Books', b._r, { Status: 'Issued' });
  return 'Issued "' + b.Title + '" to ' + m.Name + '. Due: ' + due;
}

const A = {
  books: { f: () => rows_('Books') },
  mine: { f: (p, U) => rows_('Issues').filter(i => i.MemberID === U.ref) },
  table: { staff: 1, f: p => { if (['Issues', 'Books', 'Damaged', 'Requests', 'Donations', 'WriteOff'].indexOf(p.n) < 0) throw new Error('Bad table'); if (!sh_(p.n)) return []; return rows_(p.n); } },
  members: { staff: 1, f: () => rows_('Students').map(s => member_(s.AdmNo)).concat(rows_('Teachers').map(s => member_(s.EmpID))) },
  member: { staff: 1, f: p => { const m = member_(p.id); if (!m) throw new Error('Member not found'); m.rule = rule_(m.Type); m.open = rows_('Issues').filter(i => i.MemberID === m.ID && i.Status === 'Issued'); m.photo = photo_(m.ID, m.PhotoLink); m.barred = barred_(m.ID); return m; } },
  bar: { staff: 1, f: (p, U) => { const m = member_(p.id); if (!m) throw new Error('Member not found');
    if (barred_(m.ID)) throw new Error('Already barred');
    if (!sh_('Barred')) SS.insertSheet('Barred').appendRow(['MemberID', 'MemberName', 'Reason', 'Date', 'By']);
    add_('Barred', { MemberID: m.ID, MemberName: m.Name, Reason: p.reason || '', Date: today_(), By: U.name }); return m.Name + ' is now barred from issue'; } },
  unbar: { staff: 1, f: p => { const b = barred_(p.id); if (!b) throw new Error('Not barred'); sh_('Barred').deleteRow(b._r); return 'Bar removed'; } },
  barlist: { staff: 1, f: () => sh_('Barred') ? rows_('Barred') : [] },
  cards: { staff: 1, f: p => p.ids.map(id => { const m = member_(id); m.photo = photo_(id, m.PhotoLink); return m; }) },
  issue: { staff: 1, f: p => issue_(p.bookId, p.memberId) },
  ret: { staff: 1, f: p => {
    const i = rows_('Issues').find(x => x.BookID === p.bookId && x.Status === 'Issued'); if (!i) throw new Error('This book is not issued');
    const b = rows_('Books').find(x => x.BookID === p.bookId), m = member_(i.MemberID), r = rule_(m ? m.Type : 'Student'), t = today_();
    const fine = Math.max(0, diff_(i.DueDate, t)) * (r ? +r.FinePerDay : 0);
    set_('Issues', i._r, { ReturnDate: t, Fine: fine, Status: 'Returned' });
    let d = ''; if (p.damaged) d = dmg_(b, i.MemberID, i.MemberName, 'Damaged', p.remarks, p.extra, i.IssueID); else set_('Books', b._r, { Status: 'Available' });
    return { msg: 'Received "' + b.Title + '". Late fine: ' + fine, dmg: d };
  } },
  declare: { staff: 1, f: p => {   // damaged / missing, book not coming back through the counter
    const b = rows_('Books').find(x => x.BookID === p.bookId); if (!b) throw new Error('Book not found');
    const i = rows_('Issues').find(x => x.BookID === b.BookID && x.Status === 'Issued'); let mid = String(p.memberId || '').trim(), mn = '';
    if (i) { mid = i.MemberID; mn = i.MemberName; set_('Issues', i._r, p.type === 'Missing' ? { Status: 'Missing' } : { ReturnDate: today_(), Status: 'Returned' }); }
    else if (mid) { const m = member_(mid); if (!m) throw new Error('Member not found'); mn = m.Name; }
    return { msg: p.type + ' recorded', dmg: dmg_(b, mid, mn, p.type, p.remarks, p.extra, i ? i.IssueID : '') };
  } },
  rmDmg: { staff: 1, f: p => {
    const d = rows_('Damaged').find(x => x._r === +p.row); if (!d) throw new Error('Record not found');
    const b = rows_('Books').find(x => x.BookID === d.BookID); if (b && ['Damaged', 'Missing'].indexOf(b.Status) >= 0) set_('Books', b._r, { Status: 'Available' });
    sh_('Damaged').deleteRow(d._r); return 'Removed. Book is available again.';
  } },
  writeoff: { staff: 1, f: (p, U) => {
    const b = rows_('Books').find(x => x.BookID === String(p.bookId).trim()); if (!b) throw new Error('Book not found');
    if (b.Status === 'Issued') throw new Error('Book is issued – receive it first'); if (b.Status === 'Written Off') throw new Error('Already written off');
    if (!sh_('WriteOff')) SS.insertSheet('WriteOff').appendRow(['WoID', 'Date', 'BookID', 'Title', 'Author', 'Price', 'Reason', 'OrderNo', 'By']);
    add_('WriteOff', { WoID: 'W' + Date.now(), Date: today_(), BookID: b.BookID, Title: b.Title, Author: b.Author, Price: b.Price, Reason: p.reason || '', OrderNo: p.order || '', By: U.name });
    set_('Books', b._r, { Status: 'Written Off', Remarks: 'Written off ' + today_() + (p.reason ? ': ' + p.reason : '') }); return 'Book written off – it can no longer be issued';
  } },
  invData: { staff: 1, f: p => {
    const d = rows_('Damaged').find(x => x.DmgID === p.id); if (!d) throw new Error('Record not found');
    return { d, b: rows_('Books').find(x => x.BookID === d.BookID) || {}, st: rows_('Students').find(x => x.AdmNo === d.MemberID) || null, tc: rows_('Teachers').find(x => x.EmpID === d.MemberID) || null, iss: d.IssueID ? rows_('Issues').find(x => x.IssueID === d.IssueID) || null : null };
  } },
  quota: { f: (p, U) => {
    const m = member_(U.ref); if (!m) return { max: 0, issued: 0, pending: 0, left: 0 }; const r = rule_(m.Type) || { MaxBooks: 0 };
    const iss = rows_('Issues').filter(i => i.MemberID === U.ref && i.Status === 'Issued').length, pen = rows_('Requests').filter(q => q.MemberID === U.ref && q.Status === 'Pending').length;
    return { max: +r.MaxBooks, issued: iss, pending: pen, left: Math.max(0, +r.MaxBooks - iss - pen) };
  } },
  hist: { f: (p, U, staff) => {
    const id = staff && p.id ? String(p.id).trim() : U.ref, m = member_(id || ''); if (!m) throw new Error('Member not found');
    const y = +p.sess || sessYear_(today_()), a = y + '-04-01', b = (y + 1) + '-03-31', inR = d => d >= a && d <= b;
    return { m: { ID: m.ID, Name: m.Name, Type: m.Type }, from: a, to: b,
      issues: rows_('Issues').filter(i => i.MemberID === id && inR(i.IssueDate)),
      dmg: sh_('Damaged') ? rows_('Damaged').filter(d => d.MemberID === id && inR(d.Date)) : [],
      reqs: rows_('Requests').filter(r => r.MemberID === id && inR(r.ReqDate)) };
  } },
  summary: { staff: 1, f: p => {
    const b = rows_('Books').find(x => x.BookID === p.id); if (!b) throw new Error('Book not found'); const s = S_(); let web = '', txt = '';
    const q = b.ISBN ? 'isbn:' + String(b.ISBN).replace(/[^0-9Xx]/g, '') : 'intitle:' + b.Title + (b.Author ? ' inauthor:' + b.Author : '');
    try { const j = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/books/v1/volumes?maxResults=1&q=' + encodeURIComponent(q), { muteHttpExceptions: true }).getContentText()); const v = j.items && j.items[0] && j.items[0].volumeInfo; if (v && v.description) web = v.description; } catch (e) {}
    if (s.AnthropicKey) {
      try {
        const r = JSON.parse(UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-api-key': s.AnthropicKey, 'anthropic-version': '2023-06-01' },
          payload: JSON.stringify({ model: s.AiModel || 'claude-haiku-5-5', max_tokens: 400, messages: [{ role: 'user', content: 'Write a neutral 3-4 sentence summary of this book for school library readers. Do not invent facts; if unsure keep it general.\nTitle: ' + b.Title + '\nAuthor: ' + b.Author + '\nISBN: ' + (b.ISBN || '') + '\nPublisher: ' + (b.Publisher || '') + '\nYear: ' + (b.Year || '') + '\nCategory: ' + (b.Category || '') + (web ? '\nDescription found online: ' + web.slice(0, 1200) : '') }] }) }).getContentText());
        if (r.content && r.content[0]) txt = r.content[0].text;
      } catch (e) {}
    }
    if (!txt) txt = web.slice(0, 700);
    if (!txt) throw new Error('No summary found online. Add AnthropicKey in Settings, or type the summary manually (Edit).');
    ens_('Books', ['Summary']); set_('Books', b._r, { Summary: txt }); return txt;
  } },
  payDmg: { staff: 1, f: p => { set_('Damaged', +p.row, { Status: 'Paid' }); return 1; } },
  verify: { staff: 1, f: p => { const b = rows_('Books').find(x => x.BookID === p.id); if (!b) throw new Error('Not found: ' + p.id); set_('Books', b._r, { LastVerified: today_() }); return b.Title + ' (' + b.Status + ') – Rack ' + b.Rack + '/' + b.Shelf; } },
  req: { f: (p, U) => {
    if (!U.ref) throw new Error('Only students/teachers can request');
    const bar = barred_(U.ref); if (bar) throw new Error('You are barred from issuing books' + (bar.Reason ? ': ' + bar.Reason : '') + '. Contact the librarian.');
    const mm = member_(U.ref), rr = mm && rule_(mm.Type), used = rows_('Issues').filter(i => i.MemberID === U.ref && i.Status === 'Issued').length + rows_('Requests').filter(q => q.MemberID === U.ref && q.Status === 'Pending').length;
    if (rr && used >= +rr.MaxBooks) throw new Error('Limit reached: maximum ' + rr.MaxBooks + ' books (issued + requested). Return a book first.');
    const b = rows_('Books').find(x => x.BookID === p.id);
    if (!b || b.Status !== 'Available' || b.Type === 'Reference') throw new Error('Book is not available');
    add_('Requests', { ReqID: 'R' + Date.now(), BookID: b.BookID, Title: b.Title, MemberID: U.ref, MemberName: U.name, ReqDate: today_(), Status: 'Pending' });
    set_('Books', b._r, { Status: 'Reserved' });
    return 'Request sent. The book is reserved for you – collect it from the library.';
  } },
  reqs: { f: (p, U, staff) => rows_('Requests').filter(r => staff || r.MemberID === U.ref) },
  decide: { staff: 1, f: p => {
    const r = rows_('Requests').find(x => x._r === +p.row), b = rows_('Books').find(x => x.BookID === r.BookID);
    set_('Books', b._r, { Status: 'Available' });
    if (p.ok) { try { issue_(r.BookID, r.MemberID); } catch (e) { set_('Books', b._r, { Status: 'Reserved' }); throw e; } set_('Requests', r._r, { Status: 'Approved' }); return 'Approved & issued'; }
    set_('Requests', r._r, { Status: 'Rejected' }); return 'Rejected';
  } },
  getProfile: { f: (p, U) => { const r = sh_('Profiles') ? rows_('Profiles').find(x => x.Username === U.user) : null; if (!r) return {}; return { Mobile: r.Mobile || '', Email: r.Email || '', WhatsApp: r.WhatsApp || '' }; } },
  setProfile: { f: (p, U) => {
    prof_(); const r = rows_('Profiles').find(x => x.Username === U.user), d = {};
    if (p.image) { if (p.image.length > 45000) throw new Error('Image too large'); d.Image = p.image; }
    const M = { mobile: 'Mobile', email: 'Email', whatsapp: 'WhatsApp' };
    for (const k in M) if (p[k] !== undefined) d[M[k]] = String(p[k]).trim();
    if (r) set_('Profiles', r._r, d); else add_('Profiles', Object.assign({ Username: U.user }, d));
    if (U.ref && (d.Mobile || d.Email)) {
      const tn = rows_('Students').some(x => x.AdmNo === U.ref) ? 'Students' : 'Teachers', row = rows_(tn).find(x => x[tn === 'Students' ? 'AdmNo' : 'EmpID'] === U.ref), u = {};
      if (d.Mobile) u.Mobile = d.Mobile; if (d.Email) u.Email = d.Email; if (row) set_(tn, row._r, u);
    }
    return 'Profile saved';
  } },
  chkImg: { staff: 1, f: () => {
    CacheService.getScriptCache().remove('pm'); const pm = pmap_(); let found = 0, total = 0; const miss = [];
    [['Students', 'AdmNo'], ['Teachers', 'EmpID']].forEach(x => {
      const s = sh_(x[0]); const h = hdr_(s); let ci = h.indexOf('PhotoLink');
      if (ci < 0) { ci = h.length; s.getRange(1, ci + 1).setValue('PhotoLink'); }
      CacheService.getScriptCache().remove('r_' + x[0]);
      const R = rows_(x[0]), vals = []; for (let i = 0; i < s.getLastRow() - 1; i++) vals.push(['']);
      R.forEach(r => { total++; const f = pm[String(r[x[1]]).trim()]; if (f) { found++; vals[r._r - 2][0] = 'https://drive.google.com/file/d/' + f + '/view'; } else miss.push(r[x[1]] + ' ' + (r.Name || '')); });
      if (vals.length) s.getRange(2, ci + 1, vals.length, 1).setNumberFormat('@').setValues(vals);
      CacheService.getScriptCache().remove('r_' + x[0]);
    });
    return { found, total, miss: miss.slice(0, 80) };
  } },
  saveBook: { staff: 1, f: p => {
    const d = p.data, id = String(d.BookID || '').trim(); if (!id) throw new Error('Accession No is required'); ens_('Books', ['ISBN', 'Summary']);
    const b = rows_('Books').find(x => x.BookID === id);
    if (p.isNew) { if (b) throw new Error('Accession No already exists'); d.BookID = id; if (!d.Status) d.Status = 'Available'; if (!d.EntryDate) d.EntryDate = today_(); add_('Books', d); return 'Book added'; }
    if (!b) throw new Error('Book not found');
    const s = sh_('Books'), h = hdr_(s), cur = s.getRange(b._r, 1, 1, h.length).getDisplayValues()[0];
    h.forEach((k, j) => { const ik = AL_[k] || k; if (ik !== 'BookID' && ik !== 'LastVerified' && ik in d) cur[j] = d[ik]; });
    s.getRange(b._r, 1, 1, h.length).setNumberFormat('@').setValues([cur]); return 'Book updated';
  } },
  donate: { staff: 1, f: (p, U) => {
    const id = String(p.bookId || '').trim(), per = String(p.person || '').trim();
    if (!id || !per) throw new Error('Accession No and person name are required');
    const b = rows_('Books').find(x => x.BookID === id); let title;
    if (p.dir === 'Gifted') {
      if (!b) throw new Error('Book not found'); if (b.Status !== 'Available') throw new Error('Book is ' + b.Status + ', cannot be gifted');
      set_('Books', b._r, { Status: 'Gifted', Remarks: 'Gifted to ' + per + ' on ' + today_() }); title = b.Title;
    } else if (b) { set_('Books', b._r, { Source: 'Donation: ' + per }); title = b.Title; }
    else { if (!p.title) throw new Error('Title is required for a new donated book'); add_('Books', { BookID: id, EntryDate: today_(), Title: p.title, Author: p.author || '', Price: p.price || '0', Type: 'Book', Status: 'Available', Source: 'Donation: ' + per }); title = p.title; }
    if (!sh_('Donations')) SS.insertSheet('Donations').appendRow(['DonID', 'Date', 'Direction', 'BookID', 'Title', 'Person', 'Contact', 'Remarks', 'By']);
    add_('Donations', { DonID: 'N' + Date.now(), Date: today_(), Direction: p.dir, BookID: id, Title: title, Person: per, Contact: p.contact || '', Remarks: p.remarks || '', By: U.name });
    return p.dir === 'Gifted' ? 'Gift recorded' : 'Donation recorded';
  } },
  overdue: { staff: 1, f: () => overdue_() },
  remcfg: { staff: 1, f: () => { const s = S_(); return { auto: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoRemind'), ch: s.AutoChannel || 'email', quota: MailApp.getRemainingDailyQuota(), sms: !!s.SmsApiKey }; } },
  remind: { staff: 1, f: p => { const set = {}; p.ids.forEach(i => set[i] = 1); return send_(overdue_().filter(o => set[o.IssueID]), p.ch); } },
  autoOn: { admin: 1, f: p => { delTrig_(); ScriptApp.newTrigger('autoRemind').timeBased().everyDays(1).atHour(+p.hour || 9).create(); setSet_('AutoChannel', p.ch || 'email'); return 'Daily auto-reminder enabled'; } },
  autoOff: { admin: 1, f: () => { delTrig_(); return 'Auto-reminder disabled'; } },
  heads: { admin: 1, f: p => hdr_(sh_(chk_(p.n))) },
  import: { admin: 1, f: p => {
    const n = chk_(p.n), s = sh_(n), h = hdr_(s), L = {}; h.forEach(k => L[k.toLowerCase()] = k);
    const v = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, h.length).getDisplayValues() : [], idx = {};
    const ki = h.indexOf(n === 'Books' ? 'Accession No' : h[0]); v.forEach((r, i) => idx[r[ki]] = i);
    let add = 0, upd = 0, skip = 0;
    p.rows.forEach(o0 => {
      const o = {}; for (const k in o0) { const c = L[k.trim().toLowerCase()]; if (c) o[c] = String(o0[k]).trim(); }
      const key = o[h[ki]]; if (!key) { skip++; return; }
      if (key in idx) { const r = v[idx[key]]; h.forEach((k, j) => { if (o[k]) r[j] = o[k]; }); upd++; }
      else { v.push(h.map(k => o[k] || (n === 'Books' && k === 'Status' ? 'Available' : ''))); idx[key] = v.length - 1; add++; }
    });
    if (v.length) s.getRange(2, 1, v.length, h.length).setNumberFormat('@').setValues(v);
    CacheService.getScriptCache().remove('r_' + n);
    return { add, upd, skip };
  } }
};

// ---------- Reminders & import helpers ----------
const chk_ = n => { if (['Students', 'Teachers', 'Books', 'Users'].indexOf(n) < 0) throw new Error('Sheet not allowed'); return n; };
function setSet_(k, v) { const r = rows_('Settings').find(x => x.Key === k); if (r) set_('Settings', r._r, { Value: v }); else add_('Settings', { Key: k, Value: v }); }
function delTrig_() { ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoRemind').forEach(t => ScriptApp.deleteTrigger(t)); }
function overdue_() {
  const t = today_(), R = rows_('Rules');
  return rows_('Issues').filter(i => i.Status === 'Issued' && i.DueDate < t).map(i => {
    const m = member_(i.MemberID) || {}, r = R.find(x => x.Role === m.Type) || {}, days = diff_(i.DueDate, t);
    return { IssueID: i.IssueID, MemberID: i.MemberID, Name: i.MemberName, Title: i.Title, DueDate: i.DueDate, days, fine: days * (+r.FinePerDay || 0), Email: m.Email || '', Mobile: m.Mobile || '', PhotoLink: m.PhotoLink || '' };
  });
}
function msg_(o, s) {
  return (s.ReminderText || 'Dear {name}, library book "{title}" was due on {due} and is {days} day(s) overdue. Fine: {fine}. Please return it. - {school} Library')
    .replace(/\{(\w+)\}/g, (m, k) => { const d = { name: o.Name, title: o.Title, due: o.DueDate, days: o.days, fine: (s.Currency || '₹') + o.fine, school: s.SchoolName }; return d[k] === undefined ? m : d[k]; });
}
function sms_(num, text, s) {   // Fast2SMS "quick" route. Change here for MSG91 / other gateway.
  if (!s.SmsApiKey) throw new Error('SmsApiKey missing in Settings');
  const r = UrlFetchApp.fetch('https://www.fast2sms.com/dev/bulkV2', { method: 'post', headers: { authorization: s.SmsApiKey },
    payload: { route: 'q', message: text, numbers: String(num).replace(/\D/g, '').slice(-10), language: 'english', flash: 0 }, muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText()); if (!j['return']) throw new Error(j.message || 'SMS failed');
}
function send_(list, ch) {
  const s = S_(), log = [], t = today_(); let em = 0, sm = 0, fail = 0;
  list.forEach(o => {
    const text = msg_(o, s);
    if (ch !== 'sms' && o.Email) { let r = 'OK'; try { MailApp.sendEmail({ to: o.Email, subject: 'Library book overdue: ' + o.Title, body: text, name: s.SchoolName + ' Library' }); em++; } catch (e) { r = String(e.message); fail++; } log.push([t, o.IssueID, o.MemberID, 'Email', r]); }
    if (ch !== 'email' && o.Mobile) { let r = 'OK'; try { sms_(o.Mobile, text, s); sm++; } catch (e) { r = String(e.message); fail++; } log.push([t, o.IssueID, o.MemberID, 'SMS', r]); }
  });
  if (log.length) {
    if (!sh_('ReminderLog')) SS.insertSheet('ReminderLog').appendRow(['Date', 'IssueID', 'MemberID', 'Channel', 'Result']);
    const l = sh_('ReminderLog'); l.getRange(l.getLastRow() + 1, 1, log.length, 5).setNumberFormat('@').setValues(log);
  }
  return 'Emails sent: ' + em + ' | SMS sent: ' + sm + ' | Failed: ' + fail;
}
function autoRemind() {   // runs daily from the time trigger
  const s = S_(), gap = +(s.ReminderGapDays || 3), t = today_(), L = sh_('ReminderLog') ? rows_('ReminderLog') : [];
  send_(overdue_().filter(o => !L.some(l => l.IssueID === o.IssueID && l.Result === 'OK' && diff_(l.Date, t) < gap)), s.AutoChannel || 'email');
}


// ---------- Damage / missing helpers ----------
const sessYear_ = d => +d.slice(0, 4) - (+d.slice(5, 7) < 4 ? 1 : 0);
function ens_(n, cols) {
  const s = sh_(n), h = hdr_(s);
  cols.forEach(k => { if (h.indexOf(k) < 0) { s.getRange(1, h.length + 1).setValue(k); h.push(k); } });
}
function invNo_() {
  const y = sessYear_(today_()), P = PropertiesService.getScriptProperties(), k = 'inv' + y, n = (+P.getProperty(k) || 0) + 1;
  P.setProperty(k, String(n)); return 'LIB/' + y + '-' + String(y + 1).slice(2) + '/' + String(n).padStart(3, '0');
}
function dmg_(b, mid, mname, type, remarks, extra, issueId) {
  ens_('Damaged', ['Type', 'BookCost', 'ExtraFine', 'IssueID', 'InvNo']);
  const cost = Math.round((+b.Price || 0) * +(S_().DamageChargePercent || 100) / 100), ex = Math.max(0, +extra || 0), id = 'D' + Date.now();
  add_('Damaged', { DmgID: id, BookID: b.BookID, Title: b.Title, MemberID: mid || '', MemberName: mname || '', Date: today_(), Type: type || 'Damaged', Remarks: remarks || '', BookCost: cost, ExtraFine: ex, Charge: cost + ex, Status: 'Due', IssueID: issueId || '', InvNo: invNo_() });
  set_('Books', b._r, { Status: type === 'Missing' ? 'Missing' : 'Damaged' });
  return id;
}
