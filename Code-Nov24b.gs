/**
 * PSR Reunion Hub — Google Sheet backend (version Nov 24, 2026 b)
 * --------------------------------------
 * 1. Put in the three PINs below (letters or numbers). Each one must be different.
 * 2. Save, then run "setup" once (Run button) and approve the permissions.
 * 3. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * 4. Send the Web app URL to Claude to connect the site.
 */
// HOST PIN: the host who runs this year's reunion (Setup, payments, expenses, check-in).
// After a hand-off, the new host makes their own PIN when they sign up; it is saved behind the scenes.
const ADMIN_PIN = 'CHANGE-ME';
// PSR REUNION BOARD PIN: works all year, even after the reunion is over. The board can look and print,
// and change the board email. Nothing else can be changed with it.
const COMMITTEE_PIN = 'CHANGE-ME';
// The tech PIN: for the person who looks after the system (TAM). It does everything the host can do,
// it keeps working after a new host takes over, and it is the only PIN that can use Test mode.
const TECH_PIN = 'CHANGE-ME';
// In practice, the "ask to host" request goes here (to you) instead of the PSR Reunion Board.
// For someone you invite to try it, it goes to their own email.
const TECH_EMAIL = 'tamajette@gmail.com';
// SAFETY SWITCH: while this says false, NOTHING is ever sent or shared to the PSR Reunion Board email.
// Anything meant for the board goes to TECH_EMAIL instead, marked "Would have gone to the board".
// Change it to true only when the board has approved the system and you go live.
const BOARD_LIVE = false;
// How many days after the last reunion day the page locks and waits for the next host.
const HANDOFF_DAYS = 10;
const REUNION = 'PSR Family Reunion';

const TABS = {
  Registrations: ['code','date','first','last','email','phone','city','state','payer','people','shirts','addons','donation','total','created','emailOK'],
  Payments:      ['id','date','from','via','memo','amount','code','household','method','fee','note','proofUrl','created','takenBy','balance'],
  Expenses:      ['id','date','vendor','category','amount','source','note','receiptUrl','created'],
  'Other Money In': ['id','date','name','amount','via','receivedBy','note','proofUrl','created','type','fund'],   // same order as the old Contributions tab
  Log:           ['date','time','kind','text'],
  'Check-ins':   ['id','day','code','household','name','walkin','by','created'],
  Emails:        ['date','time','direction','who','email','code','subject','note','forwardedTo'],
  Contacts:      ['email','first','last','phone','address','city','state','zip','source','optOut','added','lastSeen'],
  Settings:      ['key','value']
};
const TZ = 'America/New_York';

// What each column is called in the sheet (the code still uses the short names on the left).
const LABELS = {
  code: 'Confirmation #', date: 'Date', first: 'First Name', last: 'Last Name', email: 'Email', phone: 'Phone',
  city: 'City', state: 'State', payer: 'Paid By', people: 'People', shirts: 'Shirts', addons: 'Add-ons',
  donation: 'Donation', total: 'Total', created: 'Entered On', id: 'ID', from: 'From', via: 'Via', memo: 'Memo',
  amount: 'Amount', method: 'Method', fee: 'Fee', note: 'Note', proofUrl: 'Proof Link', vendor: 'Vendor',
  category: 'Category', source: 'Paid From', receiptUrl: 'Receipt Link', name: 'Name', receivedBy: 'Received By',
  time: 'Time', kind: 'Kind', text: 'What Happened', key: 'Key', value: 'Value', type: 'Type',
  day: 'Day', household: 'Household', walkin: 'Walk-in', by: 'Checked In By', emailOK: 'Email Confirmed',
  what: 'What It Was', fund: 'Goes To', address: 'Street Address', zip: 'Zip Code', source: 'Came From', optOut: 'No Emails', added: 'Added On', lastSeen: 'Last Registered'
};
const TAB_LABELS = { Emails: { time: 'Time', direction: 'Sent or Received', who: 'To / From', email: 'Email', code: 'Confirmation #', subject: 'Subject', note: 'Message (first lines)', forwardedTo: 'Sent On To' }, Payments: { date: "Date Rec'd", via: 'Paid Via', fee: 'Processing Fee', takenBy: 'Taken By', balance: 'Balance Left' },
                     'Other Money In': { date: "Date Rec'd", via: 'Paid Via', type: 'What It Was', name: 'From' }, Expenses: { date: 'Date Paid' }, Log: { kind: 'Status' } };
function label_(tab, k) { return (TAB_LABELS[tab] && TAB_LABELS[tab][k]) || LABELS[k] || k; }
function keyOf_(tab, h) {
  h = String(h).trim();
  const ks = TABS[tab] || [];
  for (let i = 0; i < ks.length; i++) if (h === ks[i] || h === label_(tab, ks[i])) return ks[i];
  for (let i = 0; i < ks.length; i++) if (h === LABELS[ks[i]] || h.toLowerCase() === ks[i].toLowerCase()) return ks[i];   // older column names
  return h;
}
function stamp_() { return Utilities.formatDate(new Date(), TZ, 'MM/dd/yyyy h:mm a'); }
const MONEY_COLS = ['donation','total','amount','fee','balance'];
let TEST_ = false;          // true while the committee is in Test mode
// Practice as a new host uses its own spreadsheet ("PSR Reunion Hub – Practice"), with the same tabs as the real one.
// Nothing done in practice touches the real reunion records.
let PRAC_ = false;
const PRAC_NAME = 'PSR Reunion Hub – Practice';
const TEST_PREFIX = 'Test ';        // old Test tabs inside the real sheet; setup() moves them to the test spreadsheet
const TEST_NAME = 'PSR Reunion Hub – Test';
const JSON_COLS = ['people','shirts','addons'];
const NUM_COLS = ['id','donation','total','amount','fee','balance'];

/* ---------- one-time setup ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  // Committee contributions and other money in now share one tab.
  if (ss.getSheetByName('Contributions') && !ss.getSheetByName('Other Money In')) ss.getSheetByName('Contributions').setName('Other Money In');
  if (ss.getSheetByName('Test Contributions') && !ss.getSheetByName('Test Other Money In')) ss.getSheetByName('Test Contributions').setName('Test Other Money In');
  // The Summary tabs repeated what the page shows, so they're removed.
  ['Summary', TEST_PREFIX + 'Summary'].forEach(n => { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
  moveTestTabs_(ss);
  // The Replies tab became the Emails tab (sent and received together). An empty old Replies tab is removed.
  ['', TEST_PREFIX].forEach(p => { const old = ss.getSheetByName(p + 'Replies'); if (old && !ss.getSheetByName(p + 'Emails') && old.getLastRow() <= 1) ss.deleteSheet(old); });
  Object.keys(TABS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    addMissingCols_(sh, name); formatTab_(sh, name, '#e3efe6');
    if (name === 'Registrations') plainTab_(sh, name);
  });
  fillHouseholds_();
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  try { const it = folder_().getFiles(); while (it.hasNext()) { const f = it.next(); try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {} } } catch (e) {}
  ss.setSpreadsheetTimeZone(TZ);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendQueuedEmails').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendQueuedEmails').timeBased().everyMinutes(15).create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkReplies').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkReplies').timeBased().everyMinutes(10).create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyCheck').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(7).create();
  ['Payments', 'Expenses', 'Other Money In'].forEach(name => { const sh = ss.getSheetByName(name); if (sh) linkifyTab_(sh, name); });
  const lg = ss.getSheetByName('Log'); if (lg) wordifyLog_(lg);
  orderTabs_();
  Logger.log('Setup done. Emails left today: ' + MailApp.getRemainingDailyQuota());
  if (ADMIN_PIN === 'CHANGE-ME') Logger.log('Remember to change ADMIN_PIN at the top before deploying.');
  if (TECH_PIN === 'CHANGE-ME') Logger.log('Remember to set TECH_PIN at the top, so you can use Test mode.');
  Logger.log('Test spreadsheet: ' + testSS_().getUrl());
}

// The red Test tabs used to sit inside the real sheet. They now live in their own spreadsheet
// ("PSR Reunion Hub – Test"), so the committee's sheet has no Test tabs. Their rows are copied over once.
function moveTestTabs_(ss) {
  const ts = testSS_();
  Object.keys(TABS).forEach(name => {
    if (name === 'Settings') return;
    const old = ss.getSheetByName(TEST_PREFIX + name); if (!old) return;
    const dest = ts.getSheetByName(name) || ts.insertSheet(name);
    if (dest.getLastRow() <= 1 && old.getLastRow() > 0) {
      const v = old.getDataRange().getValues();
      dest.clear(); dest.getRange(1, 1, v.length, v[0].length).setValues(v);
    }
    addMissingCols_(dest, name); formatTab_(dest, name, '#f8e1e1'); dest.setTabColor('#c0392b');
    ss.deleteSheet(old);
  });
  const blank = ts.getSheetByName('Sheet1'); if (blank && blank.getLastRow() === 0 && ts.getSheets().length > 1) ts.deleteSheet(blank);
}
function formatTab_(sh, name, color) {
  const cols = TABS[name], n = cols.length;
  sh.getRange(1, 1, 1, n).setValues([cols.map(k => label_(name, k))]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground(color);
  sh.getRange(1, 1, sh.getMaxRows(), n).setNumberFormat('@').setHorizontalAlignment('left').setVerticalAlignment('top');
  sh.autoResizeColumns(1, n);
  cols.forEach((k, i) => {
    const c = i + 1;
    if (MONEY_COLS.indexOf(k) >= 0) sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('$#,##0.00');
    if (['note', 'text', 'people', 'addons', 'shirts', 'value'].indexOf(k) >= 0) { sh.getRange(1, c, sh.getMaxRows(), 1).setWrap(true); sh.setColumnWidth(c, 260); }
    else if (sh.getColumnWidth(c) < 110) sh.setColumnWidth(c, 110);
    if (k === 'id') sh.hideColumns(c);
    if (k === 'type') sh.setColumnWidth(c, 200);
    if (URL_COLS.indexOf(k) >= 0) sh.setColumnWidth(c, 100);
    if (name === 'Emails' && k === 'subject') { sh.setColumnWidth(c, 280); sh.getRange(1, c, sh.getMaxRows(), 1).setWrap(true); }
    if (name === 'Emails' && (k === 'email' || k === 'who')) sh.setColumnWidth(c, 200);
  });
  if (name === 'Check-ins') ciTotals_(sh);
  // dates written month first (10/17/2026)
  const di = cols.indexOf('date');
  if (di >= 0 && sh.getLastRow() > 1) {
    const rg = sh.getRange(2, di + 1, sh.getLastRow() - 1, 1), v = rg.getValues();
    rg.setValues(v.map(r => [toMDY_(r[0])]));
  }
}
function toMDY_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, TZ, 'MM/dd/yyyy');
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(x || '')); return m ? m[2] + '/' + m[3] + '/' + m[1] : x;
}
function toISO_(x) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(x || ''));
  return m ? m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2) : x;
}


/* ---------- a view-only copy of the Test tabs, to share with someone trying the system ----------
   Run makeTestView once. It makes a separate spreadsheet that shows the Test tabs live
   (no contributions, no real data). Open it, click "Allow access" once, then share it as Viewer. */
function makeTestView() {
  const id = testSS_().getId();
  const v = SpreadsheetApp.create('PSR Reunion Hub – Test view');
  const tabs = ['Registrations', 'Payments', 'Expenses', 'Check-ins'];
  tabs.forEach((name, i) => {
    const sh = i === 0 ? v.getSheets()[0].setName(name) : v.insertSheet(name);
    sh.getRange('A1').setFormula('=IMPORTRANGE("' + id + '","\'' + name + '\'!A:Z")');
    sh.setFrozenRows(1);
  });
  PropertiesService.getScriptProperties().setProperty('TEST_VIEW_URL', v.getUrl());
  Logger.log('Test view made. Open it here, click Allow access, then share it as Viewer: ' + v.getUrl());
}

/* ---------- web endpoints ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.act === 'praccept' || p.act === 'prdeny') {
    let msg; try { msg = withLock_(() => practiceAnswer_(String(p.k || ''), String(p.t || ''), p.act === 'praccept')); } catch (err) { msg = { ok: false, text: 'Something went wrong: ' + err.message }; }
    return HtmlService.createHtmlOutput('<div style="font:18px/1.6 Arial,sans-serif;max-width:560px;margin:60px auto;padding:0 20px;color:#17221b">' +
      '<p style="color:#8a6a00;font-weight:bold;letter-spacing:.06em">PRACTICE</p><h1 style="font-size:26px;color:' + (msg.ok ? '#1c5a34' : '#a3252b') + '">' + (msg.ok ? 'Done' : 'Nothing was changed') + '</h1><p>' + msg.text + '</p></div>').setTitle('PSR registration page – practice');
  }
  if (p.act === 'undo' || p.act === 'accept' || p.act === 'deny') {
    let msg;
    try { msg = withLock_(() => (p.act === 'accept' ? acceptHost_ : p.act === 'deny' ? denyHost_ : undoHost_)(String(p.t || ''))); } catch (err) { msg = { ok: false, text: 'Something went wrong: ' + err.message }; }
    return HtmlService.createHtmlOutput('<div style="font:18px/1.6 Arial,sans-serif;max-width:560px;margin:60px auto;padding:0 20px;color:#17221b">' +
      '<h1 style="font-size:26px;color:' + (msg.ok ? '#1c5a34' : '#a3252b') + '">' + (msg.ok ? 'Done' : 'Nothing was changed') + '</h1><p>' + msg.text + '</p>' +
      '<p><a href="' + SITE_URL + '" style="color:#1c5a34">Go to the registration page</a></p></div>').setTitle('PSR registration page');
  }
  return json_({ ok: true, service: 'PSR Reunion Hub', ready: ADMIN_PIN !== 'CHANGE-ME' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  const action = req.action, d = req.data || {};
  const pin = String(req.pin || '').trim();
  TEST_ = false; PRAC_ = false; TRY_ = null; REAL_SS_ = null; READS_ = (action === 'all' || action === 'public') ? {} : null;
  try {
    const locked = handoffOpen_(settings_());
    // ADMIN_PIN (typed in above) is the first host's PIN. Once a new host signs up with their own PIN, it stops working.
    const tech = !!pin && TECH_PIN !== 'CHANGE-ME' && pin === String(TECH_PIN);
    const hostPin = !!pin && (tech || (ADMIN_PIN !== 'CHANGE-ME' && pin === String(ADMIN_PIN) && !PropertiesService.getScriptProperties().getProperty('HOST_PIN')) || hostPinOk_(pin));
    const committee = !hostPin && !!pin && COMMITTEE_PIN !== 'CHANGE-ME' && pin === String(COMMITTEE_PIN);
    const viewer = false;   // (the old try-only test PIN was removed; the tech PIN covers testing)
    // A tester's own PIN (sent by email from "Invite someone to try it"): practice only, in their own tryout sheet.
    let trial = !hostPin && !committee && !!pin && !!(TRY_ = tryByPin_(pin));
    // A PIN made in practice ("You're approved! Create your PIN") works too: it opens that practice as its host.
    let prHost = null;
    if (!hostPin && !committee && !trial && pin && (prHost = practiceHostByPin_(pin))) { trial = true; if (prHost.k.indexOf('try-') === 0) TRY_ = tryById_(prHost.k.slice(4)); }
    if (pin && !hostPin && !committee && !trial && pinPaused_()) return json_({ ok: false, paused: true, error: 'Too many wrong tries. Please wait 10 minutes and try again.' });
    if (!pin && req.try) { TRY_ = tryById_(req.try); if (!TRY_) return json_({ ok: false, tryEnded: true, error: 'This practice link has ended.' }); }
    if (pin && !hostPin && !committee && !trial) wrongPin_();
    const admin = tech || trial || (hostPin && !locked);  // once the reunion is over, the past host's PIN stops working (the tech PIN keeps working)
    if (hostPin && !tech && locked && action !== 'public' && !/^(newHost|forgotPin|resetPin)$/.test(action))
      return json_({ ok: false, ended: true, error: 'This reunion is over, so this PIN no longer works. The next host can start from the registration page.' });
    // Practice: the host's practice run, or the practice sign-up and check-in links made from it.
    const practice = (!!req.practice || !!TRY_) && !viewer && !committee && (admin || action === 'public' || action === 'register' || /^ci[A-Z]/.test(action));
    if (practice || trial) PRAC_ = true;   // the practice PIN always works in practice, whatever the page asks for
    if (trial && /^(saveBoard|clearTest|newHost|forgotPin|resetPin|invite|listTryouts|endTryout)$/.test(action)) return json_({ ok: false, error: 'That isn’t available in a tryout.' });
    else if (viewer) TEST_ = true;
    else if (req.test && tech) TEST_ = true;      // only the tech PIN can use Test mode
    if (action === 'public') return json_(publicData_());
    if (action === 'newHost')   return json_(TEST_ || PRAC_ ? { ok: false, error: 'Turn Test mode off first.' } : withLock_(() => newHost_(d)));
    if (action === 'forgotPin') return json_(forgotPin_(d));
    if (action === 'resetPin')  return json_(withLock_(() => resetPin_(d)));
    if (action === 'practicePin') return json_(withLock_(() => practicePin_(d)));
    if (committee) {
      if (action === 'saveBoard') return json_(withLock_(() => saveBoard_(d)));
      if (action !== 'all') return json_({ ok: false, error: 'The committee PIN can only look and print. Nothing can be changed with it.' });
      const out = withLock_(() => adminAction_('all', d));
      out.role = 'committee'; out.board = boardEmails_(); out.handoff = handoffInfo_(); delete out.settings.checkinCode; delete out.settings.zeffyUrl;
      return json_(out);
    }
    const ciCode = currentCheckinCode_();
    const helper = !!(ciCode && String(req.ci || '').trim().toLowerCase() === ciCode.toLowerCase());
    if (action === 'register') return json_(withLock_(() => register_(d, admin || viewer || helper)));
    if (/^ci[A-Z]/.test(action)) {
      if (!admin && !viewer && !helper) return json_({ ok: false, ciAuth: false, error: 'That check-in code isn’t right.' });
      if (action === 'ciFind') return json_(ciFind_(d));
      return json_(withLock_(() => {
        switch (action) {
          case 'ciSave':  return ciSave_(d, admin || viewer);
          case 'ciUndo':  return ciUndo_(d);
          case 'ciPay':   return ciPay_(d, admin || viewer);
          case 'ciEmail': return ciEmail_(d);
          case 'ciAdd':   return ciAdd_(d);
          default: throw new Error('Unknown action.');
        }
      }));
    }
    if (!admin && !viewer) return json_({ ok: false, auth: false, error: 'Committee sign-in needed.' });
    if (viewer && ['saveSettings', 'clearTest', 'addContribution', 'clearPractice'].indexOf(action) >= 0) return json_({ ok: false, error: 'That isn’t available with the test PIN.' });
    if (viewer && action === 'delete' && d.sheet === 'Other Money In') return json_({ ok: false, error: 'That isn’t available with the test PIN.' });
    if (/^(invite|listTryouts|endTryout)$/.test(action)) {
      if (!tech) return json_({ ok: false, error: 'Only the tech PIN can do that.' });
      PRAC_ = false; TEST_ = false;
      return json_(withLock_(() => action === 'invite' ? inviteTester_(d) : action === 'endTryout' ? endTryout_(d.id) : { ok: true, tryouts: listTryouts_() }));
    }
    if (action === 'practiceRequest' || action === 'practiceStatus') {
      if (!PRAC_) return json_({ ok: false, error: 'That only works in practice.' });
      return json_(action === 'practiceRequest' ? withLock_(() => practiceRequest_(d)) : practiceStatus_());
    }
    if (action === 'saveBoard') return json_(tech ? withLock_(() => saveBoard_(d)) : { ok: false, error: 'Only the PSR Reunion Board can change this email.' });
    let out = withLock_(() => adminAction_(action, d));
    if (action === 'all') { out.handoff = handoffInfo_(); out.host = host_(); out.board = boardEmails_(); out.practice = PRAC_; out.test = TEST_; if (tech) { out.role = 'tech'; out.tryouts = listTryouts_(); } if (trial) { out.role = 'trial'; out.tryEnds = TRY_ ? TRY_.expires : ''; if (prHost) out.prHostName = prHost.name; } }
    if (viewer && action === 'all') {
      out = Object.assign({}, out, { role: 'viewer', contribs: (out.contribs || []).filter(c => c.type && c.type !== 'Committee contribution'), sheetUrl: PropertiesService.getScriptProperties().getProperty('TEST_VIEW_URL') || '',
        log: (out.log || []).filter(l => !/contribution/i.test(l.text || '')) });
      out.settings = Object.assign({}, out.settings, { begin: 0, scholStart: 0, checkinCode: '', caterMin: 2000, foodPer: 30, depPSR: 500, depFaison: 0, deposit: 500, replyTo: '' }); delete out.settings.alertEmails; delete out.settings.boardEmails; delete out.host; delete out.board;
    }
    if (req.log && req.log.x) log_(req.log.k || 'info', req.log.x);
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function adminAction_(a, d) {
  const now = stamp_();
  switch (a) {
    case 'all':
      return { ok: true, sheetUrl: ss_().getUrl(), settings: ensureCheckinCode_(), regs: read_('Registrations'), pays: read_('Payments'),
               exps: read_('Expenses'), contribs: read_('Other Money In'), log: readLog_().reverse().slice(0, 80),
               emails: read_('Emails').reverse().slice(0, 500), contacts: read_('Contacts'), checkins: checkins_(), alerts: health_() };
    case 'addPayment': {
      const proofUrl = saveFile_(d.proof, d.proofName || 'payment');
      append_('Payments', Object.assign({}, d, { proofUrl: proofUrl, created: now, takenBy: d.takenBy || 'Committee' }));
      setBalance_(d.id, d.code, d.total);
      return { ok: true, proofUrl: proofUrl };
    }
    case 'addExpense': {
      const receiptUrl = saveFile_(d.receipt, d.receiptName || 'receipt');
      append_('Expenses', Object.assign({}, d, { receiptUrl: receiptUrl, created: now }));
      return { ok: true, receiptUrl: receiptUrl };
    }
    case 'importContacts': return importContacts_(d.rows || [], d.source || 'Zeffy');
    case 'setOptOut': return setOptOut_(d);
    case 'clearPractice': {           // "Start practice over": empties every tab of the practice spreadsheet
      const keep = PRAC_; PRAC_ = true;
      try { const ps = practiceSS_(); ps.getSheets().forEach(sh => { if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); }); }
      finally { PRAC_ = keep; }
      PROPS_().deleteProperty('PRAC_REQ_' + pracKey_());   // the practice host request and practice PIN start over too
      return { ok: true };
    }
    case 'sendHelperLink': {
      if (!/^\S+@\S+\.\S+$/.test(d.to || '')) throw new Error('That email address doesn’t look right.');
      const link = String(d.link || ''); if (link.indexOf(SITE_URL) !== 0) throw new Error('That link isn’t right. Refresh the page and try again.');
      const s = settings_(), title = [s.edition, fam_(s)].filter(Boolean).join(' ');
      const st = mail_({ to: d.to, subject: 'Your check-in link for the ' + title,
        body: 'Hello,\n\nThank you for helping with check-in! Open this link on your phone, tablet or laptop to check families in at the reunion. Nothing to type, just tap it:\n' + link + '\n\nThe link stops working the day after the reunion.\n\n' + senderName_(s),
        htmlBody: '<div style="font:16px/1.6 Arial,sans-serif;color:#1f2a24;max-width:560px"><p>Hello,</p><p>Thank you for helping with check-in! Tap the button on your phone, tablet or laptop to check families in at the reunion. There’s nothing to type.</p>' +
          '<p><a href="' + esc_(link) + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">Open check-in</a></p>' +
          '<p style="color:#5b6660">The link stops working the day after the reunion.</p><p>' + esc_(senderName_(s)) + '</p></div>' });
      return { ok: true, queued: st === 'queued' };
    }
    case 'addContribution': {
      const proofUrl = saveFile_(d.proof, d.proofName || 'money in');
      append_('Other Money In', Object.assign({}, d, { fund: d.fund === 'scholarship' ? 'Scholarship fund' : 'Reunion money', proofUrl: proofUrl, created: now }));
      return { ok: true, proofUrl: proofUrl };
    }
    case 'updatePayment': update_('Payments', 'id', d.id, { code: d.code, method: d.method, household: householdOf_(d.code) }); setBalance_(d.id, d.code, d.total); return { ok: true };
    case 'updateExpense': update_('Expenses', 'id', d.id, { category: d.category }); return { ok: true };
    case 'updateReg':     update_('Registrations', 'code', d.code, { email: d.email }); return { ok: true };
    case 'delete': {
      if (['Payments', 'Expenses', 'Other Money In'].indexOf(d.sheet) < 0) throw new Error('That can’t be removed.');
      delete_(d.sheet, 'id', d.id); return { ok: true };
    }
    case 'saveSettings': saveSettings_(d); return { ok: true };
    case 'clearTest': {
      testSS_().getSheets().forEach(sh => { if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
      return { ok: true };
    }
    case 'sendReceipt': {
      const r = d.reg || {};
      if (!/^\S+@\S+\.\S+$/.test(d.to || '')) throw new Error('That email address doesn’t look right.');
      const m = regEmail_(r, d.code, d.total, d.lines || [], settings_(), d.paid);
      const st = mail_({ to: d.to, subject: subj_(receiptSubject_(d.total, d.paid)), body: m.text, htmlBody: m.html });
      return { ok: true, queued: st === 'queued' };
    }
    case 'sendReminders': {
      const s = settings_(); let sent = 0, queued = 0;
      (d.items || []).forEach(it => {
        if (!/^\S+@\S+\.\S+$/.test(it.to || '')) return;
        const m = regEmail_(it.reg || {}, it.code, it.total, it.lines || [], s, it.paid, true);
        const st = mail_({ to: it.to, subject: subj_('Friendly reminder: balance due, ' + REUNION + ' (' + it.code + ')'), body: m.text, htmlBody: m.html });
        st === 'queued' ? queued++ : sent++;
      });
      return { ok: true, sent: sent, queued: queued };
    }
    case 'sendEmail':
      if (!/^\S+@\S+\.\S+$/.test(d.to || '')) throw new Error('That email address doesn’t look right.');
      const st2 = mail_({ to: d.to, subject: subj_(d.subject || REUNION), body: d.body || '' });
      return { ok: true, queued: st2 === 'queued' };
    default:
      throw new Error('Unknown action.');
  }
}

/* ---------- registration (public) ---------- */
// Same rules as the page: opens on the open date, closes the day after the last reunion day
// (or the host's close date), and stops at "Seats available" if the host turned that on.
function regState_(s) {
  const t = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!s.regOpen) return 'soon';          // only the host's switch opens registration
  let close = '';                          // the host closes it with the switch; after the reunion it closes by itself
  if (s.lastDay) { const d = new Date(s.lastDay + 'T12:00:00'); d.setDate(d.getDate() + 1); close = Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
  if (close && t >= close) return 'ended';
  // Venue capacity was taken out (Nov 14), so registration never stops for being full.
  //   if (s.capStop && Number(s.capacity) > 0) { const n = read_('Registrations').reduce((a, r) => a + (r.people || []).length, 0); if (n >= Number(s.capacity)) return 'full'; }
  return 'open';
}
function register_(d, admin) {
  const s = settings_();
  if (!admin && !TEST_) { const st = regState_(s); if (st !== 'open') return { ok: false, error: st === 'ended' ? 'Registration has ended.' : st === 'full' ? 'Registration is full.' : 'Registration isn’t open yet. Check back soon.' }; }
  const r = d.reg || {};
  if (!r.first || !r.last || !/^\S+@\S+\.\S+$/.test(r.email || '') || !(r.people || []).length)
    return { ok: false, error: 'Some required information is missing.' };
  const regs = read_('Registrations');
  const n = regs.reduce((m, x) => Math.max(m, parseInt(String(x.code).replace(/\D/g, ''), 10) || 0), 0) + 1;
  const code = 'PSR-' + ('00' + n).slice(-3);
  const total = Math.round((Number(d.total) || 0) * 100) / 100;
  append_('Registrations', Object.assign({}, r, { code: code, total: total, created: stamp_() }));
  try { upsertContacts_([{ email: r.email, first: r.first, last: r.last, phone: r.phone, city: r.city, state: r.state, lastSeen: stamp_() }], 'Registration'); } catch (e) {}
  log_('ok', 'New registration from ' + r.first + ' ' + r.last + ' (' + code + '): ' + r.people.length +
       (r.people.length === 1 ? ' person, ' : ' people, ') + money_(total) + '.');
  let emailed = true;
  try { const m = regEmail_(r, code, total, d.lines || [], s, 0);
        if (mail_({ to: r.email, subject: subj_('You’re registered! ' + REUNION + ' (' + code + ')'),
                    body: m.text, htmlBody: m.html }) === 'queued')
          log_('info', 'Daily email limit reached. ' + r.email + '’s confirmation will go out automatically when it resets.'); }
  catch (err) { emailed = false; log_('bad', 'Could not email ' + r.email + ': ' + err.message); }
  return { ok: true, code: code, emailed: emailed };
}

function esc_(x) { return String(x == null ? '' : x).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

// One layout for every family email: confirmation number, who's coming, order, payment, hotel.
// The reunion website. "Pay online" in emails opens its payment page (steps beside the Zeffy form).
const SITE_URL = 'https://psrreunion.github.io/';
function payLink_(code, amt) { return SITE_URL + '#pay=' + encodeURIComponent(code) + '&amt=' + (Number(amt) || 0).toFixed(2); }
function regEmail_(r, code, total, lines, s, paid, reminder) {
  total = Number(total) || 0; paid = Number(paid) || 0;
  const bal = Math.max(0, Math.round((total - paid) * 100) / 100);
  const full = total > 0 && bal < 0.005, zeffy = /^https:\/\/(www\.)?zeffy\.com\//i.test(s.zeffyUrl || '');
  const ed = s.edition ? s.edition + ' ' : '';
  const T = [], H = [];
  const G = '#1c5a34', INK = '#1f2a24', MUTED = '#5b6660', LINE = '#e3e6e1';
  const sec = t => '<h3 style="margin:26px 0 10px;padding:0 0 6px;font-size:17px;color:' + G + ';text-transform:uppercase;letter-spacing:.06em;border-bottom:3px solid #e0b22a">' + t + '</h3>';
  const opener = reminder
    ? 'This is a friendly reminder about your balance for the ' + ed + fam_(s) + '. Thank you for registering!'
    : 'Thank you for registering for the ' + ed + fam_(s) + '!';
  const ticket = !reminder && s.emailStyle !== 'regular';
  T.push('Hi ' + r.first + ',', '', opener, '', 'CONFIRMATION #: ' + code, '');
  H.push('<p style="margin:0 0 12px">Hi ' + esc_(r.first) + ',</p>',
    '<p style="margin:0 0 16px">' + esc_(opener) + '</p>');
  if (!ticket) H.push('<div style="background:#f3f7f2;border:1px solid #cfe0d2;border-radius:10px;padding:14px 16px;text-align:center">',
    '<div style="font-size:12px;color:' + MUTED + ';text-transform:uppercase;letter-spacing:.08em">Confirmation number</div>',
    '<div style="font-size:26px;font-weight:bold;color:' + G + ';letter-spacing:.04em;margin-top:2px">' + esc_(code) + '</div></div>');
  // who's coming
  T.push("WHO'S COMING"); (r.people || []).forEach(p => T.push('  • ' + p.name)); T.push('');
  if (!ticket) H.push(sec('Who’s coming'), '<ul style="margin:0;padding-left:20px">' + (r.people || []).map(p => '<li>' + esc_(p.name) + '</li>').join('') + '</ul>');
  // order
  T.push('YOUR ORDER');
  lines.forEach(l => T.push('  ' + l.label + (l.don ? '' : ' × ' + l.q) + ': ' + (l.a ? money_(l.a) : 'Free')));
  T.push('  Total: ' + money_(total));
  if (paid > 0) T.push('  Paid so far: ' + money_(paid));
  T.push('');
  const row = (a, b, bold) => '<tr><td style="padding:6px 0;border-bottom:1px solid ' + LINE + (bold ? ';font-weight:bold' : '') + '">' + a + '</td><td style="padding:6px 0;border-bottom:1px solid ' + LINE + ';text-align:right' + (bold ? ';font-weight:bold' : '') + '">' + b + '</td></tr>';
  H.push(sec('Your order'), '<table style="width:100%;border-collapse:collapse;font-size:15px">' +
    lines.map(l => row(esc_(l.label) + (l.don ? '' : ' <span style="color:' + MUTED + '">× ' + l.q + '</span>'), l.a ? money_(l.a) : 'Free')).join('') +
    row('Total', money_(total), true) + (paid > 0 ? row('Paid so far', money_(paid)) : '') + '</table>');
  // payment
  if (full) {
    T.push('PAYMENT', '  Paid in full. Thank you! There is nothing more to pay.', '');
    H.push(sec('Payment'), '<div style="background:#e6f4ea;border:1px solid #9fd0ad;border-radius:10px;padding:12px 16px;font-weight:bold;color:#1d6b37">✓ Paid in full. Thank you! There is nothing more to pay.</div>');
  } else if (total > 0) {
    const amt = money_(bal);
    T.push('PAYMENT', '  ' + (paid > 0 ? 'Balance due: ' : 'Amount due: ') + amt);
    H.push(sec('Payment'), '<p style="margin:0 0 10px;font-size:16px">' + (paid > 0 ? 'Balance due: ' : 'Amount due: ') + '<b>' + amt + '</b></p>');
    if (zeffy) {
      const payUrl = payLink_(code, bal);
      T.push('  Pay online: ' + payUrl, '  The payment page shows your amount and walks you through each step.', '');
      H.push('<p style="margin:0 0 10px"><a href="' + esc_(payUrl) + '" style="display:inline-block;background:' + G + ';color:#ffffff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">Pay online</a></p>',
        '<p style="margin:0;color:' + MUTED + '">The payment page shows your amount and walks you through each step.</p>');
    } else {
      T.push('  Online payment opens soon. We’ll email you as soon as you can pay.', '');
      H.push('<p style="margin:0;color:' + MUTED + '">Online payment opens soon. We’ll email you as soon as you can pay.</p>');
    }
  }
  // hotel
  if (s.hotelName) {
    const book = /^https?:\/\//i.test(s.bookUrl || '');
    T.push('HOTEL', '  ' + s.hotelName);
    if (s.hotelStreet) T.push('  ' + s.hotelStreet);
    if (s.hotelCity) T.push('  ' + s.hotelCity);
    if (s.roomRate) T.push('  ' + rateText_(s.roomRate));
    const maps = 'https://www.google.com/maps/dir/?api=1&destination=' + encodeURIComponent([s.hotelName, s.hotelStreet, s.hotelCity].filter(Boolean).join(', '));
    T.push('  Directions: ' + maps);
    T.push('  Haven’t booked your room yet? ' + (book ? 'Book your room: ' + s.bookUrl : 'Hotel booking link coming soon.'), '  Rooms are paid to the hotel, not the reunion.', '');
    const addr = [s.hotelStreet, s.hotelCity].filter(Boolean).map(esc_).join('<br>');
    H.push(sec('Hotel'), '<p style="margin:0;line-height:1.5"><b>' + esc_(s.hotelName) + '</b>' +
      (addr ? '<br><a href="' + esc_(maps) + '" style="color:' + G + ';text-decoration:underline"><img src="https://psrreunion.github.io/pin.png" width="16" height="16" alt="" style="vertical-align:-2px;margin-right:4px;border:0">' + addr + '</a>' : '') +
      (s.roomRate ? '<br>' + esc_(rateText_(s.roomRate)) : '') + '</p>',
      '<p style="margin:14px 0 6px">Haven’t booked your room yet?</p>',
      book ? '<p style="margin:0"><a href="' + esc_(s.bookUrl) + '" style="display:inline-block;background:#e3b318;color:#1f2a24;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">Book your room</a></p>'
           : '<p style="margin:0;color:' + MUTED + '">Hotel booking link coming soon.</p>',
      '<p style="margin:8px 0 0;color:' + MUTED + '">Rooms are paid to the hotel, not the reunion.</p>');
  }
  if (ticket) H.push(ticketHtml_(r, code, s));
  const from = senderName_(s);
  T.push('Questions? Reply to this email.', '', from);
  H.push('<p style="margin:24px 0 0">Questions? Just reply to this email.</p>', '<p style="margin:6px 0 0;color:' + MUTED + '">' + esc_(from) + '</p>');
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:' + INK + ';max-width:560px;margin:0 auto;line-height:1.45">' +
    (TEST_ ? '<div style="background:#fde8e8;color:#7a1414;padding:8px 12px;border-radius:8px;margin-bottom:12px;font-size:13px"><b>TEST</b>: practice email, not a real registration.</div>' : '') +
    H.join('\n') + '</div>';
  return { text: T.join('\n'), html: html };
}
// "129" becomes "Room rate: $129 a night". Words typed by the host are kept as they are.
function rateText_(v) {
  v = String(v || '').trim().replace(/^\$\s*/, '');
  if (!v) return '';
  if (/^\d+(\.\d{1,2})?$/.test(v)) return 'Room rate: $' + v + ' a night';
  return 'Room rate: ' + (/^\d/.test(v) ? '$' : '') + v;
}
// Where family replies go: the host's own email from Setup (no forwarding needed).
// Replies come back to the reunion account (no Reply-To), so checkReplies() can log each one and send it on to the host.
function replyTo_() { return ''; }
// Who emails come from. Set in Setup ("Emails come from").
function senderName_(s) {
  try { s = s || settings_(); } catch (e) { s = {}; }
  const h = String(s.host || '').trim();
  return (h ? (/^hosted by\b/i.test(h) ? h : 'Hosted by ' + h) : '') || 'PSR Host Family Committee';   // always from the hosting family in step 1
}
// Ticket-style box for the confirmation email. Everything on it comes from Setup,
// so a new host's name, theme, dates and hotel show up without editing the email.
function ticketHtml_(r, code, s) {
  const G = '#1c5a34', GOLD = '#e0b22a', MUTED = '#5b6660';
  const lab = t => '<div style="font-size:11px;color:' + MUTED + ';letter-spacing:.1em;text-transform:uppercase">' + t + '</div>';
  const people = (r.people || []).map(p => esc_(p.name)).filter(Boolean);
  const when = String(s.when || '').split('·').map(x => x.trim()).filter(Boolean);
  const qr = 'https://quickchart.io/qr?size=240&margin=1&text=' + encodeURIComponent(SITE_URL + '#ci=' + code);
  const title = [s.edition, fam_(s)].filter(Boolean).join(' ');
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border:2px solid ' + G + ';border-radius:14px;overflow:hidden;margin:22px 0 4px;font-family:Arial,Helvetica,sans-serif">' +
    '<tr><td style="background:' + G + ';color:#ffffff;padding:14px 18px">' +
      '<div style="font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:' + GOLD + ';font-weight:bold">' + esc_(title) + '</div>' +
      (s.theme ? '<div style="font-family:Georgia,serif;font-size:22px;line-height:1.2;margin-top:4px">' + esc_(s.theme) + '</div>' : '') +
    '</td></tr>' +
    '<tr><td style="padding:16px 18px 6px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
      '<td style="vertical-align:top;font-size:14px;color:#1f2a24">' +
        lab('Admit') + '<div style="font-size:20px;font-weight:bold;color:' + G + ';margin:2px 0 10px">The ' + esc_(r.last) + ' Family</div>' +
        lab('Confirmation #') + '<div style="font-size:22px;font-weight:bold;letter-spacing:.05em;margin:2px 0 10px">' + esc_(code) + '</div>' +
        lab('Who’s coming (' + people.length + ')') + '<div style="margin-top:3px;line-height:1.5">' + people.join('<br>') + '</div>' +
      '</td>' +
      '<td style="vertical-align:top;text-align:center;width:132px;padding-left:10px">' +
        '<img src="' + qr + '" width="124" height="124" alt="Check-in code ' + esc_(code) + '" style="display:block;border:1px solid #e3e6e1;border-radius:8px">' +
        '<div style="font-size:11px;color:' + MUTED + ';margin-top:5px">Check-in code</div>' +
      '</td></tr></table></td></tr>' +
    '<tr><td style="padding:0 18px"><div style="border-top:2px dashed #cfd6cf;margin:10px 0"></div></td></tr>' +
    '<tr><td style="padding:2px 18px 14px;font-size:14px;color:#1f2a24"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
      (when.length ? '<td style="vertical-align:top;padding-right:8px">' + lab('When') + '<b>' + esc_(when[0]) + '</b>' + (when[1] ? '<br><span style="color:' + MUTED + '">' + esc_(when.slice(1).join(' · ')) + '</span>' : '') + '</td>' : '') +
      (s.hotelName ? '<td style="vertical-align:top">' + lab('Where') + '<b>' + esc_(s.hotelName) + '</b>' + (s.hotelCity ? '<br><span style="color:' + MUTED + '">' + esc_(s.hotelCity) + '</span>' : '') + '</td>' : '') +
    '</tr></table>' +
    '<div style="margin-top:12px;background:#fbf3d9;border-radius:8px;padding:9px 12px;font-size:13px;color:#5a4a12">Show this at check-in. <b>No ticket? No problem.</b> Just tell us your name.</div>' +
    '</td></tr></table>';
}
function receiptSubject_(total, paid) {
  total = Number(total) || 0; paid = Number(paid) || 0;
  if (total > 0 && paid >= total - 0.005) return 'Paid in full, thank you! ' + REUNION;
  if (paid > 0) return 'Payment received, balance remaining: ' + REUNION;
  return 'Your ' + REUNION + ' registration: payment due';
}

function publicData_() {
  const regs = read_('Registrations');
  const pub = Object.assign({}, settings_()); delete pub.checkinCode; delete pub.alertEmails; delete pub.boardEmails;
  return { ok: true, settings: pub, handoff: handoffInfo_(), households: regs.length,
           people: regs.reduce((s, r) => s + ((r.people || []).length), 0) };
}

/* ---------- sheet helpers ---------- */
let REAL_SS_ = null;   // opened once per request (opening a spreadsheet is slow)
function realSS_() {
  if (REAL_SS_) return REAL_SS_;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  REAL_SS_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  return REAL_SS_;
}
let PRAC_SS_ = null, TEST_SS_ = null;
/* ---------- Tryouts: let someone else try being a host ----------
   From the tech PIN's Practice box: type a name and email, and they get an email with their own PIN and links.
   Each tryout has its own sheet, so testers never see each other's work or yours. After TRY_DAYS it ends:
   the PIN and links stop working and the sheet goes to the trash. */
const TRY_DAYS = 10;
let TRY_ = null;
function tryouts_() { try { return JSON.parse(PROPS_().getProperty('TRYOUTS') || '[]'); } catch (e) { return []; } }
function saveTryouts_(l) { PROPS_().setProperty('TRYOUTS', JSON.stringify(l)); }
function tryHash_(p) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'psr-try:' + p).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function liveTry_(t) { return !!t && String(t.expires) >= today_(); }
function tryByPin_(pin) { if (!pin) return null; const h = tryHash_(pin); return tryouts_().filter(t => t.pin === h && liveTry_(t))[0] || null; }
function tryById_(id) { id = String(id || ''); return id ? tryouts_().filter(t => t.id === id && liveTry_(t))[0] || null : null; }
function listTryouts_() { return tryouts_().map(t => ({ id: t.id, name: t.name, email: t.email, created: t.created, expires: t.expires, live: liveTry_(t) })); }
function inviteTester_(d) {
  const name = String(d.name || '').trim().replace(/\s+/g, ' '), email = String(d.email || '').trim();
  if (name.length < 2) return { ok: false, error: 'Please type their name.' };
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, error: 'Please type their email address.' };
  const list = tryouts_(); let pin;
  do { pin = String(Math.floor(10000000 + Math.random() * 90000000)); }
  while ([ADMIN_PIN, COMMITTEE_PIN, TECH_PIN].map(String).indexOf(pin) >= 0 || hostPinOk_(pin) || list.some(t => t.pin === tryHash_(pin)));
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const ss = SpreadsheetApp.create('PSR Reunion Hub – Tryout for ' + name);
  const ends = Utilities.formatDate(new Date(Date.now() + TRY_DAYS * 864e5), TZ, 'yyyy-MM-dd');
  const endsNice = Utilities.formatDate(new Date(ends + 'T12:00:00'), TZ, 'EEEE, MMMM d');
  list.push({ id: id, name: name, email: email, pin: tryHash_(pin), sheetId: ss.getId(), created: today_(), expires: ends });
  saveTryouts_(list);
  const famLink = SITE_URL + '#try=' + id, first = name.split(' ')[0];
  const body = 'Hello ' + first + ',\n\nYou’re invited to try the PSR Family Reunion registration page. Everything you do is practice only. Nothing is real and nothing is charged.\n\n' +
    'TRY BEING THE HOST\n1. Open ' + SITE_URL + '\n2. At the bottom of the page, click "Hosting committee sign-in."\n3. Type your PIN: ' + pin + '\n4. Click "Get started" and follow the steps.\n\n' +
    'TRY SIGNING UP AS A FAMILY\nOpen this link: ' + famLink + '\n(It works after you’ve set up the reunion and turned registration on in the last step.)\n\n' +
    'Your tryout ends on ' + endsNice + '. After that, your PIN and link stop working.\n\nThank you for trying it! Please let me know what was easy and what was confusing.\n';
  const html = '<div style="font:16px/1.6 Arial,sans-serif;color:#1f2a24;max-width:560px"><p>Hello ' + esc_(first) + ',</p><p>You’re invited to try the PSR Family Reunion registration page. Everything you do is <b>practice only</b>. Nothing is real and nothing is charged.</p>' +
    '<h3 style="margin:22px 0 6px;color:#1c5a34">Try being the host</h3><ol style="margin:0;padding-left:22px"><li>Open <a href="' + SITE_URL + '">' + SITE_URL + '</a></li><li>At the bottom of the page, click <b>Hosting committee sign-in</b>.</li><li>Type your PIN: <b style="font-size:20px;letter-spacing:2px">' + pin + '</b></li><li>Click <b>Get started</b> and follow the steps.</li></ol>' +
    '<h3 style="margin:22px 0 6px;color:#1c5a34">Try signing up as a family</h3><p style="margin:0 0 8px"><a href="' + esc_(famLink) + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">Open the family sign-up</a></p><p style="margin:0;color:#5b6660">This works after you’ve set up the reunion and turned registration on in the last step.</p>' +
    '<p style="margin-top:22px">Your tryout ends on <b>' + esc_(endsNice) + '</b>. After that, your PIN and link stop working.</p><p>Thank you for trying it! Please let me know what was easy and what was confusing.</p></div>';
  const st = mail_({ to: email, subject: 'You’re invited to try the PSR reunion registration page', body: body, htmlBody: html });
  log_('info', 'Sent a tryout invite to ' + name + ' (' + email + '). It ends ' + ends + '.');
  return { ok: true, queued: st === 'queued', tryouts: listTryouts_() };
}
function endTryout_(id) {
  const list = tryouts_(), t = list.filter(x => x.id === id)[0];
  if (!t) return { ok: true, tryouts: listTryouts_() };
  try { DriveApp.getFileById(t.sheetId).setTrashed(true); } catch (e) {}
  saveTryouts_(list.filter(x => x.id !== id));
  return { ok: true, tryouts: listTryouts_() };
}
function cleanTryouts_() { tryouts_().filter(t => !liveTry_(t)).forEach(t => endTryout_(t.id)); }
function testSS_() {
  if (TEST_SS_) return TEST_SS_;
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('TEST_SHEET_ID');
  if (id) { try { TEST_SS_ = SpreadsheetApp.openById(id); return TEST_SS_; } catch (e) {} }
  TEST_SS_ = SpreadsheetApp.create(TEST_NAME);
  P.setProperty('TEST_SHEET_ID', TEST_SS_.getId());
  return TEST_SS_;
}
function practiceSS_() {
  if (TRY_) { if (!TRY_._ss) TRY_._ss = SpreadsheetApp.openById(TRY_.sheetId); return TRY_._ss; }   // a tester's own tryout sheet
  if (PRAC_SS_) return PRAC_SS_;
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('PRACTICE_SHEET_ID');
  if (id) { try { PRAC_SS_ = SpreadsheetApp.openById(id); return PRAC_SS_; } catch (e) {} }
  PRAC_SS_ = SpreadsheetApp.create(PRAC_NAME);
  P.setProperty('PRACTICE_SHEET_ID', PRAC_SS_.getId());
  return PRAC_SS_;
}
function ss_() { return PRAC_ ? practiceSS_() : TEST_ ? testSS_() : realSS_(); }
function subj_(x) { return TEST_ ? '[TEST] ' + x : x; }

/* ---------- email with catch-up ----------
   Gmail lets this account send about 100 emails a day. If the limit is reached,
   the email is saved on the "Email Queue" tab and sent automatically once the limit resets
   (sendQueuedEmails runs every 15 minutes; setup() turns that on). */
const QUEUE_TAB = 'Email Queue';
function mail_(m) {
  const msg = { to: m.to, subject: m.subject, body: m.body || '', name: senderName_() }; const rt = replyTo_(); if (rt) msg.replyTo = rt;
  if (m.htmlBody) msg.htmlBody = m.htmlBody;
  if (MailApp.getRemainingDailyQuota() > 0) {
    try { MailApp.sendEmail(msg); emailLog_(msg); return 'sent'; }
    catch (err) { if (!/limit|quota|too many/i.test(err.message)) throw err; }
  }
  queueSheet_().appendRow([stamp_(), msg.to, msg.subject, msg.body, msg.htmlBody || '']);
  return 'queued';
}
// Every email the system sends is listed on the Emails tab (Sent), next to the replies that come back (Received).
function emailLog_(msg) {
  try {
    const now = new Date(), to = String(msg.to || ''), first = to.split(',')[0].trim().toLowerCase();
    const r = read_('Registrations').find(x => String(x.email || '').trim().toLowerCase() === first);
    append_('Emails', { date: Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(now, TZ, 'h:mm a'), direction: 'Sent',
      who: r ? r.first + ' ' + r.last : '', email: to, code: r ? r.code : '', subject: String(msg.subject || '').slice(0, 200),
      note: String(msg.body || '').replace(/\s+/g, ' ').trim().slice(0, 200) });
  } catch (e) {}
}
function queueSheet_() {
  const ss = realSS_();
  let sh = ss.getSheetByName(QUEUE_TAB);
  if (!sh) {
    sh = ss.insertSheet(QUEUE_TAB);
    sh.getRange(1, 1, 1, 5).setValues([['Waiting Since', 'To', 'Subject', 'Text', 'HTML']]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.setTabColor('#e69f00');
  }
  return sh;
}
function sendQueuedEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = ss_().getSheetByName(QUEUE_TAB);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    let sent = 0;
    for (let i = 0; i < rows.length; i++) {
      if (MailApp.getRemainingDailyQuota() < 1) break;
      const [, to, subject, body, html] = rows[i];
      const msg = { to: to, subject: subject, body: body, name: senderName_() }; const rt = replyTo_(); if (rt) msg.replyTo = rt;
      if (html) msg.htmlBody = html;
      try { MailApp.sendEmail(msg); sent++; emailLog_(msg); } catch (err) { break; }
    }
    if (sent) sh.deleteRows(2, sent);
  } finally { lock.releaseLock(); }
}
function sheet_(name) {
  if (TEST_) {                       // Test mode: the test spreadsheet (Setup itself is shared with the real page)
    if (name === 'Settings') { const r = realSS_().getSheetByName('Settings'); if (r) return r; }
    const ts = testSS_();
    let t = ts.getSheetByName(name);
    if (!t) { t = ts.insertSheet(name); formatTab_(t, name, '#f8e1e1'); t.setTabColor('#c0392b');
      const blank = ts.getSheetByName('Sheet1'); if (blank && blank.getLastRow() === 0) ts.deleteSheet(blank); }
    return t;
  }
  if (PRAC_) {
    const ps = practiceSS_();
    let t = ps.getSheetByName(name);
    if (!t) {
      t = ps.insertSheet(name); formatTab_(t, name, '#e3efe6');
      const blank = ps.getSheetByName('Sheet1'); if (blank && blank.getLastRow() === 0) ps.deleteSheet(blank);
    }
    return t;
  }
  let sh = ss_().getSheetByName(name);
  if (!sh && name === 'Other Money In' && ss_().getSheetByName('Contributions')) { sh = ss_().getSheetByName('Contributions').setName(name); addMissingCols_(sh, name); }
  if (!sh && (name === 'Other Money In' || name === 'Contacts')) { sh = ss_().insertSheet(name); formatTab_(sh, name, '#e3efe6'); }
  if (!sh) throw new Error('The "' + name + '" tab is missing. Run setup again.');
  return sh;
}
// While the page is loading ('all' and 'public'), each tab is read once and reused. Any write clears this.
let READS_ = null;
function read_(name) {
  if (!READS_) return readTab_(name);
  const k = (PRAC_ ? 'p:' : TEST_ && name !== 'Settings' ? 't:' : 'r:') + name;
  if (!READS_[k]) READS_[k] = readTab_(name);
  return JSON.parse(JSON.stringify(READS_[k]));
}
function readTab_(name) {
  const rng = sheet_(name).getDataRange(), v = rng.getValues();
  if (v.length < 2) return [];
  const h = v[0].map(x => keyOf_(name, x)), tz = TZ;
  const hasLinks = h.some(k => URL_COLS.indexOf(k) >= 0), rich = hasLinks ? rng.getRichTextValues() : null;
  return v.slice(1).filter(r => r.some(c => c !== '' && c !== null)).map(r => {
    const o = {};
    h.forEach((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, tz, k === 'time' ? 'h:mm a' : 'yyyy-MM-dd');
      if (k === 'date') x = toISO_(x);
      if (rich && URL_COLS.indexOf(k) >= 0 && x === 'View file') { const rt = rich[v.indexOf(r)] && rich[v.indexOf(r)][i]; x = (rt && rt.getLinkUrl()) || ''; }
      if (JSON_COLS.indexOf(k) >= 0) x = parseCol_(k, x);
      else if (NUM_COLS.indexOf(k) >= 0) x = x === '' ? 0 : Number(String(x).replace(/[$,]/g, ''));
      o[k] = x;
    });
    return o;
  });
}
function cell_(k, x) {
  if (k === 'date') x = toMDY_(x);
  if (MONEY_COLS.indexOf(k) >= 0 && x !== '' && x !== null && x !== undefined && !isNaN(Number(x))) return Number(x);
  if (JSON_COLS.indexOf(k) >= 0) x = typeof x === 'string' ? x : plainCol_(k, x);
  if (x === undefined || x === null) x = '';
  x = String(x);
  if (/^[=+@]/.test(x) || /^-[^\d]/.test(x)) x = "'" + x;   // never let typed text run as a formula
  return x;
}
/* People, Shirts and Add-ons show as plain lines in the Sheet:
   "Jane Doe (Adult)" one per line, and "M × 2" for sizes and extras. */
function plainCol_(k, x) {
  if (k === 'people') return (x || []).map(p => String(p.name || '').trim() + ' (' + cap_(p.group || 'adult') + ')' + (p.food ? ' · ' + p.food : '')).join('\n');
  const o = x || {}, ks = Object.keys(o).filter(q => o[q] !== 0 && o[q] !== '' && o[q] !== null);
  if (ks.some(q => isNaN(Number(o[q])))) return JSON.stringify(o);
  return ks.map(q => q + ' × ' + Number(o[q])).join('\n');
}
function parseCol_(k, x) {
  x = String(x == null ? '' : x).trim();
  const empty = k === 'people' ? [] : {};
  if (!x) return empty;
  if (/^[\[{]/.test(x)) { try { return JSON.parse(x); } catch (e) { return empty; } }
  const lines = x.split(/\n/).map(t => t.trim()).filter(Boolean);
  if (k === 'people') return lines.map(t => { let food = ''; const f = t.split(' · '); if (f.length > 1) { food = f.slice(1).join(' · ').trim(); t = f[0]; }
    const m = t.match(/^(.*?)\s*\(([^()]+)\)$/); const o = m ? { name: m[1], group: m[2].trim().toLowerCase() } : { name: t, group: 'adult' }; if (food) o.food = food; return o; });
  const o = {};
  lines.forEach(t => { const m = t.match(/^(.+?)\s*[×:]\s*(\d+)$/); if (m) o[m[1].trim()] = Number(m[2]); });
  return o;
}
function cap_(t) { t = String(t); return t.charAt(0).toUpperCase() + t.slice(1); }
// Rewrites any old code-looking People/Shirts/Add-ons cells as plain lines.
function plainTab_(sh, name) {
  const cols = TABS[name], n = sh.getLastRow();
  if (n < 2) return;
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => keyOf_(name, x));
  JSON_COLS.forEach(k => {
    const c = h.indexOf(k); if (c < 0) return;
    const rg = sh.getRange(2, c + 1, n - 1, 1), v = rg.getValues();
    let changed = false;
    const out = v.map(r => { const t = String(r[0] || '').trim(); if (/^[\[{]/.test(t)) { changed = true; return [plainCol_(k, parseCol_(k, t))]; } return [r[0]]; });
    if (changed) rg.setValues(out);
  });
}
function append_(name, obj) {
  READS_ = null;
  if (name === 'Payments' && obj.code && !obj.household) obj = Object.assign({}, obj, { household: householdOf_(obj.code) });
  const sh = sheet_(name), cols = TABS[name];
  sh.appendRow(cols.map(k => cell_(k, obj[k])));
  cols.forEach((k, i) => { if (URL_COLS.indexOf(k) >= 0 && obj[k]) linkCell_(sh.getRange(sh.getLastRow(), i + 1), obj[k]); });
}
/* ---------- receipt / proof links show as "View file" ---------- */
const URL_COLS = ['proofUrl', 'receiptUrl'];
function linkCell_(rg, url) {
  rg.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('View file').setLinkUrl(String(url)).build());
}
function linkifyTab_(sh, name) {
  const cols = TABS[name], n = sh.getLastRow();
  if (n < 2) return;
  cols.forEach((k, i) => {
    if (URL_COLS.indexOf(k) < 0) return;
    const rg = sh.getRange(2, i + 1, n - 1, 1), v = rg.getValues();
    v.forEach((r, j) => { if (/^https?:\/\//.test(String(r[0]))) linkCell_(sh.getRange(j + 2, i + 1), r[0]); });
  });
}
function rowOf_(name, key, val) {
  const sh = sheet_(name), v = sh.getDataRange().getValues(), col = v[0].map(x => keyOf_(name, x)).indexOf(key);
  for (let i = 1; i < v.length; i++) if (String(v[i][col]) === String(val)) return i + 1;
  throw new Error('Couldn’t find that entry. Refresh and try again.');
}
function update_(name, key, val, fields) {
  READS_ = null;
  const sh = sheet_(name), row = rowOf_(name, key, val), h = TABS[name];
  Object.keys(fields).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(row, c + 1).setValue(cell_(k, fields[k])); });
}
function delete_(name, key, val) { READS_ = null; sheet_(name).deleteRow(rowOf_(name, key, val)); }

function settings_() {
  const rows = read_('Settings'), r = rows.filter(x => x.key === 'settings')[0];
  let s = {}; try { s = r ? JSON.parse(r.value) : {}; } catch (e) { s = {}; }
  if (PRAC_) { const keep = PRAC_; PRAC_ = false; try { s.zeffyUrl = settings_().zeffyUrl || ''; } finally { PRAC_ = keep; } }
  return s;
}
// Helper check-in code: made automatically, a new one each calendar year (reunion or not).
// It works through the last reunion day in Setup and stops working the day after.
function thisYear_() { return Number(Utilities.formatDate(new Date(), TZ, 'yyyy')); }
function newCheckinCode_() {
  const abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; let c = '';
  for (let i = 0; i < 6; i++) c += abc.charAt(Math.floor(Math.random() * abc.length));
  return c;
}
function ensureCheckinCode_() {   // called when the hosting committee loads the page
  const s = settings_();
  if (s.checkinCode && Number(s.checkinCodeYear) === thisYear_()) return s;
  if (s.checkinCode && !s.checkinCodeYear) s.checkinCodeYear = thisYear_();   // keep a code the host already made this year
  else { s.checkinCode = newCheckinCode_(); s.checkinCodeYear = thisYear_(); }
  writeSettings_(s); log_('info', 'Made this year’s helper check-in code.');
  return s;
}
function currentCheckinCode_() {
  const s = settings_(), code = String(s.checkinCode || '').trim();
  if (!code || Number(s.checkinCodeYear) !== thisYear_()) return '';
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (s.lastDay && today > String(s.lastDay)) return '';   // reunion is over
  return code;
}
function saveSettings_(obj) {
  obj = obj || {};
  const old = settings_();                       // the page can't change the helper code; keep the server's
  obj.checkinCode = old.checkinCode || ''; obj.checkinCodeYear = old.checkinCodeYear || '';
  delete obj.testRetired; delete obj.testDone; delete obj.showTest;
  if (!obj.zeffyUrl && old.zeffyUrl && !obj._setZeffy) obj.zeffyUrl = old.zeffyUrl;   // only the tech PIN sees the Zeffy box
  delete obj._setZeffy;  delete obj.setupPractice;
  // A last day typed by mistake far in the past would lock the page right away, so it isn't accepted.
  if (handoffOpen_(obj) && !handoffOpen_(old)) throw new Error('The last day of the reunion can’t be more than ' + HANDOFF_DAYS + ' days ago. Please check the dates in step 1 of Setup.');
  delete obj.boardEmails;          // the PSR Reunion Board email is kept separately (see boardEmails_)
  writeSettings_(obj);
}
/* ---------- PSR Reunion Board email ----------
   New-host notices go here. Hosts can't change it. Only the board (signed in with the board's PIN)
   or the tech PIN can, and the old address is told when it changes. */
const BOARD_DEFAULT = 'portersanfordreid@gmail.com';
function boardEmails_() { return PROPS_().getProperty('BOARD_EMAILS') || BOARD_DEFAULT; }
// Where board email actually goes: the board when live, otherwise you.
function boardSendTo_() { return BOARD_LIVE ? emailList_(boardEmails_()) : [TECH_EMAIL]; }
function boardNote_() { return BOARD_LIVE ? '' : '[Would have gone to the PSR Reunion Board: ' + boardEmails_() + '] '; }
function saveBoard_(d) {
  const after = emailList_(d.emails);
  if (!after.length) return { ok: false, error: 'Please type the board’s email address.' };
  const before = emailList_(boardEmails_());
  PROPS_().setProperty('BOARD_EMAILS', after.join(', '));
  const gone = before.filter(e => after.indexOf(e) < 0);
  if (gone.length) {
    mail_({ to: BOARD_LIVE ? gone.join(',') : TECH_EMAIL, subject: 'The PSR Reunion Board email on the registration page was changed',
      body: 'Hello,\n\nThe PSR Reunion Board email on the registration page was just changed.\n\n' +
            'Before: ' + before.join(', ') + '\nNow: ' + after.join(', ') + '\n\n' +
            'This email is where new-host notices go. If the board didn’t make this change, sign in with the board’s PIN and change it back.\n' });
    log_('warn', 'The PSR Reunion Board email was changed to ' + after.join(', ') + '. A notice went to ' + gone.join(', ') + '.');
  }
  return { ok: true, board: after.join(', ') };
}
function writeSettings_(obj) {
  READS_ = null;
  const sh = sheet_('Settings'), v = sh.getDataRange().getValues(), json = JSON.stringify(obj);
  for (let i = 1; i < v.length; i++) if (v[i][0] === 'settings') { sh.getRange(i + 1, 2).setValue(json); return; }
  sh.appendRow(['settings', json]);
}
// The Log's Status column uses plain words; the code keeps short keys.
const STATUS_WORD = { ok: 'Done', info: 'Note', warn: 'Check this', bad: 'Problem' };
const STATUS_KEY = { 'Done': 'ok', 'Note': 'info', 'Check this': 'warn', 'Problem': 'bad' };
function log_(kind, text) {
  const tz = TZ, now = new Date();
  append_('Log', { date: Utilities.formatDate(now, tz, 'yyyy-MM-dd'), time: Utilities.formatDate(now, tz, 'h:mm a'),
                   kind: STATUS_WORD[kind] || STATUS_WORD.info, text: text || '' });
}
function readLog_() { return read_('Log').map(l => Object.assign({}, l, { kind: STATUS_KEY[l.kind] || l.kind || 'info' })); }
function wordifyLog_(sh) {
  const n = sh.getLastRow(); if (n < 2) return;
  const rg = sh.getRange(2, 3, n - 1, 1), v = rg.getValues();
  rg.setValues(v.map(x => [STATUS_WORD[x[0]] || x[0]]));
}
// What a household still owes after its payments, saved on each payment row.
function setBalance_(id, code, total) {
  if (!code) return;
  let t = Number(total);
  if (!(t > 0)) { const r = read_('Registrations').filter(x => x.code === code)[0]; t = r ? Number(r.total) || 0 : 0; }
  update_('Payments', 'id', id, { balance: Math.max(0, Math.round((t - paidFor_(code)) * 100) / 100) });
}

/* ---------- files (screenshots, receipts) ---------- */
function folder_() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('PSR Reunion Uploads');
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function saveFile_(dataUrl, name) {
  if (!dataUrl) return '';
  const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
  if (!m) return '';
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HHmm');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], stamp + ' ' + String(name).slice(0, 60));
  const f = folder_().createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return f.getUrl();
}

/* ---------- misc ---------- */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function money_(n) { return '$' + (Number(n) || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }


/* ---------- check-in at the reunion ----------
   The check-in screen finds a household by name, then records each person who is here
   for Friday or Saturday on the Check-ins tab. Totals show at the top right of that tab. */
function ciTotals_(sh) {
  const days = [['J1', 'Friday'], ['K1', 'Saturday']];
  days.forEach(([a, d]) => sh.getRange(a).setFormula('="' + d + ': "&COUNTIF(B2:B,"' + d + '")&" checked in"'));
  sh.getRange('L1').setFormula('="Walk-ins: "&COUNTIF(F2:F,"Yes")');
  sh.getRange('M1').setFormula('="Total here: "&COUNTA(E2:E)');
  sh.getRange('J1:M1').setFontWeight('bold').setBackground('#fdf3c4');
  [10, 11, 12, 13].forEach(c => sh.setColumnWidth(c, 190));
}
function checkins_() {
  return read_('Check-ins').filter(c => c.name).map(c => ({ id: c.id, day: c.day, code: c.code, household: c.household, name: c.name, walkin: c.walkin === 'Yes', by: c.by, created: c.created }));
}
function ciStats_(list) {
  const st = { total: list.length, Friday: 0, Saturday: 0, walkins: 0 };
  list.forEach(c => { if (st[c.day] !== undefined) st[c.day]++; if (c.walkin) st.walkins++; });
  return st;
}
function ciNorm_(x) { return String(x || '').toLowerCase().replace(/[^a-z\s]/g, ' ').replace(/\s+/g, ' ').trim(); }
function ciFind_(d) {
  const q = ciNorm_(d.q), day = d.day === 'Saturday' ? 'Saturday' : 'Friday';
  const list = checkins_();
  const code = String(d.code || '').trim().toUpperCase();
  if (d.all) return ciAll_(list);           // the whole list, kept on the check-in device in case the internet drops
  if (!q && !code) return { ok: true, results: [], stats: ciStats_(list) };
  const toks = q.split(' ');
  const out = [];
  read_('Registrations').forEach(r => {
    const names = (r.people || []).map(p => p.name).filter(Boolean);
    const hay = ciNorm_([r.first, r.last].concat(names).join(' '));
    if (code ? String(r.code).toUpperCase() !== code : !toks.every(t => hay.indexOf(t) >= 0)) return;
    const exact = names.some(n => { const nn = ciNorm_(n); return toks.every(t => nn.indexOf(t) >= 0); });
    const mine = list.filter(c => c.code === r.code), done = mine.map(c => c.name), doneDay = {};
    mine.forEach(c => doneDay[c.name] = c.day);
    out.push({ code: r.code, household: r.first + ' ' + r.last, people: names, done: done, doneDay: doneDay, exact: exact,
               reg: ciReg_(r), paid: paidFor_(r.code) });
  });
  out.sort((a, b) => (b.exact - a.exact) || a.household.localeCompare(b.household));
  return { ok: true, results: out.slice(0, 40), stats: ciStats_(list) };
}
function ciAll_(list) {
  const pays = read_('Payments'), paid = {};
  pays.forEach(p => { if (p.code) paid[p.code] = (paid[p.code] || 0) + (Number(p.amount) || 0); });
  const out = read_('Registrations').map(r => {
    const mine = list.filter(c => c.code === r.code), doneDay = {};
    mine.forEach(c => doneDay[c.name] = c.day);
    return { code: r.code, household: r.first + ' ' + r.last, people: (r.people || []).map(p => p.name).filter(Boolean),
             done: mine.map(c => c.name), doneDay: doneDay, reg: ciReg_(r), paid: paid[r.code] || 0 };
  });
  return { ok: true, all: out, stats: ciStats_(list) };
}
// Undo a check-in made by mistake (removes that person's check-in for the weekend).
function ciUndo_(d) {
  const code = String(d.code || ''), name = ciNorm_(d.name);
  if (!code || !name) return { ok: false, error: 'Nothing to undo.' };
  const sh = sheet_('Check-ins'), n = sh.getLastRow();
  if (n < 2) return { ok: true, removed: 0, stats: ciStats_(checkins_()) };
  const w = sh.getLastColumn();
  const h = sh.getRange(1, 1, 1, w).getValues()[0].map(x => keyOf_('Check-ins', x));
  const ci = h.indexOf('code'), ni = h.indexOf('name');
  const v = sh.getRange(2, 1, n - 1, w).getValues();
  let removed = 0;
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][ci]) === code && ciNorm_(v[i][ni]) === name) { sh.deleteRow(i + 2); removed++; }
  if (removed) log_('info', 'Undid the check-in for ' + d.name + ' (' + code + ').');
  return { ok: true, removed: removed, stats: ciStats_(checkins_()) };
}
function ciSave_(d, admin) {
  const day = d.day === 'Saturday' ? 'Saturday' : 'Friday';
  const list = checkins_(), now = stamp_();
  const have = {}; list.forEach(c => { have[(c.code || '') + '|' + ciNorm_(c.name)] = true; });   // once for the whole weekend
  let added = 0, base = Date.now();
  const add = (name, walkin) => {
    name = String(name || '').trim(); if (!name) return;
    const key = (d.code || '') + '|' + ciNorm_(name);
    if (have[key]) return; have[key] = true;
    append_('Check-ins', { id: base + added, day: day, code: d.code || '', household: d.household || (walkin ? 'Walk-in' : ''),
                           name: name, walkin: walkin ? 'Yes' : '', by: admin ? 'Committee' : 'Helper', created: now });
    added++;
  };
  (d.names || []).forEach(n => add(n, !!d.walkin));
  (d.walkins || []).forEach(n => add(n, true));
  if (added && d.code) { try { const r = regByCode_(d.code);     // checking in marks the family as seen on the contact list
    upsertContacts_([{ email: r.email, first: r.first, last: r.last, phone: r.phone, city: r.city, state: r.state, lastSeen: now }], 'Check-in'); } catch (e) {} }
  return { ok: true, added: added, stats: ciStats_(checkins_()) };
}

/* ---------- daily check ----------
   Runs every morning (setup() turns it on). If something looks off, it emails the reunion account
   and anyone listed under "Also email alerts to" in Setup. The same notes show on the Overview page. */
function health_() {
  const out = [];
  try {
    const regs = read_('Registrations'), pays = read_('Payments'), s = settings_();
    const byCode = {}; regs.forEach(r => byCode[r.code] = r);
    pays.filter(p => !p.code).forEach(p => out.push('A ' + money_(p.amount) + ' payment from ' + (p.from || 'someone') + ' isn’t matched to a household. Open Payments and match it.'));
    pays.filter(p => p.code && !byCode[p.code]).forEach(p => out.push('A payment from ' + (p.from || 'someone') + ' is marked ' + p.code + ', but there’s no registration with that number.'));
    const paid = {}; pays.forEach(p => { if (p.code) paid[p.code] = (paid[p.code] || 0) + (Number(p.amount) || 0); });
    regs.forEach(r => { const t = Number(r.total) || 0, pd = paid[r.code] || 0;
      if (t > 0 && pd > t + 0.01) out.push(r.first + ' ' + r.last + ' (' + r.code + ') has paid ' + money_(pd) + ', more than their ' + money_(t) + ' total. Check for a duplicate payment.'); });
    const q = ss_().getSheetByName(QUEUE_TAB);
    if (!TEST_ && q && q.getLastRow() > 1) out.push((q.getLastRow() - 1) + ' email(s) are waiting on the Email Queue tab because of the daily limit. They’ll go out on their own.');
    const since = Utilities.formatDate(new Date(Date.now() - 2 * 864e5), TZ, 'yyyy-MM-dd');
    readLog_().filter(l => l.kind === 'bad' && String(l.date) >= since).forEach(l => out.push(l.text));
    if (s.regOpen && !/^https:\/\/(www\.)?zeffy\.com\//i.test(s.zeffyUrl || '')) out.push('Registration is open, but the Zeffy payment link is missing. Families see "Online payment opens soon." Ask the reunion committee to add the link.');
  } catch (err) { out.push('The daily check ran into a problem: ' + err.message); }
  return out;
}

/* ---------- hosting committee PINs and the hand-off to the next host ----------
   - A host PIN is at least 6 numbers. It is stored scrambled (never as plain text).
   - After 5 wrong tries in 15 minutes, signing in pauses for 10 minutes.
   - HANDOFF_DAYS after the last reunion day, the page locks: the past host's PIN stops working,
     families see "Is your family hosting the next reunion?", and the next host signs up there.
   - Signing up saves last year's records to their own spreadsheet, clears the tabs, and emails
     the committee (Setup > Emails > committee emails) with an Undo button. Nothing is deleted. */
function fam_(s) { s = s || settings_(); return String(s.family || '').trim() || 'Porter-Sanford-Reid Family Reunion'; }
const PROPS_ = () => PropertiesService.getScriptProperties();
function emailList_(v) { return String(v || '').split(/[,;\s]+/).map(x => x.trim().toLowerCase()).filter(x => /^\S+@\S+\.\S+$/.test(x)).filter((x, i, a) => a.indexOf(x) === i); }
function pinHash_(pin, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pin, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function setHostPin_(pin) { const salt = Utilities.getUuid(); PROPS_().setProperty('HOST_PIN', salt + ':' + pinHash_(pin, salt)); }
function hostPinOk_(pin) {
  const v = PROPS_().getProperty('HOST_PIN'); if (!v || !pin) return false;
  const i = v.indexOf(':'); return pinHash_(String(pin), v.slice(0, i)) === v.slice(i + 1);
}
function host_() { try { return JSON.parse(PROPS_().getProperty('HOST') || 'null'); } catch (e) { return null; } }
function pinRuleError_(pin) {
  pin = String(pin || '');
  if (!/^\d+$/.test(pin)) return 'Use numbers only for your PIN.';
  if (pin.length < 6) return 'Your PIN must be at least 6 numbers.';
  if (/^(\d)\1+$/.test(pin) || '01234567890'.indexOf(pin) >= 0 || '09876543210'.indexOf(pin) >= 0) return 'That PIN is too easy to guess. Please pick a different one.';
  if (pin === String(COMMITTEE_PIN) || pin === String(TECH_PIN) || tryByPin_(pin)) return 'Please pick a different PIN.';
  return '';
}
function pinPaused_() { return CacheService.getScriptCache().get('pinPause') === '1'; }
function wrongPin_() {
  const c = CacheService.getScriptCache(), n = Number(c.get('pinBad') || 0) + 1;
  if (n >= 5) { c.put('pinPause', '1', 600); c.remove('pinBad'); try { TEST_ = false; log_('warn', 'Several wrong PINs were tried, so signing in is paused for 10 minutes.'); } catch (e) {} }
  else c.put('pinBad', String(n), 900);
}
function handoffDate_(s) {
  if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(String(s.lastDay || ''))) return '';
  const d = new Date(s.lastDay + 'T12:00:00'); d.setDate(d.getDate() + HANDOFF_DAYS);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function handoffOpen_(s) { const h = handoffDate_(s); return !!h && Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') >= h; }
function handoffInfo_() {
  const s = settings_(), yr = String(s.lastDay || '').slice(0, 4);
  return { open: handoffOpen_(s), pending: !!pendingHost_(), date: handoffDate_(s), year: yr, nextYear: yr ? String(Number(yr) + 1) : '' };
}
// A family asks to host the next reunion. Nothing changes yet: the PSR Reunion Board gets an email with
// Accept and Deny buttons. On Accept, the new host is set up and gets an email with how to get started.
function newHost_(d) {
  const s = settings_();
  if (!handoffOpen_(s)) return { ok: false, error: 'The page isn’t ready for a new host yet. It opens ' + HANDOFF_DAYS + ' days after the reunion.' };
  if (pendingHost_()) return { ok: false, error: 'Someone has already asked to host the next reunion, and the PSR Reunion Board is reviewing it. Please check back later.' };
  const name = String(d.name || '').trim().replace(/\s+/g, ' '), email = String(d.email || '').trim();
  if (name.length < 3) return { ok: false, error: 'Please type your first and last name.' };
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, error: 'Please type your email address.' };
  const ly = Number(String(s.lastDay || '').slice(0, 4)) || new Date().getFullYear(), year = String(ly + 1);   // the next reunion's year, worked out for them
  const board = boardSendTo_();
  if (!board.length) return { ok: false, error: 'The PSR Reunion Board email isn’t set, so your request can’t be sent. Please contact the board.' };
  const token = Utilities.getUuid().replace(/-/g, '');
  PROPS_().setProperty('PENDING_HOST', JSON.stringify({ t: token, name: name, email: email, year: year, asked: stamp_() }));
  const url = ScriptApp.getService().getUrl(), yes = url + '?act=accept&t=' + token, no = url + '?act=deny&t=' + token;
  const title = name + ' is asking to host the ' + year + ' ' + fam_(s);
  mail_({ to: board.join(','), subject: boardNote_() + 'Please review: ' + title,
    body: title + '.\n\nEmail: ' + email + '\n\nTo accept, open: ' + yes + '\nTo deny, open: ' + no + '\n\nNothing changes until you choose.\n',
    htmlBody: '<div style="font:16px/1.6 Arial,sans-serif;color:#17221b;max-width:560px">' +
      '<p style="font-size:18px"><b>' + esc_(name) + '</b> is asking to host the <b>' + esc_(year) + '</b> ' + esc_(fam_(s)) + '.</p>' +
      '<p>Email: ' + esc_(email) + '</p><p>Please accept or deny. Nothing changes until you choose.</p>' +
      '<p><a href="' + yes + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 26px;border-radius:10px;margin-right:10px">Accept</a>' +
      '<a href="' + no + '" style="display:inline-block;background:#a3252b;color:#fff;text-decoration:none;font-weight:bold;padding:12px 26px;border-radius:10px">Deny</a></p></div>' });
  mail_({ to: email, subject: 'We got your request to host the ' + year + ' ' + fam_(s),
    body: 'Hello ' + name.split(' ')[0] + ',\n\nThank you for offering to host! Your request was sent to the PSR Reunion Board. When they approve it, you’ll get another email with a link to create your PIN and get started.\n' });
  log_('info', name + ' asked to host the ' + year + ' reunion. Waiting for the PSR Reunion Board to accept or deny.');
  return { ok: true, pending: true };
}
/* ---------- Practice: the "ask to host" request, for real by email ----------
   In practice, the request email (with Accept and Deny) goes to TECH_EMAIL, or to a tester's own email.
   Accept also sends the "you're approved" email. Nothing touches the real reunion. */
function pracKey_() { return TRY_ ? 'try-' + TRY_.id : 'tech'; }
function practiceRequest_(d) {
  const name = String(d.name || '').trim().replace(/\s+/g, ' '), email = String(d.email || '').trim();
  if (name.length < 3) return { ok: false, error: 'Please type your first and last name.' };
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, error: 'Please type your email address.' };
  const k = pracKey_(), to = TECH_EMAIL, t = Utilities.getUuid().replace(/-/g, '');
  PROPS_().setProperty('PRAC_REQ_' + k, JSON.stringify({ t: t, name: name, email: email, status: 'pending' }));
  const url = ScriptApp.getService().getUrl(), yes = url + '?act=praccept&k=' + k + '&t=' + t, no = url + '?act=prdeny&k=' + k + '&t=' + t;
  mail_({ to: to, subject: 'PRACTICE – Please review: ' + name + ' is asking to host the next reunion',
    body: 'PRACTICE ONLY. In real life, this goes to the PSR Reunion Board.\n' + (TRY_ ? 'From the tryout you sent to ' + TRY_.email + '.\n' : '') + '\n' + name + ' is asking to host the next reunion.\nEmail: ' + email + '\n\nAccept: ' + yes + '\nDeny: ' + no + '\n',
    htmlBody: '<div style="font:16px/1.6 Arial,sans-serif;color:#17221b;max-width:560px"><p style="background:#fbf3d6;padding:8px 12px;border-radius:8px"><b>Practice only.</b> In real life, this email goes to the PSR Reunion Board.' + (TRY_ ? ' From the tryout you sent to ' + esc_(TRY_.email) + '.' : '') + '</p>' +
      '<p style="font-size:18px"><b>' + esc_(name) + '</b> is asking to host the next reunion.</p><p>Email: ' + esc_(email) + '</p><p>Please accept or deny.</p>' +
      '<p><a href="' + yes + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 26px;border-radius:10px;margin-right:10px">Accept</a>' +
      '<a href="' + no + '" style="display:inline-block;background:#a3252b;color:#fff;text-decoration:none;font-weight:bold;padding:12px 26px;border-radius:10px">Deny</a></p></div>' });
  return { ok: true, sentTo: to };
}
function practiceStatus_() {
  let r = null; try { r = JSON.parse(PROPS_().getProperty('PRAC_REQ_' + pracKey_()) || 'null'); } catch (e) {}
  return { ok: true, status: r ? r.status : 'none', name: r ? r.name : '' };
}
function practiceAnswer_(k, t, yes) {
  let r = null; try { r = JSON.parse(PROPS_().getProperty('PRAC_REQ_' + k) || 'null'); } catch (e) {}
  if (!r || r.t !== t || r.status !== 'pending') return { ok: false, text: 'This practice request was already answered.' };
  r.status = yes ? 'accepted' : 'denied'; r.pt = Utilities.getUuid().replace(/-/g, ''); PROPS_().setProperty('PRAC_REQ_' + k, JSON.stringify(r));
  const first = r.name.split(' ')[0], pinLink = SITE_URL + '#prpin=' + k + '.' + r.pt;
  if (yes) mail_({ to: r.email, subject: 'PRACTICE – You’re approved as host of the next reunion!',
    body: 'PRACTICE ONLY.\n\nHello ' + first + ',\n\nGood news! The PSR Reunion Board approved you as host of the next reunion.\n\nOpen this link to create your PIN:\n' + pinLink + '\n',
    htmlBody: '<div style="font:16px/1.6 Arial,sans-serif;color:#17221b;max-width:560px"><p style="background:#fbf3d6;padding:8px 12px;border-radius:8px"><b>Practice only.</b> Nothing real is changed.</p>' +
      '<p>Hello ' + esc_(first) + ',</p><p style="font-size:18px"><b>Good news!</b> The PSR Reunion Board approved you as host of the next reunion.</p><p>Click the button to create your PIN. You’ll use it to sign in.</p>' +
      '<p><a href="' + pinLink + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:10px">Create my PIN</a></p></div>' });
  else mail_({ to: r.email, subject: 'PRACTICE – About your request to host the next reunion',
    body: 'PRACTICE ONLY.\n\nHello ' + first + ',\n\nThank you for offering to host. The PSR Reunion Board didn’t approve this request.\n' });
  return { ok: true, text: yes ? 'Accepted. ' + esc_(r.name) + ' was emailed a link to create their PIN.' : 'Denied. ' + esc_(r.name) + ' has been told.' };
}
function practicePin_(d) {
  const parts = String(d.token || '').split('.'), k = parts[0] || '', t = parts[1] || '';
  let r = null; try { r = JSON.parse(PROPS_().getProperty('PRAC_REQ_' + k) || 'null'); } catch (e) {}
  if (!r || !(r.status === 'accepted' || r.status === 'pinmade') || !r.pt || r.pt !== t) return { ok: false, error: 'This practice link was already used or has ended.' };
  const bad = pinRuleError_(d.pin); if (bad) return { ok: false, error: bad };
  const salt = Utilities.getUuid(); r.pin = salt + ':' + pinHash_(String(d.pin).trim(), salt);
  r.status = 'pinmade'; r.pt = ''; PROPS_().setProperty('PRAC_REQ_' + k, JSON.stringify(r));   // signs in to practice as the new host
  return { ok: true };
}
function practiceHosts_() {
  const P = PROPS_().getProperties(), out = [];
  Object.keys(P).forEach(key => { if (key.indexOf('PRAC_REQ_') !== 0) return; try { const r = JSON.parse(P[key]); r.k = key.slice(9); out.push(r); } catch (e) {} });
  return out;
}
function practiceHostByPin_(pin) {
  return practiceHosts_().filter(r => {
    if (r.status !== 'pinmade' || !r.pin) return false;
    if (r.k.indexOf('try-') === 0 && !tryById_(r.k.slice(4))) return false;           // that tryout has ended
    const parts = String(r.pin).split(':'); return parts.length === 2 && pinHash_(pin, parts[0]) === parts[1];
  })[0] || null;
}
function pendingHost_() { try { return JSON.parse(PROPS_().getProperty('PENDING_HOST') || 'null'); } catch (e) { return null; } }
// The board clicked Accept: last year's records are saved, the page is cleared for the new host, and they get a welcome email.
function acceptHost_(t) {
  const p = pendingHost_();
  if (!p || !t || p.t !== t) return { ok: false, text: 'This request was already answered, or it’s no longer waiting.' };
  const s = settings_(), prevYear = String(s.lastDay || '').slice(0, 4);
  const arch = archive_(prevYear);
  clearDataTabs_();
  writeSettings_(Object.assign({}, s, { lastDay: '', startDay: '', regOpensOn: '', regClosesOn: '', earlyBird: '', regOpen: false, setupDone: false, setupStep: 0,
    replyTo: p.email, checkinCode: '', checkinCodeYear: '' }));
  setHostPin_(Utilities.getUuid());     // the old host's PIN stops working; the new host makes theirs from the email link
  PROPS_().setProperty('HOST', JSON.stringify({ name: p.name, email: p.email, year: p.year, since: stamp_() }));
  PROPS_().deleteProperty('PENDING_HOST');
  const pinTok = Utilities.getUuid().replace(/-/g, ''), pinLink = SITE_URL + '#newhost=' + pinTok;
  PROPS_().setProperty('PIN_RESET', JSON.stringify({ t: pinTok, exp: Date.now() + 14 * 864e5, first: true }));
  TEST_ = false;
  log_('ok', 'The PSR Reunion Board accepted ' + p.name + ' as host of the ' + p.year + ' reunion. Last year’s records were saved to their own spreadsheet.');
  const first = p.name.split(' ')[0], fam = fam_(s);
  mail_({ to: p.email, subject: 'You’re approved as host of the ' + p.year + ' ' + fam + '!',
    body: 'Hello ' + first + ',\n\nGood news! The PSR Reunion Board approved you as host of the ' + p.year + ' ' + fam + '.\n\nOpen this link to create your PIN. You’ll use it to sign in:\n' + pinLink + '\n\nAfter that, the page walks you through setting up the reunion step by step. The link works for 14 days.\n',
    htmlBody: '<div style="font:16px/1.6 Arial,sans-serif;color:#17221b;max-width:560px"><p>Hello ' + esc_(first) + ',</p><p style="font-size:18px"><b>Good news!</b> The PSR Reunion Board approved you as host of the <b>' + esc_(p.year) + ' ' + esc_(fam) + '</b>.</p>' +
      '<p>Click the button to create your PIN. You’ll use it to sign in.</p>' +
      '<p><a href="' + pinLink + '" style="display:inline-block;background:#1c5a34;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:10px">Create my PIN</a></p>' +
      '<p>After that, the page walks you through setting up the reunion step by step.</p><p style="color:#56655b">This button works for 14 days.</p></div>' });
  return { ok: true, text: esc_(p.name) + ' is now the host of the ' + esc_(p.year) + ' reunion. They’ve been emailed a link to create their PIN. Last year’s records, including a Ledger tab, were saved and shared with the board (view only): <a href="' + arch.getUrl() + '">last year’s records</a>.' };
}
// The board clicked Deny: the request is removed and the person is told.
function denyHost_(t) {
  const p = pendingHost_();
  if (!p || !t || p.t !== t) return { ok: false, text: 'This request was already answered, or it’s no longer waiting.' };
  PROPS_().deleteProperty('PENDING_HOST');
  TEST_ = false;
  log_('warn', 'The PSR Reunion Board denied ' + p.name + '’s request to host the ' + p.year + ' reunion.');
  const s = settings_(), board = emailList_(boardEmails_());
  mail_({ to: p.email, subject: 'About your request to host the ' + p.year + ' ' + fam_(s),
    body: 'Hello ' + p.name.split(' ')[0] + ',\n\nThank you for offering to host. The PSR Reunion Board didn’t approve this request.' + (board.length ? ' If you have questions, please contact the board at ' + board[0] + '.' : '') + '\n' });
  return { ok: true, text: 'Denied. ' + esc_(p.name) + ' has been told, and the page is waiting for a new host again.' };
}
function undoHost_(t) {
  let u = null; try { u = JSON.parse(PROPS_().getProperty('UNDO_HOST') || 'null'); } catch (e) {}
  if (!u || !t || u.t !== t) return { ok: false, text: 'This Undo link has already been used, or a newer host has signed up since.' };
  let restored = false;
  try {
    const sh = SpreadsheetApp.openById(u.archiveId).getSheetByName('Settings');
    const v = sh ? sh.getDataRange().getValues() : [];
    for (let i = 1; i < v.length; i++) if (v[i][0] === 'settings') { writeSettings_(JSON.parse(v[i][1])); restored = true; }
  } catch (e) {}
  if (u.prev && u.prev.pin) PROPS_().setProperty('HOST_PIN', u.prev.pin); else PROPS_().deleteProperty('HOST_PIN');
  if (u.prev && u.prev.host) PROPS_().setProperty('HOST', JSON.stringify(u.prev.host)); else PROPS_().deleteProperty('HOST');
  PROPS_().deleteProperty('UNDO_HOST');
  TEST_ = false;
  log_('warn', 'The committee clicked Undo: ' + u.name + ' is no longer the host. The page is waiting for a new host again.');
  return { ok: true, text: esc_(u.name) + ' no longer has access. The registration page is back to waiting for a new host. Nothing was deleted.' +
    (restored ? '' : ' (Last year’s settings couldn’t be put back automatically. They’re still saved in the records spreadsheet.)') };
}
function archive_(yr) {
  const src = ss_(), when = Utilities.formatDate(new Date(), TZ, 'MMM d, yyyy');
  const dest = SpreadsheetApp.create('PSR Reunion Hub – ' + (yr || 'past') + ' reunion records (saved ' + when + ')');
  src.getSheets().forEach(sh => { const n = sh.getName(); if (n.indexOf(TEST_PREFIX) === 0 || n === QUEUE_TAB) return; sh.copyTo(dest).setName(n); });
  const blank = dest.getSheetByName('Sheet1'); if (blank && dest.getSheets().length > 1) dest.deleteSheet(blank);
  try { ledgerTab_(dest); } catch (e) { log_('warn', 'The saved records don’t have a Ledger tab: ' + e.message); }
  // The PSR Reunion Board can always open past years' records, but can't change them.
  boardSendTo_().forEach(e => { try { dest.addViewer(e); } catch (err) { log_('warn', 'Couldn’t share last year’s records with ' + e + ': ' + err.message); } });
  let list = []; try { list = JSON.parse(PROPS_().getProperty('ARCHIVES') || '[]'); } catch (e) {}
  list.push({ year: yr, url: dest.getUrl(), saved: when }); PROPS_().setProperty('ARCHIVES', JSON.stringify(list.slice(-30)));
  return dest;
}
// A Ledger tab for the saved records: every dollar in and out of the reunion money, by date, with a running balance.
// (Scholarship money is kept separate, as on the page.)
function ledgerTab_(dest) {
  const s = settings_(), num = v => Number(v) || 0, rows = [];
  read_('Payments').forEach(p => rows.push([String(p.date || ''), 'Payment from ' + (p.household || p.from || ''), p.method || p.via || '', num(p.amount), '']));
  read_('Other Money In').forEach(c => { if (/^scholar/i.test(String(c.fund || ''))) return;
    rows.push([String(c.date || ''), (c.type || 'Money in') + (c.name ? ' – ' + c.name : ''), c.via || '', num(c.amount), '']); });
  read_('Expenses').forEach(x => rows.push([String(x.date || ''), (x.vendor || 'Expense') + (x.note ? ' – ' + x.note : ''), x.category || '', '', num(x.amount)]));
  rows.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  let bal = num(s.begin), tin = 0, tout = 0;
  const out = [['Date', 'Description', 'Category or how paid', 'In', 'Out', 'Balance'], ['', 'Starting balance', '', '', '', bal]];
  rows.forEach(r => { tin += num(r[3]); tout += num(r[4]); bal += num(r[3]) - num(r[4]); out.push(r.concat([bal])); });
  out.push(['', 'Total', '', tin, tout, bal]);
  const sh = dest.insertSheet('Ledger', 0);
  sh.getRange(1, 1, out.length, 6).setValues(out);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold'); sh.getRange(out.length, 1, 1, 6).setFontWeight('bold');
  sh.getRange(2, 4, out.length - 1, 3).setNumberFormat('$#,##0.00'); sh.setFrozenRows(1); sh.autoResizeColumns(1, 6);
}
function clearDataTabs_() {
  const ss = ss_();
  Object.keys(TABS).forEach(name => {
    if (name === 'Settings' || name === 'Contacts') return;      // the contact list carries over to the next host
    const sh = ss.getSheetByName(name); if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
}
function forgotPin_(d) {
  const c = CacheService.getScriptCache(); if (c.get('pinForgot')) return { ok: true };
  c.put('pinForgot', '1', 60);
  const email = String(d.email || '').trim().toLowerCase(), h = host_(), s = settings_();
  const known = [h && h.email, s.replyTo].map(x => String(x || '').trim().toLowerCase()).filter(Boolean);
  const pr = email ? practiceHosts_().filter(r => r.status === 'pinmade' && String(r.email || '').toLowerCase() === email)[0] : null;
  if (pr && known.indexOf(email) < 0) {
    pr.pt = Utilities.getUuid().replace(/-/g, ''); const k = pr.k; delete pr.k;
    PROPS_().setProperty('PRAC_REQ_' + k, JSON.stringify(pr));
    mail_({ to: email, subject: 'PRACTICE – Make a new PIN',
      body: 'PRACTICE ONLY.\n\nOpen this link to make a new PIN for your practice:\n' + SITE_URL + '#prpin=' + k + '.' + pr.pt + '\n\nIf you didn’t ask for this, you can ignore this email.\n' });
    return { ok: true };
  }
  if (email && known.indexOf(email) >= 0) {
    const token = Utilities.getUuid().replace(/-/g, '');
    PROPS_().setProperty('PIN_RESET', JSON.stringify({ t: token, exp: Date.now() + 3600000 }));
    TEST_ = false;
    mail_({ to: email, subject: 'Make a new PIN for the PSR registration page',
      body: 'Hello,\n\nSomeone asked to make a new PIN for the hosting committee sign-in. If that was you, open this link. It works for 1 hour:\n' +
            SITE_URL + '#resetpin=' + token + '\n\nIf you didn’t ask for this, you can ignore this email. Your PIN hasn’t changed.\n' });
    log_('info', 'A new-PIN link was emailed to the host.');
  }
  return { ok: true };       // same answer either way, so nobody can test which emails are on file
}
function resetPin_(d) {
  let r = null; try { r = JSON.parse(PROPS_().getProperty('PIN_RESET') || 'null'); } catch (e) {}
  if (!r || r.t !== String(d.token || '') || Date.now() > r.exp) return { ok: false, error: 'This link has expired or was already used. Go back to the sign-in box and click "Forgot your PIN?" again.' };
  const bad = pinRuleError_(d.pin); if (bad) return { ok: false, error: bad };
  setHostPin_(String(d.pin).trim());
  PROPS_().deleteProperty('PIN_RESET');
  TEST_ = false; log_('ok', r.first ? 'The new host created their PIN.' : 'The host made a new PIN.');
  return { ok: true };
}
/* ---------- family replies ----------
   Runs every 10 minutes (setup() turns it on). Each email that comes into the reunion account is
   saved on the Emails tab, noted in the Log, and sent on to the host's email from Setup ("Your email").
   Emails that bounce are flagged so the daily check reports them. */
function checkReplies() {
  TEST_ = false;
  const P = PropertiesService.getScriptProperties(), start = Date.now();
  const since = Number(P.getProperty('REPLIES_SINCE')) || (start - 864e5);
  let seen = []; try { seen = JSON.parse(P.getProperty('REPLIES_SEEN') || '[]'); } catch (e) {}
  const me = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const threads = GmailApp.search('in:inbox after:' + Math.floor((since - 600000) / 1000), 0, 100);
  if (threads.length) {
    const s = settings_(), hostRaw = String(s.replyTo || '').trim(), host = /^\S+@\S+\.\S+$/.test(hostRaw) ? hostRaw : '';
    const byEmail = {}; read_('Registrations').forEach(r => { if (r.email) byEmail[String(r.email).trim().toLowerCase()] = r; });
    const known = {}; try { read_('Contacts').forEach(c => { if (c.email) known[String(c.email).trim().toLowerCase()] = 1; }); } catch (e) {}
    threads.forEach(th => { const msgs = th.getMessages(), firstFrom = String((msgs[0] && msgs[0].getFrom()) || '').toLowerCase();
      const ourThread = !!me && firstFrom.indexOf(me) >= 0;   // a reply to an email the reunion system sent
      msgs.forEach(m => {
      const id = m.getId(); if (seen.indexOf(id) >= 0 || m.getDate().getTime() < since - 600000) return;
      seen.push(id);
      const fromRaw = String(m.getFrom() || ''), em = ((fromRaw.match(/<([^>]+)>/) || [])[1] || fromRaw).trim().toLowerCase();
      const subject = String(m.getSubject() || '(no subject)');
      if (!em || em === me) return;
      const isBounce = /mailer-daemon|postmaster/.test(em);
      // Only reunion mail: a reply to a reunion email, or an email from a family or contact. Everything else in the inbox (ads, other mail) is left alone.
      if (!ourThread && !byEmail[em] && !known[em] && !(isBounce && /reunion/i.test(String(m.getPlainBody() || '').slice(0, 4000)))) return;
      if (isBounce) { log_('bad', 'An email didn’t reach a family: "' + subject.slice(0, 80) + '". Check the address on the Emails tab.'); 
        append_('Emails', { date: Utilities.formatDate(m.getDate(), TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(m.getDate(), TZ, 'h:mm a'), direction: 'Bounced', who: 'Email didn’t go through', email: em, subject: subject.slice(0, 150), note: cleanBody_(m.getPlainBody()) }); return; }
      if (/no-?reply|notifications?@|zeffy\.com$|google\.com$/.test(em)) return;
      const r = byEmail[em], who = r ? r.first + ' ' + r.last : (fromRaw.replace(/<[^>]*>/, '').replace(/"/g, '').trim() || em);
      let fwd = '';
      if (host && host.toLowerCase() !== em) { try { m.forward(host, { name: senderName_() }); fwd = host; } catch (e) { log_('warn', 'Couldn’t send a reply from ' + who + ' on to ' + host + ': ' + e.message); } }
      append_('Emails', { date: Utilities.formatDate(m.getDate(), TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(m.getDate(), TZ, 'h:mm a'), direction: 'Received', who: who, email: em,
                           code: r ? r.code : '', subject: subject.slice(0, 150), note: cleanBody_(String(m.getPlainBody() || '').split(/\n\s*On .*wrote:|\n>/)[0]), forwardedTo: fwd });
      log_('info', 'Email from ' + who + (r ? ' (' + r.code + ')' : '') + ': "' + subject.slice(0, 80) + '".' + (fwd ? ' Sent on to ' + fwd + '.' : ''));
    }); });
  }
  P.setProperty('REPLIES_SINCE', String(start));
  P.setProperty('REPLIES_SEEN', JSON.stringify(seen.slice(-300)));
}
// Readable first lines of an email: no web links, codes or technical lines.
function cleanBody_(t) {
  return String(t || '').split(/\r?\n/).filter(l => !/^\s*(content-|mime-|charset=)/i.test(l)).join(' ')
    .replace(/&#?[a-z0-9]+;/gi, ' ').replace(/https?:\/\/\S+/g, '').replace(/\[\s*\]|\(\s*\)/g, '')
    .replace(/[\u200b-\u200f\u034f\u00ad\ufeff]/g, '').replace(/\s+/g, ' ').trim().slice(0, 200);
}
// Run this once (in Apps Script, pick startEmailsOver at the top, then click Run) to empty the Emails tab and start fresh.
// Only the Emails tab is cleared. Registrations, payments and everything else stay.
function startEmailsOver() {
  const sh = realSS_().getSheetByName('Emails'); if (!sh) return;
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  formatTab_(sh, 'Emails', sh.getRange(1, 1).getBackground());
  PropertiesService.getScriptProperties().setProperty('REPLIES_SINCE', String(Date.now()));
  PropertiesService.getScriptProperties().deleteProperty('REPLIES_SEEN');
}
function dailyCheck() {
  TEST_ = false; PRAC_ = false; TRY_ = null;
  try { cleanTryouts_(); } catch (e) { Logger.log('Tryouts: ' + e.message); }
  try { dateReminders_(); } catch (e) { Logger.log('Date reminders: ' + e.message); }
  const issues = health_();
  if (!issues.length) return;
  const s = settings_();
  const to = [Session.getEffectiveUser().getEmail(), s.replyTo].concat(String(s.alertEmails || '').split(/[,;\s]+/))
    .filter(e => /^\S+@\S+\.\S+$/.test(e)).filter((e, i, a) => a.indexOf(e) === i).join(',');
  if (!to) return;
  const text = 'The reunion system’s daily check found ' + issues.length + ' thing(s) to look at:\n\n' +
    issues.map((x, i) => (i + 1) + '. ' + x).join('\n') + '\n\nThese also show at the top of the Overview page when you sign in.';
  mail_({ to: to, subject: 'Reunion system: ' + issues.length + ' thing(s) to check', body: text });
}

// On the day the host picked to open (or close) registration, the host gets a reminder to flip the switch.
// Nothing opens or closes by itself.
function dateReminders_() {
  const s = settings_(), today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), h = host_();
  const to = [s.replyTo, h && h.email].map(x => String(x || '').trim()).filter(x => /^\S+@\S+\.\S+$/.test(x)).filter((x, i, a) => a.indexOf(x) === i).join(',');
  if (!to) return;
  const title = [s.edition, fam_(s)].filter(Boolean).join(' ');
  const how = 'Sign in at ' + SITE_URL + ', click Setup, then the last step (Open registration), and use the switch.';
  if (s.regOpensOn === today && !s.regOpen)
    mail_({ to: to, subject: 'Reminder: today is the day you planned to open registration', body: 'Hello,\n\nToday is the day you planned to open registration for the ' + title + '. When you’re ready, turn registration on. ' + how + '\n\nNothing opens by itself, so families can’t sign up until you turn it on.\n' });
  openNudge_(s, today, to, title, how);
  if (s.regClosesOn === today && s.regOpen)
    mail_({ to: to, subject: 'Reminder: today is your registration closing date', body: 'Hello,\n\nToday is the closing date you set for registration for the ' + title + '. When you’re ready, turn registration off. ' + how + '\n\nNothing closes by itself, so families can still sign up until you turn it off.\n' });
}
// Registration still closed? Remind the host by email. With an early-bird date: 60 and 30 days before it ends.
// Without one: 120 and 90 days before the last day of the reunion. Each reminder goes once. (The page shows the same as a pop-up.)
function openNudge_(s, today, to, title, how) {
  if (s.regOpen || TEST_ || PRAC_ || handoffOpen_(s)) return;
  const eb = /^\d{4}-\d\d-\d\d$/.test(s.earlyBird || '') ? s.earlyBird : '', ld = /^\d{4}-\d\d-\d\d$/.test(s.lastDay || '') ? s.lastDay : '';
  const ref = eb || ld; if (!ref || today > ref) return;
  const offs = eb ? [60, 30] : [120, 90];
  const day = (iso, n) => Utilities.formatDate(new Date(new Date(iso + 'T12:00:00').getTime() - n * 864e5), TZ, 'yyyy-MM-dd');
  const due = offs.filter(n => today >= day(ref, n)); if (!due.length) return;
  const key = ref + ':' + due[due.length - 1], P = PROPS_(), sent = P.getProperty('OPEN_NUDGE') || '';
  if (sent.split('|').indexOf(key) >= 0) return;
  const nice = Utilities.formatDate(new Date(ref + 'T12:00:00'), TZ, 'MMMM d, yyyy');
  const why = eb ? 'Early-bird pricing ends ' + nice + ', so families need time to sign up before prices go up.'
                 : 'The reunion is ' + Math.round((new Date(ref + 'T12:00:00') - new Date(today + 'T12:00:00')) / 864e5) + ' days away (last day ' + nice + '), so families need time to sign up.';
  mail_({ to: to, subject: 'Reminder: registration for the ' + title + ' isn’t open yet',
    body: 'Hello,\n\nRegistration for the ' + title + ' isn’t open yet. ' + why + '\n\n' + how + '\n\nThank you!\n' });
  P.setProperty('OPEN_NUDGE', (sent ? sent + '|' : '') + key);
  log_('info', 'Reminded the host that registration isn’t open yet.');
}
/* ---------- contact list ----------
   One row per household, matched by email (or by name when there's no email). It carries over to the next host.
   Registrations add to it by themselves; a Zeffy contact export can be brought in from the Records page. */
function contactKey_(c) {
  const e = String(c.email || '').trim().toLowerCase();
  return e ? e : 'name:' + (String(c.first || '') + ' ' + String(c.last || '')).trim().toLowerCase().replace(/\s+/g, ' ');
}
function upsertContacts_(rows, source) {
  const sh = sheet_('Contacts'), have = read_('Contacts'), idx = {};
  have.forEach((c, i) => { idx[contactKey_(c)] = i; });
  let added = 0, updated = 0, skipped = 0;
  const F = ['first', 'last', 'phone', 'address', 'city', 'state', 'zip'];
  rows.forEach(raw => {
    const c = {}; ['email'].concat(F).forEach(k => c[k] = String(raw[k] == null ? '' : raw[k]).trim());
    if (raw.optOut) c.optOut = 'Yes';
    if (!c.email && !(c.first && c.last)) { skipped++; return; }
    const k = contactKey_(c);
    if (idx[k] === undefined) {
      append_('Contacts', Object.assign(c, { source: source, added: stamp_(), lastSeen: raw.lastSeen || '' }));
      idx[k] = have.length; have.push(c); added++;
    } else {
      const old = have[idx[k]], fields = {};
      F.forEach(f => { if (c[f] && !old[f]) fields[f] = c[f]; });
      if (raw.lastSeen) fields.lastSeen = raw.lastSeen;
      if (c.optOut && old.optOut !== 'Yes') fields.optOut = 'Yes';
      if (Object.keys(fields).length) { updateContact_(sh, contactKey_(old), fields); Object.assign(old, fields); updated++; }
    }
  });
  return { added: added, updated: updated, skipped: skipped };
}
function updateContact_(sh, key, fields) {
  const v = sh.getDataRange().getValues(), h = v[0].map(x => keyOf_('Contacts', x));
  for (let i = 1; i < v.length; i++) {
    const o = {}; h.forEach((k, j) => o[k] = v[i][j]);
    if (contactKey_(o) === key) { Object.keys(fields).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(i + 1, c + 1).setValue(cell_(k, fields[k])); }); return; }
  }
}
// The "No emails" box on the Records page. Group emails skip anyone marked; their own receipts still go out.
function setOptOut_(d) {
  const c = { email: d.email, first: d.first, last: d.last, phone: d.phone, city: d.city, state: d.state };
  if (!String(c.email || '').trim() && !(c.first && c.last)) throw new Error('Couldn’t find that contact. Refresh and try again.');
  const have = read_('Contacts'), k = contactKey_(c), old = have.filter(x => contactKey_(x) === k)[0];
  if (!old) append_('Contacts', Object.assign(c, { source: 'Registration', added: stamp_(), optOut: d.optOut ? 'Yes' : '' }));
  else updateContact_(sheet_('Contacts'), k, { optOut: d.optOut ? 'Yes' : '' });
  log_('info', (d.optOut ? 'Marked ' : 'Unmarked ') + (c.first || '') + ' ' + (c.last || '') + ' as "No emails".');
  return { ok: true };
}
function importContacts_(rows, source) {
  if (!rows.length) return { ok: false, error: 'That file has no contacts in it.' };
  if (rows.length > 3000) return { ok: false, error: 'That file is too big. Please split it into smaller files.' };
  const r = upsertContacts_(rows, source);
  log_('ok', 'Brought in contacts from ' + source + ': ' + r.added + ' new, ' + r.updated + ' updated' + (r.skipped ? ', ' + r.skipped + ' skipped (no name or email)' : '') + '.');
  return Object.assign({ ok: true }, r);
}
function ciReg_(r) {
  return { code: r.code, date: r.date, first: r.first, last: r.last, email: r.email, phone: r.phone, city: r.city, state: r.state,
           people: r.people || [], shirts: r.shirts || {}, addons: r.addons || {}, donation: Number(r.donation) || 0, emailOK: r.emailOK || '' };
}
function paidFor_(code) {
  return read_('Payments').filter(p => p.code === code).reduce((s, p) => s + (Number(p.amount) || 0), 0);
}
function regByCode_(code) {
  const r = read_('Registrations').filter(x => x.code === code)[0];
  if (!r) throw new Error('Couldn’t find that registration. Search again.');
  return r;
}
// Payment taken at the check-in table (cash, Cash App, Zelle). A receipt email goes out right away.
function ciPay_(d, admin) {
  const r = regByCode_(d.code), amt = Math.round((Number(d.amount) || 0) * 100) / 100;
  if (!(amt > 0)) throw new Error('Type the amount they paid.');
  const via = ['Cash', 'Cash App', 'Zelle', 'Check'].indexOf(d.via) >= 0 ? d.via : 'Cash';
  // A payment saved on the device while the internet was down carries its own id, so it's never recorded twice.
  const pid = Number(d.qid) || Date.now();
  if (d.qid && read_('Payments').some(p => Number(p.id) === pid)) return { ok: true, paid: paidFor_(r.code), emailed: false, dup: true };
  append_('Payments', { id: pid, date: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), from: r.first + ' ' + r.last, via: via,
                        memo: d.late ? 'At check-in (sent when the internet came back)' : 'At check-in', amount: amt, code: r.code, method: 'At check-in', takenBy: admin ? 'Committee' : 'Helper', created: stamp_() });
  setBalance_(pid, r.code, d.total);
  const paid = paidFor_(r.code);
  log_('ok', 'Took ' + money_(amt) + ' (' + via + ') from ' + r.first + ' ' + r.last + ' (' + r.code + ') at check-in.');
  let emailed = false;
  if (/^\S+@\S+\.\S+$/.test(r.email || '')) {
    const m = regEmail_(r, r.code, Number(d.total) || Number(r.total) || 0, d.lines || [], settings_(), paid);
    mail_({ to: r.email, subject: subj_(receiptSubject_(Number(d.total) || Number(r.total) || 0, paid)), body: m.text, htmlBody: m.html });
    emailed = true;
  }
  return { ok: true, paid: paid, emailed: emailed };
}
// Confirm or correct the household's email at check-in.
function ciEmail_(d) {
  const r = regByCode_(d.code), em = String(d.email || '').trim();
  if (!/^\S+@\S+\.\S+$/.test(em)) throw new Error('That email address doesn’t look right.');
  const fields = { emailOK: stamp_() };
  if (em !== r.email) {
    fields.email = em;
    log_('info', 'Updated ' + r.first + ' ' + r.last + '’s email at check-in (' + r.code + ').');
  }
  update_('Registrations', 'code', r.code, fields);
  return { ok: true };
}
// Add someone to an existing household at check-in. Their cost goes onto that household's balance.
function ciAdd_(d) {
  const r = regByCode_(d.code), p = d.person || {};
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Type the person’s name.');
  const people = (r.people || []).concat([{ name: name, group: p.group || 'adult' }]);
  const fields = { people: people };
  if (d.total !== undefined && !isNaN(Number(d.total))) fields.total = Math.round(Number(d.total) * 100) / 100;
  update_('Registrations', 'code', r.code, fields);
  log_('info', 'Added ' + name + ' to ' + r.first + ' ' + r.last + '’s household at check-in (' + r.code + ').');
  return { ok: true, reg: ciReg_(Object.assign({}, r, { people: people })) };
}

// Keeps the tabs grouped: all the red Test tabs together first, then the real tabs.
const TAB_ORDER = ['Registrations', 'Payments', 'Expenses', 'Other Money In', 'Check-ins', 'Contacts', 'Emails', 'Log'];
function orderTabs_() {
  const ss = ss_();
  const names = TAB_ORDER.concat(['Settings', QUEUE_TAB]);
  let pos = 1;
  names.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    ss.setActiveSheet(sh); ss.moveActiveSheet(pos++);
  });
}

/* ---------- column upkeep ----------
   When a new column is added to a tab, setup() inserts it in the right place in the sheet,
   so existing rows stay lined up under the correct headers. */
function addMissingCols_(sh, name) {
  const cols = TABS[name], last = sh.getLastColumn();
  if (!last || sh.getLastRow() < 1) return;
  let head = sh.getRange(1, 1, 1, last).getValues()[0].map(h => keyOf_(name, h));
  if (!head.some(h => cols.indexOf(h) >= 0)) return;          // blank tab: formatTab_ writes the header
  cols.forEach((k, i) => {
    if (head.indexOf(k) >= 0) return;
    // put it right after the column that comes before it in TABS
    let after = -1;
    for (let j = i - 1; j >= 0; j--) { after = head.indexOf(cols[j]); if (after >= 0) break; }
    if (after >= 0 && after < head.length - 1) { sh.insertColumnAfter(after + 1); head.splice(after + 1, 0, k); sh.getRange(1, after + 2).setValue(label_(name, k)); }
    else { head.push(k); sh.getRange(1, head.length).setValue(label_(name, k)); }
  });
}
function householdOf_(code) {
  const r = read_('Registrations').filter(x => x.code === code)[0];
  return r ? r.first + ' ' + r.last : '';
}
function fillHouseholds_() {
  const pays = read_('Payments');
  pays.forEach(p => { if (p.code && !p.household) update_('Payments', 'id', p.id, { household: householdOf_(p.code) }); });
}
