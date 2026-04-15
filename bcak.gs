// ═══════════════════════════════════════════════════════════════════════════
// EGG TRACKER — Google Apps Script Backend v4
// Fixes vs v3:
//   • appendRow crash fixed — onOpen guard: only runs when Egg Buys has data
//   • Predefined sheet: [UserID, Date, Qty] — per date, not standing default
//   • User can set/OFF predefined for today or any future date
//     Blocked if admin already entered consumption for that date
//   • changePassword action — user changes own password
//   • Admin dashboard: cashInHand added
//   • getConsumptionForUserDate returns -1 if not yet entered (vs 0)
// Run setupSheets() ONCE, then redeploy as Web App
// ═══════════════════════════════════════════════════════════════════════════

const SS2 = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  USERS:       'Users',
  CONSUMPTION: 'Daily Consumption',
  PAYMENTS:    'Payments',
  BUYS:        'Egg Buys',
  QUOTAS:      'Quotas',
  SESSIONS:    'Sessions',
  DESTROY:     'Egg Destroy',
  MISC:        'Misc Cost',
  PREDEFINED:  'Predefined',   // [UserID, Date, Qty]
  FIFO_RPT:    'FIFO Report',
};

function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { return makeResponse({ success: false, error: 'Use GET.' }); }

function doGet(e) {
  try {
    if (!e.parameter.payload) return makeResponse({ status: 'ok', message: 'EggTrack API v4 ✓' });
    const body = JSON.parse(e.parameter.payload);
    const { action } = body;
    if (action === 'login') return makeResponse(handleLogin(body));
    const user = verifyToken(body.token);
    if (!user) return makeResponse({ success: false, error: 'Session expired. Please log in again.' });
    switch (action) {
      case 'getDashboard':            return makeResponse(getDashboard(body, user));
      case 'getAdminDashboard':       return makeResponse(getAdminDashboard(user));
      case 'getDailyConsumptionData': return makeResponse(getDailyConsumptionData(body, user));
      case 'saveConsumption':         return makeResponse(saveConsumption(body, user));
      case 'getPayments':             return makeResponse(getPayments(user));
      case 'addPayment':              return makeResponse(addPayment(body, user));
      case 'deletePayment':           return makeResponse(deleteById(SHEETS.PAYMENTS, body.id, user));
      case 'getEggBuys':              return makeResponse(getEggBuys(user));
      case 'addEggBuy':               return makeResponse(addEggBuy(body, user));
      case 'deleteEggBuy':            return makeResponse(deleteById(SHEETS.BUYS, body.id, user));
      case 'getEggDestroy':           return makeResponse(getEggDestroy(user));
      case 'addEggDestroy':           return makeResponse(addEggDestroy(body, user));
      case 'deleteEggDestroy':        return makeResponse(deleteById(SHEETS.DESTROY, body.id, user));
      case 'getMiscCost':             return makeResponse(getMiscCost(user));
      case 'addMiscCost':             return makeResponse(addMiscCost(body, user));
      case 'deleteMiscCost':          return makeResponse(deleteById(SHEETS.MISC, body.id, user));
      case 'getFifoReport':           return makeResponse(getFifoReport(user));
      case 'refreshFifoSheet':        return makeResponse(refreshFifoSheet(user));
      case 'getUsers':                return makeResponse(getUsers(user));
      case 'createUser':              return makeResponse(createUser(body, user));
      case 'updateUser':              return makeResponse(updateUser(body, user));
      case 'deleteUser':              return makeResponse(deleteUserFn(body, user));
      case 'changePassword':          return makeResponse(changePassword(body, user));
      case 'getQuotas':               return makeResponse(getQuotas(user));
      case 'saveQuotas':              return makeResponse(saveQuotas(body, user));
      case 'setUserQuota':            return makeResponse(setUserQuota(body, user));
      case 'getPredefined':           return makeResponse(getPredefined(body, user));
      case 'setPredefined':           return makeResponse(setPredefined(body, user));
      default: return makeResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return makeResponse({ success: false, error: 'Server error: ' + err.message });
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SETUP
// ═══════════════════════════════════════════════════════════════════════════
function setupSheets() {
  const defs = {
    [SHEETS.USERS]:       ['UserID','Username','Password','Name','Role','Target','InitBalance'],
    [SHEETS.CONSUMPTION]: ['ID','Date','UserID','Qty'],
    [SHEETS.PAYMENTS]:    ['ID','Date','UserID','Amount'],
    [SHEETS.BUYS]:        ['ID','Date','Qty','TotalTK','UnitCost'],
    [SHEETS.QUOTAS]:      ['UserID','Quota'],
    [SHEETS.SESSIONS]:    ['Token','UserID','CreatedAt'],
    [SHEETS.DESTROY]:     ['ID','Date','Qty','Note'],
    [SHEETS.MISC]:        ['ID','Date','Amount','Note'],
    [SHEETS.PREDEFINED]:  ['UserID','Date','Qty'],
    [SHEETS.FIFO_RPT]:    ['Name','Consumed Qty','Pct %','Egg Cost','Destroy Share','Misc Share','Total Cost','Total Paid','Need to Pay'],
  };
  Object.entries(defs).forEach(([name, headers]) => {
    let sh = SS2.getSheetByName(name);
    if (!sh) sh = SS2.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1,1,1,headers.length).setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold');
    }
  });
  const userSh = SS2.getSheetByName(SHEETS.USERS);
  if (userSh.getLastRow() <= 1)
    userSh.appendRow(['U0','superadmin','admin123','Super Admin','superadmin',0,0]);
  Logger.log('✓ Setup complete v4');
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTH
// ═══════════════════════════════════════════════════════════════════════════
function handleLogin({ username, password }) {
  if (!username || !password) return { success: false, error: 'Username and password required' };
  const sh = SS2.getSheetByName(SHEETS.USERS);
  if (!sh) return { success: false, error: 'Users sheet not found. Run setupSheets() first.' };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [userId, uname, pw, name, role] = rows[i];
    if (String(uname).trim() === String(username).trim() && String(pw).trim() === String(password).trim()) {
      const token = Utilities.getUuid();
      SS2.getSheetByName(SHEETS.SESSIONS).appendRow([token, String(userId), new Date().toISOString()]);
      return { success: true, session: { userId: String(userId), username: String(uname).trim(), name: String(name), role: String(role), token } };
    }
  }
  return { success: false, error: 'Invalid username or password' };
}

function verifyToken(token) {
  if (!token) return null;
  const sh = SS2.getSheetByName(SHEETS.SESSIONS);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++)
    if (String(rows[i][0]) === String(token)) return getUserById(String(rows[i][1]));
  return null;
}

function getUserById(userId) {
  const rows = SS2.getSheetByName(SHEETS.USERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++)
    if (String(rows[i][0]) === userId)
      return { userId: String(rows[i][0]), username: String(rows[i][1]), name: String(rows[i][3]),
               role: String(rows[i][4]), target: Number(rows[i][5])||0, initBalance: Number(rows[i][6])||0 };
  return null;
}

function changePassword({ oldPassword, newPassword }, user) {
  if (!oldPassword || !newPassword) return { success: false, error: 'Both passwords required' };
  const sh = SS2.getSheetByName(SHEETS.USERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(user.userId)) {
      if (String(rows[i][2]).trim() !== String(oldPassword).trim())
        return { success: false, error: 'Current password incorrect' };
      sh.getRange(i+1, 3).setValue(String(newPassword));
      return { success: true };
    }
  }
  return { success: false, error: 'User not found' };
}

// ═══════════════════════════════════════════════════════════════════════════
// PARTICIPANTS
// ═══════════════════════════════════════════════════════════════════════════
function getParticipantUsers() { return getAllUsers().filter(u => u.role !== 'superadmin'); }

// ═══════════════════════════════════════════════════════════════════════════
// USER DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function getDashboard({ date }, authUser) {
  const uid    = authUser.userId;
  const today  = date || fmtDate(new Date());
  const quota  = getQuotaForUser(uid);
  const consumed    = getConsumptionForUserDate(uid, today); // -1 = not entered yet
  const predefined  = getPredefinedForUserDate(uid, today);
  const totalPaid   = getTotalPaidForUser(uid);

  const allConsRows = SS2.getSheetByName(SHEETS.CONSUMPTION).getDataRange().getValues().slice(1)
    .filter(r => String(r[2]) === uid);
  const totalConsumed = allConsRows.reduce((s, r) => s + Number(r[3]), 0);
  const recent = allConsRows.sort((a,b) => new Date(b[1])-new Date(a[1])).slice(0,14)
    .map(r => ({ date: fmtDate(r[1]), qty: Number(r[3]) }));

  const buys = SS2.getSheetByName(SHEETS.BUYS).getDataRange().getValues().slice(1);
  const avgUnit = buys.length ? buys.reduce((s,r)=>s+Number(r[3]),0) / buys.reduce((s,r)=>s+Number(r[2]),0) : 0;
  const estimatedCost = Math.round(totalConsumed * avgUnit);
  const balance = totalPaid - estimatedCost;

  const allPredefined = getPredefinedAllForUser(uid);

  return { success: true, data: { quota, consumed, predefined, totalPaid, totalConsumed, estimatedCost, balance, recent, allPredefined } };
}

function getConsumptionForUserDate(userId, date) {
  const rows = SS2.getSheetByName(SHEETS.CONSUMPTION).getDataRange().getValues().slice(1);
  for (const r of rows)
    if (String(r[2]) === userId && fmtDate(r[1]) === date) return Number(r[3]);
  return -1; // -1 = admin has NOT entered data yet
}

function getTotalPaidForUser(userId) {
  return SS2.getSheetByName(SHEETS.PAYMENTS).getDataRange().getValues().slice(1)
    .filter(r => String(r[2]) === userId).reduce((s,r) => s+Number(r[3]), 0);
}

// ═══════════════════════════════════════════════════════════════════════════
// ADMIN DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function getAdminDashboard(user) {
  requireAdmin(user);
  const users    = getParticipantUsers();
  const buysRows = SS2.getSheetByName(SHEETS.BUYS).getDataRange().getValues().slice(1);
  const consRows = SS2.getSheetByName(SHEETS.CONSUMPTION).getDataRange().getValues().slice(1);
  const destRows = SS2.getSheetByName(SHEETS.DESTROY).getDataRange().getValues().slice(1);
  const miscRows = SS2.getSheetByName(SHEETS.MISC).getDataRange().getValues().slice(1);
  const payRows  = SS2.getSheetByName(SHEETS.PAYMENTS).getDataRange().getValues().slice(1);

  const totalBought      = buysRows.reduce((s,r) => s+Number(r[2]), 0);
  const totalSpentOnEggs = buysRows.reduce((s,r) => s+Number(r[3]), 0);
  const totalDestroyed   = destRows.reduce((s,r) => s+Number(r[2]), 0);
  const totalMisc        = miscRows.reduce((s,r) => s+Number(r[2]), 0);
  const uids = new Set(users.map(u=>u.userId));
  const totalConsumed    = consRows.filter(r=>uids.has(String(r[2]))).reduce((s,r)=>s+Number(r[3]),0);
  const stockRemaining   = totalBought - totalConsumed - totalDestroyed;

  const fifo             = computeFifo(users);
  const totalCostFifo    = fifo.report.reduce((s,r) => s+r.totalCost, 0);
  const totalCollected   = fifo.report.reduce((s,r) => s+r.totalPaid, 0);
  const totalOutstanding = fifo.report.reduce((s,r) => s+Math.max(0,r.needToPay), 0);
  const allPayments      = payRows.reduce((s,r) => s+Number(r[3]), 0);
  const cashInHand       = allPayments - totalSpentOnEggs - totalMisc;

  const userMap = {};
  users.forEach(u => { userMap[u.userId] = u.name; });
  const recentActivity = consRows.filter(r=>userMap[String(r[2])])
    .sort((a,b)=>new Date(b[1])-new Date(a[1])).slice(0,10)
    .map(r=>({ date: fmtDate(r[1]), userName: userMap[String(r[2])], qty: Number(r[3]) }));

  return { success: true, data: {
    totalBought, totalConsumed, totalDestroyed, stockRemaining,
    totalSpentOnEggs, totalMisc, totalCostFifo, totalCollected, totalOutstanding, cashInHand,
    userBalances: fifo.report.map(r=>({ name:r.name, consumedQty:r.consumedQty, totalCost:r.totalCost, totalPaid:r.totalPaid, needToPay:r.needToPay })),
    recentActivity,
  }};
}

// ═══════════════════════════════════════════════════════════════════════════
// DAILY CONSUMPTION
// ═══════════════════════════════════════════════════════════════════════════
function getDailyConsumptionData({ date }, user) {
  requireAdmin(user);
  const users = getParticipantUsers();
  const rows  = SS2.getSheetByName(SHEETS.CONSUMPTION).getDataRange().getValues().slice(1);
  const existing = {};
  rows.filter(r => fmtDate(r[1]) === date).forEach(r => { existing[String(r[2])] = Number(r[3]); });
  const hasExisting = Object.keys(existing).length > 0;
  users.forEach(u => {
    u.quota      = getQuotaForUser(u.userId);
    u.predefined = getPredefinedForUserDate(u.userId, date);
  });
  return { success: true, users, existing, hasExisting };
}

function saveConsumption({ date, entries }, user) {
  requireAdmin(user);
  const sh   = SS2.getSheetByName(SHEETS.CONSUMPTION);
  const rows = sh.getDataRange().getValues();
  entries.forEach(({ userId, qty }) => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]) === userId && fmtDate(rows[i][1]) === date) {
        sh.getRange(i+1,4).setValue(Number(qty)); rows[i][3]=Number(qty); found=true; break;
      }
    }
    if (!found && Number(qty) > 0) { const nr=[genId(),date,userId,Number(qty)]; sh.appendRow(nr); rows.push(nr); }
  });
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════════════════
// PREDEFINED — per user per date [UserID, Date, Qty]
// ═══════════════════════════════════════════════════════════════════════════
function getPredefinedForUserDate(userId, date) {
  const sh = SS2.getSheetByName(SHEETS.PREDEFINED);
  if (!sh || sh.getLastRow() <= 1) return 0;
  const rows = sh.getDataRange().getValues().slice(1);
  for (const r of rows)
    if (String(r[0]) === String(userId) && fmtDate(r[1]) === String(date)) return Number(r[2])||0;
  return 0;
}

function getPredefinedAllForUser(userId) {
  const sh = SS2.getSheetByName(SHEETS.PREDEFINED);
  if (!sh || sh.getLastRow() <= 1) return [];
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(userId))
    .map(r => ({ date: fmtDate(r[1]), qty: Number(r[2])||0 }))
    .sort((a,b) => a.date.localeCompare(b.date));
}

function getPredefined({ date }, user) {
  if (user.role === 'user') {
    return { success: true, qty: date ? getPredefinedForUserDate(user.userId, date) : 0,
             all: getPredefinedAllForUser(user.userId) };
  }
  requireAdmin(user);
  const users = getParticipantUsers();
  users.forEach(u => { u.predefined = date ? getPredefinedForUserDate(u.userId, date) : 0; });
  return { success: true, users };
}

function setPredefined({ date, qty }, user) {
  if (!date) return { success: false, error: 'Date required' };
  const today = fmtDate(new Date());
  if (user.role === 'user') {
    if (date < today) return { success: false, error: 'Cannot set predefined for past dates' };
    // Block if admin already entered for this date
    const consumed = getConsumptionForUserDate(user.userId, date);
    if (consumed >= 0) return { success: false, error: 'Admin already entered consumption for this date' };
  }
  const uid = user.userId;
  const sh  = SS2.getSheetByName(SHEETS.PREDEFINED);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(uid) && fmtDate(rows[i][1]) === String(date)) {
      sh.getRange(i+1,3).setValue(Number(qty)||0);
      return { success: true };
    }
  }
  sh.appendRow([uid, date, Number(qty)||0]);
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════════════════
// PAYMENTS
// ═══════════════════════════════════════════════════════════════════════════
function getPayments(user) {
  requireAdmin(user);
  const userMap = {};
  getParticipantUsers().forEach(u => { userMap[u.userId] = u.name; });
  return { success: true, payments: SS2.getSheetByName(SHEETS.PAYMENTS).getDataRange().getValues().slice(1)
    .filter(r => userMap.hasOwnProperty(String(r[2])))
    .map(r => ({ id:r[0], date:fmtDate(r[1]), userId:String(r[2]), userName:userMap[String(r[2])], amount:Number(r[3]) }))
    .sort((a,b) => b.date.localeCompare(a.date)) };
}

function addPayment({ userId, date, amount }, user) {
  requireAdmin(user);
  SS2.getSheetByName(SHEETS.PAYMENTS).appendRow([genId(), date, userId, Number(amount)]);
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════════════════
// EGG BUYS / DESTROY / MISC
// ═══════════════════════════════════════════════════════════════════════════
function getEggBuys(user) {
  requireAdmin(user);
  return { success: true, buys: SS2.getSheetByName(SHEETS.BUYS).getDataRange().getValues().slice(1)
    .map(r=>({id:r[0],date:fmtDate(r[1]),qty:Number(r[2]),totalTK:Number(r[3]),unitCost:Number(r[4])}))
    .sort((a,b)=>b.date.localeCompare(a.date)) };
}
function addEggBuy({ date, qty, totalTK }, user) {
  requireAdmin(user);
  const q=Number(qty), t=Number(totalTK);
  SS2.getSheetByName(SHEETS.BUYS).appendRow([genId(), date, q, t, q>0 ? t/q : 0]);
  return { success: true };
}
function getEggDestroy(user) {
  requireAdmin(user);
  return { success: true, rows: SS2.getSheetByName(SHEETS.DESTROY).getDataRange().getValues().slice(1)
    .map(r=>({id:r[0],date:fmtDate(r[1]),qty:Number(r[2]),note:r[3]||''}))
    .sort((a,b)=>b.date.localeCompare(a.date)) };
}
function addEggDestroy({ date, qty, note }, user) {
  requireAdmin(user);
  SS2.getSheetByName(SHEETS.DESTROY).appendRow([genId(), date, Number(qty), note||'']);
  return { success: true };
}
function getMiscCost(user) {
  requireAdmin(user);
  return { success: true, rows: SS2.getSheetByName(SHEETS.MISC).getDataRange().getValues().slice(1)
    .map(r=>({id:r[0],date:fmtDate(r[1]),amount:Number(r[2]),note:r[3]||''}))
    .sort((a,b)=>b.date.localeCompare(a.date)) };
}
function addMiscCost({ date, amount, note }, user) {
  requireAdmin(user);
  SS2.getSheetByName(SHEETS.MISC).appendRow([genId(), date, Number(amount), note||'']);
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════════════════
// FIFO CORE
// ═══════════════════════════════════════════════════════════════════════════
function computeFifo(users) {
  const buys = SS2.getSheetByName(SHEETS.BUYS).getDataRange().getValues().slice(1)
    .map(r=>({date:new Date(r[1]),qty:Number(r[2]),unitCost:Number(r[4])})).sort((a,b)=>a.date-b.date);
  const allCons = SS2.getSheetByName(SHEETS.CONSUMPTION).getDataRange().getValues().slice(1)
    .map(r=>({date:new Date(r[1]),userId:String(r[2]),qty:Number(r[3])})).sort((a,b)=>a.date-b.date);
  const destroyRows = SS2.getSheetByName(SHEETS.DESTROY).getDataRange().getValues().slice(1)
    .map(r=>({date:new Date(r[1]),qty:Number(r[2])})).sort((a,b)=>a.date-b.date);
  const miscRows = SS2.getSheetByName(SHEETS.MISC).getDataRange().getValues().slice(1).map(r=>({amount:Number(r[2])}));

  const queue = buys.map(b=>({qty:b.qty,unitCost:b.unitCost}));
  function dequeue(qty) {
    let rem=qty, cost=0;
    while(rem>0 && queue.length>0) { const b=queue[0]; const t=Math.min(rem,b.qty); cost+=t*b.unitCost; b.qty-=t; rem-=t; if(b.qty<=0) queue.shift(); }
    return cost;
  }
  const userCost={};
  users.forEach(u=>{userCost[u.userId]=0;});
  const timeline=[...allCons.map(c=>({...c,type:'cons'})),...destroyRows.map(d=>({...d,type:'destroy',userId:'__destroy__'}))].sort((a,b)=>a.date-b.date);
  let destroyCost=0;
  for(const ev of timeline) {
    if(ev.type==='cons') { if(!userCost.hasOwnProperty(ev.userId)) continue; userCost[ev.userId]+=dequeue(ev.qty); }
    else { destroyCost+=dequeue(ev.qty); }
  }
  const userQty={};
  users.forEach(u=>{userQty[u.userId]=0;});
  allCons.forEach(c=>{if(userQty.hasOwnProperty(c.userId)) userQty[c.userId]+=c.qty;});
  const totalQtyConsumed=Object.values(userQty).reduce((s,v)=>s+v,0);
  const totalMiscCost=miscRows.reduce((s,r)=>s+r.amount,0);
  const report=users.map(u=>{
    const qty=userQty[u.userId]||0;
    const pct=totalQtyConsumed>0 ? qty/totalQtyConsumed : 0;
    const eggCost=Math.round(userCost[u.userId]||0);
    const shareDest=Math.round(destroyCost*pct);
    const shareMisc=Math.round(totalMiscCost*pct);
    const totalCost=eggCost+shareDest+shareMisc;
    const totalPaid=getTotalPaidForUser(u.userId);
    const needToPay=(u.target||0)+totalCost-totalPaid-(u.initBalance||0);
    return {name:u.name,consumedQty:qty,pct:Math.round(pct*100),eggCost,shareDest,shareMisc,totalCost,totalPaid,target:u.target||0,initBalance:u.initBalance||0,needToPay:Math.round(needToPay)};
  });
  return {report, summary:{totalQtyConsumed, destroyCost:Math.round(destroyCost), totalMiscCost:Math.round(totalMiscCost)}};
}

function getFifoReport(user) {
  requireAdmin(user);
  const result = computeFifo(getParticipantUsers());
  return { success: true, ...result };
}

function refreshFifoSheet(user) {
  if (user) requireAdmin(user);
  const users = getParticipantUsers();
  if (!users || users.length === 0) return { success: true, message: 'No users' };
  const result = computeFifo(users);
  const sh = SS2.getSheetByName(SHEETS.FIFO_RPT);
  if (!sh) return { success: false, error: 'FIFO Report sheet missing. Run setupSheets().' };
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow-1);
  sh.appendRow(['Generated:', new Date().toLocaleString(), 'Total Consumed:', result.summary.totalQtyConsumed, 'Destroy Cost:', result.summary.destroyCost, 'Misc Cost:', result.summary.totalMiscCost, '']);
  sh.appendRow(['','','','','','','','','']);
  const hdrRow=['Name','Consumed Qty','Pct %','Egg Cost','Destroy Share','Misc Share','Total Cost','Total Paid','Need to Pay'];
  sh.appendRow(hdrRow);
  sh.getRange(sh.getLastRow(),1,1,hdrRow.length).setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold');
  result.report.forEach(r => {
    sh.appendRow([r.name,r.consumedQty,r.pct+'%',r.eggCost,r.shareDest,r.shareMisc,r.totalCost,r.totalPaid,r.needToPay]);
    sh.getRange(sh.getLastRow(),9).setFontColor(r.needToPay>0?'#b91c1c':'#166534').setFontWeight('bold');
  });
  const ds=sh.getLastRow()-result.report.length+1, de=sh.getLastRow();
  sh.appendRow(['TOTAL','','','=SUM(D'+ds+':D'+de+')','=SUM(E'+ds+':E'+de+')','=SUM(F'+ds+':F'+de+')','=SUM(G'+ds+':G'+de+')','=SUM(H'+ds+':H'+de+')','=SUM(I'+ds+':I'+de+')']);
  sh.getRange(sh.getLastRow(),1,1,9).setBackground('#F3F4F6').setFontWeight('bold');
  for(let c=1;c<=9;c++) sh.autoResizeColumn(c);
  return { success: true, message: 'FIFO Report sheet refreshed' };
}

// onOpen — only runs when there is actual egg buy data to avoid empty appendRow crash
function onOpen() {
  try {
    const buysSh = SS2.getSheetByName(SHEETS.BUYS);
    if (buysSh && buysSh.getLastRow() > 1) refreshFifoSheet(null);
  } catch(e) { Logger.log('onOpen skipped: ' + e.message); }
}

// ═══════════════════════════════════════════════════════════════════════════
// USERS CRUD
// ═══════════════════════════════════════════════════════════════════════════
function getUsers(user) { requireSuperadmin(user); return { success: true, users: getAllUsers() }; }

function getAllUsers() {
  return SS2.getSheetByName(SHEETS.USERS).getDataRange().getValues().slice(1)
    .map(r=>({userId:String(r[0]),username:String(r[1]),name:String(r[3]),role:String(r[4]),target:Number(r[5])||0,initBalance:Number(r[6])||0}));
}

function createUser(data, user) {
  requireSuperadmin(user);
  if (!data.username || !data.password) return { success: false, error: 'Username and password required' };
  if (getAllUsers().find(u=>u.username===data.username)) return { success: false, error: 'Username already exists' };
  const userId='U'+Date.now();
  SS2.getSheetByName(SHEETS.USERS).appendRow([userId,data.username,data.password,data.name||data.username,data.role||'user',Number(data.target)||0,Number(data.initBalance)||0]);
  SS2.getSheetByName(SHEETS.QUOTAS).appendRow([userId,0]);
  return { success: true, userId };
}

function updateUser(data, user) {
  requireSuperadmin(user);
  const sh=SS2.getSheetByName(SHEETS.USERS), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) {
    if(String(rows[i][0])===String(data.userId)) {
      sh.getRange(i+1,2).setValue(data.username);
      if(data.password) sh.getRange(i+1,3).setValue(data.password);
      sh.getRange(i+1,4).setValue(data.name);
      sh.getRange(i+1,5).setValue(data.role);
      sh.getRange(i+1,6).setValue(Number(data.target)||0);
      sh.getRange(i+1,7).setValue(Number(data.initBalance)||0);
      return { success: true };
    }
  }
  return { success: false, error: 'User not found' };
}

function deleteUserFn({ userId }, user) {
  requireSuperadmin(user);
  if(String(userId)===String(user.userId)) return { success: false, error: 'Cannot delete yourself' };
  const sh=SS2.getSheetByName(SHEETS.USERS), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++)
    if(String(rows[i][0])===String(userId)) { sh.deleteRow(i+1); return { success: true }; }
  return { success: false, error: 'User not found' };
}

// ═══════════════════════════════════════════════════════════════════════════
// QUOTAS
// ═══════════════════════════════════════════════════════════════════════════
function getQuotaForUser(userId) {
  const rows=SS2.getSheetByName(SHEETS.QUOTAS).getDataRange().getValues().slice(1);
  for(const r of rows) if(String(r[0])===String(userId)) return Number(r[1]);
  return 0;
}
function getQuotas(user) {
  requireAdmin(user);
  const users=getParticipantUsers();
  users.forEach(u=>{u.quota=getQuotaForUser(u.userId);});
  return { success: true, users };
}
function saveQuotas({ quotas }, user) {
  requireAdmin(user);
  quotas.forEach(({userId,quota})=>setQuotaForUser(userId,quota));
  return { success: true };
}
function setUserQuota({ userId, qty }, user) {
  if(user.role==='user') {
    if(user.userId!==userId) return { success: false, error: "Cannot edit others' quota" };
    if(new Date().getHours()>=9) return { success: false, error: 'Quota locked after 9:00 AM' };
  }
  setQuotaForUser(userId,qty);
  return { success: true };
}
function setQuotaForUser(userId, quota) {
  const sh=SS2.getSheetByName(SHEETS.QUOTAS), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++)
    if(String(rows[i][0])===String(userId)) { sh.getRange(i+1,2).setValue(Number(quota)); return; }
  sh.appendRow([userId,Number(quota)]);
}

// ═══════════════════════════════════════════════════════════════════════════
// GENERIC DELETE / UTILS
// ═══════════════════════════════════════════════════════════════════════════
function deleteById(sheetName, id, user) {
  requireAdmin(user);
  const sh=SS2.getSheetByName(sheetName), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++)
    if(String(rows[i][0])===String(id)) { sh.deleteRow(i+1); return { success: true }; }
  return { success: false, error: 'Record not found' };
}

function genId() { return Utilities.getUuid().split('-')[0].toUpperCase(); }

function fmtDate(d) {
  if(!d) return '';
  const dt=d instanceof Date ? d : new Date(d);
  if(isNaN(dt)) return String(d);
  return dt.toISOString().split('T')[0];
}

function requireAdmin(user) {
  if(!['admin','superadmin'].includes(user.role)) throw new Error('Permission denied — admin required');
}
function requireSuperadmin(user) {
  if(user.role!=='superadmin') throw new Error('Permission denied — superadmin required');
}
