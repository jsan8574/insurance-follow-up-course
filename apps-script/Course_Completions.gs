/* Deploy as a web app, execute as the owner, and allow Anyone.
   This endpoint appends completion summaries to one tab. It does not serve sheet contents. */
const COURSE_ORIGIN = 'https://rcm-back-end-lab-prep.jsan8574.chatgpt.site';
const COMPLETION_SHEET_ID = '1RvywtwqcCZuXGmgGvBAvKE8hLezLDzbeY9pJdHrz8BM';
const COMPLETION_TAB = 'Course Completions';
const TIME_HEADERS = ['Active Seconds', 'RCM Back End Lab Readiness'];
const UUID = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;

function doPost(e) {
  let nonce = '', ok = false;
  try {
    const raw = e.parameter.payload;
    if (typeof raw !== 'string' || raw.length > 1200) throw Error('Invalid payload');
    const item = JSON.parse(raw);
    if (!UUID.test(item.nonce)) throw Error('Invalid submission nonce');
    nonce = item.nonce;
    const completed = new Date(item.completedAt);
    if (!UUID.test(item.id) || typeof item.name !== 'string' ||
        !/^[\p{L}][\p{L}\p{M} .’'\-]{0,79}$/u.test(item.name) ||
        typeof item.completedAt !== 'string' || !Number.isFinite(completed.getTime()) ||
        completed.getTime() > Date.now() + 300000 ||
        !Number.isInteger(item.score) || item.score < 16 || item.score > 20 ||
        !Number.isInteger(item.attempts) || item.attempts < 1 || item.attempts > 999 ||
        item.drills !== 14 || item.cases !== 2) throw Error('Invalid completion record');
    const activeSeconds = item.activeSeconds ?? '';
    const readiness = item.score >= 18 ? 'Ready' : 'Ready With Review';
    if ((activeSeconds !== '' && (!Number.isSafeInteger(activeSeconds) || activeSeconds < 0)) ||
        (item.readiness !== undefined && item.readiness !== readiness)) throw Error('Invalid time or readiness');
    const lock = LockService.getScriptLock();
    lock.waitLock(12000);
    try {
      const tab = SpreadsheetApp.openById(COMPLETION_SHEET_ID).getSheetByName(COMPLETION_TAB);
      if (!tab) throw Error('Missing completion tab');
      const header = tab.getRange(1,11,1,TIME_HEADERS.length);
      const labels = header.getValues()[0];
      if (labels.some((label,i)=>label && label !== TIME_HEADERS[i])) throw Error('Unexpected completion columns');
      if (labels.some(label=>!label)) header.setValues([TIME_HEADERS]);
      const last = tab.getLastRow();
      const existing = last > 1 && tab.getRange(2,1,last-1,1).createTextFinder(item.id).matchEntireCell(true).findNext();
      if (existing) {
        const row = tab.getRange(existing.getRow(),1,1,12).getValues()[0];
        if (row[3] !== item.name || row[4] !== item.score || row[6] !== item.attempts ||
            new Date(row[2]).getTime() !== completed.getTime() ||
            (row[10] !== '' && activeSeconds !== '' && row[10] !== activeSeconds) ||
            (row[11] && row[11] !== readiness)) throw Error('Conflicting completion record');
        if (row[10] === '' || !row[11]){
          tab.getRange(existing.getRow(),11,1,2).setValues([[row[10] === '' ? activeSeconds : row[10], readiness]]);
          SpreadsheetApp.flush();
        }
      } else {
        tab.appendRow([item.id,new Date(),completed,item.name,item.score,item.score/20,
                       item.attempts,item.drills,item.cases,'Completed',activeSeconds,readiness]);
        SpreadsheetApp.flush();
      }
      ok = true;
    } finally {lock.releaseLock();}
  } catch (error) {
    console.error('Course completion could not be saved: ' + error.message);
  }
  const data = JSON.stringify({type:'rcm-prep-completion',nonce,ok});
  const origin = JSON.stringify(COURSE_ORIGIN);
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>' +
    'var w=window;while(w!==w.parent){w=w.parent;w.postMessage('+data+','+origin+');}</script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
