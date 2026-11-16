const ID='1EFmRYk7dq7t3tWWQyPgw9n0gMot7Iar9_W4JBhX2pNg';
const COLS=["serial","date","customer","phone","institution","caseNo","service","vehicle","driver","actingDriver","technicianName","technicianAmount","actDriverAmount","pickup","drop","distance","basic","waitingCharges","freezerOption","freezer","medicalSale","other","total","advance","balance","received","due","status","petrolExpense","otherExpense","referenceAmount","referenceName","referenceContact","billEnteredBy","billClosedBy"];
const MON_=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const ss_=()=>SpreadsheetApp.openById(ID);
const fmt_=x=>x instanceof Date?Utilities.formatDate(x,'Asia/Kolkata','yyyy-MM-dd'):x;
const json_=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const row_=(h,r)=>{const o={};h.forEach((k,i)=>o[k]=fmt_(r[i]));return o};

function doGet(e){
  const p=(e&&e.parameter)||{}, ss=ss_(), sh=ss.getSheets()[0];
  if(p.action==='reserve'){
    const lock=LockService.getScriptLock(); lock.waitLock(20000);
    try{
      const d=p.date||Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd'), pre='SVIMS/'+MON_[+d.slice(5,7)-1]+'/', rows=sh.getDataRange().getValues(); let m=0;
      for(let i=1;i<rows.length;i++){const q=String(rows[i][0]); if(q.indexOf(pre)===0) m=Math.max(m,parseInt(q.slice(pre.length))||0)}
      const serial=pre+String(m+1).padStart(2,'0'), r=sh.getLastRow()+1;
      sh.getRange(r,1,1,3).setNumberFormat('@'); sh.getRange(r,1).setValue(serial); sh.getRange(r,2).setValue(d); sh.getRange(r,28).setValue('Reserved');
      return json_({result:'success',serial});
    } finally { lock.releaseLock(); }
  }
  const v=sh.getDataRange().getValues(), h=v.shift();
  if(p.action==='has'){const i=v.findIndex(r=>String(r[0])===String(p.serial)); return json_({found:i>=0&&String(v[i][27])!=='Reserved'})}
  if(p.action==='phone'){const ph=String(p.phone); return json_({result:'success',trips:v.filter(r=>String(r[3])===ph&&String(r[27])!=='Reserved').slice(-30).map(r=>row_(h,r))})}
  if(p.action==='audit'){
    const as=ss.getSheetByName('Audit'); let a=[];
    if(as&&as.getLastRow()>1){const n=Math.min(300,as.getLastRow()-1); a=as.getRange(as.getLastRow()-n+1,1,n,5).getValues().reverse().map(r=>({at:r[0],by:r[1],action:r[2],ref:r[3],detail:r[4]}))}
    return json_({result:'success',audit:a});
  }
  const since=p.since||'', until=p.until||'', inR=d=>(!since||d>=since)&&(!until||d<=until);
  const trips=v.filter(r=>r[0]!==''&&String(r[27])!=='Reserved').map(r=>row_(h,r))
    .filter(o=>{const st=String(o.status); return inR(String(o.date))||(st!=='Paid'&&st!=='Cancelled')});
  let config=null; const cs=ss.getSheetByName('Config');
  try{ if(cs) config=JSON.parse(cs.getRange('A1').getValue()); }catch(x){}
  let expenses=[]; const es=ss.getSheetByName('Expenses');
  if(es&&es.getLastRow()>1) expenses=es.getRange(2,1,es.getLastRow()-1,11).getValues()
    .map(r=>({voucher:r[0],date:fmt_(r[1]),category:r[2],amount:r[3],note:r[4],trip:r[5],by:r[6],vehicle:r[7],odo:r[8],litres:r[9],full:r[10]}))
    .filter(x=>inR(String(x.date)));
  return json_({result:'success',trips,config,expenses});
}

function doPost(e){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const t=JSON.parse(e.parameter.postData), ss=ss_();
    if(t.__config){
      const cs=ss.getSheetByName('Config')||ss.insertSheet('Config');
      cs.getRange('A1').setNumberFormat('@').setValue(JSON.stringify(t.__config));
    } else if(t.__audit){
      const as=ss.getSheetByName('Audit')||ss.insertSheet('Audit');
      if(as.getLastRow()===0) as.appendRow(['at','by','action','ref','detail']);
      const a=t.__audit, safe=x=>String(x==null?'':x).replace(/^([=+\-@])/,"'$1");
      as.appendRow([Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd HH:mm:ss'),safe(a.by),safe(a.action),safe(a.ref),safe(a.detail)]);
    } else if(t.__expense||t.__expDelete){
      const es=ss.getSheetByName('Expenses')||ss.insertSheet('Expenses');
      es.getRange(1,1,1,11).setValues([['voucher','date','category','amount','note','trip','by','vehicle','odo','litres','full']]);
      const rows=es.getDataRange().getValues();
      if(t.__expDelete){
        for(let i=rows.length-1;i>=1;i--) if(String(rows[i][0])===String(t.__expDelete)) es.deleteRow(i+1);
      } else {
        const v=t.__expense; let r=-1;
        for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(v.voucher)){r=i+1;break}
        if(r<0) r=es.getLastRow()+1;
        es.getRange(r,1,1,2).setNumberFormat('@');
        es.getRange(r,1,1,11).setValues([[v.voucher,v.date,v.category,v.amount,v.note,v.trip,v.by,v.vehicle||'',v.odo||'',v.litres||'',v.full||'']]);
      }
    } else {
      const sh=ss.getSheets()[0], rows=sh.getDataRange().getValues(); let r=-1;
      for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(t.serial)){r=i+1;break}
      if(r<0) r=sh.getLastRow()+1;
      [2,4,6].forEach(c=>sh.getRange(r,c).setNumberFormat('@'));
      sh.getRange(r,1,1,COLS.length).setValues([COLS.map(c=>t[c]==null?'':t[c])]);
    }
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}

// Nightly backup: run setupBackup() ONCE from the editor (approve the Drive permission)
function setupBackup(){ ScriptApp.newTrigger('nightlyBackup').timeBased().everyDays(1).atHour(2).create(); }
function nightlyBackup(){
  const d=Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd');
  DriveApp.getFileById(ID).makeCopy('SVIMS billing backup '+d);
  const it=DriveApp.searchFiles('title contains "SVIMS billing backup " and trashed=false'), all=[];
  while(it.hasNext()){const f=it.next(); all.push({f,n:f.getName()})}
  all.sort((a,b)=>b.n.localeCompare(a.n)).slice(30).forEach(x=>x.f.setTrashed(true));
}
