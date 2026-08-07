/**
 * Multi-Event Ticketing System (Google Apps Script backend)
 * New standalone deployment intended for:
 * https://greenprofessionals.github.io/multi-event-tickets/
 *
 * This is deliberately separate from the existing working single-event system.
 * Run setupMultiEventSystem() once in a NEW Google Sheet, then deploy this
 * script as a Web App (Execute as: Me, Access: Anyone).
 */

const BASE_SITE_URL = 'https://greenprofessionals.github.io/multi-event-tickets';
const BASE_CLAIM_URL = BASE_SITE_URL + '/claim.html';
const BASE_VOUCHER_URL = BASE_SITE_URL + '/v.html';
const BASE_GATE_URL = BASE_SITE_URL + '/gate.html';
const BASE_ADMIN_URL = BASE_SITE_URL + '/admin.html';

const SHEETS = {
  EVENTS: 'Events', TIERS: 'Tiers', GROUPS: 'Groups', VOUCHERS: 'Vouchers',
  CLAIMS: 'Claims', CHECKINS: 'CheckIns', PAYMENTS: 'Payments', AUDIT: 'AuditLog',
  COUNTERS: 'Counters', FORM: 'EventConfigResponses'
};

// Keep these backend-only. Add/replace administrators before production use.
const ADMIN_DIRECTORY = {
  'kulane2026': 'Kula Sillah — New England Chapter (NEC)',
  'fombany2026': 'Fomba Kassoh — Administrator'
};

const EVENT_HEADERS = [
  'EventID','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime',
  'VenueName','VenueAddress','ContactPhone','WebsiteURL','DressCode','PrimaryColor',
  'GoldColor','LogoFileId','BackgroundFileId','GroupLabel','UseGroups','SerialPrefix',
  'CurrencySymbol','FooterLegalText','Capacity','Status','CreatedAt','UpdatedAt'
];
const TIER_HEADERS = ['EventID','TierKey','Label','Price','Capacity','Active'];
const GROUP_HEADERS = ['EventID','GroupName','Active'];
const VOUCHER_HEADERS = ['Timestamp','EventID','BatchID','VoucherToken','TierKey','SuggestedGroup','PrefillName','PrefillPhone','Claimed','Serial','Dispatched','RecipientEmail','RecipientPhone','SentAt','IssuedBy'];
const CLAIM_HEADERS = ['Timestamp','EventID','Serial','CheckInToken','Name','Email','Phone','GroupName','TierKey','Source','VoucherToken','Status','AmountDue','AmountPaid','PaymentStatus','PaymentMethod','PaymentNote'];
const CHECKIN_HEADERS = ['Timestamp','EventID','Serial','Name','GroupName','TierKey','Phone','PaymentStatus','PaymentMethod','AmountPaid','CheckedInBy','GateNote'];
const PAYMENT_HEADERS = ['Timestamp','EventID','Serial','Amount','Method','Status','Note','RecordedBy'];
const AUDIT_HEADERS = ['Timestamp','EventID','Action','EntityType','EntityID','Admin','Details'];
const COUNTER_HEADERS = ['EventID','CurrentNumber'];

function setupMultiEventSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.EVENTS, EVENT_HEADERS);
  ensureSheet_(ss, SHEETS.TIERS, TIER_HEADERS);
  ensureSheet_(ss, SHEETS.GROUPS, GROUP_HEADERS);
  ensureSheet_(ss, SHEETS.VOUCHERS, VOUCHER_HEADERS);
  ensureSheet_(ss, SHEETS.CLAIMS, CLAIM_HEADERS);
  ensureSheet_(ss, SHEETS.CHECKINS, CHECKIN_HEADERS);
  ensureSheet_(ss, SHEETS.PAYMENTS, PAYMENT_HEADERS);
  ensureSheet_(ss, SHEETS.AUDIT, AUDIT_HEADERS);
  ensureSheet_(ss, SHEETS.COUNTERS, COUNTER_HEADERS);

  // Seed a sample only when no events exist. It can be edited or deleted.
  const events = ss.getSheetByName(SHEETS.EVENTS);
  if (events.getLastRow() === 1) {
    const now = new Date();
    events.appendRow([
      'DEMO2026','SLPP North America','Demo Chapter','Executive Inauguration Celebration',
      'A reusable multi-event ticketing system','Saturday, September 5, 2026','6:00 PM',
      'Event Venue','123 Main Street','',BASE_SITE_URL,'Formal / Traditional','#0b3d24',
      '#c9a24b','','','Chapter','true','EV-','$','Ticket must be presented for admission · No refunds or replacements',
      500,'Draft',now,now
    ]);
    const tiers = ss.getSheetByName(SHEETS.TIERS);
    tiers.appendRow(['DEMO2026','single','Single',100,300,true]);
    tiers.appendRow(['DEMO2026','patron','Single Patron',200,120,true]);
    tiers.appendRow(['DEMO2026','double','Double Patron',300,80,true]);
    ss.getSheetByName(SHEETS.COUNTERS).appendRow(['DEMO2026',0]);
  }
  SpreadsheetApp.getUi().alert('Multi-event ticketing sheets are ready. Next run createEventConfigForm(), then deploy this script as a Web App.');
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    const row = sh.getRange(1,1,1,headers.length).getValues()[0];
    if (headers.some((h,i) => row[i] !== h)) sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
  return sh;
}

/** Creates a NEW standalone Google Form for event text/tier/group configuration.
 * Logo/background uploads are intentionally handled by config.html because Apps
 * Script FormApp cannot reliably create file-upload questions programmatically.
 */
function createEventConfigForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheets().map(s => s.getName());
  const form = FormApp.create('Multi-Event Ticket Configuration');
  form.setDescription('Create or update one event. Re-submit the same EventID to update its live configuration. Logo/background are uploaded separately from config.html.');
  const text = (t,h) => form.addTextItem().setTitle(t).setHelpText(h || '');
  const para = (t,h) => form.addParagraphTextItem().setTitle(t).setHelpText(h || '');
  text('EventID','Required unique ID, e.g. NY2026. Use the exact same ID to update an event.');
  text('OrgName','e.g. SLPP North America');
  text('ChapterName','e.g. New York Chapter');
  text('EventTitle','e.g. Executive Inauguration Celebration');
  para('Tagline','Optional subtitle');
  text('EventDate','e.g. Saturday, September 5, 2026');
  text('EventTime','e.g. 6:00 PM');
  text('VenueName','Venue name');
  text('VenueAddress','Full address');
  text('ContactPhone','Public contact number');
  text('WebsiteURL','Public event/chapter website');
  text('DressCode','Optional');
  text('PrimaryColor','Hex, e.g. #0b3d24');
  text('GoldColor','Hex, e.g. #c9a24b');
  text('SerialPrefix','e.g. NY-');
  text('CurrencySymbol','e.g. $');
  text('FooterLegalText','Small-print ticket footer');
  text('Capacity','Whole-event capacity');
  form.addMultipleChoiceItem().setTitle('UseGroups').setChoiceValues(['true','false']);
  text('GroupLabel','e.g. Chapter, Team, Table');
  para('GroupsList','One group name per line');
  for (let i=1;i<=5;i++) {
    text(`Tier${i}Name`,`Optional tier ${i} name`);
    text(`Tier${i}Price`,`Numbers only`);
    text(`Tier${i}Capacity`,`Optional tier capacity`);
  }
  form.addMultipleChoiceItem().setTitle('Status').setChoiceValues(['Draft','Active','Closed','Archived']);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  const newSheets = ss.getSheets().filter(s => !existing.includes(s.getName()));
  if (newSheets.length) newSheets[0].setName(SHEETS.FORM);
  SpreadsheetApp.getUi().alert('Configuration form created', 'Editor:\n' + form.getEditUrl() + '\n\nPublished:\n' + form.getPublishedUrl() + '\n\nLogo/background uploads are on config.html.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function safeBody_(e) { try { return JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (_) { return {}; } }
function adminName_(pass) { return ADMIN_DIRECTORY[(pass || '').toString().trim()] || ''; }
function requireAdmin_(body) { const name = adminName_(body.adminPasscode); return name ? {ok:true,name} : {ok:false,error:'Incorrect admin passcode.'}; }
function token_(bytes) { return Utilities.getUuid().replace(/-/g,'') + (bytes > 16 ? Utilities.getUuid().replace(/-/g,'') : ''); }
function normalizeId_(s) { return (s || '').toString().trim().toUpperCase().replace(/[^A-Z0-9_-]/g,''); }
function digitsOnly_(s) { return (s || '').toString().replace(/\D/g,''); }
function bool_(v) { return String(v).toLowerCase() === 'true' || v === true || v === 1; }
function money_(v) { const n = Number(v); return isFinite(n) ? n : 0; }

function rowsAsObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map((r,idx) => { const o = {_row:idx+2}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
}
function objectToRow_(headers,obj) { return headers.map(h => obj[h] === undefined ? '' : obj[h]); }
function findRowBy_(sheet, field, value) { return rowsAsObjects_(sheet).find(r => String(r[field]) === String(value)) || null; }
function updateObjectRow_(sheet, headers, rowNum, obj) { sheet.getRange(rowNum,1,1,headers.length).setValues([objectToRow_(headers,obj)]); }

function latestFormForEvent_(eventId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FORM);
  if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = headers.indexOf('EventID');
  if (idx < 0) return null;
  for (let r=data.length-1;r>=1;r--) {
    if (normalizeId_(data[r][idx]) === eventId) {
      const o={}; headers.forEach((h,i)=>o[h]=data[r][i]); return o;
    }
  }
  return null;
}

function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'America/New_York','EEEE, MMMM d, yyyy');
  const s=(v||'').toString().trim(); if (!s) return '';
  const d=new Date(s); if (!isNaN(d) && /(GMT|T\d\d:)/.test(s)) return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/New_York','EEEE, MMMM d, yyyy');
  return s;
}
function normalizeTime_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'America/New_York','h:mm a');
  const s=(v||'').toString().trim(); if (!s) return '';
  if (/^\d{1,2}:\d{2}(\s*[AP]M)?$/i.test(s)) {
    if (/[AP]M/i.test(s)) return s.toUpperCase();
    const p=s.split(':'); let h=Number(p[0]),m=p[1]; const ap=h>=12?'PM':'AM'; h=h%12||12; return `${h}:${m} ${ap}`;
  }
  const d=new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/New_York','h:mm a');
  return s;
}

function getEvent_(eventId) {
  eventId = normalizeId_(eventId);
  if (!eventId) return null;
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(SHEETS.EVENTS);
  const base=findRowBy_(sh,'EventID',eventId);
  if (!base) return null;
  const form=latestFormForEvent_(eventId);
  const e=Object.assign({},base);
  if (form) {
    ['OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','WebsiteURL','DressCode','PrimaryColor','GoldColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status'].forEach(k=>{
      if (form[k] !== undefined && String(form[k]).trim() !== '') e[k]=form[k];
    });
  }
  e.EventDate=normalizeDate_(e.EventDate);
  e.EventTime=normalizeTime_(e.EventTime);
  e.Capacity=Number(e.Capacity)||0;
  e.UseGroups=String(e.UseGroups || 'true');
  if (e.LogoFileId) e.LogoURL=driveFileToDataUrl_(e.LogoFileId);
  else e.LogoURL='';
  if (e.BackgroundFileId) e.BackgroundURL=driveFileToDataUrl_(e.BackgroundFileId);
  else e.BackgroundURL='';
  return e;
}
function publicEvent_(e) {
  if (!e) return null;
  const out={}; ['EventID','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','WebsiteURL','DressCode','PrimaryColor','GoldColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status','LogoURL','BackgroundURL'].forEach(k=>out[k]=e[k]);
  return out;
}
function getTiers_(eventId) {
  eventId=normalizeId_(eventId);
  const form=latestFormForEvent_(eventId);
  if (form) {
    const obj={};
    for (let i=1;i<=5;i++) {
      const label=(form[`Tier${i}Name`]||'').toString().trim(); if (!label) continue;
      const key=label.toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,30)||`tier${i}`;
      obj[key]={label,price:money_(form[`Tier${i}Price`]),capacity:Number(form[`Tier${i}Capacity`])||0};
    }
    if (Object.keys(obj).length) return obj;
  }
  const rows=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TIERS)).filter(r=>normalizeId_(r.EventID)===eventId && String(r.Active).toLowerCase()!=='false');
  const out={}; rows.forEach(r=>out[r.TierKey]={label:r.Label,price:money_(r.Price),capacity:Number(r.Capacity)||0}); return out;
}
function getGroups_(eventId) {
  const form=latestFormForEvent_(normalizeId_(eventId));
  if (form && String(form.GroupsList||'').trim()) return String(form.GroupsList).split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  return rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.GROUPS)).filter(r=>normalizeId_(r.EventID)===normalizeId_(eventId) && String(r.Active).toLowerCase()!=='false').map(r=>String(r.GroupName));
}
function listEvents_(includeClosed) {
  const rows=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVENTS));
  return rows.map(r=>getEvent_(r.EventID)).filter(Boolean).filter(e=>includeClosed || !['Archived'].includes(String(e.Status))).map(e=>({EventID:e.EventID,OrgName:e.OrgName,ChapterName:e.ChapterName,EventTitle:e.EventTitle,EventDate:e.EventDate,VenueName:e.VenueName,Status:e.Status,Capacity:e.Capacity,PrimaryColor:e.PrimaryColor,GoldColor:e.GoldColor,LogoURL:e.LogoURL}));
}

function driveFileToDataUrl_(fileId) {
  try { const f=DriveApp.getFileById(String(fileId)); const b=f.getBlob(); return `data:${b.getContentType()};base64,${Utilities.base64Encode(b.getBytes())}`; } catch (_) { return ''; }
}
function uploadBrandAsset_(body, admin) {
  const eventId=normalizeId_(body.eventId), e=getEvent_(eventId); if (!e) return {ok:false,error:'Event not found.'};
  const kind=body.assetType==='background'?'background':'logo';
  const data=String(body.dataUrl||''); const m=data.match(/^data:(image\/(?:png|jpeg|jpg|webp));base64,(.+)$/i); if (!m) return {ok:false,error:'Upload a PNG, JPG, or WebP image.'};
  const bytes=Utilities.base64Decode(m[2]); if (bytes.length>3*1024*1024) return {ok:false,error:'Image must be 3 MB or smaller.'};
  const ext=m[1].includes('png')?'png':m[1].includes('webp')?'webp':'jpg';
  const blob=Utilities.newBlob(bytes,m[1],`${eventId}-${kind}.${ext}`);
  const folder=getOrCreateAssetFolder_(); const file=folder.createFile(blob);
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVENTS); const row=findRowBy_(sh,'EventID',eventId); const obj=Object.assign({},row); obj[kind==='logo'?'LogoFileId':'BackgroundFileId']=file.getId(); obj.UpdatedAt=new Date(); updateObjectRow_(sh,EVENT_HEADERS,row._row,obj);
  audit_(eventId,'UPLOAD_BRAND_ASSET','Event',eventId,admin,kind);
  return {ok:true,assetType:kind};
}
function getOrCreateAssetFolder_() {
  const name='Multi-Event Ticket Assets'; const it=DriveApp.getFoldersByName(name); return it.hasNext()?it.next():DriveApp.createFolder(name);
}

function allocatedCounts_(eventId) {
  const claims=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS)).filter(r=>normalizeId_(r.EventID)===eventId && String(r.Status)!=='Revoked');
  const vouchers=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VOUCHERS)).filter(r=>normalizeId_(r.EventID)===eventId && !bool_(r.Claimed));
  const byTier={}; claims.forEach(r=>byTier[r.TierKey]=(byTier[r.TierKey]||0)+1); vouchers.forEach(r=>{if(r.TierKey)byTier[r.TierKey]=(byTier[r.TierKey]||0)+1;});
  return {total:claims.length+vouchers.length,byTier};
}
function capacityCheck_(eventId, additions) {
  const e=getEvent_(eventId); if (!e) return {ok:false,error:'Event not found.'};
  if (['Closed','Archived'].includes(String(e.Status))) return {ok:false,error:'This event is closed.'};
  const alloc=allocatedCounts_(eventId); const addTotal=Object.values(additions||{}).reduce((a,b)=>a+Number(b||0),0);
  if (e.Capacity && alloc.total+addTotal>e.Capacity) return {ok:false,error:`Event capacity exceeded. ${Math.max(0,e.Capacity-alloc.total)} place(s) remain.`};
  const tiers=getTiers_(eventId);
  for (const [k,n0] of Object.entries(additions||{})) { const n=Number(n0)||0; if (!k || !n) continue; const cap=tiers[k]&&Number(tiers[k].capacity)||0; if (cap && (alloc.byTier[k]||0)+n>cap) return {ok:false,error:`${tiers[k].label} capacity exceeded.`}; }
  return {ok:true};
}

function nextSerial_(eventId) {
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEETS.COUNTERS); let row=findRowBy_(sh,'EventID',eventId); let n=0;
    if (!row) { sh.appendRow([eventId,0]); row=findRowBy_(sh,'EventID',eventId); }
    n=Number(row.CurrentNumber)||0; n++; sh.getRange(row._row,2).setValue(n);
    const prefix=(getEvent_(eventId).SerialPrefix||eventId+'-').toString(); return prefix+String(n).padStart(3,'0');
  } finally { lock.releaseLock(); }
}
function ticketPrice_(eventId,tierKey) { const t=getTiers_(eventId)[tierKey]; return t?money_(t.price):0; }
function audit_(eventId,action,type,id,admin,details) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUDIT).appendRow([new Date(),eventId,action,type,id,admin||'',typeof details==='string'?details:JSON.stringify(details||{})]); }

function claimTicket_(body) {
  const voucherToken=String(body.voucher||'').trim(); let voucher=null; let eventId=normalizeId_(body.eventId);
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if (voucherToken) {
    voucher=findRowBy_(ss.getSheetByName(SHEETS.VOUCHERS),'VoucherToken',voucherToken);
    if (!voucher) return {ok:false,error:'Voucher not found.'};
    if (bool_(voucher.Claimed)) {
      const c=findRowBy_(ss.getSheetByName(SHEETS.CLAIMS),'Serial',voucher.Serial);
      return {ok:false,error:'This voucher was already claimed.',alreadyClaimed:true,serial:voucher.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken,eventId:voucher.EventID};
    }
    eventId=normalizeId_(voucher.EventID);
  }
  const e=getEvent_(eventId); if (!e) return {ok:false,error:'Choose a valid event.'};
  const fullName=String(body.fullName||'').trim(); if (!fullName) return {ok:false,error:'Name is required.'};
  let tier=String(body.tier||'').trim(); const tiers=getTiers_(eventId); if (voucher && voucher.TierKey) tier=String(voucher.TierKey); if (!tier) tier=Object.keys(tiers)[0]||''; if (!tiers[tier]) return {ok:false,error:'Invalid ticket type.'};
  let cap;
  if (voucher) {
    cap={ok:true};
    if (!voucher.TierKey) {
      const tc=tiers[tier]&&Number(tiers[tier].capacity)||0;
      const alloc=allocatedCounts_(eventId);
      if (tc && (alloc.byTier[tier]||0)+1>tc) cap={ok:false,error:`${tiers[tier].label} capacity exceeded.`};
    }
  } else cap=capacityCheck_(eventId,{[tier]:1});
  if (!cap.ok) return cap;
  const serial=nextSerial_(eventId), checkInToken=token_(32), price=ticketPrice_(eventId,tier);
  const paymentStatus=body.paymentStatus||'Pending', amountPaid=money_(body.amountPaid), method=String(body.paymentMethod||'');
  ss.getSheetByName(SHEETS.CLAIMS).appendRow([new Date(),eventId,serial,checkInToken,fullName,String(body.email||''),String(body.phone||''),String(body.groupName||body.chapter||''),tier,voucher?'Voucher':'Public',voucherToken,'Active',price,amountPaid,paymentStatus,method,String(body.paymentNote||'')]);
  if (voucher) { voucher.Claimed=true; voucher.Serial=serial; updateObjectRow_(ss.getSheetByName(SHEETS.VOUCHERS),VOUCHER_HEADERS,voucher._row,voucher); }
  if (amountPaid>0) ss.getSheetByName(SHEETS.PAYMENTS).appendRow([new Date(),eventId,serial,amountPaid,method,paymentStatus,String(body.paymentNote||''),'Claimant']);
  audit_(eventId,'CLAIM_TICKET','Claim',serial,fullName,{source:voucher?'Voucher':'Public'});
  return {ok:true,eventId,serial,checkInToken,tier,price,event:publicEvent_(e)};
}

function adminGenerateVouchers_(body,admin) {
  const eventId=normalizeId_(body.eventId), tiers=getTiers_(eventId), counts=body.tierCounts||{}; let openCount=Math.max(0,Number(body.openCount)||0); const add={}; let total=openCount;
  Object.keys(tiers).forEach(k=>{ const n=Math.max(0,Number(counts[k])||0); add[k]=n; total+=n; });
  if (total<1 || total>50) return {ok:false,error:'Generate between 1 and 50 vouchers per batch.'};
  const capAdd=Object.assign({},add); if (openCount) capAdd['']=openCount; const cap=capacityCheck_(eventId,capAdd); if (!cap.ok) return cap;
  const e=getEvent_(eventId); if (!e) return {ok:false,error:'Event not found.'};
  const batch=token_(16), sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VOUCHERS), now=new Date();
  const common={Timestamp:now,EventID:eventId,BatchID:batch,SuggestedGroup:String(body.groupName||''),PrefillName:String(body.distributorName||body.recipientName||''),PrefillPhone:String(body.distributorPhone||body.recipientPhone||''),Claimed:false,Serial:'',Dispatched:false,RecipientEmail:'',RecipientPhone:'',SentAt:'',IssuedBy:admin};
  Object.keys(add).forEach(k=>{for(let i=0;i<add[k];i++)sh.appendRow(objectToRow_(VOUCHER_HEADERS,Object.assign({},common,{VoucherToken:token_(16),TierKey:k})));});
  for(let i=0;i<openCount;i++)sh.appendRow(objectToRow_(VOUCHER_HEADERS,Object.assign({},common,{VoucherToken:token_(16),TierKey:''})));
  audit_(eventId,'GENERATE_VOUCHERS','Batch',batch,admin,{total,counts:add,open:openCount});
  return {ok:true,eventId,batch,count:total,distributorUrl:`${BASE_VOUCHER_URL}?batch=${encodeURIComponent(batch)}`,issuedBy:admin};
}
function getBatch_(batchId) {
  const rows=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VOUCHERS)).filter(r=>String(r.BatchID)===String(batchId)); if (!rows.length) return {ok:false,error:'Batch not found.'};
  const eventId=normalizeId_(rows[0].EventID); return {ok:true,eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),vouchers:rows.map(r=>({token:r.VoucherToken,tier:r.TierKey,claimed:bool_(r.Claimed),serial:r.Serial,dispatched:bool_(r.Dispatched),suggestedGroup:r.SuggestedGroup,prefillName:r.PrefillName,prefillPhone:r.PrefillPhone}))};
}
function getVoucher_(token) {
  const r=findRowBy_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VOUCHERS),'VoucherToken',token); if (!r) return {ok:false,error:'Voucher not found.'}; const eventId=normalizeId_(r.EventID);
  if (bool_(r.Claimed)) { const c=findRowBy_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS),'Serial',r.Serial); return {ok:false,alreadyClaimed:true,error:'Already claimed.',eventId,serial:r.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken,event:publicEvent_(getEvent_(eventId))}; }
  return {ok:true,eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),groups:getGroups_(eventId),tier:r.TierKey||'',groupName:r.SuggestedGroup||'',name:r.PrefillName||'',phone:r.PrefillPhone||''};
}
function sendVoucher_(body) {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VOUCHERS), r=findRowBy_(sh,'VoucherToken',String(body.token||'')); if (!r) return {ok:false,error:'Voucher not found.'}; if (bool_(r.Claimed)||bool_(r.Dispatched)) return {ok:false,error:'Voucher is already sent or claimed.'};
  const email=String(body.recipientEmail||'').trim(), phone=String(body.recipientPhone||'').trim(); if (!email&&!phone) return {ok:false,error:'Enter email or phone.'}; const url=`${BASE_VOUCHER_URL}?voucher=${encodeURIComponent(r.VoucherToken)}`;
  if (email) { try { MailApp.sendEmail(email,'Your event ticket voucher',`Claim your ticket: ${url}`); } catch(e) { return {ok:false,error:'Email could not be sent: '+e.message}; } }
  r.Dispatched=true;r.RecipientEmail=email;r.RecipientPhone=phone;r.SentAt=new Date();updateObjectRow_(sh,VOUCHER_HEADERS,r._row,r); audit_(r.EventID,'SEND_VOUCHER','Voucher',r.VoucherToken,'Distributor',{email,phone}); return {ok:true,url};
}

function normalizeTicketLookup_(eventId,value) {
  const s=String(value||'').trim(); const claims=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS)).filter(r=>normalizeId_(r.EventID)===eventId);
  let c=claims.find(r=>String(r.CheckInToken)===s); if (c) return c;
  c=claims.find(r=>String(r.Serial).toUpperCase()===s.toUpperCase()); if (c) return c;
  const digits=digitsOnly_(s); if (digits) c=claims.find(r=>digitsOnly_(r.Serial)===digits);
  return c||null;
}
function checkInWithTicket_(body,admin) {
  const eventId=normalizeId_(body.eventId); const c=normalizeTicketLookup_(eventId,body.serial||body.token); if (!c) return {ok:false,error:'Ticket not found for this event.'}; if (String(c.Status)==='Revoked') return {ok:false,error:'This ticket has been revoked.'};
  const check=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKINS)).find(r=>normalizeId_(r.EventID)===eventId && String(r.Serial)===String(c.Serial)); if (check) return {ok:false,alreadyCheckedIn:true,error:'Already checked in.',name:c.Name,serial:c.Serial,tier:c.TierKey,groupName:c.GroupName,checkedInAt:check.Timestamp,checkedInBy:check.CheckedInBy};
  const payStatus=String(body.paymentStatus||c.PaymentStatus||'Pending'), method=String(body.paymentMethod||c.PaymentMethod||''), amt=body.amountPaid!==undefined?money_(body.amountPaid):money_(c.AmountPaid), phone=String(body.phone||c.Phone||'');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKINS).appendRow([new Date(),eventId,c.Serial,c.Name,c.GroupName,c.TierKey,phone,payStatus,method,amt,admin,String(body.gateNote||'')]);
  updateClaimPayment_(c,payStatus,method,amt,String(body.paymentNote||''),admin);
  audit_(eventId,'CHECK_IN','Claim',c.Serial,admin,{paymentStatus:payStatus});
  return {ok:true,eventId,serial:c.Serial,name:c.Name,tier:c.TierKey,groupName:c.GroupName,paymentStatus:payStatus,paymentMethod:method,amountPaid:amt};
}
function checkInWalkIn_(body,admin) {
  const eventId=normalizeId_(body.eventId), fullName=String(body.fullName||'').trim(), tier=String(body.tier||'').trim(); if(!fullName)return{ok:false,error:'Name is required.'};
  const tiers=getTiers_(eventId); if(!tiers[tier])return{ok:false,error:'Invalid tier.'}; const cap=capacityCheck_(eventId,{[tier]:1}); if(!cap.ok)return cap;
  const serial=nextSerial_(eventId), token=token_(32), price=ticketPrice_(eventId,tier), amt=money_(body.amountPaid), status=String(body.paymentStatus||'Paid'), method=String(body.paymentMethod||'');
  const ss=SpreadsheetApp.getActiveSpreadsheet(); ss.getSheetByName(SHEETS.CLAIMS).appendRow([new Date(),eventId,serial,token,fullName,'',String(body.phone||''),String(body.groupName||''),tier,'WalkIn','','Active',price,amt,status,method,String(body.paymentNote||'')]);
  ss.getSheetByName(SHEETS.CHECKINS).appendRow([new Date(),eventId,serial,fullName,String(body.groupName||''),tier,String(body.phone||''),status,method,amt,admin,'Walk-in']);
  if(amt>0)ss.getSheetByName(SHEETS.PAYMENTS).appendRow([new Date(),eventId,serial,amt,method,status,String(body.paymentNote||''),admin]);
  audit_(eventId,'WALK_IN','Claim',serial,admin,{amountPaid:amt}); return {ok:true,eventId,serial,checkInToken:token,name:fullName,tier,price,paymentStatus:status,amountPaid:amt,event:publicEvent_(getEvent_(eventId))};
}
function updateClaimPayment_(claim,status,method,amount,note,admin) {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS); claim.PaymentStatus=status; claim.PaymentMethod=method; claim.AmountPaid=amount; if(note)claim.PaymentNote=note; updateObjectRow_(sh,CLAIM_HEADERS,claim._row,claim);
  if(amount>0) SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PAYMENTS).appendRow([new Date(),claim.EventID,claim.Serial,amount,method,status,note||'',admin]);
}

function searchGuests_(body) {
  const eventId=normalizeId_(body.eventId), q=String(body.query||'').trim().toLowerCase(); if(!q)return{ok:true,results:[]};
  const checks=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKINS)).filter(r=>normalizeId_(r.EventID)===eventId); const checked=new Set(checks.map(r=>String(r.Serial)));
  const results=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS)).filter(r=>normalizeId_(r.EventID)===eventId).filter(r=>[r.Name,r.Phone,r.Email,r.Serial].some(v=>String(v||'').toLowerCase().includes(q))).slice(0,30).map(r=>({serial:r.Serial,name:r.Name,phone:r.Phone,email:r.Email,groupName:r.GroupName,tier:r.TierKey,status:r.Status,paymentStatus:r.PaymentStatus,paymentMethod:r.PaymentMethod,amountDue:money_(r.AmountDue),amountPaid:money_(r.AmountPaid),checkedIn:checked.has(String(r.Serial))}));
  return {ok:true,results};
}
function ticketAdmin_(body,admin) {
  const eventId=normalizeId_(body.eventId), action=String(body.command||''), sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS); const c=normalizeTicketLookup_(eventId,body.serial); if(!c)return{ok:false,error:'Ticket not found.'};
  if(action==='revoke'){c.Status='Revoked';updateObjectRow_(sh,CLAIM_HEADERS,c._row,c);audit_(eventId,'REVOKE_TICKET','Claim',c.Serial,admin,'');return{ok:true,message:'Ticket revoked.'};}
  if(action==='reactivate'){c.Status='Active';updateObjectRow_(sh,CLAIM_HEADERS,c._row,c);audit_(eventId,'REACTIVATE_TICKET','Claim',c.Serial,admin,'');return{ok:true,message:'Ticket reactivated.'};}
  if(action==='reissueQr'){c.CheckInToken=token_(32);updateObjectRow_(sh,CLAIM_HEADERS,c._row,c);audit_(eventId,'REISSUE_QR','Claim',c.Serial,admin,'');return{ok:true,message:'QR credential reissued.',checkInToken:c.CheckInToken,serial:c.Serial,tier:c.TierKey,event:publicEvent_(getEvent_(eventId))};}
  if(action==='transfer'){ if(body.fullName)c.Name=String(body.fullName); if(body.phone!==undefined)c.Phone=String(body.phone); if(body.email!==undefined)c.Email=String(body.email); updateObjectRow_(sh,CLAIM_HEADERS,c._row,c);audit_(eventId,'TRANSFER_TICKET','Claim',c.Serial,admin,{name:c.Name});return{ok:true,message:'Ticket holder updated.'}; }
  if(action==='payment'){updateClaimPayment_(c,String(body.paymentStatus||c.PaymentStatus),String(body.paymentMethod||c.PaymentMethod),money_(body.amountPaid),String(body.paymentNote||''),admin);audit_(eventId,'UPDATE_PAYMENT','Claim',c.Serial,admin,{amount:body.amountPaid});return{ok:true,message:'Payment updated.'};}
  return {ok:false,error:'Unknown ticket command.'};
}

function dashboard_(eventId) {
  eventId=normalizeId_(eventId); const e=getEvent_(eventId); if(!e)return{ok:false,error:'Event not found.'}; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const claims=rowsAsObjects_(ss.getSheetByName(SHEETS.CLAIMS)).filter(r=>normalizeId_(r.EventID)===eventId && String(r.Status)!=='Revoked');
  const vouchers=rowsAsObjects_(ss.getSheetByName(SHEETS.VOUCHERS)).filter(r=>normalizeId_(r.EventID)===eventId); const checks=rowsAsObjects_(ss.getSheetByName(SHEETS.CHECKINS)).filter(r=>normalizeId_(r.EventID)===eventId);
  const paid=claims.filter(r=>String(r.PaymentStatus)==='Paid').length; const pending=claims.filter(r=>String(r.PaymentStatus)!=='Paid').length; const collected=claims.reduce((a,r)=>a+money_(r.AmountPaid),0); const due=claims.reduce((a,r)=>a+money_(r.AmountDue),0);
  const byTier={}; claims.forEach(r=>{const k=r.TierKey||'open';byTier[k]=(byTier[k]||0)+1;}); const byGroup={}; claims.forEach(r=>{const k=r.GroupName||'Unassigned';byGroup[k]=(byGroup[k]||0)+1;});
  return {ok:true,event:publicEvent_(e),issued:claims.length,pendingVouchers:vouchers.filter(r=>!bool_(r.Claimed)).length,checkedIn:checks.length,notArrived:Math.max(0,claims.length-checks.length),paid,pendingPayment:pending,collected,faceValue:due,capacity:e.Capacity,remaining:e.Capacity?Math.max(0,e.Capacity-allocatedCounts_(eventId).total):null,byTier,byGroup};
}
function report_(eventId) {
  const d=dashboard_(eventId); if(!d.ok)return d; const checks=rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKINS)).filter(r=>normalizeId_(r.EventID)===normalizeId_(eventId)); const byHour={}; checks.forEach(r=>{const dt=new Date(r.Timestamp);const key=isNaN(dt)?'Unknown':Utilities.formatDate(dt,Session.getScriptTimeZone()||'America/New_York','h a');byHour[key]=(byHour[key]||0)+1;}); return Object.assign(d,{byHour});
}
function auditList_(eventId) { return {ok:true,rows:rowsAsObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUDIT)).filter(r=>normalizeId_(r.EventID)===normalizeId_(eventId)).slice(-200).reverse().map(r=>({timestamp:r.Timestamp,action:r.Action,entityType:r.EntityType,entityId:r.EntityID,admin:r.Admin,details:r.Details}))}; }

function upsertEvent_(body,admin) {
  const eventId=normalizeId_(body.eventId); if(!eventId)return{ok:false,error:'Event ID is required.'}; const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(SHEETS.EVENTS); let r=findRowBy_(sh,'EventID',eventId); const now=new Date();
  const obj=r?Object.assign({},r):{EventID:eventId,CreatedAt:now,LogoFileId:'',BackgroundFileId:''};
  ['OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','WebsiteURL','DressCode','PrimaryColor','GoldColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status'].forEach(k=>{if(body[k]!==undefined)obj[k]=body[k];});
  obj.Status=obj.Status||'Draft';obj.UpdatedAt=now; if(r)updateObjectRow_(sh,EVENT_HEADERS,r._row,obj);else sh.appendRow(objectToRow_(EVENT_HEADERS,obj));
  if(!findRowBy_(ss.getSheetByName(SHEETS.COUNTERS),'EventID',eventId))ss.getSheetByName(SHEETS.COUNTERS).appendRow([eventId,0]);
  if(Array.isArray(body.tiers)){const tsh=ss.getSheetByName(SHEETS.TIERS); rowsAsObjects_(tsh).filter(x=>normalizeId_(x.EventID)===eventId).reverse().forEach(x=>tsh.deleteRow(x._row)); body.tiers.forEach((t,i)=>{if(!String(t.label||'').trim())return;const key=String(t.key||t.label).toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,30)||`tier${i+1}`;tsh.appendRow([eventId,key,String(t.label),money_(t.price),Number(t.capacity)||0,true]);});}
  if(Array.isArray(body.groups)){const gsh=ss.getSheetByName(SHEETS.GROUPS); rowsAsObjects_(gsh).filter(x=>normalizeId_(x.EventID)===eventId).reverse().forEach(x=>gsh.deleteRow(x._row));body.groups.map(String).map(s=>s.trim()).filter(Boolean).forEach(g=>gsh.appendRow([eventId,g,true]));}
  audit_(eventId,'UPSERT_EVENT','Event',eventId,admin,{status:obj.Status}); return{ok:true,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),groups:getGroups_(eventId)};
}
function setEventStatus_(body,admin) { const eventId=normalizeId_(body.eventId), status=String(body.status||''); if(!['Draft','Active','Closed','Archived'].includes(status))return{ok:false,error:'Invalid status.'}; const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVENTS),r=findRowBy_(sh,'EventID',eventId);if(!r)return{ok:false,error:'Event not found.'};r.Status=status;r.UpdatedAt=new Date();updateObjectRow_(sh,EVENT_HEADERS,r._row,r);audit_(eventId,'SET_EVENT_STATUS','Event',eventId,admin,status);return{ok:true,status}; }

function doGet(e) {
  const p=(e&&e.parameter)||{}, action=String(p.action||'');
  try {
    if(action==='events')return json_({ok:true,events:listEvents_(false)});
    if(action==='config'){const eventId=normalizeId_(p.event);const ev=getEvent_(eventId);if(!ev)return json_({ok:false,error:'Event not found.'});return json_({ok:true,event:publicEvent_(ev),tiers:getTiers_(eventId),groups:getGroups_(eventId)});}
    if(action==='voucher')return json_(getVoucher_(p.token));
    if(action==='batch')return json_(getBatch_(p.batch));
    return json_({ok:true,name:'Multi-Event Ticketing API'});
  } catch(err) { return json_({ok:false,error:err.message||String(err)}); }
}
function doPost(e) {
  const body=safeBody_(e), action=String(body.action||'');
  try {
    if(action==='claim') return json_(claimTicket_(body));
    if(action==='sendVoucher') return json_(sendVoucher_(body));
    if(action==='verifyPasscode'){const a=requireAdmin_(body);return json_(a);}
    const a=requireAdmin_(body); if(!a.ok)return json_(a);
    if(action==='eventsAdmin')return json_({ok:true,events:listEvents_(true)});
    if(action==='upsertEvent')return json_(upsertEvent_(body,a.name));
    if(action==='uploadBrandAsset')return json_(uploadBrandAsset_(body,a.name));
    if(action==='setEventStatus')return json_(setEventStatus_(body,a.name));
    if(action==='adminGenerateVouchers')return json_(adminGenerateVouchers_(body,a.name));
    if(action==='checkInWithTicket')return json_(checkInWithTicket_(body,a.name));
    if(action==='checkInWalkIn')return json_(checkInWalkIn_(body,a.name));
    if(action==='dashboard')return json_(dashboard_(body.eventId));
    if(action==='report')return json_(report_(body.eventId));
    if(action==='searchGuests')return json_(searchGuests_(body));
    if(action==='ticketAdmin')return json_(ticketAdmin_(body,a.name));
    if(action==='audit')return json_(auditList_(body.eventId));
    return json_({ok:false,error:'Unknown action.'});
  } catch(err) { return json_({ok:false,error:err.message||String(err)}); }
}
