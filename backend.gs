/**
 * ระบบแจ้งซ่อม โรงเรียนมุกดาหาร
 * Google Apps Script backend
 *
 * วิธีใช้:
 * 1) สร้าง Google Sheet 1 ไฟล์ แล้วเปิด Extensions > Apps Script
 * 2) วางโค้ดนี้ทั้งหมด
 * 3) Deploy > New deployment > Web app
 * 4) Execute as: Me
 * 5) Who has access: Anyone
 * 6) นำ Web app URL มาใส่ในเว็บหน้า index.html
 */

const SHEET_NAME = 'Reports';
const STAFF_PIN = '4321';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const headers = [
    'id','studentClass','studentNo','place','spot','category',
    'urgency','detail','status','createdAt','historyJson'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return 'พร้อมใช้งาน';
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ping';

  if (action === 'ping') {
    return json_({ok:true, message:'Mukdahan Repair API is online'});
  }

  if (action === 'track') {
    return json_(track_(e.parameter.code));
  }

  if (action === 'duplicate') {
    return json_(duplicate_(e.parameter.place, e.parameter.spot, e.parameter.category));
  }

  if (action === 'staff') {
    return json_(staffList_(e.parameter.pin));
  }

  return json_({ok:false,error:'unknown_action'});
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (data.action === 'report') return json_(createReport_(data));
    if (data.action === 'update') return json_(updateReport_(data));
    if (data.action === 'delete') return json_(deleteReport_(data));

    return json_({ok:false,error:'unknown_action'});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    setup();
    sh = ss.getSheetByName(SHEET_NAME);
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(function(row) {
    return {
      id: Number(row[0]),
      studentClass: String(row[1] || ''),
      studentNo: Number(row[2] || 0),
      place: String(row[3] || ''),
      spot: String(row[4] || ''),
      category: String(row[5] || ''),
      urgency: String(row[6] || 'low'),
      detail: String(row[7] || ''),
      status: String(row[8] || 'new'),
      createdAt: Number(row[9] || 0),
      history: parseHistory_(row[10])
    };
  }).filter(function(r){ return r.id; });
}

function parseHistory_(value) {
  try {
    return value ? JSON.parse(String(value)) : [];
  } catch (_) {
    return [];
  }
}

function writeRow_(r) {
  const sh = sheet_();
  sh.appendRow([
    r.id, r.studentClass, r.studentNo, r.place, r.spot, r.category,
    r.urgency, r.detail, r.status, r.createdAt, JSON.stringify(r.history)
  ]);
}

function findRow_(id) {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    if (Number(values[i][0]) === Number(id)) return i + 1;
  }
  return 0;
}

function nextId_() {
  const all = rows_();
  return all.reduce(function(max,r){ return Math.max(max,r.id || 0); }, 0) + 1;
}

function createReport_(d) {
  if (!d.studentClass || !d.studentNo || !d.place || !d.spot || !d.category || !d.detail) {
    return {ok:false,error:'ข้อมูลไม่ครบ'};
  }

  const active = rows_().filter(function(r) {
    return r.status !== 'done' &&
      r.place === String(d.place) &&
      r.spot.toLowerCase() === String(d.spot).toLowerCase() &&
      r.category === String(d.category);
  });

  const now = Date.now();
  const r = {
    id: nextId_(),
    studentClass: String(d.studentClass),
    studentNo: Number(d.studentNo),
    place: String(d.place),
    spot: String(d.spot),
    category: String(d.category),
    urgency: String(d.urgency || 'low'),
    detail: String(d.detail),
    status: 'new',
    createdAt: now,
    history: [{status:'new',at:now,note:'รับเรื่องแจ้งซ่อมแล้ว'}]
  };

  writeRow_(r);

  return {
    ok:true,
    report:safePublic_(r),
    duplicate:active.length > 0,
    duplicateCount:active.length
  };
}

function track_(code) {
  const id = Number(String(code || '').replace('#',''));
  const r = rows_().find(function(x){ return x.id === id; });
  if (!r) return {ok:false,error:'ไม่พบรหัสแจ้งซ่อมนี้'};
  return {ok:true,report:safePublic_(r)};
}

function duplicate_(place, spot, category) {
  const p = String(place || '');
  const s = String(spot || '').trim().toLowerCase();
  const c = String(category || '');

  const found = rows_().filter(function(r) {
    return r.status !== 'done' &&
      r.place === p &&
      (!s || r.spot.toLowerCase() === s) &&
      (!c || r.category === c);
  });

  return {
    ok:true,
    count:found.length,
    reports:found.map(safePublic_)
  };
}

function staffList_(pin) {
  if (String(pin || '') !== STAFF_PIN) {
    return {ok:false,error:'รหัสเจ้าหน้าที่ไม่ถูกต้อง'};
  }
  return {ok:true,reports:rows_().sort(function(a,b){return b.createdAt-a.createdAt;})};
}

function updateReport_(d) {
  if (String(d.pin || '') !== STAFF_PIN) {
    return {ok:false,error:'รหัสเจ้าหน้าที่ไม่ถูกต้อง'};
  }

  const id = Number(d.id);
  const row = findRow_(id);
  if (!row) return {ok:false,error:'ไม่พบรายการ'};

  const all = rows_();
  const r = all.find(function(x){ return x.id === id; });
  if (!r) return {ok:false,error:'ไม่พบรายการ'};

  const newStatus = String(d.status || r.status);
  const note = String(d.note || '').trim();
  const now = Date.now();

  if (r.status !== newStatus || note) {
    r.status = newStatus;
    r.history.push({
      status:newStatus,
      at:now,
      note:note || ('อัปเดตสถานะเป็น ' + newStatus)
    });
  }

  const sh = sheet_();
  sh.getRange(row,1,1,11).setValues([[
    r.id,r.studentClass,r.studentNo,r.place,r.spot,r.category,
    r.urgency,r.detail,r.status,r.createdAt,JSON.stringify(r.history)
  ]]);

  return {ok:true,report:safePublic_(r)};
}

function deleteReport_(d) {
  if (String(d.pin || '') !== STAFF_PIN) {
    return {ok:false,error:'รหัสเจ้าหน้าที่ไม่ถูกต้อง'};
  }

  const row = findRow_(Number(d.id));
  if (!row) return {ok:false,error:'ไม่พบรายการ'};
  sheet_().deleteRow(row);
  return {ok:true};
}

function safePublic_(r) {
  return {
    id:r.id,
    place:r.place,
    spot:r.spot,
    category:r.category,
    urgency:r.urgency,
    detail:r.detail,
    status:r.status,
    createdAt:r.createdAt,
    history:r.history
  };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
