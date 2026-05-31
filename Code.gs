// ===================================================
// Google Apps Script — ส่งและรับข้อมูลกับเว็บแอป
// ===================================================

var SHEET_NAME = 'ข้อมูลสำรวจ'; // ชื่อ Sheet

// ===== รับข้อมูลจากเว็บ (POST) =====
function doPost(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'วันที่บันทึก','วันที่สำรวจ','หมู่ที่','ตำบล','อำเภอ','ผู้สำรวจ',
        'หลังที่','เลขที่บ้าน','ชื่อเจ้าบ้าน',
        'โอ่งน้ำกิน(สำรวจ)','โอ่งน้ำกิน(ลูกน้ำ)',
        'น้ำใช้(สำรวจ)','น้ำใช้(ลูกน้ำ)',
        'ตู้กับข้าว(สำรวจ)','ตู้กับข้าว(ลูกน้ำ)',
        'โอ่งใหญ่(สำรวจ)','โอ่งใหญ่(ลูกน้ำ)',
        'ยางรถ(สำรวจ)','ยางรถ(ลูกน้ำ)',
        'วงบ่อ(สำรวจ)','วงบ่อ(ลูกน้ำ)',
        'อื่นๆ(สำรวจ)','อื่นๆ(ลูกน้ำ)',
        'ภาชนะรวม','ภาชนะมีลูกน้ำ','พบลูกน้ำ','มาตรการ'
      ]);
      sheet.getRange(1,1,1,27).setBackground('#0D5C3A').setFontColor('#ffffff').setFontWeight('bold');
    }

    var data = JSON.parse(e.postData.contents);
    sheet.appendRow([
      new Date(), data.date, data.moo, data.tambon, data.amphoe, data.surveyor,
      data.houseNo, data.addr, data.owner,
      data.c0s, data.c0p, data.c1s, data.c1p,
      data.c2s, data.c2p, data.c3s, data.c3p,
      data.c4s, data.c4p, data.c5s, data.c5p,
      data.c6s, data.c6p,
      data.totalC, data.posC,
      data.hasLarvae ? 'พบ' : 'ไม่พบ',
      data.measures
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== ส่งข้อมูลกลับให้เว็บ (GET) =====
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({status:'ok', data:[]}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, 27).getValues();
    var result = [];

    rows.forEach(function(row) {
      // ข้ามแถวว่าง
      if (!row[1] && !row[3]) return;

      // แปลงวันที่
      var dateVal = row[1];
      var dateStr = '';
      if (dateVal instanceof Date) {
        var y = dateVal.getFullYear();
        var m = String(dateVal.getMonth()+1).padStart(2,'0');
        var d = String(dateVal.getDate()).padStart(2,'0');
        dateStr = y+'-'+m+'-'+d;
      } else if (typeof dateVal === 'string') {
        dateStr = dateVal;
      }

      result.push({
        date:     dateStr,
        moo:      String(row[2]||''),
        tambon:   String(row[3]||''),
        amphoe:   String(row[4]||''),
        surveyor: String(row[5]||''),
        houseNo:  row[6],
        addr:     String(row[7]||''),
        owner:    String(row[8]||''),
        c0s:Number(row[9]||0),  c0p:Number(row[10]||0),
        c1s:Number(row[11]||0), c1p:Number(row[12]||0),
        c2s:Number(row[13]||0), c2p:Number(row[14]||0),
        c3s:Number(row[15]||0), c3p:Number(row[16]||0),
        c4s:Number(row[17]||0), c4p:Number(row[18]||0),
        c5s:Number(row[19]||0), c5p:Number(row[20]||0),
        c6s:Number(row[21]||0), c6p:Number(row[22]||0),
        totalC:   Number(row[23]||0),
        posC:     Number(row[24]||0),
        hasLarvae: String(row[25]||'') === 'พบ',
        measures: String(row[26]||'')
      });
    });

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok', data:result}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ทดสอบ
function testSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('✅ พร้อมใช้งาน! Spreadsheet: ' + ss.getName());
  Browser.msgBox('✅ Script พร้อมใช้งานแล้ว!\nSpreadsheet: ' + ss.getName());
}
