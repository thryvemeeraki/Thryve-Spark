// ═══════════════════════════════════════════════════════════════
//  THRYVE SPARK — Google Apps Script Backend  (v2)
//  Actions: add | check | use | submit | append
//  Deploy as: Execute as Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

var NOTIFY_EMAIL    = "thryvemeeraki@gmail.com";
var OTP_SHEET       = "OTPs";
var RESPONSES_SHEET = "Responses";
var PENDING_SHEET   = "Pending";

function doGet(e) {
  var p        = e.parameter;
  var callback = p.callback || "callback";
  var action   = p.action   || "";
  var result   = { error: "unknown action" };
  try {
    if      (action === "add")    result = addOTP(p.otp);
    else if (action === "check")  result = checkOTP(p.otp);
    else if (action === "use")    result = useOTP(p.otp, p.name);
    else if (action === "submit") result = submitForm(p.otp, p.name, p.formData, p.chunk, p.total);
    else if (action === "append") result = appendChunk(p.otp, p.formData, p.chunk, p.total);
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function addOTP(otp) {
  if (!otp) return { added: false, error: "No OTP provided" };
  var sheet = getSheet(OTP_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["OTP","Status","Created At","Used By","Used At"]);
    sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1E1A0E").setFontColor("#B8962E");
  }
  sheet.appendRow([otp,"ACTIVE",new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata"}),"",""]);
  return { added: true, otp: otp };
}

function checkOTP(otp) {
  if (!otp) return { valid: false, reason: "no_otp" };
  var sheet = getSheet(OTP_SHEET);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(otp).trim()) {
      var status = String(data[i][1]).trim().toUpperCase();
      if (status === "ACTIVE") return { valid: true };
      if (status === "USED")   return { valid: false, reason: "already_used" };
      return { valid: false, reason: "invalid_status" };
    }
  }
  return { valid: false, reason: "not_found" };
}

function useOTP(otp, name) {
  if (!otp) return { ok: false, error: "No OTP" };
  var sheet = getSheet(OTP_SHEET);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(otp).trim()) {
      sheet.getRange(i+1,2).setValue("USED");
      sheet.getRange(i+1,4).setValue(name||"");
      sheet.getRange(i+1,5).setValue(new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata"}));
      return { ok: true };
    }
  }
  return { ok: false, error: "OTP not found" };
}

function submitForm(otp, name, formDataJson, chunk, total) {
  useOTP(otp, name);
  var formData = {};
  try { formData = JSON.parse(formDataJson || "{}"); } catch(e) {}
  var totalChunks = parseInt(total) || 1;
  if (totalChunks <= 1) {
    saveToResponsesSheet(otp, name, formData);
    sendEmailNotification(otp, name, formData);
    return { ok: true };
  }
  storePendingChunk(otp, name, 0, totalChunks, formData);
  return { ok: true, pending: true };
}

function appendChunk(otp, formDataJson, chunk, total) {
  var formData = {};
  try { formData = JSON.parse(formDataJson || "{}"); } catch(e) {}
  var chunkIdx    = parseInt(chunk) || 0;
  var totalChunks = parseInt(total) || 1;
  storePendingChunk(otp, "", chunkIdx, totalChunks, formData);
  var allData = getAllPendingChunks(otp, totalChunks);
  if (allData !== null) {
    saveToResponsesSheet(otp, allData.name, allData.data);
    sendEmailNotification(otp, allData.name, allData.data);
    clearPendingChunks(otp);
  }
  return { ok: true };
}

function storePendingChunk(otp, name, chunkIdx, total, data) {
  var sheet = getSheet(PENDING_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["OTP","Name","ChunkIdx","Total","Data","Timestamp"]);
  }
  sheet.appendRow([otp, name||"", chunkIdx, total, JSON.stringify(data), new Date().toISOString()]);
}

function getAllPendingChunks(otp, total) {
  var sheet = getSheet(PENDING_SHEET);
  if (sheet.getLastRow() <= 1) return null;
  var rows   = sheet.getDataRange().getValues();
  var chunks = {};
  var name   = "";
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(otp).trim()) {
      chunks[parseInt(rows[i][2])] = rows[i][4];
      if (rows[i][1]) name = rows[i][1];
    }
  }
  if (Object.keys(chunks).length < total) return null;
  var merged = {};
  for (var c = 0; c < total; c++) {
    try {
      var chunkData = JSON.parse(chunks[c] || "{}");
      Object.keys(chunkData).forEach(function(k) { merged[k] = chunkData[k]; });
    } catch(e) {}
  }
  return { name: name, data: merged };
}

function clearPendingChunks(otp) {
  var sheet = getSheet(PENDING_SHEET);
  if (sheet.getLastRow() <= 1) return;
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === String(otp).trim()) sheet.deleteRow(i+1);
  }
}

function saveToResponsesSheet(otp, name, formData) {
  var sheet = getSheet(RESPONSES_SHEET);
  var keys  = Object.keys(formData);
  if (sheet.getLastRow() === 0) {
    var headers = ["Timestamp","OTP","Student Name"].concat(keys);
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#1E1A0E").setFontColor("#B8962E");
    sheet.setFrozenRows(1);
  }
  var existingHeaders = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  keys.forEach(function(key) {
    if (existingHeaders.indexOf(key) === -1) {
      var newCol = existingHeaders.length + 1;
      sheet.getRange(1,newCol).setValue(key).setFontWeight("bold").setBackground("#1E1A0E").setFontColor("#B8962E");
      existingHeaders.push(key);
    }
  });
  var row = new Array(existingHeaders.length).fill("");
  row[0] = new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata"});
  row[1] = otp;
  row[2] = name;
  keys.forEach(function(key) {
    var idx = existingHeaders.indexOf(key);
    if (idx !== -1) row[idx] = formData[key] || "";
  });
  sheet.appendRow(row);
  try { sheet.autoResizeColumns(1, existingHeaders.length); } catch(e) {}
}

function sendEmailNotification(otp, name, formData) {
  try {
    var subject = "✦ Thryve Spark Response — " + (name||"Student") + " (" + otp + ")";
    var html = [
      '<div style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto;border-radius:12px;overflow:hidden;border:1px solid #E2D9C4;">',
      '<div style="background:#1E1A0E;padding:24px 32px;">',
      '<h1 style="color:#B8962E;margin:0;font-size:22px;">✦ Thryve Spark Assessment</h1>',
      '<p style="color:rgba(255,255,255,0.55);margin:6px 0 0;font-size:13px;">New student submission received</p>',
      '</div>',
      '<div style="background:#FDFAF3;padding:24px 32px;">',
      '<table style="width:100%;border-collapse:collapse;margin-bottom:28px;">',
      '<tr><td style="padding:10px 14px;background:#F5EDD6;font-weight:700;color:#1E1A0E;width:35%;border:1px solid #E2D9C4;font-size:13px;">Student Name</td><td style="padding:10px 14px;border:1px solid #E2D9C4;font-size:13px;font-weight:600;">'+(name||"—")+'</td></tr>',
      '<tr><td style="padding:10px 14px;background:#F5EDD6;font-weight:700;color:#1E1A0E;border:1px solid #E2D9C4;font-size:13px;">OTP</td><td style="padding:10px 14px;border:1px solid #E2D9C4;font-size:13px;">'+otp+'</td></tr>',
      '<tr><td style="padding:10px 14px;background:#F5EDD6;font-weight:700;color:#1E1A0E;border:1px solid #E2D9C4;font-size:13px;">Submitted At</td><td style="padding:10px 14px;border:1px solid #E2D9C4;font-size:13px;">'+new Date().toLocaleString("en-IN",{timeZone:"Asia/Kolkata"})+'</td></tr>',
      '</table>'
    ];

    // Group by section
    var sections = {};
    var profileKeys = ["Full Name","Class/Grade","Contact Number","Date","School Name","Parent / Guardian","Email Address","Student DOB","Student Time of Birth","Student Place of Birth","Father DOB","Father Time of Birth","Father Place of Birth","Mother DOB","Mother Time of Birth","Mother Place of Birth"];

    Object.keys(formData).forEach(function(key) {
      var sec;
      if (profileKeys.indexOf(key) !== -1) {
        if (key.indexOf("Father") !== -1 || key.indexOf("Mother") !== -1) sec = "Parental Birth Details";
        else if (key.indexOf("Student") !== -1) sec = "Student Birth Details";
        else sec = "Student Profile";
      } else {
        var m = key.match(/^(.+?)\s+Q\d+/);
        sec = m ? m[1].trim() : "Other";
      }
      if (!sections[sec]) sections[sec] = [];
      sections[sec].push({ q: key, a: formData[key] });
    });

    Object.keys(sections).forEach(function(sec) {
      html.push(
        '<div style="margin-bottom:24px;">',
        '<h3 style="color:#B8962E;margin:0 0 10px;font-size:15px;border-bottom:2px solid #E2D9C4;padding-bottom:8px;">'+sec+'</h3>',
        '<table style="width:100%;border-collapse:collapse;">'
      );
      sections[sec].forEach(function(item, idx) {
        var bg = idx%2===0 ? '#ffffff' : '#F9F6EF';
        var displayQ = item.q.replace(/^[A-Za-z\s\/]+Q\d+:\s*/,'').trim() || item.q;
        var displayA = item.a ? String(item.a).replace(/</g,'&lt;').replace(/>/g,'&gt;') : '<span style="color:#bbb;font-style:italic;">Not answered</span>';
        html.push(
          '<tr style="background:'+bg+'">',
          '<td style="padding:9px 12px;border:1px solid #E2D9C4;font-size:12.5px;font-weight:600;color:#3A2F14;width:50%;vertical-align:top;">'+displayQ+'</td>',
          '<td style="padding:9px 12px;border:1px solid #E2D9C4;font-size:12.5px;color:#1E1A0E;vertical-align:top;">'+displayA+'</td>',
          '</tr>'
        );
      });
      html.push('</table></div>');
    });

    html.push(
      '</div>',
      '<div style="background:#1E1A0E;padding:12px 32px;text-align:center;">',
      '<p style="color:rgba(255,255,255,0.35);font-size:11px;margin:0;">Thryve Spark v4 · Secured by Google Sheets</p>',
      '</div></div>'
    );

    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, htmlBody: html.join("") });
  } catch(e) {
    Logger.log("Email error: " + e.toString());
  }
}
