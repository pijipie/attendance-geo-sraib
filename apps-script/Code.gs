// ===== Sistem Kehadiran Geofencing - Code.gs (v3) =====
// Records sheet: add header "DeviceID" in cell L1 before deploying.

function doGet(e) {
  return json_(getInitialData());
}

function doPost(e) {
  try {
    var c = JSON.parse(e.postData.contents);
    if (c.action === 'verifyStaff') return json_(verifyStaff(c));
    if (c.action === 'preview') return json_(previewDistance(c));
    return json_(processAttendance(c));
  } catch (err) {
    return json_({ success: false, status: 'Ralat Server', message: err.toString() });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isYes_(v) {
  var s = (v === undefined || v === null || v === '') ? 'YES' : v.toString().trim().toUpperCase();
  return s === 'YES' || s === 'ON' || s === 'TRUE' || s === '1';
}

function same_(a, b) {
  return a.toString().trim().toLowerCase() === b.toString().trim().toLowerCase();
}

// ---------- Initial data ----------
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { config: getConfigData(ss, false), events: getActiveEvents(ss) };
}

function getConfigData(ss, skipLogo) {
  var sheet = ss.getSheetByName("Config");
  var config = {};
  if (!sheet) return config;
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var key = data[i][0].toString().trim();
    var keyLower = key.toLowerCase().replace(/\s+/g, '');
    if (keyLower === 'logo' || keyLower === 'logourl') {
      if (skipLogo) continue; // logo is only needed by the page, not by POST requests
      var cell = sheet.getRange(i + 1, 2).getValue();
      var val = cell;
      if (cell && typeof cell === 'object' && typeof cell.getContentUrl === 'function') {
        try {
          var blob = UrlFetchApp.fetch(cell.getContentUrl()).getBlob();
          val = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
        } catch (e) { val = cell.getContentUrl(); }
      }
      config['Logo'] = val;
      config['LogoURL'] = val;
    } else {
      config[key] = data[i][1];
      if (keyLower === 'oranizationname' || keyLower === 'organizationname') config['OrganizationName'] = data[i][1];
    }
  }
  return config;
}

function getActiveEvents(ss) {
  var sheet = ss.getSheetByName("Events");
  var events = [];
  if (!sheet) return events;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && isYes_(data[i][5])) events.push({ eventId: data[i][0], eventName: data[i][1] });
  }
  return events;
}

// ---------- Lookups ----------
function findStaff_(ss, staffId) {
  var sheet = ss.getSheetByName("Staff");
  if (!sheet || !staffId) return { error: 'ID Tidak Sah' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && same_(data[i][0], staffId)) {
      if (!isYes_(data[i][3])) return { error: 'ID Tidak Aktif' };
      var id = data[i][0].toString().trim();
      return { id: id, name: data[i][1] || id, email: data[i][2] || '' };
    }
  }
  return { error: 'ID Tidak Sah' };
}

function findEvent_(ss, eventId) {
  var sheet = ss.getSheetByName("Events");
  if (!sheet || !eventId) return { error: 'Acara Tidak Sah' };
  var data = sheet.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    if (data[j][0] && same_(data[j][0], eventId)) {
      if (!isYes_(data[j][5])) return { error: 'Acara Tidak Aktif' };
      return {
        id: data[j][0].toString().trim(), name: data[j][1],
        lat: parseFloat(data[j][2]), lng: parseFloat(data[j][3]),
        radius: data[j][4] ? parseFloat(data[j][4]) : 100
      };
    }
  }
  return { error: 'Acara Tidak Sah' };
}

// ---------- Actions ----------
function verifyStaff(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = findStaff_(ss, (p.staffId || '').toString());
  if (s.error) return { success: false, status: s.error, message: 'ID Staf (' + p.staffId + ') ' + (s.error === 'ID Tidak Aktif' ? 'tidak aktif.' : 'tiada dalam rekod.') };
  return { success: true, name: s.name, email: s.email };
}

// Returns distance only; venue coordinates never leave the server.
function previewDistance(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ev = findEvent_(ss, (p.eventId || '').toString());
  if (ev.error) return { success: false, status: ev.error, message: 'Acara tidak sah / tidak aktif.' };
  var lat = parseFloat(p.latitude), lng = parseFloat(p.longitude);
  if (!isFinite(lat) || !isFinite(lng)) return { success: false, status: 'GPS Tidak Sah', message: 'Koordinat tidak sah.' };
  var d = Math.round(getDistance(lat, lng, ev.lat, ev.lng));
  return { success: true, distance: d, radius: ev.radius, inside: d <= ev.radius };
}

function processAttendance(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); // serialise writes so the duplicate check cannot be raced
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var records = ss.getSheetByName("Records");
    if (!records) {
      records = ss.insertSheet("Records");
      records.appendRow(["Timestamp", "StaffID", "Name", "Email", "EventID", "EventName",
        "Latitude", "Longitude", "Distance", "Status", "Remarks", "DeviceID"]);
    }

    var lat = parseFloat(p.latitude), lng = parseFloat(p.longitude);
    var accuracy = p.accuracy ? parseFloat(p.accuracy) : null;
    var deviceId = p.deviceId ? p.deviceId.toString().trim() : '';
    var config = getConfigData(ss, true);

    // Reject missing/NaN coordinates (previously these slipped through as "inside")
    if (!isFinite(lat) || !isFinite(lng) || lat < -90 || lat > 90 || lng < -180 || lng > 180) {
      return { success: false, status: 'GPS Tidak Sah', message: 'Koordinat GPS tidak sah atau tiada.' };
    }

    var staff = findStaff_(ss, (p.staffId || '').toString());
    if (staff.error) return { success: false, status: staff.error, message: 'ID Staf (' + p.staffId + ') ' + (staff.error === 'ID Tidak Aktif' ? 'bertaraf TIDAK AKTIF!' : 'tiada dalam rekod tab Staff!') };

    var ev = findEvent_(ss, (p.eventId || '').toString());
    if (ev.error) return { success: false, status: ev.error, message: ev.error === 'Acara Tidak Aktif' ? 'Acara ini telah ditutup / tidak aktif!' : 'Acara yang dipilih tidak wujud!' };

    var maxAcc = config['MaxGPSAccuracy'] ? parseFloat(config['MaxGPSAccuracy']) : 0;
    if (maxAcc > 0 && accuracy && accuracy > maxAcc) {
      return { success: false, status: 'GPS Lemah', message: 'Isyarat GPS terlalu lemah (±' + Math.round(accuracy) + 'm). Maksimum dibenarkan: ' + maxAcc + 'm.' };
    }

    var pd = (config['PreventDuplicate'] || '').toString().trim().toUpperCase();
    var preventDup = ['YES', 'ON', 'TRUE', '1'].indexOf(pd) >= 0;
    var dupRaw = config['Duplicate Minutes'] || config['DuplicateMinutes'] || config['Duplicate Minute'] || config['DuplicateMinute'];
    var dupMin = parseFloat(dupRaw) || 5;
    var rows = records.getDataRange().getValues();
    var now = new Date();

    for (var k = rows.length - 1; k >= 1; k--) {
      var okStatus = rows[k][9] && (rows[k][9].toString().toUpperCase() === 'BERJAYA' || rows[k][9].toString().toUpperCase() === 'HADIR');
      if (!okStatus || !rows[k][4] || !same_(rows[k][4], ev.id)) continue;

      // Buddy-punch guard: one device cannot check in a different staff ID for the same event
      if (deviceId && rows[k][11] && rows[k][11].toString() === deviceId && !same_(rows[k][1], staff.id)) {
        return { success: false, status: 'Peranti Sama', message: 'Peranti ini telah digunakan untuk ID staf lain bagi acara ini.' };
      }

      if (preventDup && same_(rows[k][1], staff.id)) {
        var t = new Date(rows[k][0]);
        if (!isNaN(t.getTime()) && (now - t) / 60000 < dupMin) {
          return { success: false, status: 'Rekod Bertindih', message: 'Anda telah mendaftar kehadiran bagi acara ini kurang dari ' + Math.ceil(dupMin) + ' minit yang lalu!' };
        }
      }
    }

    var dist = Math.round(getDistance(lat, lng, ev.lat, ev.lng));
    var inside = dist <= ev.radius;
    records.appendRow([now, staff.id, staff.name, staff.email, ev.id, ev.name, lat, lng, dist,
      inside ? "BERJAYA" : "DITOLAK", inside ? "Kehadiran Disahkan" : "Di luar kawasan (Maks: " + ev.radius + "m)", deviceId]);

    if (!inside) {
      return { success: false, status: 'Di Luar Kawasan', message: 'LOKASI DITOLAK! Anda berada ' + dist + 'm dari tempat acara (Maksimum dibenarkan: ' + ev.radius + 'm).' };
    }
    return { success: true, status: 'Berjaya', message: 'Kehadiran ' + staff.name + ' disahkan! (' + dist + 'm dari lokasi)' };
  } finally {
    lock.releaseLock();
  }
}

function getDistance(lat1, lon1, lat2, lon2) {
  var R = 6371000, rad = Math.PI / 180;
  var dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
