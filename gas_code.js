// ─── Spreadsheet ID'leri ───────────────────────────────────────────
var DICT_ID = '1Np2b7Im5q55kXHRhPc3euvlPGiZ_7anTfvYSly1hvHQ';  // FTW Dictionaries
var RESULTS = {
  'TR': '1B5F24YxM3aOedWoiB7oJ1xcO2MucRcLmwdG0pFODCGQ',         // FTW Results TR
  'EN': '1lg-u01ymZZB0nK0GOnYFt3cq6tZ39Z1LL36bKsMCM18',         // FTW Results EN
  'RU': '139sob08OmOENQfvJyXYIRbgzmlUVwzrGMl9sMBfXocA',         // FTW Results RU
  'ES': '13X9nZdrmiCfw_UxnqVbYo1D4KWba9j77YmOCOutRDjE',         // FTW Results ES
  'PT': '11nLTTSVbcxcYZATeUZoUn-PMfFLfr8nOU7sU8KreW-E',         // FTW Results PT
  'FR': '14YHoEqb12na4NQ4_CUD-O1Rz-AqJeloPvkoJ0YeBg4w',         // FTW Results FR
  'DE': '1BbKGtof9FuBgTcsD1xB3UrGP0ZGtEjkxGpCzEhW-m10'          // FTW Results DE
};

// ─── GET router ───────────────────────────────────────────────────
// Not: Frontend (index.html) CORS nedeniyle her şeyi GET ile çağırır;
// submitBatch de query param olarak gelir (words = JSON string).
function doGet(e) {
  var p = e.parameter;
  var result;
  try {
    if      (p.action === 'getWords')    result = getWords(p.task || '', p.worker || '');
    else if (p.action === 'submit')      result = submitWord(p);
    else if (p.action === 'submitBatch') result = submitBatch({
      action: 'submitBatch',
      task:   p.task,
      worker: p.worker,
      words:  JSON.parse(p.words || '[]')
    });
    else                                 result = { error: 'Unknown action: ' + p.action };
  } catch (ex) {
    result = { error: ex.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST router (batch submit için) ─────────────────────────────
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'submitBatch') {
      result = submitBatch(body);
    } else {
      result = { error: 'Unknown POST action: ' + body.action };
    }
  } catch (ex) {
    result = { error: ex.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Kelimeleri getir ─────────────────────────────────────────────
function getWords(task, worker) {
  task = task.toUpperCase();
  var lang = task.split('_')[0];

  if (!RESULTS[lang]) return { error: 'Bilinmeyen dil: ' + lang };

  var resSS    = SpreadsheetApp.openById(RESULTS[lang]);
  var tasksTab = resSS.getSheetByName(lang + '_Tasks');
  if (!tasksTab) return { error: lang + '_Tasks sekmesi bulunamadı' };

  var taskConfig = findTask(tasksTab, task);
  if (!taskConfig)        return { error: 'Task bulunamadı: ' + task };
  if (!taskConfig.active) return { error: 'Bu task aktif değil: ' + task };

  var dictSS  = SpreadsheetApp.openById(DICT_ID);
  var dictTab = dictSS.getSheetByName(lang);
  if (!dictTab) return { error: '"' + lang + '" sekmesi FTW Dictionaries\'ta bulunamadı.' };

  var count   = taskConfig.end - taskConfig.start + 1;
  var rawData = dictTab.getRange(taskConfig.start + 1, 1, count, dictTab.getLastColumn()).getValues();

  var done = getExistingResults(resSS, lang, worker, task);

  var words = [];
  rawData.forEach(function(r, idx) {
    var kelime = String(r[0] || '').trim();
    if (!kelime) return;
    var rowNum = taskConfig.start + idx;
    // Aynı kelime task içinde tekrar edebiliyor; önce row (satır no) ile eşleştir,
    // sadece row bilgisi olmayan (v8 öncesi) eski kayıtlar için kelimeye düş.
    var existingType = done.byRow[rowNum];
    if (existingType === undefined) existingType = done.byKelime[kelime];
    words.push({
      row:         rowNum,
      kelime:      kelime,
      harf_sayisi: (r[1] != null && r[1] !== '') ? r[1] : kelime.length,
      score:       (r[2] != null && r[2] !== '') ? r[2] : 0,
      type:        existingType !== undefined ? existingType : null
    });
  });

  return { words: words };
}

// ─── Tek kelime kaydet (eski uyumluluk için) ──────────────────────
function submitWord(p) {
  var task   = String(p.task   || '').toUpperCase().trim();
  var worker = String(p.worker || '').trim();
  var kelime = String(p.kelime || '').trim();
  var type   = parseInt(p.type);
  var row    = (p.row !== undefined && p.row !== null && p.row !== '') ? parseInt(p.row) : null;

  if (!task || !worker || !kelime || isNaN(type)) return { error: 'Eksik parametre' };

  var lang = task.split('_')[0];
  if (!RESULTS[lang]) return { error: 'Bilinmeyen dil: ' + lang };

  // Lock ile race condition'ı engelle
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (ex) {
    return { error: 'Sunucu meşgul, tekrar deneyin.' };
  }

  try {
    var resSS      = SpreadsheetApp.openById(RESULTS[lang]);
    var resultsTab = resSS.getSheetByName(lang + '_Results');
    if (!resultsTab) return { error: lang + '_Results sekmesi bulunamadı' };

    ensureHeader_(resultsTab);
    upsertWord_(resultsTab, RESULTS[lang], worker, task, kelime, p.harf_sayisi || kelime.length, type, row);
    return { ok: true };
  } finally {
    // flush olmadan releaseLock çağırmak, bir sonraki execution'ın bu yazmayı
    // henüz görmeden (stale veriyle) başlamasına yol açabilir — bkz. submitBatch.
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ─── Batch submit (sayfa bazında toplu kayıt) ─────────────────────
// body: { action: 'submitBatch', task, worker, words: [{kelime, harf_sayisi, type}, ...] }
function submitBatch(body) {
  var task   = String(body.task   || '').toUpperCase().trim();
  var worker = String(body.worker || '').trim();
  var words  = body.words;

  if (!task || !worker || !words || !words.length) return { error: 'Eksik parametre' };

  var lang = task.split('_')[0];
  if (!RESULTS[lang]) return { error: 'Bilinmeyen dil: ' + lang };

  // Lock ile atomik yazma garanti et
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (ex) {
    return { error: 'Sunucu meşgul, tekrar deneyin.' };
  }

  try {
    var resSS      = SpreadsheetApp.openById(RESULTS[lang]);
    var resultsTab = resSS.getSheetByName(lang + '_Results');
    if (!resultsTab) return { error: lang + '_Results sekmesi bulunamadı' };

    ensureHeader_(resultsTab);
    var sheetRowCol = ensureSheetRowIndex_(resultsTab);
    SpreadsheetApp.flush();

    var map    = headerMap_(resultsTab);
    var iScore = map.score - 1;
    var iTime  = map.timestamp - 1;
    var iRow   = map.row - 1;

    // Tüm sonuç sayfasını Apps Script'e çekme. 45 bin satırlık DE_Results
    // getValues() ile 60sn'yi aşıyordu. Sheets'in kendi filtresi sadece bu
    // worker + task satırlarını döndürür (~1sn).
    var found = queryWorkerTask_(RESULTS[lang], resultsTab, worker, task);
    var existingByRow = {};
    var existingByKelime = {};
    found.forEach(function(r) {
      if (!r.sheetRow) return;
      if (r.dictRow !== '' && r.dictRow != null) existingByRow[String(r.dictRow)] = r.sheetRow;
      else existingByKelime[r.kelime] = r.sheetRow;
    });

    var now = new Date();
    var newRows = [];
    var updatedCount = 0;

    words.forEach(function(w) {
      var kelime = String(w.kelime || '').trim();
      var type   = parseInt(w.type);
      if (!kelime || isNaN(type)) return;

      var harf   = w.harf_sayisi || kelime.length;
      var rowNum = (w.row !== undefined && w.row !== null && w.row !== '') ? parseInt(w.row) : null;
      var rowKey = rowNum !== null ? String(rowNum) : null;

      var existingRow = rowKey !== null ? existingByRow[rowKey] : undefined;
      var healedFromKelime = false;
      if (existingRow === undefined && existingByKelime.hasOwnProperty(kelime)) {
        existingRow = existingByKelime[kelime];
        healedFromKelime = true;
        delete existingByKelime[kelime]; // aynı kelime tekrar ediyorsa bu eski satır sadece bir kez "iyileştirilsin"
      }

      if (existingRow) {
        // Mevcut satırı güncelle
        resultsTab.getRange(existingRow, iScore + 1).setValue(type);
        resultsTab.getRange(existingRow, iTime + 1).setValue(now);
        if (healedFromKelime && rowKey !== null && iRow !== -1) {
          resultsTab.getRange(existingRow, iRow + 1).setValue(rowNum); // row bilgisi eksik eski kaydı tamamla
        }
        updatedCount++;
      } else {
        // Yeni satır olarak eklenecekler listesine ekle
        newRows.push([worker, task, kelime, harf, type, now, rowNum]);
      }
    });

    // Yeni satırları toplu ekle (appendRow yerine setValues — çok daha hızlı)
    if (newRows.length > 0) {
      var lastRow = resultsTab.getLastRow();
      resultsTab.getRange(lastRow + 1, 1, newRows.length, 7).setValues(newRows);
      var indexVals = [];
      for (var n = 0; n < newRows.length; n++) indexVals.push([lastRow + 1 + n]);
      resultsTab.getRange(lastRow + 1, sheetRowCol, indexVals.length, 1).setValues(indexVals);
    }

    return { ok: true, saved: words.length, updated: updatedCount, inserted: newRows.length };
  } finally {
    // KRİTİK: flush() olmadan lock.releaseLock() çağrılırsa, bu execution'ın
    // yazdığı satırlar diğer (kilidi hemen sonra alan) execution'a stale
    // görünebilir. Sonuç: iki worker aynı task'ta aynı anda kayıt gönderirse,
    // ikincisi "lastRow"u eski okuyup birincinin satırlarının üzerine yazabilir
    // (veri kaybı) veya aynı kelime için ikinci bir satır ekleyebilir (duplicate).
    // flush() burada bu execution'ın tüm yazmalarını gerçek sheet'e commit eder,
    // böylece lock'u bırakır bırakmaz bekleyen execution tutarlı veriyi okur.
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ─── Yardımcı: Header satırını garanti et ─────────────────────────
function ensureHeader_(tab) {
  if (tab.getLastRow() === 0) {
    tab.appendRow(['worker_email', 'task_id', 'kelime', 'harf_sayisi', 'score', 'timestamp', 'row']);
    return;
  }
  // v8 öncesi oluşturulmuş sheet'lerde 'row' kolonu yok — var olan veriye
  // dokunmadan sona ekle (mevcut satırlarda bu kolon boş kalır, kelimeye
  // düşülerek eşleştirilir, bkz. submitBatch/getExistingResults).
  var lastCol = tab.getLastColumn();
  var header  = tab.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('row') === -1) {
    tab.getRange(1, lastCol + 1).setValue('row');
  }
}

// ─── Yardımcı: Tek kelime upsert ─────────────────────────────────
function upsertWord_(resultsTab, spreadsheetId, worker, task, kelime, harf, type, row) {
  var sheetRowCol = ensureSheetRowIndex_(resultsTab);
  SpreadsheetApp.flush();
  var map = headerMap_(resultsTab);
  var rowKey = (row !== undefined && row !== null) ? String(row) : null;
  var found = queryWorkerTask_(spreadsheetId, resultsTab, worker, task);
  var existingRow = 0;
  var healedFromKelime = false;
  for (var i = 0; i < found.length; i++) {
    var r = found[i];
    if (!r.sheetRow) continue;
    if (rowKey !== null && r.dictRow !== '' && r.dictRow != null && String(r.dictRow) === rowKey) {
      existingRow = r.sheetRow;
      break;
    }
    if (!existingRow && (r.dictRow === '' || r.dictRow == null) && r.kelime === kelime) {
      existingRow = r.sheetRow;
      healedFromKelime = true;
    }
  }
  if (existingRow) {
    resultsTab.getRange(existingRow, map.score).setValue(type);
    resultsTab.getRange(existingRow, map.timestamp).setValue(new Date());
    if (healedFromKelime && rowKey !== null) resultsTab.getRange(existingRow, map.row).setValue(row);
    return;
  }
  var lastRow = resultsTab.getLastRow() + 1;
  resultsTab.getRange(lastRow, 1, 1, 7).setValues([[worker, task, kelime, harf, type, new Date(), row]]);
  resultsTab.getRange(lastRow, sheetRowCol).setValue(lastRow);
}

// ─── Yardımcı: Tasks sekmesinden task config bul ──────────────────
function findTask(tasksTab, task) {
  var data = tasksTab.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === task) {
      return {
        start:  parseInt(data[i][1]),
        end:    parseInt(data[i][2]),
        active: data[i][3] === true || String(data[i][3]).toUpperCase() === 'TRUE'
      };
    }
  }
  return null;
}

// ─── Yardımcı: Worker'ın bu task'taki mevcut sonuçlarını çek ──────
// Aynı kelime bir task içinde tekrar edebiliyor; bu yüzden asıl anahtar
// 'row' (dictionary'deki satır no). 'row' bilgisi olmayan (v8 öncesi) eski
// kayıtlar için kelimeye düşülen bir fallback map de dönülüyor.
function getExistingResults(resSS, lang, worker, task) {
  var byRow = {}, byKelime = {};
  var tab = resSS.getSheetByName(lang + '_Results');
  if (!tab || tab.getLastRow() < 2) return { byRow: byRow, byKelime: byKelime };

  queryWorkerTask_(resSS.getId(), tab, worker, task).forEach(function(r) {
    if (r.dictRow !== '' && r.dictRow != null) byRow[String(r.dictRow)] = r.score;
    else byKelime[r.kelime] = r.score;
  });
  return { byRow: byRow, byKelime: byKelime };
}

// Sonuç sayfasının tamamını scripte taşımak yerine Sheets sorgusu.
// Sadece bu worker + task satırları gelir.
function queryWorkerTask_(spreadsheetId, tab, worker, task) {
  var map = headerMap_(tab);
  ['worker_email', 'task_id', 'kelime', 'score', 'row'].forEach(function(name) {
    if (!map[name]) throw new Error('Kolon yok: ' + name);
  });
  var select = [colLetter_(map.row), colLetter_(map.score), colLetter_(map.kelime)];
  if (map.sheet_row) select.push(colLetter_(map.sheet_row));
  var tq = 'select ' + select.join(', ')
    + ' where ' + colLetter_(map.worker_email) + ' = ' + sqlQuote_(worker)
    + ' and ' + colLetter_(map.task_id) + ' = ' + sqlQuote_(task);
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId
    + '/gviz/tq?tqx=out:json&sheet=' + encodeURIComponent(tab.getName())
    + '&tq=' + encodeURIComponent(tq)
    + '&cb=' + Date.now();
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  var text = res.getContentText();
  if (res.getResponseCode() !== 200) throw new Error('Sonuç sorgusu HTTP ' + res.getResponseCode());
  var start = text.indexOf('{');
  var end = text.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('Sonuç sorgusu okunamadı');
  var data = JSON.parse(text.substring(start, end + 1));
  if (data.status !== 'ok') {
    var err = data.errors && data.errors[0];
    throw new Error('Sonuç sorgusu: ' + ((err && (err.detailed_message || err.message)) || 'bilinmeyen'));
  }
  var rows = (data.table && data.table.rows) || [];
  return rows.map(function(row) {
    var c = row.c || [];
    function v(idx) { return (c[idx] && c[idx].v != null) ? c[idx].v : ''; }
    return {
      dictRow: v(0),
      score: parseInt(v(1), 10),
      kelime: String(v(2)),
      sheetRow: map.sheet_row ? Number(v(3)) : 0
    };
  });
}

function headerMap_(tab) {
  var lastCol = Math.max(tab.getLastColumn(), 1);
  var header = tab.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < header.length; i++) {
    if (header[i]) map[String(header[i])] = i + 1;
  }
  return map;
}

function colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function sqlQuote_(s) {
  return "'" + String(s).replace(/'/g, "''") + "'";
}

// sheet_row: satırın kendi numarası. Güncelleme bu numarayla yapılır,
// bütün sayfa okunmaz. ARRAYFORMULA kullanılmaz; getLastRow()'u şişirirdi.
// Kolon doluysa tek hücre kontrolü, boşsa bir kerelik yazılır.
function ensureSheetRowIndex_(tab) {
  var map = headerMap_(tab);
  var col = map.sheet_row;
  if (!col) {
    col = tab.getLastColumn() + 1;
    tab.getRange(1, col).setValue('sheet_row');
  }
  var last = tab.getLastRow();
  if (last < 2) return col;
  if (Number(tab.getRange(last, col).getValue()) === last) return col;
  var CHUNK = 8000;
  for (var start = 2; start <= last; start += CHUNK) {
    var n = Math.min(CHUNK, last - start + 1);
    var vals = [];
    for (var i = 0; i < n; i++) vals.push([start + i]);
    tab.getRange(start, col, n, 1).setValues(vals);
  }
  return col;
}
