// ─── Spreadsheet ID'leri ───────────────────────────────────────────
var DICT_ID = '1Np2b7Im5q55kXHRhPc3euvlPGiZ_7anTfvYSly1hvHQ';  // FTW Dictionaries
var RESULTS = {
  'TR': '1B5F24YxM3aOedWoiB7oJ1xcO2MucRcLmwdG0pFODCGQ',         // FTW Results TR
  'EN': '1lg-u01ymZZB0nK0GOnYFt3cq6tZ39Z1LL36bKsMCM18',         // FTW Results EN
  'RU': '139sob08OmOENQfvJyXYIRbgzmlUVwzrGMl9sMBfXocA'          // FTW Results RU
};

// ─── GET router ───────────────────────────────────────────────────
function doGet(e) {
  var p = e.parameter;
  var result;
  try {
    if      (p.action === 'getWords') result = getWords(p.task || '', p.worker || '');
    else if (p.action === 'submit')   result = submitWord(p);
    else                              result = { error: 'Unknown action: ' + p.action };
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
    words.push({
      row:         taskConfig.start + idx,
      kelime:      kelime,
      harf_sayisi: (r[1] != null && r[1] !== '') ? r[1] : kelime.length,
      score:       (r[2] != null && r[2] !== '') ? r[2] : 0,
      type:        done[kelime] !== undefined ? done[kelime] : null
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
    upsertWord_(resultsTab, worker, task, kelime, p.harf_sayisi || kelime.length, type);
    return { ok: true };
  } finally {
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

    // Mevcut tüm veriyi bir kere oku (her kelime için tekrar okumamak adına)
    var data   = resultsTab.getDataRange().getValues();
    var h      = data[0];
    var iW     = h.indexOf('worker_email');
    var iT     = h.indexOf('task_id');
    var iK     = h.indexOf('kelime');
    var iScore = h.indexOf('score');
    var iTime  = h.indexOf('timestamp');

    // Mevcut satırları indexle: kelime → satır numarası
    var existingRows = {};
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iW]) === worker && String(data[i][iT]).toUpperCase() === task) {
        existingRows[String(data[i][iK])] = i + 1; // 1-indexed sheet row
      }
    }

    var now = new Date();
    var newRows = [];
    var updatedCount = 0;

    words.forEach(function(w) {
      var kelime = String(w.kelime || '').trim();
      var type   = parseInt(w.type);
      if (!kelime || isNaN(type)) return;

      var harf = w.harf_sayisi || kelime.length;
      var existingRow = existingRows[kelime];

      if (existingRow) {
        // Mevcut satırı güncelle
        resultsTab.getRange(existingRow, iScore + 1).setValue(type);
        resultsTab.getRange(existingRow, iTime + 1).setValue(now);
        updatedCount++;
      } else {
        // Yeni satır olarak eklenecekler listesine ekle
        newRows.push([worker, task, kelime, harf, type, now]);
      }
    });

    // Yeni satırları toplu ekle (appendRow yerine setValues — çok daha hızlı)
    if (newRows.length > 0) {
      var lastRow = resultsTab.getLastRow();
      resultsTab.getRange(lastRow + 1, 1, newRows.length, 6).setValues(newRows);
    }

    return { ok: true, saved: words.length, updated: updatedCount, inserted: newRows.length };
  } finally {
    lock.releaseLock();
  }
}

// ─── Yardımcı: Header satırını garanti et ─────────────────────────
function ensureHeader_(tab) {
  if (tab.getLastRow() === 0) {
    tab.appendRow(['worker_email', 'task_id', 'kelime', 'harf_sayisi', 'score', 'timestamp']);
  }
}

// ─── Yardımcı: Tek kelime upsert ─────────────────────────────────
function upsertWord_(resultsTab, worker, task, kelime, harf, type) {
  var data  = resultsTab.getDataRange().getValues();
  var h     = data[0];
  var iW    = h.indexOf('worker_email');
  var iT    = h.indexOf('task_id');
  var iK    = h.indexOf('kelime');
  var iScore = h.indexOf('score');
  var iTime = h.indexOf('timestamp');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iW]) === worker &&
        String(data[i][iT]).toUpperCase() === task &&
        String(data[i][iK]) === kelime) {
      resultsTab.getRange(i + 1, iScore + 1).setValue(type);
      resultsTab.getRange(i + 1, iTime + 1).setValue(new Date());
      return;
    }
  }
  resultsTab.appendRow([worker, task, kelime, harf, type, new Date()]);
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
function getExistingResults(resSS, lang, worker, task) {
  var map = {};
  var tab = resSS.getSheetByName(lang + '_Results');
  if (!tab || tab.getLastRow() < 2) return map;

  var data  = tab.getDataRange().getValues();
  var h     = data[0];
  var iW     = h.indexOf('worker_email');
  var iT     = h.indexOf('task_id');
  var iK     = h.indexOf('kelime');
  var iScore = h.indexOf('score');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iW]) === worker && String(data[i][iT]).toUpperCase() === task) {
      map[String(data[i][iK])] = parseInt(data[i][iScore]);
    }
  }
  return map;
}
