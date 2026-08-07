// ─── Spreadsheet ID'leri ───────────────────────────────────────────
var DICT_ID = '1Np2b7Im5q55kXHRhPc3euvlPGiZ_7anTfvYSly1hvHQ';  // FTW Dictionaries
var RESULTS = {
  'TR': '1B5F24YxM3aOedWoiB7oJ1xcO2MucRcLmwdG0pFODCGQ',         // FTW Results TR
  'EN': '1lg-u01ymZZB0nK0GOnYFt3cq6tZ39Z1LL36bKsMCM18',         // FTW Results EN
  'RU': '139sob08OmOENQfvJyXYIRbgzmlUVwzrGMl9sMBfXocA',         // FTW Results RU
  'ES': '13X9nZdrmiCfw_UxnqVbYo1D4KWba9j77YmOCOutRDjE',         // FTW Results ES
  'PT': '11nLTTSVbcxcYZATeUZoUn-PMfFLfr8nOU7sU8KreW-E'          // FTW Results PT
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
    upsertWord_(resultsTab, worker, task, kelime, p.harf_sayisi || kelime.length, type, row);
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

    // Mevcut tüm veriyi bir kere oku (her kelime için tekrar okumamak adına)
    var data   = resultsTab.getDataRange().getValues();
    var h      = data[0];
    var iW     = h.indexOf('worker_email');
    var iT     = h.indexOf('task_id');
    var iK     = h.indexOf('kelime');
    var iScore = h.indexOf('score');
    var iTime  = h.indexOf('timestamp');
    var iRow   = h.indexOf('row');

    // Mevcut satırları indexle. Aynı kelime bir task içinde tekrar edebiliyor,
    // bu yüzden birincil anahtar 'row' (dictionary'deki satır no). 'row' bilgisi
    // olmayan (v8 öncesi) eski kayıtlar geriye dönük uyumluluk için kelimeyle
    // eşleştirilir (bu eşleşme yalnızca bir kez kullanılır, bkz. delete altta).
    var existingByRow = {};
    var existingByKelime = {};
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iW]) === worker && String(data[i][iT]).toUpperCase() === task) {
        var storedRow = iRow !== -1 ? data[i][iRow] : '';
        if (storedRow !== '' && storedRow != null) {
          existingByRow[String(storedRow)] = i + 1; // 1-indexed sheet row
        } else {
          existingByKelime[String(data[i][iK])] = i + 1;
        }
      }
    }

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
function upsertWord_(resultsTab, worker, task, kelime, harf, type, row) {
  var data  = resultsTab.getDataRange().getValues();
  var h     = data[0];
  var iW    = h.indexOf('worker_email');
  var iT    = h.indexOf('task_id');
  var iK    = h.indexOf('kelime');
  var iScore = h.indexOf('score');
  var iTime = h.indexOf('timestamp');
  var iRow  = h.indexOf('row');
  var rowKey = (row !== undefined && row !== null) ? String(row) : null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iW]) !== worker || String(data[i][iT]).toUpperCase() !== task) continue;
    var storedRow  = iRow !== -1 ? data[i][iRow] : '';
    var hasRowInfo = storedRow !== '' && storedRow != null;
    var isMatch = hasRowInfo ? (rowKey !== null && String(storedRow) === rowKey)
                              : (String(data[i][iK]) === kelime);
    if (isMatch) {
      resultsTab.getRange(i + 1, iScore + 1).setValue(type);
      resultsTab.getRange(i + 1, iTime + 1).setValue(new Date());
      if (!hasRowInfo && rowKey !== null && iRow !== -1) {
        resultsTab.getRange(i + 1, iRow + 1).setValue(row); // eski kaydı row bilgisiyle tamamla
      }
      return;
    }
  }
  resultsTab.appendRow([worker, task, kelime, harf, type, new Date(), row]);
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

  var data  = tab.getDataRange().getValues();
  var h     = data[0];
  var iW     = h.indexOf('worker_email');
  var iT     = h.indexOf('task_id');
  var iK     = h.indexOf('kelime');
  var iScore = h.indexOf('score');
  var iRow   = h.indexOf('row');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iW]) === worker && String(data[i][iT]).toUpperCase() === task) {
      var score = parseInt(data[i][iScore]);
      var storedRow = iRow !== -1 ? data[i][iRow] : '';
      if (storedRow !== '' && storedRow != null) {
        byRow[String(storedRow)] = score;
      } else {
        byKelime[String(data[i][iK])] = score;
      }
    }
  }
  return { byRow: byRow, byKelime: byKelime };
}
