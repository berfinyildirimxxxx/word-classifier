// ─── Spreadsheet ID'leri ───────────────────────────────────────────
var DICT_ID = '1Np2b7Im5q55kXHRhPc3euvlPGiZ_7anTfvYSly1hvHQ';  // FTW Dictionaries
var RESULTS = {
  'TR': '1B5F24YxM3aOedWoiB7oJ1xcO2MucRcLmwdG0pFODCGQ',         // FTW Results TR
  'EN': '1lg-u01ymZZB0nK0GOnYFt3cq6tZ39Z1LL36bKsMCM18',         // FTW Results EN
  'RU': '139sob08OmOENQfvJyXYIRbgzmlUVwzrGMl9sMBfXocA'          // FTW Results RU
};

// ─── Ana router ───────────────────────────────────────────────────
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

// ─── Kelimeleri getir ─────────────────────────────────────────────
// task: "EN_001", "RU_005" gibi
function getWords(task, worker) {
  task = task.toUpperCase();               // "EN_001"
  var lang = task.split('_')[0];           // "EN"

  if (!RESULTS[lang]) return { error: 'Bilinmeyen dil: ' + lang };

  // 1. XX_Tasks sekmesinden bu task'ın start/end satırlarını bul
  var resSS    = SpreadsheetApp.openById(RESULTS[lang]);
  var tasksTab = resSS.getSheetByName(lang + '_Tasks');
  if (!tasksTab) return { error: lang + '_Tasks sekmesi bulunamadı' };

  var taskConfig = findTask(tasksTab, task);
  if (!taskConfig)        return { error: 'Task bulunamadı: ' + task };
  if (!taskConfig.active) return { error: 'Bu task aktif değil: ' + task };

  // 2. FTW Dictionaries'tan o dil sekmesini aç
  var dictSS  = SpreadsheetApp.openById(DICT_ID);
  var dictTab = dictSS.getSheetByName(lang);
  if (!dictTab) return { error: '"' + lang + '" sekmesi FTW Dictionaries\'ta bulunamadı.' };

  // start_row/end_row: 1-indexed, header hariç (1 = ilk kelime)
  // Sheet'te satır 1 header, ilk kelime satır 2 → +1 offset
  var count   = taskConfig.end - taskConfig.start + 1;
  var rawData = dictTab.getRange(taskConfig.start + 1, 1, count, dictTab.getLastColumn()).getValues();

  // 3. Bu worker'ın bu task'ta daha önce yaptıklarını yükle
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

// ─── Sınıflandırmayı kaydet ───────────────────────────────────────
function submitWord(p) {
  var task   = String(p.task   || '').toUpperCase().trim();
  var worker = String(p.worker || '').trim();
  var kelime = String(p.kelime || '').trim();
  var type   = parseInt(p.type);

  if (!task || !worker || !kelime || isNaN(type)) return { error: 'Eksik parametre' };

  var lang = task.split('_')[0];
  if (!RESULTS[lang]) return { error: 'Bilinmeyen dil: ' + lang };

  var resSS      = SpreadsheetApp.openById(RESULTS[lang]);
  var resultsTab = resSS.getSheetByName(lang + '_Results');
  if (!resultsTab) return { error: lang + '_Results sekmesi bulunamadı' };

  // Header yoksa ilk satıra yaz
  if (resultsTab.getLastRow() === 0) {
    resultsTab.appendRow(['worker_email', 'task_id', 'kelime', 'harf_sayisi', 'score', 'timestamp']);
  }

  // Aynı worker + task + kelime varsa sadece score ve timestamp güncelle
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
      return { ok: true };
    }
  }

  // Yeni satır ekle
  resultsTab.appendRow([worker, task, kelime, p.harf_sayisi || '', type, new Date()]);
  return { ok: true };
}

// ─── Yardımcı: Tasks sekmesinden task config bul ──────────────────
// Beklenen header: task_id | start_row | end_row | active
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
