/**
 * PORSENIJAR 2026 — Dashboard Peserta PGRI Kabupaten Sampang
 * Google Apps Script Web App backend.
 *
 * SETUP:
 * 1) Isi GOOGLE_SHEET_ID dan SHEET_NAME di bawah.
 * 2) Pastikan baris pertama sheet berisi header kolom.
 * 3) Deploy sebagai Web app. Jalankan sebagai pemilik deployment.
 * 4) Untuk data peserta pribadi, batasi akses web app kepada pengguna yang berwenang.
 */

const CONFIG = {
  GOOGLE_SHEET_ID: '1F6LcavHrY6ZGg5bevdDcMgLDC4xA4YyXcENY5yBXzwU',
  SHEET_NAME: 'Form Responses 1',

  // Ubah nama header di bawah bila header Google Form berbeda.
  // Sistem juga mencoba mencocokkan beberapa variasi nama secara otomatis.
  COLUMNS: {
    timestamp: 'Timestamp',       // opsional
    branch: 'PGRI Cabang',
    competition: 'Jenis Lomba',
    participant: 'Nama Peserta'
  },

  // Refresh dashboard di browser: 30 detik.
  // Cache server hanya untuk mengurangi beban; nilai 0 berarti tanpa cache.
  CACHE_SECONDS: 10,

  COMPETITIONS: [
    'Badminton',
    'Tenis Meja',
    'Catur',
    'Paduan Suara',
    'Karaoke',
    'Koreografi',
    'Seni Lukis',
    'Video Animasi Pembelajaran Mendalam dengan AI'
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PORSENIJAR 2026 | PGRI Kabupaten Sampang')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API utama yang dipanggil oleh index.html melalui google.script.run.
 * Mengembalikan data nyata dari Google Sheets, bukan data contoh.
 */
function getDashboardData(forceRefresh) {
  const startedAt = new Date();
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'PORSENIJAR_DASHBOARD_V1';
    if (!forceRefresh && CONFIG.CACHE_SECONDS > 0) {
      const cached = cache.get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        parsed.serverResponseAt = startedAt.toISOString();
        parsed.fromCache = true;
        return parsed;
      }
    }

    const spreadsheetId = String(CONFIG.GOOGLE_SHEET_ID || '').trim();
    if (!spreadsheetId || spreadsheetId === 'MASUKKAN_SPREADSHEET_ID_DI_SINI') {
      throw new Error('Konfigurasi belum lengkap: isi GOOGLE_SHEET_ID pada Code.gs.');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('Tab sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan. Periksa SHEET_NAME pada Code.gs.');
    }

    const values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length === 0) {
      return {
        success: true,
        rows: [],
        competitions: CONFIG.COMPETITIONS,
        source: ss.getName() + ' / ' + CONFIG.SHEET_NAME,
        fetchedAt: new Date().toISOString(),
        fromCache: false,
        warning: 'Sheet belum memiliki data.'
      };
    }

    const headers = values[0].map(v => normalizeHeader_(v));
    const indexes = resolveColumnIndexes_(headers);
    if (indexes.branch < 0 || indexes.competition < 0 || indexes.participant < 0) {
      throw new Error(
        'Kolom wajib tidak ditemukan. Header terbaca: ' + values[0].join(' | ') +
        '. Sesuaikan CONFIG.COLUMNS pada Code.gs.'
      );
    }

    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const line = values[i];
      const branch = cleanText_(line[indexes.branch]);
      const competitionRaw = cleanText_(line[indexes.competition]);
      const participant = cleanText_(line[indexes.participant]);
      const timestamp = indexes.timestamp >= 0 ? cleanText_(line[indexes.timestamp]) : '';

      // Abaikan baris yang sepenuhnya kosong dan baris tanpa data peserta.
      if (!branch && !competitionRaw && !participant) continue;
      if (!branch && !competitionRaw && !participant && !timestamp) continue;

      rows.push({
        timestamp: timestamp,
        branch: branch || '-',
        competition: normalizeCompetition_(competitionRaw),
        participant: participant || '-',
        sourceRow: i + 1
      });
    }

    const result = {
      success: true,
      rows: rows,
      competitions: CONFIG.COMPETITIONS,
      source: ss.getName() + ' / ' + CONFIG.SHEET_NAME,
      fetchedAt: new Date().toISOString(),
      fromCache: false,
      warning: ''
    };

    if (CONFIG.CACHE_SECONDS > 0) {
      // CacheService membatasi ukuran tiap item; kegagalan cache tidak boleh menggagalkan sinkronisasi.
      try {
        const serialized = JSON.stringify(result);
        if (serialized.length <= 90000) {
          cache.put(cacheKey, serialized, CONFIG.CACHE_SECONDS);
        }
      } catch (cacheError) {
        // Abaikan error cache; data langsung tetap dikembalikan ke dashboard.
      }
    }
    return result;
  } catch (err) {
    // Return a structured error for the UI; do not claim the sync succeeded.
    return {
      success: false,
      error: err && err.message ? err.message : String(err),
      fetchedAt: new Date().toISOString()
    };
  }
}

function resolveColumnIndexes_(headers) {
  const aliases = {
    timestamp: ['timestamp', 'stempel waktu', 'waktu', 'tanggal pengisian'],
    branch: ['pgricabang', 'cabangpgri', 'kecamatan', 'pgricabang/kecamatan', 'cabang'],
    competition: ['jenislomba', 'cabanglomba', 'lomba', 'kategori lomba', 'pilihan lomba'],
    participant: ['namapeserta', 'nama', 'namalengkap', 'namapeserta/kelompok', 'nama peserta/anggota']
  };

  function find(configured, key) {
    const wanted = normalizeHeader_(configured);
    let idx = headers.indexOf(wanted);
    if (idx >= 0) return idx;
    const choices = aliases[key] || [];
    idx = headers.findIndex(h => choices.indexOf(h) >= 0);
    return idx;
  }

  return {
    timestamp: find(CONFIG.COLUMNS.timestamp, 'timestamp'),
    branch: find(CONFIG.COLUMNS.branch, 'branch'),
    competition: find(CONFIG.COLUMNS.competition, 'competition'),
    participant: find(CONFIG.COLUMNS.participant, 'participant')
  };
}

function normalizeHeader_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function cleanText_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeCompetition_(value) {
  const raw = cleanText_(value);
  if (!raw) return '-';
  const key = raw.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

  const map = [
    { keys: ['badminton', 'bulu tangkis', 'bulutangkis'], name: 'Badminton' },
    { keys: ['tenismeja', 'pingpong'], name: 'Tenis Meja' },
    { keys: ['catur', 'chess'], name: 'Catur' },
    { keys: ['paduansuara', 'padusuara', 'choir'], name: 'Paduan Suara' },
    { keys: ['karaoke', 'menyanyi', 'vokal'], name: 'Karaoke' },
    { keys: ['koreografi', 'dance', 'tari'], name: 'Koreografi' },
    { keys: ['senilukis', 'melukis', 'lukis'], name: 'Seni Lukis' },
    { keys: ['videoanimasipembelajaranmendalamdenganai', 'videoanimasipembelajaranmendalamai',
      'animasipembelajaranmendalamdenganai', 'videopembelajaranai', 'videoanimasi', 'ai'], 
      name: 'Video Animasi Pembelajaran Mendalam dengan AI' }
  ];

  for (let i = 0; i < map.length; i++) {
    if (map[i].keys.indexOf(key) >= 0) return map[i].name;
  }
  // Match long category labels with punctuation/extra words.
  if (key.indexOf('badminton') >= 0 || key.indexOf('bulutangkis') >= 0) return 'Badminton';
  if (key.indexOf('tenismeja') >= 0) return 'Tenis Meja';
  if (key.indexOf('paduansuara') >= 0) return 'Paduan Suara';
  if (key.indexOf('koreografi') >= 0) return 'Koreografi';
  if (key.indexOf('senilukis') >= 0) return 'Seni Lukis';
  if (key.indexOf('karaoke') >= 0) return 'Karaoke';
  if (key.indexOf('catur') >= 0) return 'Catur';
  if (key.indexOf('video') >= 0 && key.indexOf('ai') >= 0) return 'Video Animasi Pembelajaran Mendalam dengan AI';
  return raw;
}

/** Utility untuk memeriksa apakah project dapat membaca spreadsheet. */
function testConnection() {
  const result = getDashboardData(true);
  Logger.log(JSON.stringify({
    success: result.success,
    source: result.source || '',
    rowCount: result.rows ? result.rows.length : 0,
    error: result.error || ''
  }, null, 2));
  return result;
}
