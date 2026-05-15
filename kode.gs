/**
 * WERUH
 * JAMU AMIS media
 * Jendela Aksi Menuju 
 * @elymulyaningsih
 */
const SPREADSHEET_ID = '1QbnhFDbowWkUZLcq7oFCd4O3hO-9Qgtfn_AG0W-swdg';
const PERTANYAAN_SHEET_NAME = 'Misi';
const PEMAIN_SHEET_NAME = 'Pemain';
const PERMAINAN_SHEET_NAME = 'Permainan';
const KREATOR_SHEET_NAME = 'Kreator';
const KREATOR_HEADERS = [
  'Email',
  'Nama Lengkap',
  'NPSN',
  'Nama Sekolah',
  'Kecamatan'
];
const PEMAIN_HEADERS = [
  'Email',
  'Nama Lengkap',
  'NPSN',
  'Nama Sekolah',
  'Kecamatan'
];
const PERMAINAN_HEADERS = [
  'Tanggal',
  'Email Pemain',
  'Nama Pemain',
  'Judul Misi',
  'Skor',
  'Daftar Jawaban'
];
const SEKOLAH_SHEET_NAME = 'Sekolah';
const PERTANYAAN_HEADERS = [
  'Mission ID',
  'Email Pembuat',
  'Judul Misi',
  'Sumber Konten',
  'Kelas',
  'Kategori',
  'Detik Berhenti',
  'Nomor Soal',
  'Pertanyaan',
  'Opsi Jawaban',
  'Kunci Jawaban',
  'Dibuat Pada',
  'Diubah Pada'
];
const LEGACY_PERTANYAAN_HEADERS = [
  'Judul',
  'Sumber',
  'Detik Berhenti',
  'Nomor Soal',
  'Pertanyaan',
  'Opsi Jawaban',
  'Kunci Jawaban',
  'Kelas',
  'Kategori',
  'Nama Pembuat',
  'Nama Sekolah',
  'Kecamatan Pembuat',
  'Mission ID',
  'Dibuat Pada',
  'Diubah Pada',
  'Dihapus Pada',
  'Email Pembuat'
];
const COL_MISSION_ID = 0;
const COL_EMAIL_PEMBUAT = 1;
const COL_JUDUL = 2;
const COL_SUMBER = 3;
const COL_KELAS = 4;
const COL_KATEGORI = 5;
const COL_DETIK = 6;
const COL_NOMOR = 7;
const COL_PERTANYAAN = 8;
const COL_OPSI = 9;
const COL_KUNCI = 10;
const COL_DIBUAT = 11;
const COL_DIUBAH = 12;

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e);
  }

  ensureKreatorSheet_();

  const requestedPage = e && e.parameter ? e.parameter.page : '';
  const allowedPages = ['index', 'kreator', 'game', 'daftar', 'login-pemain', 'daftar-pemain', 'dashboard-pemain'];
  const page = allowedPages.indexOf(requestedPage) !== -1 ? requestedPage : 'index';

  return HtmlService.createTemplateFromFile(page).evaluate()
      .setTitle('Markas Belajar Interaktif')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  return handleApiRequest_(e);
}

function handleApiRequest_(e) {
  const params = e && e.parameter ? e.parameter : {};
  const callback = params.callback;
  const response = executeApiAction_(params.action, getApiPayload_(e));
  const body = callback ? callback + '(' + JSON.stringify(response) + ');' : JSON.stringify(response);
  const mimeType = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(body).setMimeType(mimeType);
}

function getApiPayload_(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.payload) {
    try {
      return JSON.parse(params.payload);
    } catch(err) {
      return {};
    }
  }

  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch(err) {
      return {};
    }
  }

  return params;
}

function executeApiAction_(action, payload) {
  try {
    switch (action) {
      case 'saveData':
        return { success: true, data: saveData(payload) };
      case 'getMyMissions':
        return { success: true, data: getMyMissions(payload) };
      case 'getMissionDetail':
        return { success: true, data: getMissionDetail(payload) };
      case 'getMissionScores':
        return { success: true, data: getMissionScores(payload) };
      case 'getMissionById':
        return { success: true, data: getMissionById(payload) };
      case 'deleteMission':
        return { success: true, data: deleteMission(payload) };
      case 'getCreatorProfile':
        return { success: true, data: getCreatorProfile(payload) };
      case 'getKecamatanMissionDashboard':
        return { success: true, data: getKecamatanMissionDashboard() };
      case 'registerCreator':
        return { success: true, data: registerCreator(payload) };
      case 'registerPlayer':
        return { success: true, data: registerPlayer(payload) };
      case 'loginPlayer':
        return { success: true, data: loginPlayer(payload) };
      case 'getPlayerDashboard':
        return { success: true, data: getPlayerDashboard(payload) };
      case 'getSekolahByNpsn':
        return { success: true, data: getSekolahByNpsn(payload.npsn) };
      case 'getDaftarKelas':
        return { success: true, data: getDaftarKelas() };
      case 'getDaftarVideo':
        return { success: true, data: getDaftarVideo(payload.kelas || payload.kelasTerpilih) };
      case 'getSemuaMisi':
        return { success: true, data: getSemuaMisi() };
      case 'simpanJawaban':
        return { success: true, data: simpanJawaban(payload) };
      case 'getLeaderboard':
        return { success: true, data: getLeaderboard(payload.judulVideo, payload.idUser) };
      default:
        return { success: false, error: 'Action API tidak dikenal.' };
    }
  } catch(err) {
    return { success: false, error: err.message || String(err) };
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}

function getOrCreateSheet_(sheetName, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheetName === PERTANYAAN_SHEET_NAME) {
    migratePertanyaanSheetIfNeeded_(sheet);
  } else if (sheetName === KREATOR_SHEET_NAME) {
    migrateKreatorSheetIfNeeded_(sheet);
  } else if (sheetName === PEMAIN_SHEET_NAME) {
    migratePemainSheetIfNeeded_(sheet);
  }

  if (headers && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else if (headers) {
    ensureHeaders_(sheet, headers);
  }

  return sheet;
}

function ensureKreatorSheet_() {
  return getOrCreateSheet_(KREATOR_SHEET_NAME, KREATOR_HEADERS);
}

function ensurePemainSheet_() {
  return getOrCreateSheet_(PEMAIN_SHEET_NAME, PEMAIN_HEADERS);
}

function ensurePermainanSheet_() {
  return getOrCreateSheet_(PERMAINAN_SHEET_NAME, PERMAINAN_HEADERS);
}

function migratePemainSheetIfNeeded_(sheet) {
  if (sheet.getLastRow() === 0) return;

  const headerWidth = Math.max(sheet.getLastColumn(), PEMAIN_HEADERS.length, PERMAINAN_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const alreadyPemain = PEMAIN_HEADERS.every((header, index) => headers[index] === header);
  if (alreadyPemain) {
    trimExtraColumns_(sheet, PEMAIN_HEADERS.length);
    return;
  }

  const isOldPermainan = headers[0] === 'Tanggal'
    && headers[1] === 'Nama'
    && headers[2] === 'No Absen'
    && headers[3] === 'Judul Misi';
  if (!isOldPermainan) return;

  const ss = getSpreadsheet_();
  const target = ss.getSheetByName(PERMAINAN_SHEET_NAME) || ss.insertSheet(PERMAINAN_SHEET_NAME);
  if (target.getLastRow() === 0) target.appendRow(PERMAINAN_HEADERS);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const oldRows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const migratedRows = oldRows.map(row => [row[0], '', row[1], row[3], row[4], row[5]]);
    target.getRange(target.getLastRow() + 1, 1, migratedRows.length, PERMAINAN_HEADERS.length).setValues(migratedRows);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, PEMAIN_HEADERS.length).setValues([PEMAIN_HEADERS]);
  trimExtraColumns_(sheet, PEMAIN_HEADERS.length);
}

function migrateKreatorSheetIfNeeded_(sheet) {
  if (sheet.getLastRow() === 0) return;

  const headerWidth = Math.max(sheet.getLastColumn(), KREATOR_HEADERS.length, 4);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const alreadyNew = KREATOR_HEADERS.every((header, index) => headers[index] === header);
  if (alreadyNew) {
    trimExtraColumns_(sheet, KREATOR_HEADERS.length);
    return;
  }

  const isLegacy = headers[0] === 'Email'
    && headers[1] === 'Nama Lengkap'
    && headers[2] === 'Asal Sekolah'
    && headers[3] === 'Kecamatan';
  if (!isLegacy) return;

  const lastRow = sheet.getLastRow();
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 4).getValues()
    : [];

  const migratedValues = values.map(row => [
    row[0],
    row[1],
    '',
    row[2],
    row[3]
  ]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, KREATOR_HEADERS.length).setValues([KREATOR_HEADERS]);
  if (migratedValues.length) {
    sheet.getRange(2, 1, migratedValues.length, KREATOR_HEADERS.length).setValues(migratedValues);
  }
  trimExtraColumns_(sheet, KREATOR_HEADERS.length);
}

function migratePertanyaanSheetIfNeeded_(sheet) {
  if (sheet.getLastRow() === 0) return;

  const headerWidth = Math.max(sheet.getLastColumn(), LEGACY_PERTANYAAN_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const alreadyNew = PERTANYAAN_HEADERS.every((header, index) => headers[index] === header);
  if (alreadyNew) {
    trimExtraColumns_(sheet, PERTANYAAN_HEADERS.length);
    return;
  }

  const isLegacy = LEGACY_PERTANYAAN_HEADERS.every((header, index) => headers[index] === header);
  const isPreviousNew = headers[0] === 'Mission ID'
    && headers[1] === 'Email Pembuat'
    && headers[2] === 'Judul Misi'
    && headers[13] === 'Dihapus Pada';
  if (!isLegacy && !isPreviousNew) return;

  const lastRow = sheet.getLastRow();
  const sourceWidth = isLegacy ? LEGACY_PERTANYAAN_HEADERS.length : 14;
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues()
    : [];
  const migratedValues = values
    .filter(row => isLegacy ? !row[15] : !row[13])
    .map(row => isLegacy ? [
      row[12],
      row[16],
      row[0],
      row[1],
      row[7],
      row[8],
      row[2],
      row[3],
      row[4],
      row[5],
      row[6],
      row[13],
      row[14]
    ] : row.slice(0, PERTANYAAN_HEADERS.length));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, PERTANYAAN_HEADERS.length).setValues([PERTANYAAN_HEADERS]);
  if (migratedValues.length) {
    sheet.getRange(2, 1, migratedValues.length, PERTANYAAN_HEADERS.length).setValues(migratedValues);
  }
  trimExtraColumns_(sheet, PERTANYAAN_HEADERS.length);
}

function trimExtraColumns_(sheet, keepColumns) {
  const extraColumns = sheet.getMaxColumns() - keepColumns;
  if (extraColumns > 0) {
    sheet.deleteColumns(keepColumns + 1, extraColumns);
  }
}

function ensureHeaders_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  let changed = false;

  headers.forEach((header, index) => {
    if (existing[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
      changed = true;
    }
  });

  return changed;
}

function saveData(formObject) {
  try {
    ensureKreatorSheet_();
    const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
    const now = new Date();
    const missionId = cleanText_(formObject.missionId) || Utilities.getUuid();

    const row = [
      missionId,
      cleanEmail_(formObject.emailPembuat),
      cleanText_(formObject.judul),
      cleanText_(formObject.id),
      cleanText_(formObject.kelas),
      cleanText_(formObject.kategori),
      cleanText_(formObject.detik),
      cleanText_(formObject.nomor),
      cleanText_(formObject.pertanyaan),
      cleanText_(formObject.jawaban),
      cleanText_(formObject.kunci),
      now,
      now
    ];

    validateQuestionRow_(row);
    validateMissionMetadata_(row);

    const rowToUpdate = findMissionRow_(sheet, missionId);
    if (rowToUpdate > 1) {
      const existing = sheet.getRange(rowToUpdate, 1, 1, PERTANYAAN_HEADERS.length).getValues()[0];
      assertOwnerMatches_(existing, formObject);
      row[COL_DIBUAT] = existing[COL_DIBUAT] || now;
      sheet.getRange(rowToUpdate, 1, 1, row.length).setValues([row]);
      return { success: true, message: 'Misi berhasil diperbarui.', missionId: missionId };
    }

    sheet.appendRow(row);

    return { success: true, message: 'Misi berhasil disimpan.', missionId: missionId };
  } catch(e) {
    throw new Error(e.message || e);
  }
}

function cleanText_(value) {
  return value === null || value === undefined ? '' : value.toString().trim();
}

function cleanEmail_(value) {
  return cleanText_(value).toLowerCase();
}

function splitItems_(value, separator) {
  return cleanText_(value).split(separator).map(item => item.trim()).filter(String);
}

function validateQuestionRow_(row) {
  const requiredFields = [
    { index: COL_JUDUL, label: 'Judul Misi' },
    { index: COL_SUMBER, label: 'Sumber Konten' },
    { index: COL_KELAS, label: 'Kelas' },
    { index: COL_DETIK, label: 'Detik Berhenti' },
    { index: COL_NOMOR, label: 'Nomor Soal' },
    { index: COL_PERTANYAAN, label: 'Pertanyaan' },
    { index: COL_OPSI, label: 'Opsi Jawaban' },
    { index: COL_KUNCI, label: 'Kunci Jawaban' }
  ];

  requiredFields.forEach(field => {
    if (!row[field.index]) throw new Error(field.label + ' wajib diisi.');
  });

  const stops = splitItems_(row[COL_DETIK], ',');
  if (stops.some(stop => isNaN(Number(stop)))) {
    throw new Error('Detik Berhenti hanya boleh berisi angka yang dipisahkan koma.');
  }

  const nomor = splitItems_(row[COL_NOMOR], ',');
  const pertanyaan = splitItems_(row[COL_PERTANYAAN], '|');
  const opsi = splitItems_(row[COL_OPSI], '|');
  const kunci = splitItems_(row[COL_KUNCI], '|');

  if (nomor.length !== pertanyaan.length || pertanyaan.length !== opsi.length || pertanyaan.length !== kunci.length) {
    throw new Error('Jumlah Nomor Soal, Pertanyaan, Opsi Jawaban, dan Kunci Jawaban harus sama.');
  }
}

function validateMissionMetadata_(row) {
  const requiredFields = [
    { index: COL_KATEGORI, label: 'Kategori' },
    { index: COL_EMAIL_PEMBUAT, label: 'Email Pembuat' }
  ];

  requiredFields.forEach(field => {
    if (!row[field.index]) throw new Error(field.label + ' wajib diisi.');
  });

  if (!isBelajarEmail_(row[COL_EMAIL_PEMBUAT])) {
    throw new Error('Email Pembuat harus memakai domain .belajar.id.');
  }
}

function isBelajarEmail_(email) {
  return /^[^\s@]+@[a-z0-9.-]+\.belajar\.id$/i.test(cleanEmail_(email));
}

function ownerKey_(data) {
  return cleanEmail_(data.emailPembuat);
}

function rowOwnerKey_(row) {
  return cleanEmail_(row[COL_EMAIL_PEMBUAT]);
}

function assertOwnerMatches_(row, payload) {
  if (rowOwnerKey_(row) !== ownerKey_(payload)) {
    throw new Error('Misi ini hanya bisa diubah oleh pembuat yang sama.');
  }
}

function findMissionRow_(sheet, missionId) {
  if (!missionId || sheet.getLastRow() < 2) return -1;

  const ids = sheet.getRange(2, COL_MISSION_ID + 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (cleanText_(ids[i][0]) === missionId) return i + 2;
  }

  return -1;
}

function getSekolahSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SEKOLAH_SHEET_NAME);
  if (!sheet) throw new Error("Sheet 'Sekolah' tidak ditemukan.");
  return sheet;
}

function normalizeHeader_(value) {
  return cleanText_(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function findColumnIndex_(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const normalized = normalizeHeader_(headers[i]);
    if (candidates.some(candidate => normalized === normalizeHeader_(candidate))) return i;
  }
  return -1;
}

function getSekolahByNpsn(npsn) {
  const targetNpsn = cleanText_(npsn);
  if (!targetNpsn) throw new Error('NPSN wajib diisi.');

  const sheet = getSekolahSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('Data sekolah belum tersedia.');

  const headers = values[0];
  const npsnIndex = findColumnIndex_(headers, ['NPSN']);
  const namaIndex = findColumnIndex_(headers, ['Nama Sekolah', 'Nama Satuan Pendidikan', 'Nama']);
  const kecamatanIndex = findColumnIndex_(headers, ['Kecamatan', 'Nama Kecamatan']);

  if (npsnIndex < 0 || namaIndex < 0 || kecamatanIndex < 0) {
    throw new Error('Header sheet Sekolah harus memiliki NPSN, Nama Sekolah, dan Kecamatan.');
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (cleanText_(row[npsnIndex]) === targetNpsn) {
      return {
        npsn: targetNpsn,
        namaSekolah: cleanText_(row[namaIndex]),
        kecamatan: cleanText_(row[kecamatanIndex])
      };
    }
  }

  throw new Error('NPSN tidak ditemukan!');
}

function getMyMissions(payload) {
  ensureKreatorSheet_();
  const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const key = ownerKey_(payload);
  if (!key || !isBelajarEmail_(key) || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERTANYAAN_HEADERS.length).getValues();
  const playCounts = getMissionPlayCountsByTitle_();
  return data
      .filter(row => rowOwnerKey_(row) === key)
      .map(row => ({
        missionId: row[COL_MISSION_ID],
        judul: row[COL_JUDUL],
      videoId: row[COL_SUMBER] ? row[COL_SUMBER].toString() : '',
      id: row[COL_SUMBER] ? row[COL_SUMBER].toString() : '',
      source: row[COL_SUMBER] ? row[COL_SUMBER].toString() : '',
        sumber: row[COL_SUMBER] ? row[COL_SUMBER].toString() : '',
        kelas: row[COL_KELAS],
        kategori: row[COL_KATEGORI],
        jumlahDimainkan: playCounts[cleanText_(row[COL_JUDUL])] || 0,
        updatedAt: row[COL_DIUBAH] ? new Date(row[COL_DIUBAH]).toISOString() : ''
      }))
      .sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

function getMissionDetail(payload) {
  ensureKreatorSheet_();
  const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const rowNumber = findMissionRow_(sheet, cleanText_(payload.missionId));
  if (rowNumber < 2) throw new Error('Misi tidak ditemukan.');

  const row = sheet.getRange(rowNumber, 1, 1, PERTANYAAN_HEADERS.length).getValues()[0];
  assertOwnerMatches_(row, payload);

  return {
    missionId: row[COL_MISSION_ID],
    judul: row[COL_JUDUL],
    id: row[COL_SUMBER],
    detik: row[COL_DETIK],
    nomor: row[COL_NOMOR],
    pertanyaan: row[COL_PERTANYAAN],
    jawaban: row[COL_OPSI],
    kunci: row[COL_KUNCI],
    kelas: row[COL_KELAS],
    kategori: row[COL_KATEGORI],
    emailPembuat: row[COL_EMAIL_PEMBUAT]
  };
}

function getMissionScores(payload) {
  ensureKreatorSheet_();
  const missionSheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const rowNumber = findMissionRow_(missionSheet, cleanText_(payload.missionId));
  if (rowNumber < 2) throw new Error('Misi tidak ditemukan.');

  const missionRow = missionSheet.getRange(rowNumber, 1, 1, PERTANYAAN_HEADERS.length).getValues()[0];
  assertOwnerMatches_(missionRow, payload);

  const missionTitle = cleanText_(missionRow[COL_JUDUL]);
  const gameSheet = ensurePermainanSheet_();
  const scores = [];

  if (gameSheet.getLastRow() >= 2) {
    const rows = gameSheet.getRange(2, 1, gameSheet.getLastRow() - 1, PERMAINAN_HEADERS.length).getValues();
    rows.forEach(row => {
      if (cleanText_(row[3]) !== missionTitle) return;

      scores.push({
        tanggal: row[0] ? new Date(row[0]).toISOString() : '',
        emailPemain: cleanEmail_(row[1]),
        namaPemain: cleanText_(row[2]),
        skor: Number(row[4]) || 0
      });
    });
  }

  scores.sort((a, b) => {
    if (b.skor !== a.skor) return b.skor - a.skor;
    return (a.tanggal || '').localeCompare(b.tanggal || '');
  });

  return {
    missionId: cleanText_(missionRow[COL_MISSION_ID]),
    judul: missionTitle,
    totalPemain: scores.length,
    scores
  };
}

function getMissionById(payload) {
  const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const rowNumber = findMissionRow_(sheet, cleanText_(payload.missionId));
  if (rowNumber < 2) throw new Error('Misi tidak ditemukan.');

  const row = sheet.getRange(rowNumber, 1, 1, PERTANYAAN_HEADERS.length).getValues()[0];
  return formatMissionForGame_(row);
}

function formatMissionForGame_(row) {
  return {
    missionId: row[COL_MISSION_ID],
    judul: row[COL_JUDUL],
    videoId: row[COL_SUMBER] ? row[COL_SUMBER].toString() : '',
    stops: row[COL_DETIK] ? row[COL_DETIK].toString().split(',').map(Number) : [],
    pertanyaan: row[COL_PERTANYAAN] ? row[COL_PERTANYAAN].toString().split('|').map(p => p.trim()) : [],
    opsi: row[COL_OPSI] ? row[COL_OPSI].toString().split('|').map(o => o.split(',').map(opt => opt.trim())) : [],
    kunci: row[COL_KUNCI] ? row[COL_KUNCI].toString().split('|').map(k => k.trim()) : []
  };
}

function deleteMission(payload) {
  ensureKreatorSheet_();
  const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const rowNumber = findMissionRow_(sheet, cleanText_(payload.missionId));
  if (rowNumber < 2) throw new Error('Misi tidak ditemukan.');

  const row = sheet.getRange(rowNumber, 1, 1, PERTANYAAN_HEADERS.length).getValues()[0];
  assertOwnerMatches_(row, payload);

  sheet.deleteRow(rowNumber);
  return { success: true, message: 'Misi berhasil dihapus.' };
}

function registerCreator(payload) {
  const sheet = ensureKreatorSheet_();
  const sekolah = getSekolahByNpsn(payload.npsn);
  const creator = {
    email: cleanEmail_(payload.email || payload.emailPembuat),
    namaLengkap: cleanText_(payload.namaLengkap),
    npsn: cleanText_(payload.npsn),
    namaSekolah: sekolah.namaSekolah,
    kecamatan: sekolah.kecamatan
  };

  validateCreator_(creator);

  const row = [
    creator.email,
    creator.namaLengkap,
    creator.npsn,
    creator.namaSekolah,
    creator.kecamatan
  ];
  const rowToUpdate = findCreatorRow_(sheet, creator.email);

  if (rowToUpdate > 1) {
    throw new Error('EMAIL_ALREADY_REGISTERED');
  }

  sheet.appendRow(row);
  return { success: true, message: 'Pendaftaran kreator berhasil disimpan.' };
}

function validateCreator_(creator) {
  if (!creator.email) throw new Error('Email wajib diisi.');
  if (!isBelajarEmail_(creator.email)) throw new Error('Email harus memakai domain .belajar.id.');
  if (!creator.namaLengkap) throw new Error('Nama Lengkap wajib diisi.');
  if (!creator.npsn) throw new Error('NPSN wajib diisi.');
  if (!creator.namaSekolah) throw new Error('Nama Sekolah wajib diisi.');
  if (!creator.kecamatan) throw new Error('Kecamatan wajib diisi.');
}

function getCreatorProfile(payload) {
  const sheet = ensureKreatorSheet_();
  const email = cleanEmail_(payload.email || payload.emailPembuat);
  if (!email) throw new Error('Email kreator wajib diisi.');

  const rowNumber = findCreatorRow_(sheet, email);
  if (rowNumber < 2) throw new Error('CREATOR_NOT_REGISTERED');

  const row = sheet.getRange(rowNumber, 1, 1, KREATOR_HEADERS.length).getValues()[0];
  return {
    email: cleanEmail_(row[0]),
    namaLengkap: cleanText_(row[1]),
    npsn: cleanText_(row[2]),
    namaSekolah: cleanText_(row[3]),
    kecamatan: cleanText_(row[4])
  };
}

function registerPlayer(payload) {
  const sheet = ensurePemainSheet_();
  const sekolah = getSekolahByNpsn(payload.npsn);
  const player = {
    email: cleanEmail_(payload.email),
    namaLengkap: cleanText_(payload.namaLengkap),
    npsn: cleanText_(payload.npsn),
    namaSekolah: sekolah.namaSekolah,
    kecamatan: sekolah.kecamatan
  };

  validatePlayer_(player);
  if (findPlayerRow_(sheet, player.email) > 1) throw new Error('EMAIL_ALREADY_REGISTERED');

  sheet.appendRow([
    player.email,
    player.namaLengkap,
    player.npsn,
    player.namaSekolah,
    player.kecamatan
  ]);
  return { success: true, message: 'Pendaftaran pemain berhasil disimpan.' };
}

function loginPlayer(payload) {
  const sheet = ensurePemainSheet_();
  const email = cleanEmail_(payload.email);
  if (!email) throw new Error('Email wajib diisi.');
  if (!isBelajarEmail_(email)) throw new Error('Email harus memakai domain .belajar.id.');

  const rowNumber = findPlayerRow_(sheet, email);
  if (rowNumber < 2) throw new Error('PLAYER_NOT_REGISTERED');

  const row = sheet.getRange(rowNumber, 1, 1, PEMAIN_HEADERS.length).getValues()[0];
  return {
    email: cleanEmail_(row[0]),
    namaLengkap: cleanText_(row[1]),
    npsn: cleanText_(row[2]),
    namaSekolah: cleanText_(row[3]),
    kecamatan: cleanText_(row[4])
  };
}

function getPlayerDashboard(payload) {
  const playerSheet = ensurePemainSheet_();
  const email = cleanEmail_(payload.email || payload.emailPemain || payload.idUser);
  if (!email) throw new Error('Email pemain wajib diisi.');

  const rowNumber = findPlayerRow_(playerSheet, email);
  if (rowNumber < 2) throw new Error('PLAYER_NOT_REGISTERED');

  const playerRow = playerSheet.getRange(rowNumber, 1, 1, PEMAIN_HEADERS.length).getValues()[0];
  const player = {
    email: cleanEmail_(playerRow[0]),
    namaLengkap: cleanText_(playerRow[1]),
    npsn: cleanText_(playerRow[2]),
    namaSekolah: cleanText_(playerRow[3]),
    kecamatan: cleanText_(playerRow[4])
  };

  const missionSources = getMissionSourcesByTitle_();

  const gameSheet = ensurePermainanSheet_();
  const history = [];
  if (gameSheet.getLastRow() >= 2) {
    const rows = gameSheet.getRange(2, 1, gameSheet.getLastRow() - 1, PERMAINAN_HEADERS.length).getValues();
    rows.forEach(row => {
      if (cleanEmail_(row[1]) !== email) return;

      history.push({
        tanggal: row[0] ? new Date(row[0]).toISOString() : '',
        namaPemain: cleanText_(row[2]),
        judulMisi: cleanText_(row[3]),
        videoId: missionSources[cleanText_(row[3])] || '',
        skor: Number(row[4]) || 0,
        daftarJawaban: cleanText_(row[5])
      });
    });
  }

  history.sort((a, b) => (b.tanggal || '').localeCompare(a.tanggal || ''));

  const totalMisi = new Set(history.map(item => item.judulMisi).filter(Boolean)).size;
  const skorTertinggi = history.reduce((max, item) => Math.max(max, item.skor), 0);
  const rataRataSkor = history.length
    ? Math.round(history.reduce((sum, item) => sum + item.skor, 0) / history.length)
    : 0;

  return {
    player,
    summary: {
      totalMain: history.length,
      totalMisi,
      skorTertinggi,
      rataRataSkor
    },
    history
  };
}

function getMissionSourcesByTitle_() {
  const sheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const sources = {};
  if (sheet.getLastRow() < 2) return sources;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERTANYAAN_HEADERS.length).getValues();
  rows.forEach(row => {
    const title = cleanText_(row[COL_JUDUL]);
    if (title && !sources[title]) sources[title] = cleanText_(row[COL_SUMBER]);
  });

  return sources;
}

function validatePlayer_(player) {
  if (!player.email) throw new Error('Email wajib diisi.');
  if (!isBelajarEmail_(player.email)) throw new Error('Email harus memakai domain .belajar.id.');
  if (!player.namaLengkap) throw new Error('Nama Lengkap wajib diisi.');
  if (!player.npsn) throw new Error('NPSN wajib diisi.');
  if (!player.namaSekolah) throw new Error('Nama Sekolah wajib diisi.');
  if (!player.kecamatan) throw new Error('Kecamatan wajib diisi.');
}

function findPlayerRow_(sheet, email) {
  if (!email || sheet.getLastRow() < 2) return -1;

  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (cleanEmail_(emails[i][0]) === email) return i + 2;
  }

  return -1;
}

function findCreatorRow_(sheet, email) {
  if (!email || sheet.getLastRow() < 2) return -1;

  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (cleanEmail_(emails[i][0]) === email) return i + 2;
  }

  return -1;
}

function getCreatorNamesByEmail_() {
  const sheet = ensureKreatorSheet_();
  const names = {};
  if (sheet.getLastRow() < 2) return names;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, KREATOR_HEADERS.length).getValues();
  rows.forEach(row => {
    const email = cleanEmail_(row[0]);
    if (email) {
      names[email] = {
        namaLengkap: cleanText_(row[1]),
        namaSekolah: cleanText_(row[3]),
        kecamatan: cleanText_(row[4])
      };
    }
  });

  return names;
}

function getKecamatanMissionDashboard() {
  ensureKreatorSheet_();
  const missionSheet = getOrCreateSheet_(PERTANYAAN_SHEET_NAME, PERTANYAAN_HEADERS);
  const creatorInfo = getCreatorNamesByEmail_();
  const counts = getAllKecamatanCounts_();
  const creatorsByDistrict = {};

  if (missionSheet.getLastRow() >= 2) {
    const rows = missionSheet.getRange(2, 1, missionSheet.getLastRow() - 1, PERTANYAAN_HEADERS.length).getValues();
    rows.forEach(row => {
      if (!row[COL_MISSION_ID] || !row[COL_JUDUL]) return;

      const ownerEmail = rowOwnerKey_(row);
      const creator = creatorInfo[ownerEmail] || {};
      const kecamatan = cleanText_(creator.kecamatan) || 'Tidak diketahui';
      if (counts[kecamatan] === undefined) counts[kecamatan] = 0;
      counts[kecamatan]++;

      if (!creatorsByDistrict[kecamatan]) creatorsByDistrict[kecamatan] = {};
      const creatorKey = ownerEmail || 'unknown';
      if (!creatorsByDistrict[kecamatan][creatorKey]) {
        creatorsByDistrict[kecamatan][creatorKey] = {
          email: ownerEmail || '-',
          namaLengkap: cleanText_(creator.namaLengkap) || ownerEmail || 'Tidak diketahui',
          namaSekolah: cleanText_(creator.namaSekolah) || '-',
          totalMisi: 0
        };
      }
      creatorsByDistrict[kecamatan][creatorKey].totalMisi++;
    });
  }

  return Object.keys(counts)
    .sort((a, b) => {
      if (counts[b] !== counts[a]) return counts[b] - counts[a];
      return a.localeCompare(b);
    })
    .map(kecamatan => ({
      kecamatan,
      totalMisi: counts[kecamatan],
      kreator: Object.keys(creatorsByDistrict[kecamatan] || {})
        .map(key => creatorsByDistrict[kecamatan][key])
        .sort((a, b) => {
          if (b.totalMisi !== a.totalMisi) return b.totalMisi - a.totalMisi;
          return a.namaLengkap.localeCompare(b.namaLengkap);
        })
    }));
}

function getAllKecamatanCounts_() {
  const counts = {};

  try {
    const sheet = getSekolahSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return counts;

    const kecamatanIndex = findColumnIndex_(values[0], ['Kecamatan', 'Nama Kecamatan']);
    if (kecamatanIndex < 0) return counts;

    values.slice(1).forEach(row => {
      const kecamatan = cleanText_(row[kecamatanIndex]);
      if (kecamatan && counts[kecamatan] === undefined) counts[kecamatan] = 0;
    });
  } catch (err) {
    return counts;
  }

  return counts;
}

function getMissionPlayCountsByTitle_() {
  const sheet = getOrCreateSheet_(PERMAINAN_SHEET_NAME, PERMAINAN_HEADERS);
  const counts = {};
  if (sheet.getLastRow() < 2) return counts;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PERMAINAN_HEADERS.length).getValues();
  rows.forEach(row => {
    const title = cleanText_(row[3]);
    if (title) counts[title] = (counts[title] || 0) + 1;
  });

  return counts;
}

/**
 * MENGAMBIL DAFTAR KELAS
 */
function getDaftarKelas() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(PERTANYAAN_SHEET_NAME);
    if (!sheet) return ["Sheet 'Misi' tidak ada!"];
    
    const data = sheet.getDataRange().getValues();
    let kelasUnik = [];
    for (let i = 1; i < data.length; i++) {
      let kls = data[i][COL_KELAS] ? data[i][COL_KELAS].toString().trim() : "";
      if (kls && kelasUnik.indexOf(kls) === -1) kelasUnik.push(kls);
    }
    return kelasUnik.sort();
  } catch(e) { 
    return ["Error: " + e.message]; 
  }
}

/**
 * MENGAMBIL DATA VIDEO & PERTANYAAN
 */
function getDaftarVideo(kelasTerpilih) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(PERTANYAAN_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    
    return data.slice(1)
      .filter(r => r[COL_KELAS] && r[COL_KELAS].toString().trim() === kelasTerpilih)
      .map(formatMissionForGame_);
  } catch(e) { return []; }
}

function getSemuaMisi() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(PERTANYAAN_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const creatorNames = getCreatorNamesByEmail_();
    const playCounts = getMissionPlayCountsByTitle_();

    return data.slice(1)
      .filter(row => row[COL_MISSION_ID] && row[COL_JUDUL] && row[COL_SUMBER])
      .map(row => {
        const mission = formatMissionForGame_(row);
        mission.kelas = cleanText_(row[COL_KELAS]);
        mission.kategori = cleanText_(row[COL_KATEGORI]);
        const creatorInfo = creatorNames[rowOwnerKey_(row)] || {};
        mission.namaKreator = creatorInfo.namaLengkap || rowOwnerKey_(row) || '-';
        mission.namaSekolahKreator = creatorInfo.namaSekolah || '-';
        mission.kecamatanKreator = creatorInfo.kecamatan || '-';
        mission.dibuatPada = row[COL_DIBUAT] ? new Date(row[COL_DIBUAT]).toISOString() : '';
        mission.jumlahDimainkan = playCounts[cleanText_(row[COL_JUDUL])] || 0;
        return mission;
      })
      .sort((a, b) => (b.dibuatPada || '').localeCompare(a.dibuatPada || ''));
  } catch(e) { return []; }
}

function simpanJawaban(data) {
  try {
    let sheet = ensurePermainanSheet_();

    const values = sheet.getDataRange().getValues();
    let rowToUpdate = -1;
    let jawabanLama = "";
    const emailPemain = cleanEmail_(data.emailPemain || data.idUser);

    for (let i = 1; i < values.length; i++) {
      if (cleanEmail_(values[i][1]) === emailPemain && values[i][3] === data.judulVideo) {
        rowToUpdate = i + 1;
        jawabanLama = values[i][5] || ""; // Ambil jawaban yang sudah tersimpan sebelumnya
        break;
      }
    }

    // Gabungkan jawaban lama dengan yang baru agar tersimpan SEMUA
    // Format: "Soal 1: A | Soal 2: B"
    let jawabanBaru = jawabanLama ? jawabanLama + " , " + data.jawabanAsli : data.jawabanAsli;

    if (rowToUpdate !== -1) {
      sheet.getRange(rowToUpdate, 1).setValue(new Date());
      sheet.getRange(rowToUpdate, 5).setValue(data.skor); // Update Skor terbaru
      sheet.getRange(rowToUpdate, 6).setValue(jawabanBaru); // Simpan semua jawaban
    } else {
      sheet.appendRow([new Date(), emailPemain, data.nama, data.judulVideo, data.skor, data.jawabanAsli]);
    }
    return "Tersimpan";
  } catch(e) { return "Gagal: " + e.message; }
}

/** * MENGAMBIL PERINGKAT & TOP 10 (BERDASARKAN JUDUL MISI)
 */
function getLeaderboard(judulVideo, idUser) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(PERMAINAN_SHEET_NAME);
    if (!sheet) return { myRank: "-", total: 0, top10: [] };
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang header

    // 1. FILTER: Hanya ambil data yang judul misinya COCOK
    const targetUser = cleanEmail_(idUser);
    let filtered = data.filter(r => r[3] === judulVideo)
      .map(r => ({ 
        email: cleanEmail_(r[1]),
        nama: r[2], 
        skor: Number(r[4]) || 0,
        waktu: new Date(r[0]).getTime()
      }));

    // 2. SORTING: Skor tertinggi di atas, jika skor sama, yang lebih cepat selesai di atas
    filtered.sort((a, b) => {
      if (b.skor !== a.skor) return b.skor - a.skor;
      return a.waktu - b.waktu;
    });

    // 3. CARI POSISI USER: Cari di urutan ke berapa idUser berada
    let myRank = "-";
    for (let i = 0; i < filtered.length; i++) {
      if (filtered[i].email === targetUser) {
        myRank = i + 1;
        break;
      }
    }

    // 4. AMBIL TOP 10 KHUSUS MISI INI
    let top10 = filtered.slice(0, 10).map((r, i) => ({
      rank: i + 1,
      nama: r.nama,
      skor: r.skor,
      absen: r.email
    }));

    return { 
      myRank: myRank, 
      total: filtered.length, 
      top10: top10,
      judul: judulVideo 
    };
  } catch(e) { 
    return { myRank: "-", total: 0, top10: [] }; 
  }
}
