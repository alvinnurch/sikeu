/**
 * PLATFORM WORKSHOP TERINTEGRASI — Backend Google Apps Script
 * Islamidotco & Wahid Foundation
 *
 * Basis data  : Google Spreadsheet (sheet-sheet di bawah)
 * Cara pakai  : Extensions → Apps Script → tempel berkas ini → Deploy → Web app
 * Panduan     : lihat PANDUAN-DEPLOY.md
 *
 * ── Struktur sheet ────────────────────────────────────────────────
 *  Sesi            sesi_id | judul | deskripsi | narasumber | waktu | tempat | format |
 *                  aktif | materi_ditampilkan | urutan
 *  Materi          materi_id | sesi_id | jenis | judul | deskripsi | pengisi | dipakai |
 *                  terbuka | url | nama_berkas | meeting_id | passcode | prompt |
 *                  jenis_kuis | berbatas_waktu | mode_waktu | detik | boleh_ulang | kode_barcode | urutan
 *      jenis       : checkin | ppt | zoom | quiz | wordcloud | feedback | form
 *      mode_waktu  : per (per soal) | total (seluruh soal)
 *  Soal            soal_id | materi_id | pertanyaan | opsi_a | opsi_b | opsi_c | opsi_d |
 *                  jawaban (A-D) | urutan
 *  IsianFormulir   isian_id | materi_id | label | jenis | wajib | pilihan | urutan
 *      jenis       : short | long | number | choice | image
 *  Peserta         peserta_id | nama | kode | kelompok | instansi | hp | email
 *  Checkin         sesi_id | peserta_id | waktu | metode
 *  JawabanKuis     sesi_id | materi_id | peserta_id | skor | jawaban | waktu
 *  WordCloud       sesi_id | materi_id | peserta_id | kata | waktu
 *  Feedback        sesi_id | materi_id | peserta_id | emoji | catatan | waktu
 *  JawabanForm     sesi_id | materi_id | peserta_id | isian_id | nilai | nama_berkas | waktu
 *
 * ── Titik akhir (endpoint) ────────────────────────────────────────
 *  GET  ?action=db                     → seluruh data untuk halaman (sesi, materi, jawaban)
 *  GET  ?action=admin&password=…       → sama, termasuk daftar peserta
 *  POST {action:'login', name, code, group}
 *  POST {action:'checkin', sessionId, blockId, id, token}
 *  POST {action:'quiz', sessionId, blockId, id, answers:[…]}
 *  POST {action:'word', sessionId, blockId, id, word}
 *  POST {action:'feedback', sessionId, blockId, id, emoji, note}
 *  POST {action:'form', sessionId, blockId, id, values:{isian_id:nilai}, fileNames:{}}
 *  POST {action:'upload', sessionId, blockId, fieldId, id, fileName, mimeType, dataBase64}
 *  POST {action:'putDb', password, db}  → admin menyimpan seluruh susunan sesi & materi
 */

var ADMIN_PASSWORD = 'sayaadminnya';   // ← ganti bila perlu
var UPLOAD_FOLDER  = 'Workshop Unggahan';

var SS = SpreadsheetApp.getActiveSpreadsheet();

var SHEETS = {
  Sesi: ['sesi_id', 'judul', 'deskripsi', 'narasumber', 'waktu', 'tempat', 'format', 'aktif', 'materi_ditampilkan', 'urutan'],
  Materi: ['materi_id', 'sesi_id', 'jenis', 'judul', 'deskripsi', 'pengisi', 'dipakai', 'terbuka', 'url', 'nama_berkas', 'meeting_id', 'passcode', 'prompt', 'jenis_kuis', 'berbatas_waktu', 'mode_waktu', 'detik', 'boleh_ulang', 'kode_barcode', 'urutan'],
  Soal: ['soal_id', 'materi_id', 'pertanyaan', 'opsi_a', 'opsi_b', 'opsi_c', 'opsi_d', 'jawaban', 'urutan'],
  IsianFormulir: ['isian_id', 'materi_id', 'label', 'jenis', 'wajib', 'pilihan', 'urutan'],
  Peserta: ['peserta_id', 'nama', 'kode', 'kelompok', 'instansi', 'hp', 'email'],
  Checkin: ['sesi_id', 'peserta_id', 'waktu', 'metode'],
  JawabanKuis: ['sesi_id', 'materi_id', 'peserta_id', 'skor', 'jawaban', 'waktu'],
  WordCloud: ['sesi_id', 'materi_id', 'peserta_id', 'kata', 'waktu'],
  Feedback: ['sesi_id', 'materi_id', 'peserta_id', 'emoji', 'catatan', 'waktu'],
  JawabanForm: ['sesi_id', 'materi_id', 'peserta_id', 'isian_id', 'nilai', 'nama_berkas', 'waktu']
};

/* ══════════ MENU & PENYIAPAN ══════════ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Workshop')
    .addItem('1. Siapkan spreadsheet', 'setupSpreadsheet')
    .addItem('2. Isi contoh sesi', 'seedContoh')
    .addItem('Buat kode peserta yang kosong', 'generateCodes')
    .addSeparator()
    .addItem('Kosongkan semua jawaban', 'resetJawaban')
    .addToUi();
}

function setupSpreadsheet() {
  Object.keys(SHEETS).forEach(function (name) {
    var sh = SS.getSheetByName(name) || SS.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[name]);
      sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold').setBackground('#F3EDE5');
      sh.setFrozenRows(1);
    }
  });
  SpreadsheetApp.getUi().alert('Sheet siap. Lanjutkan dengan "2. Isi contoh sesi", lalu isi sheet Peserta.');
}

function seedContoh() {
  setupSpreadsheet();
  var sesi = SS.getSheetByName('Sesi');
  if (sesi.getLastRow() < 2) {
    sesi.appendRow(['s1', 'Literasi Digital & Moderasi Beragama', 'Workshop sehari untuk penggerak komunitas.', 'Dr. Nurul Hidayati, M.A.', 'Kamis, 17 September 2026 · 08.30–15.30', 'Aula Wahid Foundation, Jakarta', 'Tatap muka', true, '', 1]);
    sesi.appendRow(['s2', 'Narasi Damai di Media Sosial', 'Sesi daring lanjutan.', 'Ahmad Fauzi, S.Sos.', 'Sabtu, 26 September 2026 · 19.30–21.00', 'Zoom Meeting', 'Daring', false, '', 2]);
  }
  var mat = SS.getSheetByName('Materi');
  if (mat.getLastRow() < 2) {
    [['m0', 's1', 'checkin', 'Check in peserta', 'Pindai barcode di layar proyektor untuk mencatat kehadiran.', 'Panitia', true, false, '', '', '', '', '', '', false, '', '', false, 'K7M2P', 1],
     ['m1', 's1', 'ppt', 'Paparan narasumber', 'Materi pembuka sesi.', 'Dr. Nurul Hidayati, M.A.', true, false, '', 'Materi-Sesi.pptx', '', '', '', '', false, '', '', false, '', 2],
     ['m2', 's1', 'quiz', 'Pre test', 'Pemetaan pemahaman awal.', 'Tim fasilitator', true, false, '', '', '', '', '', 'pre', true, 'per', 30, true, '', 3],
     ['m3', 's1', 'wordcloud', 'Word cloud', 'Kata kunci peserta tampil di layar.', 'Fasilitator', true, false, '', '', '', '', 'Satu kata yang Anda ingat dari sesi ini?', '', false, '', '', false, '', 4],
     ['m4', 's1', 'quiz', 'Post test', 'Pengukuran setelah materi.', 'Tim fasilitator', true, false, '', '', '', '', '', 'post', true, 'total', 300, false, '', 5],
     ['m5', 's1', 'form', 'Rencana aksi', 'Rumuskan rencana tindak lanjut.', 'Fasilitator kelompok', true, false, '', '', '', '', '', '', false, '', '', false, '', 6],
     ['m6', 's1', 'feedback', 'Umpan balik sesi', 'Pilih ekspresi dan beri catatan.', 'Panitia', true, false, '', '', '', '', '', '', false, '', '', false, '', 7],
     ['m7', 's2', 'checkin', 'Check in peserta daring', 'Barcode ditampilkan di layar bersama.', 'Panitia', true, false, '', '', '', '', '', '', false, '', '', false, 'R4TQ8', 1],
     ['m8', 's2', 'zoom', 'Ruang Zoom sesi daring', 'Masuk 10 menit sebelum mulai.', 'Panitia', true, false, 'https://zoom.us/j/0000000000', '', '000 0000 0000', 'damai26', '', '', false, '', '', false, '', 2],
     ['m9', 's2', 'feedback', 'Umpan balik sesi daring', 'Pilih ekspresi yang mewakili pengalaman Anda.', 'Panitia', true, false, '', '', '', '', '', '', false, '', '', false, '', 3]
    ].forEach(function (r) { mat.appendRow(r); });
  }
  var soal = SS.getSheetByName('Soal');
  if (soal.getLastRow() < 2) {
    var q = [
      ['Moderasi beragama paling tepat dipahami sebagai…', 'Sikap tengah yang adil dan tidak berlebihan', 'Mencampur ajaran semua agama', 'Menghindari pembahasan agama', 'Mengikuti pendapat mayoritas', 'A'],
      ['Langkah pertama saat menerima informasi yang meragukan adalah…', 'Langsung meneruskan ke grup', 'Tabayun — memeriksa sumber aslinya', 'Menghapus tanpa membaca', 'Menilai dari jumlah pengirim', 'B'],
      ['Literasi digital mencakup kemampuan…', 'Mengoperasikan perangkat saja', 'Mengakses, menilai, dan memproduksi informasi secara etis', 'Menambah jumlah pengikut', 'Memakai aplikasi terbaru', 'B'],
      ['Ujaran kebencian di ruang digital sebaiknya direspons dengan…', 'Balasan serupa agar setara', 'Pelaporan kanal dan narasi alternatif yang santun', 'Diamkan selamanya', 'Menyebarkannya sebagai contoh', 'B'],
      ['Indikator konten hoaks yang paling kuat adalah…', 'Judul provokatif tanpa sumber yang dapat diverifikasi', 'Ditulis panjang', 'Banyak dibagikan', 'Memuat foto', 'A']
    ];
    ['m2', 'm4'].forEach(function (mid) {
      q.forEach(function (r, i) { soal.appendRow([mid + '-q' + (i + 1), mid, r[0], r[1], r[2], r[3], r[4], r[5], i + 1]); });
    });
  }
  var isi = SS.getSheetByName('IsianFormulir');
  if (isi.getLastRow() < 2) {
    [['i1', 'm5', 'Judul rencana aksi', 'short', true, '', 1],
     ['i2', 'm5', 'Uraian langkah', 'long', true, '', 2],
     ['i3', 'm5', 'Perkiraan penerima manfaat', 'number', false, '', 3],
     ['i4', 'm5', 'Waktu pelaksanaan', 'choice', true, 'Bulan ini,1–3 bulan,3–6 bulan', 4],
     ['i5', 'm5', 'Foto atau sketsa pendukung', 'image', false, '', 5]
    ].forEach(function (r) { isi.appendRow(r); });
  }
  var ps = SS.getSheetByName('Peserta');
  if (ps.getLastRow() < 2) {
    [['p1', 'Nama Peserta Contoh', 'WS-1001', 'Kelompok A', 'Komunitas Contoh', '08120000000', 'contoh@email.id']]
      .forEach(function (r) { ps.appendRow(r); });
  }
  SpreadsheetApp.getUi().alert('Contoh sesi terisi. Ganti isinya sesuai workshop Anda, lalu isi sheet Peserta.');
}

function generateCodes() {
  var sh = SS.getSheetByName('Peserta');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    if (!rows[i][0]) sh.getRange(i + 1, 1).setValue('p' + i + '-' + new Date().getTime());
    if (!rows[i][2]) sh.getRange(i + 1, 3).setValue('WS-' + Math.floor(1000 + Math.random() * 9000));
  }
}

function resetJawaban() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Kosongkan semua jawaban peserta?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ['Checkin', 'JawabanKuis', 'WordCloud', 'Feedback', 'JawabanForm'].forEach(function (n) {
    var sh = SS.getSheetByName(n);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
  ui.alert('Jawaban dikosongkan. Susunan sesi dan materi tetap.');
}

/* ══════════ UTILITAS ══════════ */

function rows_(name) {
  var sh = SS.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var head = data.shift();
  return data.filter(function (r) { return String(r[0]).length; }).map(function (r) {
    var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
}
function bool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1'; }
function head_(name) { var sh = SS.getSheetByName(name); return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }
function upsert_(name, keyCols, obj) {
  var sh = SS.getSheetByName(name), head = head_(name);
  var data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  var row = head.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  for (var i = 0; i < data.length; i++) {
    var same = keyCols.every(function (k) { return String(data[i][head.indexOf(k)]) === String(obj[k]); });
    if (same) { sh.getRange(i + 2, 1, 1, head.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}
function rewrite_(name, objs) {
  var sh = SS.getSheetByName(name), head = SHEETS[name];
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  if (!objs.length) return;
  var rows = objs.map(function (o) { return head.map(function (h) { return o[h] !== undefined && o[h] !== null ? o[h] : ''; }); });
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}
function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function jam_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH.mm'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function letters_() { return ['A', 'B', 'C', 'D']; }

/* ══════════ BACA: bentuk data siap dipakai halaman ══════════ */

function buildDb_(withIdentity) {
  var soal = rows_('Soal'), isian = rows_('IsianFormulir'), materi = rows_('Materi');
  var sesiRows = rows_('Sesi').sort(function (a, b) { return (a.urutan || 0) - (b.urutan || 0); });

  var sessions = sesiRows.map(function (s) {
    return {
      id: s.sesi_id, title: s.judul, desc: s.deskripsi, speaker: s.narasumber, time: s.waktu,
      venue: s.tempat, mode: s.format || 'Tatap muka', active: s.materi_ditampilkan || null,
      run: null,
      blocks: materi.filter(function (m) { return String(m.sesi_id) === String(s.sesi_id); })
        .sort(function (a, b) { return (a.urutan || 0) - (b.urutan || 0); })
        .map(function (m) {
          var b = {
            id: m.materi_id, type: m.jenis, title: m.judul, desc: m.deskripsi, presenter: m.pengisi,
            enabled: bool_(m.dipakai), open: bool_(m.terbuka)
          };
          if (m.jenis === 'ppt') { b.url = m.url; b.fileName = m.nama_berkas; }
          if (m.jenis === 'zoom') { b.url = m.url; b.meetingId = m.meeting_id; b.passcode = m.passcode; }
          if (m.jenis === 'wordcloud') b.prompt = m.prompt;
          if (m.jenis === 'checkin') b.token = String(m.kode_barcode || '').toUpperCase();
          if (m.jenis === 'quiz') {
            b.quizKind = m.jenis_kuis || 'pre';
            b.timed = bool_(m.berbatas_waktu);
            b.timeMode = m.mode_waktu || 'per';
            b.seconds = Number(m.detik) || (b.timeMode === 'total' ? 300 : 30);
            b.allowRetake = bool_(m.boleh_ulang);
            b.questions = soal.filter(function (q) { return String(q.materi_id) === String(m.materi_id); })
              .sort(function (a, c) { return (a.urutan || 0) - (c.urutan || 0); })
              .map(function (q) {
                return {
                  id: q.soal_id, text: q.pertanyaan,
                  options: [q.opsi_a, q.opsi_b, q.opsi_c, q.opsi_d].map(function (x) { return x === '' || x == null ? '' : String(x); }),
                  answer: Math.max(0, 'ABCD'.indexOf(String(q.jawaban).toUpperCase()))
                };
              });
          }
          if (m.jenis === 'form') {
            b.fields = isian.filter(function (f) { return String(f.materi_id) === String(m.materi_id); })
              .sort(function (a, c) { return (a.urutan || 0) - (c.urutan || 0); })
              .map(function (f) {
                return {
                  id: f.isian_id, label: f.label, type: f.jenis || 'short', required: bool_(f.wajib),
                  options: String(f.pilihan || '').split(',').map(function (x) { return x.trim(); }).filter(String)
                };
              });
          }
          return b;
        })
    };
  });

  var responses = {};
  sessions.forEach(function (s) { responses[s.id] = { checkins: {}, quiz: {}, words: [], feedback: {}, forms: {} }; });
  rows_('Checkin').forEach(function (r) { if (responses[r.sesi_id]) responses[r.sesi_id].checkins[r.peserta_id] = r.waktu; });
  rows_('JawabanKuis').forEach(function (r) {
    var R = responses[r.sesi_id]; if (!R) return;
    if (!R.quiz[r.materi_id]) R.quiz[r.materi_id] = {};
    R.quiz[r.materi_id][r.peserta_id] = {
      score: Number(r.skor) || 0,
      answers: String(r.jawaban || '').split(',').map(function (x) { return x === '' ? null : Number(x); })
    };
  });
  rows_('WordCloud').forEach(function (r) {
    if (responses[r.sesi_id]) responses[r.sesi_id].words.push({ blockId: r.materi_id, word: String(r.kata), by: r.peserta_id });
  });
  rows_('Feedback').forEach(function (r) {
    var R = responses[r.sesi_id]; if (!R) return;
    if (!R.feedback[r.materi_id]) R.feedback[r.materi_id] = {};
    R.feedback[r.materi_id][r.peserta_id] = { emoji: r.emoji, note: r.catatan || '' };
  });
  rows_('JawabanForm').forEach(function (r) {
    var R = responses[r.sesi_id]; if (!R) return;
    if (!R.forms[r.materi_id]) R.forms[r.materi_id] = {};
    if (!R.forms[r.materi_id][r.peserta_id]) R.forms[r.materi_id][r.peserta_id] = {};
    R.forms[r.materi_id][r.peserta_id][r.isian_id] = r.nilai;
    if (r.nama_berkas) R.forms[r.materi_id][r.peserta_id][r.isian_id + '_nama'] = r.nama_berkas;
  });

  var aktif = sesiRows.filter(function (s) { return bool_(s.aktif); })[0];
  var run = null;
  try { run = JSON.parse(PropertiesService.getScriptProperties().getProperty('run') || 'null'); } catch (e) {}
  if (run && run.sessionId) {
    sessions.forEach(function (s) { if (s.id === run.sessionId) s.run = { blockId: run.blockId, start: run.start, seconds: run.seconds, timeMode: run.timeMode }; });
  }

  return {
    sessions: sessions,
    activeSessionId: aktif ? aktif.sesi_id : (sessions[0] ? sessions[0].id : null),
    participants: rows_('Peserta').map(function (p) {
      return { id: p.peserta_id, name: p.nama, code: String(p.kode).toUpperCase(), group: p.kelompok, org: p.instansi, phone: p.hp, email: p.email };
    }),
    responses: responses
  };
}

/* ══════════ TULIS: admin menyimpan susunan sesi & materi ══════════ */

function putDb_(db) {
  var sesi = [], materi = [], soal = [], isian = [], peserta = [], run = null;
  var checkin = [], kuis = [], kata = [], fb = [], formRows = [];
  var R = db.responses || {};
  Object.keys(R).forEach(function (sid) {
    var r = R[sid] || {};
    Object.keys(r.checkins || {}).forEach(function (pid) {
      checkin.push({ sesi_id: sid, peserta_id: pid, waktu: r.checkins[pid], metode: 'barcode' });
    });
    Object.keys(r.quiz || {}).forEach(function (mid) {
      Object.keys(r.quiz[mid]).forEach(function (pid) {
        var x = r.quiz[mid][pid];
        kuis.push({ sesi_id: sid, materi_id: mid, peserta_id: pid, skor: x.score, jawaban: (x.answers || []).map(function (a) { return a == null ? '' : a; }).join(','), waktu: now_() });
      });
    });
    (r.words || []).forEach(function (w) {
      kata.push({ sesi_id: sid, materi_id: w.blockId, peserta_id: w.by, kata: w.word, waktu: now_() });
    });
    Object.keys(r.feedback || {}).forEach(function (mid) {
      Object.keys(r.feedback[mid]).forEach(function (pid) {
        var f = r.feedback[mid][pid];
        fb.push({ sesi_id: sid, materi_id: mid, peserta_id: pid, emoji: f.emoji, catatan: f.note || '', waktu: now_() });
      });
    });
    Object.keys(r.forms || {}).forEach(function (mid) {
      Object.keys(r.forms[mid]).forEach(function (pid) {
        var ans = r.forms[mid][pid];
        Object.keys(ans).forEach(function (fid) {
          if (fid.indexOf('_nama') >= 0) return;
          formRows.push({ sesi_id: sid, materi_id: mid, peserta_id: pid, isian_id: fid, nilai: ans[fid], nama_berkas: ans[fid + '_nama'] || '', waktu: now_() });
        });
      });
    });
  });
  (db.sessions || []).forEach(function (s, si) {
    sesi.push({
      sesi_id: s.id, judul: s.title, deskripsi: s.desc, narasumber: s.speaker, waktu: s.time,
      tempat: s.venue, format: s.mode, aktif: db.activeSessionId === s.id,
      materi_ditampilkan: s.active || '', urutan: si + 1
    });
    if (s.run && s.run.blockId) run = { sessionId: s.id, blockId: s.run.blockId, start: s.run.start, seconds: s.run.seconds, timeMode: s.run.timeMode };
    (s.blocks || []).forEach(function (b, bi) {
      materi.push({
        materi_id: b.id, sesi_id: s.id, jenis: b.type, judul: b.title, deskripsi: b.desc, pengisi: b.presenter,
        dipakai: !!b.enabled, terbuka: !!b.open, url: b.url || '', nama_berkas: b.fileName || '',
        meeting_id: b.meetingId || '', passcode: b.passcode || '', prompt: b.prompt || '',
        jenis_kuis: b.quizKind || '', berbatas_waktu: !!b.timed, mode_waktu: b.timeMode || '',
        detik: b.seconds || '', boleh_ulang: !!b.allowRetake, kode_barcode: b.token || '', urutan: bi + 1
      });
      (b.questions || []).forEach(function (q, qi) {
        soal.push({
          soal_id: q.id, materi_id: b.id, pertanyaan: q.text,
          opsi_a: (q.options || [])[0] || '', opsi_b: (q.options || [])[1] || '',
          opsi_c: (q.options || [])[2] || '', opsi_d: (q.options || [])[3] || '',
          jawaban: letters_()[q.answer] || 'A', urutan: qi + 1
        });
      });
      (b.fields || []).forEach(function (f, fi) {
        isian.push({
          isian_id: f.id, materi_id: b.id, label: f.label, jenis: f.type,
          wajib: !!f.required, pilihan: (f.options || []).join(','), urutan: fi + 1
        });
      });
    });
  });
  (db.participants || []).forEach(function (p) {
    peserta.push({ peserta_id: p.id, nama: p.name, kode: p.code, kelompok: p.group, instansi: p.org, hp: p.phone, email: p.email });
  });
  rewrite_('Sesi', sesi);
  rewrite_('Materi', materi);
  rewrite_('Soal', soal);
  rewrite_('IsianFormulir', isian);
  rewrite_('Peserta', peserta);
  if (Object.keys(R).length) {
    rewrite_('Checkin', checkin);
    rewrite_('JawabanKuis', kuis);
    rewrite_('WordCloud', kata);
    rewrite_('Feedback', fb);
    rewrite_('JawabanForm', formRows);
  }
  PropertiesService.getScriptProperties().setProperty('run', JSON.stringify(run));
}

/* ══════════ WEB APP ══════════ */

function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || 'db';
  if (a === 'db') return json_({ ok: true, data: buildDb_(true) });
  if (a === 'admin') {
    if (e.parameter.password !== ADMIN_PASSWORD) return json_({ ok: false, error: 'unauthorized' });
    return json_({ ok: true, data: buildDb_(true) });
  }
  return json_({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var p = JSON.parse((e.postData && e.postData.contents) || '{}');
    var admin = p.password === ADMIN_PASSWORD;

    switch (p.action) {

      case 'login': {
        var f = rows_('Peserta').filter(function (x) {
          return String(x.kode).toUpperCase() === String(p.code).toUpperCase() &&
                 String(x.nama).toLowerCase().trim() === String(p.name).toLowerCase().trim() &&
                 String(x.kelompok) === String(p.group);
        })[0];
        if (!f) return json_({ ok: false, error: 'Data login tidak sesuai.' });
        return json_({ ok: true, data: { id: f.peserta_id, name: f.nama, group: f.kelompok, code: f.kode } });
      }

      case 'checkin': {
        var blok = rows_('Materi').filter(function (m) { return String(m.materi_id) === String(p.blockId); })[0];
        if (!blok) return json_({ ok: false, error: 'Materi check in tidak ditemukan.' });
        if (!bool_(blok.terbuka)) return json_({ ok: false, error: 'Check in belum dibuka fasilitator.' });
        var want = String(blok.kode_barcode || '').toUpperCase();
        var got = String(p.token || '').toUpperCase().trim()
          .replace('WS-CHECKIN:' + String(p.sessionId).toUpperCase() + ':', '');
        if (!want || got !== want) return json_({ ok: false, error: 'Kode barcode tidak cocok.' });
        upsert_('Checkin', ['sesi_id', 'peserta_id'], { sesi_id: p.sessionId, peserta_id: p.id, waktu: jam_(), metode: 'barcode' });
        return json_({ ok: true });
      }

      case 'quiz': {
        var qs = rows_('Soal').filter(function (q) { return String(q.materi_id) === String(p.blockId); })
          .sort(function (a, c) { return (a.urutan || 0) - (c.urutan || 0); });
        var correct = 0;
        qs.forEach(function (q, i) { if ('ABCD'.indexOf(String(q.jawaban).toUpperCase()) === p.answers[i]) correct++; });
        var score = qs.length ? Math.round(correct / qs.length * 100) : 0;
        upsert_('JawabanKuis', ['sesi_id', 'materi_id', 'peserta_id'], {
          sesi_id: p.sessionId, materi_id: p.blockId, peserta_id: p.id, skor: score,
          jawaban: (p.answers || []).map(function (x) { return x == null ? '' : x; }).join(','), waktu: now_()
        });
        return json_({ ok: true, data: { score: score } });
      }

      case 'word': {
        var mine = rows_('WordCloud').filter(function (w) {
          return String(w.materi_id) === String(p.blockId) && String(w.peserta_id) === String(p.id);
        });
        if (mine.length >= 3) return json_({ ok: false, error: 'Kuota tiga kata sudah terpakai.' });
        SS.getSheetByName('WordCloud').appendRow([p.sessionId, p.blockId, p.id, String(p.word).trim().toLowerCase(), now_()]);
        return json_({ ok: true });
      }

      case 'feedback':
        upsert_('Feedback', ['sesi_id', 'materi_id', 'peserta_id'], {
          sesi_id: p.sessionId, materi_id: p.blockId, peserta_id: p.id,
          emoji: p.emoji, catatan: p.note || '', waktu: now_()
        });
        return json_({ ok: true });

      case 'form':
        Object.keys(p.values || {}).forEach(function (fid) {
          if (fid.indexOf('_nama') >= 0) return;
          upsert_('JawabanForm', ['sesi_id', 'materi_id', 'peserta_id', 'isian_id'], {
            sesi_id: p.sessionId, materi_id: p.blockId, peserta_id: p.id, isian_id: fid,
            nilai: p.values[fid], nama_berkas: (p.values[fid + '_nama'] || ''), waktu: now_()
          });
        });
        return json_({ ok: true });

      case 'upload': {
        var it = DriveApp.getFoldersByName(UPLOAD_FOLDER);
        var folder = it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER);
        var blob = Utilities.newBlob(Utilities.base64Decode(p.dataBase64), p.mimeType, p.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        upsert_('JawabanForm', ['sesi_id', 'materi_id', 'peserta_id', 'isian_id'], {
          sesi_id: p.sessionId, materi_id: p.blockId, peserta_id: p.id, isian_id: p.fieldId,
          nilai: file.getUrl(), nama_berkas: p.fileName, waktu: now_()
        });
        return json_({ ok: true, data: { url: file.getUrl() } });
      }

      case 'putDb':
        if (!admin) return json_({ ok: false, error: 'unauthorized' });
        putDb_(p.db || {});
        return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
