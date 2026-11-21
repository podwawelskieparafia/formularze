/**
 * @OnlyCurrentDoc
 * Skrypt ma dostęp wyłącznie do tego jednego arkusza.
 */

/**
 * Formularze "Szukam pomocy" i "Chcę pomóc" -> Google Sheets
 *
 * 1. Uruchom raz funkcję setup (zakładki, nagłówki, formatowanie).
 * 2. Wdróż -> Nowe wdrożenie -> Aplikacja internetowa
 *    Wykonaj jako: Ja, Kto ma dostęp: Każdy.
 *
 * Skrypt tylko PRZYJMUJE zgłoszenia. Nie ma funkcji doGet, więc z zewnątrz
 * nie da się odczytać żadnych danych z arkusza.
 */

const MIN_FILL_MS = 3000;   // formularz wysłany szybciej niż 3 s po otwarciu = bot
const MAX_PER_HOUR = 30;    // maks. liczba zgłoszeń na godzinę (łącznie)

const TYPY = ['Dla siebie', 'Zgłaszam kogoś'];
const OSOBA_WIE = ['', 'Tak', 'Nie', 'Nie wiem'];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const STATUSES = ['Nowe', 'W trakcie', 'Połączono', 'Zamknięte'];
const ADMIN_COLS = ['Status', 'Opublikować', 'Opis publiczny', 'Uwagi koordynatora'];

// [parametr z formularza, nagłówek kolumny, maks. długość]
const FORMS = {
  szukam: {
    sheet: 'Szukam pomocy',
    fields: [
      ['typ', 'Kto zgłasza', 30],
      ['imie', 'Imię / pseudonim', 80],
      ['telefon', 'Telefon', 30],
      ['email', 'E-mail', 120],
      ['osoba', 'Osoba potrzebująca', 200],
      ['relacja', 'Relacja', 100],
      ['kontakt_osoby', 'Kontakt do osoby', 200],
      ['osoba_wie', 'Osoba wie o zgłoszeniu', 20],
      ['kategorie', 'Rodzaj pomocy', 500],
      ['opis', 'Opis potrzeby', 3000],
      ['zgoda_kontakt', 'Zgoda na przekazanie kontaktu', 5],
      ['zgoda_publikacja', 'Zgoda na publikację', 5]
    ],
    required: ['typ', 'kategorie', 'opis']
  },
  chce: {
    sheet: 'Chcę pomóc',
    fields: [
      ['imie', 'Imię / pseudonim', 80],
      ['telefon', 'Telefon', 30],
      ['email', 'E-mail', 120],
      ['kategorie', 'Rodzaj pomocy', 500],
      ['opis', 'Opis pomocy', 3000],
      ['zgoda_kontakt', 'Zgoda na przekazanie kontaktu', 5],
      ['zgoda_publikacja', 'Zgoda na publikację', 5]
    ],
    required: ['kategorie', 'opis']
  }
};

function headers_(cfg) {
  return ['Data'].concat(cfg.fields.map(f => f[1]), ADMIN_COLS);
}

function sheet_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(cfg.sheet);
}

// ---------- Konfiguracja arkusza (uruchom raz) ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Europe/Warsaw');
  const names = Object.keys(FORMS).map(k => FORMS[k].sheet);

  Object.keys(FORMS).forEach(function (key, i) {
    const cfg = FORMS[key];
    let sh = ss.getSheetByName(cfg.sheet);
    if (!sh) {
      const first = ss.getSheets()[0];
      if (i === 0 && first.getLastRow() === 0 && names.indexOf(first.getName()) === -1) {
        sh = first;
        sh.setName(cfg.sheet);
      } else {
        sh = ss.insertSheet(cfg.sheet);
      }
    }
    const headers = headers_(cfg);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    const n = headers.length;
    const rows = sh.getMaxRows() - 1;
    const head = sh.getRange(1, 1, 1, n);
    head.setFontWeight('bold').setBackground('#e8eaed').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);

    // Kolumny koordynatora oznaczone na żółto
    const firstAdmin = n - ADMIN_COLS.length + 1;
    sh.getRange(1, firstAdmin, 1, ADMIN_COLS.length).setBackground('#fff2cc');

    headers.forEach(function (h, idx) {
      const col = idx + 1;
      let w = 140;
      if (/^Opis|Uwagi/.test(h)) w = 360;
      if (h === 'Rodzaj pomocy') w = 220;
      if (h === 'Data') w = 130;
      if (/^Zgoda|wie o/.test(h)) w = 110;
      sh.setColumnWidth(col, w);
      if (/^Opis|Uwagi|Rodzaj/.test(h)) sh.getRange(2, col, rows, 1).setWrap(true);
    });
    sh.getRange(2, 1, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange(2, 1, rows, n).setVerticalAlignment('top');

    const statusCol = headers.indexOf('Status') + 1;
    sh.getRange(2, statusCol, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
    const pubCol = headers.indexOf('Opublikować') + 1;
    sh.getRange(2, pubCol, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build());
  });
}

// ---------- Przyjmowanie zgłoszeń ----------

function doPost(e) {
  try {
    return handlePost_((e && e.parameter) || {});
  } catch (err) {
    // Szczegóły błędu tylko w dzienniku skryptu, nie w odpowiedzi.
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function handlePost_(p) {
  // Pułapka na boty: pole niewidoczne dla ludzi musi być puste.
  // Botowi odpowiadamy "ok", żeby nie próbował dalej.
  if (p.website) return json_({ ok: true });

  if (!Object.prototype.hasOwnProperty.call(FORMS, p.form)) return json_({ ok: false, error: 'invalid' });
  const cfg = FORMS[p.form];

  // Formularz wysłany zbyt szybko po otwarciu strony.
  const elapsed = Number(p.elapsed);
  if (!(elapsed >= MIN_FILL_MS)) return json_({ ok: false, error: 'too_fast' });

  if (p.zgoda_rodo !== 'Tak') return json_({ ok: false, error: 'invalid' });

  const values = Object.create(null);
  cfg.fields.forEach(function (f) { values[f[1]] = clean_(p[f[0]], f[2]); });
  values['Zgoda na przekazanie kontaktu'] = p.zgoda_kontakt === 'Tak' ? 'Tak' : 'Nie';
  values['Zgoda na publikację'] = p.zgoda_publikacja === 'Tak' ? 'Tak' : 'Nie';

  const missing = cfg.required.some(function (k) {
    const f = cfg.fields.find(x => x[0] === k);
    return !values[f[1]];
  });
  const tel = values['Telefon'], mail = values['E-mail'];
  if (missing || (!tel && !mail)) return json_({ ok: false, error: 'invalid' });

  // Te same reguły co w formularzu, sprawdzone ponownie po stronie serwera.
  if (tel && !(/^[0-9 +()\-]+$/.test(tel) && tel.replace(/\D/g, '').length >= 7)) return json_({ ok: false, error: 'invalid' });
  if (mail && !EMAIL_RE.test(mail)) return json_({ ok: false, error: 'invalid' });
  if ('Kto zgłasza' in values && TYPY.indexOf(values['Kto zgłasza']) === -1) return json_({ ok: false, error: 'invalid' });
  if ('Osoba wie o zgłoszeniu' in values && OSOBA_WIE.indexOf(values['Osoba wie o zgłoszeniu']) === -1) return json_({ ok: false, error: 'invalid' });

  // Tekst wpisany przez ludzi: zabezpieczenie przed formułami (także po eksporcie do CSV/Excela).
  Object.keys(values).forEach(function (h) {
    if (h !== 'Telefon') values[h] = safeText_(values[h]);
  });
  // Telefon zawiera tylko cyfry, spacje, +, -, ( ). Zapisujemy go jako tekst.
  if (tel) values['Telefon'] = "'" + tel;

  values['Data'] = new Date();
  values['Status'] = 'Nowe';
  values['Opublikować'] = false;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    // Limit zgłoszeń na godzinę.
    const cache = CacheService.getScriptCache();
    const key = 'n' + Math.floor(Date.now() / 3600000);
    const count = Number(cache.get(key) || 0);
    if (count >= MAX_PER_HOUR) return json_({ ok: false, error: 'rate_limited' });
    cache.put(key, String(count + 1), 3600);

    const sh = sheet_(cfg);
    if (!sh) return json_({ ok: false, error: 'not_configured' });
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => (h in values ? values[h] : ''));
    const target = nextRow_(sh, row.length);
    if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    sh.getRange(target, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// Pierwszy całkowicie pusty wiersz pod nagłówkiem (niezaznaczone pole wyboru = puste).
// Nie używamy appendRow: pola wyboru w kolumnie "Opublikować" Arkusze uznają
// za wypełnione, więc appendRow dopisywał zgłoszenia na samym dole arkusza.
// Szukamy od góry: zapomniane wiersze na dole nie przeszkadzają,
// a żaden wiersz z jakąkolwiek treścią nie zostanie nadpisany.
function nextRow_(sh, ncols) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const data = sh.getRange(2, 1, last - 1, ncols).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i].every(v => v === '' || v === false || v === null)) return i + 2;
  }
  return last + 1;
}

// ---------- Pomocnicze ----------

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

// Tekst zaczynający się od =, +, -, @ (także w wersji pełnej szerokości)
// Arkusze lub Excel mogłyby potraktować jako formułę. Dodajemy apostrof,
// który zostaje w treści, więc formuła nie zadziała także po eksporcie do CSV.
function safeText_(s) {
  if (typeof s !== 'string' || !s) return s;
  return /^[=+\-@\t\r\uFF1D\uFF0B\uFF0D\uFF20]/.test(s) ? "''" + s : s;
}

function json_(o) { return text_(JSON.stringify(o)); }

function text_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
