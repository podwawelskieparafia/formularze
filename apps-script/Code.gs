/**
 * @OnlyCurrentDoc
 * Skrypt ma dostęp wyłącznie do tego jednego arkusza.
 */

/**
 * Formularze "Szukam pomocy" i "Chcę pomóc" -> Google Sheets
 *
 * 1. Uruchom funkcję setup (zakładki, nagłówki, formatowanie). Można ją uruchamiać
 *    ponownie: dopisze brakujące kolumny na końcu i nie ruszy istniejących danych.
 * 2. Wdróż -> Nowe wdrożenie -> Aplikacja internetowa
 *    Wykonaj jako: Ja, Kto ma dostęp: Każdy.
 *    Po każdej zmianie kodu: Wdróż -> Zarządzaj wdrożeniami -> Edytuj (ołówek)
 *    -> Wersja: Nowa wersja -> Wdróż. Adres /exec zostaje ten sam.
 *
 * doPost przyjmuje zgłoszenia z obu formularzy.
 *
 * doGet zwraca WYŁĄCZNIE listę próśb, które koordynator zaznaczył w kolumnie
 * "Opublikować". Z każdej prośby
 * wychodzą tylko: ID, data dodania, rodzaj pomocy i "Opis publiczny" wpisany przez
 * koordynatora. Imiona, kontakty i oryginalny opis nigdy nie opuszczają arkusza.
 *
 * Gdy wolontariusz wybierze prośbę w formularzu "Chcę pomóc", skrypt odznacza
 * "Opublikować", ustawia status "Wolontariusz zgłoszony" i wpisuje wolontariusza,
 * więc prośba znika z listy. Żeby wróciła na listę, wystarczy znów zaznaczyć "Opublikować".
 *
 * Kolumny "Wybrana prośba" (u wolontariusza) i "Wolontariusz" (przy prośbie) są
 * linkami: kliknięcie przenosi do odpowiedniego wiersza w drugiej zakładce. Wiersz jest
 * szukany po ID, więc linki działają także po sortowaniu i usuwaniu wierszy.
 */

const MIN_FILL_MS = 3000;   // formularz wysłany szybciej niż 3 s po otwarciu = bot
const MAX_PER_HOUR = 30;    // maks. liczba zgłoszeń na godzinę (łącznie)
const TZ = 'Europe/Warsaw';

const LIST_KEY = 'prosby';  // lista próśb w pamięci podręcznej...
const LIST_TTL = 300;       // ...do 5 minut, czyszczona po każdej edycji arkusza i każdej rezerwacji
const LIST_MAX = 50;

// Te same kategorie co w app.js.
const CATEGORIES = [
  'Zakupy i sprawunki',
  'Transport (np. do lekarza)',
  'Drobne naprawy w domu',
  'Sprzątanie i pomoc w domu',
  'Odwiedziny i rozmowa',
  'Żywność, odzież, rzeczy',
  'Pomoc w urzędzie i formalnościach',
  'Nauka i korepetycje',
  'Inne'
];

const TYPY = ['Dla siebie', 'Zgłaszam kogoś'];
const OSOBA_WIE = ['', 'Tak', 'Nie', 'Nie wiem'];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const ID_RE = /^[a-z0-9]{6,12}$/;

const CLAIMED = 'Wolontariusz zgłoszony';
const STATUSES = ['Nowe', 'W trakcie', CLAIMED, 'Połączono', 'Zamknięte'];
const CLOSED = ['Połączono', 'Zamknięte'];   // prośby z tym statusem nie są pokazywane
const ADMIN_COLS = ['Status', 'Opublikować', 'Opis publiczny', 'Uwagi koordynatora'];
const MATCH_COLS = ['Wolontariusz', 'Wybrana prośba'];   // kto do czego się zgłosił (na zielono)
// Kolumny po zgodach usuniętych z formularzy; setup usuwa je z arkusza.
const OBSOLETE_COLS = ['Zgoda na przekazanie kontaktu', 'Zgoda na publikację'];

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
      ['opis', 'Opis potrzeby', 3000]
    ],
    required: ['typ', 'kategorie', 'opis'],
    extra: ['Wolontariusz', 'ID']
  },
  chce: {
    sheet: 'Chcę pomóc',
    fields: [
      ['imie', 'Imię / pseudonim', 80],
      ['telefon', 'Telefon', 30],
      ['email', 'E-mail', 120],
      ['kategorie', 'Rodzaj pomocy', 500],
      ['opis', 'Opis pomocy', 3000]
    ],
    required: ['kategorie', 'opis'],
    extra: ['Wybrana prośba', 'ID']
  }
};

function headers_(cfg) {
  return ['Data'].concat(cfg.fields.map(f => f[1]), ADMIN_COLS, cfg.extra);
}

function sheet_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(cfg.sheet);
}

// ---------- Konfiguracja arkusza (uruchom raz i po każdej aktualizacji kodu) ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
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
    } else {
      // Arkusz z danymi: dopisujemy tylko brakujące kolumny, na końcu.
      let have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      // Usuwamy kolumny zgód, których formularze już nie mają (od prawej, żeby numery się nie przesuwały).
      for (let c = have.length; c >= 1; c--) {
        if (OBSOLETE_COLS.indexOf(have[c - 1]) !== -1) sh.deleteColumn(c);
      }
      have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const add = headers.filter(h => have.indexOf(h) === -1);
      if (add.length) sh.getRange(1, have.length + 1, 1, add.length).setValues([add]);
    }

    // Dalej pracujemy na faktycznym układzie kolumn (koordynator mógł je przestawić).
    const actual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const n = actual.length;
    const rows = sh.getMaxRows() - 1;
    sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#e8eaed').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);

    actual.forEach(function (h, idx) {
      const col = idx + 1;
      // Kolumny koordynatora na żółto, dopasowanie wolontariusza do prośby na zielono
      if (ADMIN_COLS.indexOf(h) !== -1) sh.getRange(1, col).setBackground('#fff2cc');
      if (MATCH_COLS.indexOf(h) !== -1) sh.getRange(1, col).setBackground('#d9ead3');
      let w = 140;
      if (/^Opis|Uwagi/.test(h)) w = 360;
      if (h === 'Rodzaj pomocy') w = 220;
      if (h === 'Data') w = 130;
      if (/^Zgoda|wie o/.test(h)) w = 110;
      if (MATCH_COLS.indexOf(h) !== -1) w = 280;
      if (h === 'ID') w = 90;
      sh.setColumnWidth(col, w);
      if (/^Opis|Uwagi|Rodzaj/.test(h) || MATCH_COLS.indexOf(h) !== -1) sh.getRange(2, col, rows, 1).setWrap(true);
    });
    const dataCol = actual.indexOf('Data') + 1;
    if (dataCol) sh.getRange(2, dataCol, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange(2, 1, rows, n).setVerticalAlignment('top');

    const statusCol = actual.indexOf('Status') + 1;
    sh.getRange(2, statusCol, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
    const pubCol = actual.indexOf('Opublikować') + 1;
    sh.getRange(2, pubCol, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build());
  });

  // Wiersze zapisane przed tą wersją dostają ID (lista na stronie, linki między zakładkami).
  ensureIds_(sheet_(FORMS.szukam));
  ensureIds_(sheet_(FORMS.chce));
  CacheService.getScriptCache().remove(LIST_KEY);
}

// Prosty wyzwalacz (działa sam, bez instalowania): po każdej edycji arkusza lista
// na stronie odświeża się od razu, a prośby wpisane ręcznie dostają ID.
function onEdit(e) {
  CacheService.getScriptCache().remove(LIST_KEY);
  const sh = e && e.range ? e.range.getSheet() : null;
  if (sh && (sh.getName() === FORMS.szukam.sheet || sh.getName() === FORMS.chce.sheet)) ensureIds_(sh);
}

// ---------- Lista próśb dla formularza "Chcę pomóc" ----------

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'prosby') return json_({ ok: true, items: publicRequests_() });
    return json_({ ok: false, error: 'invalid' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function publicRequests_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(LIST_KEY);
  if (hit) return JSON.parse(hit);

  const sh = sheet_(FORMS.szukam);
  const items = [];
  if (sh) {
    const t = table_(sh);
    t.rows.forEach(function (r) {
      if (!isOpen_(t, r)) return;
      const d = r[t.col['Data']];
      const isDate = Object.prototype.toString.call(d) === '[object Date]';
      items.push({
        sort: isDate ? d.getTime() : Number.MAX_SAFE_INTEGER,
        // Tylko te pola trafiają na stronę:
        id: String(r[t.col['ID']]),
        data: isDate ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : '',
        kategorie: categories_(r[t.col['Rodzaj pomocy']]),
        opis: String(r[t.col['Opis publiczny']] || '').trim().slice(0, 600)
      });
    });
  }
  // Najdłużej czekające na górze.
  items.sort((a, b) => a.sort - b.sort);
  const out = items.slice(0, LIST_MAX).map(it => ({ id: it.id, data: it.data, kategorie: it.kategorie, opis: it.opis }));
  cache.put(LIST_KEY, JSON.stringify(out), LIST_TTL);
  return out;
}

// Prośba jest na liście, gdy: ma ID, koordynator zaznaczył "Opublikować"
// i prośba nie jest połączona/zamknięta.
function isOpen_(t, r) {
  const c = t.col;
  return ID_RE.test(String(r[c['ID']]))
    && r[c['Opublikować']] === true
    && CLOSED.indexOf(r[c['Status']]) === -1;
}

function categories_(v) {
  const s = String(v || '').trim();
  const found = CATEGORIES.filter(c => s.indexOf(c) !== -1);
  return found.length ? found : (s ? [s.slice(0, 100)] : []);
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

  // Wolontariusz wybrał konkretną prośbę z listy (tylko formularz "Chcę pomóc").
  const claimId = p.form === 'chce' ? clean_(p.zgloszenie, 20) : '';
  if (claimId && !ID_RE.test(claimId)) return json_({ ok: false, error: 'invalid' });

  const values = Object.create(null);
  cfg.fields.forEach(function (f) { values[f[1]] = clean_(p[f[0]], f[2]); });

  const missing = cfg.required.some(function (k) {
    if (claimId && k === 'kategorie') return false;   // rodzaj pomocy weźmiemy z prośby
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
  values['ID'] = newId_();   // prośba i wolontariusz: do listy na stronie i do linków między zakładkami

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

    // Rezerwacja prośby. Jesteśmy pod blokadą, więc dwie osoby nie zarezerwują tej samej:
    // druga zobaczy już odznaczone "Opublikować" i dostanie "taken".
    let claim = null;
    if (claimId) {
      claim = findOpenRequest_(claimId);
      if (!claim) return json_({ ok: false, error: 'taken' });
      const rodzaj = String(claim.r[claim.t.col['Rodzaj pomocy']] || '');
      const pub = String(claim.r[claim.t.col['Opis publiczny']] || '').trim();
      values['Rodzaj pomocy'] = safeText_(rodzaj);
      values['Wybrana prośba'] = jump_(claim.sh, claim.t.col['ID'] + 1, claimId, claimId + ': ' + (pub || rodzaj));
    }

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => (h in values ? values[h] : ''));
    const target = nextRow_(sh, row.length);
    if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    sh.getRange(target, 1, 1, row.length).setValues([row]);

    if (claim) markClaimed_(claim, values, tel, mail, sh, headers.indexOf('ID') + 1);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function findOpenRequest_(id) {
  const sh = sheet_(FORMS.szukam);
  if (!sh) return null;
  const t = table_(sh);
  for (let i = 0; i < t.rows.length; i++) {
    const r = t.rows[i];
    if (String(r[t.col['ID']]) === id) return isOpen_(t, r) ? { sh: sh, t: t, r: r, row: i + 2 } : null;
  }
  return null;
}

// Prośba znika z listy: odznaczamy "Opublikować", zmieniamy status i dopisujemy
// wolontariusza (poprzedni wpis zostaje, gdyby koordynator opublikował prośbę ponownie).
// Komórka "Wolontariusz" prowadzi do wiersza najnowszego wolontariusza w "Chcę pomóc".
function markClaimed_(c, values, tel, mail, volSheet, volIdCol) {
  const col = c.t.col;
  function set(h, v) { if (col[h] !== undefined) c.sh.getRange(c.row, col[h] + 1).setValue(v); }
  const who = [values['Imię / pseudonim'] || '(bez imienia)', tel ? 'tel. ' + tel : '', mail]
    .filter(Boolean).join(', ');
  const line = who + ' (' + Utilities.formatDate(values['Data'], TZ, 'dd.MM.yyyy HH:mm') + ')';
  const prev = col['Wolontariusz'] !== undefined ? String(c.r[col['Wolontariusz']] || '') : '';
  set('Opublikować', false);
  set('Status', CLAIMED);
  const label = prev ? prev + '\n' + line : line;
  set('Wolontariusz', volIdCol ? jump_(volSheet, volIdCol, values['ID'], label) : safeText_(label));
  CacheService.getScriptCache().remove(LIST_KEY);
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
    if (data[i].every(isEmpty_)) return i + 2;
  }
  return last + 1;
}

// Nadaje ID każdej prośbie z treścią, która go jeszcze nie ma.
function ensureIds_(sh) {
  if (!sh || sh.getLastRow() < 2) return;
  const t = table_(sh);
  const c = t.col['ID'];
  if (c === undefined) return;
  t.rows.forEach(function (r, i) {
    if (r[c] !== '' && r[c] !== null) return;
    if (r.some((v, j) => j !== c && !isEmpty_(v))) sh.getRange(i + 2, c + 1).setValue(newId_());
  });
}

// ---------- Pomocnicze ----------

// Nagłówki i wszystkie wiersze arkusza; col['Nazwa kolumny'] = indeks (od 0).
function table_(sh) {
  const ncols = sh.getLastColumn();
  const last = sh.getLastRow();
  const head = ncols ? sh.getRange(1, 1, 1, ncols).getValues()[0] : [];
  const col = Object.create(null);
  head.forEach(function (h, i) { if (h !== '' && !(h in col)) col[h] = i; });
  const rows = last >= 2 && ncols ? sh.getRange(2, 1, last - 1, ncols).getValues() : [];
  return { head: head, col: col, rows: rows };
}

// Formuła-link do wiersza w zakładce sh, w którym kolumna idCol (od 1) zawiera id.
// Wiersz jest szukany przez MATCH, więc link działa po sortowaniu i usuwaniu wierszy.
// Gdy wiersza już nie ma, komórka pokazuje sam tekst.
function jump_(sh, idCol, id, label) {
  const L = colLetter_(idCol);
  const ref = "'" + sh.getName().replace(/'/g, "''") + "'!" + L + ':' + L;
  const text = lit_(label);
  return '=IFERROR(HYPERLINK("#gid=' + sh.getSheetId() + '&range=A"&MATCH(' + lit_(id) + ',' + ref + ',0),' +
    text + '),' + text + ')';
}

// Dowolny tekst jako literał w formule: cudzysłowy podwojone, znaki sterujące usunięte,
// nowe wiersze jako CHAR(10). Treść wpisana w formularzu nie może więc zmienić formuły.
function lit_(s) {
  return String(s).replace(/\r/g, '').split('\n')
    .map(p => '"' + p.replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/"/g, '""') + '"')
    .join('&CHAR(10)&');
}

function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function isEmpty_(v) { return v === '' || v === false || v === null; }

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

// Tekst zaczynający się od =, +, -, @ (także w wersji pełnej szerokości)
// Arkusze lub Excel mogłyby potraktować jako formułę. Dodajemy apostrof,
// który zostaje w treści, więc formuła nie zadziała także po eksporcie do CSV.
function safeText_(s) {
  if (typeof s !== 'string' || !s) return s;
  return /^[=+\-@\t\r＝＋－＠]/.test(s) ? "''" + s : s;
}

function json_(o) { return text_(JSON.stringify(o)); }

function text_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
