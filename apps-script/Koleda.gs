/**
 * @OnlyCurrentDoc
 * Skrypt ma dostęp wyłącznie do tego jednego arkusza.
 */

/**
 * Formularz "Wizyta duszpasterska" (kolęda) -> Google Sheets
 *
 * Osobny arkusz i osobny skrypt, niezależne od formularzy Zespołu Charytatywnego.
 *
 * 1. W nowym arkuszu: Rozszerzenia -> Apps Script, wklej ten kod, zapisz
 *    i uruchom funkcję setup (zakładki, nagłówki, formatowanie). Można ją uruchamiać
 *    ponownie: dopisze brakujące kolumny na końcu i nie ruszy istniejących danych.
 * 2. Wdróż -> Nowe wdrożenie -> Aplikacja internetowa
 *    Wykonaj jako: Ja, Kto ma dostęp: Każdy.
 *    Adres /exec trzeba wpisać w app.js jako KOLEDA_URL.
 *    Po każdej zmianie kodu: Wdróż -> Zarządzaj wdrożeniami -> Edytuj (ołówek)
 *    -> Wersja: Nowa wersja -> Wdróż. Adres /exec zostaje ten sam.
 *
 * Nowy rok bez zmiany kodu: w zakładce "Ustawienia" wpisz nowy rok (np. 2027/28)
 * i nowy termin. Formularz pokaże nowy rok, a zgłoszenia zaczną trafiać do nowej
 * zakładki "Zgłoszenia 2027/28" (powstaje sama). Poprzednie lata zostają w swoich zakładkach.
 * Po dniu z pola "Zgłoszenia do" formularz przestaje przyjmować zgłoszenia.
 *
 * Księża: w zakładce "Ustawienia" w kolumnie "Księża" (E) wpisuje się kapłanów, jeden pod
 * drugim. W kolumnie "Ksiądz" przy zgłoszeniu wybiera się kapłana z tej listy.
 *
 * "Sprawdź termin": po podaniu adresu strona pokazuje "Data wizyty", "Godzina wizyty" i "Ksiądz",
 * gdy Status = "Zaplanowane" (albo informację, że wizyta już się odbyła). W każdym innym
 * przypadku odpowiedź jest ta sama ("termin nieustalony"), więc po adresie nie da się
 * sprawdzić, kto złożył zgłoszenie. Imiona, telefony i uwagi nie opuszczają arkusza.
 */

const MIN_FILL_MS = 3000;   // formularz wysłany szybciej niż 3 s po otwarciu = bot
const MAX_PER_HOUR = 60;    // maks. liczba zgłoszeń na godzinę (łącznie)
const MAX_LOOKUPS_PER_HOUR = 500;   // maks. liczba sprawdzeń terminu na godzinę (łącznie)
const TZ = 'Europe/Warsaw';
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Zakładka z ustawieniami: w kolumnie A nazwa ustawienia, w kolumnie B wartość.
const SETTINGS = 'Ustawienia';
const SET_ROK = 'Rok';
const SET_TERMIN = 'Zgłoszenia do';
const ROK_RE = /^\d{4}([\/-]\d{2}|[\/-]\d{4})?$/;   // 2026/27, 2026/2027, 2027
const PREFIX = 'Zgłoszenia ';                         // zakładka roku: "Zgłoszenia 2026/27"
const DEFAULTS = [
  [SET_ROK, '2026/27', 'Rok wizyty duszpasterskiej, np. 2027/28. Formularz pokazuje go w tytule, a zgłoszenia trafiają do zakładki „Zgłoszenia <rok>” (nowa zakładka powstaje sama).'],
  [SET_TERMIN, new Date(2026, 11, 20, 12), 'Ostatni dzień przyjmowania zgłoszeń (pokazywany w formularzu). Następnego dnia formularz się zamyka. Puste pole = bez terminu.']
];

// Ulice parafii (te same co na liście w formularzu). Na ulicach z ODD_ONLY
// do parafii należą tylko numery nieparzyste.
const STREETS = ['Komandosów', 'Ludwinowska', 'Rozdroże', 'Słomiana', 'Swoboda', 'Turecka', 'Zatorska'];
const ODD_ONLY = ['Słomiana'];

const SASIAD = 'Mieszkanie sąsiada';
const TYPY = ['Moje mieszkanie', SASIAD];
const PORY = ['Przed południem (9–14)', 'Po południu (16–21)', 'Obojętnie'];
const STATUSES = ['Nowe', 'Zaplanowane', 'Po wizycie', 'Odwołane'];
// Kilka zgłoszeń tego samego mieszkania: pokazujemy najważniejsze.
const RANK = { zaplanowane: 0, po: 1, brak: 2 };
const PLANNED = 'Zaplanowane';
const ADMIN_COLS = ['Status', 'Data wizyty', 'Godzina wizyty', 'Ksiądz', 'Uwagi kancelarii'];   // dla kancelarii (na żółto)
// Lista księży w zakładce "Ustawienia": kolumna E, od 2. wiersza.
const PRIESTS_HEAD = 'Księża';
const PRIESTS_COL = 5;
const PRIESTS_MAX = 50;
// Poprzednie nazwy kolumn; setup zamienia je na obecne (dane zostają).
const RENAMED_COLS = { 'Termin wizyty': 'Data wizyty' };

// [parametr z formularza, nagłówek kolumny, maks. długość]
// Adres na początku: po posortowaniu Ulica -> Nr domu -> Nr mieszkania wychodzi trasa.
const FIELDS = [
  ['ulica', 'Ulica', 80],
  ['nr_domu', 'Nr domu', 10],
  ['nr_mieszkania', 'Nr mieszkania', 10],
  ['imie', 'Zgłaszający', 100],
  ['telefon', 'Telefon', 30],
  ['email', 'E-mail', 120],
  ['typ', 'Czyje mieszkanie', 30],
  ['sasiad', 'Sąsiad', 100],
  ['sasiad_tel', 'Telefon sąsiada', 30],
  ['pora', 'Pora odwiedzin', 40],
  ['dluzsze', 'Dłuższe spotkanie', 3],
  ['uwagi', 'Uwagi do parafii', 2000]
];
const REQUIRED = ['ulica', 'nr_domu', 'imie', 'telefon', 'typ', 'pora'];
const HEADERS = ['Data zgłoszenia'].concat(FIELDS.map(f => f[1]), ADMIN_COLS);

// ---------- Konfiguracja arkusza (uruchom raz i po każdej aktualizacji kodu) ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  setupSettings_(ss);
  const s = settings_();
  // Zakładka bieżącego roku: nowa albo istniejąca (dopisuje kolumny, odświeża formatowanie).
  if (s.rok) prepare_(yearSheet_(s.rok));
}

// Zakładka "Ustawienia": brakujące ustawienia dopisujemy, istniejących wartości nie ruszamy.
function setupSettings_(ss) {
  let sh = ss.getSheetByName(SETTINGS);
  if (!sh) {
    // Pusty pierwszy arkusz ("Arkusz1") dostaje naszą nazwę, zamiast tworzyć drugi.
    const first = ss.getSheets()[0];
    if (first.getLastRow() === 0) { sh = first; sh.setName(SETTINGS); }
    else sh = ss.insertSheet(SETTINGS);
  }
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, 3).setValues([['Ustawienie', 'Wartość', 'Opis']]);
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]).trim());
  // Ostatni wiersz z ustawieniem w kolumnie A (lista księży w kolumnie E może być dłuższa).
  let lastA = keys.length;
  while (lastA > 0 && !keys[lastA - 1]) lastA--;
  DEFAULTS.forEach(function (d) {
    let row = keys.indexOf(d[0]) + 1;
    if (!row) {
      row = ++lastA;
      keys[row - 1] = d[0];
      // Rok jako tekst: inaczej Arkusze mogłyby zamienić np. "2027/01" na datę.
      if (d[0] === SET_ROK) sh.getRange(row, 2).setNumberFormat('@');
      sh.getRange(row, 1, 1, 3).setValues([d]);
    }
    if (d[0] === SET_ROK) sh.getRange(row, 2).setNumberFormat('@');
    if (d[0] === SET_TERMIN) {
      // Data z kalendarzykiem (dwuklik na komórce).
      sh.getRange(row, 2).setNumberFormat('dd.mm.yyyy')
        .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
    }
  });
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(2, 2, DEFAULTS.length, 1).setBackground('#fff2cc').setFontWeight('bold');
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 560);
  sh.getRange(1, 3, DEFAULTS.length + 1, 1).setWrap(true);

  // Lista księży (kolumna E): z niej wybiera się kapłana w kolumnie "Ksiądz".
  const head = sh.getRange(1, PRIESTS_COL);
  if (String(head.getValues()[0][0]).trim() === '') head.setValue(PRIESTS_HEAD);
  head.setFontWeight('bold').setBackground('#e8eaed')
    .setNote('Wpisz kapłanów jeden pod drugim, np. „ks. Jan Kowalski”. Pojawią się na liście w kolumnie „Ksiądz” w zakładce zgłoszeń i na stronie w „Sprawdź termin”.');
  sh.getRange(2, PRIESTS_COL, PRIESTS_MAX, 1).setBackground('#fff2cc').setNumberFormat('@');
  sh.setColumnWidth(PRIESTS_COL - 1, 30); sh.setColumnWidth(PRIESTS_COL, 220);
  sh.setFrozenRows(1);
}

// Zakładka danego roku; gdy jej nie ma, powstaje (na początku, jako pierwsza).
function yearSheet_(rok) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = PREFIX + rok;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, 0);
    prepare_(sh);
  }
  return sh;
}

// Nagłówki i formatowanie zakładki ze zgłoszeniami.
function prepare_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    let have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    // Kolumny o dawnych nazwach dostają obecne.
    have.forEach(function (h, idx) {
      const to = RENAMED_COLS[h];
      if (to && have.indexOf(to) === -1) { sh.getRange(1, idx + 1).setValue(to); have[idx] = to; }
    });
    // Brakujące kolumny kancelarii wstawiamy zaraz za poprzednią
    // (np. "Godzina wizyty" za "Data wizyty", "Ksiądz" za "Godzina wizyty").
    for (let i = 1; i < ADMIN_COLS.length; i++) {
      const prev = have.indexOf(ADMIN_COLS[i - 1]) + 1;
      if (prev && have.indexOf(ADMIN_COLS[i]) === -1) {
        sh.insertColumnAfter(prev);
        sh.getRange(1, prev + 1).setValue(ADMIN_COLS[i]);
        have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      }
    }
    // Pozostałe brakujące kolumny dopisujemy na końcu.
    const add = HEADERS.filter(h => have.indexOf(h) === -1);
    if (add.length) sh.getRange(1, have.length + 1, 1, add.length).setValues([add]);
  }

  // Dalej pracujemy na faktycznym układzie kolumn (kancelaria mogła je przestawić).
  const actual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const n = actual.length;
  const rows = sh.getMaxRows() - 1;
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#e8eaed').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  actual.forEach(function (h, idx) {
    const col = idx + 1;
    if (ADMIN_COLS.indexOf(h) !== -1) sh.getRange(1, col).setBackground('#fff2cc');
    let w = 130;
    if (/^Uwagi/.test(h)) w = 320;
    if (h === 'Ulica' || h === 'Zgłaszający' || h === 'Sąsiad' || h === 'E-mail') w = 170;
    if (/^Nr /.test(h) || h === 'Dłuższe spotkanie') w = 90;
    if (h === 'Data wizyty' || h === 'Godzina wizyty') w = 115;
    if (h === 'Ksiądz') w = 190;
    if (h === 'Pora odwiedzin' || h === 'Czyje mieszkanie') w = 150;
    sh.setColumnWidth(col, w);
    if (/^Uwagi/.test(h)) sh.getRange(2, col, rows, 1).setWrap(true);
    if (h === 'Data wizyty') {
      // Data z kalendarzykiem (dwuklik na komórce).
      sh.getRange(2, col, rows, 1).setNumberFormat('dd.mm.yyyy')
        .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
      sh.getRange(1, col).setNote('Dwuklik na komórce otwiera kalendarz. Termin widać na stronie w zakładce „Sprawdź termin”, gdy Status = Zaplanowane.');
    }
    if (h === 'Godzina wizyty') {
      // Format "Godzina" (gg:mm). Przedział typu 16:00–18:00 też można wpisać: zostanie tekstem.
      // Na stronę idzie godzina dokładnie tak, jak widać w komórce.
      sh.getRange(2, col, rows, 1).clearDataValidations().setNumberFormat('hh:mm');
      sh.getRange(1, col).setNote('Np. 16:00 (albo przedział 16:00–18:00).');
    }
    if (h === 'Ksiądz') {
      // Lista rozwijana z kolumny "Księża" w "Ustawieniach" (nowe nazwiska pojawiają się same).
      const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
      const r = sh.getRange(2, col, rows, 1).setNumberFormat('@');
      if (set) r.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(set.getRange(2, PRIESTS_COL, PRIESTS_MAX, 1), true).setAllowInvalid(false).build());
      sh.getRange(1, col).setNote('Wybierz z listy. Listę księży zmienia się w zakładce „Ustawienia”, kolumna „Księża”.');
    }
  });
  const dataCol = actual.indexOf('Data zgłoszenia') + 1;
  if (dataCol) sh.getRange(2, dataCol, rows, 1).setNumberFormat('dd.mm.yyyy hh:mm');
  sh.getRange(2, 1, rows, n).setVerticalAlignment('top');
  const statusCol = actual.indexOf('Status') + 1;
  if (statusCol) sh.getRange(2, statusCol, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
}

// Prosty wyzwalacz (działa sam): po zmianie roku w "Ustawieniach"
// od razu powstaje zakładka nowego roku.
function onEdit(e) {
  const sh = e && e.range ? e.range.getSheet() : null;
  if (!sh || sh.getName() !== SETTINGS) return;
  const s = settings_();
  if (s.rok) yearSheet_(s.rok);
}

// ---------- Ustawienia ----------

function settings_() {
  const out = { rok: '', termin: null };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    const k = String(r[0]).trim();
    if (k === SET_ROK) out.rok = String(r[1]).trim();
    if (k === SET_TERMIN) out.termin = date_(r[1]);
  });
  if (!ROK_RE.test(out.rok)) out.rok = '';
  return out;
}

// Zgłoszenia przyjmujemy, gdy rok jest wpisany, a termin nie minął (dzień terminu jeszcze tak).
function isOpen_(s) {
  if (!s.rok) return false;
  if (!s.termin) return true;
  return day_(new Date()) <= day_(s.termin);
}

function day_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

// Data z komórki: obiekt Date albo tekst "20.12.2026".
function date_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  const m = /^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/.exec(String(v || '').trim());
  return m ? new Date(+m[3], +m[2] - 1, +m[1], 12) : null;
}

// ---------- Strona formularza pyta o rok i termin ----------

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action !== 'ustawienia') return json_({ ok: false, error: 'invalid' });
    const s = settings_();
    return json_({ ok: true, rok: s.rok, termin: s.termin ? day_(s.termin) : '', otwarte: isOpen_(s) });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
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
  if (p.form === 'sprawdz') return lookup_(p);
  if (p.form !== 'koleda') return json_({ ok: false, error: 'invalid' });

  // Formularz wysłany zbyt szybko po otwarciu strony.
  const elapsed = Number(p.elapsed);
  if (!(elapsed >= MIN_FILL_MS)) return json_({ ok: false, error: 'too_fast' });

  const s = settings_();
  if (!isOpen_(s)) return json_({ ok: false, error: 'closed' });

  if (p.zgoda_rodo !== 'Tak') return json_({ ok: false, error: 'invalid' });

  const v = Object.create(null);
  FIELDS.forEach(function (f) { v[f[0]] = clean_(p[f[0]], f[2]); });

  // Te same reguły co w formularzu, sprawdzone ponownie po stronie serwera.
  if (REQUIRED.some(k => !v[k])) return json_({ ok: false, error: 'invalid' });
  if (!phoneOk_(v.telefon)) return json_({ ok: false, error: 'invalid' });
  if (v.email && !EMAIL_RE.test(v.email)) return json_({ ok: false, error: 'invalid' });
  if (TYPY.indexOf(v.typ) === -1 || PORY.indexOf(v.pora) === -1) return json_({ ok: false, error: 'invalid' });
  if (STREETS.indexOf(v.ulica) === -1) return json_({ ok: false, error: 'invalid' });
  if (ODD_ONLY.indexOf(v.ulica) !== -1 && parseInt(v.nr_domu, 10) % 2 !== 1) return json_({ ok: false, error: 'invalid' });
  if (['', 'Tak', 'Nie'].indexOf(v.dluzsze) === -1) return json_({ ok: false, error: 'invalid' });
  if (v.typ !== SASIAD) { v.sasiad = ''; v.sasiad_tel = ''; }
  if (v.sasiad_tel && !phoneOk_(v.sasiad_tel)) return json_({ ok: false, error: 'invalid' });

  const values = Object.create(null);
  FIELDS.forEach(function (f) {
    const k = f[0], x = v[k];
    if (k === 'telefon' || k === 'sasiad_tel') values[f[1]] = x ? "'" + x : '';   // telefon jako tekst
    else if (k === 'nr_domu' || k === 'nr_mieszkania') values[f[1]] = number_(x);
    else values[f[1]] = safeText_(x);
  });
  values['Data zgłoszenia'] = new Date();
  values['Status'] = 'Nowe';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    // Limit zgłoszeń na godzinę.
    const cache = CacheService.getScriptCache();
    const key = 'n' + Math.floor(Date.now() / 3600000);
    const count = Number(cache.get(key) || 0);
    if (count >= MAX_PER_HOUR) return json_({ ok: false, error: 'rate_limited' });
    cache.put(key, String(count + 1), 3600);

    const sh = yearSheet_(s.rok);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(headers.map(h => (h in values ? values[h] : '')));
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// ---------- "Sprawdź termin" ----------

// Szuka zgłoszenia w zakładce bieżącego roku po adresie (ulica, nr domu, nr mieszkania).
// Zwraca tylko: zaplanowane (data, godzina), po wizycie albo "brak" - to samo dla
// zgłoszenia nowego, odwołanego i nieistniejącego.
function lookup_(p) {
  const ulica = clean_(p.ulica, 80);
  const dom = nr_(clean_(p.nr_domu, 10));
  const lokal = nr_(clean_(p.nr_mieszkania, 10));
  if (STREETS.indexOf(ulica) === -1 || !dom) return json_({ ok: false, error: 'invalid' });

  // Limit sprawdzeń na godzinę (przybliżony, bez blokady).
  const cache = CacheService.getScriptCache();
  const key = 'q' + Math.floor(Date.now() / 3600000);
  const count = Number(cache.get(key) || 0);
  if (count >= MAX_LOOKUPS_PER_HOUR) return json_({ ok: false, error: 'rate_limited' });
  cache.put(key, String(count + 1), 3600);

  const s = settings_();
  const sh = s.rok ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PREFIX + s.rok) : null;
  if (!sh || sh.getLastRow() < 2) return json_({ ok: true, wynik: 'brak' });
  const rng = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn());
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();   // godzina dokładnie tak, jak widać w arkuszu
  const col = Object.create(null);
  vals[0].forEach(function (h, i) { if (h !== '' && !(h in col)) col[h] = i; });
  if (col['Ulica'] === undefined || col['Nr domu'] === undefined) return json_({ ok: true, wynik: 'brak' });
  const cell = (r, h) => (col[h] === undefined ? '' : r[col[h]]);

  let best = null;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(cell(r, 'Ulica')).trim() !== ulica) continue;
    if (nr_(cell(r, 'Nr domu')) !== dom || nr_(cell(r, 'Nr mieszkania')) !== lokal) continue;
    const res = visit_(r, disp[i], cell);
    if (!best || RANK[res.wynik] < RANK[best.wynik]) best = res;
  }
  return json_(Object.assign({ ok: true }, best || { wynik: 'brak' }));
}

function visit_(r, d, cell) {
  const status = String(cell(r, 'Status')).trim();
  if (status === PLANNED) {
    const data = date_(cell(r, 'Data wizyty'));
    // "Zaplanowane" bez daty: dla mieszkańca to wciąż "termin nieustalony".
    if (!data) return { wynik: 'brak' };
    return {
      wynik: 'zaplanowane',
      data: day_(data),
      godzina: String(cell(d, 'Godzina wizyty')).trim().slice(0, 40),
      ksiadz: String(cell(d, 'Ksiądz')).trim().slice(0, 80)
    };
  }
  if (status === 'Po wizycie') return { wynik: 'po' };
  return { wynik: 'brak' };
}

// Numer domu / mieszkania do porównania: "15 A" = "15a", 18 (liczba) = "18".
function nr_(v) {
  return String(v == null ? '' : v).toLowerCase().replace(/\s+/g, '');
}

// ---------- Pomocnicze ----------

function phoneOk_(s) {
  return /^[0-9 +()\-]+$/.test(s) && s.replace(/\D/g, '').length >= 7;
}

// Numer domu / mieszkania: same cyfry zapisujemy jako liczbę (poprawne sortowanie),
// resztę jako tekst, bo np. "3/5" Arkusze zamieniłyby na datę.
function number_(s) {
  if (!s) return '';
  if (/^\d+$/.test(s)) return Number(s);
  const t = safeText_(s);
  return t === s ? "'" + s : t;
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

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
