(function () {
  "use strict";

  // ===== KONFIGURACJA =====
  // Adres aplikacji internetowej Google Apps Script (kończy się na /exec).
  // Formularze Zespołu Charytatywnego („Chcę pomóc”, „Szukam pomocy”):
  var SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxQnUtSA23rEhKbNgvlfnlBWYWvwjS6bMwVsO6O4FVpeqNTVxMkUMAGH14kWY5hqvay/exec";
  // Wizyta duszpasterska (kolęda): osobny arkusz i osobny skrypt (apps-script/Koleda.gs).
  var KOLEDA_URL = "https://script.google.com/macros/s/AKfycbxbXBmHd4u0hBwesS0JlLwnqRpZdhCYyJYuSQ042EM_8SXRYGP7ezel9hBkd3wf5-Dldw/exec";

  // Kategorie pomocy (wspólne dla obu formularzy).
  var CATEGORIES = [
    "Zakupy i sprawunki",
    "Transport (np. do lekarza)",
    "Drobne naprawy w domu",
    "Sprzątanie i pomoc w domu",
    "Odwiedziny i rozmowa",
    "Żywność, odzież, rzeczy",
    "Pomoc w urzędzie i formalnościach",
    "Nauka i korepetycje",
    "Inne"
  ];

  var params = new URLSearchParams(location.search);
  var theme = params.get("theme");
  document.documentElement.setAttribute("data-theme",
    theme === "dark" || theme === "auto" ? theme : "light");

  // Automatyczna wysokość: strona podaje swoją wysokość stronie-rodzicowi (iframe).
  function postHeight() {
    if (window.parent === window) return;
    var h = Math.ceil(document.body.getBoundingClientRect().height);
    window.parent.postMessage({ type: "iframe-height", height: h }, "*");
  }
  if ("ResizeObserver" in window) new ResizeObserver(postHeight).observe(document.body);
  window.addEventListener("load", postHeight);
  // Strona-rodzic może poprosić o wysokość (np. z atrybutu onload ramki, który
  // zaczyna nasłuchiwać dopiero po załadowaniu formularza).
  window.addEventListener("message", function (e) {
    if (e.source === window.parent && e.data && e.data.type === "iframe-height-request") postHeight();
  });

  // Kategorie jako pola wyboru
  document.querySelectorAll("[data-categories]").forEach(function (box) {
    CATEGORIES.forEach(function (c) {
      var label = document.createElement("label");
      label.className = "check";
      var input = document.createElement("input");
      input.type = "checkbox";
      input.name = "kategorie";
      input.value = c;
      var span = document.createElement("span");
      span.textContent = c;
      label.append(input, span);
      box.appendChild(label);
    });
  });

  var form = document.querySelector("form[data-form]");
  if (form) initForm(form);

  function active(el) { return !el.closest("[hidden]"); }

  // ===== Lista próśb w formularzu „Chcę pomóc” =====
  // Serwer zwraca tylko prośby zaznaczone przez koordynatora do publikacji
  // i tylko dane anonimowe: rodzaj pomocy, opis publiczny, datę dodania.
  function initRequests(form) {
    var box = form.querySelector("[data-requests]");
    if (!box) return null;
    if (!SCRIPT_URL) { box.hidden = true; return null; }
    var list = box.querySelector(".requests");
    var loading = box.querySelector("[data-loading]");
    var empty = box.querySelector("[data-empty]");
    var more = box.querySelector("[data-more]");
    var unpick = box.querySelector("[data-unpick]");
    var intro = box.querySelector("[data-requests-note]");
    var note = form.querySelector("[data-picked]");
    var SHOW = 3;
    var items = [];
    var expanded = false;
    var loaded = false;
    var dateFmt = new Intl.DateTimeFormat("pl-PL", { day: "numeric", month: "long" });

    function fmtDate(s) {
      var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || "");
      return m ? dateFmt.format(new Date(+m[1], +m[2] - 1, +m[3])) : "";
    }

    function card(it) {
      var label = document.createElement("label");
      label.className = "request";
      var input = document.createElement("input");
      input.type = "radio";
      input.name = "zgloszenie";
      input.value = it.id;
      var body = document.createElement("span");
      body.className = "request-body";
      var cats = document.createElement("span");
      cats.className = "request-cats";
      cats.textContent = (it.kategorie || []).join(" · ");
      var text = document.createElement("span");
      text.className = "request-text";
      // Spójniki i przyimki jednoliterowe nie zostają na końcu wiersza.
      text.textContent = (it.opis || "Szczegóły przekaże koordynator.")
        .replace(/(^|\s)([aiouwzAIOUWZ])\s+/g, "$1$2\u00a0")
        .replace(/(\u00a0)([aiouwzAIOUWZ])\s+/g, "$1$2\u00a0");
      if (!it.opis) text.classList.add("is-muted");
      body.append(cats, text);
      var d = fmtDate(it.data);
      if (d) {
        var date = document.createElement("span");
        date.className = "request-date";
        date.textContent = "Dodano " + d;
        body.append(date);
      }
      var pick = document.createElement("span");
      pick.className = "request-pick";
      pick.setAttribute("aria-hidden", "true");
      label.append(input, body, pick);
      return label;
    }

    function render() {
      var sel = selected();
      list.textContent = "";
      loading.hidden = true;
      empty.hidden = items.length > 0;
      var shown = expanded ? items : items.slice(0, SHOW);
      shown.forEach(function (it) { list.appendChild(card(it)); });
      if (sel) {
        var input = list.querySelector('input[value="' + CSS.escape(sel) + '"]');
        if (input) input.checked = true;
      }
      more.textContent = "Pokaż wszystkie (" + items.length + ")";
      update();
    }

    function selected() {
      var c = list.querySelector("input:checked");
      return c ? c.value : "";
    }

    function uncheck() {
      var c = list.querySelector("input:checked");
      if (c) c.checked = false;
    }

    // Po wyborze prośby zostaje na liście tylko ona, a pytanie „W czym możesz pomóc?”
    // znika, bo rodzaj pomocy wynika z prośby. „Anuluj wybór” przywraca wszystko.
    function update() {
      var has = !!selected();
      list.querySelectorAll(".request").forEach(function (l) {
        l.hidden = has && !l.querySelector("input").checked;
      });
      more.hidden = has || expanded || items.length <= SHOW;
      unpick.hidden = !has;
      // Pusta lista: bez podpowiedzi „Wybierz prośbę…”, zostaje tylko żółta informacja.
      intro.hidden = has || (loaded && !items.length);
      note.hidden = !has;
      form.querySelectorAll("[data-general]").forEach(function (el) { el.hidden = has; });
      postHeight();
    }

    list.addEventListener("change", update);
    more.addEventListener("click", function () { expanded = true; render(); });
    unpick.addEventListener("click", function () { uncheck(); update(); });

    function load() {
      return fetch(SCRIPT_URL + "?action=prosby", { cache: "no-store" })
        .then(function (r) { return r.json(); })
        .then(function (res) {
          if (!res || !res.ok || !Array.isArray(res.items)) throw new Error("bad");
          items = res.items;
          loaded = true;
          render();
        })
        .catch(function () {
          // Bez listy formularz działa jak dotąd (ogólna oferta pomocy).
          box.hidden = true;
          uncheck();
          update();
        });
    }
    load();

    return {
      selected: selected,
      // Rodzaj pomocy wybranej prośby (pole „W czym możesz pomóc?” jest wtedy ukryte).
      categories: function () {
        var id = selected();
        var it = items.find(function (x) { return x.id === id; });
        return it ? (it.kategorie || []) : [];
      },
      // Ktoś inny zdążył zgłosić się do tej prośby: odświeżamy listę bez niej.
      taken: function () {
        uncheck();
        update();
        return load().then(function () {
          box.scrollIntoView({ behavior: "smooth", block: "start" });
        });
      },
      reset: function () {
        expanded = false;
        list.textContent = "";
        update();
        load();
      }
    };
  }

  // ===== Formularz „Wizyta duszpasterska”: rok i termin z arkusza (zakładka „Ustawienia”) =====
  // Po zmianie roku w arkuszu strona sama pokazuje nowy rok; po terminie zamiast formularza
  // jest informacja „Zapisy zakończone”. Gdy ustawień nie da się pobrać, formularz działa
  // normalnie (serwer i tak sprawdza termin).
  function initSeason(form, url) {
    var closed = document.querySelector("[data-closed]");
    if (!closed || !url) return null;
    var dateFmt = new Intl.DateTimeFormat("pl-PL", { day: "numeric", month: "long", year: "numeric" });
    function close() {
      var box = document.querySelector("[data-termin-box]");
      if (box) box.hidden = true;   // „przyjmujemy do …” już nieaktualne
      form.hidden = true;
      closed.hidden = false;
      postHeight();
    }
    fetch(url + "?action=ustawienia", { cache: "no-store" })
      .then(function (r) { return r.json(); })
      .then(function (s) {
        if (!s || !s.ok) return;
        document.querySelectorAll("[data-rok]").forEach(function (el) { el.textContent = s.rok || ""; });
        var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.termin || "");
        var box = document.querySelector("[data-termin-box]");
        if (m && box && s.otwarte) {
          box.querySelector("[data-termin]").textContent =
            (dateFmt.format(new Date(+m[1], +m[2] - 1, +m[3])) + " r.").replace(/ /g, "\u00a0");
          box.hidden = false;
        }
        if (!s.otwarte) close();
        postHeight();
      })
      .catch(function () {});
    return { close: close };
  }

  function initForm(form) {
    var started = Date.now();
    var kind = form.getAttribute("data-form");
    var url = kind === "koleda" ? KOLEDA_URL : SCRIPT_URL;
    var submitBtn = form.querySelector('button[type="submit"]');
    var submitLabel = submitBtn.textContent;
    var msg = form.querySelector(".form-msg");
    var done = document.querySelector(".done");
    var requests = initRequests(form);
    var season = initSeason(form, url);

    // Pokazywanie pól zależnie od wybranego typu zgłoszenia
    function applyTyp() {
      var checked = form.querySelector('input[name="typ"]:checked');
      var key = checked ? checked.getAttribute("data-key") : "";
      form.querySelectorAll("[data-only]").forEach(function (el) {
        el.hidden = el.getAttribute("data-only").split(" ").indexOf(key) === -1;
      });
      form.querySelectorAll("[data-after-typ]").forEach(function (el) { el.hidden = !key; });
      postHeight();
    }
    form.querySelectorAll('input[name="typ"]').forEach(function (r) {
      r.addEventListener("change", function () {
        applyTyp();
        clearErr(r.closest(".field"));
      });
    });
    applyTyp();

    function setErr(field, text) {
      var e = field.querySelector(".err");
      if (e) e.textContent = text || "";
      field.classList.toggle("has-err", !!text);
    }
    function clearErr(field) { if (field) setErr(field, ""); }

    var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    function phoneOk(v) {
      return /^[0-9 +()\-]+$/.test(v) && v.replace(/\D/g, "").length >= 7;
    }

    function check(field) {
      var rule = field.getAttribute("data-req");
      var custom = field.getAttribute("data-msg");
      var inputs = field.querySelectorAll("input, textarea, select");
      if (rule === "text") {
        return inputs[0].value.trim() ? "" : (custom || "To pole jest wymagane.");
      }
      if (rule === "radio") {
        return field.querySelector("input:checked") ? "" : (custom || "Wybierz jedną z opcji.");
      }
      if (rule === "any") {
        return field.querySelector("input:checked") ? "" : (custom || "Wybierz co najmniej jedną opcję.");
      }
      if (rule === "consent") {
        return inputs[0].checked ? "" : (custom || "Ta zgoda jest wymagana, abyśmy mogli przyjąć zgłoszenie.");
      }
      // Numer domu; na ulicy oznaczonej data-odd (Słomiana) tylko numery nieparzyste.
      if (rule === "house") {
        var nr = inputs[0].value.trim();
        if (!nr) return custom || "Podaj numer domu.";
        var street = form.querySelector('select[name="ulica"] option:checked');
        if (street && street.hasAttribute("data-odd") && parseInt(nr, 10) % 2 !== 1) {
          return "Tylko numery nieparzyste.";
        }
        return "";
      }
      if (rule === "email-opt") {
        var m = inputs[0].value.trim();
        return !m || EMAIL_RE.test(m) ? "" : "Sprawdź adres e-mail.";
      }
      if (rule === "tel" || rule === "tel-opt") {
        var v = inputs[0].value.trim();
        if (!v) return rule === "tel" ? (custom || "Podaj numer telefonu.") : "";
        return phoneOk(v) ? "" : "Sprawdź numer telefonu.";
      }
      if (rule === "contact" || rule === "contact-opt") {
        var tel = field.querySelector('input[type="tel"]').value.trim();
        var mail = field.querySelector('input[type="email"]').value.trim();
        if (rule === "contact" && !tel && !mail) return custom || "Podaj telefon lub e-mail, abyśmy mogli się skontaktować.";
        if (tel && !phoneOk(tel)) return "Sprawdź numer telefonu.";
        if (mail && !EMAIL_RE.test(mail)) return "Sprawdź adres e-mail.";
      }
      return "";
    }

    function validate() {
      var first = null;
      form.querySelectorAll("[data-req]").forEach(function (field) {
        if (!active(field)) { clearErr(field); return; }
        var text = check(field);
        setErr(field, text);
        if (text && !first) first = field;
      });
      if (first) {
        var focusable = first.querySelector("input, textarea, select");
        if (focusable) focusable.focus({ preventScroll: true });
        first.scrollIntoView({ behavior: "smooth", block: "center" });
        postHeight();
      }
      return !first;
    }

    // Usuwanie błędu po poprawieniu pola
    form.addEventListener("input", function (e) {
      var field = e.target.closest("[data-req]");
      if (field && field.classList.contains("has-err") && !check(field)) clearErr(field);
    });
    // Zmiana wyboru może naprawić też inne pole (np. inna ulica -> numer domu już pasuje).
    form.addEventListener("change", function () {
      form.querySelectorAll("[data-req].has-err").forEach(function (field) {
        if (!check(field)) clearErr(field);
      });
    });

    function collect() {
      var data = new URLSearchParams();
      data.set("form", kind);
      data.set("elapsed", String(Date.now() - started));
      data.set("website", form.querySelector('[name="website"]').value);
      var cats = [];
      form.querySelectorAll("input[name], textarea[name], select[name]").forEach(function (el) {
        if (el.name === "website" || !active(el)) return;
        if (el.type === "checkbox") {
          if (el.name === "kategorie") { if (el.checked) cats.push(el.value); }
          else data.set(el.name, el.checked ? "Tak" : "Nie");
        } else if (el.type === "radio") {
          if (el.checked) data.set(el.name, el.value);
        } else {
          data.set(el.name, el.value.trim());
        }
      });
      if (!cats.length && requests) cats = requests.categories();
      data.set("kategorie", cats.join(", "));
      return data;
    }

    function errorText(code) {
      switch (code) {
        case "too_fast": return "Formularz został wysłany zbyt szybko. Odczekaj chwilę i spróbuj ponownie.";
        case "rate_limited": return "Otrzymaliśmy teraz bardzo dużo zgłoszeń. Spróbuj ponownie za godzinę.";
        case "invalid": return "Uzupełnij wymagane pola i spróbuj ponownie.";
        case "closed": return "Zapisy zostały już zakończone.";
        case "taken": return "Ktoś przed chwilą zgłosił się do tej prośby. Lista została odświeżona: wybierz inną prośbę albo wyślij ogólną ofertę pomocy.";
        default: return "Coś poszło nie tak. Spróbuj ponownie za chwilę.";
      }
    }

    form.addEventListener("submit", function (e) {
      e.preventDefault();
      msg.textContent = "";
      if (!validate()) return;
      if (!url) {
        msg.textContent = "Formularz nie jest jeszcze podłączony. Spróbuj ponownie później.";
        return;
      }
      submitBtn.disabled = true;
      submitBtn.textContent = "Wysyłanie…";
      var picked = requests ? requests.selected() : "";
      fetch(url, { method: "POST", body: collect() })
        .then(function (r) { return r.json(); })
        .then(function (res) {
          if (res && res.ok) {
            var dg = done.querySelector("[data-done-general]");
            var dr = done.querySelector("[data-done-request]");
            if (dg && dr) { dg.hidden = !!picked; dr.hidden = !picked; }
            form.hidden = true;
            done.hidden = false;
            done.focus();
            window.scrollTo(0, 0);
            postHeight();
          } else {
            msg.textContent = errorText(res && res.error);
            if (res && res.error === "taken" && requests) requests.taken();
            if (res && res.error === "closed" && season) season.close();
          }
        })
        .catch(function () {
          msg.textContent = "Nie udało się wysłać formularza. Sprawdź połączenie z internetem i spróbuj ponownie.";
        })
        .then(function () {
          submitBtn.disabled = false;
          submitBtn.textContent = submitLabel;
        });
    });

    var again = done && done.querySelector("[data-again]");
    if (again) again.addEventListener("click", function () {
      form.reset();
      form.querySelectorAll("[data-req]").forEach(clearErr);
      msg.textContent = "";
      started = Date.now();
      applyTyp();
      done.hidden = true;
      form.hidden = false;
      if (requests) requests.reset();
      postHeight();
    });
  }
})();
