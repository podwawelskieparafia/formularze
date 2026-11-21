(function () {
  "use strict";

  // ===== KONFIGURACJA =====
  // Adres aplikacji internetowej Google Apps Script (kończy się na /exec).
  var SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzsSKqr0vUuMj4HB0Z5D8hJVTzVQKIUOZKLmRaMEwlgzYleFCrBjXU6YNagKACJ8sSWGw/exec";

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
  var list = document.querySelector("[data-list]");
  if (list) initList(list);

  function active(el) { return !el.closest("[hidden]"); }

  function initForm(form) {
    var started = Date.now();
    var kind = form.getAttribute("data-form");
    var submitBtn = form.querySelector('button[type="submit"]');
    var submitLabel = submitBtn.textContent;
    var msg = form.querySelector(".form-msg");
    var done = document.querySelector(".done");

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
      var inputs = field.querySelectorAll("input, textarea");
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
        var focusable = first.querySelector("input, textarea");
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
    form.addEventListener("change", function (e) {
      var field = e.target.closest("[data-req]");
      if (field && field.classList.contains("has-err") && !check(field)) clearErr(field);
    });

    function collect() {
      var data = new URLSearchParams();
      data.set("form", kind);
      data.set("elapsed", String(Date.now() - started));
      data.set("website", form.querySelector('[name="website"]').value);
      var cats = [];
      form.querySelectorAll("input[name], textarea[name]").forEach(function (el) {
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
      data.set("kategorie", cats.join(", "));
      return data;
    }

    function errorText(code) {
      switch (code) {
        case "too_fast": return "Formularz został wysłany zbyt szybko. Odczekaj chwilę i spróbuj ponownie.";
        case "rate_limited": return "Otrzymaliśmy teraz bardzo dużo zgłoszeń. Spróbuj ponownie za godzinę.";
        case "invalid": return "Uzupełnij wymagane pola i spróbuj ponownie.";
        default: return "Coś poszło nie tak. Spróbuj ponownie za chwilę.";
      }
    }

    form.addEventListener("submit", function (e) {
      e.preventDefault();
      msg.textContent = "";
      if (!validate()) return;
      if (!SCRIPT_URL) {
        msg.textContent = "Formularz nie jest jeszcze podłączony. Spróbuj ponownie później.";
        return;
      }
      submitBtn.disabled = true;
      submitBtn.textContent = "Wysyłanie…";
      fetch(SCRIPT_URL, { method: "POST", body: collect() })
        .then(function (r) { return r.json(); })
        .then(function (res) {
          if (res && res.ok) {
            form.hidden = true;
            done.hidden = false;
            done.focus();
            window.scrollTo(0, 0);
            postHeight();
          } else {
            msg.textContent = errorText(res && res.error);
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
      postHeight();
    });
  }

  function initList(box) {
    var typ = params.get("typ") === "chce" ? "chce" : "szukam";
    var title = box.querySelector(".list-title");
    var status = box.querySelector(".list-status");
    var ul = box.querySelector(".items");
    title.textContent = typ === "chce" ? "Oferty pomocy" : "Aktualne potrzeby";

    if (!SCRIPT_URL) { status.textContent = "Lista nie jest jeszcze podłączona."; return; }
    status.textContent = "Wczytywanie…";

    fetch(SCRIPT_URL + "?list=" + typ)
      .then(function (r) { return r.json(); })
      .then(function (res) {
        var items = (res && res.ok && res.items) || [];
        status.textContent = items.length ? "" : (typ === "chce"
          ? "Obecnie nie ma opublikowanych ofert pomocy."
          : "Obecnie nie ma opublikowanych potrzeb.");
        items.forEach(function (it) {
          var li = document.createElement("li");
          li.className = "item";
          var head = document.createElement("div");
          head.className = "item-head";
          var tags = document.createElement("div");
          tags.className = "tags";
          String(it.kategorie || "").split(",").forEach(function (c) {
            c = c.trim();
            if (!c) return;
            var t = document.createElement("span");
            t.className = "tag";
            t.textContent = c;
            tags.appendChild(t);
          });
          var date = document.createElement("span");
          date.className = "item-date";
          date.textContent = it.data || "";
          head.append(tags, date);
          var p = document.createElement("p");
          p.className = "item-text";
          p.textContent = it.opis;
          li.append(head, p);
          ul.appendChild(li);
        });
        postHeight();
      })
      .catch(function () {
        status.textContent = "Nie udało się wczytać listy. Spróbuj odświeżyć stronę.";
      });
  }
})();
