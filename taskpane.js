"use strict";

// Etappe 0: prüft, ob die Kernmechanik des Add-ins auf dem jeweiligen Gerät
// funktioniert – Zitat als Inhaltssteuerelement einfügen, Klick darauf erkennen,
// PDF an der richtigen Stelle öffnen und den Dokumenttext auslesen.

const TAG_PREFIX = "lit:cit:";

// Ein fest eingebautes Wissenselement, bis es echte Daten gibt (ab Etappe 1).
const DEMO = {
  id: "demo-1",
  label: "[1, S. 2]",
  title: "Beispielquelle (Test-PDF)",
  pdf: "test/beispiel.pdf",
  page: 2,
  quote: "Gute Literaturverwaltung spart beim Schreiben viel Zeit.",
};
const ELEMENTS = { [DEMO.id]: DEMO };

const $ = (id) => document.getElementById(id);
let selectedElement = null;

function log(message) {
  const time = new Date().toLocaleTimeString("de-DE");
  $("log").value += `[${time}] ${message}\n`;
  $("log").scrollTop = $("log").scrollHeight;
}

function describeError(err) {
  if (!err) return "unbekannter Fehler";
  const parts = [err.name, err.code, err.message].filter(Boolean);
  if (err.debugInfo && err.debugInfo.errorLocation) parts.push(`bei ${err.debugInfo.errorLocation}`);
  return parts.join(" – ");
}

function addEnv(label, value) {
  const dt = document.createElement("dt");
  dt.textContent = label;
  const dd = document.createElement("dd");
  dd.textContent = value;
  $("env").append(dt, dd);
  log(`${label}: ${value}`);
}

function viewerUrl(element) {
  const url = new URL("viewer.html", location.href);
  url.searchParams.set("file", element.pdf);
  url.searchParams.set("page", String(element.page));
  url.searchParams.set("q", element.quote);
  return url.toString();
}

// ---------- 1. Umgebung ----------

function reportEnvironment(info) {
  addEnv("Anwendung", String(info.host));
  addEnv("Plattform", String(info.platform));
  const diag = Office.context.diagnostics || {};
  if (diag.version) addEnv("Office-Version", diag.version);
  addEnv("Browser", navigator.userAgent);

  const supported = [];
  for (let minor = 1; minor <= 9; minor++) {
    if (Office.context.requirements.isSetSupported("WordApi", `1.${minor}`)) supported.push(`1.${minor}`);
  }
  addEnv("WordApi", supported.length ? supported.join(", ") : "keine");
  addEnv("Dialog-API", Office.context.requirements.isSetSupported("DialogApi", "1.1") ? "ja" : "nein");
}

// ---------- 2. Zitat einfügen ----------

async function insertCitation() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const range = selection.insertText(DEMO.label, Word.InsertLocation.replace);
      const cc = range.insertContentControl();
      cc.tag = TAG_PREFIX + DEMO.id;
      cc.title = DEMO.title;
      cc.appearance = Word.ContentControlAppearance.boundingBox;
      // Danach weiterschreiben können, ohne im Zitat zu landen.
      cc.getRange(Word.RangeLocation.after).select(Word.SelectionMode.end);
      await context.sync();
    });
    log("✅ Zitat eingefügt.");
  } catch (err) {
    log(`❌ Zitat einfügen fehlgeschlagen: ${describeError(err)}`);
  }
}

// ---------- 3. Klick auf Zitat erkennen ----------

let selectionTimer = null;

function onSelectionChanged() {
  // Beim Tippen feuert das Ereignis sehr oft – kurz abwarten.
  clearTimeout(selectionTimer);
  selectionTimer = setTimeout(checkSelection, 250);
}

async function checkSelection() {
  try {
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,title,text");
      await context.sync();

      const isCitation = !cc.isNullObject && cc.tag && cc.tag.startsWith(TAG_PREFIX);
      const element = isCitation ? ELEMENTS[cc.tag.slice(TAG_PREFIX.length)] : null;
      if (element === selectedElement) return;
      selectedElement = element;
      showSelected(element, isCitation ? cc.tag : null);
    });
  } catch (err) {
    log(`❌ Auswahl prüfen fehlgeschlagen: ${describeError(err)}`);
  }
}

function showSelected(element, tag) {
  const box = $("selected");
  if (element) {
    box.classList.remove("muted");
    box.textContent = "";
    const title = document.createElement("strong");
    title.textContent = element.title;
    const quote = document.createElement("div");
    quote.textContent = `„${element.quote}“ – S. ${element.page}`;
    box.append(title, quote);
    log(`✅ Zitat erkannt (${tag}).`);
  } else {
    box.classList.add("muted");
    box.textContent = tag ? `Unbekanntes Zitat (${tag}).` : "Kein Zitat ausgewählt.";
  }
  $("btn-open-pdf").disabled = !element;
}

// ---------- PDF öffnen ----------

function openPdf() {
  if (!selectedElement) return;
  const url = viewerUrl(selectedElement);
  // Im Browser eingebettet öffnen: umgeht Popup-Blocker (wichtig für Safari auf dem iPad).
  const inIframe = Office.context.platform === Office.PlatformType.OfficeOnline;
  Office.context.ui.displayDialogAsync(url, { height: 85, width: 70, displayInIframe: inIframe }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      log(`✅ PDF-Dialog geöffnet (${inIframe ? "eingebettet" : "eigenes Fenster"}).`);
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") dialog.close();
        else log(`Viewer meldet: ${arg.message}`);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        log(`Dialog-Ereignis: ${arg.error}`);
      });
    } else {
      log(`❌ Dialog fehlgeschlagen (${result.error.code}: ${result.error.message}) – versuche neues Fenster.`);
      const win = window.open(url, "_blank");
      log(win ? "✅ PDF in neuem Fenster geöffnet." : "❌ Auch das neue Fenster wurde blockiert.");
    }
  });
}

// ---------- 4. Dokument lesen ----------

async function readDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      const controls = body.contentControls;
      controls.load("items/tag,items/text");
      await context.sync();

      const citations = controls.items.filter((cc) => cc.tag && cc.tag.startsWith(TAG_PREFIX));
      const text = body.text.replace(/\s+/g, " ").trim();
      log(`✅ Dokument gelesen: ${text.length} Zeichen, ${citations.length} Zitat(e).`);
      log(`Anfang: ${text.slice(0, 200)}${text.length > 200 ? " …" : ""}`);
    });
  } catch (err) {
    log(`❌ Dokument lesen fehlgeschlagen: ${describeError(err)}`);
  }
}

// ---------- Protokoll kopieren ----------

async function copyLog() {
  const text = $("log").value;
  try {
    await navigator.clipboard.writeText(text);
    log("Protokoll kopiert.");
  } catch {
    // Zwischenablage ist in Add-in-Rahmen oft gesperrt – dann markieren zum manuellen Kopieren.
    $("log").focus();
    $("log").select();
    log("Automatisches Kopieren gesperrt – Text ist markiert, bitte manuell kopieren.");
  }
}

// ---------- Start ----------

$("btn-copy").addEventListener("click", copyLog);

if (typeof Office === "undefined") {
  $("not-in-word").hidden = false;
  log("office.js konnte nicht geladen werden.");
} else {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
      $("not-in-word").hidden = false;
      log("Nicht in Word geöffnet.");
      return;
    }
    reportEnvironment(info);

    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      log("❌ WordApi 1.3 fehlt – das Add-in braucht mindestens diese Version.");
      return;
    }

    $("btn-insert").disabled = false;
    $("btn-read").disabled = false;
    $("btn-insert").addEventListener("click", insertCitation);
    $("btn-open-pdf").addEventListener("click", openPdf);
    $("btn-read").addEventListener("click", readDocument);

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) log("✅ Auswahl-Ereignis registriert.");
      else log(`❌ Auswahl-Ereignis nicht verfügbar: ${result.error.message}`);
    });
  });
}
