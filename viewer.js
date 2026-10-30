// PDF-Ansicht: öffnet ein PDF auf einer bestimmten Seite und hebt eine Textstelle hervor.
// Parameter: ?file=<pfad>&page=<nr>&q=<zitierter text>

import * as pdfjsLib from "./vendor/pdfjs/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdfjs/pdf.worker.min.mjs", import.meta.url).toString();

const params = new URLSearchParams(location.search);
const file = params.get("file");
const query = params.get("q") || "";
let pageNumber = Number(params.get("page")) || 1;
let pdf = null;

const $ = (id) => document.getElementById(id);

function setStatus(text) {
  $("status").textContent = text;
}

// Text vergleichbar machen: Kleinbuchstaben, Leerraum zu einem Leerzeichen.
function normalize(text) {
  return text.toLowerCase().replace(/\s+/g, " ").trim();
}

// Sucht `needle` im Textinhalt der Seite und liefert pro betroffenem Textstück
// den Anteil [start, end) als Bruch der Stücklänge.
function findMatch(items, needle) {
  const target = normalize(needle);
  if (!target) return [];

  // Seitentext normalisiert zusammensetzen und für jedes Zeichen merken,
  // aus welchem Textstück und von welcher Position es stammt.
  let text = "";
  const origin = [];
  const push = (ch, itemIndex, charIndex) => {
    if (/\s/.test(ch)) {
      if (text.length === 0 || text.endsWith(" ")) return;
      ch = " ";
    }
    text += ch.toLowerCase();
    origin.push([itemIndex, charIndex]);
  };
  items.forEach((item, i) => {
    for (let c = 0; c < item.str.length; c++) push(item.str[c], i, c);
    push(" ", i, item.str.length);
  });

  const at = text.indexOf(target);
  if (at < 0) return [];

  const spans = new Map();
  for (let k = at; k < at + target.length; k++) {
    const [i, c] = origin[k];
    if (c >= items[i].str.length) continue;
    const span = spans.get(i) || { start: c, end: c + 1 };
    span.start = Math.min(span.start, c);
    span.end = Math.max(span.end, c + 1);
    spans.set(i, span);
  }
  return [...spans].map(([i, { start, end }]) => ({
    item: items[i],
    from: start / items[i].str.length,
    to: end / items[i].str.length,
  }));
}

async function renderPage() {
  const page = await pdf.getPage(pageNumber);
  const container = $("page");
  const available = Math.min(document.documentElement.clientWidth - 32, 1000);
  const baseViewport = page.getViewport({ scale: 1 });
  const viewport = page.getViewport({ scale: available / baseViewport.width });
  const ratio = window.devicePixelRatio || 1;

  const canvas = document.createElement("canvas");
  canvas.width = Math.floor(viewport.width * ratio);
  canvas.height = Math.floor(viewport.height * ratio);
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;
  container.replaceChildren(canvas);
  container.style.width = `${viewport.width}px`;

  await page.render({
    canvasContext: canvas.getContext("2d"),
    viewport,
    transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : null,
  }).promise;

  $("page-label").textContent = `Seite ${pageNumber} / ${pdf.numPages}`;
  $("prev").disabled = pageNumber <= 1;
  $("next").disabled = pageNumber >= pdf.numPages;

  if (!query) {
    setStatus("");
    return;
  }
  const content = await page.getTextContent();
  const matches = findMatch(content.items, query);
  if (!matches.length) {
    setStatus("Textstelle auf dieser Seite nicht gefunden.");
    return;
  }

  let first = null;
  for (const { item, from, to } of matches) {
    const [, , c, d, e, f] = pdfjsLib.Util.transform(viewport.transform, item.transform);
    const height = Math.hypot(c, d);
    const width = item.width * viewport.scale;
    const mark = document.createElement("div");
    mark.className = "mark";
    mark.style.left = `${e + width * from}px`;
    mark.style.top = `${f - height}px`;
    mark.style.width = `${width * (to - from)}px`;
    mark.style.height = `${height * 1.2}px`;
    container.append(mark);
    first = first || mark;
  }
  first.scrollIntoView({ block: "center", behavior: "smooth" });
  setStatus("Textstelle markiert.");
}

async function main() {
  if (!file) {
    setStatus("Kein PDF angegeben.");
    return;
  }
  try {
    pdf = await pdfjsLib.getDocument(new URL(file, location.href).toString()).promise;
    pageNumber = Math.min(Math.max(pageNumber, 1), pdf.numPages);
    await renderPage();
  } catch (err) {
    setStatus(`Fehler: ${err.message}`);
  }
}

$("prev").addEventListener("click", () => { pageNumber--; renderPage(); });
$("next").addEventListener("click", () => { pageNumber++; renderPage(); });
main();
