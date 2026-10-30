# Literatur-Add-in für Word – Prototyp (Etappe 0)

Ziel des Projekts: ein Word-Add-in (Word im Browser, auch auf dem iPad), mit dem man
PDFs markiert, daraus Wissenselemente macht, sie einer Gliederung zuordnet, im
IEEE-Stil zitiert und **vom Zitat in Word direkt zur Stelle im PDF springt** – mit
KI-Unterstützung durch Claude.

Aktueller Stand: Prototyp für **Etappe 0**.

## Was Etappe 0 testet

| Test | Prüft |
|---|---|
| 1. Umgebung | Läuft das Add-in? Welche Word-API-Versionen gibt es auf dem Gerät? |
| 2. Zitat einfügen | Zitat als Inhaltssteuerelement mit unsichtbarer Kennung einfügen |
| 3. Vom Zitat zum PDF | Klick ins Zitat wird erkannt, PDF öffnet sich auf Seite 2 mit markierter Stelle |
| 4. Dokument lesen | Text und Zitate des Dokuments auslesen (Grundlage für den Claude-Zugriff) |

## Anleitung zum Testen

1. **GitHub Pages aktivieren** (einmalig): Repo auf GitHub → *Settings* → *Pages* →
   *Source: Deploy from a branch* → Branch `main`, Ordner `/ (root)` → *Save*.
   Das Repo muss dafür öffentlich sein (private Repos brauchen ein kostenpflichtiges
   GitHub-Konto). Im Repo liegt nur Programmcode – PDFs und Texte gehören nie hierher.
   Nach 1–2 Minuten sollte diese Seite ein PDF mit gelber Markierung zeigen:
   <https://felix7w.github.io/Literatur-AddIn-GHRep/viewer.html?file=test/beispiel.pdf&page=2&q=Gute%20Literaturverwaltung>
2. **Manifest herunterladen**: [`manifest.xml`](manifest.xml) öffnen → *Download raw file*.
3. **Word im Browser** öffnen (<https://www.office.com>, privates Microsoft-Konto) und ein
   neues leeres Dokument anlegen.
4. **Add-in hochladen**: *Start* → *Add-Ins* → *Weitere Add-Ins* → *Meine Add-Ins* →
   *Mein Add-In hochladen* → `manifest.xml` auswählen. (Die Menüs heißen je nach
   Word-Version leicht anders.)
5. Auf *Start* → **Literatur** klicken – die Seitenleiste öffnet sich.
6. Ein paar Wörter tippen, dann die Tests 2–4 der Reihe nach durchführen.
7. *Protokoll kopieren* und an Claude schicken.
8. Dasselbe auf dem **iPad in Safari** wiederholen.

## Aufbau

| Datei | Zweck |
|---|---|
| `manifest.xml` | Beschreibt das Add-in für Word (Name, Button, Adressen) |
| `taskpane.html` / `taskpane.js` | Seitenleiste mit den Tests |
| `viewer.html` / `viewer.js` | PDF-Ansicht mit Sprung zur Seite und Markierung (pdf.js) |
| `commands.html` | Pflichtdatei für den Button im Manifest |
| `vendor/pdfjs/` | pdf.js 4.10.38 (Legacy-Build für ältere Safari-Versionen), Apache-2.0 |
| `test/beispiel.pdf` | Test-PDF; das Test-Zitat verweist auf einen Satz auf Seite 2 |
| `tools/make-assets.py` | Erzeugt Icons und Test-PDF neu |

Lokal ansehen (ohne Word): `python3 -m http.server` im Repo-Ordner, dann
`http://localhost:8000/viewer.html?file=test/beispiel.pdf&page=2&q=Gute` öffnen.

## Gesamtplan

| # | Etappe | Ergebnis |
|---|---|---|
| **0** | Grundgerüst und Risiko-Test | Add-in läuft in Word im Browser und auf dem iPad (aktueller Stand) |
| 1 | Quellen | Projekte, Quellen per DOI/ISBN importieren, PDFs ablegen |
| 2 | PDF-Reader | Markieren (auch Apple Pencil) → Wissenselement (Typ, Titel, Text, Seite, Position) |
| 3 | Gliederung | Gliederung anlegen, Wissenselemente zuordnen |
| 4 | Zitieren in Word | IEEE-Belege (citeproc-js + CSL), automatische Neunummerierung, Literaturverzeichnis, Sprung ins PDF |
| 5 | Claude-Anbindung | Claude kann Dokument, PDFs und Markierungen lesen |
| 6 | Feinschliff | weitere Zitierstile, Offline-Zwischenspeicher, Import aus Citavi (RIS/BibTeX) |
