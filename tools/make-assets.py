"""Erzeugt die Icons und das Test-PDF für den Etappe-0-Prototyp.

Aufruf (aus dem Repo-Wurzelverzeichnis):
    python3 tools/make-assets.py
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
TEST = ROOT / "test"

# Der Satz, den das Test-Zitat im PDF markiert (muss zu DEMO.quote in taskpane.js passen).
QUOTE = "Gute Literaturverwaltung spart beim Schreiben viel Zeit."


def make_icons():
    ASSETS.mkdir(exist_ok=True)
    for size in (16, 32, 64, 80, 128):
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        r = max(2, size // 6)
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=(31, 87, 165, 255))
        try:
            font = ImageFont.truetype("DejaVuSans-Bold.ttf", int(size * 0.42))
        except OSError:
            font = ImageFont.load_default()
        text = "[1]" if size >= 32 else "1"
        box = d.textbbox((0, 0), text, font=font)
        w, h = box[2] - box[0], box[3] - box[1]
        d.text(((size - w) / 2 - box[0], (size - h) / 2 - box[1]), text, font=font, fill="white")
        img.save(ASSETS / f"icon-{size}.png")


def make_test_pdf():
    TEST.mkdir(exist_ok=True)
    c = canvas.Canvas(str(TEST / "beispiel.pdf"), pagesize=A4)
    c.setTitle("Beispielquelle für den Add-in-Test")
    width, height = A4
    pages = [
        [
            "Beispielquelle – Seite 1",
            "",
            "Dieses PDF dient nur zum Testen des Literatur-Add-ins.",
            "Das Test-Zitat im Word-Dokument verweist auf eine Stelle auf Seite 2.",
        ],
        [
            "Beispielquelle – Seite 2",
            "",
            "Wissenschaftliches Arbeiten besteht zu einem großen Teil aus Lesen.",
            QUOTE,
            "Wer Zitate direkt beim Lesen sauber erfasst, muss später nicht suchen.",
        ],
        [
            "Beispielquelle – Seite 3",
            "",
            "Ende des Testdokuments.",
        ],
    ]
    for lines in pages:
        y = height - 80
        for i, line in enumerate(lines):
            c.setFont("Helvetica-Bold" if i == 0 else "Helvetica", 16 if i == 0 else 12)
            c.drawString(72, y, line)
            y -= 24
        c.showPage()
    c.save()


if __name__ == "__main__":
    make_icons()
    make_test_pdf()
    print("Icons und Test-PDF erzeugt.")
