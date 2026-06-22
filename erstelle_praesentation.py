from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pptx.util as util

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

# Farben
DARK_BLUE = RGBColor(0x1A, 0x23, 0x4E)   # Nachtblau
MID_BLUE  = RGBColor(0x2E, 0x4A, 0x7A)
ACCENT    = RGBColor(0x5B, 0xC8, 0xFA)   # Hellblau
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY= RGBColor(0xF0, 0xF4, 0xF8)
YELLOW    = RGBColor(0xFF, 0xD7, 0x00)

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, text, left, top, width, height,
                font_size=20, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    from pptx.util import Inches
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape

def add_bullet_slide(prs, title, bullets, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BLUE)
    # Accent bar links
    add_rect(slide, 0, 0, 0.08, 7.5, ACCENT)
    # Titel-Box
    add_rect(slide, 0.2, 0.2, 12.9, 1.1, MID_BLUE)
    add_textbox(slide, title, 0.35, 0.28, 12.5, 0.9,
                font_size=28, bold=True, color=ACCENT)
    if subtitle:
        add_textbox(slide, subtitle, 0.35, 1.1, 12.5, 0.4,
                    font_size=14, italic=True, color=LIGHT_GRAY)
    y = 1.6
    for bullet in bullets:
        add_textbox(slide, "▸  " + bullet, 0.5, y, 12.5, 0.55,
                    font_size=18, color=WHITE)
        y += 0.6
    return slide

# ── Folie 1: Titelfolie ──────────────────────────────────────────────────────
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide1, DARK_BLUE)
add_rect(slide1, 0, 0, 13.33, 0.12, ACCENT)
add_rect(slide1, 0, 7.38, 13.33, 0.12, ACCENT)
add_rect(slide1, 0, 2.5, 13.33, 2.8, MID_BLUE)
add_textbox(slide1, "Das große Buch vom Schlaf",
            0.5, 2.6, 12.33, 1.2,
            font_size=40, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_textbox(slide1, "Matthew Walker",
            0.5, 3.7, 12.33, 0.6,
            font_size=24, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide1, "Buchpräsentation • Kapitel 2–3",
            0.5, 4.3, 12.33, 0.5,
            font_size=16, italic=True, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
add_textbox(slide1, "Sachbuch • Goldmann Verlag • 2018",
            0.5, 6.7, 12.33, 0.5,
            font_size=13, color=ACCENT, align=PP_ALIGN.CENTER)

# ── Folie 2: Über den Autor ───────────────────────────────────────────────────
add_bullet_slide(prs,
    "Über den Autor: Matthew Walker",
    [
        "Neurowissenschaftler & Schlafforscher",
        "Professor an der UC Berkeley (Kalifornien, USA)",
        "Direktor des Center for Human Sleep Science",
        "Forscht seit über 20 Jahren zum Thema Schlaf",
        'Sein Ziel: eine weltweite "Schlafrevolution" ausloesen',
    ],
    subtitle="Wer steckt hinter dem Buch?"
)

# ── Folie 3: Worum geht es? ───────────────────────────────────────────────────
add_bullet_slide(prs,
    "Worum geht es?",
    [
        "Grundfrage: Warum müssen alle Lebewesen schlafen?",
        "Kapitel 2: Der innere Schlaf-Wach-Rhythmus (circadianer Rhythmus)",
        "Kapitel 2: Die Rolle von Melatonin & Koffein",
        "Kapitel 3: Was ist Schlaf wirklich? – Non-REM & REM",
        "Botschaft: Schlaf ist die wichtigste Gesundheitsinvestition",
    ],
    subtitle="Inhalt – Kapitel 2 & 3"
)

# ── Folie 4: Circadianer Rhythmus ─────────────────────────────────────────────
add_bullet_slide(prs,
    "Kapitel 2: Der circadiane Rhythmus",
    [
        "Unser Körper hat eine innere 24-Stunden-Uhr",
        "Gesteuert durch Tageslicht → bei Dunkelheit: Melatonin-Ausschüttung",
        "Melatonin = Signal für den Körper: Es ist Nacht",
        "Melatonin löst NICHT den Schlaf aus – es zeigt nur die Uhrzeit an",
        "Störungen durch: künstliches Licht, Bildschirme, Schichtarbeit",
    ],
    subtitle="Wie weiß unser Körper, wann er schlafen soll?"
)

# ── Folie 5: Koffein ──────────────────────────────────────────────────────────
add_bullet_slide(prs,
    "Kapitel 2: Koffein & Schlaf",
    [
        "Im Laufe des Tages sammelt sich Adenosin im Gehirn an → macht müde",
        "Koffein blockiert die Adenosin-Rezeptoren → unterdrückt Müdigkeit",
        "Wichtig: Koffein beseitigt die Müdigkeit NICHT – nur aufgeschoben!",
        "Wenn Koffein abbaut → aufgestaute Müdigkeit kommt auf einmal zurück",
        '= der bekannte "Kaffee-Absturz"',
        "Koffein ist das am meisten konsumierte psychoaktive Mittel der Welt",
    ],
    subtitle="Warum macht Kaffee wach – und was passiert danach?"
)

# ── Folie 6: Schlafphasen ─────────────────────────────────────────────────────
add_bullet_slide(prs,
    "Kapitel 3: Die Schlafphasen",
    [
        "Schlaf ist KEIN passiver Zustand – das Gehirn ist hochaktiv",
        "Non-REM-Schlaf (Tiefschlaf): Körper erholt sich, Gelerntes wird gespeichert",
        "REM-Schlaf (Traumschlaf): emotionale Verarbeitung, Kreativität",
        "Beide Phasen wechseln sich in 90-Minuten-Zyklen ab",
        "Beide sind unverzichtbar – keine Phase kann die andere ersetzen",
        "Schlafmangel beschädigt BEIDE Phasen",
    ],
    subtitle="Was passiert wirklich, wenn wir schlafen?"
)

# ── Folie 7: Themen & Botschaften ────────────────────────────────────────────
add_bullet_slide(prs,
    "Themen & Botschaft des Buches",
    [
        "Thema 1: Schlaf als Grundlage der Gesundheit – nicht verhandelbar",
        "Thema 2: Moderner Lebensstil zerstört unseren natürlichen Rhythmus",
        "Thema 3: Koffein – unterschätztes Schlafgift",
        "Kernbotschaft: Schlaf ist keine Schwäche – es ist biologische Notwendigkeit",
        'Walker kaempft gegen den Mythos: "Wenig schlafen = produktiv sein"',
    ],
    subtitle="Was will uns der Autor sagen?"
)

# ── Folie 8: Persönliche Meinung ─────────────────────────────────────────────
add_bullet_slide(prs,
    "Meine persönliche Meinung",
    [
        "Gut lesbar trotz wissenschaftlicher Tiefe – viele Alltagsbeispiele",
        "Lieblingspassage: die Koffein-Erklärung (Kap. 2) – hat meine Sicht verändert",
        "Walker wirkt authentisch – er hat selbst früher zu wenig geschlafen",
        "Wunsch: mehr praktische Tipps schon in den ersten Kapiteln",
        "Leseempfehlung: für jeden, der zu wenig schläft oder Kaffee trinkt 😄",
    ],
    subtitle="Was denke ich über das Buch?"
)

# ── Folie 9: Empfehlung ───────────────────────────────────────────────────────
add_bullet_slide(prs,
    "Empfehlung",
    [
        "✔  Wissenschaftlich fundiert – überzeugende Beweise aus echten Studien",
        "✔  Verständlich geschrieben – auch ohne Vorkenntnisse gut lesbar",
        "✔  Lebenswichtige Informationen – kann wirklich Gesundheit verbessern",
        "✔  Geeignet für: alle, die denken, 5–6 Stunden Schlaf sind genug",
        "✘  Manchmal zu viele Zahlen hintereinander – kann etwas trocken wirken",
    ],
    subtitle="Würde ich das Buch weiterempfehlen?"
)

# ── Folie 10: Abschluss ───────────────────────────────────────────────────────
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide10, DARK_BLUE)
add_rect(slide10, 0, 0, 13.33, 0.12, ACCENT)
add_rect(slide10, 0, 7.38, 13.33, 0.12, ACCENT)
add_rect(slide10, 1.5, 2.3, 10.33, 3.2, MID_BLUE)
add_textbox(slide10, "Danke fürs Zuhören!",
            1.5, 2.5, 10.33, 1.0,
            font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
add_textbox(slide10, '"Schlaf ist das maechtigste Mittel,\ndas wir haben, um Gehirn und Koerper zu erneuern."',
            1.5, 3.5, 10.33, 1.2,
            font_size=18, italic=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide10, "– Matthew Walker",
            1.5, 4.7, 10.33, 0.4,
            font_size=14, color=ACCENT, align=PP_ALIGN.CENTER)
add_textbox(slide10, "Habt ihr Fragen?",
            1.5, 5.3, 10.33, 0.5,
            font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

prs.save("/home/user/claude-code/Buchpraesentation_Schlaf.pptx")
print("Fertig!")
