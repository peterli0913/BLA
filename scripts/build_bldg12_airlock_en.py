#!/usr/bin/env python3
"""English poster: Building 12 / Peptide Building 2 logistics-platform airlock.

Slide text becomes English. Floor-plan labels keep the Chinese and gain an
English line in nearby white space, without covering walls, doors, or pressure tags.
"""

import io
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

SRC = Path("/workspace/12号楼增加物流平台缓冲间-10-14(1).pptx")
OUT = Path("/workspace/12号楼增加物流平台缓冲间-10-14_EN.pptx")
FONT = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 15)
FONT_SM = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 13)
FONT_XS = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 11)
INK = (11, 59, 107)  # readable navy, distinct from the black CAD text

# "auto": search white space around the Chinese box.
# "at": draw at a fixed point, only if that patch is clear.
LABELS = {
    6: [  # before
        ("auto", (250, 88, 350, 116), "Logistics platform", FONT),
        ("auto", (210, 318, 310, 352), "Logistics corridor", FONT_SM),
        ("auto", (100, 520, 165, 590), "Vestibule", FONT),
        ("auto", (320, 455, 430, 560), "Elevator 1", FONT_SM),
        ("auto", (530, 430, 610, 475), "Stair 2", FONT_SM),
    ],
    8: [  # after
        ("auto", (360, 22, 470, 58), "Logistics platform", FONT_SM),
        ("at", 118, 108, "Airlock", FONT_SM),
        ("auto", (175, 220, 280, 265), "Logistics corridor", FONT_SM),
        ("at", 78, 448, "Vestibule", FONT_SM),
        ("at", 300, 292, "Elevator 1", FONT_SM),
        ("auto", (460, 310, 545, 370), "Stair 2", FONT_SM),
    ],
    10: [  # level 2. Elevator and vestibule have no clear patch; captions carry them.
        ("auto", (8, 115, 70, 175), "Air shaft", FONT_SM),
        ("auto", (355, 140, 435, 210), "Air shaft", FONT_SM),
        ("at", 148, 358, "Corridor", FONT_XS),
        ("at", 356, 304, "Elec. shaft", FONT_XS),
        ("auto", (255, 120, 360, 200), "Stair 2", FONT_SM),
    ],
    11: [  # level 3. Stair / exhaust / electrical already have English.
        ("at", 250, 294, "Vestibule", FONT_XS),
    ],
    12: [  # level 4
        ("at", 250, 294, "Vestibule", FONT_XS),
    ],
}

# Replace the CAD word "Tubewells" (a mistranslation of 排风管井) and nothing else.
TUBEWELLS = {
    (437, 382): [((4, 186, 84, 205), (6, 188)), ((334, 141, 430, 156), (338, 143))],
    (441, 386): [((4, 220, 86, 240), (6, 222))],
}


def is_clear_pixel(rgb):
    r, g, b = (int(v) for v in rgb)
    if r > 232 and g > 232 and b > 232:
        return True
    # Light-green room highlight on the after plan. Do not treat pink pressure tags as clear.
    return g > 220 and r > 120 and b < 180 and g > r and g > b


def ink_fraction(im, box):
    x1, y1, x2, y2 = [int(v) for v in box]
    x1, y1 = max(0, x1), max(0, y1)
    x2, y2 = min(im.width, x2), min(im.height, y2)
    if x2 - x1 < 4 or y2 - y1 < 4:
        return 1.0
    a = np.asarray(im.crop((x1, y1, x2, y2)))
    clear = np.array([[is_clear_pixel(px) for px in row] for row in a])
    return 1.0 - float(clear.mean())


def text_size(text, font):
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    b = dummy.textbbox((0, 0), text, font=font)
    return b[2] - b[0], b[3] - b[1]


def draw_at(im, x, y, text, font, limit=0.04):
    tw, th = text_size(text, font)
    box = (x - 1, y - 1, x + tw + 2, y + th + 2)
    if box[0] < 1 or box[1] < 1 or box[2] > im.width - 1 or box[3] > im.height - 1:
        return False
    if ink_fraction(im, box) > limit:
        return False
    ImageDraw.Draw(im).text((x, y), text, font=font, fill=INK)
    return True


def draw_label(im, anchor, text, font):
    """Put English in the first nearby patch that does not cover the drawing."""
    x1, y1, x2, y2 = anchor
    tw, th = text_size(text, font)
    cx = (x1 + x2) // 2
    cy = (y1 + y2) // 2
    candidates = [
        (cx - tw // 2, y2 + 3),
        (cx - tw // 2, y1 - th - 3),
        (x2 + 4, cy - th // 2),
        (x1 - tw - 4, cy - th // 2),
        (cx - tw // 2, y2 + th + 6),
        (cx - tw // 2, y1 - 2 * th - 6),
    ]
    for x, y in candidates:
        if draw_at(im, x, y, text, font):
            return True
    return False


def cover_tubewells(im):
    """Replace the CAD word Tubewells. Leave 排风管井 in place."""
    draw = ImageDraw.Draw(im)
    for rect, (x, y) in TUBEWELLS.get(im.size, []):
        draw.rectangle(rect, fill=(255, 255, 255))
        draw.text((x, y), "Exhaust shaft", font=FONT_XS, fill=INK)


def annotate(blob, shape_index):
    im = Image.open(io.BytesIO(blob)).convert("RGB")
    cover_tubewells(im)
    missed = []
    for item in LABELS.get(shape_index, []):
        if item[0] == "auto":
            _, box, text, font = item
            ok = draw_label(im, box, text, font)
        else:
            _, x, y, text, font = item
            ok = draw_at(im, x, y, text, font, limit=0.08)
        if not ok:
            missed.append(text)
    out = io.BytesIO()
    im.save(out, format="PNG")
    return out.getvalue(), missed, im


def replace_picture(shape, blob):
    blip = shape._element.blipFill.blip
    rid = blip.get(qn("r:embed"))
    shape.part.related_part(rid)._blob = blob


def set_para(paragraph, text, size=None, bold=None):
    runs = paragraph.runs
    if not runs:
        run = paragraph.add_run()
        run.text = text
    else:
        runs[0].text = text
        for run in runs[1:]:
            run.text = ""
        run = runs[0]
    run.font.name = "Arial"
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    rpr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        node = rpr.find(qn(f"a:{tag}"))
        if node is not None:
            node.set("typeface", "Arial")


def write_cell(cell, text, size, bold):
    tf = cell.text_frame
    tf.word_wrap = True
    paragraph = tf.paragraphs[0]
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    set_para(paragraph, text, size=size, bold=bold)


def build():
    prs = Presentation(str(SRC))
    slide = prs.slides[0]
    missed = {}
    for idx in LABELS:
        shape = slide.shapes[idx]
        blob, miss, _im = annotate(shape.image.blob, idx)
        replace_picture(shape, blob)
        if miss:
            missed[idx] = miss

    # Title
    set_para(slide.shapes[1].text_frame.paragraphs[0],
             "Peptide Building 2 — New Logistics Platform Airlock", size=32, bold=True)
    set_para(slide.shapes[2].text_frame.paragraphs[0],
             "Current state of the material-platform airlock", size=26, bold=True)
    set_para(slide.shapes[3].text_frame.paragraphs[0],
             "Material transfer flow", size=26, bold=True)
    for caption, picture, text in (
        (slide.shapes[7], slide.shapes[6], "Logistics platform before the change"),
        (slide.shapes[9], slide.shapes[8], "After the change (vestibule upgraded to a CNC area)"),
        (slide.shapes[14], slide.shapes[10], "Level 2 logistics corridor"),
        (slide.shapes[13], slide.shapes[11], "Level 3 logistics corridor"),
        (slide.shapes[15], slide.shapes[12], "Level 4 logistics corridor"),
    ):
        caption.left = picture.left
        caption.width = picture.width
        paragraph = caption.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        set_para(paragraph, text, size=14, bold=True)
    set_para(
        slide.shapes[16].text_frame.paragraphs[0],
        "Logistics corridors in the building are classified as general areas and are controlled to the CNC standard.",
        size=16,
        bold=True,
    )

    status = [
        ["Action", "Current status", "Target date"],
        ["Change-control document", "Approved", "N/A"],
        ["Airlock structure", "Construction complete", "N/A"],
        ["Access control", "Installed and entered in the system", "N/A"],
        ["HVAC", "Ductwork complete", "N/A"],
        ["Interlock", "Commissioning", "10/16/2026"],
        ["Differential pressure", "Commissioning", "10/18/2026"],
        ["Qualification", "IQ in progress", "10/20/2026"],
    ]
    table = slide.shapes[4].table
    for r, row in enumerate(status):
        for c, text in enumerate(row):
            write_cell(table.cell(r, c), text, 15 if r else 16, bold=True)

    flow = [
        ["Step", "Control"],
        ["Staffing", "Two dedicated operators: one external, one internal, for receipt and transfer."],
        ["Platform check", "The external operator checks the delivery record against the outer label (identity, quantity, and pack), then calls the internal operator and moves the material into the airlock."],
        ["Airlock check", "Only one material for one project at a time. The internal operator completes the second barcode check in the airlock."],
        ["Clean and transfer", "After the check, clean the material in the airlock and move it at once into the logistics corridor, then by the logistics elevator to the suite. Do not leave material staged."],
    ]
    table = slide.shapes[5].table
    heights = [Emu(520000), Emu(620000), Emu(980000), Emu(820000), Emu(980000)]
    for r, row in enumerate(flow):
        table.rows[r].height = heights[r]
        for c, text in enumerate(row):
            write_cell(table.cell(r, c), text, 14 if r else 16, bold=(r == 0 or c == 0))

    prs.save(str(OUT))
    print(OUT)
    print("missed labels", missed)


if __name__ == "__main__":
    build()
