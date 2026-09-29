#!/usr/bin/env python3
"""Chinese version of Water Supply Examples.pdf: figures untouched, explanatory text translated in place."""

import pymupdf

SRC = "/workspace/Water Supply Examples.pdf"
OUT = "/workspace/deliverables/Water_Supply_Examples_中文.pdf"
FONT_DIR = "/usr/share/fonts/truetype/wqy"

TEXT = "#0e4a65"
TEXT_RGB = tuple(int(TEXT[i:i + 2], 16) / 255 for i in (1, 3, 5))
BLUE = "#00b0f0"
ORANGE = "#ffc000"

CSS = f"""
@font-face {{font-family: cjk; src: url(wqy-microhei.ttc);}}
* {{font-family: cjk; font-size: 11pt; color: {TEXT}; text-align: center; line-height: 1.35;}}
p {{margin: 0 0 2pt 0;}}
b {{font-size: 12pt;}}
"""
HEADING = "操作顺序："
# (page, rect) around the English heading underline, which is line art rather than text
OLD_UNDERLINES = [(0, (270.5, 185, 272.7, 290.5))]

# (page, rect in unrotated page space covering the English paragraph, Chinese HTML)
BLOCKS = [
    (0, (455, 525, 565, 780),
     "<p>注：由于需要留出卡箍安装空间并避开焊接过渡区，空间有限，"
     "只能让三通的直通方向（上图，<span style='color:" + BLUE + "'>蓝色图示</span>）"
     "或分支方向（下图，<span style='color:" + ORANGE + "'>橙色图示</span>）"
     "其中之一做到最短距离，无法两者兼顾。因此，需要慎重确定最短距离支腿的布置位置。</p>"),
    (0, (258, 60, 522, 412),
     "<p><b>操作顺序：</b></p>"
     "<p>1）按现行做法，通过软管接口接入供水。</p>"
     "<p>2）经“Waste Common”路径，对公共总管进行使用前冲洗。</p>"
     "<p>3）关闭“Waste Common”路径，打开通往该用水点排放路径的阀门。"
     "例如，向“Client 1”供水时，打开通往“Waste 1”的路径。</p>"
     "<p>4）经用水点排放路径，对该用水点专用支管进行使用前冲洗。</p>"
     "<p>5）关闭用水点排放路径，打开通往用水点的路径，供给所需水量。"
     "加水结束后，关闭用水点供水路径，并关闭供水。</p>"
     "<p>6）按现行做法断开供水软管、排空并存放。</p>"
     "<p>7）用氮气将公共总管吹扫至“Waste Common”路径，再将用水点支管吹扫至用水点排放路径，"
     "最后吹扫 LPD 路径。关闭氮气供应。</p>"),
    (1, (458, 515, 515, 755),
     "<p>注：零滞留（Zero-Static，即“零死角”）阀门可取代“最短距离三通 + 标准阀门”的组合。</p>"),
    (2, (445, 512, 528, 760),
     "<p>注：ISG 阀门有多种朝向／结构形式，需根据现场实施情况仔细确认所需型号。"
     "该阀门可在两个方向上实现零滞留功能，并进一步隔离各用水点。</p>"),
]


def main():
    doc = pymupdf.open(SRC)
    archive = pymupdf.Archive(FONT_DIR)
    for pno in sorted({b[0] for b in BLOCKS}):
        page = doc[pno]
        rotation = page.rotation
        page.set_rotation(0)
        blocks = [b for b in BLOCKS if b[0] == pno]
        for _, rect, _ in blocks:
            page.add_redact_annot(pymupdf.Rect(rect), fill=False)
        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
        )
        for _, rect in [u for u in OLD_UNDERLINES if u[0] == pno]:
            page.add_redact_annot(pymupdf.Rect(rect), fill=False)
            page.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,
                graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
            )
        for _, rect, html in blocks:
            spare, scale = page.insert_htmlbox(
                pymupdf.Rect(rect), html, css=CSS, archive=archive, rotate=90
            )
            assert spare >= 0 and scale == 1, (pno, rect, spare, scale)
        # CSS underline is misplaced in rotated boxes, so draw it; text runs bottom-to-top here.
        for hit in page.search_for(HEADING):
            x = hit.x1 - 1.5
            page.draw_line((x, hit.y0), (x, hit.y1), color=TEXT_RGB, width=0.8)
        page.set_rotation(rotation)
    doc.set_metadata({**doc.metadata, "title": "Water Supply Examples（中文）"})
    doc.save(OUT, garbage=3, deflate=True)
    print(OUT)


if __name__ == "__main__":
    main()
