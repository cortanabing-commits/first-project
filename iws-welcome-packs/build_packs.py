"""Build the IWS welcome packs (Word + PDF).

Run:  python3 build_packs.py
Needs: python-docx, and LibreOffice (soffice) for the PDF versions.

All brand colours live in THEME below. To match another brand, change the
hex values there and re-run; every heading, band, table and callout updates.
Anything in [square brackets] is a placeholder to fill in before sending.
"""

import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

OUT = Path(__file__).parent

THEME = {
    # Taken from the IWS logo and tax invoice template
    "primary": "1E3A5F",      # IWS navy: titles, heading bars, table headers
    "secondary": "E07A2E",    # IWS orange: section numbers, bullets, top bar
    "accent": "E07A2E",
    "orange_text": "C07034",  # darker orange for small text (readable on white)
    "tint": "F7F9FC",         # pale blue-grey: info callouts, zebra rows
    "warn_tint": "FDF1E7",    # pale orange: "important" callouts, pills
    "line": "E3E8EF",         # table borders
    "text": "1F2933",
    "muted": "5B6770",
    "font": "Inter",
}

COMPANY = "Interactive Workforce Solutions"
SHORT = "IWS"
LOGO = Path(__file__).parent / "iws_logo.png"
ADDRESS = "2 Butterleaf Way, North Kellyville NSW 2155"
ABN = "15 373 790 652"
EMAIL = "michael.sibanda@interactivews.com"
PHONE = "+61 493 716 007"
VERSION = "Version 1.0  |  October 2026"


def rgb(hexstr):
    return RGBColor.from_string(hexstr)


# ---------- low-level helpers ----------

def shade(cell, hexstr):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexstr)
    tcPr.append(shd)


def cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("bottom", bottom), ("start", left), ("end", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def table_borders(table, color="E3E8EF", size=4, inside=True):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    for e in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{e}")
        if e in edges:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def no_borders(table):
    table_borders(table, inside=False, size=0, color="FFFFFF")
    for e in table._tbl.tblPr.find(qn("w:tblBorders")):
        e.set(qn("w:val"), "nil")


def left_bar(cell, hexstr, size=36):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for e in ["top", "bottom", "right"]:
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "nil")
        b.append(el)
    el = OxmlElement("w:left")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(size))
    el.set(qn("w:color"), hexstr)
    b.append(el)
    tcPr.append(b)


def keep_row_together(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    trPr.append(el)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def set_col_widths(table, widths_cm):
    """Fixed layout plus explicit grid, so Word and LibreOffice both honour widths."""
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        old = tblPr.find(qn(tag))
        if old is not None:
            tblPr.remove(old)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tw.set(qn("w:type"), "dxa")
    tblPr.append(tw)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)


def add_field(run, instr):
    for tag, text in (("begin", None), ("instr", instr), ("separate", None), ("end", None)):
        if tag == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        run._r.append(el)


def write_rich(par, text, size=10.5, color=None, bold=False, italic=False):
    """Write text into a paragraph. **bold** spans and [placeholders] are styled."""
    color = color or THEME["text"]
    import re
    tokens = re.split(r"(\*\*.+?\*\*|\[[^\]]+\])", text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            r = par.add_run(tok[2:-2])
            r.bold = True
            r.font.color.rgb = rgb(color)
        elif tok.startswith("[") and tok.endswith("]"):
            r = par.add_run(tok)
            r.bold = True
            r.font.color.rgb = rgb("6B4E00")
            rPr = r._r.get_or_add_rPr()
            hl = OxmlElement("w:shd")
            hl.set(qn("w:val"), "clear")
            hl.set(qn("w:color"), "auto")
            hl.set(qn("w:fill"), "FFF2A8")
            rPr.append(hl)
        else:
            r = par.add_run(tok)
            r.bold = bold
            r.font.color.rgb = rgb(color)
        r.italic = italic
        r.font.size = Pt(size)
    return par


# ---------- document-level building blocks ----------

class Pack:
    def __init__(self, title, audience_label):
        self.doc = Document()
        self.title = title
        self.audience = audience_label
        self.section_no = 0
        self._styles()

    def _styles(self):
        st = self.doc.styles["Normal"]
        st.font.name = THEME["font"]
        st.element.rPr.rFonts.set(qn("w:eastAsia"), THEME["font"])
        st.font.size = Pt(10.5)
        st.font.color.rgb = rgb(THEME["text"])
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.line_spacing = 1.15
        for name, size, color in (("Heading 1", 18, THEME["primary"]),
                                  ("Heading 2", 13, THEME["orange_text"])):
            h = self.doc.styles[name]
            h.font.name = THEME["font"]
            h.element.rPr.rFonts.set(qn("w:ascii"), THEME["font"])
            h.element.rPr.rFonts.set(qn("w:hAnsi"), THEME["font"])
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.color.rgb = rgb(color)
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(4)
            h.paragraph_format.keep_with_next = True
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        for side in ("left_margin", "right_margin"):
            setattr(sec, side, Cm(2.0))
        sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.6)

    # --- cover (mirrors the IWS tax invoice layout) ---
    def brand_bar(self, container, height_pt=3):
        t = container.add_table(rows=1, cols=2) if container is self.doc else container.add_table(1, 2, Cm(17))
        no_borders(t)
        a, b = t.rows[0].cells
        shade(a, THEME["primary"])
        shade(b, THEME["secondary"])
        set_col_widths(t, [12.4, 4.6])
        for c in (a, b):
            cell_margins(c, 0, 0, 0, 0)
            par = c.paragraphs[0]
            par.paragraph_format.space_after = Pt(0)
            par.paragraph_format.line_spacing = Pt(height_pt)
            par.add_run("").font.size = Pt(1)
        return t

    def cover(self, kicker, title, subtitle, bullets):
        d = self.doc
        sec = d.sections[0]
        sec.different_first_page_header_footer = True
        self.brand_bar(d, 5)

        # logo + company block
        t = d.add_table(rows=1, cols=2)
        no_borders(t)
        set_col_widths(t, [4.5, 12.5])
        lc, rc = t.rows[0].cells
        cell_margins(lc, 400, 0, 0, 0)
        cell_margins(rc, 400, 0, 0, 0)
        lc.paragraphs[0].add_run().add_picture(str(LOGO), width=Cm(3.4))
        rp = rc.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = rp.add_run(COMPANY)
        r.bold, r.font.size = True, Pt(15)
        r.font.color.rgb = rgb(THEME["primary"])
        for line in (ADDRESS, f"ABN {ABN}", f"{EMAIL}  ·  {PHONE}"):
            q = rc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            q.paragraph_format.space_after = Pt(1)
            r = q.add_run(line)
            r.font.size = Pt(9.5)
            r.font.color.rgb = rgb(THEME["muted"])

        for _ in range(6):
            d.add_paragraph()

        # orange pill kicker
        pt = d.add_table(rows=1, cols=1)
        no_borders(pt)
        set_col_widths(pt, [6.2])
        pc = pt.rows[0].cells[0]
        shade(pc, THEME["warn_tint"])
        cell_margins(pc, 60, 60, 180, 180)
        r = pc.paragraphs[0].add_run(kicker)
        r.bold, r.font.size = True, Pt(10.5)
        r.font.color.rgb = rgb(THEME["orange_text"])
        pc.paragraphs[0].paragraph_format.space_after = Pt(0)

        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(title)
        r.bold, r.font.size = True, Pt(40)
        r.font.color.rgb = rgb(THEME["primary"])
        p = d.add_paragraph()
        r = p.add_run(subtitle)
        r.font.size = Pt(14)
        r.font.color.rgb = rgb(THEME["muted"])

        for _ in range(11):
            d.add_paragraph()

        # navy "inside this pack" panel, styled like the invoice's Amount Due box
        t2 = d.add_table(rows=1, cols=1)
        no_borders(t2)
        set_col_widths(t2, [17.0])
        c2 = t2.rows[0].cells[0]
        shade(c2, THEME["primary"])
        cell_margins(c2, 300, 300, 400, 400)
        p = c2.paragraphs[0]
        r = p.add_run("INSIDE THIS PACK")
        r.font.size, r.bold = Pt(9), True
        r.font.color.rgb = rgb("F3B37E")
        rPr = r._r.get_or_add_rPr()
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), "30")
        rPr.append(sp)
        p = c2.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("   ·   ".join(bullets))
        r.font.size = Pt(11)
        r.font.color.rgb = rgb("FFFFFF")

        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        r = p.add_run(VERSION)
        r.font.size = Pt(9)
        r.font.color.rgb = rgb(THEME["muted"])

    def header_footer(self):
        sec = self.doc.sections[0]
        hdr = sec.header
        self.brand_bar(hdr, 3)
        # logo + name on the left, pack title on the right
        hdr._element.remove(hdr.paragraphs[0]._p)
        ht = hdr.add_table(1, 2, Cm(17))
        no_borders(ht)
        set_col_widths(ht, [9.0, 8.0])
        lc, rc = ht.rows[0].cells
        for c in (lc, rc):
            cell_margins(c, 100, 60, 0, 0)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
        lp = lc.paragraphs[0]
        lp.add_run().add_picture(str(LOGO), height=Cm(0.75))
        r = lp.add_run("  " + COMPANY)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = rgb(THEME["primary"])
        rp = rc.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = rp.add_run(self.title)
        r.font.size = Pt(9)
        r.font.color.rgb = rgb(THEME["muted"])
        hdr.add_paragraph().paragraph_format.space_after = Pt(0)

        fp = sec.footer.paragraphs[0]
        fp.paragraph_format.tab_stops.add_tab_stop(Cm(16.9), WD_TAB_ALIGNMENT.RIGHT)
        r = fp.add_run(f"{COMPANY}  ·  ABN {ABN}  ·  {EMAIL}  ·  {PHONE}")
        r.font.size = Pt(8)
        r.font.color.rgb = rgb(THEME["muted"])
        r = fp.add_run("\tPage ")
        r.font.size = Pt(8.5)
        r.bold = True
        r.font.color.rgb = rgb(THEME["orange_text"])
        r = fp.add_run()
        r.font.size = Pt(8.5)
        r.bold = True
        r.font.color.rgb = rgb(THEME["orange_text"])
        add_field(r, "PAGE")
        pPr = fp._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        bt = OxmlElement("w:top")
        bt.set(qn("w:val"), "single")
        bt.set(qn("w:sz"), "4")
        bt.set(qn("w:color"), THEME["line"])
        bt.set(qn("w:space"), "6")
        bdr.append(bt)
        pPr.append(bdr)

    # --- content blocks ---
    def h1(self, text, new_page=False):
        self.section_no += 1
        t = self.doc.add_table(rows=1, cols=2)
        no_borders(t)
        a, b = t.rows[0].cells
        shade(a, THEME["secondary"])
        shade(b, THEME["primary"])
        cell_margins(a, 120, 120, 120, 120)  # orange number, navy title, as on the invoice
        cell_margins(b, 120, 120, 220, 120)
        set_col_widths(t, [1.4, 15.6])
        pa = a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pa.add_run(f"{self.section_no:02d}")
        r.bold, r.font.size = True, Pt(14)
        r.font.color.rgb = rgb("FFFFFF")
        if new_page:
            pa.paragraph_format.page_break_before = True
        pb = b.paragraphs[0]
        r = pb.add_run(text)
        r.bold, r.font.size = True, Pt(15)
        r.font.color.rgb = rgb("FFFFFF")
        for c in (a, b):
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def h2(self, text):
        self.doc.add_heading(text, level=2)

    def p(self, text, **kw):
        return write_rich(self.doc.add_paragraph(), text, **kw)

    def bullets(self, items, mark="•", color=None):
        color = color or THEME["secondary"]
        for it in items:
            par = self.doc.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.6)
            par.paragraph_format.first_line_indent = Cm(-0.45)
            par.paragraph_format.space_after = Pt(3)
            r = par.add_run(mark + "  ")
            r.bold = True
            r.font.color.rgb = rgb(color)
            write_rich(par, it)

    def checklist(self, items):
        self.bullets(items, mark="☐", color=THEME["primary"])

    def steps(self, items):
        t = self.doc.add_table(rows=0, cols=2)
        no_borders(t)
        for i, it in enumerate(items, 1):
            row = t.add_row()
            keep_row_together(row)
            a, b = row.cells
            shade(a, THEME["primary"])
            cell_margins(a, 80, 80, 60, 60)
            cell_margins(b, 80, 80, 160, 60)
            pa = a.paragraphs[0]
            pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = pa.add_run(str(i))
            r.bold, r.font.size = True, Pt(12)
            r.font.color.rgb = rgb("FFFFFF")
            a.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            write_rich(b.paragraphs[0], it)
            spacer = t.add_row()
            for c in spacer.cells:
                c.paragraphs[0].paragraph_format.space_after = Pt(0)
                c.paragraphs[0].runs or c.paragraphs[0].add_run("").font.size
                c.paragraphs[0].paragraph_format.line_spacing = Pt(4)
        set_col_widths(t, [0.9, 16.1])
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def callout(self, title, body, kind="info"):
        fill = THEME["tint"] if kind == "info" else THEME["warn_tint"]
        bar = THEME["primary"] if kind == "info" else THEME["secondary"]
        t = self.doc.add_table(rows=1, cols=1)
        no_borders(t)
        c = t.rows[0].cells[0]
        set_col_widths(t, [17.0])
        shade(c, fill)
        left_bar(c, bar)
        cell_margins(c, 140, 140, 240, 200)
        keep_row_together(t.rows[0])
        p = c.paragraphs[0]
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = rgb(THEME["primary"])
        body = body if isinstance(body, list) else [body]
        for b in body:
            write_rich(c.add_paragraph(), b, size=10)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def table(self, header, rows, widths):
        t = self.doc.add_table(rows=1, cols=len(header))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_borders(t)
        hr = t.rows[0]
        repeat_header(hr)
        for i, h in enumerate(header):
            c = hr.cells[i]
            shade(c, THEME["primary"])
            cell_margins(c, 90, 90, 120, 120)
            r = c.paragraphs[0].add_run(h)
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = rgb("FFFFFF")
        for n, row in enumerate(rows):
            tr = t.add_row()
            keep_row_together(tr)
            for i, val in enumerate(row):
                c = tr.cells[i]
                cell_margins(c, 80, 80, 120, 120)
                if n % 2 == 1:
                    shade(c, "F4F8F9")
                par = c.paragraphs[0]
                par.paragraph_format.space_after = Pt(0)
                write_rich(par, val, size=9.5, bold=(i == 0 and len(header) > 2))
        set_col_widths(t, widths)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def signature_block(self, lines):
        t = self.doc.add_table(rows=0, cols=2)
        table_borders(t)
        for label in lines:
            row = t.add_row()
            keep_row_together(row)
            a, b = row.cells
            shade(a, THEME["tint"])
            cell_margins(a, 180, 180, 140, 140)
            cell_margins(b, 180, 180, 140, 140)
            r = a.paragraphs[0].add_run(label)
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = rgb(THEME["primary"])
        set_col_widths(t, [5.5, 11.5])

    def save(self, stem):
        self.header_footer()
        path = OUT / f"{stem}.docx"
        self.doc.save(path)
        return path


# =====================================================================
# Pack 1: Contractors and Support Workers
# =====================================================================

def worker_pack():
    k = Pack("Welcome Pack: Contractors and Support Workers", "Worker Welcome Pack")
    k.cover(
        "Welcome to the team",
        "Contractor and Support Worker Welcome Pack",
        "Everything you need before, during and after your first shift with IWS.",
        ["About IWS", "Your engagement", "Before your first shift", "Connecteam",
         "Code of Conduct", "Safety and incidents", "Pay", "Key contacts", "Sign-off"],
    )

    # 1 Welcome
    k.h1("Welcome to IWS", new_page=True)
    k.p("Hi [Worker first name],")
    k.p("Welcome to Interactive Workforce Solutions. You are joining a team that places skilled, "
        "caring support workers with disability and community care providers across [service region, e.g. Sydney's Hills District and surrounds]. "
        "The people you support are at the centre of everything we do, and so are you. We want you to feel "
        "prepared, safe and backed up on every shift.")
    k.p("This pack explains how we work, what we expect, and what you can expect from us. Please read it "
        "all, keep it handy, and sign the acknowledgement on the last page.")
    k.p("If anything is unclear, call us. No question is too small.")
    k.p("**Michael Sibanda**\nDirector, Interactive Workforce Solutions")

    k.callout("Our promise to you", [
        "**Clear information** before every placement: site, contact, participant needs.",
        "**Real support on shift:** an on-call person you can reach whenever you are working.",
        "**Paid correctly and on time,** every time.",
        "**A safe workplace.** You will never be asked to do work you are not trained or cleared for.",
    ])

    # 2 About
    k.h1("About IWS")
    k.h2("What we do")
    k.p("IWS supplies casual support workers and independent contractors to registered NDIS providers and "
        "other care organisations (our **customers**). You will usually work at a customer's site, such as a "
        "supported independent living house, a day program or in the community, following that customer's "
        "support plans and site rules.")
    k.h2("Our values")
    k.table(["Value", "What it looks like on shift"], [
        ["Dignity first", "The person you support makes the choices about their own life. You support, you do not take over."],
        ["Safe every time", "If something is unsafe, you stop, make the person safe and speak up. We will back you."],
        ["Reliable", "You turn up on time, clock in, and finish the shift you accepted. If you can't, you tell us early."],
        ["Honest", "You record what actually happened, including mistakes. Honest notes protect everyone."],
        ["Learning", "Ask questions, do the training, and tell us what would help you do the job better."],
    ], [4.0, 13.0])

    # 3 Engagement type
    k.h1("Your engagement: casual employee or contractor")
    k.p("IWS engages people in one of two ways. Your offer letter or agreement says which applies to you. "
        "The rules are different, so read the column that matches your engagement.")
    k.table(["Topic", "Casual employee", "ABN contractor"], [
        ["Agreement", "Casual employment contract", "Independent contractor agreement"],
        ["Pay rates", "SCHADS Award (Social, Community, Home Care and Disability Services Industry Award 2010), at your classification level, plus penalty rates and allowances", "Agreed rate in your contractor agreement. You invoice IWS."],
        ["Tax", "IWS withholds tax (PAYG). Give us a TFN declaration.", "You manage your own tax. IWS does not withhold unless you have not quoted an ABN."],
        ["GST", "Not applicable", "Add GST only if you are registered for GST. Tell us your GST status."],
        ["Superannuation", "IWS pays 12% super into your fund. Since 1 July 2026, super is paid with each pay run (Payday Super).", "Usually your own responsibility. Contractors paid mainly for their labour may still be owed super by law, so IWS assesses this for each contractor."],
        ["Insurance", "Covered by IWS workers compensation insurance", "You must hold your own public liability and personal accident (income protection) cover and give us certificates."],
        ["Timesheets", "Clock in and out in Connecteam", "Clock in and out in Connecteam, then invoice from the approved hours"],
        ["Turning down shifts", "You may decline shifts that are offered", "You may decline shifts that are offered"],
    ], [3.2, 6.9, 6.9])
    k.callout("Not sure which you are?",
              "Ask us before your first shift. Being treated as the wrong type of worker can affect your pay, "
              "super, tax and insurance. NSW law can treat some contractors as workers for workers compensation, "
              "so we check each arrangement.", kind="warn")

    # 4 Before first shift
    k.h1("Before your first shift")
    k.p("You cannot be rostered until every item below is uploaded to Connecteam and checked by IWS. "
        "We record expiry dates and remind you before anything runs out.")
    k.h2("Everyone")
    k.checklist([
        "**NDIS Worker Screening Check** cleared, and linked to IWS as an employer",
        "**NDIS Worker Orientation Module** \"Quality, Safety and You\" completed (certificate)",
        "Signed **NDIS Code of Conduct** acknowledgement",
        "Current **First Aid (HLTAID011)** and **CPR** (CPR refreshed every 12 months)",
        "**Infection prevention and control** training, and any vaccinations a customer requires",
        "**Manual handling** training",
        "Signed **confidentiality and privacy** agreement",
        "**Working With Children Check**, if you may support anyone under 18",
        "**Driver licence, registration and comprehensive insurance**, if you will drive participants",
        "Proof of your **right to work in Australia**",
        "Emergency contact details",
    ])
    k.h2("Casual employees also need")
    k.checklist(["Signed employment contract", "Tax file number declaration",
                 "Superannuation choice form (or we pay into our default fund)", "Bank details"])
    k.h2("Contractors also need")
    k.checklist(["Signed contractor agreement", "Active ABN (we check it on ABN Lookup)",
                 "Certificate of currency: public liability insurance",
                 "Certificate of currency: personal accident or income protection cover",
                 "GST registration status", "Bank details for invoice payments"])
    k.callout("Some sites need extra training",
              "Medication support, mealtime and swallowing (dysphagia) support, epilepsy, PEG feeding and "
              "positive behaviour support are common examples. We will only place you where your training "
              "matches the people you support.")

    # 5 Connecteam
    k.h1("Connecteam: your work app")
    k.p("IWS uses **Connecteam** for rosters, clocking in and out, documents and messages. "
        "Download it from the App Store or Google Play and log in with the invite we send you.")
    k.steps([
        "**Accept your invite** and complete the onboarding pack: upload every document on the checklist.",
        "**Check your shifts.** Accept or decline offered shifts as soon as you can. Accepted shifts appear in your schedule.",
        "**Clock in when you arrive** at the site, not before. Location must be on so your clock-in is matched to the site.",
        "**Clock out when you finish.** Add a short note if the shift ran over or something changed.",
        "**Read messages and updates.** Site changes, policies and reminders are sent through the app.",
        "**Keep documents current.** Upload renewed certificates before they expire, or you can't be rostered.",
    ])
    k.callout("Why clocking in matters",
              "Your clock-in and clock-out times are your timesheet. They decide your pay, and they are how we "
              "bill the customer. Forgotten clock-ins delay pay. If you forget, message IWS the same day.",
              kind="warn")

    # 6 First shift
    k.h1("Your first shift at a customer site")
    k.h2("The day before")
    k.bullets([
        "Check the site address, parking, entry instructions and start time in Connecteam.",
        "Save the **shift lead's number** and the **IWS on-call number** in your phone.",
        "Plan to arrive **10 minutes early**. Bring your ID, phone charger and water.",
    ])
    k.h2("On arrival: your site induction")
    k.p("The shift lead or house manager will induct you using the **IWS Customer Site Induction Checklist**. "
        "You must not work alone until it is signed off. It covers:")
    k.bullets([
        "Meeting the people you support, how each person likes to be greeted and communicates",
        "Who is in charge, sign-in, keys, door codes and secure areas",
        "Fire exits, the evacuation plan and each person's evacuation needs",
        "Where to find support plans, risk assessments, health plans and progress notes",
        "Site safety: manual handling equipment, PPE, chemicals, working alone",
    ])
    k.callout("Read the plans before you support",
              "Read every plan for each person you support before you give that support: support plan, "
              "risk assessment, health and mealtime plans, and any behaviour support plan. If you have not "
              "been shown a plan, ask for it. If you are still unsure, call IWS.")

    # 7 Code of Conduct
    k.h1("The NDIS Code of Conduct")
    k.p("Every person delivering NDIS supports must follow the NDIS Code of Conduct. Breaches can lead to "
        "bans from working in the sector, fines and other penalties. In plain language, you must:")
    k.table(["#", "The rule", "What it means for you"], [
        ["1", "Respect people's rights to freedom of expression, self-determination and decision-making", "Offer choices, wait for answers, and respect a \"no\"."],
        ["2", "Respect people's privacy", "Knock before entering. Keep personal information private. No photos on your personal phone."],
        ["3", "Provide supports and services in a safe and competent way, with care and skill", "Follow the plans. Only do tasks you are trained and authorised to do."],
        ["4", "Act with integrity, honesty and transparency", "Record what really happened. Never accept money, gifts or loans from participants."],
        ["5", "Promptly take steps to raise and act on concerns about quality and safety", "If something isn't right, tell the shift lead and IWS straight away."],
        ["6", "Take all reasonable steps to prevent and respond to violence, exploitation, neglect and abuse", "Make the person safe, call 000 if needed, then report. Do not stay silent."],
        ["7", "Take all reasonable steps to prevent and respond to sexual misconduct", "Keep professional boundaries at all times. Report any concern immediately."],
    ], [0.8, 7.6, 8.6])

    k.h2("Professional boundaries")
    k.bullets([
        "Do not arrange private work, paid or unpaid, with participants, families or customers. All work goes through IWS.",
        "Do not connect with participants or families on personal social media.",
        "Do not post anything about the people you support, sites or customers online.",
        "Do not bring visitors, children or pets to a shift.",
        "Do not lend or borrow money, or sell anything to participants.",
        "No alcohol or drugs before or during a shift. Tell us about any medicine that could affect your work.",
    ])

    # 8 Safety
    k.h1("Working safely")
    k.p("Under the NSW Work Health and Safety Act 2011, IWS and the customer both have a duty to keep you safe, "
        "and we work together on it. Your duties are to:")
    k.bullets([
        "**Take reasonable care** of your own health and safety",
        "**Take reasonable care** that what you do (or don't do) doesn't harm others",
        "**Follow reasonable safety instructions** and site policies",
    ])
    k.callout("You can stop unsafe work",
              "If a task poses a serious and immediate risk, make the person you support safe, stop the task, "
              "and call the shift lead and IWS straight away. You will not be penalised for raising a genuine safety concern.",
              kind="warn")
    k.table(["Risk", "What to do"], [
        ["Manual handling", "Use only the transfer method and equipment in the person's plan. No unplanned lifting. Ask for help."],
        ["Behaviours of concern", "Know the person's early warning signs and the strategies in their plan. Know your exit. Step back when unsafe."],
        ["Infection control", "Wash or sanitise hands before and after support. Use gloves and PPE as shown. Stay home if you are unwell."],
        ["Working alone", "Keep your phone charged. Follow the site's check-in process. Call on-call if you feel unsafe."],
        ["Fatigue", "Take your breaks. Tell us if you are too tired to work safely. Don't accept shifts that leave no rest."],
        ["Driving", "Do a pre-start check, use seatbelts and wheelchair restraints, and never use your phone while driving."],
        ["Chemicals", "Use only labelled products, as directed. Know where the safety data sheets are kept."],
    ], [4.0, 13.0])

    # 9 Incidents
    k.h1("Incidents, injuries and concerns")
    k.p("An **incident** is anything that harmed, or could have harmed, a person: a fall, a medication error, "
        "an injury, a behaviour that hurt someone, abuse, neglect, or a near miss.")
    k.steps([
        "**Make everyone safe.** Give first aid if needed.",
        "**Call 000** for any emergency: serious injury, a medical emergency, or a crime in progress.",
        "**Tell the shift lead** immediately.",
        "**Call IWS on-call** immediately. Do not wait until the end of the shift.",
        "**Write it down** the same shift, on the customer's incident form. Write facts only: what you saw, heard and did, with times.",
        "**Keep it confidential.** Only discuss it with the customer, IWS or the authorities.",
    ])
    k.callout("Serious incidents have legal time limits",
              ["Some incidents (for example a death, serious injury, abuse or neglect, unlawful sexual or physical "
               "contact, sexual misconduct, or an unauthorised restrictive practice) are **reportable incidents**. "
               "The registered provider must notify the NDIS Commission, usually within 24 hours.",
               "Your job is to tell the shift lead and IWS **straight away** so this can happen on time."],
              kind="warn")
    k.h2("If you are injured")
    k.bullets([
        "Get first aid or medical treatment.",
        "Tell the shift lead and IWS the **same day**, even if it seems minor.",
        "Casual employees: IWS will guide you through a workers compensation claim.",
        "Contractors: notify your own personal accident insurer, and tell us.",
    ])

    # 10 Pay
    k.h1("Pay, invoices and shifts")
    k.table(["", "Casual employees", "Contractors"], [
        ["How hours are recorded", "Connecteam clock-in and clock-out", "Connecteam clock-in and clock-out"],
        ["Pay or invoice cycle", "[Weekly / fortnightly], period ends [day]", "Invoice IWS [weekly / fortnightly] for approved hours"],
        ["When paid", "[Payday, e.g. Wednesday after period ends]", "Within [X] days of a correct invoice"],
        ["Payslip", "Emailed or in Connecteam each pay", "Remittance advice emailed"],
        ["Super", "12%, paid each pay run", "See Section 3"],
        ["Pay query", "Message IWS within 7 days of payday", "Message IWS within 7 days of payment"],
    ], [4.2, 6.4, 6.4])
    k.h2("A correct contractor invoice includes")
    k.bullets(["Your name, ABN and contact details", "Invoice number and date",
               "Shift dates, site, hours and rate", "GST shown separately (only if GST-registered)",
               "Bank account details"])
    k.h2("If you can't work a shift")
    k.bullets([
        "Tell IWS **as early as possible**, and at least [X hours] before the shift. Call; don't just message.",
        "Same-day sickness: **call the IWS on-call number**, not the customer only.",
        "Update your **availability** in Connecteam so we only offer shifts you can do.",
    ])

    # 11 Privacy, complaints
    k.h1("Privacy, feedback and complaints")
    k.h2("Privacy")
    k.bullets([
        "Only access information you need for the shift. Never take files, photos or notes away from the site.",
        "Don't share login details. Lock your phone.",
        "Talk about participants only with people who need to know, and never in public.",
    ])
    k.h2("Raising a concern or complaint")
    k.p("We want to hear about anything that worries you: about a site, a customer, a colleague, or IWS itself.")
    k.steps([
        "Talk to the IWS coordinator or on-call person.",
        "If not resolved, email the Director at michael.sibanda@interactivews.com.",
        "You can also contact the **NDIS Quality and Safeguards Commission** on **1800 035 544** or at ndiscommission.gov.au.",
    ])
    k.p("You will not be treated unfairly for raising a concern in good faith. "
        "For workplace rights, you can also contact the Fair Work Ombudsman on 13 13 94.")
    k.callout("Looking after yourself",
              "Support work can be emotionally heavy. After a distressing shift, ask for a debrief. "
              "[If IWS offers an Employee Assistance Program, add the provider and number here.]")

    # 12 Contacts
    k.h1("Key contacts", new_page=True)
    k.table(["Who", "When to contact", "Details"], [
        ["Emergency", "Life-threatening situation, fire, crime in progress", "**000**"],
        ["IWS on-call", "During any shift: incidents, sick calls, running late, any worry", "[On-call phone, hours]"],
        ["IWS office", "Rosters, documents, general questions", "+61 493 716 007\nmichael.sibanda@interactivews.com"],
        ["Pay and invoices", "Pay queries, invoices, super", "michael.sibanda@interactivews.com"],
        ["Director", "Complaints or anything not resolved", "Michael Sibanda\nmichael.sibanda@interactivews.com"],
        ["NDIS Commission", "Complaints and concerns about NDIS supports", "1800 035 544  |  ndiscommission.gov.au"],
        ["Fair Work Ombudsman", "Pay and workplace rights", "13 13 94  |  fairwork.gov.au"],
        ["SafeWork NSW", "Workplace safety advice", "13 10 50  |  safework.nsw.gov.au"],
    ], [3.2, 6.0, 7.8])
    k.callout("Save these two numbers now",
              "**IWS on-call:** [On-call phone]      **Emergency:** 000")

    # 13 Acknowledgement
    k.h1("Worker acknowledgement", new_page=True)
    k.p("Please sign and upload this page to Connecteam, or return it to IWS before your first shift.")
    k.p("I confirm that I have:")
    k.checklist([
        "read and understood this Welcome Pack",
        "read and agree to follow the NDIS Code of Conduct",
        "provided all the documents listed for my engagement type, and will keep them current",
        "downloaded Connecteam and know how to clock in and out",
        "saved the IWS on-call number in my phone",
        "understood that I must report incidents and concerns to the shift lead and IWS immediately",
        "understood that I must not arrange private work with participants, families or customers",
    ])
    k.p("Engagement type:   ☐ Casual employee     ☐ ABN contractor")
    k.signature_block(["Full name", "Signature", "Date", "IWS representative", "IWS signature and date"])

    return k.save("IWS_Welcome_Pack_Contractors_and_Support_Workers")


# =====================================================================
# Pack 2: Customers
# =====================================================================

def customer_pack():
    k = Pack("Welcome Pack: Customers", "Customer Welcome Pack")
    k.cover(
        "Welcome to IWS",
        "Customer Welcome Pack",
        "How we work with you to deliver safe, reliable, quality support.",
        ["About IWS", "Our workers", "Booking shifts", "Your role on site",
         "Incidents", "Rates and invoices", "Feedback", "Key contacts", "Getting started"],
    )

    k.h1("Welcome", new_page=True)
    k.p("Dear [Customer name],")
    k.p("Thank you for choosing Interactive Workforce Solutions. We supply screened, trained support workers "
        "so your participants get consistent, quality support and your team gets reliable cover.")
    k.p("This pack explains how to book workers, what we check before anyone arrives at your site, how we share "
        "safety duties with you, and how invoicing works. Your **service agreement** with IWS sets out the formal terms; "
        "where this pack and the agreement differ, the agreement applies.")
    k.p("**Michael Sibanda**\nDirector, Interactive Workforce Solutions")
    k.callout("Our promise to you", [
        "**Only fully screened workers.** No worker is rostered until every check is complete and recorded.",
        "**Clear communication.** Shift confirmations, changes and any issue, as soon as we know.",
        "**A real person on call** whenever an IWS worker is on shift at your site.",
        "**Accurate invoices** backed by clock-in and clock-out records.",
    ])

    k.h1("About IWS")
    k.p("IWS supplies casual support workers and independent contractors to registered NDIS providers and "
        "care organisations. Our workers work at your sites, under your support plans and site procedures, "
        "with IWS managing their screening, training records, rostering, pay and conduct.")
    k.table(["Service", "What it covers"], [
        ["Regular rostered shifts", "Ongoing shifts with a consistent worker or small team, so participants see familiar faces"],
        ["Short-notice cover", "Sick leave, emergencies and unexpected gaps"],
        ["Planned cover", "Annual leave, school holidays, new participants or extra support periods"],
        ["Shift types", "[Day, evening, active night, sleepover, weekend, public holiday: confirm what IWS offers]"],
        ["Support types", "[Personal care, community access, SIL, high intensity daily activities: confirm what IWS offers]"],
    ], [5.0, 12.0])

    k.h1("Our workers: what we check", new_page=True)
    k.p("Before a worker is offered a shift at your site, IWS sights the original document, records the expiry date "
        "and stores it in our workforce system. You can ask to see any worker's compliance summary at any time.")
    k.table(["Check", "Detail"], [
        ["NDIS Worker Screening Check", "Cleared and verified on the NDIS Worker Screening Database. We can link the worker to your organisation if you require it."],
        ["NDIS Worker Orientation Module", "\"Quality, Safety and You\" certificate"],
        ["NDIS Code of Conduct", "Signed acknowledgement"],
        ["First Aid and CPR", "HLTAID011, with CPR refreshed every 12 months"],
        ["Infection control", "Training certificate, plus vaccinations where your site requires them"],
        ["Manual handling", "Training record"],
        ["Working With Children Check", "Where any participant or household member is under 18"],
        ["Driving", "Licence, registration and comprehensive insurance, where the worker drives participants"],
        ["Right to work", "Verified"],
        ["Contractors only", "Active ABN, public liability and personal accident cover (certificates of currency on file)"],
        ["Specialist training", "Matched to your participants' needs: e.g. medication, mealtime/dysphagia, epilepsy, behaviour support [list what IWS can supply]"],
    ], [5.0, 12.0])
    k.callout("Tell us what your site needs",
              "If your participants need specific skills, gender preferences, languages or cultural matches, tell us "
              "when you book. We will only place workers who meet them.")

    k.h1("Booking shifts")
    k.h2("How to request a worker")
    k.p("Book by phone on **+61 493 716 007** or email **michael.sibanda@interactivews.com**. For each shift, please give us:")
    k.checklist([
        "Site name and address, and parking or entry instructions",
        "Date, start and finish times, and shift type (e.g. active night or sleepover)",
        "Number of workers and any skills, training or gender requirements",
        "Name and phone number of the shift lead the worker reports to",
        "Any current risks the worker must know before arriving",
    ])
    k.h2("What happens next")
    k.steps([
        "**We confirm** the booking and the worker's name by [method], usually within [X hours].",
        "**The worker receives** the site details and shift lead contact through our app, Connecteam.",
        "**The worker clocks in** on arrival and clocks out at finish. These records support your invoice.",
        "**On the first shift** at your site, your shift lead completes the IWS Customer Site Induction Checklist with the worker.",
    ])
    k.table(["Request type", "Notice we need"], [
        ["Regular rostered shifts", "[e.g. 7 days]"],
        ["One-off shifts", "[e.g. 48 hours]"],
        ["Urgent cover", "Call [on-call number]. We'll do our best and tell you quickly if we can't fill it."],
    ], [6.0, 11.0])

    k.h1("Your role as the host site", new_page=True)
    k.p("Under the NSW Work Health and Safety Act 2011, when an IWS worker is at your site, **both IWS and your "
        "organisation owe that worker a duty of care**. The law requires us to consult, cooperate and coordinate "
        "with each other. In practice, we ask you to:")
    k.table(["You provide", "Why it matters"], [
        ["A site induction on each worker's first shift, using the IWS checklist", "Workers cannot work alone until they know the site, the people and the emergency plan"],
        ["Access to support plans, risk assessments and health plans", "Workers must read the plans before giving support"],
        ["A named shift lead the worker reports to", "Clear direction, and someone to escalate to on shift"],
        ["Safe systems of work, equipment and PPE", "Hoists, slide sheets, gloves and other equipment in the participant's plan"],
        ["Information about any new or changed hazard", "So we can brief the worker and adjust the placement"],
        ["Your incident form and process", "So incidents are recorded the same way as your own staff"],
    ], [8.0, 9.0])
    k.callout("Please don't ask workers to step outside their training",
              "For example: giving medication they are not trained or authorised to give, using a restrictive practice "
              "that is not authorised, or performing a two-person transfer alone. Workers are told to stop and call IWS "
              "if this happens.", kind="warn")

    k.h1("Incidents and concerns")
    k.p("If an incident happens while an IWS worker is on shift, the worker will:")
    k.bullets([
        "make the participant safe and call 000 if needed",
        "tell your shift lead immediately",
        "call IWS on-call immediately",
        "complete your incident report the same shift",
    ])
    k.p("As the registered provider, **your organisation notifies reportable incidents to the NDIS Commission** "
        "within the required time (usually 24 hours). IWS will cooperate fully: we provide worker statements, "
        "rostering and clock-in records, and take any action needed with the worker.")
    k.h2("Please tell us")
    k.bullets([
        "about any incident involving an IWS worker, **on the same day**",
        "about any concern with a worker's conduct, skills or reliability, so we can act quickly",
        "if you would prefer a particular worker not to return to your site",
        "if you'd like to request a regular worker, so participants get consistent support",
    ])

    k.h1("Rates, timesheets and invoices")
    k.table(["Item", "How it works"], [
        ["Rates", "Hourly rates by shift type (for example Weekday AM, Weekday PM, night, Saturday, Sunday, public holiday and sleepover), as set out in your rate schedule. [Attach rate schedule]"],
        ["Minimum shift", "[e.g. 2 hours] per worker per shift"],
        ["Hours charged", "Based on the worker's clock-in and clock-out records, as confirmed by your shift lead"],
        ["Overtime and extensions", "Shifts extended at your request are charged at the applicable rate. Please tell IWS of any extension."],
        ["Travel and kilometres", "[If charged: rate per km for participant transport, per your rate schedule]"],
        ["Invoice cycle", "Weekly, per site, with a line for every shift: date, worker, shift type, rate and hours. Hours found after a week is invoiced (for example a shift that ran past its rostered finish) are billed on a supplementary invoice at the standard rate."],
        ["Payment terms", "[e.g. 14 days] from invoice date"],
        ["GST", "IWS is registered for GST. Rates are quoted excluding GST, and 10% GST is added on each tax invoice."],
        ["How to pay", "By EFT to the account shown on each invoice. Please use the invoice number as the payment reference."],
        ["Invoice disputes", "Raise within [7 days] of the invoice. Undisputed amounts remain payable on time."],
    ], [4.2, 12.8])
    k.h2("Cancellations")
    k.table(["When you cancel", "What is charged"], [
        ["More than [X hours] before the shift", "No charge"],
        ["Less than [X hours] before the shift", "[e.g. Minimum shift or full shift charge]"],
        ["Worker turned away on arrival", "[e.g. Full shift charge]"],
    ], [8.0, 9.0])
    k.p("If IWS needs to cancel or replace a worker, we will tell you as early as possible and offer a replacement.")
    k.callout("Engaging our workers directly",
              "Please speak to us before offering an IWS worker direct employment or private work. Your service "
              "agreement explains how this is handled. [Confirm the clause in your service agreement.]")

    k.h1("Privacy and confidentiality")
    k.bullets([
        "IWS workers sign a confidentiality agreement and follow the NDIS Code of Conduct.",
        "We only share worker information you need to manage safe placements.",
        "Please share participant information with our workers on a need-to-know basis, and keep records at your site.",
        "IWS handles personal information under the Privacy Act 1988 and the Australian Privacy Principles. [Link to IWS privacy policy]",
    ])

    k.h1("Feedback and complaints", new_page=True)
    k.p("We want to know what's working and what isn't. Feedback helps us place the right workers with you.")
    k.steps([
        "Call or email your IWS coordinator. Most issues are fixed on the same day.",
        "If not resolved, contact the Director at michael.sibanda@interactivews.com. We will acknowledge within [2 business days] and aim to resolve within [X business days].",
        "Participants and families can also contact the **NDIS Quality and Safeguards Commission** on **1800 035 544**.",
    ])

    k.h1("Key contacts")
    k.table(["Who", "When to contact", "Details"], [
        ["IWS on-call", "Urgent cover, incidents, worker running late or unwell", "[On-call phone, hours]"],
        ["Bookings and rosters", "New bookings, changes, regular worker requests", "+61 493 716 007\nmichael.sibanda@interactivews.com"],
        ["Accounts", "Invoices, rates, payment queries", "michael.sibanda@interactivews.com"],
        ["Director", "Service agreement, escalations and complaints", "Michael Sibanda\nmichael.sibanda@interactivews.com"],
        ["NDIS Commission", "Complaints and concerns about NDIS supports", "1800 035 544  |  ndiscommission.gov.au"],
    ], [3.2, 6.0, 7.8])

    k.h1("Getting started checklist", new_page=True)
    k.p("To set up your account, please complete and return the following to **michael.sibanda@interactivews.com**:")
    k.checklist([
        "Signed service agreement and rate schedule",
        "Accounts contact and invoicing email address",
        "List of sites, with addresses, access and parking details",
        "Shift lead or house manager name and phone number for each site",
        "Skills, training and worker preferences for each site",
        "Copies of key site policies workers must follow (e.g. incident, medication, emergency)",
        "Any vaccination or extra screening requirements",
        "Confirmation of who will complete the IWS Customer Site Induction Checklist at each site",
    ])
    k.h2("Customer details")
    k.signature_block(["Organisation name", "ABN", "NDIS registration number", "Main contact and role",
                       "Phone and email", "Accounts email", "Signature and date"])

    return k.save("IWS_Welcome_Pack_Customers")


def to_pdf(paths):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT), *map(str, paths)],
                   check=True, capture_output=True)


if __name__ == "__main__":
    files = [worker_pack(), customer_pack()]
    to_pdf(files)
    for f in files:
        print(f.name, "and", f.with_suffix(".pdf").name)
