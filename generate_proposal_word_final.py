"""
generate_proposal_word  — FINAL VERSION
========================================
Strategy: Clone the original InCorp_Proposal_2026.docx, then:
  1. Remove children 73–121  (old letter + scope + fees — the dynamic section)
  2. Inject freshly generated dynamic content at position 73
  3. Save

This preserves:
  • Cover page         (children 0–72)   — all images, drawings, fonts INTACT
  • Static pages 2-4   (children 28–72)  — About InCorp, Snapshot, Locations, Accreditations INTACT
  • Static pages 14-21 (children 122–276) — Annexures + T&C INTACT

NO PDF-to-image conversion. NO temp files. Pure XML.
"""

import copy
import io
import os
import zipfile
from datetime import datetime

from docx import Document as DocxDoc
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT as WD_VA
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# ─────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────
RED   = RGBColor(0xC0, 0x00, 0x00)
BLUE  = RGBColor(0x00, 0x20, 0x60)
BLACK = RGBColor(0x00, 0x00, 0x00)
GREY  = RGBColor(0x66, 0x66, 0x66)

# Original docx body-child indices
DYNAMIC_START = 73   # first child of old letter (date paragraph)
DYNAMIC_END   = 122  # first child of Annexure 1 (exclusive)

# ─────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────

def format_currency(amount):
    if not amount or amount in ('', '0'):
        return ''
    try:
        return f'{int(float(amount)):,}'
    except:
        return ''


def fc(v):
    return format_currency(v) or ''


# ─────────────────────────────────────────
# LOW-LEVEL XML BUILDERS
# These return lxml Element objects that can
# be inserted directly into the document body.
# ─────────────────────────────────────────

def _make_pPr(bold=False, italic=False, size_half=18, color_hex=None,
              font_name='Microsoft Sans Serif', align='left',
              space_before=0, space_after=0,
              left_indent=0, first_line=0, style_id=None):
    """Build a <w:pPr> element."""
    pPr = OxmlElement('w:pPr')
    if style_id:
        pStyle = OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'), style_id)
        pPr.append(pStyle)
    if space_before or space_after:
        spacing = OxmlElement('w:spacing')
        if space_before:
            spacing.set(qn('w:before'), str(int(space_before * 20)))
        if space_after:
            spacing.set(qn('w:after'), str(int(space_after * 20)))
        pPr.append(spacing)
    if left_indent or first_line:
        ind = OxmlElement('w:ind')
        if left_indent:
            ind.set(qn('w:left'), str(int(left_indent * 20)))
        if first_line:
            ind.set(qn('w:firstLine') if first_line > 0 else qn('w:hanging'),
                    str(abs(int(first_line * 20))))
        pPr.append(ind)
    jc_map = {'left': 'left', 'center': 'center', 'right': 'right',
              'justify': 'both', 'both': 'both'}
    jc_val = jc_map.get(align, 'left')
    if jc_val != 'left':
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), jc_val)
        pPr.append(jc)
    return pPr


def _make_rPr(bold=False, italic=False, size_half=18, color_hex=None,
              font_name='Microsoft Sans Serif', underline=False):
    """Build a <w:rPr> element."""
    rPr = OxmlElement('w:rPr')
    if font_name:
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement('w:b'))
    if italic:
        rPr.append(OxmlElement('w:i'))
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_half))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size_half))
    rPr.append(szCs)
    if color_hex:
        color = OxmlElement('w:color')
        color.set(qn('w:val'), color_hex.lstrip('#'))
        rPr.append(color)
    return rPr


def _p(text='', bold=False, italic=False, size_pt=9, color_hex=None,
       font='Microsoft Sans Serif', align='left',
       sb=0, sa=0, li=0, fi=0, underline=False):
    """Build a complete <w:p> element."""
    p = OxmlElement('w:p')
    pPr = _make_pPr(align=align, space_before=sb, space_after=sa,
                    left_indent=li, first_line=fi)
    p.append(pPr)
    if text:
        r = OxmlElement('w:r')
        rPr = _make_rPr(bold=bold, italic=italic, size_half=int(size_pt * 2),
                        color_hex=color_hex, font_name=font, underline=underline)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        r.append(t)
        p.append(r)
    return p


def _p_multirun(runs, align='left', sb=0, sa=0, li=0, fi=0):
    """Build a <w:p> with multiple runs. runs = list of (text, bold, italic, size_pt, color_hex, font)"""
    p = OxmlElement('w:p')
    pPr = _make_pPr(align=align, space_before=sb, space_after=sa,
                    left_indent=li, first_line=fi)
    p.append(pPr)
    for (text, bold, italic, size_pt, color_hex, font) in runs:
        if not text:
            continue
        r = OxmlElement('w:r')
        rPr = _make_rPr(bold=bold, italic=italic, size_half=int(size_pt * 2),
                        color_hex=color_hex, font_name=font)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        r.append(t)
        p.append(r)
    return p


def _page_break():
    """<w:p> containing a page break."""
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p


def _spacer(sa=6):
    return _p(sa=sa)


def _h1(text):
    return _p(text, bold=True, size_pt=12, color_hex='C00000',
              font='Roboto', sb=12, sa=10)


def _h2(text):
    return _p(text, bold=True, size_pt=10, color_hex='C00000',
              font='Roboto', sb=10, sa=8, li=4)


def _body(text, align='both'):
    return _p(text, size_pt=9, font='Microsoft Sans Serif', align=align, sa=2)


def _bold_body(text):
    return _p(text, bold=True, size_pt=9, font='Roboto', align='both', sa=2)


def _ital(text, size_pt=8):
    return _p(text, italic=True, size_pt=size_pt, font='Roboto', sb=2, sa=2)


def _bul(text, size_pt=9):
    return _p(text, size_pt=size_pt, font='Microsoft Sans Serif',
              align='both', li=10, fi=-10, sa=2)


def _note_label():
    p = OxmlElement('w:p')
    pPr = _make_pPr()
    p.append(pPr)
    r = OxmlElement('w:r')
    rPr = _make_rPr(bold=True, size_half=18, color_hex='002060',
                    font_name='Microsoft Sans Serif', underline=True)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = 'Note:'
    r.append(t)
    p.append(r)
    return p


def _note_block(items):
    elems = [_note_label(), _spacer(2)]
    for item in items:
        elems.append(_bul(f'\u2022 {item}'))
    elems.append(_spacer(2))
    return elems


def _hourly():
    return [
        _ital('* Any other services not specifically quoted above and not specifically agreed separately shall be chargeable as under:'),
        _ital('For Partner: USD 300 per Hour', 8),
        _ital('For Associates: USD 200 per Hour', 8),
    ]


# ─────────────────────────────────────────
# TABLE BUILDER
# ─────────────────────────────────────────

def _set_cell_width(tc, width_dxa):
    tcPr = tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(width_dxa)))
    tcW.set(qn('w:type'), 'dxa')


def _set_cell_borders(tc, is_header=False):
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old)
    tcB = OxmlElement('w:tcBorders')
    if is_header:
        for side in ['top', 'left', 'right']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'nil')
            tcB.append(b)
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single')
        bot.set(qn('w:sz'), '6')
        bot.set(qn('w:color'), '000000')
        tcB.append(bot)
    else:
        for side in ['top', 'left', 'right', 'bottom']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '4')
            b.set(qn('w:color'), '000000')
            tcB.append(b)
    tcPr.append(tcB)


def _set_cell_valign(tc, valign='top'):
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:vAlign')):
        tcPr.remove(old)
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), valign)
    tcPr.append(va)


def _set_cell_margins(tc, top=80, bottom=80, left=120, right=120):
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(old)
    mar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        mar.append(m)
    tcPr.append(mar)


def _fill_cell(tc, paragraphs, is_header=False, valign='top', width_dxa=None):
    """Fill a table cell with a list of <w:p> elements."""
    # Remove default empty paragraph
    for old_p in tc.findall(qn('w:p')):
        tc.remove(old_p)
    for p_elem in paragraphs:
        tc.append(copy.deepcopy(p_elem) if isinstance(p_elem, etree._Element) else p_elem)
    if not tc.findall(qn('w:p')):
        tc.append(OxmlElement('w:p'))
    _set_cell_borders(tc, is_header)
    _set_cell_valign(tc, valign)
    _set_cell_margins(tc)
    if width_dxa:
        _set_cell_width(tc, width_dxa)


def _make_table(col_widths_inches, rows_data):
    """
    Build a complete <w:tbl> element.

    rows_data: list of rows, each row = list of cells
    Each cell = dict:
        paragraphs: list of <w:p> elements
        is_header: bool
        valign: 'top'/'center'
    """
    tbl = OxmlElement('w:tbl')

    # Table properties
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)
    tblW = OxmlElement('w:tblW')
    total_dxa = int(sum(w * 1440 for w in col_widths_inches))
    tblW.set(qn('w:w'), str(total_dxa))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    # No table-level borders (cell-level handles it)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    tbl.append(tblPr)

    # Column grid
    tblGrid = OxmlElement('w:tblGrid')
    for w in col_widths_inches:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(w * 1440)))
        tblGrid.append(gc)
    tbl.append(tblGrid)

    # Rows
    for row_data in rows_data:
        tr = OxmlElement('w:tr')
        for ci, cell_data in enumerate(row_data):
            tc = OxmlElement('w:tc')
            paragraphs = cell_data.get('paragraphs', [_p()])
            is_hdr = cell_data.get('is_header', False)
            valign = cell_data.get('valign', 'top')
            w_inch = col_widths_inches[ci] if ci < len(col_widths_inches) else 1.0
            _fill_cell(tc, paragraphs, is_header=is_hdr,
                       valign=valign, width_dxa=int(w_inch * 1440))
            tr.append(tc)
        tbl.append(tr)

    return tbl


# ─────────────────────────────────────────
# SECTION LETTER ASSIGNMENT
# ─────────────────────────────────────────

def get_section_letters(data):
    section_keys = [
        ('handover',
         data.get('sectionA') == 'on' and any([
             data.get('includeHandover') == 'on',
             data.get('includeDueDiligence') == 'on'])),
        ('incorporation',
         data.get('sectionB') == 'on' and any([
             data.get(k) == 'on' for k in [
                 'includeIncorporation', 'includeGST', 'includeFCGPR',
                 'includeROC', 'includeIEC', 'includePT', 'includeBEN',
                 'includeMGT', 'includePAN', 'includeTrademark',
                 'includeForeignPAN', 'includeBankAssist',
                 'includeRegOffice', 'includeNomineeDir']])),
        ('accounting',
         data.get('sectionC') == 'on' and any([
             data.get(k) == 'on' for k in [
                 'includeAdvanceTax', 'includeTDS', 'includeIncomeTax',
                 'includeGSTComp', 'includeCompanyLaw', 'includeRBIFiling',
                 'includeMasterFiling', 'includeAcctSetup', 'includeAcctMaint',
                 'includeFinStmt', 'includePayrollSetup', 'includeShopPOSH',
                 'includePayrollProc', 'includeLabourLaw', 'includeAnnualReturns']])),
        ('transfer',
         data.get('sectionD') == 'on' and any([
             data.get('includeBenchmarking') == 'on',
             data.get('includeIntercompany') == 'on'])),
    ]
    letters = {}
    counter = 0
    for key, selected in section_keys:
        if selected:
            letters[key] = chr(65 + counter)
            counter += 1
    return letters


# ─────────────────────────────────────────
# DYNAMIC CONTENT BUILDER
# Returns list of lxml elements to inject
# ─────────────────────────────────────────

def build_dynamic_elements(data):
    """Generate all dynamic page elements as a list of lxml <w:p>/<w:tbl> elements."""
    elems = []

    def add(*items):
        for item in items:
            if isinstance(item, list):
                elems.extend(item)
            else:
                elems.append(item)

    # ── FORMAT DATE ──
    proposal_date = data.get('proposalDate', datetime.now().strftime('%Y-%m-%d'))
    try:
        fd = datetime.strptime(proposal_date, '%Y-%m-%d').strftime('%d. %m. %Y')
    except:
        fd = proposal_date

    # ── LETTER ────────────────────────────────────────────────────
    add(_body(fd))
    add(_spacer(6))
    add(_body(data.get('clientName', 'Client Name')))
    add(_body(data.get('clientDesignation', 'Client Designation')))
    add(_body(data.get('clientCompany', 'Client Company Name')))
    for f in ['clientAddress1', 'clientAddress2', 'clientAddress3']:
        if data.get(f):
            add(_body(data[f]))
    if not any(data.get(f) for f in ['clientAddress1', 'clientAddress2', 'clientAddress3']):
        add(_body(data.get('clientAddress', '')))
    add(_spacer(10))
    add(_body(f"Dear {data.get('clientName', 'XXXX').split()[0]},"))
    add(_spacer(6))
    # RE: FEE PROPOSAL — red bold
    add(_p('RE: FEE PROPOSAL', bold=True, size_pt=9, color_hex='C00000',
           font='Roboto', align='both', sa=4))
    # Letter body
    for line in [
        'We are pleased to be presenting our proposal to you.',
        '',
        ("Our team of experienced professionals work very closely with clients on various "
         "corporate, accounting, compliance and governance matter and identify the unique "
         "requirements of individual organizations. As a strong believer of long-term "
         "partnerships, we are committed to providing tailored solutions that not only meet "
         "our clients\u2019 objectives, but also giving them a peace of mind to focus on "
         "their core businesses."),
        '',
        ("The following pages outline our services tailor made to you and we trust that our "
         "proposal meets your expectations. We are excited to work with you and look forward "
         "to a long and mutually beneficial working relationship with you and the company."),
        '',
        'Yours Sincerely and on behalf of In.Corp,',
        '', '', '',
    ]:
        add(_body(line))
    for txt in ['CA Bansi Shah',
                'Lead \u2013 International clients group',
                'InCorp Advisory Services Pvt Ltd']:
        add(_bold_body(txt))
    add(_page_break())

    # ── SCOPE OF SERVICES ─────────────────────────────────────────
    add(_h1('SCOPE OF SERVICES'))
    add(_spacer(1))
    for line in data.get('scopeOfServices', '').split('\n'):
        if line.strip():
            add(_bul(f'\u2022 {line.strip()}'))
    add(_spacer(6))

    # ── FEES INTRO ────────────────────────────────────────────────
    add(_h1('FEES'))
    fee_type = data.get('feeType', '')
    if fee_type == 'ongoing':
        fee_line = ('Our fee structure includes ongoing charges that may be billed '
                    'monthly, quarterly, or annually.')
    else:
        fee_line = 'Our fee structure includes initial setup fees.'
    add(_body(
        f"This section outlines the estimated fees for InCorp\u2019s services of your "
        f"company. {fee_line} Additionally, fees may be incurred based on the time spent "
        f"on specific tasks or on a per-instance basis. For any additional services not "
        f"encompassed by this proposal that may incur, additional charges, we will receive "
        f"your approval before any work commences. Please note that all fees mentioned are "
        f"in US Dollars, exclusive of the prevailing Goods and Services Tax (GST) / Value "
        f"Added Tax (VAT)."))
    add(_spacer(4))

    letters = get_section_letters(data)

    # ── A. HANDOVER ───────────────────────────────────────────────
    if 'handover' in letters:
        add(_h2(f"{letters['handover']}. One time Handover Service"))
        add(_spacer(1))
        add(_body(
            f"Since the company has been in existence since "
            f"{data.get('companyYear', 'YYYY')}, we shall need to undertake a handover of "
            f"the current financial, secretarial, payroll and other records of the company "
            f"from current service provider."))
        add(_spacer(6))

        # Header row
        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Frequency', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
            {'paragraphs': [_p('Fee (In USD)', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        rows = [h_row]

        if data.get('includeHandover') == 'on':
            fee = fc(data.get('handoverFee', '0')) or '500'
            rows.append([
                {'paragraphs': [
                    _p('Handover from erstwhile service provider of various records under laws as mentioned below. This process does not entail conducting a due diligence.',
                       bold=True, size_pt=9, color_hex='C00000', font='Roboto', align='both'),
                    _bul('\u2022 GST laws/regulations'),
                    _bul('\u2022 Income Tax Act, 1961'),
                    _bul("\u2022 Company\u2019s Act, 2013"),
                    _bul('\u2022 Foreign Exchange Rules & Regulations'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(data.get('handoverFrequency', 'One-time'), size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeDueDiligence') == 'on':
            fee = fc(data.get('dueDiligenceFee', '0')) or '500'
            rows.append([
                {'paragraphs': [
                    _p('Basic due diligence from perspective of*\u2013', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _bul("\u2022 Company\u2019s Act, 2013"),
                    _bul('\u2022 Income Tax Act, 1961'),
                    _bul('\u2022 Goods and Service Tax Act, 2017'),
                    _bul('\u2022 Foreign Exchange Management Act, 1999'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(data.get('ddFrequency', 'One-time'), size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                {'paragraphs': [_p(fee + ' per year', size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if len(rows) > 1:
            add(_make_table([4.4, 1.5, 1.5], rows))
            add(_spacer(6))
            add(_ital('*Any fees for rectification (or) completion of pending past compliances shall attract additional fees and we shall seek your approval prior to commencement of that work.'))
            add(_spacer(6))
        add(*_note_block([
            'All fees quoted above exclude 18% GST',
            'Professional fees exclude any fees towards regularisation of past non compliances.',
            'Advance of 100% of the above selected option.',
        ]))
        add(*_hourly())

    # ── B. INCORPORATION ──────────────────────────────────────────
    if 'incorporation' in letters:
        add(_spacer(6))
        add(_h2(f"{letters['incorporation']}. Incorporation / Secretarial Service and Mandatory Registrations post Incorporation"))
        add(_spacer(2))

        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('One-time Fee', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        rows = [h_row]

        if data.get('includeIncorporation') == 'on':
            fee = fc(data.get('incorporationFee', '0')) or '1500'
            rows.append([
                {'paragraphs': [
                    _p('Incorporation', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _bul('\u2022 PAN of the company included'),
                    _bul('\u2022 TAN of the company included'),
                    _bul("\u2022 Employees\u2019 Provident Fund and Miscellaneous Provision Act, Employees\u2019 State Insurance Corporation Act included"),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeGST') == 'on':
            fee = fc(data.get('gstRegFee', '0')) or '350'
            rows.append([
                {'paragraphs': [
                    _p('Goods & Service Tax (GST)', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _body('Registration of single location with GST authorities.'),
                    _ital('Registration of every additional location with the GST authorities shall cost USD 100'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeFCGPR') == 'on':
            fee = (fc(data.get('fcgprFee', '0')) or '1250') + ' per applicant'
            rows.append([
                {'paragraphs': [
                    _p('FCGPR Filing with Reserve Bank of India', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _body('Filing of Forms and declaration with RBI as required under FEMA'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeROC') == 'on':
            fee = fc(data.get('rocComplianceFee', '0')) or '500'
            rows.append([
                {'paragraphs': [
                    _p('Statutory Compliances with Registrar of Companies under Companies Act:', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _bul('\u2022 Drafting of first board meeting documents'),
                    _bul('\u2022 Guidance on capital infusion in bank account'),
                    _bul('\u2022 File form with Ministry for commencement of business (COC)'),
                    _bul('\u2022 Preparation of statutory shareholders register'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if len(rows) > 1:
            add(_make_table([5.7, 1.5], rows))
            add(_spacer(8))
        add(*_note_block([
            'All fees quoted above exclude 18% GST.',
            'Professional fees exclude all out-of-pocket expenses like filing fees, courier expenses, apostilling & notary cost to any authorities/departments, statutory fees payable to Registrar of companies (ROC) towards incorporation etc. other than those mentioned above.',
            'Advance of 100% of the above selected option.',
            'On finalization of shareholding structure, we shall be able to guide on compliances needed for issuance of share certificates and shall share a separate fee quote for the same.',
        ]))
        add(*_hourly())
        add(_spacer(4))

        # Optional registrations
        add(_h2('Optional registrations required post incorporation (One-time)'))
        add(_spacer(3))
        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Fees (In USD)', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        opt_rows = [h_row]
        for cb, fk, dflt, lbl in [
            ('includeIEC', 'iecFee', '200', 'Import Export Code (IEC Code)'),
            ('includePT', 'ptFee', '200', "Profession Tax (PT) \u2022 Payments and return filing for company, its employees until the company\u2019s certificate of commencement is obtained"),
            ('includeBEN', 'benFee', '250', 'Submission of for Significant Beneficial Ownership via form BEN-2'),
            ('includeMGT', 'mgtFee', '250', 'Filing of requisite forms with ROC (Form MGT 4, MGT 5, MGT 6)'),
            ('includePAN', 'panCardFee', '300', 'Physical PAN Card of the company'),
            ('includeTrademark', 'trademarkFee', '350', 'Trademark Registration (exclusive of disbursement fees)'),
            ('includeForeignPAN', 'foreignPanFee', '200 per director', 'PAN for foreign director'),
            ('includeBankAssist', 'bankAssistFee', '250', 'Assistance in opening of bank account'),
        ]:
            if data.get(cb) == 'on':
                fee = fc(data.get(fk, '0')) or dflt.split()[0]
                if fk == 'foreignPanFee' and fc(data.get(fk, '0')):
                    fee += ' per director'
                opt_rows.append([
                    {'paragraphs': [_p(lbl, bold=True, size_pt=9, color_hex='C00000', font='Roboto')], 'valign': 'top'},
                    {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                ])
        if len(opt_rows) > 1:
            add(_make_table([5.7, 1.5], opt_rows))
            add(_spacer(6))
        add(_ital("*For every new director\u2019s professional tax no., there shall be additional cost of $100 per director"))
        add(_ital('*Digital signature certificate (DSC) token can be obtained at a cost of USD 200 per applicant.'))
        add(_spacer(10))
        add(*_note_block([
            'All fees quoted above exclude 18% GST.',
            'Professional fees exclude all out-of-pocket expenses like filing fees, courier expenses, apostilling & notary cost to any authorities/departments, statutory fees payable to Registrar of companies (ROC) towards incorporation etc. other than those mentioned above.',
            'Advance of 100% of the above selected option.',
        ]))
        add(*_hourly())
        add(_spacer(4))

        # Nominee Director
        add(_h2('Nominee Director and Registered Office Address Service'))
        add(_spacer(4))
        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Monthly Fee (in USD)', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        nom_rows = [h_row]

        if data.get('includeRegOffice') == 'on':
            fee = fc(data.get('registeredOfficeFee', '0')) or '300'
            nom_rows.append([
                {'paragraphs': [
                    _p('Registered Office Service', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _body('A refundable Security deposit @USD 2500 applies**. Refundable upon cessation of Registered office service.'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeNomineeDir') == 'on':
            fee = fc(data.get('nomineeDirectorFee', '0')) or '350'
            nom_rows.append([
                {'paragraphs': [
                    _p('Nominee Director Service', bold=True, size_pt=9, color_hex='C00000', font='Roboto'),
                    _body('A refundable Security deposit per nominee @USD 5000 applies*. Refundable upon cessation of Nominee Director Service'),
                    _body("Director\u2019s fee for attending a physical or recorded or live board meeting @USD300 per director per board meeting"),
                    _body("Every nominee director needs to be protected under a director\u2019s indemnity policy. Premium of indemnity bond to be charged on actual basis. InCorp shall enter into a separate nominee directors\u2019 agreement at the time of engagement."),
                    _body('To ensure the removal of a nominee director from registrations ***with various authorities where required, InCorp must be notified at least three months in advance.'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if len(nom_rows) > 1:
            add(_make_table([5.7, 1.5], nom_rows))
            add(_spacer(4))
            add(_ital("*Failure to engage InCorp\u2019s services for regular compliances of the company post the setup such as tax, secretarial, FEMA etc. shall result in forfeiture of the security deposit received against nominee director and registered office services.", 8))
            add(_ital('**Any fees for rectification (or) completion of pending past compliances shall attract additional fees and we shall seek your approval prior to commencement of that work.', 8))
            add(_ital('*** The Nominee Director shall not sign any return, forms or documents relating to any statutory filing nor will be appointed as the authorized signatory to any of the bank accounts of the entity or under GST, Income Tax any other government portal.', 8))
            add(_spacer(6))
            add(*_note_block([
                'All fees quoted above exclude 18% GST.',
                'The Nominee Director will not be involved in day-to-day affairs / management of the Company.',
                'The service of Registered office & Nominee director is offered on discretionary basis only for temporary basis of 6 months.',
                "Failure to engage InCorp\u2019s services for regular compliances of the company post the setup such as tax, secretarial, FEMA etc. shall result in forfeiture of the security deposit received against registered office and nominee director services.",
                'Professional fees exclude all out-of-pocket expenses like filing fees, courier expenses, apostilling & notary cost to any authorities/departments, statutory fees payable to Registrar of companies (ROC) towards incorporation etc. other than those mentioned above.',
                'Advance of 100% of the above selected option.',
            ]))
            add(*_hourly())

    # ── C. ACCOUNTING / TAX / PAYROLL ─────────────────────────────
    if 'accounting' in letters:
        add(_spacer(6))
        add(_h2(f"{letters['accounting']}. Accounting / Tax / Payroll / Annual Compliance Services"))
        add(_spacer(5))
        add(_p(
            "If the number of transactions are not known while preparing the proposal then "
            "'Depending on the estimated volume of transactions, business nature, products and "
            "services rendered by the company and actual requirements, the below fees are being "
            "quoted based on certain assumptions. Fees will be adjusted once InCorp scopes out "
            "the details with the client",
            size_pt=8, font='Microsoft Sans Serif'))
        add(_spacer(6))

        ta = 0  # total annual
        to = 0  # total one-time

        def adt(freq, raw_fee):
            nonlocal ta, to
            try:
                fee = float(str(raw_fee).split()[0].replace(',', ''))
                fl = freq.lower()
                if 'one' in fl:
                    to += fee
                elif 'month' in fl:
                    ta += fee * 12
                elif 'quarter' in fl:
                    ta += fee * 4
                elif 'annual' in fl or 'year' in fl:
                    ta += fee
            except:
                pass

        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Frequency', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
            {'paragraphs': [_p('Notes', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Fees (in USD)', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        acc_rows = [h_row]

        section_labels = {
            'includeAdvanceTax': 'Direct tax compliances',
            'includeGSTComp': 'Indirect tax compliances',
            'includeCompanyLaw': 'Company Law',
            'includeRBIFiling': 'Foreign Exchange laws',
            'includeAcctSetup': 'Accounting',
            'includePayrollSetup': 'Payroll',
        }

        entries = [
            ('includeAdvanceTax', 'advanceTaxFee', '200', 'advanceTaxFrequency', 'Quarterly',
             '1) Advance tax Compliances\n\u2022 Quarterly calculations and payment'),
            ('includeTDS', 'tdsFee', '200', 'tdsFrequency', 'Monthly/Quarterly',
             '2) TDS compliances:\n\u2022 Calculation and Payment of TDS\n\u2022 Filing of TDS Returns\n(The above excludes cost of revisions of TDS returns.)'),
            ('includeIncomeTax', 'incomeTaxReturnFee', '500', 'incomeTaxFrequency', 'Annual',
             '3) Annual Income tax return\nComputation and filing of Annual Income tax Return\n4) Statement of Financial Transactions (SFT) \u2013 Basic Reporting'),
            ('includeGSTComp', 'gstComplianceFee', '250', 'gstFrequency', 'Monthly and Annual',
             '1) GST Compliances:\n\u2022 Calculations and payment of GST\n\u2022 Filing of monthly GST Returns\n(The above excludes cost of revisions of GST returns.)'),
            ('includeCompanyLaw', 'companyLawFee', '200', 'companyLawFrequency', 'Monthly',
             'Company Law Compliances (Scope as per Annexure 1)\nAssistance on conduction of virtual board meeting \u2013 USD 150 per board meeting'),
            ('includeRBIFiling', 'rbiFilingFee', '450', 'rbiFilingFrequency', 'Annual',
             'Annual Filings with Reserve bank of India'),
            ('includeMasterFiling', 'masterFilingFee', '350', 'masterFilingFrequency', 'Annual',
             'Annual Master Filing Form 3CEAA Part A (Basic Reporting)'),
            ('includeAcctSetup', 'accountingSetupFee', '300', 'acctSetupFrequency', 'One time',
             'Setup of accounting software\n\u2022 Liaison with the software expert for the setup\n\u2022 Ensure due configuration of the software with applicable laws\n\u2022 Short tutorial on guidance with respect to use of accounting software'),
            ('includeAcctMaint', 'accountingMaintenanceFee', '200', 'acctMaintFrequency', 'Monthly',
             'Accounting and maintenance of books of accounts:\n\u2022 Data entry in accounting software\n\u2022 Weekly processing of Bank Reconciliation\n\u2022 Weekly processing of Purchase invoices'),
            ('includeFinStmt', 'financialStatementsFee', '500', 'finStmtFrequency', 'Annual',
             '\u2022 Preparation of the financial Statements as per the Indian accounting Standards\n\u2022 Liaising with auditors for audit, compliance and related matters'),
            ('includePayrollSetup', 'payrollSetupFee', '500', 'payrollSetupFrequency', 'One time',
             'Payroll Setup (Scope as per Annexure 2)'),
            ('includeShopPOSH', 'shopPOSHFee', '0', 'shopPOSHFrequency', 'One time',
             '1. Obtaining Shop and establishment registration\n2. Drafting of POSH (Prevention of Sexual Harassment at Workplace) policy'),
            ('includePayrollProc', 'payrollProcessingFee', '125', 'payrollProcFrequency', 'Monthly',
             'Payroll Processing (Scope as per Annexure 3)'),
            ('includeLabourLaw', 'labourLawFee', '200', 'labourLawFrequency', 'Monthly',
             'Labour Law Compliances \u2022 Payments and return filing under:\n\u2022 Provident Fund\n\u2022 Employees State Insurance Corporation\n\u2022 Profession Tax\n\u2022 Labor Welfare Fund\n(for employees upto 20 \u2013 fixed fee)'),
            ('includeAnnualReturns', 'annualReturnsFee', '200', 'annualReturnsFrequency', 'Annual',
             'Annual Return under the following labor law compliances:\n\u2022 Sexual Harassment of Women at Workplace Act, 2013\n\u2022 Shop and Establishment Act\n\u2022 Maternity Act\n\u2022 Gratuity Act'),
        ]

        for cb, fk, dflt, fqk, dffq, notes in entries:
            if data.get(cb) == 'on':
                raw_fee = fc(data.get(fk, '0')) or dflt
                freq = data.get(fqk, dffq)
                adt(freq, raw_fee)
                fl = freq.lower()
                sfx = (' per month' if 'month' in fl
                       else (' per annum' if ('annual' in fl or 'year' in fl)
                       else ' one time'))
                fee_display = raw_fee + sfx
                svc_label = section_labels.get(cb, '')
                svc_color = 'C00000' if svc_label else None

                # Notes: split on \n into multiple paragraphs
                note_paras = []
                for line in notes.split('\n'):
                    if line.strip():
                        note_paras.append(_body(line))
                    else:
                        note_paras.append(_spacer(2))

                acc_rows.append([
                    {'paragraphs': [_p(svc_label, bold=True, size_pt=9, color_hex=svc_color, font='Roboto', align='center')] if svc_label else [_p()], 'valign': 'center'},
                    {'paragraphs': [_p(freq, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                    {'paragraphs': note_paras, 'valign': 'top'},
                    {'paragraphs': [_p(fee_display, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                ])

        # Total rows
        acc_rows.append([
            {'paragraphs': [_p('Total costs (excluding one time costs)', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center', 'is_header': False},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p(f'{int(ta):,} per annum', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
        ])
        acc_rows.append([
            {'paragraphs': [_p('One-time costs', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p(f'{int(to):,} one time', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
        ])

        if len(acc_rows) > 1:
            add(_make_table([1.3, 1.3, 3.2, 1.4], acc_rows))
            add(_spacer(6))

        add(_ital('*The above quotation fee is for approx.20 transactions per month'))
        add(_spacer(4))
        add(_ital("^InCorp\u2019s empanelled audit partners can offer the services of statutory audit, tax audit and GST audit services. The quotes for the same can be provided separately."))
        add(_spacer(4))
        add(_ital("^Audit partner firms (Jayesh Sanghrajka & Associates, Manish Modi & Associates) shall be able to assist on that front. The estimated statutory fee quote for the first FY shall be between USD 2500 TO USD 3500."))
        add(_spacer(4))
        add(*_note_block([
            'All fees quoted above exclude 18% GST.',
            'Professional fees exclude all out-of-pocket expenses like filing fees, courier expenses, government/statutory fees etc.',
            'Advance of 100% of the above selected option',
        ]))
        add(_ital('*** Any other services not specifically quoted above and not specifically agreed separately shall be chargeable as under:'))
        add(_ital('For Partner: USD 300 per Hour', 8))
        add(_ital('For Associates: USD 200 per Hour', 8))
        add(_spacer(10))

    # ── D. TRANSFER PRICING ───────────────────────────────────────
    if 'transfer' in letters:
        add(_h2(f"{letters['transfer']}. Transfer Pricing compliances"))
        add(_spacer(4))

        h_row = [
            {'paragraphs': [_p('Services', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Frequency', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
            {'paragraphs': [_p('Notes', bold=True, size_pt=10, font='Microsoft Sans Serif')], 'is_header': True},
            {'paragraphs': [_p('Fee (In USD)', bold=True, size_pt=10, font='Microsoft Sans Serif', align='center')], 'is_header': True},
        ]
        tp_rows = [h_row]
        tp_total = 0

        if data.get('includeBenchmarking') == 'on':
            fee = (fc(data.get('benchmarkingFee', '0')) or '1500') + ' per business activity'
            try:
                tp_total += int(fee.split()[0].replace(',', ''))
            except:
                pass
            tp_rows.append([
                {'paragraphs': [_p('Benchmarking', bold=True, size_pt=9, color_hex='C00000', font='Roboto')], 'valign': 'center'},
                {'paragraphs': [_p('One-time', size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                {'paragraphs': [
                    _body('1. Assistance in conducting Functional, Asset and Risk Analysis of the proposed transaction to be entered between related parties.'),
                    _body("2. Assisting in arriving at the arm\u2019s length price or margin range that may be applicable to the proposed transaction."),
                    _body('3. Preparation of final benchmarking report*.'),
                ], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        if data.get('includeIntercompany') == 'on':
            fee = fc(data.get('intercompanyAgreementFee', '0')) or '1500'
            try:
                tp_total += int(str(fee).replace(',', ''))
            except:
                pass
            tp_rows.append([
                {'paragraphs': [_p('Inter-company agreement', bold=True, size_pt=9, color_hex='C00000', font='Roboto')], 'valign': 'center'},
                {'paragraphs': [_p('One-time', size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
                {'paragraphs': [_body(
                    'Drafting and finalizing of Inter-company service agreement covering detailed '
                    'description of service to be provided, components to be included while '
                    'calculating cost of services, Invoicing period, Receivable cycle, withholding, '
                    'ownership rights, effective date of agreement, indemnity etc. in compliance '
                    'with the Transfer Pricing regulations defined under Income tax laws and other '
                    'applicable Indian laws')], 'valign': 'top'},
                {'paragraphs': [_p(fee, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            ])

        tp_rows.append([
            {'paragraphs': [_p('Total one-time costs', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p()], 'valign': 'center'},
            {'paragraphs': [_p(f'{tp_total:,}', bold=True, size_pt=9, font='Microsoft Sans Serif', align='center')], 'valign': 'center'},
        ])

        if len(tp_rows) > 1:
            add(_make_table([1.4, 1.1, 3.3, 1.4], tp_rows))
            add(_spacer(6))
            add(_ital("*Please note that the above benchmarking report will not be transfer pricing documentation as required to be maintained under transfer pricing regulations."))

        add(_spacer(6))
        add(*_note_block([
            'All fees quoted above exclude 18% GST.',
            'Professional fees exclude all out-of-pocket expenses.',
            'Advance of 100% of the above selected option.',
        ]))
        add(_ital('* Any other services not specifically quoted above shall be chargeable as under:'))
        add(_ital('For Partner: USD 300 per Hour', 8))
        add(_ital('For Associates: USD 200 per Hour', 8))
        add(_spacer(4))
        add(_ital("^ InCorp\u2019s empanelled audit partners can assist with the transfer pricing reporting & audit (applicable for companies having intercompany transactions). The quotes for the same can be provided separately."))

    return elems


# ─────────────────────────────────────────
# MAIN: clone original docx + replace dynamic section
# ─────────────────────────────────────────

def generate_word_doc(data, original_docx_path, output_path=None):
    """
    Clone original docx, replace dynamic section (children 73-121)
    with freshly generated content. Returns output_path.
    """
    # Step 1: Clone the entire original docx in memory
    # We do this by reading bytes and creating a new Document from them
    with open(original_docx_path, 'rb') as f:
        original_bytes = f.read()
    
    doc = DocxDoc(io.BytesIO(original_bytes))
    body = doc.element.body
    children = list(body)

    # Step 2: Remove old dynamic children (73 to 121 inclusive)
    # We must find them fresh since indices shift as we remove
    # Collect the actual element references first
    to_remove = children[DYNAMIC_START:DYNAMIC_END]
    for el in to_remove:
        body.remove(el)

    # Step 3: Build new dynamic elements
    new_elems = build_dynamic_elements(data)

    # Step 4: Find insertion point — after child 72 (now at index 72 since 73+ removed)
    # Get fresh children list
    current_children = list(body)
    # Insert point is after index 72 (0-based), i.e. before index 73
    # But after removal, what was child 72 is still child 72
    insert_after_elem = current_children[72]  # last cover/static-front child

    # Insert all new elements after insert_after_elem
    ref = insert_after_elem
    for elem in new_elems:
        ref.addnext(elem)
        ref = elem  # move pointer forward

    # Step 5: Update company name in cover page
    # Child 22 has the old company name text — update it
    try:
        company_name = data.get('clientCompany', '')
        if company_name:
            cover_children = list(body)  # fresh after insertions
            # Find the company name paragraph (was child 22 in original)
            # It contains 'BW WATER INDIA PVT LTD' or similar
            for child in cover_children[:30]:
                if child.tag == qn('w:p'):
                    text = ''.join(t.text or '' for t in child.iter(qn('w:t')))
                    # Match the old company name — any ALL CAPS text in cover area
                    # that is not a fixed label
                    fixed = {'LEADING ASIA PACIFIC CORPORATE SOLUTIONS PROVIDER',
                             'INCORP GROUP PROPOSAL', 'Table of Contents', ''}
                    if text.strip() and text.strip() not in fixed:
                        # This is likely the client company paragraph — update it
                        for t_elem in child.iter(qn('w:t')):
                            if t_elem.text and t_elem.text.strip() and t_elem.text.strip() not in fixed:
                                t_elem.text = company_name.upper()
                                break
                        break
    except Exception as e:
        print(f'Warning: could not update company name in cover: {e}')

    # Step 6: Save
    if output_path is None:
        client = data.get('clientCompany', 'Client').replace(' ', '_')
        output_path = f'InCorp_Proposal_{client}_{datetime.now().strftime("%Y%m%d")}.docx'

    doc.save(output_path)
    print(f'✅ Word document saved: {output_path}')
    return output_path


# ─────────────────────────────────────────
# FLASK ROUTE — drop-in replacement
# ─────────────────────────────────────────

def generate_proposal_word_route(request, send_file, jsonify, BASE_DIR):
    """
    Flask route handler.
    Usage in app.py:
        from generate_proposal_word_final import generate_proposal_word_route

        @app.route('/generate_proposal_word', methods=['POST'])
        def generate_proposal_word():
            return generate_proposal_word_route(request, send_file, jsonify, BASE_DIR)

    Requires: InCorp_Proposal_2026.docx in BASE_DIR
    """
    import traceback
    try:
        data = request.json or {}
        client_name = data.get('clientCompany', 'Client').replace(' ', '_')
        output_filename = (f'InCorp_Proposal_{client_name}_'
                           f'{datetime.now().strftime("%Y%m%d")}.docx')
        output_path = os.path.join(BASE_DIR, output_filename)

        original_docx = os.path.join(BASE_DIR, 'InCorp_Proposal_2026.docx')
        if not os.path.exists(original_docx):
            return jsonify({'error': f'Template not found: {original_docx}'}), 500

        generate_word_doc(
            data=data,
            original_docx_path=original_docx,
            output_path=output_path,
        )

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=output_filename,
        )

    except Exception as e:
        print(f'Error generating Word: {e}')
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500