import streamlit as st
import pandas as pd
import altair as alt
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- FONT CONSTANTS ---
BODY_FONT = "Verdana"
BODY_SIZE = 9
HEADER_FONT = "Verdana"
HEADER_SIZE = 10

# Matches a line that starts with manual numbering, e.g. "1. ", "2) ", "(3) ",
# nested numbering like "1.2. " / "2.3) ", and also "1.Text" with no space
# after the marker (a common typo in pasted Excel content). The (?!\d) guard
# stops a plain decimal like "9.8 is the CVSS score" from being misread as
# marker "9." followed by "8 is...".
NUMBERED_LINE_RE = re.compile(r'^\s*\(?\d+(?:\.\d+)*[\.\)](?!\d)\s*')

# --- RISK BADGE COLOR SCHEME (hex, no '#') ---
RISK_COLORS = {
    "CRITICAL": {"bg": "C00000", "text": "FFFFFF"},
    "HIGH":     {"bg": "EE0000", "text": "FFFFFF"},
    "MEDIUM":   {"bg": "FFC000", "text": "000000"},
    "LOW":      {"bg": "92D050", "text": "000000"},
    "INFO":     {"bg": "00B0F0", "text": "000000"},
}
FALLBACK_RISK_COLOR = {"bg": "D9D9D9", "text": "000000"}  # used for unrecognized ratings

# Sort order for "arrange findings by severity, descending": lower number = higher severity = sorts first.
# Anything not in this map (typos, "Best Practice", etc.) sorts after Info, last.
SEVERITY_ORDER = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3, "INFO": 4}


def normalize_risk_key(risk_text):
    """Single source of truth for matching a raw Excel risk value (e.g.
    'Informational', ' high ') to one of our standard severity buckets."""
    key = str(risk_text).strip().upper()
    if key.startswith("INFORM"):  # e.g. "Informational" -> INFO
        key = "INFO"
    return key


# Relative widths for the 5 summary-table columns (must sum to the usable
# page width — 6.5" assumes default 1" margins on a Letter page).
SUMMARY_TABLE_COL_WIDTHS = [Inches(0.9), Inches(0.9), Inches(3.0), Inches(1.0), Inches(0.7)]


# --- HELPER FUNCTIONS FOR FONT / TABLE STYLING ---

def apply_font(font, name=BODY_FONT, size=BODY_SIZE, bold=None, color=None):
    """
    Force a font onto a python-docx Font object (works for both run.font
    and style.font).

    python-docx's font.name setter only writes the w:ascii / w:hAnsi
    attributes of <w:rFonts>. Word's default template links Title/Heading
    styles to *theme* fonts (asciiTheme="majorHAnsi" etc.), and those theme
    attributes can cause Word to keep showing the template font instead of
    Verdana. This helper sets ascii/hAnsi/eastAsia/cs explicitly AND strips
    any theme attributes, so the font change is guaranteed to stick.

    `color`, if given, is a hex string without '#' (e.g. "000000").
    """
    font.name = name
    font.size = Pt(size)
    if bold is not None:
        font.bold = bold
    if color is not None:
        font.color.rgb = RGBColor.from_string(color)

    rPr = font._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)

    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rFonts.set(qn(attr), name)
    for theme_attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if rFonts.get(qn(theme_attr)) is not None:
            del rFonts.attrib[qn(theme_attr)]



def set_cell_background(cell, fill_color):
    """Injects XML to set the background color of a Word table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)


def set_cell_border(cell, **kwargs):
    """
    Set one or more borders on a table cell.
    Usage: set_cell_border(cell, top={"sz": 24, "val": "single", "color": "000000"},
                                  bottom={"sz": 6, "val": "single", "color": "000000"})
    `sz` is in eighths of a point (e.g. 24 = 3pt, 6 = 0.75pt).
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if not edge_data:
            continue
        tag = qn(f'w:{edge}')
        element = tcBorders.find(tag)
        if element is None:
            element = OxmlElement(f'w:{edge}')
            tcBorders.append(element)
        for key in ('sz', 'val', 'color', 'space'):
            if key in edge_data:
                element.set(qn(f'w:{key}'), str(edge_data[key]))


def get_risk_style(risk_text):
    """Looks up the badge background/text color for a risk rating, with a
    neutral gray fallback for any value not in RISK_COLORS (e.g. typos or
    a rating like 'Best Practice' that isn't part of the standard scheme)."""
    return RISK_COLORS.get(normalize_risk_key(risk_text), FALLBACK_RISK_COLOR)


def sort_by_severity(df, risk_col):
    """Returns a copy of df ordered Critical > High > Medium > Low > Info,
    with anything unrecognized sorted last. Index is reset so finding
    numbers (3.1, 3.2...) follow the new order, not the original row order."""
    if risk_col not in df.columns:
        return df
    rank = df[risk_col].apply(lambda r: SEVERITY_ORDER.get(normalize_risk_key(r), len(SEVERITY_ORDER)))
    return (
        df.assign(_severity_rank=rank)
          .sort_values('_severity_rank', kind='stable')
          .drop(columns='_severity_rank')
          .reset_index(drop=True)
    )


def set_table_column_widths(table, widths):
    """python-docx needs width set on the table, AND on every cell in each
    column, AND autofit disabled — otherwise Word ignores it and auto-sizes
    columns based on content instead."""
    table.autofit = False
    table.allow_autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
    for idx, width in enumerate(widths):
        table.columns[idx].width = width


def add_risk_summary_table(doc, risk, overall_score, cvss_vector, owasp_top10, cwe_id):
    """Builds the borderless 'badge' summary table: a bold header row with a
    thick top rule, and one centered data row with the Risk Rating cell
    color-coded by severity and the remaining cells on a light gray band."""
    headers = ["Risk Rating", "Overall Score", "CVSS Vector", "OWASP Top 10", "CWE ID"]
    values = [str(risk).upper(), str(overall_score), str(cvss_vector), str(owasp_top10), str(cwe_id)]
    style = get_risk_style(risk)

    table = doc.add_table(rows=2, cols=5)
    set_table_column_widths(table, SUMMARY_TABLE_COL_WIDTHS)

    # Header row
    for col, text in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = text
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        for run in para.runs:
            apply_font(run.font, BODY_FONT, BODY_SIZE, bold=True)
        set_cell_border(
            cell,
            top={"sz": 24, "val": "single", "color": "000000"},     # thick top rule
            bottom={"sz": 6, "val": "single", "color": "000000"},   # thin separator under header
        )

    # Data row
    for col, text in enumerate(values):
        cell = table.cell(1, col)
        cell.text = text
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        is_risk_cell = (col == 0)
        for run in para.runs:
            apply_font(run.font, BODY_FONT, BODY_SIZE, bold=is_risk_cell)
            if is_risk_cell:
                run.font.color.rgb = RGBColor.from_string(style["text"])
        set_cell_background(cell, style["bg"] if is_risk_cell else "F2F2F2")
        set_cell_border(cell, bottom={"sz": 6, "val": "single", "color": "000000"})  # thin bottom rule

    return table


def add_styled_heading(doc, text, level):
    """doc.add_heading wrapper that forces Verdana + black color on every run
    it creates. Heading 3 also gets 6pt space-after (Word's default Heading
    styles otherwise use a theme accent color and tighter spacing)."""
    heading = doc.add_heading(text, level)
    for run in heading.runs:
        apply_font(run.font, HEADER_FONT, HEADER_SIZE, color="000000")
    if level == 3:
        heading.paragraph_format.space_after = Pt(6)
    return heading


def add_text_block(doc, value, justify=False):
    """
    Writes Excel cell content as one or more paragraphs (splitting on line
    breaks, since a single docx paragraph doesn't render embedded '\\n'
    characters as line breaks).

    Any line that looks like a manually-numbered item ("1. ...", "2) ...")
    gets list-style paragraph formatting: hanging indent (Left 0.63cm,
    Hanging 0.63cm), Spacing Before 0pt / After 10pt, Line spacing
    Multiple 1.15 — matching Word's Paragraph dialog settings. Every line
    is justified when `justify=True`, list item or not.

    The gap between the number/marker and the item text is normalized to a
    single space, so inconsistent source spacing (e.g. "1.  Do X" with two
    spaces vs "2. Do Y" with one) doesn't show up as a visual misalignment.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        text = "N/A"
    else:
        text = str(value)

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        lines = [text]

    for line in lines:
        match = NUMBERED_LINE_RE.match(line)
        if match:
            marker = match.group(0).strip()
            rest = line[match.end():].strip()
            line = f"{marker} {rest}" if rest else marker

        p = doc.add_paragraph(line)
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if match:
            pf = p.paragraph_format
            pf.left_indent = Cm(0.63)
            pf.first_line_indent = -Cm(0.63)  # hanging indent
            pf.space_before = Pt(0)
            pf.space_after = Pt(10)
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            pf.line_spacing = 1.15
        for run in p.runs:
            apply_font(run.font, BODY_FONT, BODY_SIZE)


def add_styled_paragraph(doc, text="", style=None, alignment=None):
    """doc.add_paragraph wrapper that forces Verdana body font on every run it creates."""
    p = doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
    if alignment is not None:
        p.alignment = alignment
    for run in p.runs:
        apply_font(run.font, BODY_FONT, BODY_SIZE)
    return p


# (Heading shown in the report, source Excel column, justify the body text?)
# NOTE: the Excel column names here must match your "Risk Register" tab headers
# exactly — mismatches (e.g. "Observation" vs "Observations") silently fall
# back to "N/A" instead of erroring, which is why Observation/Implication
# were showing up blank before.
FINDING_SECTIONS = [
    ("Observation",          "Observations",              True),
    ("Implication",          "Implications",              True),
    ("Recommendations",      "Recommendations",           True),
    ("Management Comments",  "Management Comments",       False),
    ("Follow-up Comments",   "Post Review Observations",  True),
    ("Status",               "Status",                    False),
]


# --- FUNCTION TO GENERATE WORD DOC ---
def create_word_report(df):
    doc = Document()

    # --- 1. SET GLOBAL STYLES TO VERDANA (belt-and-suspenders with the
    #         per-run apply_font() calls below, since this also covers any
    #         text Word itself generates, e.g. TOC fields or future edits) ---
    apply_font(doc.styles['Normal'].font, BODY_FONT, BODY_SIZE)
    apply_font(doc.styles['Title'].font, HEADER_FONT, HEADER_SIZE, color="000000")
    for i in range(1, 4):
        apply_font(doc.styles[f'Heading {i}'].font, HEADER_FONT, HEADER_SIZE, color="000000")
    doc.styles['Heading 3'].paragraph_format.space_after = Pt(6)
    apply_font(doc.styles['Table Grid'].font, BODY_FONT, BODY_SIZE)
    apply_font(doc.styles['List Bullet'].font, BODY_FONT, BODY_SIZE)

    add_styled_heading(doc, 'Penetration Testing Report', 0)

    # Executive Summary Section
    add_styled_heading(doc, '1. Executive Summary', 1)
    add_styled_paragraph(doc, f"Total Vulnerabilities Identified: {len(df)}")

    risk_col = "CVSS Risk Rating" if "CVSS Risk Rating" in df.columns else "Risk Rating"
    if risk_col in df.columns:
        risk_counts = df[risk_col].value_counts()
        add_styled_paragraph(doc, "Vulnerabilities by Risk Rating:")
        for risk_level, count in risk_counts.items():
            add_styled_paragraph(doc, f"- {risk_level}: {count}", style='List Bullet')

    # Detailed Findings Section
    add_styled_heading(doc, '3. Detailed Findings', 1)

    df = sort_by_severity(df, risk_col)

    for index, row in df.iterrows():
        # Extract data
        title = row.get('Issue Title', f'Unknown Finding #{index+1}')
        risk = row.get(risk_col, 'Unknown')
        cvss_score = row.get('CVSS Score', 'N/A')
        cvss_vector = row.get('CVSS Vector', 'N/A')
        category = row.get('OWASP Top 10', 'N/A')
        cwe_id = row.get('CWE ID', 'N/A')
        affected_modules = row.get('Affected Module / URL', 'N/A')

        # Finding Heading (3.1, 3.2...)
        add_styled_heading(doc, f"3.{index+1} {title}", 2)

        # Risk summary badge table (Risk Rating / Overall Score / CVSS Vector / OWASP Top 10 / CWE ID)
        add_risk_summary_table(doc, risk, cvss_score, cvss_vector, category, cwe_id)

        add_styled_paragraph(doc)

        # Affected Module(s) — shown right before Observation
        add_styled_heading(doc, "Affected Module(s)", 3)
        add_text_block(doc, affected_modules)

        # Write Sections
        for heading_label, col_name, justify in FINDING_SECTIONS:
            add_styled_heading(doc, heading_label, 3)
            content = row.get(col_name, 'N/A')
            add_text_block(doc, content, justify=justify)

        doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- STREAMLIT WEB DASHBOARD ---

st.set_page_config(page_title="VAPT Report Generator", layout="wide")

st.title("🛡️ Penetration Testing Report Generator")
st.write("Upload the VAPT Tracking List. The system will read the findings from the **Risk Register** tab and generate a dashboard and a downloadable Word Report.")
st.divider()

uploaded_file = st.file_uploader("Upload VAPT Excel file (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Risk Register tab
        df = pd.read_excel(uploaded_file, sheet_name="Risk Register", header=1)

        # Clean the data
        if 'Issue Title' in df.columns:
            df = df.dropna(subset=['Issue Title'])

        st.success("✅ File processed successfully!")

        risk_col = "CVSS Risk Rating" if "CVSS Risk Rating" in df.columns else "Risk Rating"
        df = sort_by_severity(df, risk_col)

        # --- WEB DASHBOARD DISPLAY ---
        st.header("📊 Executive Summary")
        col1, col2 = st.columns([1, 2])
        with col1:
            st.metric("Total Vulnerabilities Identified", len(df))
        with col2:
            if risk_col in df.columns:
                risk_counts = df[risk_col].value_counts()
                chart_df = pd.DataFrame({
                    "Risk Rating": risk_counts.index.astype(str),
                    "Count": risk_counts.values,
                })
                # Reuse the exact same color lookup that colors the Risk Rating
                # badge in the Word report, so the web chart always matches it.
                chart_df["Color"] = chart_df["Risk Rating"].apply(
                    lambda r: f"#{get_risk_style(r)['bg']}"
                )

                color_scale = alt.Scale(
                    domain=chart_df["Risk Rating"].tolist(),
                    range=chart_df["Color"].tolist(),
                )
                chart = (
                    alt.Chart(chart_df)
                    .mark_bar()
                    .encode(
                        x=alt.X("Risk Rating:N", sort="-y", title="Risk Rating"),
                        y=alt.Y("Count:Q", title="Count"),
                        color=alt.Color("Risk Rating:N", scale=color_scale, legend=None),
                        tooltip=["Risk Rating", "Count"],
                    )
                )
                st.altair_chart(chart, use_container_width=True)

        st.divider()
        st.header("📝 Detailed Findings (Web Preview)")

        for index, row in df.iterrows():
            title = row.get('Issue Title', f'Unknown Finding #{index+1}')
            risk = row.get(risk_col, 'Unknown')
            with st.expander(f"3.{index+1} [{str(risk).upper()}] {title}"):
                st.write("**Affected Module(s):**", row.get('Affected Module / URL', 'N/A'))
                st.write("**Observation:**", row.get('Observations', 'N/A'))
                st.write("**Implications:**", row.get('Implications', 'N/A'))
                st.write("**Recommendations:**", row.get('Recommendations', 'N/A'))
                st.write("**Management Comments:**", row.get('Management Comments', 'N/A'))
                st.write("**Follow-up Comments:**", row.get('Post Review Observations', 'N/A'))
                st.write("**Status:**", row.get('Status', 'Open')) # Added to preview as well

        # --- WORD DOCUMENT GENERATION & DOWNLOAD ---
        st.divider()
        st.header("📥 Export Report")
        st.write("Click below to download the fully formatted Microsoft Word document containing all findings.")

        word_file = create_word_report(df)

        st.download_button(
            label="📄 Download Word Document (.docx)",
            data=word_file,
            file_name="VAPT_Penetration_Testing_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.info("Please ensure the uploaded Excel file contains a tab named exactly 'Risk Register'.")