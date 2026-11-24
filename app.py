from __future__ import annotations

import io
import shutil
from pathlib import Path

import fitz
import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="PDF to Word Layout Engine V2", page_icon="📄", layout="wide")
st.title("PDF → Editable Word | Layout Engine V2")
st.caption("Coordinate-based Malayalam OCR reconstruction. Output contains editable text, not page images.")

st.warning(
    "This engine reconstructs editable text lines at their detected page coordinates. "
    "It does not embed page screenshots or add a duplicate OCR section. Exact typography "
    "still depends on OCR accuracy and fonts installed on the computer."
)

with st.sidebar:
    st.header("Reconstruction settings")
    dpi = st.slider("OCR render DPI", 200, 500, 350, 25)
    languages = st.text_input("Tesseract languages", "mal+eng")
    min_conf = st.slider("Minimum OCR confidence", 0, 95, 15, 5)
    tesseract_path = st.text_input(
        "Tesseract executable",
        r"C:\Users\megha.sabu\Documents\Python Files Folder\PDF COVERTED\Tesseract-OCR\tesseract.exe"
    )
    font_name = st.text_input("Word font", "Noto Sans Malayalam")
    font_size = st.slider("Base font size (pt)", 7.0, 18.0, 10.0, 0.5)
    line_height_scale = st.slider("OCR line-height scale", 0.7, 1.5, 1.0, 0.05)
    left_adjust = st.slider("Horizontal calibration (pt)", -30, 30, 0, 1)
    top_adjust = st.slider("Vertical calibration (pt)", -30, 30, 0, 1)
    keep_page_count = st.checkbox("Force one Word section per PDF page", True)

pdf_file = st.file_uploader("Upload the original PDF", type=["pdf"])

def set_font(run, name, size):
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for slot in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{slot}"), name)
    # Request complex-script font handling for Malayalam.
    rpr.set(qn("w:lang"), "ml-IN")

def ocr_page(page, dpi, lang, min_conf):
    import pytesseract
    from pytesseract import Output

    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    data = pytesseract.image_to_data(
        image, lang=lang, config="--psm 6",
        output_type=Output.DICT
    )

    rows = {}
    for i, raw in enumerate(data["text"]):
        token = raw.strip()
        if not token:
            continue
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if conf < min_conf:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        rows.setdefault(key, []).append({
            "text": token,
            "x": int(data["left"][i]) / scale,
            "y": int(data["top"][i]) / scale,
            "w": int(data["width"][i]) / scale,
            "h": int(data["height"][i]) / scale,
            "conf": conf,
        })

    lines = []
    for words in rows.values():
        words.sort(key=lambda w: w["x"])
        text = " ".join(w["text"] for w in words)
        x = min(w["x"] for w in words)
        y = min(w["y"] for w in words)
        h = max(w["h"] for w in words)
        avg_conf = sum(w["conf"] for w in words) / len(words)
        lines.append({"text": text, "x": x, "y": y, "h": h, "conf": avg_conf})
    lines.sort(key=lambda v: (v["y"], v["x"]))
    return lines, image

if pdf_file and st.button("Reconstruct editable Word", type="primary"):
    try:
        import pytesseract
        exe = Path(tesseract_path.strip().strip('"')) if tesseract_path.strip() else None
        if exe and exe.exists():
            pytesseract.pytesseract.tesseract_cmd = str(exe)
        elif shutil.which("tesseract") is None:
            st.error(
                "Tesseract OCR executable was not found. Install Tesseract for Windows "
                "and Malayalam (`mal`) + English (`eng`) language data. The PDF is not converted."
            )
            st.stop()

        available = pytesseract.get_languages(config="")
        requested = [x.strip() for x in languages.split("+") if x.strip()]
        missing = [x for x in requested if x not in available]
        if missing:
            st.error(f"Missing Tesseract language data: {missing}. Available: {available}")
            st.stop()

        pdf = fitz.open(stream=pdf_file.getvalue(), filetype="pdf")
        if not pdf.page_count:
            st.error("PDF contains no pages.")
            st.stop()

        doc = Document()
        # Configure first page section, then one new-page section for each later PDF page.
        results = []
        progress = st.progress(0)
        status = st.empty()

        for page_idx, page in enumerate(pdf):
            status.write(f"Analyzing PDF page {page_idx + 1}/{pdf.page_count}")
            lines, rendered = ocr_page(page, dpi, languages, min_conf)
            results.append(lines)

            if page_idx == 0:
                sec = doc.sections[0]
            else:
                sec = doc.add_section(WD_SECTION.NEW_PAGE)

            # PDF points are Word points; set exact original page size.
            sec.page_width = Pt(page.rect.width)
            sec.page_height = Pt(page.rect.height)
            sec.top_margin = Pt(0)
            sec.bottom_margin = Pt(0)
            sec.left_margin = Pt(0)
            sec.right_margin = Pt(0)
            sec.header_distance = Pt(0)
            sec.footer_distance = Pt(0)

            # Each OCR line is an editable paragraph. Vertical gaps are derived
            # from OCR top coordinates; horizontal position from the left bounding box.
            previous_baseline = 0.0
            for line in lines:
                y = max(0.0, line["y"] + top_adjust)
                x = max(0.0, line["x"] + left_adjust)
                height = max(1.0, line["h"] * line_height_scale)

                para = doc.add_paragraph()
                fmt = para.paragraph_format
                fmt.left_indent = Pt(x)
                fmt.right_indent = Pt(4)
                fmt.space_before = Pt(max(0.0, y - previous_baseline))
                fmt.space_after = Pt(0)
                fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                fmt.line_spacing = Pt(max(font_size * 1.15, height))
                fmt.keep_together = True
                fmt.widow_control = False

                run = para.add_run(line["text"])
                set_font(run, font_name, font_size)
                previous_baseline = y + height

            # Explicitly end section/page with page break only via section break
            # on the next loop iteration; no extra break paragraph is inserted.
            progress.progress((page_idx + 1) / pdf.page_count)

        # Prevent Word from adding unintended blank content paragraphs.
        output = io.BytesIO()
        doc.core_properties.title = "Editable PDF layout reconstruction"
        doc.core_properties.subject = "Coordinate-based OCR reconstruction"
        doc.core_properties.comments = (
            "Editable text reconstructed from OCR coordinates. Proofread OCR and compare "
            "against the source PDF; exact font reproduction is not guaranteed."
        )
        doc.save(output)
        output.seek(0)

        st.success("Editable reconstruction completed.")
        st.metric("Source PDF pages", pdf.page_count)
        st.metric("Recognized editable lines", sum(len(x) for x in results))

        with st.expander("OCR review by page"):
            for i, lines in enumerate(results, 1):
                st.markdown(f"**Page {i}: {len(lines)} lines**")
                st.text("\n".join(line["text"] for line in lines))

        st.download_button(
            "Download editable Word document",
            data=output.getvalue(),
            file_name=f"{Path(pdf_file.name).stem}_layout_reconstructed.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        st.info(
            "The DOCX contains editable text paragraphs only. It does not insert PDF page "
            "images or a second transcription section. OCR mistakes must be corrected manually. "
            "For best Malayalam rendering, install Noto Sans Malayalam on the computer opening Word."
        )

    except Exception as exc:
        st.exception(exc)
else:
    st.caption("Upload the original PDF, verify OCR settings, and start reconstruction.")
