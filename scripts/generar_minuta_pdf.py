"""
Genera PDF de minuta con arte alineado a Solicitud_La_Internacional_Definitivo_v2.pdf
(Arial, título #1B3A5C, cabecera/pie tipo documento corporativo La Internacional).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Colores extraídos del PDF modelo (título y pies)
TITLE_BLUE = colors.HexColor("#1B3A5C")
FOOTER_GRAY = colors.HexColor("#555555")
MUTED_LINE = colors.HexColor("#888888")

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MD = ROOT / "reuniones" / "2026" / "MIN-POR-ACT-20260414-01.md"
DEFAULT_OUT = ROOT / "reuniones" / "2026" / "MIN-POR-ACT-20260414-01.pdf"

# Referencia leída del markdown (cabecera PDF)
_REF_DOC: str = ""


def _parse_referencia(md_text: str) -> str:
    """Extrae el valor de | **Referencia** | código | en la tabla inicial."""
    for line in md_text.splitlines():
        if "Referencia" not in line or "|" not in line:
            continue
        cells = [c.strip() for c in line.split("|")]
        cells = [c for c in cells if c]
        if len(cells) >= 2 and "Referencia" in cells[0]:
            return cells[1].strip()
    return ""

# Windows Arial (mismo criterio que el Word/PDF modelo)
_ARIAL = Path(r"C:\Windows\Fonts\arial.ttf")
_ARIAL_BD = Path(r"C:\Windows\Fonts\arialbd.ttf")


def _register_fonts() -> None:
    if not _ARIAL.is_file() or not _ARIAL_BD.is_file():
        raise FileNotFoundError(
            "No se encontraron Arial en C:\\Windows\\Fonts. Instale fuentes o ajuste rutas en el script."
        )
    pdfmetrics.registerFont(TTFont("Arial", str(_ARIAL)))
    pdfmetrics.registerFont(TTFont("Arial-Bold", str(_ARIAL_BD)))


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle(
            name="CoverTitle",
            fontName="Arial-Bold",
            fontSize=19,
            leading=22,
            textColor=TITLE_BLUE,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "cover_sub": ParagraphStyle(
            name="CoverSub",
            fontName="Arial",
            fontSize=11,
            leading=14,
            textColor=colors.black,
            alignment=TA_CENTER,
            spaceAfter=4,
        ),
        "cover_line": ParagraphStyle(
            name="CoverLine",
            fontName="Arial",
            fontSize=10,
            leading=13,
            textColor=FOOTER_GRAY,
            alignment=TA_CENTER,
            spaceAfter=16,
        ),
        "h2": ParagraphStyle(
            name="H2",
            fontName="Arial-Bold",
            fontSize=12,
            leading=15,
            textColor=TITLE_BLUE,
            spaceBefore=12,
            spaceAfter=8,
        ),
        "body": ParagraphStyle(
            name="Body",
            fontName="Arial",
            fontSize=11,
            leading=14,
            textColor=colors.black,
            alignment=TA_JUSTIFY,
            spaceAfter=8,
        ),
        "bullet": ParagraphStyle(
            name="Bullet",
            parent=base["Normal"],
            fontName="Arial",
            fontSize=11,
            leading=14,
            leftIndent=18,
            bulletIndent=8,
            spaceAfter=6,
        ),
        "footer": ParagraphStyle(
            name="Foot",
            fontName="Arial",
            fontSize=8,
            textColor=FOOTER_GRAY,
            alignment=TA_CENTER,
        ),
        "small_meta": ParagraphStyle(
            name="SmallMeta",
            fontName="Arial",
            fontSize=9,
            leading=11,
            textColor=colors.black,
        ),
    }


def _escape_xml(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _md_inline_to_xml(s: str) -> str:
    """**bold** y `code` a XML para Paragraph."""
    s = _escape_xml(s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"`([^`]+)`", r"<font name='Courier' size='9'>\1</font>", s)
    return s


def _parse_simple_md(md_text: str) -> list:
    """Parse mínimo del md de minuta: tablas markdown y secciones ##."""
    st = _styles()
    story: list = []
    lines = md_text.splitlines()
    i = 0
    in_table = False
    table_rows: list[list[str]] = []

    def flush_table() -> None:
        nonlocal table_rows, in_table
        if not table_rows:
            return
        data = [[Paragraph(_md_inline_to_xml(a), st["small_meta"]) for a in row] for row in table_rows]
        t = Table(data, colWidths=[1.9 * inch, 4.5 * inch])
        t.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "Arial", 9),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LINEABOVE", (0, 0), (-1, 0), 0.25, MUTED_LINE),
                    ("LINEBELOW", (0, -1), (-1, -1), 0.25, MUTED_LINE),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(t)
        story.append(Spacer(1, 10))
        table_rows = []
        in_table = False

    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("|") and "|" in line[1:]:
            if re.match(r"^\|[\s\-:|]+\|$", line.strip()):
                i += 1
                continue
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            in_table = True
            table_rows.append(cells)
            i += 1
            continue
        if in_table:
            flush_table()

        if line.startswith("## "):
            title = line[3:].strip()
            story.append(Paragraph(_md_inline_to_xml(title), st["h2"]))
            i += 1
            continue
        if line.strip() == "---":
            story.append(Spacer(1, 6))
            i += 1
            continue
        if line.strip().startswith("- "):
            items: list[str] = []
            while i < len(lines) and lines[i].strip().startswith("- "):
                items.append(lines[i].strip()[2:])
                i += 1
            for it in items:
                story.append(Paragraph("• " + _md_inline_to_xml(it), st["body"]))
            continue
        if line.strip().startswith("*") and line.strip().endswith("*") and line.count("*") >= 2:
            story.append(Paragraph(_md_inline_to_xml(line.strip().strip("*").strip()), st["body"]))
            i += 1
            continue
        if not line.strip():
            i += 1
            continue
        if line.startswith("# "):
            i += 1
            continue
        story.append(Paragraph(_md_inline_to_xml(line), st["body"]))
        i += 1

    if in_table:
        flush_table()
    return story


def _cover_block(md_text: str, styles: dict) -> list:
    """Primera página estilo portada del modelo: títulos centrados + meta."""
    lines = md_text.splitlines()
    title_line = ""
    sub_parts: list[str] = []
    for i, line in enumerate(lines):
        if line.startswith("# "):
            title_line = line[2:].strip()
        if line.strip().startswith("**La Internacional"):
            sub_parts.append(line.strip())
        if line.strip().startswith("| **Referencia**"):
            rest = lines[i:]
            block = []
            for L in rest:
                if L.strip().startswith("|") and not re.match(r"^\|[\s\-:|]+\|$", L.strip()):
                    block.append(L)
                elif block and not L.strip():
                    break
                elif block and L.strip().startswith("##"):
                    break
            story = [
                Spacer(1, 0.35 * inch),
                Paragraph(_md_inline_to_xml(title_line), styles["cover_title"]),
                Spacer(1, 0.12 * inch),
            ]
            for sp in sub_parts:
                story.append(Paragraph(_md_inline_to_xml(sp), styles["cover_sub"]))
            story.append(
                Paragraph(
                    "<i>Documento de apoyo al seguimiento de reuniones y reportería actuarial</i>",
                    styles["cover_line"],
                )
            )
            story.append(Spacer(1, 0.08 * inch))
            # Tabla meta desde markdown
            rows = []
            for L in block:
                if "---" in L or not L.strip().startswith("|"):
                    continue
                raw = [c.strip() for c in L.strip().strip("|").split("|")]
                if len(raw) >= 2:
                    rows.append([raw[0].replace("**", ""), raw[1]])
            if rows:
                data = [
                    [
                        Paragraph(f"<b>{_escape_xml(a)}</b>", styles["small_meta"]),
                        Paragraph(_md_inline_to_xml(b), styles["small_meta"]),
                    ]
                    for a, b in rows
                ]
                t = Table(data, colWidths=[1.55 * inch, 4.85 * inch])
                t.setStyle(
                    TableStyle(
                        [
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("BOX", (0, 0), (-1, -1), 0.5, TITLE_BLUE),
                            ("INNERGRID", (0, 0), (-1, -1), 0.25, MUTED_LINE),
                            ("LEFTPADDING", (0, 0), (-1, -1), 8),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                            ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F5F7FA")),
                        ]
                    )
                )
                story.append(t)
            story.append(Spacer(1, 0.2 * inch))
            story.append(Paragraph(f"<b>Versión:</b> 1.0 · <b>Clasificación:</b> Confidencial", styles["cover_line"]))
            story.append(PageBreak())
            return story
    return [Paragraph("Error parseando portada — revise el markdown.", styles["body"]), PageBreak()]


def _header_footer(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFillColor(FOOTER_GRAY)
    canvas.setFont("Arial", 8)
    width, height = LETTER
    ref = _REF_DOC or "—"
    # Cabecera interior (como página 3 del modelo)
    canvas.drawString(54, height - 42, "La Internacional de Seguros, S.A.")
    canvas.drawRightString(
        width - 54,
        height - 42,
        f"Minuta de reunión — Mesa Actuarial · {ref}",
    )
    # Pie
    canvas.setFont("Arial", 8)
    text = f"Confidencial · Página {canvas.getPageNumber()}"
    canvas.drawCentredString(width / 2, 36, text)
    canvas.restoreState()


def _first_page(canvas, doc) -> None:
    """Portada: sin cabecera duplicada; solo pie opcional."""
    canvas.saveState()
    canvas.setFillColor(FOOTER_GRAY)
    canvas.setFont("Arial", 8)
    width, _ = LETTER
    canvas.drawCentredString(width / 2, 36, "Confidencial · Página 1")
    canvas.restoreState()


def build_pdf(md_path: Path, out_path: Path) -> None:
    global _REF_DOC
    _register_fonts()
    styles = _styles()
    md_text = md_path.read_text(encoding="utf-8")
    _REF_DOC = _parse_referencia(md_text)

    story: list = _cover_block(md_text, styles)

    # Cuerpo: desde primera ## después de meta
    body_md = md_text
    # Recortar hasta después de primera tabla de participantes
    if "## 1." in body_md:
        body_md = body_md.split("## 1.", 1)[1]
        body_md = "## 1." + body_md
    story.extend(_parse_simple_md(body_md))

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=LETTER,
        leftMargin=54,
        rightMargin=54,
        topMargin=56,
        bottomMargin=50,
        title="Minuta de reunión — Mesa Actuarial",
        author="Angel Colmenares",
    )

    doc.build(story, onFirstPage=_first_page, onLaterPages=_header_footer)


def main() -> None:
    md = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_MD
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_OUT
    if not md.is_file():
        print(f"No existe: {md}", file=sys.stderr)
        sys.exit(1)
    out.parent.mkdir(parents=True, exist_ok=True)
    build_pdf(md, out)
    print(f"PDF generado: {out}")


if __name__ == "__main__":
    main()
