from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from io import BytesIO


def _fmt_money(v):
    try:
        if v is None or v == '':
            return ''
        n = float(v)
        return f"${n:,.2f}"
    except Exception:
        return str(v or '')


def generate_pdf(results: dict, selected_date: str) -> bytes:
    """Generate a PDF report with a clean, tabular format suitable for printing.

    The layout matches the requested style: title line, header row with light background,
    right-aligned numeric columns, and consistent column widths.
    """
    rows = results.get('rows', [])
    snapshots = results.get('snapshots', [])
    snap_map = {s.get('Account'): s for s in snapshots}

    # map rows by account+date
    row_map = {}
    for r in rows:
        key = (r.get('Account') or '') + '|' + (r.get('End Date') or r.get('Date') or r.get('LastDate') or '')
        row_map[key] = r

    accounts = sorted(list({r.get('Account') for r in rows if r.get('Account')}))

    buffer = BytesIO()
    page_size = landscape(letter)
    doc = SimpleDocTemplate(buffer, pagesize=page_size, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('title', parent=styles['Heading2'], alignment=1, spaceAfter=6)
    subtitle_style = ParagraphStyle('subtitle', parent=styles['Normal'], alignment=1, textColor=colors.HexColor('#666'), spaceAfter=12)

    title = Paragraph(f"Valor de Mercado - {selected_date}", title_style)
    subtitle = Paragraph('', subtitle_style)
    elements.append(title)
    elements.append(subtitle)

    # Table header
    header = ["Portafolio", "Prev Month End", "Valor PrevMonthEnd", f"{selected_date}", "Valor", "Diff", "% Dif"]
    data = [header]

    for acct in accounts:
        snap = snap_map.get(acct, {})
        prev_date = snap.get('PrevMonthEnd', '')
        prev_val = snap.get('PrevMarketValue', '')
        key = acct + '|' + (selected_date or '')
        row = row_map.get(key, {})
        sel_val = row.get('Total Market Value', '') if row else ''

        # attempt numeric calculations
        try:
            prev_n = float(prev_val) if prev_val not in (None, '') else None
        except Exception:
            prev_n = None
        try:
            sel_n = float(sel_val) if sel_val not in (None, '') else None
        except Exception:
            sel_n = None

        if prev_n is not None and sel_n is not None:
            diff = sel_n - prev_n
            pct = (diff / prev_n * 100) if prev_n != 0 else None
        else:
            diff = None
            pct = None

        data.append([
            acct or '',
            prev_date or '',
            _fmt_money(prev_n),
            selected_date or '',
            _fmt_money(sel_n),
            _fmt_money(diff),
            (f"{pct:.2f}%" if pct is not None else '')
        ])

    # Column widths (points): sum should fit page width minus margins
    total_width = page_size[0] - doc.leftMargin - doc.rightMargin
    col_widths = [total_width * w for w in (0.20, 0.12, 0.14, 0.12, 0.18, 0.12, 0.12)]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#333')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor('#dfe7ef')),
        ('ALIGN', (2,1), (2,-1), 'RIGHT'),
        ('ALIGN', (4,1), (5,-1), 'RIGHT'),
        ('ALIGN', (6,1), (6,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ])

    # alternate row background
    for i in range(1, len(data)):
        if i % 2 == 0:
            tbl_style.add('BACKGROUND', (0,i), (-1,i), colors.HexColor('#fbfdff'))

    table.setStyle(tbl_style)
    elements.append(table)

    # Build document
    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf
