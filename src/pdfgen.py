from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO


def generate_pdf(results: dict, selected_date: str) -> bytes:
    """Generate a simple PDF report from results dict and selected date.

    Returns PDF bytes.
    """
    rows = results.get('rows', [])
    snapshots = results.get('snapshots', [])
    # map snapshots by account
    snap_map = {s.get('Account'): s for s in snapshots}

    # map rows by account+date
    row_map = {}
    for r in rows:
        key = (r.get('Account') or '') + '|' + (r.get('End Date') or r.get('Date') or r.get('LastDate') or '')
        row_map[key] = r

    accounts = sorted(list({r.get('Account') for r in rows if r.get('Account')}))

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph(f"Valor de Mercado - {selected_date}", styles['Heading2'])
    elements.append(title)
    elements.append(Spacer(1, 8))

    data = [["Portafolio", "Prev Month End", "Valor PrevMonthEnd", f"{selected_date}", "Valor", "Diff", "% Dif"]]
    for acct in accounts:
        snap = snap_map.get(acct, {})
        prev_date = snap.get('PrevMonthEnd', '')
        prev_val = snap.get('PrevMarketValue', '')
        key = acct + '|' + (selected_date or '')
        row = row_map.get(key, {})
        sel_val = row.get('Total Market Value', '') if row else ''
        try:
            diff = float(sel_val) - float(prev_val)
        except Exception:
            diff = ''
        try:
            pct = (diff / float(prev_val) * 100) if (diff != '' and prev_val not in (None, '', 0)) else ''
        except Exception:
            pct = ''
        data.append([acct, prev_date, f"{prev_val}", selected_date, f"{sel_val}", f"{diff}", f"{pct}" if pct == '' else f"{pct:.2f}%"])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f4f8')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#333')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dfe7ef')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (2,0), (-1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))

    elements.append(table)
    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf
