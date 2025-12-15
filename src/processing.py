import io
import csv
import datetime
from typing import List, Optional

import pandas as pd


def process_retornos_csv(contents: bytes, account: Optional[str] = None) -> List[dict]:
    """Replicar la transformación de PowerQuery para RetornosV21.csv

    - Saltea 5 filas iniciales
    - Promueve headers
    - Convierte tipos de columnas (se mapearán según el M script)
    - Filtra por Account si se suministra
    """
    # Leer como texto con encoding cp1252 (1252) similar al PowerQuery
    text = contents.decode('cp1252', errors='replace')
    # Detectar delimitador (coma o punto y coma u otro) en la línea de encabezado (después de saltar 5 filas)
    lines = text.splitlines()
    header_line = ''
    if len(lines) > 5:
        header_line = lines[5]
    else:
        header_line = '\n'.join(lines)

    # Intentar detectar delimitador con csv.Sniffer
    dialect = None
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(header_line)
        sep = dialect.delimiter
    except Exception:
        sep = ','

    # Leer con pandas usando el delimitador detectado
    df = pd.read_csv(io.StringIO(text), skiprows=5, header=0, sep=sep, encoding='cp1252', engine='python')

    # Normalizar columnas: rename to known names if present
    expected = ["Account", "Begin Date", "End Date", "Perf. Class", "Settlement Date Cash Balance", "Total Market Value", "TWRR", "TWRR M-T-D", "TWRR Y-T-D", "TWRR 3 month", "TWRR 1 yr.", "TWRR 3 yr. Ann.", "TWRR Incept. Ann.", "TWRR w/ Fees", "TWRR w/Fees M-T-D", "TWRR w/Fees Y-T-D", "Total Earnings"]
    # If df has more or fewer columns, try to align by position
    if len(df.columns) >= len(expected):
        df = df.iloc[:, :len(expected)]
        df.columns = expected
    else:
        # fallback: keep current names
        df.columns = list(df.columns)

    # Convert types
    for col in ["Begin Date", "End Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    for col in ["Settlement Date Cash Balance", "Total Market Value", "TWRR", "TWRR M-T-D", "TWRR Y-T-D", "TWRR 3 month", "TWRR 1 yr.", "TWRR 3 yr. Ann.", "TWRR Incept. Ann.", "TWRR w/ Fees", "TWRR w/Fees M-T-D", "TWRR w/Fees Y-T-D", "Total Earnings"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')

    # Filter by account
    if account and "Account" in df.columns:
        df = df[df['Account'] == account]

    # Remove columns matching PowerQuery remove: "TWRR w/ Fees", "TWRR w/Fees M-T-D", "TWRR w/Fees Y-T-D"
    for c in ["TWRR w/ Fees", "TWRR w/Fees M-T-D", "TWRR w/Fees Y-T-D"]:
        if c in df.columns:
            df = df.drop(columns=[c])

    # Compute additional aggregates similar to the Informe sheet
    # Example: summary by Account: latest Total Market Value, Total Earnings, and TWRR metrics
    summaries = []
    if 'Account' in df.columns:
        for acct, g in df.groupby('Account'):
            latest_row = g.sort_values('End Date').iloc[-1]
            summaries.append({
                'Account': acct,
                'Date': latest_row['End Date'].strftime('%Y-%m-%d') if not pd.isnull(latest_row['End Date']) else None,
                'Total Market Value': float(latest_row['Total Market Value']) if 'Total Market Value' in g.columns else None,
                'Total Earnings': float(latest_row['Total Earnings']) if 'Total Earnings' in g.columns else None,
                'TWRR': float(latest_row['TWRR']) if 'TWRR' in g.columns else None,
            })

    # Return both rows and a computed informe-style summary
    # Rows: convert values to serializable
    results = []
    for _, row in df.iterrows():
        r = {}
        for col in df.columns:
            val = row[col]
            if pd.isnull(val):
                v = None
            elif isinstance(val, (pd.Timestamp, datetime.datetime)):
                v = val.strftime('%Y-%m-%d')
            elif isinstance(val, (float, int)):
                # convert numpy types to python native
                v = float(val) if not pd.isna(val) else None
            else:
                v = str(val)
            r[col if isinstance(col, str) else str(col)] = v
        results.append(r)

    return {'rows': results, 'count': len(results), 'informe': summaries}
