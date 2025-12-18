import io
import csv
import datetime
from typing import List, Optional

import pandas as pd


def process_retornos_csv(contents: bytes, account: Optional[str] = None, support_files: Optional[dict] = None) -> dict:
    """Replicar la transformación de PowerQuery para RetornosV21.csv

    - Saltea 5 filas iniciales
    - Promueve headers
    - Convierte tipos de columnas (se mapearán según el M script)
    - Filtra por Account si se suministra
    """
    # Leer como texto con encoding cp1252 (1252) similar al PowerQuery
    text = contents.decode('cp1252', errors='replace')
    lines = text.splitlines()
    # Try to locate the header row. Some files (PowerQuery output) skip 5 rows,
    # but uploaded CSVs may have header on the first line. Search first 10 lines
    # for a line that contains expected header names (e.g. 'Account').
    header_idx = None
    header_line = ''
    for i, ln in enumerate(lines[:10]):
        if 'Account' in ln and ('Total' in ln or 'Total Market' in ln or 'End Date' in ln):
            header_idx = i
            header_line = ln
            break
    if header_idx is None:
        # default to PowerQuery behavior (skip 5 rows) if enough lines exist,
        # otherwise assume header is first line
        if len(lines) > 5:
            header_idx = 5
            header_line = lines[5]
        else:
            header_idx = 0
            header_line = lines[0] if lines else ''
    # Detect delimiter using csv.Sniffer on the detected header_line
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(header_line) if header_line else None
        sep = dialect.delimiter if dialect else ','
    except Exception:
        sep = ','
    # Read with pandas using the detected header row index.
    # Be defensive: uploaded files may not match PowerQuery output. Try multiple fallbacks
    # before giving up.
    if not text.strip():
        # empty content
        return {'rows': [], 'count': 0, 'informe': [], 'snapshots': [], 'benchmarks': []}

    try:
        df = pd.read_csv(io.StringIO(text), skiprows=header_idx, header=0, sep=sep, encoding='cp1252', engine='python')
    except Exception:
        # try reading without skipping rows (header at top)
        try:
            df = pd.read_csv(io.StringIO(text), header=0, sep=',', encoding='cp1252', engine='python')
        except Exception:
            # final fallback: let pandas infer delimiter with default engine
            try:
                df = pd.read_csv(io.StringIO(text), header=0, encoding='cp1252')
            except Exception:
                # can't parse file
                raise
    # If df has no columns (sometimes pandas returns empty), raise a clear error
    if df is None or df.shape[1] == 0:
        raise pd.errors.EmptyDataError('No columns parsed from uploaded CSV')

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

    # Market value snapshots: end of previous month (relative to overall latest date) and last available date
    snapshots = []
    if 'End Date' in df.columns and 'Total Market Value' in df.columns and len(df) > 0:
        overall_last = df['End Date'].max()
        if not pd.isnull(overall_last):
            # previous month end relative to overall_last
            first_of_month = overall_last.replace(day=1)
            prev_month_end = first_of_month - pd.Timedelta(days=1)

            total_prev = 0.0
            total_last = 0.0
            for acct, g in df.groupby('Account'):
                # value at or before prev_month_end
                g_sorted = g.sort_values('End Date')
                prev_rows = g_sorted[g_sorted['End Date'] <= prev_month_end]
                if not prev_rows.empty:
                    prev_val = float(prev_rows.iloc[-1]['Total Market Value']) if not pd.isnull(prev_rows.iloc[-1]['Total Market Value']) else None
                    prev_date = prev_rows.iloc[-1]['End Date']
                else:
                    prev_val = None
                    prev_date = None

                last_row = g_sorted.iloc[-1]
                last_val = float(last_row['Total Market Value']) if not pd.isnull(last_row['Total Market Value']) else None
                last_date = last_row['End Date']

                diff = None
                pct = None
                if prev_val is not None and last_val is not None:
                    try:
                        diff = last_val - prev_val
                        pct = (diff / prev_val) * 100 if prev_val != 0 else None
                    except Exception:
                        diff = None
                        pct = None

                snapshots.append({
                    'Account': acct,
                    'PrevMonthEnd': prev_month_end.strftime('%Y-%m-%d'),
                    'PrevDate': prev_date.strftime('%Y-%m-%d') if prev_date is not None else None,
                    'PrevMarketValue': prev_val,
                    'LastDate': last_date.strftime('%Y-%m-%d') if not pd.isnull(last_date) else None,
                    'LastMarketValue': last_val,
                    'Diff': diff,
                    'PctDiff': pct,
                })

                if prev_val is not None:
                    total_prev += prev_val
                if last_val is not None:
                    total_last += last_val

            # add total row
            snapshots.append({
                'Account': 'Total',
                'PrevMonthEnd': prev_month_end.strftime('%Y-%m-%d'),
                'PrevDate': None,
                'PrevMarketValue': total_prev,
                'LastDate': overall_last.strftime('%Y-%m-%d') if not pd.isnull(overall_last) else None,
                'LastMarketValue': total_last,
                'Diff': (total_last - total_prev) if (total_prev is not None) else None,
                'PctDiff': ((total_last - total_prev) / total_prev * 100) if (total_prev not in (None, 0)) else None,
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

    # Sanitize outputs: convert NaN/inf to None and convert numpy/pandas types to native Python
    import math
    def _clean(v):
        # pandas NA / numpy nan -> None
        try:
            if pd.isna(v):
                return None
        except Exception:
            pass
        if isinstance(v, (pd.Timestamp, datetime.datetime)):
            return v.strftime('%Y-%m-%d')
        if isinstance(v, datetime.date):
            return v.strftime('%Y-%m-%d')
        if isinstance(v, (int, float)):
            # reject NaN/inf
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                return None
            return float(v) if isinstance(v, (int, float)) else v
        return v

    # clean summaries
    clean_summaries = []
    for s in summaries:
        cs = {k: _clean(v) for k, v in s.items()}
        clean_summaries.append(cs)

    # clean snapshots
    clean_snapshots = []
    for s in snapshots:
        cs = {k: _clean(v) for k, v in s.items()}
        clean_snapshots.append(cs)

    # rows already converted but ensure types cleaned
    clean_results = []
    for r in results:
        cr = {k: _clean(v) for k, v in r.items()}
        clean_results.append(cr)

    # derive benchmarks from any accounts that include 'BMK' in the Account name
    benchmarks = []
    try:
        if 'Account' in df.columns and 'TWRR' in df.columns:
            bmk_df = df[df['Account'].str.contains('BMK', case=False, na=False)]
            for acct, g in bmk_df.groupby('Account'):
                latest = g.sort_values('End Date').iloc[-1]
                benchmarks.append({
                    'Account': acct,
                    'Date': latest['End Date'].strftime('%Y-%m-%d') if not pd.isnull(latest['End Date']) else None,
                    'TWRR': float(latest['TWRR']) if 'TWRR' in latest and not pd.isnull(latest['TWRR']) else None,
                })
    except Exception:
        benchmarks = []

    clean_benchmarks = []
    for b in benchmarks:
        cb = {k: _clean(v) for k, v in b.items()}
        clean_benchmarks.append(cb)

    return {'rows': clean_results, 'count': len(clean_results), 'informe': clean_summaries, 'snapshots': clean_snapshots, 'benchmarks': clean_benchmarks}
