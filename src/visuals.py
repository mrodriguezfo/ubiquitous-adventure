import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import io
import time
from typing import List, Dict, Any
import os

STATIC_DIR = os.path.join(os.path.dirname(__file__), 'static')

def generate_informe_images(informe: List[Dict[str, Any]]) -> List[str]:
    """Genera imágenes que representen la hoja Informe.

    - Si `informe` contiene filas, genera:
      1) una tabla PNG con las filas (top summary)
      2) un gráfico PNG con Total Market Value por Account

    Devuelve lista de rutas relativas a /static (p. ej. ['/static/gen_...png']).
    """
    if not informe:
        return []

    ts = int(time.time() * 1000)
    paths = []

    # 1) Tabla imagen
    fig, ax = plt.subplots(figsize=(10, max(1, 0.5 * len(informe))))
    ax.axis('off')
    # Prepare table data
    cols = ['Account', 'Date', 'Total Market Value', 'Total Earnings', 'TWRR']
    table_data = []
    for r in informe:
        table_data.append([
            r.get('Account', ''),
            r.get('Date', ''),
            '' if r.get('Total Market Value') is None else f"{r.get('Total Market Value'):,}",
            '' if r.get('Total Earnings') is None else f"{r.get('Total Earnings'):,}",
            '' if r.get('TWRR') is None else f"{r.get('TWRR'):.6f}"
        ])
    table = ax.table(cellText=table_data, colLabels=cols, loc='center', cellLoc='left')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.2)
    out_path = os.path.join(STATIC_DIR, f'gen_informe_table_{ts}.png')
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    paths.append('/static/' + os.path.basename(out_path))

    # 2) Bar chart: Total Market Value by Account
    accounts = [r.get('Account') for r in informe]
    values = [r.get('Total Market Value') or 0 for r in informe]
    fig2, ax2 = plt.subplots(figsize=(10, 4))
    ax2.bar(accounts, values, color='tab:blue')
    ax2.set_ylabel('Total Market Value')
    ax2.set_title('Total Market Value por Account')
    ax2.tick_params(axis='x', rotation=45)
    out_path2 = os.path.join(STATIC_DIR, f'gen_informe_chart_{ts}.png')
    fig2.tight_layout()
    fig2.savefig(out_path2, dpi=150, bbox_inches='tight')
    plt.close(fig2)
    paths.append('/static/' + os.path.basename(out_path2))

    return paths


def generate_snapshots_image(snapshots: List[Dict[str, Any]]) -> List[str]:
    """Genera una imagen tipo tabla con los snapshots de mercado (prev month end y last).

    Devuelve lista con la ruta relativa a /static.
    """
    if not snapshots:
        return []
    ts = int(time.time() * 1000)
    cols = ['Account', 'PrevMonthEnd', 'PrevMarketValue', 'LastDate', 'LastMarketValue', 'Diff', 'PctDiff']
    table_data = []
    for s in snapshots:
        table_data.append([
            s.get('Account', ''),
            s.get('PrevMonthEnd', ''),
            '' if s.get('PrevMarketValue') is None else f"{s.get('PrevMarketValue'):,}",
            s.get('LastDate', ''),
            '' if s.get('LastMarketValue') is None else f"{s.get('LastMarketValue'):,}",
            '' if s.get('Diff') is None else f"{s.get('Diff'):,}",
            '' if s.get('PctDiff') is None else f"{s.get('PctDiff'):.2f}%",
        ])
    fig, ax = plt.subplots(figsize=(12, max(1, 0.5 * len(table_data))))
    ax.axis('off')
    table = ax.table(cellText=table_data, colLabels=cols, loc='center', cellLoc='right')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.1)
    out_path = os.path.join(STATIC_DIR, f'gen_snapshots_table_{ts}.png')
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    return ['/static/' + os.path.basename(out_path)]
