"""Zeitreihen-Graphen aus der Excel-Datei erzeugen."""

from datetime import datetime, timedelta

import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import numpy as np

from Visualize.Visualize import Visualize
from src.excel import get_workbook


def _safe_float(row: tuple, idx: int) -> float:
    """Gibt float(row[idx]) zurück, oder NaN falls der Index fehlt oder der Wert leer ist."""
    if idx >= len(row) or row[idx] is None:
        return float("nan")
    try:
        return float(row[idx])
    except (TypeError, ValueError):
        return float("nan")


def _plot(title: str, signals: list[tuple[list, str]], ylabel: str, x: np.ndarray) -> None:
    # Signale überspringen, bei denen alle Werte NaN sind (z. B. neue Spalten in alten Zeilen)
    valid = [(v, lbl) for v, lbl in signals if not np.all(np.isnan(np.array(v, dtype=float)))]
    if not valid:
        return

    x_ticks = np.linspace(x[0], x[-1], 5)
    viz = Visualize(title, y_label=ylabel, x_label="Datum")
    for values, label in valid:
        arr = np.array(values, dtype=float)
        mask = ~np.isnan(arr)
        viz.add_signal(arr[mask], x_axis=x[mask], label=label, scatter=True)

    _orig = viz._redraw

    def _redraw_with_dates(n: int) -> None:
        _orig(n)
        viz._ax.xaxis.set_major_locator(mticker.FixedLocator(x_ticks))
        viz._ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%y"))
        viz._ax.tick_params(axis="x", rotation=45)

    viz._redraw = _redraw_with_dates
    viz.show()


def show_graphs(days: int) -> None:
    """Öffnet Diagrammfenster für alle Rohstoffe und Wechselkurse."""
    _, ws = get_workbook()
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    if not rows:
        print("Keine Daten in der Excel-Datei vorhanden.")
        return

    cutoff = datetime.now() - timedelta(days=days)

    dates        = []
    latex_eur    = []
    cotton_eur   = []
    nitrile_eur  = []
    rate_usd     = []
    rate_cny     = []
    oil_wti_eur  = []
    oil_brent_eur= []
    gas_ttf      = []
    pe_eur       = []
    rate_myr     = []
    rate_thb     = []

    for row in rows:
        if len(row) < 9 or row[0] is None:
            continue
        dt = row[0] if isinstance(row[0], datetime) else datetime.fromisoformat(str(row[0]))
        if dt < cutoff:
            continue
        try:
            dates.append(mdates.date2num(dt))
        except Exception:
            continue

        latex_eur    .append(_safe_float(row,  2))
        cotton_eur   .append(_safe_float(row,  4))
        nitrile_eur  .append(_safe_float(row,  6))
        rate_usd     .append(_safe_float(row,  7))
        rate_cny     .append(_safe_float(row,  8))
        oil_wti_eur  .append(_safe_float(row, 10))
        oil_brent_eur.append(_safe_float(row, 12))
        gas_ttf      .append(_safe_float(row, 13))
        pe_eur       .append(_safe_float(row, 15))
        rate_myr     .append(_safe_float(row, 16))
        rate_thb     .append(_safe_float(row, 17))

    if not dates:
        print(f"Keine Daten in den letzten {days} Tagen vorhanden.")
        return

    x = np.array(dates)

    _plot(
        f"Rohstoffpreise in EUR/t – letzte {days} Tage",
        [
            (latex_eur,   "Latex (EUR/t)"),
            (cotton_eur,  "Baumwolle (EUR/t)"),
            (nitrile_eur, "Nitril (EUR/t)"),
            (pe_eur,      "Polyethylen (EUR/t)"),
        ],
        "EUR/t",
        x,
    )

    _plot(
        f"Ölpreis – letzte {days} Tage",
        [
            (oil_wti_eur,   "WTI (EUR/bbl)"),
            (oil_brent_eur, "Brent (EUR/bbl)"),
        ],
        "EUR/bbl",
        x,
    )

    _plot(
        f"Erdgas TTF – letzte {days} Tage",
        [(gas_ttf, "TTF (EUR/MWh)")],
        "EUR/MWh",
        x,
    )

    _plot(
        f"Wechselkurse EUR/USD & EUR/CNY – letzte {days} Tage",
        [
            (rate_usd, "1 EUR in USD"),
            (rate_cny, "1 EUR in CNY"),
        ],
        "Kurs",
        x,
    )

    _plot(
        f"Wechselkurse USD/MYR & USD/THB – letzte {days} Tage",
        [
            (rate_myr, "1 USD in MYR"),
            (rate_thb, "1 USD in THB"),
        ],
        "Kurs",
        x,
    )
