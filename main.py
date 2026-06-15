"""
Preise für Latex, Baumwolle, Nitril, Öl WTI & Brent, Erdgas TTF und Polyethylen
abrufen, in EUR umrechnen und in prices.xlsx speichern.

Verwendung:
  uv run main.py               – Preise abrufen und speichern
  uv run main.py -graph        – Graphen anzeigen (Standard: letzte 30 Tage)
  uv run main.py -graph -7d    – Graphen für die letzten 7 Tage
  uv run main.py -graph -90d   – Graphen für die letzten 90 Tage
"""

import re
import sys

from src.excel import append_prices, safe_convert, isnan
from src.graphs import show_graphs
from src.scrapers import (
    fetch_exchange_rates,
    fetch_latex,
    fetch_cotton,
    fetch_nitrile,
    fetch_oil_wti,
    fetch_oil_brent,
    fetch_natural_gas_ttf,
    fetch_polyethylene,
)


def _parse_days(argv: list[str]) -> int:
    for arg in argv:
        m = re.match(r"^-(\d+)[dD]$", arg)
        if m:
            return int(m.group(1))
    return 30


def main() -> None:
    argv = sys.argv[1:]

    if "-graph" in argv:
        days = _parse_days(argv)
        print(f"Zeige Graphen für die letzten {days} Tage …")
        show_graphs(days)
        return

    print("Rufe aktuelle Preise und Wechselkurse ab …")

    rates = fetch_exchange_rates()
    usd_myr = round(rates["MYR"] / rates["USD"], 4)
    usd_thb = round(rates["THB"] / rates["USD"], 4)
    print(f"  ✓ Wechselkurse: 1 EUR = {rates['USD']} USD  |  1 EUR = {rates['CNY']} CNY")
    print(f"                  1 USD = {usd_myr} MYR   |  1 USD = {usd_thb} THB")

    latex = cotton = nitrile = float("nan")
    oil_wti = oil_brent = gas_ttf = polyethylene = float("nan")

    fetchers = [
        ("Latex",       fetch_latex,           lambda v: f"{v} USD/t  →  {safe_convert(v, 'USD/t', rates)} EUR/t"),
        ("Baumwolle",   fetch_cotton,          lambda v: f"{v} USc/lb  →  {safe_convert(v, 'USc/lb', rates)} EUR/t"),
        ("Nitril",      fetch_nitrile,         lambda v: f"{v} CNY/t  →  {safe_convert(v, 'CNY/t', rates)} EUR/t"),
        ("Öl WTI",      fetch_oil_wti,         lambda v: f"{v} USD/bbl  →  {safe_convert(v, 'USD/bbl', rates)} EUR/bbl"),
        ("Öl Brent",    fetch_oil_brent,       lambda v: f"{v} USD/bbl  →  {safe_convert(v, 'USD/bbl', rates)} EUR/bbl"),
        ("Erdgas TTF",  fetch_natural_gas_ttf, lambda v: f"{v} EUR/MWh"),
        ("Polyethylen", fetch_polyethylene,    lambda v: f"{v} CNY/t  →  {safe_convert(v, 'CNY/t', rates)} EUR/t"),
    ]

    results = {}
    for name, fetcher, fmt in fetchers:
        try:
            val = fetcher()
            results[name] = val
            print(f"  ✓ {name:<14} {fmt(val)}")
        except Exception as e:
            results[name] = float("nan")
            print(f"  ✗ {name}: {e}")

    latex        = results["Latex"]
    cotton       = results["Baumwolle"]
    nitrile      = results["Nitril"]
    oil_wti      = results["Öl WTI"]
    oil_brent    = results["Öl Brent"]
    gas_ttf      = results["Erdgas TTF"]
    polyethylene = results["Polyethylen"]

    if all(isnan(v) for v in results.values()):
        print("\nAlle Abrufe fehlgeschlagen – nichts gespeichert.")
        sys.exit(1)

    append_prices(latex, cotton, nitrile, oil_wti, oil_brent, gas_ttf, polyethylene, rates)


if __name__ == "__main__":
    main()
