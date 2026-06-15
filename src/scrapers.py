"""Alle Web-Scraper für Rohstoffpreise und Wechselkurse."""

import re
from base64 import b64encode

import cloudscraper
import requests
from bs4 import BeautifulSoup

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
_LGM_API = "https://www.lgm.gov.my/webv2api/api/rubberprice/currentprice"
_LGM_AUTH = "Basic " + b64encode(b"FOB:LgMF0b$2025").decode()


# ---------------------------------------------------------------------------
# Wechselkurse
# ---------------------------------------------------------------------------

def fetch_exchange_rates() -> dict[str, float]:
    """EUR-basierte Wechselkurse abrufen (kostenlose API).

    Rückgabe: USD, CNY, MYR, THB jeweils als „1 EUR = X".
    USD/MYR und USD/THB werden in main.py als Kreuzrate berechnet.
    """
    try:
        resp = requests.get(
            "https://api.exchangerate-api.com/v4/latest/EUR",
            headers=_HEADERS,
            timeout=10,
        )
        resp.raise_for_status()
        r = resp.json()["rates"]
        return {
            "USD": float(r["USD"]),
            "CNY": float(r["CNY"]),
            "MYR": float(r["MYR"]),
            "THB": float(r["THB"]),
        }
    except Exception as e:
        print(f"  ! Wechselkurs-Abruf fehlgeschlagen ({e}), verwende Näherungswerte")
        return {"USD": 1.08, "CNY": 7.80, "MYR": 4.72, "THB": 37.50}


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _onvista_price(url: str, name: str) -> float:
    """Aktuellen Kurs via <data class='…text-4xl…font-bold…' value='…'> von onvista.de."""
    scraper = cloudscraper.create_scraper()
    resp = scraper.get(url, headers=_HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    tag = soup.find("data", class_=lambda c: c and "text-4xl" in c and "font-bold" in c)
    if tag and tag.get("value"):
        return float(tag["value"])
    raise ValueError(f"{name}: Preis-Element nicht auf {url} gefunden")


def _tradingeconomics_price(url: str, name: str, row_label: str | None = None) -> float:
    """Preis von tradingeconomics.com abrufen.

    Wenn row_label gesetzt ist, wird die erste Tabellenzeile mit diesem Text
    gesucht (für Seiten mit mehreren Rohstoffen).
    """
    scraper = cloudscraper.create_scraper()
    resp = scraper.get(url, headers=_HEADERS, timeout=25)
    resp.raise_for_status()
    text = resp.text

    # Spezifische Tabellenzeile suchen
    if row_label:
        soup = BeautifulSoup(text, "html.parser")
        for row in soup.find_all("tr"):
            if row_label in row.get_text():
                td = row.find("td", id="p")
                if td:
                    raw = td.get_text(strip=True).replace(",", "")
                    return float(raw)

    # JSON-Muster (Realtime-Wert)
    m = re.search(r'"last"\s*:\s*([\d.]+)', text)
    if m:
        val = float(m.group(1))
        if val > 0:
            return val

    raise ValueError(f"{name}: Kein Preis auf {url} gefunden")


# ---------------------------------------------------------------------------
# Rohstoffe – bestehend
# ---------------------------------------------------------------------------

def fetch_latex() -> float:
    """Latex in Bulk (USD/t) von der LGM-API (Malaysian Government)."""
    resp = requests.get(
        _LGM_API,
        headers={**_HEADERS, "Authorization": _LGM_AUTH},
        timeout=15,
    )
    resp.raise_for_status()
    for item in resp.json():
        if item.get("grade") == "Latex in Bulk":
            return float(item["sellersUs"])
    raise ValueError("Grade 'Latex in Bulk' nicht in LGM-Antwort gefunden")


def fetch_cotton() -> float:
    """Baumwolle Future (USc/lb) von onvista.de."""
    return _onvista_price(
        "https://www.onvista.de/rohstoffe/Baumwolle-Future-21474937",
        "Baumwolle",
    )


def fetch_nitrile() -> float:
    """Nitril/Butadien Benchmark (CNY/t) von 100ppi.com."""
    scraper = cloudscraper.create_scraper()
    resp = scraper.get("https://www.100ppi.com/vane/detail-886.html", timeout=20)
    resp.raise_for_status()
    m = re.search(r"基准价[为：:]([\d,]+\.?\d*)元/吨", resp.text)
    if not m:
        m = re.search(r"([\d,]{4,}\.?\d*)元/吨", resp.text)
    if not m:
        raise ValueError("Nitril-Preis-Muster nicht auf 100ppi gefunden")
    return float(m.group(1).replace(",", ""))


# ---------------------------------------------------------------------------
# Rohstoffe – neu
# ---------------------------------------------------------------------------

def fetch_oil_wti() -> float:
    """WTI Rohöl (USD/bbl) von onvista.de (finanzen.net-Gruppe)."""
    return _onvista_price(
        "https://www.onvista.de/rohstoffe/Oel-WTI-Future-6988820",
        "Öl WTI",
    )


def fetch_oil_brent() -> float:
    """Brent Rohöl (USD/bbl) von onvista.de (finanzen.net-Gruppe)."""
    return _onvista_price(
        "https://www.onvista.de/rohstoffe/Oel-Brent-Future-6988832",
        "Öl Brent",
    )


def fetch_natural_gas_ttf() -> float:
    """Dutch TTF Natural Gas Future (EUR/MWh) von tradingeconomics.com.

    Entspricht inhaltlich der ING-Markets-Seite 'Dutch TTF Natural Gas Future'.
    """
    return _tradingeconomics_price(
        "https://tradingeconomics.com/commodity/eu-natural-gas",
        "TTF Erdgas",
    )


def fetch_polyethylene() -> float:
    """Polyethylen (CNY/t) von tradingeconomics.com.

    Preis bezieht sich auf den asiatischen Markt (chinesischer PE-Futures-Preis).
    Umrechnung in EUR/t erfolgt über den EUR/CNY-Kurs.
    """
    return _tradingeconomics_price(
        "https://tradingeconomics.com/commodity/polyethylene",
        "Polyethylen",
        row_label="Polyethylene",
    )
