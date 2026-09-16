#!/usr/bin/env python3
"""
car_tracker.py
--------------
Monitoruje oferty samochodów z OLX, Otomoto, Autoplac i Facebook Marketplace
zapisane w pliku offers.csv.
Dla każdej oferty:
  - pobiera aktualną cenę, przebieg i lokalizację
  - porównuje z poprzednim odczytem (historia zmian cen)
  - status jest zawsze jednym z dwóch: "DOSTĘPNE" albo "BRAK OGŁOSZENIA"
    (na "BRAK OGŁOSZENIA" zmienia się tylko przy twardym dowodzie - HTTP 404
    albo komunikacie strony o nieaktualności; wszelkie błędy/blokady zostają
    jako "DOSTĘPNE" z opisem w Uwagach, żeby nie było fałszywych alarmów)
  - zapisuje wszystko do tracker.xlsx (kolorowanie, arkusz z historią)

Uruchamiaj cyklicznie (np. co godzinę) przez Harmonogram zadań (Windows)
lub cron (Mac/Linux) - patrz README.md.

WAŻNE: OLX i Otomoto regularnie zmieniają strukturę HTML swoich stron
oraz mogą blokować automatyczne zapytania (captcha, blokada IP przy zbyt
częstym odpytywaniu). Skrypt korzysta w pierwszej kolejności z danych
strukturalnych (JSON-LD), które są najbardziej odporne na zmiany
wizualne strony, a dopiero w drugiej kolejności z selektorów CSS jako
fallback. Jeśli po jakimś czasie przestanie działać - to najpewniej
strona zmieniła strukturę i selektory (SELECTORS_* poniżej) trzeba
będzie zaktualizować.
"""

import csv
import html
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Brakuje openpyxl. Zainstaluj: pip install openpyxl")
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent
OFFERS_CSV = BASE_DIR / "offers.csv"
OUTPUT_XLSX = BASE_DIR / "tracker.xlsx"
LOG_FILE = BASE_DIR / "tracker.log"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8",
}

SOLD_PHRASES = [
    "ogłoszenie zakończone", "ogłoszenie nieaktualne", "ogłoszenie wygasło",
    "to ogłoszenie jest nieaktualne", "oferta niedostępna", "not found",
    "strona nie została znaleziona", "ten produkt nie jest już dostępny",
    "this listing is no longer available", "ogłoszenie zostało usunięte",
    "content isn't available", "content not found",
]


def log(msg: str):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_offers():
    if not OFFERS_CSV.exists():
        log(f"Nie znaleziono {OFFERS_CSV}. Utwórz plik offers.csv wg wzoru z README.")
        sys.exit(1)
    with open(OFFERS_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def extract_json_ld(soup):
    """Szuka danych structured data (schema.org) - najbardziej stabilne źródło ceny."""
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "{}")
        except (json.JSONDecodeError, TypeError):
            continue
        candidates = data if isinstance(data, list) else [data]
        for item in candidates:
            offers = item.get("offers") if isinstance(item, dict) else None
            if offers:
                offers = offers if isinstance(offers, list) else [offers]
                for off in offers:
                    price = off.get("price")
                    if price:
                        mileage = (
                            item.get("mileageFromOdometer", {}).get("value")
                            if isinstance(item.get("mileageFromOdometer"), dict)
                            else item.get("mileageFromOdometer")
                        )
                        image = item.get("image")
                        if isinstance(image, list):
                            image = image[0] if image else None
                        elif isinstance(image, dict):
                            image = image.get("url")
                        return {
                            "price": str(price),
                            "currency": off.get("priceCurrency", "PLN"),
                            "title": item.get("name"),
                            "mileage": str(mileage) if mileage else None,
                            "image": image,
                        }
    return None


def extract_image_fallback(soup):
    """Plan B na zdjęcie: tag og:image (meta tag do udostępniania w social media,
    prawie zawsze obecny i wskazuje na pierwsze/główne zdjęcie oferty)."""
    tag = soup.find("meta", property="og:image")
    if tag and tag.get("content"):
        return tag["content"]
    img = soup.find("img")
    if img and img.get("src"):
        return img["src"]
    return None


def parse_price_fallback(soup, site):
    """Selektory CSS jako plan B, gdyby JSON-LD nie było dostępne.
    UWAGA: te selektory trzeba czasem poprawić ręcznie, bo strony się zmieniają."""
    text = soup.get_text(" ", strip=True)
    match = re.search(r"(\d[\d\s]{2,10})\s*zł", text)
    price = match.group(1).replace(" ", "") if match else None

    mileage_match = re.search(r"(\d[\d\s]{2,7})\s*km\b", text)
    mileage = mileage_match.group(1).replace(" ", "") if mileage_match else None

    location = None
    if site == "olx":
        loc_el = soup.select_one('[data-testid="location-date"]')
        if loc_el:
            location = loc_el.get_text(strip=True).split(" - ")[0]
    elif site == "otomoto":
        loc_el = soup.select_one('[data-testid="location"]') or soup.select_one(".offer-meta__location")
        if loc_el:
            location = loc_el.get_text(strip=True)

    return price, location, mileage


def detect_site(url: str) -> str:
    if "olx.pl" in url:
        return "olx"
    if "otomoto.pl" in url:
        return "otomoto"
    if "facebook.com" in url:
        return "facebook"
    if "autoplac.pl" in url:
        return "autoplac"
    return "inne"


def site_display_name(site: str) -> str:
    return {
        "olx": "OLX",
        "otomoto": "Otomoto",
        "facebook": "Facebook Marketplace",
        "autoplac": "Autoplac",
    }.get(site, "Inna strona")


def check_offer(url: str):
    """Zwraca dict: status, price, currency, location, title, mileage, image, note.

    Status jest celowo uproszczony do dwóch wartości:
      - "DOSTĘPNE"       - domyślny stan; ustawiany zawsze, chyba że mamy
                           TWARDY dowód, że ogłoszenie zniknęło
      - "BRAK OGŁOSZENIA" - tylko gdy strona zwróciła 404 albo w jej treści
                           wprost pojawił się komunikat o nieaktualności

    To celowa zasada bezpieczeństwa: błąd sieci, blokada strony (403),
    czy nieudane odczytanie ceny NIE oznaczają, że oferta zniknęła - więc
    w takich przypadkach zostaje "DOSTĘPNE" z opisem problemu w kolumnie
    "Uwagi", zamiast fałszywie informować Cię, że auto zostało sprzedane."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    except requests.RequestException as e:
        return {"status": "DOSTĘPNE", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None,
                "note": f"Nie udało się połączyć ({e}) - spróbuję ponownie przy kolejnym sprawdzeniu"}

    if resp.status_code == 404:
        return {"status": "BRAK OGŁOSZENIA", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None, "note": "HTTP 404"}

    if resp.status_code == 403:
        return {"status": "DOSTĘPNE", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None,
                "note": ("Strona zablokowała zapytanie (HTTP 403) - często dotyczy serwerów "
                         "GitHub Actions, które mają współdzielone adresy IP. To NIE znaczy, "
                         "że oferta zniknęła, tylko że nie udało się jej teraz sprawdzić.")}

    if resp.status_code >= 400:
        return {"status": "DOSTĘPNE", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None,
                "note": f"Serwer odpowiedział błędem HTTP {resp.status_code} - spróbuję ponownie później"}

    lower_text = resp.text.lower()
    if any(phrase in lower_text for phrase in SOLD_PHRASES):
        return {"status": "BRAK OGŁOSZENIA", "price": None, "currency": None,
                "location": None, "title": None, "mileage": None, "image": None,
                "note": "wykryto frazę o nieaktualności ogłoszenia"}

    soup = BeautifulSoup(resp.text, "html.parser")
    site = detect_site(url)

    structured = extract_json_ld(soup)
    _, location, fallback_mileage = parse_price_fallback(soup, site)
    image = extract_image_fallback(soup)
    if structured:
        price = structured["price"]
        currency = structured["currency"]
        title = structured.get("title")
        mileage = structured.get("mileage") or fallback_mileage
        image = structured.get("image") or image
    else:
        price, location, mileage = parse_price_fallback(soup, site)
        currency = "PLN"
        title_el = soup.find("h1")
        title = title_el.get_text(strip=True) if title_el else None

    if not price:
        note = "Strona odpowiedziała, ale nie udało się automatycznie znaleźć ceny - wpisz ją ręcznie."
        if site == "facebook":
            note = ("Facebook Marketplace zwykle wymaga zalogowania, żeby zobaczyć szczegóły "
                    "ogłoszenia, więc automatyczne pobranie ceny często się tu nie uda - "
                    "wpisz cenę i przebieg ręcznie.")
        return {"status": "DOSTĘPNE", "price": None, "currency": None,
                "location": location, "title": title, "mileage": mileage, "image": image,
                "note": note}

    return {"status": "DOSTĘPNE", "price": price, "currency": currency,
            "location": location, "title": title, "mileage": mileage, "image": image, "note": ""}


def load_previous_data():
    """Wczytuje poprzedni stan z tracker.xlsx (jeśli istnieje), żeby wykryć zmiany cen."""
    if not OUTPUT_XLSX.exists():
        return {}
    wb = load_workbook(OUTPUT_XLSX)
    if "Oferty" not in wb.sheetnames:
        return {}
    ws = wb["Oferty"]
    prev = {}
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_dict = dict(zip(headers, row))
        if row_dict.get("URL"):
            prev[row_dict["URL"]] = row_dict
    return prev


def write_xlsx(results, history_entries):
    if OUTPUT_XLSX.exists():
        wb = load_workbook(OUTPUT_XLSX)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    # --- Arkusz "Oferty" (aktualny stan) ---
    if "Oferty" in wb.sheetnames:
        wb.remove(wb["Oferty"])
    ws = wb.create_sheet("Oferty", 0)

    columns = ["Nazwa", "URL", "Zdjęcie", "Status", "Cena", "Waluta", "Przebieg (km)", "Lokalizacja",
               "Ostatnia zmiana ceny", "Ostatnie sprawdzenie", "Uwagi"]
    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for r in results:
        ws.append([
            r["title"] or "", r["url"], r.get("image") or "", r["status"], r["price"] or "",
            r["currency"] or "", r.get("mileage") or "", r["location"] or "",
            r["price_change"], r["checked_at"], r["note"],
        ])
        row_idx = ws.max_row
        if r["status"] == "BRAK OGŁOSZENIA":
            fill = red
        elif r["status"] == "DOSTĘPNE" and not r["price"]:
            fill = yellow  # dostępne, ale nie udało się odczytać ceny - do ręcznej weryfikacji
        else:
            fill = green
        for col in range(1, len(columns) + 1):
            ws.cell(row=row_idx, column=col).fill = fill

    for i, col in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(col) + 4)

    # --- Arkusz "Historia cen" ---
    if "Historia cen" in wb.sheetnames:
        hist_ws = wb["Historia cen"]
    else:
        hist_ws = wb.create_sheet("Historia cen")
        hist_ws.append(["Data", "URL", "Poprzednia cena", "Nowa cena", "Zmiana"])
        for cell in hist_ws[1]:
            cell.font = Font(bold=True)

    for entry in history_entries:
        hist_ws.append(entry)

    wb.save(OUTPUT_XLSX)


def status_dot_html(status: str, has_price: bool) -> str:
    if status == "BRAK OGŁOSZENIA":
        return '<span class="dot dot-sold"></span>Brak ogłoszenia'
    if not has_price:
        return '<span class="dot dot-unknown"></span>Dostępne (sprawdź cenę ręcznie)'
    return '<span class="dot dot-active"></span>Dostępne'


def generate_html(results):
    """Tworzy index.html - interaktywną stronę publikowaną przez GitHub Pages.
    Dane pobiera skrypt uruchamiany przez GitHub Actions (działa bez Twojego
    komputera). Strona pozwala też dodawać/edytować/usuwać oferty wprost
    z przeglądarki - zmiany zapisują się do offers.csv w repozytorium przez
    GitHub API (wymaga wklejenia własnego tokenu dostępu, patrz panel
    Ustawienia na stronie)."""

    active_count = sum(1 for r in results if r["status"] == "DOSTĘPNE")
    sold_count = sum(1 for r in results if r["status"] == "BRAK OGŁOSZENIA")
    prices = [float(r["price"]) for r in results if r["price"]]
    avg_price = f'{sum(prices)/len(prices):,.0f} zł'.replace(",", " ") if prices else "—"

    rows_html = []
    for r in results:
        price_txt = f'{float(r["price"]):,.0f} zł'.replace(",", " ") if r["price"] else "—"
        mileage_txt = f'{int(float(r["mileage"])):,} km'.replace(",", " ") if r.get("mileage") else "—"
        change_txt = f'<span class="delta">{r["price_change"]}</span>' if r["price_change"] else ""
        row_class = "row-sold" if r["status"] == "BRAK OGŁOSZENIA" else ""
        status_key = "sold" if r["status"] == "BRAK OGŁOSZENIA" else "active"
        thumb_html = (
            f'<img class="thumb" src="{html.escape(r["image"])}" alt="" loading="lazy" '
            f'onerror="this.style.display=\'none\'">'
            if r.get("image") else '<div class="thumb thumb-empty">brak<br>zdjęcia</div>'
        )
        safe_url = html.escape(r["url"], quote=True)
        safe_name = html.escape(r.get("title") or "", quote=True)
        rows_html.append(f"""
        <tr class="{row_class}" data-status="{status_key}" data-url="{safe_url}" data-name="{safe_name}">
          <td>
            <div class="offer-cell">
              {thumb_html}
              <div>
                <div class="offer-name">{r['title'] or 'Bez nazwy'}</div>
                <a href="{r['url']}" target="_blank" rel="noopener">zobacz ogłoszenie ↗</a>
              </div>
            </div>
          </td>
          <td class="mono">{price_txt}{change_txt}</td>
          <td class="mono dim">{mileage_txt}</td>
          <td class="dim">{r['location'] or '—'}</td>
          <td>{status_dot_html(r['status'], bool(r['price']))}</td>
          <td class="dim">{r['checked_at']}</td>
          <td class="row-actions">
            <button class="icon-btn" title="Edytuj" onclick="startEdit(this)">✎</button>
            <button class="icon-btn icon-btn-danger" title="Usuń" onclick="deleteOffer(this)">✕</button>
          </td>
        </tr>""")

    html_doc = f"""<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monitor ofert aut</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
  :root{{
    --bg:#181b20; --panel:#1f2329; --panel-2:#262b32; --line:#333941;
    --text:#f1f3f5; --dim:#8b93a1; --dim-2:#5f6672;
    --amber:#f0a93b; --amber-dim:rgba(240,169,59,.14);
    --green:#5cc296; --green-bg:rgba(92,194,150,.12);
    --red:#e2694e; --red-bg:rgba(226,105,78,.12);
    --yellow:#d9b45a; --yellow-bg:rgba(217,180,90,.12);
    --radius:10px;
  }}
  *{{box-sizing:border-box;}}
  body{{
    margin:0; background:var(--bg); color:var(--text);
    font-family:'Inter',system-ui,sans-serif;
    padding:40px 24px 80px;
  }}
  .wrap{{max-width:1040px;margin:0 auto;}}

  header{{display:flex;justify-content:space-between;align-items:center;
    flex-wrap:wrap;gap:14px;margin-bottom:26px;}}
  .brand{{display:flex;align-items:center;gap:12px;}}
  .brand-mark{{width:38px;height:38px;border-radius:10px;background:linear-gradient(155deg,var(--amber),#c9822a);
    display:flex;align-items:center;justify-content:center;font-size:19px;flex-shrink:0;}}
  h1{{font-family:'Space Grotesk';font-weight:700;font-size:26px;margin:0;letter-spacing:0.2px;}}
  .updated{{color:var(--dim);font-size:12px;margin-top:2px;}}
  .updated b{{color:var(--amber);font-weight:600;}}

  .top-actions{{display:flex;gap:8px;}}

  button{{font-family:'Inter';cursor:pointer;border:none;}}
  .btn{{
    font-size:13px;font-weight:600;padding:9px 16px;border-radius:8px;
    background:var(--panel-2);color:var(--text);border:1px solid var(--line);
    transition:border-color .15s;
  }}
  .btn:hover{{border-color:var(--dim-2);}}
  .btn-primary{{background:var(--amber);color:#1b1e24;border:1px solid var(--amber);}}
  .btn-primary:hover{{opacity:.9;border-color:var(--amber);}}
  .btn-ghost{{background:transparent;}}

  .gauges{{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--line);
    border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;margin-bottom:24px;}}
  .gauge{{background:var(--panel);padding:18px 20px;}}
  .gauge-value{{font-family:'JetBrains Mono';font-size:25px;font-weight:700;color:var(--amber);line-height:1;}}
  .gauge-label{{font-size:11px;color:var(--dim);margin-top:6px;}}

  .card{{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);
    padding:20px;margin-bottom:20px;}}
  .card h2{{font-family:'Space Grotesk';font-size:15px;font-weight:600;margin:0 0 14px;color:var(--text);}}
  .card-hint{{font-size:12px;color:var(--dim);margin-top:10px;line-height:1.5;}}
  .card-hint a{{color:var(--amber);}}

  .form-row{{display:flex;gap:10px;flex-wrap:wrap;}}
  input[type=text], input[type=url], input[type=password]{{
    background:var(--bg);border:1px solid var(--line);color:var(--text);
    padding:10px 12px;border-radius:7px;font-size:13.5px;font-family:'Inter';
  }}
  input:focus{{outline:none;border-color:var(--amber);}}
  #offer-url{{flex:2;min-width:220px;}}
  #offer-name{{flex:1;min-width:150px;}}
  #settings-panel input{{width:100%;margin-bottom:10px;}}
  #settings-panel{{display:none;}}
  #settings-panel.open{{display:block;}}
  label{{font-size:12px;color:var(--dim);display:block;margin-bottom:5px;}}

  .status-msg{{font-size:12.5px;padding:9px 12px;border-radius:7px;margin-top:12px;display:none;}}
  .status-msg.show{{display:block;}}
  .status-msg.ok{{background:var(--green-bg);color:var(--green);}}
  .status-msg.err{{background:var(--red-bg);color:var(--red);}}
  .status-msg.info{{background:var(--amber-dim);color:var(--amber);}}

  .filters{{display:flex;gap:8px;margin-bottom:14px;}}
  .filter-btn{{
    background:transparent;border:1px solid var(--line);color:var(--dim);
    font-size:12.5px;font-weight:600;padding:6px 14px;border-radius:20px;
  }}
  .filter-btn.active{{background:var(--amber);border-color:var(--amber);color:#1b1e24;}}

  .table-wrap{{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;}}
  table{{width:100%;border-collapse:collapse;}}
  thead th{{text-align:left;font-size:11px;letter-spacing:.3px;color:var(--dim);
    font-weight:600;padding:12px 14px;border-bottom:1px solid var(--line);background:var(--panel-2);}}
  tbody tr{{border-bottom:1px solid var(--line);}}
  tbody tr:last-child{{border-bottom:none;}}
  tbody tr:hover{{background:var(--panel-2);}}
  tbody tr.row-sold{{opacity:.5;}}
  td{{padding:14px;font-size:13.5px;vertical-align:middle;}}
  .offer-cell{{display:flex;gap:12px;align-items:center;}}
  .thumb{{width:60px;height:45px;border-radius:6px;object-fit:cover;flex-shrink:0;background:var(--bg);border:1px solid var(--line);}}
  .thumb-empty{{display:flex;align-items:center;justify-content:center;font-size:9px;color:var(--dim-2);text-align:center;line-height:1.3;}}
  .offer-name{{font-weight:600;margin-bottom:2px;}}
  a{{color:var(--dim);font-size:11.5px;text-decoration:none;}}
  a:hover{{color:var(--amber);}}
  .mono{{font-family:'JetBrains Mono';font-size:14px;}}
  .dim{{color:var(--dim);font-size:12.5px;}}
  .delta{{font-size:11px;color:var(--amber);margin-left:7px;}}

  .dot{{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:7px;}}
  .dot-active{{background:var(--green);box-shadow:0 0 0 3px var(--green-bg);}}
  .dot-sold{{background:var(--red);box-shadow:0 0 0 3px var(--red-bg);}}
  .dot-unknown{{background:var(--yellow);box-shadow:0 0 0 3px var(--yellow-bg);}}

  .row-actions{{white-space:nowrap;text-align:right;}}
  .icon-btn{{background:transparent;color:var(--dim);border:1px solid var(--line);
    width:28px;height:28px;border-radius:6px;font-size:13px;margin-left:5px;}}
  .icon-btn:hover{{color:var(--text);border-color:var(--dim-2);}}
  .icon-btn-danger:hover{{color:var(--red);border-color:var(--red);}}

  .empty{{text-align:center;padding:60px 20px;color:var(--dim);}}
  footer{{margin-top:26px;color:var(--dim-2);font-size:11px;text-align:center;}}

  @media (max-width:640px){{
    .gauges{{grid-template-columns:repeat(2,1fr);}}
    thead{{display:none;}}
    table, tbody, tr, td{{display:block;width:100%;}}
    tbody tr{{padding:14px;}}
    td{{padding:4px 0;}}
    .row-actions{{text-align:left;margin-top:8px;}}
  }}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="brand">
      <div class="brand-mark">🚗</div>
      <div>
        <h1>Monitor ofert aut</h1>
        <div class="updated">Ostatnie sprawdzenie: <b>{datetime.now().strftime('%Y-%m-%d %H:%M')}</b> &middot; odświeża się co godzinę</div>
      </div>
    </div>
    <div class="top-actions">
      <button class="btn btn-ghost" onclick="toggleSettings()">⚙ Ustawienia</button>
    </div>
  </header>

  <div class="gauges">
    <div class="gauge"><div class="gauge-value">{len(results)}</div><div class="gauge-label">śledzonych ofert</div></div>
    <div class="gauge"><div class="gauge-value">{active_count}</div><div class="gauge-label">dostępnych</div></div>
    <div class="gauge"><div class="gauge-value">{sold_count}</div><div class="gauge-label">brak ogłoszenia</div></div>
    <div class="gauge"><div class="gauge-value">{avg_price}</div><div class="gauge-label">średnia cena aktywnych</div></div>
  </div>

  <div class="card" id="settings-panel">
    <h2>Ustawienia — token dostępu do GitHub</h2>
    <label>Personal Access Token (fine-grained, uprawnienia: Contents: Read&write, Actions: Read&write, tylko dla tego repo)</label>
    <input type="password" id="gh-token" placeholder="github_pat_...">
    <button class="btn btn-primary" onclick="saveToken()">Zapisz token w tej przeglądarce</button>
    <div class="card-hint">
      Token tworzysz na github.com → kliknij swój awatar → Settings → Developer settings →
      Personal access tokens → Fine-grained tokens → Generate new token. W "Repository access"
      wybierz tylko to repozytorium, a w "Permissions" ustaw Contents: Read and write oraz
      Actions: Read and write. Token zapisuje się tylko lokalnie w tej przeglądarce — nikomu
      go nie wysyłamy poza bezpośrednie zapytania do api.github.com.
      <a href="https://github.com/settings/personal-access-tokens/new" target="_blank">Utwórz token →</a>
    </div>
  </div>

  <div class="card">
    <h2 id="form-title">Dodaj ofertę</h2>
    <div class="form-row">
      <input type="url" id="offer-url" placeholder="Link do oferty (OLX / Otomoto)...">
      <input type="text" id="offer-name" placeholder="Nazwa (opcjonalnie)">
      <button class="btn btn-primary" id="submit-btn" onclick="submitOffer()">Dodaj</button>
      <button class="btn btn-ghost" id="cancel-edit-btn" style="display:none" onclick="cancelEdit()">Anuluj</button>
    </div>
    <div class="status-msg" id="status-msg"></div>
  </div>

  <div class="filters">
    <button class="filter-btn active" onclick="filterRows('all', this)">Wszystkie</button>
    <button class="filter-btn" onclick="filterRows('active', this)">Dostępne</button>
    <button class="filter-btn" onclick="filterRows('sold', this)">Brak ogłoszenia</button>
  </div>

  {"<div class='table-wrap'><table><thead><tr><th>Oferta</th><th>Cena</th><th>Przebieg</th><th>Lokalizacja</th><th>Status</th><th>Sprawdzono</th><th></th></tr></thead><tbody id='offers-body'>" + ''.join(rows_html) + "</tbody></table></div>" if results else '<div class="empty">Brak ofert. Dodaj pierwszy link powyżej.</div>'}

  <footer>Dane pobierane automatycznie przez GitHub Actions co godzinę, niezależnie od tego czy masz włączony komputer.</footer>
</div>

<script>
const OWNER = location.hostname.split('.')[0];
const REPO = location.pathname.split('/').filter(Boolean)[0] || '';
const CSV_PATH = 'offers.csv';
const WORKFLOW_FILE = 'check-offers.yml';
let editingUrl = null;

function getToken(){{ return localStorage.getItem('gh_pat') || ''; }}

function toggleSettings(){{
  document.getElementById('settings-panel').classList.toggle('open');
  const t = getToken();
  if(t) document.getElementById('gh-token').value = t;
}}

function saveToken(){{
  const t = document.getElementById('gh-token').value.trim();
  if(!t){{ return; }}
  localStorage.setItem('gh_pat', t);
  showStatus('Token zapisany w tej przeglądarce.', 'ok');
  document.getElementById('settings-panel').classList.remove('open');
}}

function showStatus(msg, kind){{
  const el = document.getElementById('status-msg');
  el.textContent = msg;
  el.className = 'status-msg show ' + kind;
}}

function filterRows(status, btn){{
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('tbody tr').forEach(row => {{
    row.style.display = (status === 'all' || row.dataset.status === status) ? '' : 'none';
  }});
}}

function parseCSV(text){{
  const lines = text.replace(/\\r\\n/g, '\\n').split('\\n').filter(l => l.length > 0);
  if(lines.length === 0) return [];
  const rows = lines.slice(1);
  return rows.map(line => {{
    const fields = [];
    let cur = '', inQuotes = false;
    for(let i = 0; i < line.length; i++){{
      const c = line[i];
      if(inQuotes){{
        if(c === '"' && line[i+1] === '"'){{ cur += '"'; i++; }}
        else if(c === '"'){{ inQuotes = false; }}
        else {{ cur += c; }}
      }} else {{
        if(c === '"'){{ inQuotes = true; }}
        else if(c === ','){{ fields.push(cur); cur = ''; }}
        else {{ cur += c; }}
      }}
    }}
    fields.push(cur);
    return {{ name: fields[0] || '', url: fields[1] || '' }};
  }});
}}

function csvEscape(field){{
  if(field.includes(',') || field.includes('"') || field.includes('\\n')){{
    return '"' + field.replace(/"/g, '""') + '"';
  }}
  return field;
}}

function stringifyCSV(rows){{
  const lines = ['nazwa,url'];
  rows.forEach(r => lines.push(csvEscape(r.name) + ',' + csvEscape(r.url)));
  return lines.join('\\n') + '\\n';
}}

function utf8ToBase64(str){{
  return btoa(unescape(encodeURIComponent(str)));
}}
function base64ToUtf8(str){{
  return decodeURIComponent(escape(atob(str)));
}}

async function ghRequest(path, options){{
  const token = getToken();
  if(!token){{
    showStatus('Najpierw ustaw token w Ustawieniach (⚙ w prawym górnym rogu).', 'err');
    document.getElementById('settings-panel').classList.add('open');
    throw new Error('no token');
  }}
  const resp = await fetch(`https://api.github.com/repos/${{OWNER}}/${{REPO}}${{path}}`, {{
    ...options,
    headers: {{
      'Authorization': `Bearer ${{token}}`,
      'Accept': 'application/vnd.github+json',
      ...(options && options.headers ? options.headers : {{}})
    }}
  }});
  if(!resp.ok){{
    const body = await resp.text();
    throw new Error(`GitHub API ${{resp.status}}: ${{body.slice(0,200)}}`);
  }}
  return resp.status === 204 ? null : resp.json();
}}

async function getOffersFile(){{
  const data = await ghRequest(`/contents/${{CSV_PATH}}`, {{ method: 'GET' }});
  return {{ rows: parseCSV(base64ToUtf8(data.content)), sha: data.sha }};
}}

async function saveOffersFile(rows, sha, message){{
  await ghRequest(`/contents/${{CSV_PATH}}`, {{
    method: 'PUT',
    body: JSON.stringify({{
      message,
      content: utf8ToBase64(stringifyCSV(rows)),
      sha,
      branch: 'main'
    }})
  }});
}}

async function triggerWorkflow(){{
  try{{
    await ghRequest(`/actions/workflows/${{WORKFLOW_FILE}}/dispatches`, {{
      method: 'POST',
      body: JSON.stringify({{ ref: 'main' }})
    }});
    return {{ ok: true }};
  }}catch(e){{
    return {{ ok: false, error: e.message }};
  }}
}}

function startEdit(btn){{
  const row = btn.closest('tr');
  editingUrl = row.dataset.url;
  document.getElementById('offer-url').value = editingUrl;
  document.getElementById('offer-name').value = row.dataset.name;
  document.getElementById('form-title').textContent = 'Edytuj ofertę';
  document.getElementById('submit-btn').textContent = 'Zapisz zmiany';
  document.getElementById('cancel-edit-btn').style.display = '';
  window.scrollTo({{top:0, behavior:'smooth'}});
}}

function cancelEdit(){{
  editingUrl = null;
  document.getElementById('offer-url').value = '';
  document.getElementById('offer-name').value = '';
  document.getElementById('form-title').textContent = 'Dodaj ofertę';
  document.getElementById('submit-btn').textContent = 'Dodaj';
  document.getElementById('cancel-edit-btn').style.display = 'none';
}}

async function submitOffer(){{
  const url = document.getElementById('offer-url').value.trim();
  const name = document.getElementById('offer-name').value.trim();
  if(!url){{ showStatus('Wklej najpierw link do oferty.', 'err'); return; }}

  const btn = document.getElementById('submit-btn');
  btn.disabled = true;
  showStatus('Zapisuję do repozytorium...', 'info');
  try{{
    const {{ rows, sha }} = await getOffersFile();
    let newRows;
    if(editingUrl){{
      newRows = rows.map(r => r.url === editingUrl ? {{ name: name || r.name, url }} : r);
    }} else {{
      if(rows.some(r => r.url === url)){{
        showStatus('Ta oferta jest już na liście.', 'err');
        btn.disabled = false;
        return;
      }}
      newRows = [...rows, {{ name, url }}];
    }}
    await saveOffersFile(newRows, sha, editingUrl ? 'Edycja oferty ze strony' : 'Dodanie oferty ze strony');
    const trigger = await triggerWorkflow();
    if(trigger.ok){{
      showStatus('Zapisano! Sprawdzanie ruszyło - odśwież tę stronę (Ctrl+Shift+R) za ok. 1-2 minuty.', 'ok');
    }} else {{
      showStatus('Zapisano w repozytorium, ALE nie udało się automatycznie uruchomić sprawdzania (' + trigger.error + '). Sprawdź czy token ma uprawnienie "Actions: Read and write", albo odpal ręcznie w zakładce Actions → Run workflow. W najgorszym razie zadziała samo o pełnej godzinie.', 'err');
    }}
    cancelEdit();
  }}catch(e){{
    showStatus('Błąd: ' + e.message, 'err');
  }}
  btn.disabled = false;
}}

async function deleteOffer(btn){{
  const row = btn.closest('tr');
  const url = row.dataset.url;
  if(!confirm('Usunąć tę ofertę z listy?')) return;
  btn.disabled = true;
  showStatus('Usuwam...', 'info');
  try{{
    const {{ rows, sha }} = await getOffersFile();
    const newRows = rows.filter(r => r.url !== url);
    await saveOffersFile(newRows, sha, 'Usunięcie oferty ze strony');
    const trigger = await triggerWorkflow();
    if(trigger.ok){{
      showStatus('Usunięto! Odśwież stronę (Ctrl+Shift+R) za ok. 1-2 minuty, żeby zobaczyć zaktualizowaną listę.', 'ok');
    }} else {{
      showStatus('Usunięto z repozytorium, ALE nie udało się automatycznie uruchomić odświeżenia strony (' + trigger.error + '). Odpal ręcznie w zakładce Actions → Run workflow, albo poczekaj do pełnej godziny.', 'err');
    }}
    row.remove();
  }}catch(e){{
    showStatus('Błąd: ' + e.message, 'err');
  }}
}}
</script>
</body>
</html>"""

    out_path = BASE_DIR / "index.html"
    out_path.write_text(html_doc, encoding="utf-8")
    return out_path


def main():
    offers = load_offers()
    prev_data = load_previous_data()
    results = []
    history_entries = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    for offer in offers:
        url = offer["url"].strip()
        if not url:
            continue
        log(f"Sprawdzam: {url}")
        info = check_offer(url)

        prev = prev_data.get(url)
        price_change = ""
        if prev and info["price"]:
            try:
                old_price = float(str(prev.get("Cena", "")).replace(" ", "").replace(",", "."))
                new_price = float(str(info["price"]).replace(" ", "").replace(",", "."))
                if old_price and new_price and old_price != new_price:
                    diff = new_price - old_price
                    price_change = f"{'+' if diff > 0 else ''}{diff:.0f} zł"
                    history_entries.append([now, url, old_price, new_price, price_change])
                    log(f"  Zmiana ceny: {old_price} -> {new_price} ({price_change})")
            except (ValueError, TypeError):
                pass

        title = info["title"] or (offer.get("nazwa") or "")
        results.append({
            "title": title,
            "url": url,
            "status": info["status"],
            "price": info["price"],
            "currency": info["currency"],
            "location": info["location"],
            "mileage": info.get("mileage"),
            "image": info.get("image"),
            "price_change": price_change,
            "checked_at": now,
            "note": info["note"],
        })

    write_xlsx(results, history_entries)
    dashboard_path = generate_html(results)
    log(f"Gotowe. Zapisano {len(results)} ofert do {OUTPUT_XLSX.name} i {dashboard_path.name}")


if __name__ == "__main__":
    main()
