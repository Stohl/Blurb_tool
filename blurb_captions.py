#!/usr/bin/env python3
"""
Reads EXIF from images inside a .blurb and fills text boxes with captions.

Format: Location · weekday date · time · temp°C symbol · age
        Caption from EXIF ImageDescription

Text boxes containing ### are protected (not modified). Other text boxes that do
not receive new text (for example more boxes than images, or no image on a page)
are cleared automatically.

Workflow:
  python3 blurb_captions.py file.blurb

Writes captions to file-new.blurb and does not modify the original .blurb.
CSV: file.csv (same name as the book without a trailing -new on the stem).

Works on .blurb – no extract/pack needed.

Requires: Pillow.
"""

import csv
import html
import io
import json
import re
import shutil
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional

from PIL import Image

# ── User toggles (set True/False here) ───────────────────────────────────────
RESIZE_LAYOUT = True    # True = run internal resize of text/image containers in bbf2.xml
ENABLE_GPS = True       # If False: no place name in captions; lat/lon still in CSV / used for weather when on
ENABLE_WEATHER = True   # If False: do not fetch/write weather

# Reference date for age / days until birthday in captions (see format_age). None = disable.
BIRTH_DATE = date(2023, 5, 14)

# ── Language  ─────────────────────────────────────────────────────────
# Set LANGUAGE to one of the codes defined in _LANGUAGES below.
# To add a new language: copy one of the existing blocks, give it a new 2-letter
# code (e.g. "DE", "FR", "NO"), translate the strings, and set LANGUAGE to that code.
LANGUAGE = "EN"  # "EN" = English, "SV" = Svenska, 
SHORT_TEXT = False  # True = abbreviate weekday/month/age text

_LANGUAGES: dict[str, dict] = {
    "EN": {
        "weekdays":       ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"),
        "weekdays_short": ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"),
        "months":         ("January", "February", "March", "April", "May", "June",
                           "July", "August", "September", "October", "November", "December"),
        "months_short":   ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
        "day":            "day",
        "days":           "days",
        "week":           "week",
        "weeks":          "weeks",
        "month":          "month",
        "months_age":     "months",
        "year":           "year",
        "years":          "years",
        "birthday":       "birthday",
        "remaining":      "remaining",
        "abbr": [
            ("weeks",  "wk"),
            ("week",   "wk"),
            ("months", "mo"),
            ("month",  "mo"),
            ("days",   "d"),
            ("years",   "y"),
            ("year",   "y"),
        ],
    },
        "SV": {
        "weekdays":       ("måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag", "söndag"),
        "weekdays_short": ("mån", "tis", "ons", "tor", "fre", "lör", "sön"),
        "months":         ("januari", "februari", "mars", "april", "maj", "juni",
                           "juli", "augusti", "september", "oktober", "november", "december"),
        "months_short":   ("jan", "feb", "mar", "apr", "maj", "jun",
                           "jul", "aug", "sep", "okt", "nov", "dec"),
        # Age strings – singular and plural where needed
        "day":            "dag",
        "days":           "dagar",
        "week":           "vecka",
        "weeks":          "veckor",
        "month":          "månad",
        "months_age":     "månader",
        "year":           "år",
        "years":          "år",
        # Special age labels
        "birthday":       "födelsedagen",
        "remaining":      "kvar",
        # Short replacements applied by _shorty()
        # Each entry is (full_word, short_form)
        "abbr": [
            ("veckor",   "v"),
            ("vecka",    "v"),
            ("månader",  "mån"),
            ("månad",    "mån"),
            ("dagar",    "d"),
        ],
    },
    # ── Add your language here ────────────────────────────────────────────────
    # "DE": {
    #     "weekdays":       ("Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"),
    #     "weekdays_short": ("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"),
    #     "months":         ("Januar", "Februar", "März", "April", "Mai", "Juni",
    #                        "Juli", "August", "September", "Oktober", "November", "Dezember"),
    #     "months_short":   ("Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
    #                        "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"),
    #     "day":            "Tag",    "days":       "Tage",
    #     "week":           "Woche",  "weeks":      "Wochen",
    #     "month":          "Monat",  "months_age": "Monate",
    #     "year":           "Jahr",   "years":      "Jahre",
    #     "birthday":       "Geburtstag",
    #     "remaining":      "noch",
    #     "abbr": [
    #         ("Wochen", "Wo"), ("Woche", "Wo"),
    #         ("Monate", "Mo"), ("Monat", "Mo"),
    #         ("Tage",   "T"),
    #     ],
    # },
}

# Resolve active language – fall back to English if code is unknown
_LANG = _LANGUAGES.get(LANGUAGE, _LANGUAGES["EN"])


@dataclass
class ExifData:
    """EXIF data from one image."""
    datetime_str: str = ""
    caption: str = ""
    gps: Optional[tuple[float, float]] = None


def get_exif_from_blob(blob: bytes) -> ExifData:
    """Read EXIF from an image blob. Returns ExifData."""
    try:
        img = Image.open(io.BytesIO(blob))
        exif = img.getexif() if img else None
        if not exif:
            return ExifData()

        def _str(val) -> str:
            if val is None:
                return ""
            if isinstance(val, bytes):
                return val.decode("utf-8", errors="replace").strip("\x00")
            return str(val).strip()

        dt = _str(exif.get(36867)) or _str(exif.get(306))  # DateTimeOriginal, DateTime
        caption = _str(exif.get(270))  # ImageDescription

        gps = None
        try:
            gps_ifd = exif.get_ifd(0x8825)  # GPS IFD
            if gps_ifd:
                lat = gps_ifd.get(2)  # GPSLatitude
                lon = gps_ifd.get(4)  # GPSLongitude
                lat_ref = gps_ifd.get(1, "N")
                lon_ref = gps_ifd.get(3, "E")
                if lat and lon:
                    def to_deg(v):
                        d, m, s = v
                        return float(d) + float(m) / 60 + float(s) / 3600
                    lat_val = to_deg(lat) * (-1 if str(lat_ref) == "S" else 1)
                    lon_val = to_deg(lon) * (-1 if str(lon_ref) == "W" else 1)
                    gps = (lat_val, lon_val)
        except Exception:
            pass

        return ExifData(datetime_str=dt, caption=caption, gps=gps)
    except Exception:
        return ExifData()


def _fix_mojibake(s: str) -> str:
    """Fix UTF-8 text that was mis-decoded as Latin-1."""
    if not s or "Ã" not in s:
        return s
    try:
        return s.encode("latin-1").decode("utf-8")
    except Exception:
        return s


# These are resolved from _LANG so that all language logic flows through _LANGUAGES above.
_SV_WEEKDAYS      = _LANG["weekdays"]
_SV_MONTHS        = _LANG["months"]
_SV_WEEKDAYS_KORT = _LANG["weekdays_short"]
_SV_MONTHS_KORT   = _LANG["months_short"]

_WMO_SYMBOLS = {
    0: "☀︎", 1: "☀︎", 2: "⛅", 3: "☁︎", 45: "🌫", 48: "🌫",
    51: "🌧", 53: "🌧", 55: "🌧", 56: "🌧", 57: "🌧",
    61: "🌧", 63: "🌧", 65: "🌧", 66: "🌧", 67: "🌧",
    71: "❄︎", 73: "❄︎", 75: "❄︎", 77: "❄︎",
    80: "🌦", 81: "🌦", 82: "🌦", 85: "🌨", 86: "🌨",
    95: "⛈", 96: "⛈", 99: "⛈",
}


def _shorty(s: str) -> str:
    """Abbreviates weekday/month/age words when SHORT_TEXT is True."""
    if not SHORT_TEXT or not s:
        return s
    result = s
    # Weekdays
    for full, kort in zip(_SV_WEEKDAYS, _SV_WEEKDAYS_KORT):
        result = result.replace(full, kort)
    # Months (longer first to avoid substring collisions)
    for full, kort in zip(_SV_MONTHS, _SV_MONTHS_KORT):
        result = result.replace(full, kort)
    # Age words – defined per language in _LANGUAGES[...]["abbr"]
    for full, kort in _LANG["abbr"]:
        result = result.replace(full, kort)
    return result


def format_date_long(dt_str: str) -> tuple[str, str]:
    """Returns (weekday day month year, HH:MM)."""
    s = (dt_str or "").strip().replace(":", "-", 2)
    parts = s.split()
    date_part = parts[0] if parts else ""
    time_part = parts[1][:5] if len(parts) > 1 and len(parts[1]) >= 5 else ""
    dparts = date_part.split("-")
    if len(dparts) >= 3:
        try:
            y, m, d = int(dparts[0]), int(dparts[1]), int(dparts[2])
            if 1 <= m <= 12:
                dt = datetime(y, m, d)
                weekdays = _SV_WEEKDAYS_KORT if SHORT_TEXT else _SV_WEEKDAYS
                months = _SV_MONTHS_KORT if SHORT_TEXT else _SV_MONTHS
                weekday = weekdays[dt.weekday()]
                month = months[m - 1]
                return (f"{weekday} {d} {month} {y}", time_part)
        except (ValueError, IndexError):
            pass
    return ("", time_part)


def _months_between(earlier: date, later: date) -> int:
    """Whole calendar months from earlier to later (later >= earlier)."""
    mo = (later.year - earlier.year) * 12 + (later.month - earlier.month)
    if later.day < earlier.day:
        mo -= 1
    return max(0, mo)


def _years_between(earlier: date, later: date) -> int:
    """Whole calendar years from earlier to later (later >= earlier)."""
    y = later.year - earlier.year
    if (later.month, later.day) < (earlier.month, earlier.day):
        y -= 1
    return max(0, y)


def _sv_dagar(n: int) -> str:
    return f"1 {_LANG['day']}" if n == 1 else f"{n} {_LANG['days']}"


def _sv_veckor(delta_days: int) -> str:
    w, r = divmod(delta_days, 7)
    if w == 1 and r == 0:
        return f"1 {_LANG['week']}"
    if r == 0:
        return f"{w} {_LANG['weeks']}"
    if w == 0:
        return _sv_dagar(r)
    w_word = _LANG["week"] if w == 1 else _LANG["weeks"]
    return f"{w} {w_word} {_sv_dagar(r)}"


def _sv_månader(n: int) -> str:
    return f"1 {_LANG['month']}" if n == 1 else f"{n} {_LANG['months_age']}"


def format_age(dt_str: str) -> str:
    """Age or time until birth relative to BIRTH_DATE.

    Tiers: <14 days → days; <8 weeks → weeks; <24 months → months; ≥2 years → years.
    """
    if BIRTH_DATE is None:
        return ""
    s = (dt_str or "").strip().replace(":", "-", 2)
    parts = s.split()
    date_part = parts[0] if parts else ""
    dparts = date_part.split("-")
    if len(dparts) < 3:
        return ""
    try:
        y, m, d = int(dparts[0]), int(dparts[1]), int(dparts[2])
        photo = date(y, m, d)
    except (ValueError, IndexError):
        return ""
    delta = (photo - BIRTH_DATE).days

    if delta == 0:
        return _LANG["birthday"]

    if delta < 0:
        # Photo before birth: countdown to BIRTH_DATE
        du = abs(delta)
        if du < 14:
            return f"{_sv_dagar(du)} {_LANG['remaining']}"
        if du < 56:
            return f"{_sv_veckor(du)} {_LANG['remaining']}"
        mu = _months_between(photo, BIRTH_DATE)
        if mu < 24:
            return f"{_sv_månader(mu)} {_LANG['remaining']}"
        yu = _years_between(photo, BIRTH_DATE)
        y_word = _LANG["year"] if yu == 1 else _LANG["years"]
        return f"{yu} {y_word} {_LANG['remaining']}"

    # Age after birth
    if delta < 14:
        return _sv_dagar(delta)
    if delta < 56:
        return _sv_veckor(delta)
    months_age = _months_between(BIRTH_DATE, photo)
    years_age = _years_between(BIRTH_DATE, photo)
    if years_age >= 2:
        return f"{years_age} {_LANG['years']}"
    return _sv_månader(months_age)


def _get_city_from_coords(lat: float, lon: float) -> str:
    """Reverse geocoding via Nominatim. Returns empty string on failure."""
    try:
        from geopy.geocoders import Nominatim
        from geopy.exc import GeocoderTimedOut, GeocoderServiceError
    except ImportError:
        return ""
    try:
        geolocator = Nominatim(user_agent="blurb-exif")
        location = geolocator.reverse(f"{lat}, {lon}", timeout=10, addressdetails=True)
        if location is None or not location.raw:
            return ""
        address = location.raw.get("address", {})
        for key in ("city", "town", "village", "suburb", "hamlet"):
            val = address.get(key, "").strip()
            if val:
                return val
        mun = address.get("municipality", "").strip()
        if mun:
            return mun.replace(" kommun", "").replace(" Kommun", "").strip()
        return ""
    except (GeocoderTimedOut, GeocoderServiceError, Exception):
        return ""


def _get_weather_from_openmeteo(
    lat: float, lon: float, date_str: str, hour: int
) -> tuple[Optional[float], Optional[str]]:
    """Fetch historical weather data from Open-Meteo."""
    try:
        url = (
            "https://archive-api.open-meteo.com/v1/archive"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={date_str}&end_date={date_str}"
            "&hourly=temperature_2m,weather_code"
            "&timezone=Europe/Stockholm"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "blurb-exif/1.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode())
        hourly = data.get("hourly", {})
        times = hourly.get("time", [])
        temps = hourly.get("temperature_2m", [])
        codes = hourly.get("weather_code", [])
        target = f"{date_str}T{hour:02d}:00"
        for i, t in enumerate(times):
            if t.startswith(target) or t == target:
                temp = temps[i] if i < len(temps) else None
                code = int(codes[i]) if i < len(codes) else 0
                symbol = _WMO_SYMBOLS.get(code, "☁︎")
                return (temp, symbol)
        return (None, None)
    except Exception:
        return (None, None)


def _parse_lat_lon(lat_s: str, lon_s: str) -> Optional[tuple[float, float]]:
    """Parse two coordinate strings. Both must be present and numeric."""
    if not (lat_s or "").strip() or not (lon_s or "").strip():
        return None
    try:
        return (
            float(lat_s.strip().replace(",", ".")),
            float(lon_s.strip().replace(",", ".")),
        )
    except ValueError:
        return None


def _gps_from_csv_or_exif(
    csv_lat: str, csv_lon: str, exif_gps: Optional[tuple[float, float]]
) -> Optional[tuple[float, float]]:
    """Prefer CSV coordinates when both columns are set (manual override); else EXIF."""
    parsed = _parse_lat_lon(csv_lat, csv_lon)
    if parsed is not None:
        return parsed
    return exif_gps


def get_plats_väder(
    img_filename: str,
    gps: Optional[tuple[float, float]],
    dt_str: str,
    city_map: dict[str, str],
    weather_map: dict[str, str],
) -> tuple[str, str]:
    """Get location/weather from cache, otherwise via geocoding + Open-Meteo."""
    plats = city_map.get(img_filename, "") if ENABLE_GPS else ""
    väder = weather_map.get(img_filename, "") if ENABLE_WEATHER else ""

    if (not ENABLE_GPS) and (not ENABLE_WEATHER):
        return ("", "")

    if gps and (ENABLE_GPS or ENABLE_WEATHER) and (not plats or not väder):
        lat, lon = gps
        if ENABLE_GPS and not plats:
            plats = _get_city_from_coords(lat, lon)
            if plats:
                city_map[img_filename] = plats
            time.sleep(1.1)
        if ENABLE_WEATHER and not väder:
            parts = (dt_str or "").replace(":", "-", 2).split()
            date_part = parts[0] if parts else ""
            hour = 12
            if len(parts) > 1:
                try:
                    hour = int(parts[1].split(":")[0])
                except ValueError:
                    pass
            if date_part and len(date_part.split("-")) == 3:
                temp, symbol = _get_weather_from_openmeteo(lat, lon, date_part, hour)
                if temp is not None and symbol:
                    väder = f"{int(round(temp))}°C {symbol}"
                    weather_map[img_filename] = väder
                time.sleep(0.2)
    return (plats, väder)


# Resize settings
RESIZE_PAGE_HEIGHT = 594.0
RESIZE_TOP_MARGIN = 30.0
RESIZE_BOTTOM_MARGIN = 30.0
RESIZE_GAP = 4.0
RESIZE_EVEN_LEFT_MARGIN = 30.0
RESIZE_EVEN_RIGHT_MARGIN = 50.0
RESIZE_ODD_LEFT_MARGIN = 50.0
RESIZE_ODD_RIGHT_MARGIN = 30.0
RESIZE_FONT_SIZE = 12.0
RESIZE_LINE_HEIGHT = RESIZE_FONT_SIZE * 1.28
RESIZE_PADDING_V = 1.5
RESIZE_MIN_TEXT_HEIGHT = RESIZE_LINE_HEIGHT + RESIZE_PADDING_V * 2
RESIZE_COL_TOLERANCE = 8.0
RESIZE_CHAR_WIDTH_FACTOR = 0.50
RESIZE_SHORT_TEXT_FUDGE = 1.08
RESIZE_LONG_TEXT_FUDGE = 1.02
RESIZE_FUDGE_LINE_CUTOFF = 3


class _ResizeTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.paragraphs = []
        self._current = []

    def handle_starttag(self, tag, attrs):
        if tag == "p":
            self._current = []

    def handle_endtag(self, tag):
        if tag == "p":
            self.paragraphs.append("".join(self._current))
            self._current = []

    def handle_data(self, data):
        self._current.append(data)


def _resize_calc_text_height(cdata_content: str, container_width: float) -> float:
    html_content = cdata_content.strip()
    if not html_content:
        return RESIZE_MIN_TEXT_HEIGHT

    parser = _ResizeTextExtractor()
    try:
        parser.feed(html_content)
    except Exception:
        pass

    paragraphs = parser.paragraphs
    if not paragraphs:
        return RESIZE_MIN_TEXT_HEIGHT

    norm = []
    for p in paragraphs:
        t = (p or "").replace("\xa0", " ").replace("\u200b", "").strip()
        t = re.sub(r"\s+", " ", t)
        norm.append(t)
    while norm and not norm[-1]:
        norm.pop()
    if not norm:
        return RESIZE_MIN_TEXT_HEIGHT

    char_w = RESIZE_FONT_SIZE * RESIZE_CHAR_WIDTH_FACTOR
    chars_per_line = max(1, int(container_width / char_w))
    total_lines = 0
    for text in norm:
        if not text:
            total_lines += 1
            continue
        words = text.split()
        line_len = 0
        lines = 1
        for word in words:
            w = len(word) + 1
            if line_len + w > chars_per_line and line_len > 0:
                lines += 1
                line_len = w
            else:
                line_len += w
        total_lines += lines

    height = total_lines * RESIZE_LINE_HEIGHT + RESIZE_PADDING_V * 2
    height *= (
        RESIZE_SHORT_TEXT_FUDGE
        if total_lines <= RESIZE_FUDGE_LINE_CUTOFF
        else RESIZE_LONG_TEXT_FUDGE
    )
    return max(height, RESIZE_MIN_TEXT_HEIGHT)


def _resize_get_cdata(text_elem) -> str:
    return (text_elem.text or "").strip()


def _resize_group_by_column(containers):
    groups = []
    for c in containers:
        x = float(c.get("x", 0))
        matched = False
        for group in groups:
            if abs(x - group[0]) <= RESIZE_COL_TOLERANCE:
                group[1].append(c)
                group[0] = sum(float(cc.get("x", 0)) for cc in group[1]) / len(group[1])
                matched = True
                break
        if not matched:
            groups.append([x, [c]])
    groups.sort(key=lambda g: g[0])
    return [(g[0], g[1]) for g in groups]


def _resize_col_bbox(containers):
    xs = [float(c.get("x", 0)) for c in containers]
    xe = [float(c.get("x", 0)) + float(c.get("width", 0)) for c in containers]
    return (min(xs), max(xe))


def _resize_scale_column(containers, old_xmin, old_width, new_xmin, new_width):
    if old_width <= 0:
        return
    s = new_width / old_width
    for c in containers:
        ox = float(c.get("x", 0))
        ow = float(c.get("width", 0))
        nx = new_xmin + (ox - old_xmin) * s
        nw = ow * s
        c.set("x", str(nx))
        c.set("width", str(nw))


def _resize_side_margins_for_page(page_number):
    try:
        n = int(page_number)
    except Exception:
        return RESIZE_EVEN_LEFT_MARGIN, RESIZE_EVEN_RIGHT_MARGIN
    if n % 2 == 0:
        return RESIZE_EVEN_LEFT_MARGIN, RESIZE_EVEN_RIGHT_MARGIN
    return RESIZE_ODD_LEFT_MARGIN, RESIZE_ODD_RIGHT_MARGIN


def _resize_process_page(page_elem, page_width):
    containers = page_elem.findall("container")
    if not containers:
        return

    columns = _resize_group_by_column(containers)

    standard_cols = []
    for x_center, col_containers in columns:
        images = [c for c in col_containers if c.get("type") == "image"]
        texts = [c for c in col_containers if c.get("type") == "text"]
        if len(images) == 2 and len(texts) == 2:
            standard_cols.append((x_center, images + texts))

    if len(standard_cols) == 2 and page_width:
        standard_cols.sort(key=lambda t: t[0])
        (_, left_ctrs), (_, right_ctrs) = standard_cols
        left_margin, right_margin = _resize_side_margins_for_page(page_elem.get("number"))
        l_xmin, l_xmax = _resize_col_bbox(left_ctrs)
        r_xmin, r_xmax = _resize_col_bbox(right_ctrs)
        w_l = l_xmax - l_xmin
        w_r = r_xmax - r_xmin
        w_sum = w_l + w_r
        w_avail = float(page_width) - left_margin - right_margin - RESIZE_GAP

        if w_l > 0 and w_r > 0 and w_avail > 0 and w_sum > 0:
            new_w_l = w_avail * (w_l / w_sum)
            new_w_r = w_avail - new_w_l
            new_l_xmin = left_margin
            new_r_xmin = left_margin + new_w_l + RESIZE_GAP
            _resize_scale_column(left_ctrs, l_xmin, w_l, new_l_xmin, new_w_l)
            _resize_scale_column(right_ctrs, r_xmin, w_r, new_r_xmin, new_w_r)
            columns = _resize_group_by_column(containers)

    for _, col_containers in columns:
        images = [c for c in col_containers if c.get("type") == "image"]
        texts = [c for c in col_containers if c.get("type") == "text"]
        if len(images) != 2 or len(texts) != 2:
            continue

        images.sort(key=lambda c: float(c.get("y", 0)))
        texts.sort(key=lambda c: float(c.get("y", 0)))
        width = float(images[0].get("width", 268))

        text_data = []
        for t in texts:
            te = t.find("text")
            cdata = _resize_get_cdata(te) if te is not None else ""
            h = _resize_calc_text_height(cdata, width)
            text_data.append({"h": h})

        h_text1 = text_data[0]["h"]
        h_text2 = text_data[1]["h"]
        available = (
            RESIZE_PAGE_HEIGHT
            - RESIZE_TOP_MARGIN
            - RESIZE_BOTTOM_MARGIN
            - RESIZE_GAP * 3
            - h_text1
            - h_text2
        )
        h_img = max(available / 2, 10)

        y = RESIZE_TOP_MARGIN
        images[0].set("y", str(y))
        images[0].set("height", str(h_img))
        y += h_img + RESIZE_GAP

        texts[0].set("y", str(y))
        texts[0].set("height", str(h_text1))
        y += h_text1 + RESIZE_GAP

        images[1].set("y", str(y))
        images[1].set("height", str(h_img))
        y += h_img + RESIZE_GAP

        texts[1].set("y", str(y))
        texts[1].set("height", str(h_text2))


_ATTR_FLOAT_RE = r"[-+]?\d+(?:\.\d+)?"


def _fmt_num(x: float) -> str:
    """
    Format floats in a stable, compact way without exponent.
    Prefer integers when the value is very close to an int.
    """
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    s = f"{x:.6f}".rstrip("0").rstrip(".")
    return s if s else "0"


def _find_container_opening_by_id(
    content: str, container_id: str, start_pos: int = 0
) -> Optional[tuple[int, int]]:
    """
    Return (tag_start, tag_end) for the opening '<container ...>' tag whose id= matches.
    tag_end is an exclusive index (one past '>').
    """
    needle = f'id="{container_id}"'
    pos = content.find(needle, start_pos)
    if pos < 0:
        return None

    tag_start = content.rfind("<container", 0, pos)
    if tag_start < 0:
        return None

    tag_end = content.find(">", tag_start)
    if tag_end < 0:
        return None

    opening = content[tag_start : tag_end + 1]
    if needle not in opening:
        return _find_container_opening_by_id(content, container_id, start_pos=pos + len(needle))

    return (tag_start, tag_end + 1)


def _set_attr_in_opening_tag(opening: str, attr: str, value: str) -> str:
    """
    Update or insert attr="value" in a <container ...> opening tag.
    Only touches that single attribute; preserves the rest verbatim.
    """
    pat = re.compile(rf'(\s{re.escape(attr)}="){_ATTR_FLOAT_RE}(")')
    if pat.search(opening):
        return pat.sub(rf"\g<1>{value}\2", opening, count=1)

    if opening.endswith("/>"):
        return opening[:-2] + f' {attr}="{value}"' + "/>"
    if opening.endswith(">"):
        return opening[:-1] + f' {attr}="{value}"' + ">"
    return opening


def patch_container_attrs_by_id(
    content: str,
    container_id: str,
    *,
    x: Optional[float] = None,
    y: Optional[float] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
) -> str:
    loc = _find_container_opening_by_id(content, container_id)
    if not loc:
        return content

    tag_start, tag_end = loc
    opening = content[tag_start:tag_end]

    if x is not None:
        opening = _set_attr_in_opening_tag(opening, "x", _fmt_num(x))
    if y is not None:
        opening = _set_attr_in_opening_tag(opening, "y", _fmt_num(y))
    if width is not None:
        opening = _set_attr_in_opening_tag(opening, "width", _fmt_num(width))
    if height is not None:
        opening = _set_attr_in_opening_tag(opening, "height", _fmt_num(height))

    return content[:tag_start] + opening + content[tag_end:]


def _container_dim_map(root) -> dict[str, tuple[Optional[float], Optional[float], Optional[float], Optional[float]]]:
    """
    Build map: container_id -> (x, y, width, height) as floats when present.
    """
    m: dict[str, tuple[Optional[float], Optional[float], Optional[float], Optional[float]]] = {}
    for c in root.iter("container"):
        cid = c.get("id")
        if not cid:
            continue
        def _f(name: str) -> Optional[float]:
            v = c.get(name)
            if v is None or v == "":
                return None
            try:
                return float(v)
            except ValueError:
                return None
        m[cid] = (_f("x"), _f("y"), _f("width"), _f("height"))
    return m


def _apply_resize_layout(content: str) -> tuple[str, int]:
    """
    Apply resize logic while preserving the original XML formatting/CDATA by patching
    only container x/y/width/height attributes in the original text.
    """
    root_before = ET.fromstring(content)
    before = _container_dim_map(root_before)

    root_after = ET.fromstring(content)
    page_width = float(root_after.get("width", "693") or 693.0)
    pages_processed = 0
    for section in root_after.iter("section"):
        for page in section.findall("page"):
            _resize_process_page(page, page_width)
            pages_processed += 1

    after = _container_dim_map(root_after)

    # Patch only changed numeric attributes; everything else stays byte-identical.
    patched = content
    for cid, (bx, by, bw, bh) in before.items():
        ax, ay, aw, ah = after.get(cid, (None, None, None, None))
        changes = {}

        def _changed(a: Optional[float], b: Optional[float]) -> bool:
            if a is None and b is None:
                return False
            if a is None or b is None:
                return True
            return abs(a - b) > 1e-9

        if _changed(ax, bx) and ax is not None:
            changes["x"] = ax
        if _changed(ay, by) and ay is not None:
            changes["y"] = ay
        if _changed(aw, bw) and aw is not None:
            changes["width"] = aw
        if _changed(ah, bh) and ah is not None:
            changes["height"] = ah

        if changes:
            patched = patch_container_attrs_by_id(patched, cid, **changes)

    return patched, pages_processed


SPAN = '<span class="font-montserrat" style="font-size:12px;color:#000000;">'


def build_caption_html(
    plats: str,
    datum_long: str,
    klockslag: str,
    väder: str,
    ålder: str,
    bildtext: str,
) -> str:
    """Bygger HTML enligt BookWright-format."""
    parts = ['<p class="align-left line-height-qt">']
    if plats:
        parts.append(f"{SPAN}<strong>{html.escape(plats)}</strong></span>")
    meta_parts = [datum_long, klockslag, väder]
    meta = " · ".join(p for p in meta_parts if p)
    if meta or ålder:
        sep = " · " if plats else ""
        parts.append(f"{SPAN}{sep}{html.escape(meta)} · </span>")
    if ålder:
        parts.append(f"{SPAN}<em>{html.escape(ålder)}</em></span>")
    parts.append("</p>")
    if bildtext:
        parts.append("<p class=\"align-left line-height-qt\">")
        parts.append(f"{SPAN}{html.escape(bildtext)}</span></p>")
    return "".join(parts).replace("]]>", "]]]]><![CDATA[>")


def _bildsida_csv_path(blurb_path: Path) -> Path:
    """Path to CSV next to .blurb: stem without a trailing '-new' + .csv."""
    stem = blurb_path.stem
    if stem.endswith("-new"):
        stem = stem[: -len("-new")]
    return blurb_path.parent / (stem + ".csv")


def load_bildsida_csv(blurb_path: Path) -> tuple[dict[str, str], dict[str, str], list[list[str]]]:
    """Reads the CSV file. Returns (city_map, weather_map, rows).
    Format: filename, page_number, city, weather, lat, lon
    The same filename may appear in multiple rows (one per placement/page). city_map/weather
    are stored per filename; later rows overwrite earlier ones (same metadata)."""
    city_map: dict[str, str] = {}
    weather_map: dict[str, str] = {}
    rows: list[list[str]] = []
    csv_path = _bildsida_csv_path(blurb_path)
    if not csv_path.is_file():
        return (city_map, weather_map, rows)
    with csv_path.open(encoding="utf-8", newline="") as f:
        for row in csv.reader(f):
            if not row or not row[0].strip():
                continue
            fn = row[0].strip()
            sid = row[1].strip() if len(row) > 1 else ""
            stad = row[2].strip() if len(row) > 2 else ""
            väder = row[3].strip() if len(row) > 3 else ""
            lat = row[4].strip() if len(row) > 4 else ""
            lon = row[5].strip() if len(row) > 5 else ""
            rows.append([fn, sid, stad, väder, lat or "", lon or ""])
            if stad:
                city_map[fn] = stad
            if väder:
                weather_map[fn] = väder
    return (city_map, weather_map, rows)


def save_bildsida_csv(blurb_path: Path, rows: list[list[str]]) -> None:
    """Saves the .csv file."""
    csv_path = _bildsida_csv_path(blurb_path)
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows(rows)


def _row_filename_basename(r: list[str]) -> str:
    fn = (r[0] if r else "").strip()
    return Path(fn).name if fn and "/" in fn else fn


def _csv_row_is_front_cover(r: list[str]) -> bool:
    """Page -1 in Blurb is front cover; it should not be included in the .csv."""
    if not r or len(r) < 2:
        return False
    return (r[1] or "").strip() == "-1"


def _preserve_manual_csv_values(
    new_rows: list[list[str]],
    existing_rows: list[list[str]],
) -> None:
    """
    Preserve manually edited city/weather/lat/lon from existing CSV.
    Matches rows by filename occurrence order (1st, 2nd, ...) in document order.
    Lat/lon: if either column was set in CSV, keep both columns from that row (no overwrite from EXIF).
    """
    existing_occ_idx: dict[str, int] = {}
    existing_lookup: dict[tuple[str, int], tuple[str, str, str, str]] = {}

    for r in existing_rows:
        fn = _row_filename_basename(r)
        if not fn:
            continue
        i = existing_occ_idx.get(fn, 0)
        existing_occ_idx[fn] = i + 1
        city = (r[2] if len(r) > 2 else "").strip()
        weather = (r[3] if len(r) > 3 else "").strip()
        lat = (r[4] if len(r) > 4 else "").strip()
        lon = (r[5] if len(r) > 5 else "").strip()
        existing_lookup[(fn, i)] = (city, weather, lat, lon)

    new_occ_idx: dict[str, int] = {}
    for r in new_rows:
        fn = _row_filename_basename(r)
        if not fn:
            continue
        i = new_occ_idx.get(fn, 0)
        new_occ_idx[fn] = i + 1
        old_city, old_weather, old_lat, old_lon = existing_lookup.get(
            (fn, i), ("", "", "", "")
        )
        while len(r) < 6:
            r.append("")
        if old_city:
            r[2] = old_city
        if old_weather:
            r[3] = old_weather
        if old_lat or old_lon:
            r[4] = old_lat
            r[5] = old_lon


def sync_csv_sids_from_content(content: str, rows: list[list[str]]) -> None:
    """Sets page number (column 2) from each image position in layout.
    Multiple rows with the same filename match occurrence 1, 2, … in document order.
    Previously the page was only updated when a text box was found; moved images
    without a text box then kept stale page numbers in the CSV."""
    page_re = re.compile(r'<page[^>]*number="(-?\d+)"')
    src_pattern = re.compile(r'<image[^>]+src="([^"]+)"')
    occ: list[tuple[str, str, str]] = []
    for m in src_pattern.finditer(content):
        img_src = m.group(1)
        fn = Path(img_src).name if "/" in img_src else img_src
        pages = page_re.findall(content[: m.start()])
        sid = pages[-1] if pages else ""
        occ.append((img_src, fn, sid))

    prev_counts: dict[tuple[str, str], int] = {}
    for r in rows:
        if not r or not (r[0] or "").strip():
            continue
        raw = (r[0] or "").strip()
        want = _row_filename_basename(r)
        key = (want, raw)
        k = prev_counts.get(key, 0)
        prev_counts[key] = k + 1

        matched = 0
        for img_src, fn, sid in occ:
            if fn != want and img_src != raw:
                continue
            if matched == k:
                if sid:
                    while len(r) < 2:
                        r.append("")
                    r[1] = sid
                break
            matched += 1


_FLOAT_RE = r'([\d.]+)'


def _find_container_attrs(content: str, pos: int, ctype: str) -> Optional[dict[str, float]]:
    """Returns {x, y, width, height} for a container."""
    start = content.rfind("<container", 0, pos)
    if start < 0:
        return None
    tag_end = content.find(">", start)
    opening = content[start : tag_end + 1]
    if f'type="{ctype}"' not in opening:
        return None
    attrs = {}
    for name in ("x", "y", "width", "height"):
        m = re.search(rf'{name}="{_FLOAT_RE}"', opening)
        if m:
            try:
                attrs[name] = float(m.group(1))
            except ValueError:
                pass
    return attrs if attrs else None


def _get_page_bounds(content: str, pos: int) -> tuple[int, int]:
    """Returns (page_start, page_end) for the page containing pos."""
    page_start = content.rfind("<page", 0, pos)
    if page_start < 0:
        page_start = 0
    page_end = content.find("</page>", pos)
    if page_end < 0:
        page_end = len(content)
    else:
        page_end += len("</page>")
    return (page_start, page_end)


def _get_page_pairs(
    content: str,
    page_start: int,
    page_end: int,
) -> list[tuple[int, tuple[int, int]]]:
    """
    Collects all images and text boxes on the page in document order.
    Pairs 1:1 in the order they appear – first image with first text box, etc.
    """
    images, texts = _get_page_images_and_texts(content, page_start, page_end)
    if not images or not texts:
        return []
    n = min(len(images), len(texts))
    return [(images[i], texts[i]) for i in range(n)]


def _get_page_images_and_texts(
    content: str,
    page_start: int,
    page_end: int,
) -> tuple[list[int], list[tuple[int, int]]]:
    """Returns (images [img_pos for each image slot], texts [(inner_start, inner_end)]) for the page.
    Counts all type=\"image\" containers (including empty). img_pos = src span start if filled, else
    container start (for empty slots, img_pos is not used for text pairing)."""
    page_tag = content[page_start : page_start + 300]
    num_m = re.search(r'number="(-?\d+)"', page_tag)
    if num_m and int(num_m.group(1)) < 1:
        return ([], [])  # Omslag/master

    images: list[int] = []
    texts: list[tuple[int, int]] = []
    src_re = re.compile(r'<image[^>]+src="([^"]+)"')

    pos = page_start
    while pos < page_end:
        container = content.find("<container", pos)
        if container < 0 or container >= page_end:
            break
        tag_end = content.find(">", container)
        opening = content[container : tag_end + 1]

        if 'type="image"' in opening:
            img_m = src_re.search(content, container, container + 500)
            # Count all image slots (empty + filled). img_pos = match start for pairing.
            images.append(img_m.start() if img_m else container)
        elif 'type="text"' in opening:
            text_start = content.find("<text", container)
            if text_start >= 0:
                cdata_start = content.find("<![CDATA[", text_start)
                if cdata_start >= 0 and cdata_start < container + 2000:
                    inner_start = cdata_start + len("<![CDATA[")
                    inner_end = content.find("]]>", inner_start)
                    if inner_end >= 0:
                        texts.append((inner_start, inner_end))
                else:
                    text_open_end = content.find(">", text_start)
                    text_close = content.find("</text>", text_open_end + 1 if text_open_end >= 0 else text_start)
                    if text_open_end >= 0 and text_close >= 0:
                        texts.append((text_open_end + 1, text_close))

        pos = content.find("</container>", container) + 1

    return (images, texts)


def _get_all_pages(content: str) -> list[tuple[int, int, int]]:
    """Returns [(page_num, page_start, page_end), ...] for all pages."""
    result: list[tuple[int, int, int]] = []
    pos = 0
    while True:
        page_start = content.find("<page", pos)
        if page_start < 0:
            break
        tag_end = content.find(">", page_start)
        if tag_end < 0:
            break
        # Check if the page is self-closing (<page .../>)
        if content[tag_end - 1 : tag_end + 1] == "/>":
            page_end = tag_end + 1
        else:
            page_end = content.find("</page>", page_start)
            if page_end < 0:
                break
            page_end += len("</page>")
        page_tag = content[page_start : page_start + 300]
        num_m = re.search(r'number="(-?\d+)"', page_tag)
        page_num = int(num_m.group(1)) if num_m else -1
        result.append((page_num, page_start, page_end))
        pos = page_end
    return result


def _find_text_cdata(
    content: str,
    img_pos: int,
    page_start: int,
    page_end: int,
    page_pairs: Optional[dict[int, tuple[int, int]]] = None,
) -> Optional[tuple[int, int]]:
    """
    Finds the text box for an image. If page_pairs is provided (img_pos -> (start,end)), it is used.
    Otherwise fallback: nearest vertical candidate with the same x.
    """
    if page_pairs is not None and img_pos in page_pairs:
        return page_pairs[img_pos]

    img_attrs = _find_container_attrs(content, img_pos, "image")
    if not img_attrs:
        return None
    img_x = int(round(img_attrs.get("x", 0)))
    img_center_y = img_attrs.get("y", 0) + img_attrs.get("height", 0) / 2

    page_tag = content[page_start : page_start + 300]
    num_m = re.search(r'number="(-?\d+)"', page_tag)
    if num_m and int(num_m.group(1)) < 1:
        return None

    candidates: list[tuple[float, int, int]] = []
    pos = page_start
    while pos < page_end:
        container = content.find("<container", pos)
        if container < 0 or container >= page_end:
            break
        tag_end = content.find(">", container)
        opening = content[container : tag_end + 1]
        if 'type="text"' not in opening:
            pos = container + 1
            continue

        attrs = _find_container_attrs(content, container, "text")
        txt_x = int(round(attrs.get("x", 0))) if attrs else 0
        txt_y = attrs.get("y", 0) if attrs else 0
        txt_h = attrs.get("height", 0) if attrs else 0
        txt_center_y = txt_y + txt_h / 2

        text_start = content.find("<text", container)
        if text_start < 0:
            pos = container + 1
            continue

        cdata_start = content.find("<![CDATA[", text_start)
        if cdata_start >= 0 and cdata_start < container + 2000:
            inner_start = cdata_start + len("<![CDATA[")
            inner_end = content.find("]]>", inner_start)
            if inner_end < 0:
                pos = container + 1
                continue
        else:
            text_open_end = content.find(">", text_start)
            text_close = content.find("</text>", text_open_end + 1 if text_open_end >= 0 else text_start)
            if text_open_end < 0 or text_close < 0:
                pos = container + 1
                continue
            inner_start = text_open_end + 1
            inner_end = text_close

        if txt_x == img_x:
            priority = abs(img_center_y - txt_center_y)
        else:
            priority = 1000 + abs(img_center_y - txt_center_y)
        candidates.append((priority, inner_start, inner_end))
        pos = content.find("</container>", container) + 1

    if not candidates:
        return None
    candidates.sort(key=lambda c: c[0])
    return (candidates[0][1], candidates[0][2])


def run(source_blurb: Path) -> tuple[int, Optional[Path]]:
    """
    Reads source .blurb, fills text boxes with EXIF captions, writes to source-new.blurb only.
    Returns (number of caption cells written, output path or None if nothing was written).
    """
    import sqlite3

    source_blurb = source_blurb.resolve()
    output_blurb = source_blurb.parent / f"{source_blurb.stem}-new{source_blurb.suffix}"

    conn = sqlite3.connect(source_blurb)
    cursor = conn.cursor()

    cursor.execute("SELECT filecontent FROM Files WHERE filepath = 'bbf2.xml'")
    row = cursor.fetchone()
    if not row:
        conn.close()
        return 0, None

    content = row[0].decode("utf-8")

    # Image files: filepath -> blob
    cursor.execute("SELECT filepath, filecontent FROM Files WHERE filepath LIKE 'images/%'")
    images_db: dict[str, bytes] = {}
    for filepath, blob in cursor.fetchall():
        images_db[Path(filepath).name] = blob
        images_db[filepath] = blob

    conn.close()

    city_map, weather_map, _loaded_csv_rows = load_bildsida_csv(source_blurb)
    uses_cdata_text = "<![CDATA[" in content
    rows: list[list[str]] = []
    page_re = re.compile(r'<page[^>]*number="(-?\d+)"')
    src_pattern = re.compile(r'<image[^>]+src="([^"]+)"')
    csv_occ_lookup: dict[tuple[str, int], tuple[str, str]] = {}
    csv_ll_occ_lookup: dict[tuple[str, int], tuple[str, str]] = {}
    csv_occ_counts: dict[str, int] = {}
    for r in _loaded_csv_rows:
        fn = _row_filename_basename(r)
        if not fn:
            continue
        i = csv_occ_counts.get(fn, 0)
        csv_occ_counts[fn] = i + 1
        city = (r[2] if len(r) > 2 else "").strip()
        weather = (r[3] if len(r) > 3 else "").strip()
        lat_c = (r[4] if len(r) > 4 else "").strip()
        lon_c = (r[5] if len(r) > 5 else "").strip()
        csv_occ_lookup[(fn, i)] = (city, weather)
        csv_ll_occ_lookup[(fn, i)] = (lat_c, lon_c)

    replacements: list[tuple[int, int, str, str]] = []  # (start, end, new_html, img_file)
    used_text_ranges: set[tuple[int, int]] = set()
    page_bildtext_written: dict[tuple[int, int], set[str]] = {}  # (page_start,page_end) -> captions already written on page

    # Build per-page pairing cache: img_pos -> (text_start, text_end)
    page_pairs: dict[int, tuple[int, int]] = {}
    page_cache: dict[tuple[int, int], list] = {}
    images_in_blurb: set[str] = set()
    for m in src_pattern.finditer(content):
        img_src = m.group(1)
        fn = Path(img_src).name if "/" in img_src else img_src
        images_in_blurb.add(fn)
        page_start, page_end = _get_page_bounds(content, m.start())
        key = (page_start, page_end)
        if key not in page_cache:
            page_cache[key] = _get_page_pairs(content, page_start, page_end)
        for img_pos, cdata_range in page_cache[key]:
            page_pairs[img_pos] = cdata_range

    image_occ_seen: dict[str, int] = {}
    for m in src_pattern.finditer(content):
        img_file = m.group(1)
        img_pos = m.start()
        img_file_base = Path(img_file).name if "/" in img_file else img_file
        occ_i = image_occ_seen.get(img_file_base, 0)
        image_occ_seen[img_file_base] = occ_i + 1
        manual_city, manual_weather = csv_occ_lookup.get((img_file_base, occ_i), ("", ""))
        csv_lat, csv_lon = csv_ll_occ_lookup.get((img_file_base, occ_i), ("", ""))
        if not ENABLE_GPS:
            manual_city = ""
        if not ENABLE_WEATHER:
            manual_weather = ""
        blob = images_db.get(img_file) or images_db.get(f"images/{img_file}")
        if not blob:
            print(f"  Skipping {img_file} (image missing in .blurb)")
            continue

        exif = get_exif_from_blob(blob)
        exif.caption = _fix_mojibake(exif.caption)
        gps_for_api = _gps_from_csv_or_exif(csv_lat, csv_lon, exif.gps)
        if not exif.caption and not exif.datetime_str:
            print(f"  {img_file}: no EXIF data")
            continue

        page_start, page_end = _get_page_bounds(content, img_pos)
        found = _find_text_cdata(content, img_pos, page_start, page_end, page_pairs)
        if not found:
            plats_pre, väder_pre = get_plats_väder(
                img_file, gps_for_api, exif.datetime_str, city_map, weather_map
            )
            if manual_city:
                plats_pre = manual_city
            if manual_weather:
                väder_pre = manual_weather
            lat_pre = f"{exif.gps[0]:.6f}" if exif.gps else ""
            lon_pre = f"{exif.gps[1]:.6f}" if exif.gps else ""
            pages_pre = page_re.findall(content[:img_pos])
            sid_pre = pages_pre[-1] if pages_pre else ""
            rows.append([img_file, sid_pre, plats_pre or "", väder_pre or "", lat_pre, lon_pre])
            city_map[img_file] = plats_pre
            weather_map[img_file] = väder_pre
            print(f"  {img_file}: no matching text box found")
            continue

        start, end = found

        plats, väder = get_plats_väder(
            img_file, gps_for_api, exif.datetime_str, city_map, weather_map
        )
        if manual_city:
            plats = manual_city
        if manual_weather:
            väder = manual_weather

        lat_str = f"{exif.gps[0]:.6f}" if exif.gps else ""
        lon_str = f"{exif.gps[1]:.6f}" if exif.gps else ""

        pages = page_re.findall(content[:img_pos])
        sid = pages[-1] if pages else ""
        rows.append([img_file, sid, plats if ENABLE_GPS else "", väder if ENABLE_WEATHER else "", lat_str, lon_str])
        city_map[img_file] = plats
        weather_map[img_file] = väder

        if "###" in content[start:end]:
            print(f"  {img_file}: text box contains ###, skipping")
            continue
        if (start, end) in used_text_ranges:
            print(f"  {img_file}: text box already used, skipping")
            continue

        bildtext = exif.caption or ""
        page_key = (page_start, page_end)
        if page_key not in page_bildtext_written:
            page_bildtext_written[page_key] = set()
        if bildtext and bildtext in page_bildtext_written[page_key]:
            caption_bildtext = ""  # Write only title (location/date/weather), not caption text
            print(f"  {img_file}: same caption already on page, writing title only")
        else:
            caption_bildtext = bildtext
            if bildtext:
                page_bildtext_written[page_key].add(bildtext)

        used_text_ranges.add((start, end))
        datum_long, klockslag = format_date_long(exif.datetime_str)
        ålder = _shorty(format_age(exif.datetime_str))

        caption_html = build_caption_html(
            plats if ENABLE_GPS else "",
            datum_long,
            klockslag,
            väder if ENABLE_WEATHER else "",
            ålder,
            caption_bildtext,
        )
        new_text_payload = caption_html if uses_cdata_text else html.escape(caption_html)
        replacements.append((start, end, new_text_payload, img_file))
        preview = f"{plats} · {datum_long}"[:50] + "..." if (plats or datum_long) else exif.caption[:40] + "..."
        print(f"  {img_file}: {preview}")

    # Clear text boxes that did not get new text (except ###)
    for page_num, page_start, page_end in _get_all_pages(content):
        if page_num < 1:
            continue
        images, texts = _get_page_images_and_texts(content, page_start, page_end)
        for tr in texts:
            if tr in used_text_ranges:
                continue
            if "###" in content[tr[0]:tr[1]]:
                continue
            replacements.append((tr[0], tr[1], "", ""))

    n_captions = sum(1 for r in replacements if r[2])

    if not replacements:
        return 0, None

    shutil.copy2(source_blurb, output_blurb)

    for start, end, new, _ in sorted(replacements, key=lambda r: r[0], reverse=True):
        content = content[:start] + new + content[end:]

    if rows:
        # Remove rows for images that no longer exist in the book
        rows_filtered = [r for r in rows if _row_filename_basename(r) in images_in_blurb]
        _preserve_manual_csv_values(rows_filtered, _loaded_csv_rows)
        sync_csv_sids_from_content(content, rows_filtered)
        rows_filtered = [r for r in rows_filtered if not _csv_row_is_front_cover(r)]
        save_bildsida_csv(output_blurb, rows_filtered)

    if RESIZE_LAYOUT:
        try:
            content, pages_processed = _apply_resize_layout(content)
            print(f"  Resize layout complete: {pages_processed} pages.")
        except Exception as e:
            print(f"  Warning: resize layout failed ({e}).")

    conn = sqlite3.connect(output_blurb)
    conn.execute(
        "UPDATE Files SET filecontent = ?, filesize = ? WHERE filepath = 'bbf2.xml'",
        (content.encode("utf-8"), len(content.encode("utf-8"))),
    )
    conn.commit()
    conn.close()

    return n_captions, output_blurb


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    blurb_path = Path(sys.argv[1]).expanduser().resolve()
    if not blurb_path.exists():
        print(f"File {blurb_path} was not found.")
        sys.exit(1)

    n, out_path = run(blurb_path)
    if out_path is not None:
        print(f"\nUpdated {n} captions in {out_path}")
    else:
        print("\nNo output file written (nothing to update or missing bbf2.xml).")


if __name__ == "__main__":
    main()