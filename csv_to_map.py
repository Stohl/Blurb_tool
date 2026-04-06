#!/usr/bin/env python3
"""
Build a single full-page HTML map from a book CSV (same columns as blurb_captions output).

CSV columns: filename, page, city, weather, lat, lon [, optional extra ...]

Usage:
  python3 csv_to_map.py album.csv

Writes <stem>_map.html next to the CSV and opens it in the default browser.

The HTML includes a Map style dropdown in the top-left; initial style follows
DEFAULT_STYLE in this file.
"""

from __future__ import annotations

import argparse
import csv
import json
import webbrowser
from html import escape
from pathlib import Path
from typing import Any, List, Optional

# MapTiler (https://cloud.maptiler.com/) — set your own key for tiles to load in the browser.
# The script always runs and writes HTML; MapTiler maps will not load without a valid key.
MAPTILER_KEY = "API-Key"


def _style_url(map_id: str) -> str:
    return f"https://api.maptiler.com/maps/{map_id}/style.json?key={MAPTILER_KEY}"


# Preset style ids: https://docs.maptiler.com/cloud/api/maps/
MAP_STYLES = {
    "streets": _style_url("streets-v2"),
    "outdoor": _style_url("outdoor-v2"),
    "satellite": _style_url("satellite"),
    "basic": _style_url("basic-v2"),
    "winter": _style_url("winter-v2"),
    "toner": _style_url("toner-v2"),
    "aquarelle": _style_url("aquarelle-v4"),
    "custom": f"https://api.maptiler.com/maps/0123/style.json?key={MAPTILER_KEY}",
}

DEFAULT_STYLE = "streets"

# Pin: line from dot to page label (px). Set 0 to hide.
PIN_LINE_WIDTH = 22
# Cluster bubble: target max characters per line (breaks after comma between page numbers when possible).
CLUSTER_CHARS_PER_LINE = 22


def _parse_csv(csv_path: Path) -> List[dict[str, Any]]:
    """Rows with lat/lon, page id, filename, optional city (for popup)."""
    rows: List[dict[str, Any]] = []
    with csv_path.open(encoding="utf-8", newline="") as f:
        for row in csv.reader(f):
            if not row or not row[0].strip():
                continue
            fn = row[0].strip()
            page = row[1].strip() if len(row) > 1 else ""
            city = row[2].strip() if len(row) > 2 else ""
            lat_str = row[4].strip() if len(row) > 4 else ""
            lon_str = row[5].strip() if len(row) > 5 else ""
            lat: Optional[float] = None
            lon: Optional[float] = None
            if lat_str and lon_str:
                try:
                    lat = float(lat_str.replace(",", "."))
                    lon = float(lon_str.replace(",", "."))
                except ValueError:
                    pass
            rows.append(
                {
                    "filename": fn,
                    "page": page,
                    "city": city,
                    "lat": lat,
                    "lon": lon,
                }
            )
    return rows


def create_map_html(
    rows: List[dict[str, Any]],
    output_path: Path,
    initial_style_key: str,
) -> None:
    """Write one HTML file: MapLibre map + clustered markers."""
    with_coords = [r for r in rows if r["lat"] is not None and r["lon"] is not None]
    if not with_coords:
        print("No rows with valid latitude/longitude. Nothing to map.")
        return

    avg_lat = sum(r["lat"] for r in with_coords) / len(with_coords)
    avg_lon = sum(r["lon"] for r in with_coords) / len(with_coords)

    features: List[dict[str, Any]] = []
    for r in with_coords:
        fn = r["filename"]
        sid = r["page"] or ""
        city = r["city"] or ""
        popup_parts = [escape(fn)]
        if city:
            popup_parts.append(escape(city))
        popup_label = " — ".join(popup_parts)
        features.append(
            {
                "type": "Feature",
                "geometry": {
                    "type": "Point",
                    "coordinates": [r["lon"], r["lat"]],
                },
                "properties": {
                    "sid": sid,
                    "popup": popup_label,
                },
            }
        )

    geojson_str = json.dumps({"type": "FeatureCollection", "features": features})
    line_display = "none" if PIN_LINE_WIDTH == 0 else "block"
    label_left = 6 + PIN_LINE_WIDTH
    style_url = MAP_STYLES[initial_style_key]
    map_styles_json = json.dumps(MAP_STYLES)
    style_options_html = "".join(
        f'<option value="{escape(k)}"{" selected" if k == initial_style_key else ""}>{escape(k.title())}</option>'
        for k in MAP_STYLES
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Photo map</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap" rel="stylesheet" />
  <link href="https://unpkg.com/maplibre-gl@3/dist/maplibre-gl.css" rel="stylesheet" />
  <script src="https://unpkg.com/maplibre-gl@3/dist/maplibre-gl.js"></script>
  <script src="https://unpkg.com/supercluster@7.1.5/dist/supercluster.min.js"></script>
  <style>
    html, body {{ height: 100%; margin: 0; padding: 0; font-family: 'Montserrat', sans-serif; position: relative; }}
    #map {{ height: 100%; width: 100%; }}
    .pin-icon {{ background: transparent; border: none; overflow: visible; }}
    .pin-wrap {{ position: relative; width: 1px; height: 1px; overflow: visible; }}
    .pin-dot {{ position: absolute; left: -4px; top: -4px; width: 8px; height: 8px; background: #222; border: 2px solid #fff; border-radius: 50%; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }}
    .pin-line {{ position: absolute; left: 4px; top: 0; width: {PIN_LINE_WIDTH}px; height: 1px; background: rgba(0,0,0,0.15); display: {line_display}; }}
    .pin-label {{ position: absolute; left: {label_left}px; top: -11px; background: #fff; color: #222; border-radius: 14px; padding: 1px 7px; font-size: 13px; font-weight: 700; border: 1px solid #ddd; white-space: nowrap; }}
    .pin-label--cluster {{ white-space: normal; line-height: 1; max-width: min(560px, 92vw); font-size: 11.5px; padding: 1px 6px; letter-spacing: -0.02em; }}
    .pin-label--cluster .cluster-line {{ display: block; white-space: nowrap; line-height: 1.05; margin: 0; padding: 0; }}
    .pin-label--single {{ font-size: 13px; padding: 1px 6px; }}
    .map-style-bar {{
      position: absolute; z-index: 1; top: 10px; left: 10px;
      display: flex; align-items: center; gap: 8px;
      background: rgba(255,255,255,0.95); border: 1px solid #ddd; border-radius: 8px;
      padding: 6px 10px; font-size: 13px; font-weight: 600; color: #333;
      box-shadow: 0 1px 4px rgba(0,0,0,0.12);
    }}
    .map-style-bar label {{ cursor: pointer; margin: 0; }}
    .map-style-bar select {{
      font: inherit; font-weight: 500; padding: 4px 8px; border-radius: 6px;
      border: 1px solid #ccc; background: #fff; cursor: pointer; min-width: 140px;
    }}
  </style>
</head>
<body>
  <div id="map"></div>
  <div class="map-style-bar" aria-label="Map style">
    <label for="map-style-select">Map</label>
    <select id="map-style-select">{style_options_html}</select>
  </div>
  <script>
    var pointsData = {geojson_str};
    var mapStyles = {map_styles_json};
    var map = new maplibregl.Map({{
      container: 'map',
      style: '{style_url}',
      center: [{avg_lon}, {avg_lat}],
      zoom: 11
    }});
    map.addControl(new maplibregl.NavigationControl(), 'top-right');
    var clusterCharsPerLine = {CLUSTER_CHARS_PER_LINE};
    var markers = [];
    var index = new Supercluster({{ radius: 50, maxZoom: 18 }});
    index.load(pointsData.features);
    function formatPagesForCluster(pages) {{
      if (pages.length === 0) return '';
      var nbsp = String.fromCharCode(160);
      var sep = ',';
      var maxLen = clusterCharsPerLine;
      var full = 'p.' + nbsp + pages.map(function(p) {{ return String(p); }}).join(sep);
      var tokens = full.split(sep);
      var lines = [];
      var line = '';
      function pushChunks(s) {{
        for (var c = 0; c < s.length; c += maxLen) {{
          lines.push('<span class="cluster-line">' + s.slice(c, c + maxLen) + '</span>');
        }}
      }}
      for (var ti = 0; ti < tokens.length; ti++) {{
        var piece = tokens[ti];
        if (line.length === 0) {{
          line = piece;
          if (line.length > maxLen) {{ pushChunks(line); line = ''; }}
          continue;
        }}
        var add = sep + piece;
        if (line.length + add.length <= maxLen) {{
          line += add;
        }} else {{
          lines.push('<span class="cluster-line">' + line + '</span>');
          line = piece;
          if (line.length > maxLen) {{ pushChunks(line); line = ''; }}
        }}
      }}
      if (line.length) {{
        if (line.length > maxLen) pushChunks(line);
        else lines.push('<span class="cluster-line">' + line + '</span>');
      }}
      return lines.join('');
    }}
    function createPinEl(labelHtml, isCluster) {{
      var el = document.createElement('div');
      el.className = 'pin-icon';
      var labelClass = isCluster ? 'pin-label pin-label--cluster' : 'pin-label pin-label--single';
      el.innerHTML = '<div class="pin-wrap"><span class="pin-dot"></span><span class="pin-line"></span><span class="' + labelClass + '">' + labelHtml + '</span></div>';
      return el;
    }}
    function updateMarkers() {{
      markers.forEach(function(m) {{ m.remove(); }});
      markers = [];
      var bbox = map.getBounds().toArray().flat();
      var zoom = map.getZoom();
      var clusters = index.getClusters(bbox, Math.floor(zoom));
      clusters.forEach(function(cluster) {{
        var coords = cluster.geometry.coordinates;
        var el, popupHtml;
        if (cluster.properties.cluster) {{
          var leaves = index.getLeaves(cluster.properties.cluster_id, Infinity);
          var pages = leaves.map(function(l) {{ return l.properties.sid || ''; }}).filter(Boolean);
          pages = Array.from(new Set(pages)).sort(function(a,b) {{
            var ai = parseInt(a, 10), bi = parseInt(b, 10);
            if (isNaN(ai) && isNaN(bi)) return String(a).localeCompare(String(b));
            if (isNaN(ai)) return 1;
            if (isNaN(bi)) return -1;
            return ai - bi;
          }});
          el = createPinEl(formatPagesForCluster(pages) || String(cluster.properties.point_count), true);
          popupHtml = leaves.map(function(l) {{ return l.properties.popup; }}).join('<br>');
        }} else {{
          var p = cluster.properties.sid || '';
          el = createPinEl(p ? ('p.' + String.fromCharCode(160) + p) : '', false);
          popupHtml = cluster.properties.popup || '';
        }}
        var m = new maplibregl.Marker({{ element: el }})
          .setLngLat(coords)
          .setPopup(new maplibregl.Popup({{ maxWidth: '320px' }}).setHTML(popupHtml))
          .addTo(map);
        markers.push(m);
      }});
    }}
    map.on('load', updateMarkers);
    map.on('style.load', updateMarkers);
    map.on('moveend', updateMarkers);
    document.getElementById('map-style-select').addEventListener('change', function(e) {{
      var key = e.target.value;
      if (mapStyles[key]) map.setStyle(mapStyles[key]);
    }});
  </script>
</body>
</html>
"""

    output_path.write_text(html, encoding="utf-8")
    print(f"Wrote {output_path}")
    webbrowser.open(output_path.as_uri())


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build an HTML map from a book CSV (filename, page, city, weather, lat, lon)."
    )
    parser.add_argument("csv_file", type=str, help="Path to CSV file")
    args = parser.parse_args()
    csv_path = Path(args.csv_file).expanduser().resolve()
    if not csv_path.is_file():
        print(f"File not found: {csv_path}")
        raise SystemExit(1)

    rows = _parse_csv(csv_path)
    print(f"Read {len(rows)} row(s) from {csv_path.name}")

    out = csv_path.parent / (csv_path.stem + "_map.html")
    create_map_html(rows, out, DEFAULT_STYLE)


if __name__ == "__main__":
    main()
