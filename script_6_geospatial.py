"""
SCRIPT 6 — GEOSPATIAL ANALYSIS & MAPPING
================================================
What it does:
  - Plots all GPS collection points on an interactive map
  - Color-codes points: valid (green), invalid (red), out-of-bounds (orange)
  - Shows each worker's points in a different color
  - Creates a heatmap of collection density
  - Calculates coverage area
  - Exports points as GeoJSON (usable in QGIS / ArcGIS)
  - Exports points as Shapefile

Input : cleaned_data.csv
Output: map_all_points.html      ← open in browser (interactive)
        map_by_worker.html       ← one layer per worker
        heatmap.html             ← density heatmap
        points_export.geojson    ← for QGIS/ArcGIS
"""

import pandas as pd
import numpy as np
import json
import os
from math import radians, sin, cos, sqrt, atan2

# ── CONFIG ──────────────────────────────────────────────────────────────
INPUT_FILE   = "cleaned_data.csv"
ACCURACY_THRESH = 10.0
BOUNDARY = {"lat_min":-1.35,"lat_max":-1.25,"lon_min":36.80,"lon_max":36.85}

WORKER_COLORS = [
    "#E74C3C","#3498DB","#2ECC71","#F39C12","#9B59B6",
    "#1ABC9C","#E67E22","#34495E","#E91E63","#00BCD4"
]

# ── HELPERS ──────────────────────────────────────────────────────────────
def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1; dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1)*cos(lat2)*sin(dlon/2)**2
    return R * 2 * atan2(sqrt(a), sqrt(1-a))

def is_valid(row):
    try:
        lat, lon, acc = float(row["latitude"]), float(row["longitude"]), float(row["accuracy_m"])
        return (acc < ACCURACY_THRESH and
                BOUNDARY["lat_min"] < lat < BOUNDARY["lat_max"] and
                BOUNDARY["lon_min"] < lon < BOUNDARY["lon_max"])
    except:
        return False

# ── 1. LOAD & PREP ───────────────────────────────────────────────────────
print("[LOAD] Loading data...")
df = pd.read_csv(INPUT_FILE)
df["latitude"]   = pd.to_numeric(df["latitude"],   errors="coerce")
df["longitude"]  = pd.to_numeric(df["longitude"],  errors="coerce")
df["accuracy_m"] = pd.to_numeric(df["accuracy_m"], errors="coerce")
df = df.dropna(subset=["latitude","longitude"])

df["valid"]   = df.apply(is_valid, axis=1)
df["in_bbox"] = (df["latitude"].between(BOUNDARY["lat_min"], BOUNDARY["lat_max"]) &
                 df["longitude"].between(BOUNDARY["lon_min"], BOUNDARY["lon_max"]))

workers   = df["worker_id"].unique()
w_color   = {wid: WORKER_COLORS[i % len(WORKER_COLORS)] for i, wid in enumerate(workers)}
center_lat = df["latitude"].mean()
center_lon = df["longitude"].mean()

print(f"   {len(df)} points loaded | {df['valid'].sum()} valid | center: ({center_lat:.4f}, {center_lon:.4f})")

# ── 2. GEOJSON EXPORT ────────────────────────────────────────────────────
print("\n Exporting GeoJSON...")
features = []
for _, row in df.iterrows():
    features.append({
        "type": "Feature",
        "geometry": {
            "type": "Point",
            "coordinates": [float(row["longitude"]), float(row["latitude"])]
        },
        "properties": {
            "worker_id"     : str(row.get("worker_id","")),
            "worker_name"   : str(row.get("worker_name","")),
            "team"          : str(row.get("team","")),
            "date"          : str(row.get("date","")),
            "point_number"  : int(row.get("point_number", 0)),
            "accuracy_m"    : float(row["accuracy_m"]) if pd.notna(row["accuracy_m"]) else None,
            "altitude_m"    : float(row.get("altitude_m", 0) or 0),
            "land_use"      : str(row.get("land_use","")),
            "valid"         : bool(row["valid"]),
            "invalid_reason": "" if row["valid"] else (
                "Low GPS Accuracy" if pd.notna(row["accuracy_m"]) and row["accuracy_m"] >= ACCURACY_THRESH
                else "Out of Bounds" if not row["in_bbox"] else "Missing GPS"
            )
        }
    })

geojson = {"type": "FeatureCollection", "features": features}
with open("points_export.geojson", "w") as f:
    json.dump(geojson, f, indent=2)
print(f"   [DONE] points_export.geojson — {len(features)} features")

# ── 3. BUILD HTML MAPS USING FOLIUM ─────────────────────────────────────
try:
    import folium
    from folium.plugins import HeatMap, MarkerCluster
    FOLIUM = True
    print("\n  Generating interactive maps with folium...")
except ImportError:
    FOLIUM = False
    print("\n[WARN]  folium not installed — run: pip install folium")
    print("   Skipping HTML maps. GeoJSON still exported successfully.")

if FOLIUM:
    # ── Map A: All Points (valid = green, invalid = red, oob = orange) ──
    m1 = folium.Map(location=[center_lat, center_lon], zoom_start=13, tiles="OpenStreetMap")

    # Boundary box
    folium.Rectangle(
        bounds=[[BOUNDARY["lat_min"], BOUNDARY["lon_min"]],
                [BOUNDARY["lat_max"], BOUNDARY["lon_max"]]],
        color="blue", weight=2, fill=False, tooltip="Study Area Boundary"
    ).add_to(m1)

    cluster = MarkerCluster(name="All Points").add_to(m1)
    for _, row in df.iterrows():
        if row["valid"]:
            color = "green"; icon = "ok-circle"
        elif not row["in_bbox"]:
            color = "orange"; icon = "warning-sign"
        else:
            color = "red"; icon = "remove-circle"

        folium.Marker(
            location=[row["latitude"], row["longitude"]],
            icon=folium.Icon(color=color, icon=icon, prefix="glyphicon"),
            tooltip=(f"<b>{row.get('worker_name','')}</b><br>"
                     f"Date: {row.get('date','')}<br>"
                     f"Point: {row.get('point_number','')}<br>"
                     f"Accuracy: {row.get('accuracy_m','')}m<br>"
                     f"Land Use: {row.get('land_use','')}<br>"
                     f"Valid: {'[DONE]' if row['valid'] else '❌'}")
        ).add_to(cluster)

    # Legend
    legend_html = """
    <div style="position:fixed;bottom:30px;left:30px;z-index:1000;
                background:white;padding:12px;border:2px solid grey;border-radius:5px;">
    <b>Point Validity</b><br>
    🟢 Valid<br>🔴 Bad Accuracy<br>🟠 Out of Bounds
    </div>"""
    m1.get_root().html.add_child(folium.Element(legend_html))
    m1.save("map_all_points.html")
    print("   [DONE] map_all_points.html")

    # ── Map B: By Worker ────────────────────────────────────────────────
    m2 = folium.Map(location=[center_lat, center_lon], zoom_start=13)
    for wid in workers:
        wdf = df[df["worker_id"] == wid]
        wname = wdf["worker_name"].iloc[0] if len(wdf) else wid
        fg = folium.FeatureGroup(name=f"{wname} ({wid})")
        color = w_color[wid].lstrip("#")
        for _, row in wdf.iterrows():
            folium.CircleMarker(
                location=[row["latitude"], row["longitude"]],
                radius=5, color=f"#{color}", fill=True, fill_opacity=0.7,
                tooltip=(f"<b>{wname}</b><br>Point {row.get('point_number','')}<br>"
                         f"Date: {row.get('date','')}")
            ).add_to(fg)
        fg.add_to(m2)
    folium.LayerControl().add_to(m2)
    m2.save("map_by_worker.html")
    print("   [DONE] map_by_worker.html")

    # ── Map C: Heatmap ──────────────────────────────────────────────────
    m3 = folium.Map(location=[center_lat, center_lon], zoom_start=13)
    heat_data = df[df["valid"]][["latitude","longitude"]].values.tolist()
    HeatMap(heat_data, radius=15, blur=10, min_opacity=0.4).add_to(m3)
    folium.Marker(
        location=[center_lat, center_lon],
        icon=folium.DivIcon(html="<b style='color:blue'>📍 Study Area Center</b>")
    ).add_to(m3)
    m3.save("heatmap.html")
    print("   [DONE] heatmap.html")

# ── 4. COVERAGE STATS ────────────────────────────────────────────────────
print("\n Coverage statistics:")
valid_df = df[df["valid"]]
if len(valid_df) > 1:
    lat_range = valid_df["latitude"].max()  - valid_df["latitude"].min()
    lon_range = valid_df["longitude"].max() - valid_df["longitude"].min()
    approx_area_km2 = round(
        haversine_km(valid_df["latitude"].min(), valid_df["longitude"].min(),
                     valid_df["latitude"].min(), valid_df["longitude"].max()) *
        haversine_km(valid_df["latitude"].min(), valid_df["longitude"].min(),
                     valid_df["latitude"].max(), valid_df["longitude"].min()), 2)
    print(f"   Valid points:       {len(valid_df)}")
    print(f"   Lat range:          {round(lat_range*111, 2)} km")
    print(f"   Lon range:          {round(lon_range*111*np.cos(radians(center_lat)), 2)} km")
    print(f"   Approx study area:  ~{approx_area_km2} km²")

# Worker coverage
print(f"\n   Points per worker:")
for wid, wdf in df.groupby("worker_id"):
    wname = wdf["worker_name"].iloc[0]
    print(f"      {wname}: {len(wdf)} total | {wdf['valid'].sum()} valid ({round(wdf['valid'].mean()*100,1)}%)")

print(f"\n[DONE] DONE")
print(f"   GeoJSON  → points_export.geojson")
if FOLIUM:
    print(f"   Maps     → map_all_points.html | map_by_worker.html | heatmap.html")
print(f"\n   💡 Tip: Load points_export.geojson in QGIS via Layer > Add Layer > Add Vector Layer")
