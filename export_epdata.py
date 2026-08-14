"""
Export EPDataSample.xlsx (MasterList sheet) -> EPdata.csv for the TorchStone
Executive Protection dashboard.

Follows the same conventions as the SEA dashboard's PHLdata export pipeline:
- openpyxl in read_only=True, data_only=True mode (cached values only, never
  recompute / never write back to the source workbook)
- UTF-8 with BOM (utf-8-sig) for safe handling in both Excel and the browser
- Dates formatted as 'DD Month YYYY' via .strftime('%d %B %Y')
- A derived 'Location' column (City, State/Province, Country - empty parts
  dropped) for consistent table/popup display, analogous to how SEA derives
  'Actor Categories' from 'Actors' during export.

Output columns (11):
Date, Location, City, Country, Domestic/Int'l, Latitude, Longitude,
Target, Tactic, Venue, Content, Source
"""
import csv
import openpyxl

SRC = "/mnt/user-data/uploads/EPDataSample.xlsx"
OUT = "/home/claude/torchstone/EPdata.csv"

wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
ws = wb["MasterList"]
rows = list(ws.iter_rows(values_only=True))
header = rows[0]
idx = {h: i for i, h in enumerate(header) if h}

def cell(row, name):
    v = row[idx[name]]
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return v

out_rows = []
for row in rows[1:]:
    if row[idx["Date"]] is None:
        continue  # skip trailing empty rows

    date_val = cell(row, "Date")
    date_str = date_val.strftime("%d %B %Y") if hasattr(date_val, "strftime") else str(date_val)

    city = cell(row, "City")
    state = cell(row, "State/Province")
    country = cell(row, "Country")
    location_parts = [p for p in [city, state, country] if p]
    location = ", ".join(location_parts)

    out_rows.append({
        "Date": date_str,
        "Location": location,
        "City": city,
        "Country": country,
        "Domestic/Int'l": cell(row, "Domestic/Int'l"),
        "Latitude": cell(row, "Latitude"),
        "Longitude": cell(row, "Longitude"),
        "Target": cell(row, "Target"),
        "Tactic": cell(row, "Tactic"),
        "Venue": cell(row, "Venue"),
        "Content": cell(row, "Content"),
        "Source": cell(row, "Source"),
    })

fieldnames = ["Date", "Location", "City", "Country", "Domestic/Int'l",
              "Latitude", "Longitude", "Target", "Tactic", "Venue", "Content", "Source"]

with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(out_rows)

print(f"Wrote {len(out_rows)} rows to {OUT}")
