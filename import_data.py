import pandas as pd
import json

df = pd.read_excel("sites.xlsx")

data = []

for i, row in df.iterrows():
    try:
        item = {
            "id": int(i + 1),
            "center": str(row["Center"]).strip(),
            "location": str(row["Location"]).strip(),
            "W": int(row["W"]),
            "H": int(row["H"]),
            "type": str(row["Type"]).strip(),
            "sq_ft": int(row["Sq. Ft."]),
            "lat_long": str(row["Lat. & Long."]).strip(),
            "rate": str(row.get("Rate per month", "")).strip(),
            "image": f"images/{i+1}.jpg"
        }

        data.append(item)

    except Exception as e:
        print(f"Error in row {i+1}: {e}")

with open("data.json", "w") as f:
    json.dump(data, f, indent=2)

print(f"Imported {len(data)} sites successfully")
