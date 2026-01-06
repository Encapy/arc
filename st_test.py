import streamlit as st
import pandas as pd
from io import BytesIO
import re

df_final = pd.read_excel("crafting_table.xlsx")

st.title("🧮 Crafting Material Summierer")

# ---------------------------------------------------------
# 1) Basisnamen extrahieren (Stitcher I → Stitcher)
# ---------------------------------------------------------
def get_base_name(name):
    return re.sub(r"\s+[IVX]+$", "", str(name)).strip()

df_final["BaseName"] = df_final["Name"].apply(get_base_name)

# ---------------------------------------------------------
# 2) Alle Versionen pro Basisnamen sammeln
#    z.B. "Stitcher": ["Stitcher I", "Stitcher II", ...]
# ---------------------------------------------------------
grouped_versions = (
    df_final.groupby("BaseName")["Name"]
    .apply(list)
    .to_dict()
)

# ---------------------------------------------------------
# 3) Dropdown zeigt NUR Basisnamen
# ---------------------------------------------------------
base_item_names = sorted(grouped_versions.keys())

selected_items = st.multiselect("Wähle Items zum Craften:", base_item_names)

# ---------------------------------------------------------
# 4) Anzahl pro Basisitem
# ---------------------------------------------------------
item_counts = {}
for item in selected_items:
    count = st.number_input(f"Anzahl für '{item}':", min_value=1, value=1)
    item_counts[item] = count

material_totals = None

# ---------------------------------------------------------
# 5) Rekursive Materialberechnung
# ---------------------------------------------------------
if item_counts:

    # Materialspalten bestimmen
    material_cols = [c for c in df_final.columns if c not in ["Name", "Station", "BaseName"]]

    # Leere Summe starten
    total_materials = pd.Series(0, index=material_cols, dtype=float)

    for base_item, count in item_counts.items():

        versions = grouped_versions[base_item]  # alle Versionen I–IV

        df_versions = df_final[df_final["Name"].isin(versions)].copy()
        df_versions = df_versions.set_index("Name")[material_cols]

        # Summiere alle Versionen
        summed = df_versions.sum()

        # Multipliziere mit gewünschter Anzahl
        total_materials += summed * count

    # Nur > 0 anzeigen
    material_totals = total_materials[total_materials > 0].sort_index()

    # ---------------------------------------------------------
    # Entferne alle Item-Namen aus der Materialliste
    # (z. B. Stitcher I, Stitcher II, ...)
    # ---------------------------------------------------------
    all_item_names = df_final["Name"].unique().tolist()
    material_totals = material_totals.drop(labels=[m for m in material_totals.index if m in all_item_names], errors="ignore")


    st.subheader("📦 Gesamtmaterialbedarf (inkl. Upgrade-Logik)")
    st.dataframe(
        material_totals.reset_index().rename(columns={"index": "Material", 0: "Anzahl"})
    )
else:
    st.info("Bitte wähle mindestens ein Item aus.")

# ---------------------------------------------------------
# 6) Export-Funktion
# ---------------------------------------------------------
if material_totals is not None and not material_totals.empty:
    einkaufsliste = material_totals.reset_index()
    einkaufsliste.columns = ['Material', 'Anzahl']
    einkaufsliste = einkaufsliste.sort_values('Material')

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        einkaufsliste.to_excel(writer, index=False, sheet_name='Materialliste')
    buffer.seek(0)

    st.download_button(
        label="📥 Export Resources (Excel)",
        data=buffer,
        file_name="Liste_materialien.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
