import streamlit as st
import pandas as pd
import io
import os

# Titel og beskrivelse
st.title("ZPP Skabelon-Udfylder")
st.write("Upload dine 4 datakilder, og appen vil udfylde standardskabelonen korrekt.")

# Indlæs standardskabelon direkte fra appens projektmappe
TEMPLATE_PATH = "ZPP_standard_template.xlsx"
if not os.path.exists(TEMPLATE_PATH):
    st.error("Standardskabelonen mangler i projektmappen. Tilføj filen og genstart appen.")
else:
    # Upload datafiler
    with st.sidebar:
        data1 = st.file_uploader("Upload fil 1 (DKK)", type=["xlsx", "csv"])
        data2 = st.file_uploader("Upload fil 2 (EUR)", type=["xlsx", "csv"])
        data3 = st.file_uploader("Upload fil 3 (SEK)", type=["xlsx", "csv"])
        data4 = st.file_uploader("Upload fil 4 (Landed)", type=["xlsx", "csv"])

    if data1 and data2 and data3 and data4:
        # Læs standardskabelon
        template_df = pd.read_excel(TEMPLATE_PATH)

        # Hjælpefunktion til at læse filer fleksibelt
        def read_file(file):
            if file.name.endswith(".csv"):
                return pd.read_csv(file)
            return pd.read_excel(file)

        df1 = read_file(data1)
        df2 = read_file(data2)
        df3 = read_file(data3)
        df4 = read_file(data4)

        # Tjek at nødvendige kolonner findes
        missing_cols = []
        for df, cols in [
            (df1, ["Style Name", "Wholesale Price DKK", "Recommended Retail Price DKK"]),
            (df2, ["Style Name", "Wholesale Price EUR", "Recommended Retail Price EUR"]),
            (df3, ["Style Name", "Wholesale Price SEK", "Recommended Retail Price SEK"]),
            (df4, ["Style Name", "Landed"])
        ]:
            for col in cols:
                if col not in df.columns:
                    missing_cols.append(col)

        if missing_cols:
            st.error(f"Følgende nødvendige kolonner mangler i dine uploads: {', '.join(missing_cols)}")
        else:
            # Fletning baseret på relevante kolonner
            merged = df2.copy()
            merged = merged.merge(df1[["Style Name", "Wholesale Price DKK", "Recommended Retail Price DKK"]], on="Style Name", how="left")
            merged = merged.merge(df3[["Style Name", "Wholesale Price SEK", "Recommended Retail Price SEK"]], on="Style Name", how="left")
            merged = merged.merge(df4[["Style Name", "Landed"]], on="Style Name", how="left")

            # Kopier de nødvendige kolonner ind i skabelonen
            relevant_cols = [
                "Style No", "Style Name", "Brand", "Type", "Category", "Quality", "Color",
                "Size", "Qty", "Barcode", "Weight", "Country", "Customs Tariff No", "Season",
                "Delivery", "Wholesale Price EUR", "Recommended Retail Price EUR",
                "Wholesale Price DKK", "Recommended Retail Price DKK",
                "Wholesale Price SEK", "Recommended Retail Price SEK"
            ]

            for col in relevant_cols:
                if col in merged.columns:
                    template_df[col] = merged[col]

            # Tilføj Landed til Costs DKK kolonne
            template_df["Costs DKK"] = merged["Landed"]

            # Formater barcode korrekt
            if "Barcode" in template_df.columns:
                template_df["Barcode"] = template_df["Barcode"].astype(str).str.replace(".0", "", regex=False)

            # Download-knap med ekstra ark
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                template_df.to_excel(writer, index=False, sheet_name="DKK")
                df2.to_excel(writer, index=False, sheet_name="EUR")
                df3.to_excel(writer, index=False, sheet_name="SEK")
                df4.to_excel(writer, index=False, sheet_name="Landed cost")
            towrite.seek(0)
            st.download_button("Download udfyldt skabelon", towrite, file_name="udfyldt_skabelon.xlsx")
