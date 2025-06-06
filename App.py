import streamlit as st
import pandas as pd
import io
import os

# Titel og beskrivelse
st.title("ZPP Skabelon-Udfylder")
st.write("Upload ordrebekræftelse og de 4 datakilder, så samles og udfyldes skabelonen automatisk.")

# Indlæs standardskabelon direkte fra appens projektmappe
TEMPLATE_PATH = "ZPP_standard_template.xlsx"
if not os.path.exists(TEMPLATE_PATH):
    st.error("Standardskabelonen mangler i projektmappen. Tilføj filen og genstart appen.")
else:
    # Centreret layout for ordrebekræftelse
    st.markdown("<div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
    order_file = st.file_uploader("\n📦 Upload ordrebekræftelse (hovedark)", type=["xlsx", "csv"], key="order")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='max-width: 600px; margin: 0 auto;'>", unsafe_allow_html=True)
    
    data1 = st.file_uploader("Upload fil 1 (DKK)", type=["xlsx", "csv"], key="dkk")
    data2 = st.file_uploader("Upload fil 2 (EUR)", type=["xlsx", "csv"], key="eur")
    data3 = st.file_uploader("Upload fil 3 (SEK)", type=["xlsx", "csv"], key="sek")
    data4 = st.file_uploader("Upload fil 4 (Landed)", type=["xlsx", "csv"], key="landed")
    st.markdown("</div>", unsafe_allow_html=True)

    if order_file and data1 and data2 and data3 and data4:
        # Læs standardskabelon
        template_df = pd.read_excel(TEMPLATE_PATH)

        # Hjælpefunktion til at læse filer fleksibelt
        def read_file(file):
            if file.name.endswith(".csv"):
                return pd.read_csv(file)
            return pd.read_excel(file)

        df_order = read_file(order_file)
        df1 = read_file(data1)
        df2 = read_file(data2)
        df3 = read_file(data3)
        df4 = read_file(data4)

        # Tjek at nødvendige kolonner findes
        missing_cols = []
        for df, cols in [
            (df1, ["Style Name", "Wholesale Price DKK", "Recommended Retail Price DKK"]),
            (df2, ["Style Name", "Style No", "Brand", "Type", "Category", "Quality", "Color", "Size", "Qty", "Barcode", "Weight", "Country", "Customs Tariff No", "Season", "Delivery", "Wholesale Price EUR", "Recommended Retail Price EUR"]),
            (df3, ["Style Name", "Wholesale Price SEK", "Recommended Retail Price SEK"]),
            (df4, ["Style Name", "Landed"]),
            (df_order, ["Style Name", "Barcode"])
        ]:
            for col in cols:
                if col not in df.columns:
                    missing_cols.append(col)

        if missing_cols:
            st.error(f"Følgende nødvendige kolonner mangler i dine uploads: {', '.join(missing_cols)}")
        else:
            # Start med EUR-data som basis
            merged = df2.copy()
            merged = merged.merge(df1[["Style Name", "Wholesale Price DKK", "Recommended Retail Price DKK"]], on="Style Name", how="left")
            merged = merged.merge(df3[["Style Name", "Wholesale Price SEK", "Recommended Retail Price SEK"]], on="Style Name", how="left")
            merged = merged.merge(df4[["Style Name", "Landed"]], on="Style Name", how="left")

            # Filtrér kun rækker fra order-filen
            relevant_rows = df_order[["Style Name", "Barcode"]].drop_duplicates()
            merged = pd.merge(relevant_rows, merged, on=["Style Name", "Barcode"], how="left")

            # Fjern dubletter
            merged = merged.drop_duplicates(subset=["Style Name", "Barcode"])

            # Ny DataFrame med samme kolonner som templatestruktur
            final_df = pd.DataFrame(columns=template_df.columns)
            for col in final_df.columns:
                if col in merged.columns:
                    final_df[col] = merged[col]

            if "Costs DKK" in final_df.columns and "Landed" in merged.columns:
                final_df["Costs DKK"] = merged["Landed"]

            if "Barcode" in final_df.columns:
                final_df["Barcode"] = final_df["Barcode"].astype(str).str.replace(".0", "", regex=False)

            # Download-knap med ekstra ark
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name="Ordrebekræftelse")
                df2.to_excel(writer, index=False, sheet_name="EUR")
                df3.to_excel(writer, index=False, sheet_name="SEK")
                df4.to_excel(writer, index=False, sheet_name="Landed cost")
            towrite.seek(0)
            st.download_button("Download udfyldt skabelon", towrite, file_name="udfyldt_skabelon.xlsx")
