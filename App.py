import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="ZPP Skabelon-Udfylder", layout="centered")

# Titel og beskrivelse
st.markdown("""
    <style>
        html, body, .main, .block-container {
            margin: 0 auto !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 1rem !important;
        }
        [data-testid="stFileUploader"] {
            padding: 0.5rem !important;
            margin-bottom: 0.2rem !important;
        }
        [data-testid="stFileUploader"] > div > div {
            height: 135px !important;
        }
        .title {
            text-align: center;
            font-size: 1.5em;
            font-weight: 700;
            margin: 0.2em 0 0.3em 0;
        }
        .subtitle {
            text-align: center;
            font-size: 0.9em;
            color: #444;
            margin-bottom: 0.6em;
        }
        h5 {
            margin: 0.6rem 0 0.3rem 0 !important;
        }
        hr {
            margin: 0.8rem 0 !important;
        }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>ZPP Skabelon-Udfylder</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Upload ordrebekræftelse og de 4 datakilder, så samles og udfyldes skabelonen automatisk.</div>", unsafe_allow_html=True)

TEMPLATE_PATH = "ZPP_standard_template.xlsx"
if not os.path.exists(TEMPLATE_PATH):
    st.error("Standardskabelonen mangler i projektmappen. Tilføj filen og genstart appen.")
else:
    st.markdown("<h5>📦 Ordrebekræftelse</h5>", unsafe_allow_html=True)
    order_file = st.file_uploader("Upload ordrebekræftelse (hovedark)", type=["xlsx", "csv"], key="order")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h5>📂 Datakilder</h5>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        data1 = st.file_uploader("Upload fil 1 (DKK)", type=["xlsx", "csv"], key="dkk")
        data3 = st.file_uploader("Upload fil 3 (SEK)", type=["xlsx", "csv"], key="sek")
    with col2:
        data2 = st.file_uploader("Upload fil 2 (EUR)", type=["xlsx", "csv"], key="eur")
        data4 = st.file_uploader("Upload fil 4 (Landed)", type=["xlsx", "csv"], key="landed")

    if order_file and data1 and data2 and data3 and data4:
        template_df = pd.read_excel(TEMPLATE_PATH)

        def read_file(file):
            return pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)

        df_order = read_file(order_file)
        df1 = read_file(data1)
        df2 = read_file(data2)
        df3 = read_file(data3)
        df4 = read_file(data4)

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
            merged = df2.copy()
            merged = merged.merge(df1[["Style Name", "Wholesale Price DKK", "Recommended Retail Price DKK"]], on="Style Name", how="left")
            merged = merged.merge(df3[["Style Name", "Wholesale Price SEK", "Recommended Retail Price SEK"]], on="Style Name", how="left")
            merged = merged.merge(df4[["Style Name", "Landed"]], on="Style Name", how="left")

            relevant_rows = df_order[["Style Name", "Barcode"]].drop_duplicates()
            merged = pd.merge(relevant_rows, merged, on=["Style Name", "Barcode"], how="left")
            merged = merged.drop_duplicates(subset=["Style Name", "Barcode"])

            final_df = pd.DataFrame(columns=template_df.columns)
            for col in final_df.columns:
                if col in merged.columns:
                    final_df[col] = merged[col]

            if "Costs DKK" in final_df.columns and "Landed" in merged.columns:
                final_df["Costs DKK"] = merged["Landed"]

            if "Barcode" in final_df.columns:
                final_df["Barcode"] = final_df["Barcode"].astype(str).str.replace(".0", "", regex=False)

            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name="Ordrebekræftelse")
                df2.to_excel(writer, index=False, sheet_name="EUR")
                df3.to_excel(writer, index=False, sheet_name="SEK")
                df4.to_excel(writer, index=False, sheet_name="Landed cost")
            towrite.seek(0)
            st.download_button("Download udfyldt skabelon", towrite, file_name="udfyldt_skabelon.xlsx")
