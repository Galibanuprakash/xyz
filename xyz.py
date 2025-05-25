import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Coupling Selector Tool", layout="wide")
st.title("🔧 Coupling Selector Tool")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Main-Data", header=1)
    df.columns = df.columns.str.strip()  # Remove leading/trailing spaces

    # Sidebar Inputs
    st.sidebar.header("Filter Options")
    driver_input = st.sidebar.text_input("Driver")
    driven_input = st.sidebar.text_input("Driven")
    model_input = st.sidebar.text_input("Coupling Model")

    power_min = st.sidebar.number_input("Min Power (kW)", value=0.0)
    power_max = st.sidebar.number_input("Max Power (kW)", value=1e6)
    speed_min = st.sidebar.number_input("Min Speed (RPM)", value=0)
    speed_max = st.sidebar.number_input("Max Speed (RPM)", value=1e6)
    dbse_min = st.sidebar.number_input("Min DBSE/DBFF (mm)", value=0.0)
    dbse_max = st.sidebar.number_input("Max DBSE/DBFF (mm)", value=1e6)

    if st.sidebar.button("🔍 Search Couplings"):
        df_filtered = df.copy()

        # Case-insensitive exact matches
        if driver_input:
            df_filtered = df_filtered[df_filtered["Driver"].astype(str).str.lower() == driver_input.lower()]
        if driven_input:
            df_filtered = df_filtered[df_filtered["Driven"].astype(str).str.lower() == driven_input.lower()]
        if model_input:
            df_filtered = df_filtered[df_filtered["Coupling \nModel"].astype(str).str.lower() == model_input.lower()]

        # Convert range filters
        df_filtered["Power (kW)"] = pd.to_numeric(df_filtered["Power (kW)"], errors="coerce")
        df_filtered["Speed (RPM)"] = pd.to_numeric(df_filtered["Speed (RPM)"], errors="coerce")
        df_filtered["DBSE /DBFF (mm)"] = pd.to_numeric(df_filtered["DBSE /DBFF (mm)"], errors="coerce")

        df_filtered = df_filtered[
            df_filtered["Power (kW)"].between(power_min, power_max, inclusive="both") &
            df_filtered["Speed (RPM)"].between(speed_min, speed_max, inclusive="both") &
            df_filtered["DBSE /DBFF (mm)"].between(dbse_min, dbse_max, inclusive="both")
        ]

        # Output columns
        output_columns = [
            "Sl # / Couplig #", "OEM (Buyer)", "Drawing \nno", "Driver", "Driven",
            "Driver Coupling  \nType",
            "Driver Connection Type \n(taper/keyed/Angled/ counterbore /stepped /other)",
            "Driver - If keyed type, Single / double/taper ratio",
            "Driver End shaft dia", "Driver End hub Boss dia",
            "Driver Hub Pull-up distance (mm)", "Driver Shaft Juncture Capacity (kNm)",
            "Driver side Flange size- OD", "Driver side Flange size- PCD",
            "Driver side  Flange - Location size",
            "Driven coupling type",
            "Driven Connection Type (taper/keyed/Angled/ counterbore /stepped /other)",
            "Driven - If keyed type, Single / double/taper ratio",
            "Driven End shaft dia", "Driven Hub boss diameter",
            "Driven Hub Pull-up distance (mm)", "Driven Shaft Juncture Capacity (kNm)",
            "Driven side Flange size- OD", "Driven side Flange size- PCD",
            "Driven side  Flange - Location size",
            "Coupling \nModel", "PCD-1", "PCD-2",
            "Power (kW)", "Speed (RPM)",
            "Cyclic Torque requirement (yes / No)", "SCT (kNm)",
            "Torsional Stiffness (MNm/rad)", "DBSE /DBFF (mm)",
            "Total Weight\n(Kg)"
        ]

        # Filter only available columns
        output_columns = [col for col in output_columns if col in df_filtered.columns]
        df_result = df_filtered[output_columns].copy()

        # Replace dash and blank strings with NaN
        df_result.replace("-", pd.NA, inplace=True)
        df_result.replace("", pd.NA, inplace=True)

        # Drop rows where all values are missing (completely blank or all dashes)
        df_result.dropna(how='all', inplace=True)

        if df_result.empty:
            st.warning("❌ No matching records found.")
        else:
            st.success(f"✅ {len(df_result)} valid records found.")
            st.dataframe(df_result, use_container_width=True)

            # Prepare cleaned output
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name="Filtered")

            st.download_button(
                label="📥 Download Excel",
                data=output.getvalue(),
                file_name="cleaned_couplings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("⬆️ Upload an Excel file to get started.")
