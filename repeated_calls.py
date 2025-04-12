
import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image

logo = Image.open("logo.png")
st.image(logo, use_container_width=False)


st.title("Repeated Service Calls Analyzer")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("Select Sheet", xls.sheet_names)
    df = xls.parse(sheet)

    st.write("Detected columns:", df.columns.tolist())

    # Flexible column name matching
    col_call_id = next((col for col in df.columns if "קריאה" in col and "מס" in col), None)
    col_device_id = next((col for col in df.columns if "מכשיר" in col), None)
    col_call_date = next((col for col in df.columns if "תאריך" in col and "קריאה" in col), None)

    if col_call_id and col_device_id and col_call_date:
        df[col_call_date] = pd.to_datetime(df[col_call_date], errors='coerce')
        df = df.sort_values(by=[col_device_id, col_call_date])

        df['Previous Date'] = df.groupby(col_device_id)[col_call_date].shift(1)
        df['Days Since Last Call'] = (df[col_call_date] - df['Previous Date']).dt.days

        repeated_df = df[df['Days Since Last Call'] <= 30].copy()

        st.subheader("Repeated Calls Detected (within 30 days for same device)")
        st.write(repeated_df)

        # Prepare download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            repeated_df.to_excel(writer, index=False, sheet_name='Repeated Calls')
        output.seek(0)

        st.download_button(
            label="Download Result as Excel",
            data=output,
            file_name="repeated_calls_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ One or more required columns were not found automatically.")
        st.markdown("Expected to find columns like:")
        st.code("מס' קריאה  |  מספר מכשיר  |  תאריך קריאה", language="markdown")
