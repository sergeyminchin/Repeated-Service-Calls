
import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Repeated Service Calls Analyzer")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    if "מס' קריאה" in df.columns and "מספר מכשיר" in df.columns and "תאריך קריאה" in df.columns:
        df['תאריך קריאה'] = pd.to_datetime(df['תאריך קריאה'], errors='coerce')
        df = df.sort_values(by=['מספר מכשיר', 'תאריך קריאה'])

        df['Previous Date'] = df.groupby('מספר מכשיר')['תאריך קריאה'].shift(1)
        df['Days Since Last Call'] = (df['תאריך קריאה'] - df['Previous Date']).dt.days

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
        st.error("The file is missing one of the required columns: 'מס' קריאה', 'מספר מכשיר', or 'תאריך קריאה'")
