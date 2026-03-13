import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Summary Tool")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("Preview Data")
    st.dataframe(df)

    if st.button("Generate Report"):

        result = (
            df.groupby("Product")["LDSO"]
            .agg(
                Count="count",
                LDSO_List=lambda x: ", ".join(x.astype(str))
            )
            .reset_index()
        )

        st.subheader("Result")
        st.dataframe(result)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, index=False)

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )