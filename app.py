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

        # Sheet 1 : Summary
        summary = (
            df.groupby("Product")["LDSO"]
            .agg(
                Count="count",
                LDSO_List=lambda x: ", ".join(x.astype(str))
            )
            .reset_index()
        )

        # Sheet 2 : Detail (ข้อมูล original)
        detail = df.copy()

        # Sheet 3 : LDSO Horizontal
        horizontal_data = []

        for product, group in df.groupby("Product"):
            row = [product] + list(group["LDSO"])
            horizontal_data.append(row)

        max_len = max(len(r) for r in horizontal_data)

        for r in horizontal_data:
            r.extend([""] * (max_len - len(r)))

        columns = ["Product"] + [f"LDSO_{i}" for i in range(1, max_len)]

        horizontal = pd.DataFrame(horizontal_data, columns=columns)

        st.subheader("Summary")
        st.dataframe(summary)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            summary.to_excel(writer, sheet_name="Summary", index=False)
            detail.to_excel(writer, sheet_name="Detail", index=False)
            horizontal.to_excel(writer, sheet_name="LDSO_Horizontal", index=False)

        st.download_button(
            label="Download Excel Report",
            data=output.getvalue(),
            file_name="report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
