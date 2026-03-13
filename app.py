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

        # Group data
        grouped = df.groupby("Product")["LDSO"].apply(list)

        rows = []

        for product, ldso_list in grouped.items():
            row = [product, len(ldso_list)] + ldso_list
            rows.append(row)

        max_len = max(len(r) for r in rows)

        for r in rows:
            r.extend([""] * (max_len - len(r)))

        columns = ["Product", "Count"] + [f"LDSO_{i}" for i in range(1, max_len-1)]

        summary = pd.DataFrame(rows, columns=columns)

        detail = df.copy()

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            summary.to_excel(writer, sheet_name="Summary", index=False)
            detail.to_excel(writer, sheet_name="Detail", index=False)

        st.subheader("Summary")
        st.dataframe(summary)

        st.download_button(
            label="Download Excel Report",
            data=output.getvalue(),
            file_name="report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
