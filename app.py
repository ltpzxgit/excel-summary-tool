import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Jira LDSO Report Generator")

uploaded_file = st.file_uploader("Upload Jira Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("Preview Jira Data")
    st.dataframe(df)

    if st.button("Generate Report"):

        # -----------------------------
        # Sheet3 : Jira_LDSO (Raw data)
        # -----------------------------
        jira_ldso = df.copy()

        # -----------------------------
        # Sheet2 : Mapping
        # -----------------------------
        grouped = df.groupby("Service Name")["LDSO"].apply(list)

        rows = []

        for service, ldso_list in grouped.items():
            row = [service, len(ldso_list)] + ldso_list
            rows.append(row)

        max_len = max(len(r) for r in rows)

        for r in rows:
            r.extend([""] * (max_len - len(r)))

        mapping = pd.DataFrame(rows)

        # ตั้งชื่อ column
        mapping.columns = ["Service Name", "Total"] + [""]*(len(mapping.columns)-2)

        # -----------------------------
        # Sheet1 : Summary
        # -----------------------------
        summary = (
            df.groupby(["Type", "Rank", "Status"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )

        # -----------------------------
        # Export Excel
        # -----------------------------
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            summary.to_excel(writer, sheet_name="Summary", index=False)
            mapping.to_excel(writer, sheet_name="Mapping", index=False)
            jira_ldso.to_excel(writer, sheet_name="Jira_LDSO", index=False)

        st.success("Report generated")

        st.download_button(
            label="Download Report",
            data=output.getvalue(),
            file_name="Jira_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
