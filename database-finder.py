import streamlit as st
import pandas as pd
import tabulate  # required for pandas.DataFrame.to_markdown()
import streamlit.components.v1 as components
from openpyxl import load_workbook


# Load the Excel file
df = pd.read_excel("data/database_checklist.xlsx", header=None)

# Extract headers and data ranges
content_types = df.iloc[0, 1:].dropna().tolist()
data_start_row = 2

# Streamlit UI
st.title("Business Database Finder")
# st.markdown("Select content types from the drop down menu below.")

# Label and dropdown on the same line to avoid extra padding
st.subheader("Select content types:", divider=False, width="stretch")
# st.markdown("**⬇️ Choose one or more content types from the list below:**")
selected_types = st.multiselect(
    "Select content types", content_types, label_visibility="collapsed"
)

if selected_types:
    # Match columns for selected content types
    matching_cols = [
        df.columns[df.iloc[0] == col].tolist()[0] for col in selected_types
    ]

    # Extract names and hyperlinks from Excel directly
    wb = load_workbook("data/database_checklist.xlsx")
    ws = wb.active

    names = []
    urls = []

    for row in ws.iter_rows(min_row=3, min_col=1, max_col=1):
        cell = row[0]
        names.append(cell.value)
        if cell.hyperlink:
            urls.append(cell.hyperlink.target)
        else:
            urls.append("")

    content_data = df.iloc[data_start_row:, matching_cols].copy()
    content_data.columns = selected_types
    result = pd.DataFrame({"Database": names, "URL": urls})
    result = pd.concat([result, content_data.reset_index(drop=True)], axis=1)

    # Filter for rows that match all selected types
    for col in selected_types:
        if col == "OTHER":
            result = result[
                result[col].notna() & result[col].astype(str).str.strip().ne("")
            ]
        else:
            result = result[
                result[col].notna()
                & result[col].astype(str).str.strip().str.lower().eq("y")
            ]
    result = result.reset_index(drop=True)
    result = result.apply(
        lambda col: col.map(
            lambda x: "✅" if isinstance(x, str) and x.strip().lower() == "y" else x
        )
    )

    result["Database"] = [
        (
            f"[{name}]({url})"
            if isinstance(url, str) and ("http" in url or "www." in url)
            else name
        )
        for name, url in zip(result["Database"], result["URL"])
    ]
    result.drop(columns=["URL"], inplace=True)

    # Display output with proper formatting using markdown to render links
    st.write(
        f"### Databases matching **all** selected content types ({len(result)} found):"
    )
    st.markdown(result.to_markdown(index=False), unsafe_allow_html=True)
else:
    st.info("⬆️ Select your content types above.")
