import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel AI Dashboard", layout="wide")

# Helper functions
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

# UI - Upload files
st.title("📊 AI-Enhanced Excel Sheet Dashboard")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and len(uploaded_files) >= 2:
    st.success(f"{len(uploaded_files)} file(s) uploaded")

    dfs = {}
    for file in uploaded_files:
        raw_df = pd.read_excel(file, header=None)
        headers = raw_df.iloc[2]  # Assuming 3rd row is header
        df = pd.read_excel(file, skiprows=3)
        df.columns = headers
        dfs[file.name] = df

    sheet_names = list(dfs.keys())

    st.subheader("🔧 Select Operation")
    operation = st.selectbox("Choose an operation", [
        "Compare Columns", "Find Discrepancies", "Apply Formula", "Format Column"
    ])

    if operation in ["Compare Columns", "Find Discrepancies"]:
        col1 = st.selectbox("Select file 1", sheet_names)
        col2 = st.selectbox("Select file 2", sheet_names)

        df1 = dfs[col1]
        df2 = dfs[col2]

        st.write("### Select Columns to Compare")
        
        # Reference column selection (common identifier)
        st.write("#### Select Reference Column (Common Identifier)")
        ref_col1 = st.selectbox("Reference column from file 1", df1.columns)
        ref_col2 = st.selectbox("Reference column from file 2", df2.columns)
        
        # Columns to compare
        st.write("#### Select Columns to Compare")
        col1_select = st.selectbox("Column to compare from file 1", df1.columns)
        col2_select = st.selectbox("Column to compare from file 2", df2.columns)

        if st.button("🔍 Run Comparison"):
            # Merge on reference columns
            merged = pd.merge(
                df1, 
                df2, 
                left_on=ref_col1, 
                right_on=ref_col2, 
                how="outer", 
                suffixes=('_file1', '_file2')
            )
            
            # Create comparison result
            comparison_df = pd.DataFrame({
                'Reference_Value': merged[ref_col1],
                f'{col1_select}_file1': merged[f'{col1_select}_file1'],
                f'{col2_select}_file2': merged[f'{col2_select}_file2'],
                'Match': merged[f'{col1_select}_file1'] == merged[f'{col2_select}_file2']
            })
            
            # Find differences
            differences = comparison_df[~comparison_df['Match']]
            
            st.success(f"Found {len(differences)} difference(s).")
            st.dataframe(comparison_df)
            
            # Download button
            excel_bytes = to_excel(comparison_df)
            st.download_button(
                label="📥 Download Comparison Results",
                data=excel_bytes,
                file_name="comparison_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif operation == "Apply Formula":
        selected_file = st.selectbox("Select file", sheet_names)
        df = dfs[selected_file]

        st.write("### Enter a pandas-compatible formula:")
        formula = st.text_input("Example: df['Total'] = df['Price'] * df['Quantity']")

        if st.button("⚙️ Apply Formula"):
            try:
                exec(formula, {"df": df})
                st.success("Formula applied!")
                st.dataframe(df)
                dfs[selected_file] = df  # Update in memory
            except Exception as e:
                st.error(f"Error: {e}")

    elif operation == "Format Column":
        selected_file = st.selectbox("Select file", sheet_names)
        df = dfs[selected_file]

        col = st.selectbox("Select column to format", df.columns)
        fmt_option = st.selectbox("Formatting Option", ["Uppercase", "Lowercase", "Titlecase", "Strip Spaces"])

        if st.button("🧼 Format Column"):
            if fmt_option == "Uppercase":
                df[col] = df[col].astype(str).str.upper()
            elif fmt_option == "Lowercase":
                df[col] = df[col].astype(str).str.lower()
            elif fmt_option == "Titlecase":
                df[col] = df[col].astype(str).str.title()
            elif fmt_option == "Strip Spaces":
                df[col] = df[col].astype(str).str.strip()

            st.dataframe(df)
            dfs[selected_file] = df

else:
    st.warning("Please upload at least 2 Excel files.")

