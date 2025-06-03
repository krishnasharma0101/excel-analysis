# import streamlit as st
# import pandas as pd
# import requests
# import json
# import re
# import io
# import os
# import contextlib
# from dotenv import load_dotenv

# # Load environment variables
# load_dotenv()

# # OpenRouter API configuration
# OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
# API_URL = "https://openrouter.ai/api/v1/chat/completions"
# MODEL = "tngtech/deepseek-r1t-chimera:free"

# def process_with_gemini(prompt, data_description):
#     headers = {
#         "Authorization": f"Bearer {OPENROUTER_API_KEY}",
#         "Content-Type": "application/json"
#     }

#     data = {
#         "model": MODEL,
#         "messages": [
#             {
#                 "role": "user",
#                 "content": f"{prompt}\n\nData Description:\n{data_description}"
#             }
#         ]
#     }

#     try:
#         response = requests.post(API_URL, headers=headers, json=data)
#         response.raise_for_status()
#         content = response.json()["choices"][0]["message"]["content"]

#         # Extract Python code block
#         code_match = re.search(r"```(?:python)?\n(.*?)```", content, re.DOTALL)
#         if code_match:
#             return code_match.group(1).strip()
#         else:
#             return content.strip()
#     except Exception as e:
#         return f"Error processing request: {str(e)}"

# def execute_user_code(user_code, dfs):
#     output = io.StringIO()
#     local_vars = {"dfs": dfs, "pd": pd}

#     try:
#         with contextlib.redirect_stdout(output):
#             exec(user_code, {}, local_vars)
#     except Exception as e:
#         return f"Error during code execution: {str(e)}"

#     return output.getvalue() or "Code executed successfully. No output to show."

# def main():
#     st.title("Excel Data Analysis Dashboard")
#     st.write("Upload your Excel files and provide instructions for analysis")

#     # File upload
#     uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)

#     if uploaded_files:
#         dfs = {}
#         data_description = ""

#         # Read and describe uploaded files
#         for file in uploaded_files:
#             df = pd.read_excel(file)
#             dfs[file.name] = df
#             data_description += f"\nFile: {file.name}\n"
#             data_description += f"Columns: {', '.join(df.columns)}\n"
#             data_description += f"Number of rows: {len(df)}\n"
#             data_description += f"Sample data:\n{df.head().to_string(index=False)}\n"

#         # Preview uploaded files
#         st.subheader("Uploaded Files Preview")
#         for name, df in dfs.items():
#             st.write(f"**{name}**")
#             st.write(f"Shape: {df.shape}")
#             st.dataframe(df.head())

#         # User prompt
#         st.subheader("Analysis Instructions")
#         prompt = st.text_area("Enter your instructions for data analysis",
#                               placeholder="Example: Compare the data between files and find discrepancies in the 'amount' column")

#         # Process
#         if st.button("Process Data"):
#             if prompt:
#                 with st.spinner("Sending to Gemini and processing..."):
#                     code = process_with_gemini(prompt, data_description)
#                     st.subheader("Generated Python Code")
#                     st.code(code, language="python")

#                     result = execute_user_code(code, dfs)
#                     st.subheader("Execution Result")
#                     st.text(result)
#             else:
#                 st.warning("Please enter instructions for analysis")

# if __name__ == "__main__":
#     main()


# import streamlit as st
# import pandas as pd
# import requests
# import json
# import os
# import re
# import io
# import contextlib
# import time
# from dotenv import load_dotenv

# # Load environment variables
# load_dotenv()

# # OpenRouter API configuration
# OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
# API_URL = "https://openrouter.ai/api/v1/chat/completions"
# MODEL = "tngtech/deepseek-r1t-chimera:free"

# def process_with_gemini(prompt, data_description, retries=3, delay=5):
#     headers = {
#         "Authorization": f"Bearer {OPENROUTER_API_KEY}",
#         "Content-Type": "application/json"
#     }

#     data = {
#         "model": MODEL,
#         "messages": [
#             {
#                 "role": "user",
#                 "content": f"{prompt}\n\nData Description:\n{data_description}\nReturn Python code only in a single code block without extra explanation."
#             }
#         ]
#     }

#     for attempt in range(retries):
#         try:
#             response = requests.post(API_URL, headers=headers, json=data)
#             response.raise_for_status()
#             content = response.json()["choices"][0]["message"]["content"]

#             # Extract Python code block
#             code_match = re.search(r"```(?:python)?\n(.*?)```", content, re.DOTALL)
#             if code_match:
#                 return code_match.group(1).strip()
#             else:
#                 return content.strip()
#         except requests.exceptions.HTTPError as e:
#             if response.status_code == 429 and attempt < retries - 1:
#                 time.sleep(delay)
#                 continue
#             return f"Error processing request: {str(e)}"
#         except Exception as e:
#             return f"Error processing request: {str(e)}"

# def execute_user_code(user_code, dfs):
#     output = io.StringIO()
#     local_vars = {"dfs": dfs, "pd": pd}

#     try:
#         with contextlib.redirect_stdout(output):
#             exec(user_code, {}, local_vars)
#     except SyntaxError as e:
#         return f"Syntax Error: {e.text.strip()} on line {e.lineno}"
#     except Exception as e:
#         return f"Runtime Error: {str(e)}"

#     return output.getvalue() or "Code executed successfully. No output to show."

# def main():
#     st.title("Excel Data Analysis Dashboard")
#     st.write("Upload your Excel files and provide instructions for analysis")

#     uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)

#     if uploaded_files:
#         dfs = {}
#         data_description = ""

#         for file in uploaded_files:
#             df = pd.read_excel(file)
#             dfs[file.name] = df
#             data_description += f"\nFile: {file.name}\n"
#             data_description += f"Columns: {', '.join(df.columns)}\n"
#             data_description += f"Number of rows: {len(df)}\n"
#             data_description += f"Sample data:\n{df.head().to_string()}\n"

#         st.subheader("Uploaded Files Information")
#         for name, df in dfs.items():
#             st.write(f"**{name}**")
#             st.write(f"Shape: {df.shape}")
#             st.dataframe(df.head())

#         st.subheader("Analysis Instructions")
#         prompt = st.text_area("Enter your instructions for data analysis", 
#                               placeholder="Example: Compare DOBs by SSMID between the two files and find mismatches.")

#         if st.button("Process Data"):
#             if prompt:
#                 with st.spinner("Processing your request..."):
#                     code = process_with_gemini(prompt, data_description)

#                     # Replace pd.read_excel(...) with dfs['filename']
#                     for filename in dfs:
#                         pattern = f"pd.read_excel\\(['\"]{re.escape(filename)}['\"].*?\\)"
#                         replacement = f"dfs['{filename}']"
#                         code = re.sub(pattern, replacement, code)

#                     st.subheader("Generated Python Code")
#                     st.code(code, language='python')

#                     result = execute_user_code(code, dfs)

#                     st.subheader("Execution Result")
#                     st.text(result)
#             else:
#                 st.warning("Please enter instructions for analysis.")

# if __name__ == "__main__":
#     main()



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
        col1_select = st.selectbox("Column from file 1", df1.columns)
        col2_select = st.selectbox("Column from file 2", df2.columns)

        if st.button("🔍 Run Comparison"):
            merged = pd.merge(df1, df2, left_on=col1_select, right_on=col2_select, how="outer", suffixes=('_file1', '_file2'), indicator=True)
            differences = merged[merged['_merge'] != 'both']
            st.success(f"Found {len(differences)} difference(s).")
            st.dataframe(differences)

            if st.button("⬇️ Download Result"):
                file_name = st.text_input("Enter filename for download (without .xlsx)", value="comparison_result")
                if st.button("📥 Confirm Download"):
                    excel_bytes = to_excel(differences)
                    st.download_button(
                        label="Download Excel File",
                        data=excel_bytes,
                        file_name=f"{file_name}.xlsx",
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

