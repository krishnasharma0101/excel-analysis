import streamlit as st
import pandas as pd
from io import BytesIO
import requests
import json
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# OpenRouter API configuration
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
API_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "tngtech/deepseek-r1t-chimera:free"

st.set_page_config(page_title="Excel Analysis Tool", layout="wide")

def get_ai_suggestion(dfs, user_query, operation_type):
    # Prepare data description
    data_description = "DataFrame Information:\n"
    for name, df in dfs.items():
        # Convert column names to strings
        columns_str = ', '.join(str(col) for col in df.columns)
        data_description += f"""
        File: {name}
        Columns: {columns_str}
        Number of rows: {len(df)}
        Sample data:
        {df.head().to_string()}
        """
    
    prompt = f"""You are an expert in pandas and data analysis. Given the following DataFrame information and user query, 
    provide a pandas formula or analysis approach that accomplishes the task. Return ONLY the code without any explanation.
    
    Operation Type: {operation_type}
    User Query: {user_query}
    
    {data_description}
    
    Important: Use the exact file names from the data description in your code. For example, if the files are named 'file1.xlsx' and 'file2.xlsx',
    use dfs['file1.xlsx'] and dfs['file2.xlsx'] in your code.
    
    Return the code in this format:
    For single DataFrame operations:
    df['new_column'] = formula
    
    For multiple DataFrame operations:
    result = pd.merge(...)  # Use the exact file names from dfs dictionary
    """
    
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": MODEL,
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ]
    }
    
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"Error getting AI suggestion: {str(e)}"

# Helper functions
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

# UI - Upload files
st.title("📊 Excel Analysis Tool")

# Sidebar
with st.sidebar:
    st.header("📁 File Upload")
    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
    
    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded")
    
    st.header("🔧 Operations")
    operation = st.radio("Choose an operation", [
        "Compare Columns",
        "Find Discrepancies",
        "Apply Formula",
        "Format Column",
        "AI Analysis"
    ])

# Main content
if uploaded_files and len(uploaded_files) >= 2:
    dfs = {}
    for file in uploaded_files:
        raw_df = pd.read_excel(file, header=None)
        headers = raw_df.iloc[2]  # Assuming 3rd row is header
        df = pd.read_excel(file, skiprows=3)
        df.columns = headers
        dfs[file.name] = df

    sheet_names = list(dfs.keys())

    if operation in ["Compare Columns", "Find Discrepancies"]:
        st.header("🔍 Column Comparison")
        
        # AI Assistant for comparison
        st.subheader("🤖 AI Assistant")
        comparison_query = st.text_area("Describe what you want to compare", 
                                      placeholder="Example: Compare student names and their grades between two files")
        
        # Store AI suggestion in session state
        if 'ai_suggestion' not in st.session_state:
            st.session_state.ai_suggestion = None
        if 'comparison_result' not in st.session_state:
            st.session_state.comparison_result = None
        
        if st.button("Get AI Suggestion"):
            if comparison_query:
                with st.spinner("Getting AI suggestion..."):
                    st.session_state.ai_suggestion = get_ai_suggestion(dfs, comparison_query, "comparison")
        
        # Display and apply suggestion
        if st.session_state.ai_suggestion:
            st.write("#### Suggested Code:")
            st.code(st.session_state.ai_suggestion, language="python")
            
            if st.button("Apply Suggestion"):
                try:
                    # Create a new DataFrame for the result
                    result = pd.DataFrame()
                    # Execute the suggestion
                    exec(st.session_state.ai_suggestion, {"dfs": dfs, "pd": pd, "result": result})
                    st.session_state.comparison_result = result
                    st.success("Analysis completed!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        # Display results
        if st.session_state.comparison_result is not None:
            st.write("#### Analysis Results:")
            st.dataframe(st.session_state.comparison_result)
            
            # Download button for results
            excel_bytes = to_excel(st.session_state.comparison_result)
            st.download_button(
                label="📥 Download Results",
                data=excel_bytes,
                file_name="ai_analysis_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Manual comparison
        st.subheader("Manual Comparison")
        col1, col2 = st.columns(2)
        
        with col1:
            file1 = st.selectbox("Select file 1", sheet_names)
            df1 = dfs[file1]
            ref_col1 = st.selectbox("Reference column from file 1", df1.columns)
            col1_select = st.selectbox("Column to compare from file 1", df1.columns)
        
        with col2:
            file2 = st.selectbox("Select file 2", sheet_names)
            df2 = dfs[file2]
            ref_col2 = st.selectbox("Reference column from file 2", df2.columns)
            col2_select = st.selectbox("Column to compare from file 2", df2.columns)

        if st.button("Run Comparison"):
            merged = pd.merge(
                df1, 
                df2, 
                left_on=ref_col1, 
                right_on=ref_col2, 
                how="outer", 
                suffixes=('_file1', '_file2')
            )
            
            comparison_df = pd.DataFrame({
                'Reference_Value': merged[ref_col1],
                f'{col1_select}_file1': merged[f'{col1_select}_file1'],
                f'{col2_select}_file2': merged[f'{col2_select}_file2'],
                'Match': merged[f'{col1_select}_file1'] == merged[f'{col2_select}_file2']
            })
            
            differences = comparison_df[~comparison_df['Match']]
            
            st.success(f"Found {len(differences)} difference(s).")
            st.dataframe(comparison_df)
            
            excel_bytes = to_excel(comparison_df)
            st.download_button(
                label="📥 Download Comparison Results",
                data=excel_bytes,
                file_name="comparison_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif operation == "Apply Formula":
        st.header("📝 Formula Application")
        
        # AI Formula Assistant
        st.subheader("🤖 AI Formula Assistant")
        formula_query = st.text_area("Describe what you want to do with the data", 
                                   placeholder="Example: Calculate the total price by multiplying quantity and price columns")
        
        # Store AI suggestion in session state
        if 'formula_suggestion' not in st.session_state:
            st.session_state.formula_suggestion = None
        
        if st.button("Get AI Formula"):
            if formula_query:
                with st.spinner("Getting AI suggestion..."):
                    st.session_state.formula_suggestion = get_ai_suggestion(dfs, formula_query, "formula")
        
        # Display and apply suggestion
        if st.session_state.formula_suggestion:
            st.write("#### Suggested Formula:")
            st.code(st.session_state.formula_suggestion, language="python")
            
            if st.button("Apply Formula"):
                try:
                    exec(st.session_state.formula_suggestion, {"dfs": dfs, "pd": pd})
                    st.success("Formula applied successfully!")
                    # Display the updated DataFrame
                    for name, df in dfs.items():
                        st.write(f"#### Updated {name}")
                        st.dataframe(df)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        # Manual Formula Input
        st.subheader("Manual Formula Input")
        selected_file = st.selectbox("Select file", sheet_names)
        df = dfs[selected_file]
        
        st.write("#### Current DataFrame Information")
        st.write(f"Columns: {', '.join(str(col) for col in df.columns)}")
        st.write(f"Number of rows: {len(df)}")
        st.write("Sample data:")
        st.dataframe(df.head())
        
        formula = st.text_input("Enter your pandas formula")
        if st.button("Apply Manual Formula"):
            try:
                exec(formula, {"df": df, "pd": pd})
                st.success("Formula applied!")
                st.dataframe(df)
                dfs[selected_file] = df
            except Exception as e:
                st.error(f"Error: {e}")

    elif operation == "Format Column":
        st.header("🧹 Column Formatting")
        
        # AI Formatting Assistant
        st.subheader("🤖 AI Formatting Assistant")
        format_query = st.text_area("Describe how you want to format the data", 
                                  placeholder="Example: Convert all names to title case and remove extra spaces")
        
        # Store AI suggestion in session state
        if 'format_suggestion' not in st.session_state:
            st.session_state.format_suggestion = None
        
        if st.button("Get Formatting Suggestion"):
            if format_query:
                with st.spinner("Getting AI suggestion..."):
                    st.session_state.format_suggestion = get_ai_suggestion(dfs, format_query, "formatting")
        
        # Display and apply suggestion
        if st.session_state.format_suggestion:
            st.write("#### Suggested Formatting:")
            st.code(st.session_state.format_suggestion, language="python")
            
            if st.button("Apply Formatting"):
                try:
                    exec(st.session_state.format_suggestion, {"dfs": dfs, "pd": pd})
                    st.success("Formatting applied successfully!")
                    # Display the updated DataFrame
                    for name, df in dfs.items():
                        st.write(f"#### Updated {name}")
                        st.dataframe(df)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        # Manual Formatting
        st.subheader("Manual Formatting")
        selected_file = st.selectbox("Select file", sheet_names)
        df = dfs[selected_file]
        
        col = st.selectbox("Select column to format", df.columns)
        fmt_option = st.selectbox("Formatting Option", ["Uppercase", "Lowercase", "Titlecase", "Strip Spaces"])
        
        if st.button("Apply Format"):
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

    elif operation == "AI Analysis":
        st.header("🤖 AI Analysis")
        
        analysis_query = st.text_area("Describe what analysis you want to perform", 
                                    placeholder="Example: Find all students with grades above 90 and calculate their average age")
        
        # Store AI suggestion in session state
        if 'analysis_suggestion' not in st.session_state:
            st.session_state.analysis_suggestion = None
        if 'analysis_result' not in st.session_state:
            st.session_state.analysis_result = None
        
        if st.button("Run AI Analysis"):
            if analysis_query:
                with st.spinner("Performing AI analysis..."):
                    st.session_state.analysis_suggestion = get_ai_suggestion(dfs, analysis_query, "analysis")
        
        # Display and apply suggestion
        if st.session_state.analysis_suggestion:
            st.write("#### Analysis Code:")
            st.code(st.session_state.analysis_suggestion, language="python")
            
            if st.button("Execute Analysis"):
                try:
                    # Create a new DataFrame for the result
                    result = pd.DataFrame()
                    # Execute the suggestion
                    exec(st.session_state.analysis_suggestion, {"dfs": dfs, "pd": pd, "result": result})
                    st.session_state.analysis_result = result
                    st.success("Analysis completed successfully!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        # Display results
        if st.session_state.analysis_result is not None:
            st.write("#### Analysis Results:")
            st.dataframe(st.session_state.analysis_result)
            
            # Download button for results
            excel_bytes = to_excel(st.session_state.analysis_result)
            st.download_button(
                label="📥 Download Results",
                data=excel_bytes,
                file_name="ai_analysis_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.warning("Please upload at least 2 Excel files.")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <p>Developed by Krishna Sharma</p>
    <p>
        <a href="https://x.com/kkrishnnaaa01" target="_blank">Twitter</a> | 
        <a href="https://www.linkedin.com/in/krishna-sharma-7953b42a2" target="_blank">LinkedIn</a>
    </p>
</div>
""", unsafe_allow_html=True)

