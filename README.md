# Excel Data Analysis Dashboard

This Streamlit application allows you to upload multiple Excel files and analyze them using the Gemini model through OpenRouter API.

## Features

- Upload multiple Excel files (.xlsx, .xls)
- View file previews and basic information
- Provide custom instructions for data analysis
- Process data using Gemini model
- Get formatted analysis results

## Setup

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Create a `.env` file in the project root and add your OpenRouter API key:
```
OPENROUTER_API_KEY=your_api_key_here
```

3. Run the Streamlit app:
```bash
streamlit run app.py
```

## Usage

1. Launch the application using the command above
2. Upload one or more Excel files using the file uploader
3. Review the file previews and information
4. Enter your analysis instructions in the text area
5. Click "Process Data" to get the analysis results

## Example Instructions

- "Compare the data between files and find discrepancies in the 'amount' column"
- "Format the date columns in all files to YYYY-MM-DD"
- "Find duplicate entries across all files"
- "Calculate the sum of 'sales' column for each file"

## Note

Make sure your Excel files are properly formatted and contain the expected data structure for accurate analysis. 