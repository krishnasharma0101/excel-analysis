# Excel Analysis Tool with AI Integration

A powerful Streamlit-based web application that combines Excel file analysis with AI capabilities using the OpenRouter API. This tool allows users to perform complex data analysis, comparisons, and transformations on Excel files with the help of AI assistance.

## 🌟 Features

### File Operations
- Upload multiple Excel files
- Support for standard Excel (.xlsx) format
- Automatic header detection and data loading

### AI-Powered Analysis
- Natural language processing for data operations
- Smart formula suggestions
- Intelligent data comparison
- Automated discrepancy detection
- AI-assisted formatting recommendations

### Data Analysis Tools
1. **Compare Columns**
   - Compare data between multiple Excel files
   - AI-assisted comparison suggestions
   - Manual column comparison with reference keys
   - Download comparison results

2. **Find Discrepancies**
   - Identify differences between datasets
   - AI-powered discrepancy detection
   - Detailed discrepancy reports

3. **Apply Formula**
   - AI Formula Assistant for natural language formula generation
   - Manual formula input
   - Real-time formula application
   - Support for complex pandas operations

4. **Format Column**
   - AI Formatting Assistant
   - Multiple formatting options:
     - Uppercase
     - Lowercase
     - Titlecase
     - Strip Spaces
   - Custom formatting through AI suggestions

5. **AI Analysis**
   - Natural language query processing
   - Custom analysis generation
   - Download analysis results

## 🚀 Live Demo

Try the application online: [Excel Analysis Tool](https://excel-analysis.streamlit.app/)

## 🛠️ Setup and Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-analysis-tool.git
cd excel-analysis-tool
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your OpenRouter API key:
```
OPENROUTER_API_KEY=your_api_key_here
```

4. Run the application:
```bash
streamlit run app.py
```

## 📋 Requirements

- Python 3.7+
- Streamlit
- Pandas
- OpenRouter API key
- Other dependencies listed in `requirements.txt`

## 💻 Usage

1. **Upload Files**
   - Click "Browse files" to upload Excel files
   - Support for multiple file uploads

2. **Choose Operation**
   - Select from available operations in the sidebar
   - Each operation has both AI-assisted and manual options

3. **AI Assistance**
   - Describe your task in natural language
   - Get AI-generated suggestions
   - Apply suggestions with one click

4. **Manual Operations**
   - Select files and columns
   - Apply formulas or formatting
   - Download results

## 🔒 Security

- API keys are stored in environment variables
- No data is stored on the server
- All processing is done in-memory

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👨‍💻 Developer

- **Krishna Sharma**
  - [Twitter](https://x.com/kkrishnnaaa01)
  - [LinkedIn](https://www.linkedin.com/in/krishna-sharma-7953b42a2)

## 🙏 Acknowledgments

- Streamlit for the amazing web framework
- OpenRouter for AI capabilities
- Pandas for powerful data manipulation 
