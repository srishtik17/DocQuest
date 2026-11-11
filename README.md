# DocQuest – From Documents to Decisions 📄📊

**DocQuest** is a beginner-built, AI-powered document assistant and data analytics platform developed using Python and Streamlit.

The project brings document interaction and data visualization together in one workspace. Users can upload documents, ask questions about their content, and explore structured data through interactive dashboards.

This is an early version of the project, built as a hands-on learning experience. More improvements and features are planned.

## 🚀 Features

### 1. Document Assistant
- Upload and interact with supported document formats.
- Ask questions about uploaded documents.
- Get context-aware answers using the Gemini API.
- Retrieve relevant document information using semantic search.

### 2. Data Analytics Workspace
- Upload CSV and Excel files.
- Explore and process structured data.
- Generate interactive dashboards and visualizations.
- View key metrics, summaries and data insights.

### 3. AI-Powered Document Search
- Generate text embeddings using Gemini.
- Use FAISS for efficient similarity search.
- Retrieve relevant content to support document-based answers.

## 🛠️ Technologies Used

| Technology | Purpose |
|---|---|
| Python | Core programming language |
| Streamlit | Application interface |
| Gemini API | AI-powered responses and embeddings |
| FAISS | Vector similarity search |
| Pandas | Data processing and analysis |
| Plotly | Interactive charts and dashboards |
| LangChain Text Splitters | Document text chunking |

## 📂 Project Structure

```text
DocQuest/
│
├── DocGenius/
│   ├── PDFChat_Combined.py
│   └── ...
│
├── Artifacts/
│
├── app.py
├── requirements.txt
└── README.md
```

*Note: The project structure may evolve as new features are added.*

## ⚙️ Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/srishtik17/DocQuest.git
cd DocQuest
```

### 2. Create a virtual environment (optional)

```bash
python -m venv venv
```

Activate it:

**Mac/Linux**
```bash
source venv/bin/activate
```

**Windows**
```bash
venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure your Gemini API key

Create a `.env` file in the project root:

```env
GEMINI_API=your_gemini_api_key
```

Replace the placeholder with your own API key.

**Important:** Never upload your `.env` file or expose your API key on GitHub.

### 5. Run the application

```bash
streamlit run DocGenius/PDFChat_Combined.py
```

Open the local URL displayed in your terminal.

## 🌐 Live Demo

[Try DocQuest](https://deploy-gc24z8ssappdn4ucs33vbl.streamlit.app/)

## 💻 GitHub Repository

[View DocQuest on GitHub](https://github.com/srishtik17/DocQuest)

## 📚 What I Learned

Building DocQuest helped me gain practical experience with:

- Integrating the Gemini API into a Python application.
- Understanding embeddings and semantic search.
- Working with FAISS for vector-based retrieval.
- Processing structured data using Pandas.
- Creating interactive visualizations using Plotly.
- Building an application interface with Streamlit.
- Handling files, debugging errors and deploying an application.

## 🔮 Future Improvements

DocQuest is currently an early version and a continuous learning project.

Some areas I would like to improve:

- Enhancing the user interface and user experience.
- Improving document processing and response quality.
- Adding more analytics and visualization options.
- Improving application performance and error handling.
- Introducing additional features based on testing and feedback.

## 👩‍💻 Author

**Srishti Kumari**

B.Tech Computer Science Student

[GitHub](https://github.com/srishtik17) · [LinkedIn](https://www.linkedin.com/in/srishti-kumari-749a91383/)

---

*Built with curiosity, learning and a willingness to improve.*
