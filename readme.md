# PDF to MS Word Replica Converter

A professional Flask-based web application designed to replicate the layout, structure, and formatting of a specific PDF document into a perfectly readable and editable MS Word (.docx) file.

## 🚀 Live Demo
**URL:** [https://pdf-to-docx-converter-0brd.onrender.com](https://pdf-to-docx-converter-0brd.onrender.com)

## 🎯 Task Objective
The goal was to create an exact replica of a provided PDF, ensuring that spacing, alignment, line breaks, headings, and overall table structure were maintained using Python.

## ✨ Key Features
* **Exact Structural Replication:** Uses a custom mapping logic to recreate complex tables and nested headers.
* **Advanced Formatting:** Implemented custom XML (`OxmlElement`) to handle specific table borders and cell alignments that standard libraries don't support out-of-the-box.
* **In-Memory Processing:** Utilizes `io.BytesIO` for handling file uploads and downloads, ensuring no temporary files are stored on the server (Stateless Architecture).
* **Hyperlink Detection:** Automatically converts email strings into functional `mailto:` hyperlinks in the generated document.
* **Responsive UI:** A simple, clean interface for users to upload files and receive results instantly.



## 🛠️ Tech Stack
* **Language:** Python 3.x
* **Framework:** Flask
* **Libraries:** * `python-docx`: For professional document generation and XML manipulation.
    * `pypdf`: For precise text extraction from source PDFs.
* **Deployment:** Render (with Gunicorn for production-grade performance).

## 📖 My Approach
1.  **Parsing:** Analyzed the PDF structure to identify key data points. Used line-by-line extraction with `pypdf` to map data into specific document segments.
2.  **Document Engineering:** * Created a "Helper Function" architecture to keep code clean and reusable (e.g., `set_border`, `write_label`, `write_value`).
    * Manually calculated column widths (`Inches(0.35)`, `Inches(1.8)`, etc.) to match the visual proportions of the original PDF.
3.  **Deployment Logic:** Configured a `Procfile` and `requirements.txt` to ensure the application scales correctly in a cloud environment.

## 📂 Project Structure
```text
├── app.py              # Core logic, PDF parsing, and DOCX generation
├── requirements.txt    # List of dependencies for deployment
├── Procfile            # Instructions for the Render/Heroku web server
├── templates/
│   └── index.html      # Frontend HTML for file uploading
└── README.md           # Project documentation



##🛠️ Local Setup
```text
Clone the repo:

Bash

git clone [https://github.com/Omkar1979/PDF_TO_DOCX-Converter.git](https://github.com/Omkar1979/PDF_TO_DOCX-Converter.git)
cd PDF_TO_DOCX-Converter
Install requirements:

Bash

pip install -r requirements.txt
Run locally:

Bash

python app.py