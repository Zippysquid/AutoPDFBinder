📄 Instant Print-Ready Legal PDF Bundler

This script **automates document bundling** by converting an entire directory of Word and PDF files into a **single, print-ready PDF**—fully formatted with:

- ✅ **Table of contents** (auto-numbered with clickable links)
- ✅ **Cover pages for each document**
- ✅ **Bates numbering** for easy referencing
- ✅ **Automatic bookmarks for quick navigation**
- ✅ **Consistent, professional formatting**

Ideal for **lawyers, paralegals, and legal assistants — saving hours of manual work! 🚀


---

## 🛠 Features

- 📂 **Batch Processing** – Processes an entire directory in one go
- 📝 **Word to PDF Conversion** – Automatically converts DOCX files to PDFs
- 📑 **Table of Contents** – Generates a structured, clickable TOC with numbered pages
- 🔢 **Bates Numbering** – Applies sequential page numbering for legal documents
- 🔖 **PDF Bookmarks** – Adds easy navigation points to the final document
- 📏 **Consistent Formatting** – Ensures a clean, professional look with standard margins & fonts
- 🖨 **Ready to Print** – Outputs a fully formatted PDF that requires no further adjustments

---

---

## 🚀 Installation

### **Download the Script**

1. Clone or download this repository to a folder of your choice (e.g., `C:\AutoPDFBinder`).

### **Requirements**

Ensure you have **Python 3.8+** installed along with the necessary dependencies.

```sh
pip install -r requirements.txt
```

If you are running this on Windows, make sure **Microsoft Word is installed** (for DOCX to PDF conversion).

---

## ▶️ How to Run the Script

### **Step 1: Prepare Your Files**

1. Place all your **Word (.docx) and PDF (.pdf) files** in a single folder.
2. Ensure the files are named properly for correct order in the bundle.

### **Step 2: Run the Script**

#### **Windows (Double Click Method)**

1. Navigate to the folder where you extracted **AutoPDFBinder**.
2. Locate the **autopdfbinder.py** file.
3. **Right-click** on `autopdfbinder.py` and select **Open with > Python** (or double-click if Python is set as the default).

#### **Command Line Method (Recommended for Logs)**

1. Open **Command Prompt (cmd)**.
2. Navigate to the script folder:
   ```sh
   cd C:\AutoPDFBinder
   ```
3. Run the script:
   ```sh
   python autopdfbinder.py
   ```

### **Step 3: Output Files**

- `final_output.pdf` – The fully compiled and formatted document, ready to print.
- `script_log.txt` – Log file for troubleshooting.

---

## 🔧 Configuration

You can modify settings at the top of the script:

- `BATES_START = 1` – Change starting page number
- `BATES_FONT_SIZE = 14` – Adjust Bates numbering size
- `OUTPUT_DIR = "output"` – Define where output files are saved

---

## 📌 Example

Imagine you have the following directory:

```
📁 Case_Files/
   ├── Contract.docx
   ├── Exhibit_A.pdf
   ├── Witness_Statement.docx
   ├── Evidence/
       ├── Photo1.pdf
       ├── Report.docx
```

Running the script will generate:

- A **single merged PDF** containing all documents
- A **table of contents** listing all files with corresponding Bates numbers
- **Cover pages** for each document
- **Bookmarks** for easy navigation

---

## ⚠️ Notes

- **Requires Windows** (due to Microsoft Word COM automation)
- Ensure Word documents are properly formatted before conversion
- Large batches may take time to process—allow a few minutes

---

## 📜 License

This project is open-source under the MIT License.

---

## 💬 Support & Contributions

Feel free to **submit issues, suggestions, or pull requests** to improve this tool!

---

## 📦 Setting Up in GitHub Codespaces

To run this script in GitHub Codespaces:

1. Open a new **GitHub Codespace** for your repository.
2. Run the following commands in the terminal:
   ```sh
   pip install -r requirements.txt
   ```
3. Place all Word and PDF files in your working directory.
4. Execute the script:
   ```sh
   python autopdfbinder.py
   ```
5. Download `final_output.pdf` from the workspace when processing is complete.
