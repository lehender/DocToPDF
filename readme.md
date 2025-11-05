# DocToPDF

A simple, portable desktop app for converting Office files to PDF... No installation required!  
Built with **PySide6** and **LibreOffice headless mode** for clean, offline conversions.

---

## ✨ Features
- 🧩 Drag-and-drop interface for DOCX, XLSX, PPTX, ODT, ODS, and ODP files  
- 🪄 Converts multiple files at once  
- 📂 Choose a custom output folder or default to the source directory  
- 💡 No internet or Microsoft Office required  
- 🎨 Compact, modern UI with solid color cards and smooth scaling  
- ⚙️ Fully open-source (Python + Qt)  

---

## 🚀 How to Use
1. Download the **Portable ZIP** from the [Releases](../../releases/latest) page  
2. Extract the folder anywhere  
3. Run `DocToPDF.exe`  
4. Drag and drop your Office files onto the window or click **Choose Files**  

Converted PDFs will appear in the same folder as the originals (or your chosen output folder).

---

## 🛠️ Building from Source
Requirements:
```bash
pip install PySide6 pyinstaller
