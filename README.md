---

# Local Office Suite

A lightweight desktop application for working with **Word**, **Excel**, **PowerPoint**, and **PDF** files locally. Supports creating, editing, uploading, and saving office documents with basic functionality.

---

## Features

### Word Editor

* Create and edit `.docx` files.
* Change font family, size, and apply **bold**, **italic**, and **underline** styles.
* Upload existing Word documents and edit them.
* Save your edits to a new `.docx` file.

### Excel Editor

* View and edit Excel files (`.xlsx`) in a spreadsheet-style interface.
* Add new rows, edit cells directly.
* Perform simple numeric calculations: sum, average, min, max on selected cells.
* Upload and save Excel files.

### PowerPoint Editor

* Create and edit `.pptx` presentations.
* Add slides with title and content.
* Edit existing slides.
* Upload existing PowerPoint files and save them.

### PDF Merge

* Upload multiple PDF files and merge them into a single PDF.

---

## Installation

1. Make sure Python **3.9+** is installed on your system.
2. Install required packages using pip:

```bash
pip install tkinter
pip install python-docx
pip install openpyxl
pip install python-pptx
pip install PyPDF2
```

> Note: `tkinter` is usually included with Python, but if not, install via your package manager.

---

## Usage

1. Run the application:

```bash
python office_suite.py
```

2. The main window offers four buttons to access:

   * Word
   * Excel
   * PowerPoint
   * PDF Merge

3. **Word Editor**

   * Use the text area to type.
   * Select text and click **B/I/U** for bold, italic, underline.
   * Use "Upload Word" to load existing `.docx` files.
   * Use "Save Word" to save the document.

4. **Excel Editor**

   * Double-click a cell to edit its content.
   * Click **Add Row** to add a new row.
   * Upload existing `.xlsx` files or save your current sheet.

5. **PowerPoint Editor**

   * Add new slides using the title and content fields.
   * Select a slide from the list to edit.
   * Upload existing `.pptx` files or save the presentation.

6. **PDF Merge**

   * Add PDF files using **Add PDF Files**.
   * Remove files if needed.
   * Click **Merge PDFs** to create a single PDF.

---

## Keyboard Shortcuts (Excel)

* `Ctrl + m` → Sum selected numeric cells
* `Ctrl + a` → Average selected numeric cells
* `Ctrl + x` → Minimum of selected numeric cells
* `Ctrl + z` → Maximum of selected numeric cells

---

## License

This project is open-source and free to use.

---
