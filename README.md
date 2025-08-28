# Handbook-Generator-Python-Word-COM-Automation-
Python translation of legacy Visual Basic scripts for generating a Microsoft Word handbook from HTML files. Automates importing content, structuring chapters, inserting lessons, and building a dynamic table of contents.

Overview

This project was developed at SAIC as part of an initiative to modernize legacy Visual Basic scripts by converting them into Python. The script automates the process of generating a complete Microsoft Word handbook from a collection of HTML files.

The resulting document includes:

  A Title Page
  
  A dynamically generated Table of Contents (TOC)
  
  Multiple chapters sourced from HTML content
  
  Student Handouts inserted sequentially
  
  Specialized lesson modules (M270A1 and M142) added with TOC references
  
By leveraging the pywin32 library, the script programmatically controls Word via COM automation, ensuring consistent formatting, pagination, and TOC alignment.


Key Features
📄 Automated Word Document Creation – Opens Word, creates a new document, and manages content insertion.

📑 Dynamic Table of Contents – Titles and page numbers automatically aligned with right-margin tab stops.

📂 HTML Importing – Inserts chapters and lesson content directly from .htm / .html files.

✂️ Section & Page Break Management – Ensures each section starts cleanly on a new page.

🧾 Student Handouts Support – Iteratively loads all lesson handouts (e.g., lesson0.htm to lesson100.htm).

🚀 M270A1 & M142 Lesson Modules – Imports specialized training lessons into the handbook with single TOC entries.

How It Works
  Title Page – Inserts a pre-defined HTML file as the cover/title page.
  Table of Contents – Adds a TOC header and dynamically updates entries as chapters/lessons are imported.
  Chapters – Imports core chapters such as System Overview, Student Console Operations, Special Functions, etc.
  Student Handouts – Sequentially inserts all available lesson handouts from a designated folder.
  Advanced Lessons – Adds M270A1 and M142 lesson modules as distinct handbook sections.
  Cleanup – Removes extra breaks or spacing to finalize the document.

Requirements
  Python 3.x
  Microsoft Word (Windows only)
  pywin32 - (pip install pywin32)
 
Example Workflow
  Place your HTML content in the designated folder structure (e.g., handbook/, Student Handout/, M270A1/procedures/).
  Update the file paths in the script (e.g., title_page_path, base_path).
  Run the script:
    python handbook_generator.py
  Microsoft Word will open, populate the handbook, and display the result.

Potential Improvements
  Make file paths configurable via JSON or YAML instead of hardcoding.
  Add logging instead of print statements.
  Error handling for missing files or invalid HTML.
  Cross-platform compatibility (currently Windows + Word only).

Use Case
 Originally, this Python script was used to replace Visual Basic scripts that generated training handbooks for military systems. Converting to Python provided:
    Easier maintainability
    Compatibility with modern workflows
    Improved automation and flexibility
