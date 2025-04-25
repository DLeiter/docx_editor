# DOCX File Format Technical Specifications

A DOCX file is actually a ZIP archive that follows the Office Open XML (OOXML) standard. Understanding this structure allows you to programmatically manipulate these files at a deeper level than just using high-level libraries.

## Basic Structure of a DOCX File

When you extract a .docx file, you'll find:

1. **_rels** directory - Contains relationship information
2. **docProps** directory - Document properties/metadata
3. **word** directory - The main content
   - **document.xml** - Main document content
   - **styles.xml** - Style definitions
   - **numbering.xml** - Numbering definitions
   - **settings.xml** - Document settings
   - **media** folder - Images and embedded objects
   - **theme** folder - Theme information
   - **footer*.xml** - Footer content
   - **header*.xml** - Header content
4. **[Content_Types].xml** - Defines content types in the document

## Programmatically Editing DOCX Files

You can edit DOCX files programmatically using several approaches:

### 1. Using python-docx (High-Level)

The `python-docx` library (which you're already using) provides a high-level API:

```python
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Create or load a document
doc = Document("example.docx")

# Metadata/properties
doc.core_properties.author = "Your Name"
doc.core_properties.title = "Document Title"

# Add paragraphs and control formatting
para = doc.add_paragraph("Sample text")
para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# Text formatting
run = para.add_run("Bold text")
run.bold = True
run.font.size = Pt(12)
run.font.name = "Arial"

# Working with sections for headers/footers
section = doc.sections[0]
header = section.header
header.paragraphs[0].text = "My Header"

# Save changes
doc.save("modified.docx")
```

### 2. Direct XML Manipulation (Low-Level)

For more advanced control, you can manipulate the XML directly:

```python
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

# Open the docx file as a zip
with zipfile.ZipFile("example.docx") as docx_zip:
    # Extract document.xml
    xml_content = docx_zip.read("word/document.xml")
    
    # Parse XML
    tree = ET.parse(BytesIO(xml_content))
    root = tree.getroot()
    
    # Namespaces in OOXML
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    # Modify content (example: find and replace text)
    for paragraph in root.findall('.//w:p', namespaces):
        for text_element in paragraph.findall('.//w:t', namespaces):
            if text_element.text and "OldText" in text_element.text:
                text_element.text = text_element.text.replace("OldText", "NewText")
    
    # Create new in-memory file with modified content
    modified_docx = BytesIO()
    with zipfile.ZipFile(modified_docx, 'w') as output_zip:
        # Copy all files from original zip except document.xml
        for item in docx_zip.infolist():
            if item.filename != "word/document.xml":
                output_zip.writestr(item, docx_zip.read(item.filename))
        
        # Add modified document.xml
        output_zip.writestr("word/document.xml", ET.tostring(root))
    
    # Save the modified file
    with open("modified.docx", "wb") as f:
        f.write(modified_docx.getvalue())
```

### 3. Using the python-docx-template Library

For template-based document generation:

```python
from docxtpl import DocxTemplate

doc = DocxTemplate("template.docx")
context = {
    'title': 'Document Title',
    'name': 'John Doe',
    'items': ['Item 1', 'Item 2', 'Item 3']
}
doc.render(context)
doc.save("generated.docx")
```

## Advanced DOCX Features You Can Manipulate Programmatically

1. **Document Properties**: Title, author, keywords, creation date, etc.
2. **Styles**: Define and apply paragraph and character styles
3. **Tables**: Create tables with specific formatting, merged cells
4. **Headers/Footers**: Different headers for first page, odd/even pages
5. **Sections**: Different layouts in different parts of the document
6. **Page Layout**: Margins, page orientation, columns
7. **Content Controls**: Form fields and data binding
8. **Tracked Changes**: Work with revision tracking

## Implementation Tips

1. **Use python-docx for most operations**: It handles the complex XML structure for you
2. **For advanced needs, use docx's oxml module**: Access to underlying XML
3. **Use styles where possible**: More maintainable than direct formatting
4. **Be careful with sections**: They control page layout, headers/footers
5. **Test thoroughly**: DOCX formatting can have unexpected interactions

## Sample Implementation for Document Properties Editor

```python
def edit_document_properties(self):
    """Edit document properties/metadata"""
    if not self.document:
        messagebox.showinfo("No Document", "Please open or create a document first.")
        return
        
    props_dialog = tk.Toplevel(self.root)
    props_dialog.title("Document Properties")
    
    # Create fields for common properties
    tk.Label(props_dialog, text="Title:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    title_var = tk.StringVar(value=self.document.core_properties.title or "")
    tk.Entry(props_dialog, textvariable=title_var, width=30).grid(row=0, column=1, padx=5, pady=5)
    
    tk.Label(props_dialog, text="Author:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    author_var = tk.StringVar(value=self.document.core_properties.author or "")
    tk.Entry(props_dialog, textvariable=author_var, width=30).grid(row=1, column=1, padx=5, pady=5)
    
    # Add more fields as needed
    
    def apply_properties():
        self.document.core_properties.title = title_var.get()
        self.document.core_properties.author = author_var.get()
        # Set other properties
        props_dialog.destroy()
        self.status_var.set("Document properties updated")
    
    tk.Button(props_dialog, text="Apply", command=apply_properties).grid(row=10, column=0, columnspan=2, pady=10)
```

## Useful OOXML References

1. [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) - Official OOXML specification
2. [Python-docx Documentation](https://python-docx.readthedocs.io/en/latest/) - Comprehensive guide to the python-docx library
3. [OpenXML SDK Documentation](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk) - Microsoft's documentation on the OpenXML SDK

This information should help with understanding the internal structure of DOCX files and provide guidance for more advanced programmatic manipulation of these documents in your editor.
