#!/usr/bin/env python3
"""
DOCX Editor - A program to edit Word documents with advanced features
"""
import os
import sys
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, colorchooser, font
from tkinter.constants import *
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from PIL import Image, ImageTk
import base64
from io import BytesIO

class DocxEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Editor")
        self.root.geometry("1000x700")
        
        # Document variables
        self.current_file = None
        self.document = None
        self.document_images = []
        
        # Text formatting state variables
        self.current_font_family = "Arial"
        self.current_font_size = 11
        self.current_bold = False
        self.current_italic = False
        self.current_underline = False
        self.current_alignment = "left"
        self.current_color = "#000000"
        
        # Table management
        self.tables = []
        self.current_table = None
        
        # Paragraph styling
        self.paragraph_styles = []
        
        # Create all widgets and interface elements
        self.create_menu()
        self.create_widgets()
        self.create_formatting_toolbar()
        self.create_tabs()
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # File menu
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.new_document)
        filemenu.add_command(label="Open", command=self.open_file)
        filemenu.add_command(label="Save", command=self.save_file)
        filemenu.add_command(label="Save As", command=self.save_file_as)
        filemenu.add_separator()
        filemenu.add_command(label="Export as PDF", command=self.export_pdf)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        
        # Edit menu
        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Undo", command=self.undo)
        editmenu.add_command(label="Redo", command=self.redo)
        editmenu.add_separator()
        editmenu.add_command(label="Cut", command=lambda: self.text_editor.event_generate("<<Cut>>"))
        editmenu.add_command(label="Copy", command=lambda: self.text_editor.event_generate("<<Copy>>"))
        editmenu.add_command(label="Paste", command=lambda: self.text_editor.event_generate("<<Paste>>"))
        editmenu.add_separator()
        editmenu.add_command(label="Find/Replace", command=self.find_replace_dialog)
        editmenu.add_separator()
        editmenu.add_command(label="Clear All", command=self.clear_text)
        menubar.add_cascade(label="Edit", menu=editmenu)
        
        # Format menu
        formatmenu = tk.Menu(menubar, tearoff=0)
        
        # Text formatting submenu
        text_formatting_menu = tk.Menu(formatmenu, tearoff=0)
        text_formatting_menu.add_command(label="Font...", command=self.font_dialog)
        text_formatting_menu.add_command(label="Text Color...", command=self.text_color_dialog)
        text_formatting_menu.add_separator()
        text_formatting_menu.add_command(label="Bold", command=self.toggle_bold)
        text_formatting_menu.add_command(label="Italic", command=self.toggle_italic)
        text_formatting_menu.add_command(label="Underline", command=self.toggle_underline)
        formatmenu.add_cascade(label="Text", menu=text_formatting_menu)
        
        # Paragraph formatting submenu
        paragraph_menu = tk.Menu(formatmenu, tearoff=0)
        paragraph_menu.add_command(label="Align Left", command=lambda: self.set_alignment("left"))
        paragraph_menu.add_command(label="Align Center", command=lambda: self.set_alignment("center"))
        paragraph_menu.add_command(label="Align Right", command=lambda: self.set_alignment("right"))
        paragraph_menu.add_command(label="Justify", command=lambda: self.set_alignment("justify"))
        paragraph_menu.add_separator()
        paragraph_menu.add_command(label="Paragraph Styles...", command=self.paragraph_style_dialog)
        formatmenu.add_cascade(label="Paragraph", menu=paragraph_menu)
        
        menubar.add_cascade(label="Format", menu=formatmenu)
        
        # Insert menu
        insertmenu = tk.Menu(menubar, tearoff=0)
        insertmenu.add_command(label="Image...", command=self.insert_image)
        insertmenu.add_command(label="Table...", command=self.insert_table_dialog)
        insertmenu.add_separator()
        insertmenu.add_command(label="Page Break", command=self.insert_page_break)
        insertmenu.add_command(label="Hyperlink...", command=self.insert_hyperlink)
        menubar.add_cascade(label="Insert", menu=insertmenu)
        
        # Table menu
        tablemenu = tk.Menu(menubar, tearoff=0)
        tablemenu.add_command(label="Insert Table...", command=self.insert_table_dialog)
        tablemenu.add_command(label="Edit Table...", command=self.edit_table_dialog)
        tablemenu.add_separator()
        tablemenu.add_command(label="Add Row", command=self.add_table_row)
        tablemenu.add_command(label="Add Column", command=self.add_table_column)
        tablemenu.add_separator()
        tablemenu.add_command(label="Delete Row", command=self.delete_table_row)
        tablemenu.add_command(label="Delete Column", command=self.delete_table_column)
        menubar.add_cascade(label="Table", menu=tablemenu)
        
        # View menu
        viewmenu = tk.Menu(menubar, tearoff=0)
        viewmenu.add_command(label="Document Properties", command=self.document_properties)
        viewmenu.add_separator()
        viewmenu.add_command(label="Zoom In", command=self.zoom_in)
        viewmenu.add_command(label="Zoom Out", command=self.zoom_out)
        viewmenu.add_command(label="Reset Zoom", command=self.reset_zoom)
        menubar.add_cascade(label="View", menu=viewmenu)
        
        # Help menu
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_about)
        helpmenu.add_command(label="Help", command=self.show_help)
        menubar.add_cascade(label="Help", menu=helpmenu)
        
        self.root.config(menu=menubar)
    
    def create_widgets(self):
        # Frame for buttons
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Open button
        self.open_button = tk.Button(self.button_frame, text="Open Document", command=self.open_file)
        self.open_button.pack(side=tk.LEFT, padx=5)
        
        # Save button
        self.save_button = tk.Button(self.button_frame, text="Save Document", command=self.save_file)
        self.save_button.pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("No file open")
        self.status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Text editor
        self.text_editor = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=80, height=30)
        self.text_editor.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
    
    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if file_path:
            try:
                self.document = Document(file_path)
                self.current_file = file_path
                
                # Extract text from document
                text_content = ""
                for para in self.document.paragraphs:
                    text_content += para.text + "\n"
                
                # Update UI
                self.text_editor.delete(1.0, tk.END)
                self.text_editor.insert(tk.END, text_content)
                self.status_var.set(f"Opened: {os.path.basename(file_path)}")
                messagebox.showinfo("Success", f"Opened {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def save_file(self):
        if not self.current_file:
            self.save_file_as()
        else:
            self.save_document(self.current_file)
    
    def save_file_as(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.save_document(file_path)
            self.current_file = file_path
    
    def save_document(self, file_path):
        try:
            # Create a new document if we don't have one
            if not self.document:
                self.document = Document()
            
            # Get text content from editor
            text_content = self.text_editor.get(1.0, tk.END)
            
            # Clear existing content
            for para in list(self.document.paragraphs):
                p = para._element
                p.getparent().remove(p)
                para._p = para._element = None
            
            # Add new content
            for line in text_content.split('\n'):
                self.document.add_paragraph(line)
            
            # Save the document
            self.document.save(file_path)
            self.status_var.set(f"Saved: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"Saved to {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def clear_text(self):
        self.text_editor.delete(1.0, tk.END)

def main():
    root = tk.Tk()
    app = DocxEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
