# Required installations:
# pip install pandas tabula-py openpyxl pillow tqdm pdf2image PyMuPDF PyPDF2 camelot-py opencv-python pytesseract pyocr reportlab

# Add these imports at the top
import uuid
import hashlib
import datetime
import winreg
import socket

def install_missing_packages():
    import subprocess
    import sys
    
    required_packages = [
        'pandas', 'tabula-py', 'openpyxl', 'pillow', 'tqdm',
        'pdf2image', 'PyMuPDF', 'PyPDF2', 'camelot-py', 'opencv-python',
        'pytesseract', 'pyocr', 'reportlab'
    ]
    
    for package in required_packages:
        try:
            __import__(package.replace('-', '_').split('[')[0])
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Try to install missing packages
install_missing_packages()

# Import statements
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tabula.io import read_pdf
import pandas as pd
import os
import sys
from datetime import datetime
from PIL import Image, ImageTk, ImageDraw, ImageEnhance
from tqdm import tqdm
import threading
import fitz  # PyMuPDF
import camelot
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfMerger, PdfWriter
import cv2
import numpy as np
import json
import pytesseract
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red, blue, yellow

# Add these functions before the PDFToExcelConverter class
def get_hardware_id():
    """Generate a unique hardware ID"""
    system_info = [
        platform.node(),
        platform.machine(),
        str(uuid.getnode()),  # MAC address
        platform.processor()
    ]
    return hashlib.md5(''.join(system_info).encode()).hexdigest()

def verify_license(key):
    """Verify license key against hardware ID"""
    hardware_id = get_hardware_id()
    expected_key = hashlib.sha256(hardware_id.encode()).hexdigest()[:16]
    return key == expected_key

class PDFToExcelConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.check_license()
        
        self.title("PDF Table Extractor Professional")
        self.state('zoomed')  # Start maximized
        self.configure(bg='#f0f0f0')
        
        # Initialize variables
        self.current_page = 1
        self.pdf_path = None
        self.password = None
        self.zoom_level = 1.0
        self.annotations = []
        self.undo_stack = []
        self.current_tool = None
        
        self.create_menu()
        self.create_toolbar()
        self.create_main_interface()
        self.bind_events()
        
    def check_license(self):
        """Verify license and hardware ID"""
        config_file = 'license.json'
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
                if not verify_license(config.get('license_key', '')):
                    self.show_activation_dialog()
                expiry = datetime.datetime.strptime(config.get('expiry', '2000-01-01'), '%Y-%m-%d')
                if datetime.datetime.now() > expiry:
                    messagebox.showerror("Error", "License has expired!")
                    self.destroy()
        except FileNotFoundError:
            self.show_activation_dialog()

    def show_activation_dialog(self):
        """Show dialog for license key entry"""
        hardware_id = get_hardware_id()
        dialog = tk.Toplevel(self)
        dialog.title("Software Activation")
        dialog.geometry("400x200")
        
        tk.Label(dialog, text="Hardware ID:").pack(pady=5)
        tk.Label(dialog, text=hardware_id).pack(pady=5)
        tk.Label(dialog, text="Enter License Key:").pack(pady=5)
        
        key_entry = tk.Entry(dialog, width=40)
        key_entry.pack(pady=5)
        
        def validate_key():
            if verify_license(key_entry.get()):
                with open('license.json', 'w') as f:
                    json.dump({
                        'license_key': key_entry.get(),
                        'expiry': (datetime.datetime.now() + datetime.timedelta(days=365)).strftime('%Y-%m-%d')
                    }, f)
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Invalid license key!")
                self.destroy()
        
        tk.Button(dialog, text="Activate", command=validate_key).pack(pady=10)
        dialog.transient(self)
        dialog.grab_set()
        self.wait_window(dialog)

    def create_menu(self):
        menubar = tk.Menu(self)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open PDF", command=self.open_pdf)
        file_menu.add_command(label="Save PDF", command=self.save_pdf)
        file_menu.add_command(label="Export as Image", command=self.export_as_image)
        file_menu.add_separator()
        file_menu.add_command(label="Merge PDFs", command=self.merge_pdfs)
        file_menu.add_command(label="Split PDF", command=self.split_pdf)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Add Text", command=lambda: self.set_tool("text"))
        edit_menu.add_command(label="Add Image", command=self.add_image)
        edit_menu.add_command(label="Delete Page", command=self.delete_current_page)
        edit_menu.add_command(label="Rotate Page", command=self.rotate_page)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Highlight", command=lambda: self.set_tool("highlight"))
        tools_menu.add_command(label="Underline", command=lambda: self.set_tool("underline"))
        tools_menu.add_command(label="Draw", command=lambda: self.set_tool("draw"))
        tools_menu.add_command(label="Sticky Note", command=lambda: self.set_tool("note"))
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        self.config(menu=menubar)

    def create_toolbar(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X)
        
        # Zoom controls
        ttk.Button(toolbar, text="Zoom In", command=self.zoom_in).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Zoom Out", command=self.zoom_out).pack(side=tk.LEFT)
        
        # Page navigation
        ttk.Button(toolbar, text="Previous", command=self.prev_page).pack(side=tk.LEFT)
        self.page_var = tk.StringVar()
        ttk.Entry(toolbar, textvariable=self.page_var, width=5).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Go", command=self.go_to_page).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Next", command=self.next_page).pack(side=tk.LEFT)

    def create_main_interface(self):
        # Create main container
        self.main_container = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel - PDF preview
        self.create_preview_panel()
        
        # Right panel - Controls
        self.create_control_panel()
        
    def create_preview_panel(self):
        preview_frame = ttk.Frame(self.main_container)
        self.main_container.add(preview_frame, weight=2)
        
        # PDF Preview
        self.preview_canvas = tk.Canvas(preview_frame, bg='gray')
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Navigation
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.pack(fill=tk.X)
        
        ttk.Button(nav_frame, text="←", command=self.prev_page).pack(side=tk.LEFT)
        self.page_label = ttk.Label(nav_frame, text="Page: 1")
        self.page_label.pack(side=tk.LEFT)
        ttk.Button(nav_frame, text="→", command=self.next_page).pack(side=tk.LEFT)
        
    def create_control_panel(self):
        control_frame = ttk.Frame(self.main_container)
        self.main_container.add(control_frame, weight=1)
        
        # Extraction Options
        options_frame = ttk.LabelFrame(control_frame, text="Advanced Options")
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Engine selection
        ttk.Label(options_frame, text="Extraction Engine:").pack()
        self.engine_var = tk.StringVar(value="auto")
        engines = ["Auto", "Tabula", "Camelot", "OCR"]
        ttk.OptionMenu(options_frame, self.engine_var, *engines).pack()
        
        # Table detection options
        self.lattice_var = tk.BooleanVar(value=True)
        self.stream_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Detect Bordered Tables", 
                       variable=self.lattice_var).pack()
        ttk.Checkbutton(options_frame, text="Detect Borderless Tables", 
                       variable=self.stream_var).pack()
        
        # Page selection
        pages_frame = ttk.LabelFrame(options_frame, text="Page Selection")
        pages_frame.pack(fill=tk.X, padx=5, pady=5)
        self.pages_var = tk.StringVar(value="all")
        ttk.Radiobutton(pages_frame, text="All Pages", 
                       variable=self.pages_var, value="all").pack()
        ttk.Radiobutton(pages_frame, text="Current Page", 
                       variable=self.pages_var, value="current").pack()
        ttk.Radiobutton(pages_frame, text="Custom Range", 
                       variable=self.pages_var, value="custom").pack()
        self.page_range = ttk.Entry(pages_frame)
        self.page_range.pack()

        # Output Options
        output_frame = ttk.LabelFrame(control_frame, text="Output Options")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.format_var = tk.StringVar(value="xlsx")
        ttk.Radiobutton(output_frame, text="Excel (XLSX)", 
                       variable=self.format_var, value="xlsx").pack()
        ttk.Radiobutton(output_frame, text="CSV", 
                       variable=self.format_var, value="csv").pack()
        
        # Conversion button
        self.convert_btn = ttk.Button(control_frame, text="Convert", 
                                      command=self.start_conversion)
        self.convert_btn.pack(pady=10)
        
        # Status and progress
        self.status_label = ttk.Label(control_frame, text="Ready")
        self.status_label.pack()
        self.progress = ttk.Progressbar(control_frame, mode='determinate')
        self.progress.pack(fill=tk.X, padx=5)

    def detect_tables(self):
        if not self.pdf_path:
            return
            
        try:
            # Use OpenCV to detect table structures
            image = convert_from_path(self.pdf_path, first_page=self.current_page,
                                    last_page=self.current_page)[0]
            img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            # Table detection logic
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            edges = cv2.Canny(gray, 50, 150, apertureSize=3)
            lines = cv2.HoughLinesP(edges, 1, np.pi/180, 100, minLineLength=100, maxLineGap=10)
            
            # Draw detected tables on preview
            for line in lines:
                x1, y1, x2, y2 = line[0]
                cv2.line(img_cv, (x1, y1), (x2, y2), (0, 255, 0), 2)
            
            # Update preview with detected tables
            self.update_preview(Image.fromarray(cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB)))
            
        except Exception as e:
            messagebox.showerror("Error", f"Table detection failed: {str(e)}")

    def open_pdf(self):
        """Open and load a PDF file"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.pdf_path = file_path
                self.pdf_document = fitz.open(file_path)
                self.total_pages = len(self.pdf_document)
                self.current_page = 1
                self.update_preview()
                self.status_label.config(text=f"Loaded PDF: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF: {str(e)}")

    def save_pdf(self):
        """Save PDF with annotations"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            if output_path:
                # Create new PDF with annotations
                writer = PdfWriter()
                reader = PdfReader(self.pdf_path)
                
                for i in range(len(reader.pages)):
                    page = reader.pages[i]
                    # Add annotations for this page
                    page_annotations = [a for a in self.annotations if a['page'] == i + 1]
                    for annot in page_annotations:
                        if annot['type'] == 'text':
                            page.insert_text(annot['coords'], annot['text'])
                        elif annot['type'] == 'highlight':
                            page.add_highlight_annot(annot['coords'])
                    writer.add_page(page)
                
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                messagebox.showinfo("Success", "PDF saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def export_as_image(self):
        """Export current page as image"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg")]
            )
            if output_path:
                image = convert_from_path(self.pdf_path, first_page=self.current_page,
                                        last_page=self.current_page)[0]
                image.save(output_path)
                messagebox.showinfo("Success", "Image exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export image: {str(e)}")

    def add_image(self):
        """Add image to current page"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if file_path:
            self.annotations.append({
                'type': 'image',
                'path': file_path,
                'coords': (100, 100),  # Default position
                'page': self.current_page
            })
            self.update_preview()

    def delete_current_page(self):
        """Delete current page from PDF"""
        if not self.pdf_path or not hasattr(self, 'pdf_document'):
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        
        if messagebox.askyesno("Confirm", "Delete current page?"):
            try:
                self.pdf_document.delete_page(self.current_page - 1)
                self.total_pages = len(self.pdf_document)
                if self.current_page > self.total_pages:
                    self.current_page = self.total_pages
                self.update_preview()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete page: {str(e)}")

    def rotate_page(self):
        """Rotate current page"""
        if not self.pdf_path or not hasattr(self, 'pdf_document'):
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        
        try:
            page = self.pdf_document[self.current_page - 1]
            page.set_rotation((page.rotation + 90) % 360)
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to rotate page: {str(e)}")

    def zoom_in(self):
        """Zoom in preview"""
        self.zoom_level *= 1.2
        self.update_preview()

    def zoom_out(self):
        """Zoom out preview"""
        self.zoom_level /= 1.2
        self.update_preview()

    def go_to_page(self):
        """Go to specific page number"""
        try:
            page = int(self.page_var.get())
            if 1 <= page <= self.total_pages:
                self.current_page = page
                self.update_preview()
                self.page_label.config(text=f"Page: {self.current_page}")
            else:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Invalid page number!")

    def draw_annotation(self, event):
        """Handle drawing annotations"""
        if self.current_tool == "draw":
            self.annotations.append({
                'type': 'draw',
                'coords': (event.x, event.y),
                'page': self.current_page
            })
            self.update_preview()

    def finish_annotation(self, event):
        """Finish current annotation"""
        if self.current_tool:
            self.undo_stack.append(self.annotations[:])
            self.current_tool = None
            self.preview_canvas.config(cursor="arrow")

    def save_settings(self):
        """Save current settings to a config file"""
        settings = {
            'engine': self.engine_var.get(),
            'lattice': self.lattice_var.get(),
            'stream': self.stream_var.get(),
            'format': self.format_var.get()
        }
        try:
            with open('pdf_converter_settings.json', 'w') as f:
                json.dump(settings, f)
            messagebox.showinfo("Success", "Settings saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def ocr_preprocess(self):
        """Preprocess current page for OCR"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
        try:
            image = convert_from_path(self.pdf_path, first_page=self.current_page,
                                    last_page=self.current_page)[0]
            # Apply preprocessing
            img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            denoised = cv2.fastNlMeansDenoising(gray)
            thresh = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
            # Update preview
            self.update_preview(Image.fromarray(thresh))
        except Exception as e:
            messagebox.showerror("Error", f"OCR preprocessing failed: {str(e)}")

    def prev_page(self):
        """Go to previous page"""
        if hasattr(self, 'pdf_document') and self.current_page > 1:
            self.current_page -= 1
            self.update_preview()
            self.page_label.config(text=f"Page: {self.current_page}")

    def next_page(self):
        """Go to next page"""
        if hasattr(self, 'pdf_document') and self.current_page < self.total_pages:
            self.current_page += 1
            self.update_preview()
            self.page_label.config(text=f"Page: {self.current_page}")

    def update_preview(self, image=None):
        """Update the preview panel with current page"""
        if image is None and hasattr(self, 'pdf_document'):
            page = self.pdf_document[self.current_page - 1]
            pix = page.get_pixmap()
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        if image:
            # Resize image to fit canvas
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            if canvas_width > 1 and canvas_height > 1:  # Canvas has been rendered
                ratio = min(canvas_width/image.width, canvas_height/image.height)
                new_size = (int(image.width*ratio), int(image.height*ratio))
                image = image.resize(new_size, Image.Resampling.LANCZOS)
                
                self.photo = ImageTk.PhotoImage(image)
                self.preview_canvas.delete("all")
                self.preview_canvas.create_image(
                    canvas_width//2, canvas_height//2,
                    image=self.photo, anchor="center"
                )

    def start_conversion(self):
        """Start the conversion process"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return

        try:
            # Determine pages to process
            if self.pages_var.get() == "all":
                pages = "all"
            elif self.pages_var.get() == "current":
                pages = str(self.current_page)
            else:
                pages = self.page_range.get()

            # Start conversion in a separate thread
            self.convert_btn.config(state='disabled')
            self.progress.start()
            
            def conversion_thread():
                try:
                    # Extract tables based on selected engine
                    if self.engine_var.get() == "Camelot":
                        tables = camelot.read_pdf(self.pdf_path, pages=pages)
                    else:
                        tables = read_pdf(self.pdf_path, pages=pages,
                                        lattice=self.lattice_var.get(),
                                        stream=self.stream_var.get())
                    
                    if tables:
                        # Save to selected format
                        output_path = os.path.splitext(self.pdf_path)[0]
                        if self.format_var.get() == "xlsx":
                            output_path += "_converted.xlsx"
                            with pd.ExcelWriter(output_path) as writer:
                                for i, table in enumerate(tables):
                                    table.df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
                        else:
                            output_path += "_converted.csv"
                            pd.concat([table.df for table in tables]).to_csv(output_path, index=False)
                        
                        messagebox.showinfo("Success", f"Saved to {output_path}")
                    else:
                        messagebox.showinfo("Info", "No tables found in the selected pages.")
                
                except Exception as e:
                    messagebox.showerror("Error", f"Conversion failed: {str(e)}")
                finally:
                    self.progress.stop()
                    self.convert_btn.config(state='normal')
            
            threading.Thread(target=conversion_thread, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start conversion: {str(e)}")

    def set_tool(self, tool_name):
        """Set current annotation tool"""
        self.current_tool = tool_name
        self.preview_canvas.config(cursor="crosshair" if tool_name else "arrow")

    def add_annotation(self, event):
        """Handle annotation creation based on current tool"""
        if not self.current_tool or not hasattr(self, 'pdf_document'):
            return
            
        x, y = event.x, event.y
        if self.current_tool == "highlight":
            self.annotations.append({
                'type': 'highlight',
                'coords': (x, y, x+100, y+20),
                'page': self.current_page
            })
        elif self.current_tool == "text":
            text = simpledialog.askstring("Add Text", "Enter text:")
            if text:
                self.annotations.append({
                    'type': 'text',
                    'text': text,
                    'coords': (x, y),
                    'page': self.current_page
                })
        self.update_preview()

    def merge_pdfs(self):
        """Merge multiple PDFs"""
        files = filedialog.askopenfilenames(
            filetypes=[("PDF files", "*.pdf")]
        )
        if files:
            try:
                merger = PdfMerger()
                for pdf in files:
                    merger.append(pdf)
                    
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")]
                )
                if output_path:
                    merger.write(output_path)
                    merger.close()
                    messagebox.showinfo("Success", "PDFs merged successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs: {str(e)}")

    def split_pdf(self):
        """Split PDF into separate pages"""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please open a PDF first!")
            return
            
        try:
            output_dir = filedialog.askdirectory(title="Select output directory")
            if output_dir:
                pdf = PdfReader(self.pdf_path)
                writer = PdfWriter()
                
                for i in range(len(pdf.pages)):
                    writer.add_page(pdf.pages[i])
                    output_path = os.path.join(output_dir, f'page_{i+1}.pdf')
                    with open(output_path, 'wb') as output_file:
                        writer.write(output_file)
                
                messagebox.showinfo("Success", f"PDF split into {len(pdf.pages)} pages!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to split PDF: {str(e)}")

    def perform_ocr(self):
        """Perform OCR on current page"""
        if not self.pdf_path:
            return
            
        try:
            image = convert_from_path(self.pdf_path, first_page=self.current_page,
                                    last_page=self.current_page)[0]
            text = pytesseract.image_to_string(image)
            
            # Create searchable PDF
            output_path = os.path.splitext(self.pdf_path)[0] + "_searchable.pdf"
            pdf = canvas.Canvas(output_path)
            pdf.drawString(10, 10, text)
            pdf.save()
            
            messagebox.showinfo("Success", "OCR completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"OCR failed: {str(e)}")

    def bind_events(self):
        self.preview_canvas.bind("<Button-1>", self.add_annotation)
        self.preview_canvas.bind("<B1-Motion>", self.draw_annotation)
        self.preview_canvas.bind("<ButtonRelease-1>", self.finish_annotation)

if __name__ == "__main__":
    app = PDFToExcelConverter()
    app.mainloop()
