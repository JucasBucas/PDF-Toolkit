import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import os
import sys
from PIL import Image, ImageTk
import io
import pdfplumber
from docx import Document
import pandas as pd
import threading
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import datetime
import pymupdf

class PDFToolkitGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Toolkit Pro")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)
        
        self.current_file = tk.StringVar()
        self.progress = tk.DoubleVar()
        self.selected_files = []
        
        self.setup_ui()
        self.apply_theme()
    
    def apply_theme(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=6, relief='flat', background='#0078d4', foreground='white')
        style.map('TButton', background=[('active', '#005a9e')])
        style.configure('TLabelframe', background='#f0f0f0', relief='solid')
        style.configure('TLabelframe.Label', background='#f0f0f0')
        style.configure('TProgressbar', background='#0078d4')
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=4, pady=(0, 15), sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(title_frame, text="📄 PDF Toolkit Pro", font=('Arial', 20, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        version_label = ttk.Label(title_frame, text="v2.0", font=('Arial', 10))
        version_label.pack(side=tk.RIGHT, padx=(0, 10))

        file_frame = ttk.LabelFrame(main_frame, text="📁 File Selection", padding="12")
        file_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 12))
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="📂 Browse File", command=self.browse_file, width=15).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="📁 Browse Folder", command=self.browse_folder, width=15).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="❌ Clear All", command=self.clear_files, width=12).pack(side=tk.RIGHT)
        
        entry_frame = ttk.Frame(file_frame)
        entry_frame.pack(fill=tk.X)
        
        self.file_entry = ttk.Entry(entry_frame, textvariable=self.current_file, font=('Arial', 10))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(entry_frame, text="🔍 Preview", command=self.preview_selected, width=10).pack(side=tk.RIGHT)

        file_list_frame = ttk.LabelFrame(main_frame, text="📋 Selected Files", padding="12")
        file_list_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.file_list_text = scrolledtext.ScrolledText(file_list_frame, height=5, width=100, font=('Courier', 9))
        self.file_list_text.pack(fill=tk.BOTH, expand=True)
        self.file_list_text.insert(tk.END, "No files selected")
        self.file_list_text.config(state=tk.DISABLED)

        ops_frame = ttk.LabelFrame(main_frame, text="⚙️ PDF Operations", padding="15")
        ops_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ops = [
            ("🖼️ PDF to Images", self.pdf_to_images_gui),
            ("📝 PDF to Word", self.pdf_to_word_gui),
            ("📊 Extract Tables", self.extract_tables_gui),
            ("📄 PDF to Text", self.pdf_to_text_gui),
            ("🖼️ Images to PDF", self.images_to_pdf_gui),
            ("🔗 Merge PDFs", self.merge_pdfs_gui),
            ("✂️ Split PDF", self.split_pdf_gui),
            ("📑 Extract Pages", self.extract_pages_gui),
            ("🔒 Protect PDF", self.protect_pdf_gui),
            ("🔓 Unlock PDF", self.unlock_pdf_gui),
            ("📏 Compress PDF", self.compress_pdf_gui),
            ("🔄 Rotate PDF", self.rotate_pdf_gui)
        ]
        
        row, col = 0, 0
        for text, command in ops:
            ttk.Button(ops_frame, text=text, command=command, width=18).grid(
                row=row, column=col, padx=8, pady=8, sticky=(tk.W, tk.E)
            )
            col += 1
            if col > 2:
                col = 0
                row += 1

        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 5))
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress, maximum=100, length=400)
        self.progress_bar.pack(fill=tk.X, expand=True)
        
        self.progress_label = ttk.Label(progress_frame, text="🟢 Ready")
        self.progress_label.pack()

        preview_frame = ttk.LabelFrame(main_frame, text="👁️ Preview", padding="12")
        preview_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=12, width=100, font=('Arial', 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.status_label = ttk.Label(status_frame, text="Ready | 0 files selected", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, ipady=5)

        for i in range(4):
            main_frame.columnconfigure(i, weight=1)
        main_frame.rowconfigure(5, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def get_unique_filename(self, base_path, extension):
        counter = 1
        base_name = os.path.splitext(base_path)[0]
        new_path = f"{base_name}{extension}"
        
        while os.path.exists(new_path):
            new_path = f"{base_name}_{counter}{extension}"
            counter += 1
        
        return new_path

    def get_unique_folder(self, base_folder):
        counter = 1
        new_folder = base_folder
        
        while os.path.exists(new_folder):
            new_folder = f"{base_folder}_{counter}"
            counter += 1
        
        return new_folder

    def browse_file(self):
        filenames = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filenames:
            self.selected_files = list(filenames)
            self.current_file.set(filenames[0] if filenames else "")
            self.update_file_list()
            self.update_status()

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            pdf_files = []
            for root_dir, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root_dir, file))
            
            self.selected_files = pdf_files
            self.current_file.set(folder if pdf_files else "")
            self.update_file_list()
            self.update_status()

    def preview_selected(self):
        if self.selected_files:
            self.update_preview(self.selected_files[0])

    def update_file_list(self):
        self.file_list_text.config(state=tk.NORMAL)
        self.file_list_text.delete(1.0, tk.END)
        if self.selected_files:
            for f in self.selected_files[:10]:
                name = os.path.basename(f)
                size = os.path.getsize(f) / 1024
                self.file_list_text.insert(tk.END, f"• {name} ({size:.1f} KB)\n")
            if len(self.selected_files) > 10:
                self.file_list_text.insert(tk.END, f"... and {len(self.selected_files) - 10} more files")
        else:
            self.file_list_text.insert(tk.END, "No files selected")
        self.file_list_text.config(state=tk.DISABLED)

    def clear_files(self):
        self.selected_files = []
        self.current_file.set("")
        self.update_file_list()
        self.preview_text.delete(1.0, tk.END)
        self.update_status()

    def update_preview(self, filename):
        try:
            if filename.lower().endswith('.pdf'):
                with pdfplumber.open(filename) as pdf:
                    preview_text = f"📄 {os.path.basename(filename)}\n"
                    preview_text += f"📏 Size: {os.path.getsize(filename)/1024:.1f} KB\n"
                    preview_text += f"📑 Pages: {len(pdf.pages)}\n"
                    preview_text += "─" * 50 + "\n\n"
                    
                    if len(pdf.pages) > 0:
                        first_page = pdf.pages[0]
                        text = first_page.extract_text()
                        if text:
                            lines = text.split('\n')[:10]
                            preview_text += '\n'.join(lines)
                            if len(text.split('\n')) > 10:
                                preview_text += "\n\n... (more content)"
                    
                    self.preview_text.delete(1.0, tk.END)
                    self.preview_text.insert(1.0, preview_text)
        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"⚠️ Preview error: {str(e)}")

    def update_progress(self, value, text=""):
        self.progress.set(value)
        if text:
            self.progress_label.config(text=text)
        self.root.update_idletasks()

    def update_status(self):
        count = len(self.selected_files)
        status = f"🟢 Ready | {count} file{'s' if count != 1 else ''} selected"
        self.status_label.config(text=status)

    def run_in_thread(self, func, *args):
        thread = threading.Thread(target=func, args=args)
        thread.daemon = True
        thread.start()

    def validate_pdf_file(self):
        if not self.selected_files:
            messagebox.showwarning("No File", "Please select a PDF file first")
            return False
        return True

    def pdf_to_images_gui(self):
        if not self.validate_pdf_file():
            return
        
        dpi = simpledialog.askinteger("Image Quality", "Enter DPI (72-300):", minvalue=72, maxvalue=300, initialvalue=150)
        if not dpi:
            return
        
        fmt = simpledialog.askstring("Image Format", "Enter format (png/jpg):", initialvalue="png")
        if not fmt or fmt.lower() not in ['png', 'jpg', 'jpeg']:
            messagebox.showerror("Error", "Format must be png or jpg")
            return
        
        output_folder = filedialog.askdirectory(title="Select output folder")
        if not output_folder:
            return
        
        self.run_in_thread(self.pdf_to_images_thread, self.selected_files[0], output_folder, dpi, fmt.lower())

    def pdf_to_images_thread(self, pdf_path, output_folder, dpi, fmt):
        try:
            self.update_progress(0, "Converting PDF to images...")
            
            doc = pymupdf.open(pdf_path)
            total_pages = len(doc)
            zoom = dpi / 72
            
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            image_folder = os.path.join(output_folder, f"{base_name}_images")
            os.makedirs(image_folder, exist_ok=True)
            
            for page_num in range(total_pages):
                page = doc[page_num]
                mat = pymupdf.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                
                image_filename = f"page_{page_num + 1:03d}.{fmt}"
                output_path = os.path.join(image_folder, image_filename)
                pix.save(output_path)
                
                progress = (page_num + 1) / total_pages * 100
                self.update_progress(progress, f"Processed page {page_num + 1}/{total_pages}")
            
            doc.close()
            self.update_progress(100, "✅ Conversion completed!")
            messagebox.showinfo("Success", f"PDF converted to {fmt.upper()} images in {image_folder}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    def pdf_to_word_gui(self):
        if not self.validate_pdf_file():
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.pdf_to_word_thread, self.selected_files[0], output_file)

    def pdf_to_word_thread(self, pdf_path, output_file):
        try:
            self.update_progress(0, "Converting PDF to Word...")
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                doc = Document()
                
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        doc.add_paragraph(text)
                    
                    if i == 0:
                        metadata = pdf.metadata
                        if metadata:
                            doc.core_properties.title = metadata.get('Title', '')
                            doc.core_properties.author = metadata.get('Author', '')
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Processed page {i + 1}/{total_pages}")
                
                doc.save(output_file)
            
            self.update_progress(100, "✅ Conversion completed!")
            messagebox.showinfo("Success", f"PDF converted to Word: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    def extract_tables_gui(self):
        if not self.validate_pdf_file():
            return
        
        output_folder = filedialog.askdirectory(title="Select output folder")
        if not output_folder:
            return
        
        self.run_in_thread(self.extract_tables_thread, self.selected_files[0], output_folder)

    def extract_tables_thread(self, pdf_path, output_folder):
        try:
            self.update_progress(0, "Extracting tables from PDF...")
            
            all_tables = []
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    
                    for table_num, table_data in enumerate(tables):
                        if table_data:
                            df = pd.DataFrame(table_data)
                            df = df.dropna(how='all').dropna(axis=1, how='all')
                            
                            if not df.empty:
                                all_tables.append({
                                    'page': page_num + 1,
                                    'table_num': table_num + 1,
                                    'dataframe': df
                                })
                    
                    progress = (page_num + 1) / total_pages * 100
                    self.update_progress(progress, f"Processed page {page_num + 1}/{total_pages}")
            
            if all_tables:
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                table_folder = os.path.join(output_folder, f"{base_name}_tables")
                os.makedirs(table_folder, exist_ok=True)
                
                with pd.ExcelWriter(os.path.join(table_folder, "all_tables.xlsx")) as writer:
                    for table_info in all_tables:
                        df = table_info['dataframe']
                        page = table_info['page']
                        table_num = table_info['table_num']
                        
                        sheet_name = f"Page_{page}_T{table_num}"[:31]
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        csv_file = os.path.join(table_folder, f"table_p{page}_t{table_num}.csv")
                        df.to_csv(csv_file, index=False, encoding='utf-8')
                
                self.update_progress(100, f"✅ Extracted {len(all_tables)} tables!")
                messagebox.showinfo("Success", f"Extracted {len(all_tables)} tables to {table_folder}")
            else:
                self.update_progress(100, "ℹ️ No tables found")
                messagebox.showinfo("Info", "No tables found in the PDF")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Extraction failed: {str(e)}")

    def pdf_to_text_gui(self):
        if not self.validate_pdf_file():
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        include_metadata = messagebox.askyesno("Include Metadata", "Include PDF metadata in output?")
        
        self.run_in_thread(self.pdf_to_text_thread, self.selected_files[0], output_file, include_metadata)

    def pdf_to_text_thread(self, pdf_path, output_file, include_metadata):
        try:
            self.update_progress(0, "Extracting text from PDF...")
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                all_text = []
                
                if include_metadata:
                    metadata = pdf.metadata
                    if metadata:
                        all_text.append("=== PDF METADATA ===\n")
                        for key, value in metadata.items():
                            all_text.append(f"{key}: {value}\n")
                        all_text.append("\n")
                
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_text.append(f"\n=== Page {i+1} ===\n\n")
                        all_text.append(text)
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Processed page {i + 1}/{total_pages}")
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.writelines(all_text)
            
            self.update_progress(100, "✅ Text extraction completed!")
            messagebox.showinfo("Success", f"Text extracted to: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Text extraction failed: {str(e)}")

    def images_to_pdf_gui(self):
        image_files = filedialog.askopenfilenames(
            title="Select images to convert to PDF",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff *.gif")]
        )
        if not image_files:
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.images_to_pdf_thread, image_files, output_file)

    def images_to_pdf_thread(self, image_files, output_file):
        try:
            self.update_progress(0, "Converting images to PDF...")
            
            images = []
            total_images = len(image_files)
            
            for i, image_file in enumerate(image_files):
                img = Image.open(image_file)
                if img.mode in ['RGBA', 'LA']:
                    rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                    rgb_img.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else img.split()[-1])
                    img = rgb_img
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                images.append(img)
                
                progress = (i + 1) / total_images * 100
                self.update_progress(progress, f"Processed image {i + 1}/{total_images}")
            
            if images:
                images[0].save(output_file, save_all=True, append_images=images[1:], resolution=100.0)
            
            self.update_progress(100, "✅ PDF created successfully!")
            messagebox.showinfo("Success", f"PDF created: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"PDF creation failed: {str(e)}")

    def merge_pdfs_gui(self):
        if len(self.selected_files) < 2:
            pdf_files = filedialog.askopenfilenames(
                title="Select PDFs to merge (minimum 2)",
                filetypes=[("PDF files", "*.pdf")]
            )
            if not pdf_files or len(pdf_files) < 2:
                messagebox.showerror("Error", "Please select at least 2 PDF files")
                return
            files_to_merge = pdf_files
        else:
            files_to_merge = self.selected_files
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.merge_pdfs_thread, files_to_merge, output_file)

    def merge_pdfs_thread(self, pdf_files, output_file):
        try:
            self.update_progress(0, "Merging PDFs...")
            
            merger = PdfMerger()
            total_files = len(pdf_files)
            
            for i, pdf_file in enumerate(pdf_files):
                merger.append(pdf_file)
                progress = (i + 1) / total_files * 100
                self.update_progress(progress, f"Merged {i + 1}/{total_files} files")
            
            merger.write(output_file)
            merger.close()
            
            self.update_progress(100, "✅ PDFs merged successfully!")
            messagebox.showinfo("Success", f"PDFs merged into: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Merge failed: {str(e)}")

    def split_pdf_gui(self):
        if not self.validate_pdf_file():
            return
        
        pdf_path = self.selected_files[0]
        output_folder = filedialog.askdirectory(title="Select output folder")
        if not output_folder:
            return
        
        self.run_in_thread(self.split_pdf_thread, pdf_path, output_folder)

    def split_pdf_thread(self, pdf_path, output_folder):
        try:
            self.update_progress(0, "Splitting PDF...")
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)
                
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                split_folder = os.path.join(output_folder, f"{base_name}_pages")
                os.makedirs(split_folder, exist_ok=True)
                
                for i in range(total_pages):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    
                    output_file = os.path.join(split_folder, f"page_{i+1:03d}.pdf")
                    with open(output_file, 'wb') as output_pdf:
                        writer.write(output_pdf)
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Split page {i + 1}/{total_pages}")
            
            self.update_progress(100, "✅ PDF split successfully!")
            messagebox.showinfo("Success", f"PDF split into {total_pages} pages in {split_folder}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Split failed: {str(e)}")

    def extract_pages_gui(self):
        if not self.validate_pdf_file():
            return
        
        with pdfplumber.open(self.selected_files[0]) as pdf:
            total_pages = len(pdf.pages)
        
        pages_input = simpledialog.askstring(
            "Extract Pages", 
            f"Enter page numbers (1-{total_pages})\nExamples: 1,3,5 or 2-7 or 1,3-5,8:"
        )
        if not pages_input:
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.extract_pages_thread, self.selected_files[0], pages_input, output_file)

    def extract_pages_thread(self, pdf_path, pages_input, output_file):
        try:
            self.update_progress(0, "Extracting pages...")
            
            page_ranges = self.parse_page_ranges(pages_input)
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                writer = PdfWriter()
                
                total_pages = len(page_ranges)
                for i, page_num in enumerate(page_ranges):
                    if 1 <= page_num <= len(reader.pages):
                        writer.add_page(reader.pages[page_num - 1])
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Extracted page {i + 1}/{total_pages}")
                
                with open(output_file, 'wb') as output_pdf:
                    writer.write(output_pdf)
            
            self.update_progress(100, "✅ Pages extracted successfully!")
            messagebox.showinfo("Success", f"Pages extracted to: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Extraction failed: {str(e)}")

    def parse_page_ranges(self, pages_input):
        pages = []
        for part in pages_input.replace(' ', '').split(','):
            if '-' in part:
                start_end = part.split('-')
                if len(start_end) == 2:
                    start, end = map(int, start_end)
                    pages.extend(range(start, end + 1))
            else:
                try:
                    pages.append(int(part))
                except ValueError:
                    continue
        return sorted(set(pages))

    def protect_pdf_gui(self):
        if not self.validate_pdf_file():
            return
        
        password = simpledialog.askstring("Protect PDF", "Enter password:", show='*')
        if not password:
            return
        
        confirm = simpledialog.askstring("Confirm Password", "Confirm password:", show='*')
        if password != confirm:
            messagebox.showerror("Error", "Passwords don't match")
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.protect_pdf_thread, self.selected_files[0], output_file, password)

    def protect_pdf_thread(self, pdf_path, output_file, password):
        try:
            self.update_progress(0, "Protecting PDF...")
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                writer = PdfWriter()
                
                for page in reader.pages:
                    writer.add_page(page)
                
                writer.encrypt(password)
                
                with open(output_file, 'wb') as output_pdf:
                    writer.write(output_pdf)
            
            self.update_progress(100, "✅ PDF protected successfully!")
            messagebox.showinfo("Success", f"Protected PDF saved: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Protection failed: {str(e)}")

    def unlock_pdf_gui(self):
        if not self.validate_pdf_file():
            return
        
        password = simpledialog.askstring("Unlock PDF", "Enter password:", show='*')
        if not password:
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.unlock_pdf_thread, self.selected_files[0], output_file, password)

    def unlock_pdf_thread(self, pdf_path, output_file, password):
        try:
            self.update_progress(0, "Unlocking PDF...")
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                
                if reader.is_encrypted:
                    if reader.decrypt(password):
                        writer = PdfWriter()
                        
                        for page in reader.pages:
                            writer.add_page(page)
                        
                        with open(output_file, 'wb') as output_pdf:
                            writer.write(output_pdf)
                        
                        self.update_progress(100, "✅ PDF unlocked successfully!")
                        messagebox.showinfo("Success", f"Unlocked PDF saved: {output_file}")
                    else:
                        self.update_progress(0, "❌ Incorrect password")
                        messagebox.showerror("Error", "Incorrect password")
                else:
                    self.update_progress(0, "❌ PDF is not encrypted")
                    messagebox.showinfo("Info", "PDF is not encrypted")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Unlock failed: {str(e)}")

    def compress_pdf_gui(self):
        if not self.validate_pdf_file():
            return
        
        quality = simpledialog.askinteger(
            "Compress PDF", 
            "Enter compression level (1-100):\n100 = Best quality, 1 = Smallest size",
            minvalue=1, 
            maxvalue=100, 
            initialvalue=75
        )
        if not quality:
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        self.run_in_thread(self.compress_pdf_thread, self.selected_files[0], output_file, quality)

    def compress_pdf_thread(self, pdf_path, output_file, quality):
        try:
            self.update_progress(0, "Compressing PDF...")
            
            original_size = os.path.getsize(pdf_path) / 1024
            
            doc = pymupdf.open(pdf_path)
            
            for page in doc:
                page.get_pixmap(dpi=72 * quality / 100)
            
            doc.save(output_file, garbage=4, deflate=True, clean=True)
            doc.close()
            
            new_size = os.path.getsize(output_file) / 1024
            reduction = ((original_size - new_size) / original_size) * 100
            
            self.update_progress(100, f"✅ Compression completed!")
            messagebox.showinfo(
                "Success", 
                f"Compressed PDF saved: {output_file}\n"
                f"Original: {original_size:.1f} KB\n"
                f"Compressed: {new_size:.1f} KB\n"
                f"Reduction: {reduction:.1f}%"
            )
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Compression failed: {str(e)}")

    def rotate_pdf_gui(self):
        if not self.validate_pdf_file():
            return
        
        angle = simpledialog.askinteger(
            "Rotate PDF", 
            "Enter rotation angle (degrees):",
            minvalue=-360, 
            maxvalue=360, 
            initialvalue=90
        )
        if angle is None:
            return
        
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not output_file:
            return
        
        pages_input = simpledialog.askstring(
            "Rotate Pages", 
            "Enter page numbers to rotate (leave empty for all pages):"
        )
        
        self.run_in_thread(self.rotate_pdf_thread, self.selected_files[0], output_file, angle, pages_input)

    def rotate_pdf_thread(self, pdf_path, output_file, angle, pages_input):
        try:
            self.update_progress(0, "Rotating PDF...")
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                writer = PdfWriter()
                
                if pages_input:
                    pages_to_rotate = self.parse_page_ranges(pages_input)
                else:
                    pages_to_rotate = list(range(1, len(reader.pages) + 1))
                
                total_pages = len(reader.pages)
                rotated_count = 0
                
                for i in range(total_pages):
                    page = reader.pages[i]
                    
                    if (i + 1) in pages_to_rotate:
                        page.rotate(angle)
                        rotated_count += 1
                    
                    writer.add_page(page)
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Processed page {i + 1}/{total_pages}")
                
                with open(output_file, 'wb') as output_pdf:
                    writer.write(output_pdf)
            
            self.update_progress(100, f"✅ Rotated {rotated_count} pages!")
            messagebox.showinfo("Success", f"Rotated PDF saved: {output_file}")
        except Exception as e:
            self.update_progress(0, f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Rotation failed: {str(e)}")

def main():
    root = tk.Tk()
    app = PDFToolkitGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
