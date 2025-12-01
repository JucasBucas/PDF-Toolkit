import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
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

class PDFToolkitGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Toolkit Pro")
        self.root.geometry("900x700")
        
        # Variables
        self.current_file = tk.StringVar()
        self.progress = tk.DoubleVar()
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF Toolkit", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Browse File", command=self.browse_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="Browse Folder", command=self.browse_folder).grid(row=0, column=1, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.current_file, width=80)
        self.file_entry.grid(row=0, column=2, sticky=(tk.W, tk.E))
        
        # File list display
        file_list_frame = ttk.LabelFrame(main_frame, text="Selected Files", padding="10")
        file_list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_list_text = scrolledtext.ScrolledText(file_list_frame, height=4, width=80)
        self.file_list_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.file_list_text.insert(tk.END, "No files selected")
        self.file_list_text.config(state=tk.DISABLED)
        
        # Operations frame
        ops_frame = ttk.LabelFrame(main_frame, text="PDF Operations", padding="10")
        ops_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Row 1
        ttk.Button(ops_frame, text="PDF to Images", command=self.pdf_to_images_gui).grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="PDF to Word", command=self.pdf_to_word_gui).grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="Extract Tables", command=self.extract_tables_gui).grid(row=0, column=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Row 2
        ttk.Button(ops_frame, text="PDF to Text", command=self.pdf_to_text_gui).grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="Images to PDF", command=self.images_to_pdf_gui).grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="Merge PDFs", command=self.merge_pdfs_gui).grid(row=1, column=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Row 3
        ttk.Button(ops_frame, text="Split PDF", command=self.split_pdf_gui).grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="Extract Pages", command=self.extract_pages_gui).grid(row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(ops_frame, text="Clear Files", command=self.clear_files).grid(row=2, column=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        self.progress_label = ttk.Label(main_frame, text="Ready")
        self.progress_label.grid(row=5, column=0, columnspan=3)
        
        # Preview area
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
        preview_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=10, width=80)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(6, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

    def get_unique_filename(self, base_path, extension):
        """Generate unique filename by adding counter if file exists"""
        counter = 1
        base_name = os.path.splitext(base_path)[0]
        new_path = f"{base_name}{extension}"
        
        while os.path.exists(new_path):
            new_path = f"{base_name}_{counter}{extension}"
            counter += 1
        
        return new_path

    def get_unique_folder(self, base_folder):
        """Generate unique folder name by adding counter if folder exists"""
        counter = 1
        new_folder = base_folder
        
        while os.path.exists(new_folder):
            new_folder = f"{base_folder}_{counter}"
            counter += 1
        
        return new_folder

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.current_file.set(filename)
            self.update_file_list([filename])
            self.update_preview(filename)
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            self.current_file.set(folder)
            pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            self.update_file_list(pdf_files)
    
    def update_file_list(self, files):
        self.file_list_text.config(state=tk.NORMAL)
        self.file_list_text.delete(1.0, tk.END)
        if files:
            for f in files:
                self.file_list_text.insert(tk.END, f + '\n')
        else:
            self.file_list_text.insert(tk.END, "No files selected")
        self.file_list_text.config(state=tk.DISABLED)
    
    def clear_files(self):
        self.current_file.set("")
        self.update_file_list([])
        self.preview_text.delete(1.0, tk.END)
    
    def update_preview(self, filename):
        try:
            if filename.lower().endswith('.pdf'):
                with pdfplumber.open(filename) as pdf:
                    preview_text = f"PDF Preview: {os.path.basename(filename)}\n"
                    preview_text += f"Pages: {len(pdf.pages)}\n\n"
                    
                    if len(pdf.pages) > 0:
                        first_page = pdf.pages[0]
                        text = first_page.extract_text()
                        if text:
                            preview_text += text[:500] + "..." if len(text) > 500 else text
                    
                    self.preview_text.delete(1.0, tk.END)
                    self.preview_text.insert(1.0, preview_text)
                    
        except Exception as e:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, f"Preview error: {str(e)}")
    
    def update_progress(self, value, text=""):
        self.progress.set(value)
        if text:
            self.progress_label.config(text=text)
        self.root.update_idletasks()
    
    def run_in_thread(self, func, *args):
        thread = threading.Thread(target=func, args=args)
        thread.daemon = True
        thread.start()
    
    def pdf_to_images_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        # Create output folder with timestamp
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_folder = f"{base_name}_images"
        output_folder = self.get_unique_folder(default_folder)
        
        self.run_in_thread(self.pdf_to_images_thread, pdf_path, output_folder)
    
    def pdf_to_images_thread(self, pdf_path, output_folder):
        try:
            self.update_progress(0, "Converting PDF to images...")
            
            import pymupdf
            doc = pymupdf.open(pdf_path)
            total_pages = len(doc)
            
            os.makedirs(output_folder, exist_ok=True)
            
            for page_num in range(total_pages):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2))
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                
                # Generate unique filename for each page
                image_filename = f"page_{page_num + 1}.png"
                output_path = os.path.join(output_folder, image_filename)
                img.save(output_path)
                
                progress = (page_num + 1) / total_pages * 100
                self.update_progress(progress, f"Processed page {page_num + 1}/{total_pages}")
            
            doc.close()
            self.update_progress(100, "Conversion completed!")
            messagebox.showinfo("Success", f"PDF converted to images in {output_folder}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
    
    def pdf_to_word_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        # Generate unique output filename
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_file = f"{base_name}_converted.docx"
        output_file = self.get_unique_filename(default_file, ".docx")
        
        self.run_in_thread(self.pdf_to_word_thread, pdf_path, output_file)
    
    def pdf_to_word_thread(self, pdf_path, output_file):
        try:
            self.update_progress(0, "Converting PDF to Word...")
            
            pdf = pdfplumber.open(pdf_path)
            doc = Document()
            total_pages = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
                
                progress = (i + 1) / total_pages * 100
                self.update_progress(progress, f"Processed page {i + 1}/{total_pages}")
            
            doc.save(output_file)
            pdf.close()
            
            self.update_progress(100, "Conversion completed!")
            messagebox.showinfo("Success", f"PDF converted to Word: {output_file}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
    
    def extract_tables_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        # Create unique output folder
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_folder = f"{base_name}_tables"
        output_folder = self.get_unique_folder(default_folder)
        
        self.run_in_thread(self.extract_tables_thread, pdf_path, output_folder)
    
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
                os.makedirs(output_folder, exist_ok=True)
                
                for table_info in all_tables:
                    df = table_info['dataframe']
                    page = table_info['page']
                    table_num = table_info['table_num']
                    
                    # Generate unique filenames
                    csv_base = f"table_p{page}_t{table_num}"
                    excel_base = f"table_p{page}_t{table_num}"
                    
                    csv_filename = self.get_unique_filename(os.path.join(output_folder, csv_base), ".csv")
                    excel_filename = self.get_unique_filename(os.path.join(output_folder, excel_base), ".xlsx")
                    
                    df.to_csv(csv_filename, index=False, encoding='utf-8')
                    df.to_excel(excel_filename, index=False)
                
                if len(all_tables) > 1:
                    combined_base = "all_tables_combined"
                    combined_filename = self.get_unique_filename(os.path.join(output_folder, combined_base), ".xlsx")
                    with pd.ExcelWriter(combined_filename) as writer:
                        for table_info in all_tables:
                            sheet_name = f"Page_{table_info['page']}_Table_{table_info['table_num']}"[:31]
                            table_info['dataframe'].to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.update_progress(100, f"Extracted {len(all_tables)} tables!")
                messagebox.showinfo("Success", f"Extracted {len(all_tables)} tables to {output_folder}")
            else:
                self.update_progress(100, "No tables found")
                messagebox.showinfo("Info", "No tables found in the PDF")
                
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Extraction failed: {str(e)}")
    
    def pdf_to_text_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        # Generate unique output filename
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_file = f"{base_name}_extracted.txt"
        output_file = self.get_unique_filename(default_file, ".txt")
        
        self.run_in_thread(self.pdf_to_text_thread, pdf_path, output_file)
    
    def pdf_to_text_thread(self, pdf_path, output_file):
        try:
            self.update_progress(0, "Extracting text from PDF...")
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                all_text = []
                
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_text.append(f"--- Page {i+1} ---\n")
                        all_text.append(text)
                        all_text.append("\n\n")
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Processed page {i + 1}/{total_pages}")
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.writelines(all_text)
            
            self.update_progress(100, "Text extraction completed!")
            messagebox.showinfo("Success", f"Text extracted to: {output_file}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Text extraction failed: {str(e)}")
    
    def images_to_pdf_gui(self):
        image_files = filedialog.askopenfilenames(
            title="Select images to convert to PDF",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        if not image_files:
            return
        
        # Generate unique output filename
        default_file = "images_combined.pdf"
        output_file = self.get_unique_filename(default_file, ".pdf")
        
        self.run_in_thread(self.images_to_pdf_thread, image_files, output_file)
    
    def images_to_pdf_thread(self, image_files, output_file):
        try:
            self.update_progress(0, "Converting images to PDF...")
            
            images = []
            total_images = len(image_files)
            
            for i, image_file in enumerate(image_files):
                img = Image.open(image_file)
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                images.append(img)
                
                progress = (i + 1) / total_images * 100
                self.update_progress(progress, f"Processed image {i + 1}/{total_images}")
            
            if images:
                images[0].save(output_file, save_all=True, append_images=images[1:])
            
            self.update_progress(100, "PDF created successfully!")
            messagebox.showinfo("Success", f"PDF created: {output_file}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"PDF creation failed: {str(e)}")
    
    def merge_pdfs_gui(self):
        pdf_files = filedialog.askopenfilenames(
            title="Select PDFs to merge",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_files or len(pdf_files) < 2:
            messagebox.showerror("Error", "Please select at least 2 PDF files to merge")
            return
        
        # Generate unique output filename
        default_file = "merged_document.pdf"
        output_file = self.get_unique_filename(default_file, ".pdf")
        
        self.run_in_thread(self.merge_pdfs_thread, pdf_files, output_file)
    
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
            
            self.update_progress(100, "PDFs merged successfully!")
            messagebox.showinfo("Success", f"PDFs merged into: {output_file}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Merge failed: {str(e)}")
    
    def split_pdf_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        # Create unique output folder
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_folder = f"{base_name}_split_pages"
        output_folder = self.get_unique_folder(default_folder)
        
        self.run_in_thread(self.split_pdf_thread, pdf_path, output_folder)
    
    def split_pdf_thread(self, pdf_path, output_folder):
        try:
            self.update_progress(0, "Splitting PDF...")
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                total_pages = len(reader.pages)
                
                os.makedirs(output_folder, exist_ok=True)
                
                for i in range(total_pages):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    
                    # Generate unique filename for each page
                    page_filename = f"page_{i+1}.pdf"
                    output_file = os.path.join(output_folder, page_filename)
                    
                    with open(output_file, 'wb') as output_pdf:
                        writer.write(output_pdf)
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Split page {i + 1}/{total_pages}")
            
            self.update_progress(100, "PDF split successfully!")
            messagebox.showinfo("Success", f"PDF split into {total_pages} pages in {output_folder}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Split failed: {str(e)}")
    
    def extract_pages_gui(self):
        pdf_path = self.current_file.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file")
            return
        
        pages_input = tk.simpledialog.askstring("Extract Pages", "Enter page numbers (e.g., 1,3,5-8):")
        if not pages_input:
            return
        
        # Generate unique output filename
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        default_file = f"{base_name}_extracted_pages.pdf"
        output_file = self.get_unique_filename(default_file, ".pdf")
        
        self.run_in_thread(self.extract_pages_thread, pdf_path, pages_input, output_file)
    
    def extract_pages_thread(self, pdf_path, pages_input, output_file):
        try:
            self.update_progress(0, "Extracting pages...")
            
            page_ranges = self.parse_page_ranges(pages_input)
            
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                writer = PdfWriter()
                
                total_pages = len(page_ranges)
                for i, page_num in enumerate(page_ranges):
                    if 0 <= page_num - 1 < len(reader.pages):
                        writer.add_page(reader.pages[page_num - 1])
                    
                    progress = (i + 1) / total_pages * 100
                    self.update_progress(progress, f"Extracted page {i + 1}/{total_pages}")
                
                with open(output_file, 'wb') as output_pdf:
                    writer.write(output_pdf)
            
            self.update_progress(100, "Pages extracted successfully!")
            messagebox.showinfo("Success", f"Pages extracted to: {output_file}")
            
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            messagebox.showerror("Error", f"Extraction failed: {str(e)}")
    
    def parse_page_ranges(self, pages_input):
        pages = []
        for part in pages_input.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages.extend(range(start, end + 1))
            else:
                pages.append(int(part))
        return pages

def main():
    root = tk.Tk()
    app = PDFToolkitGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
