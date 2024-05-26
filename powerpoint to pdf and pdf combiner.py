import os
import win32com.client
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# PowerPoint to PDF conversion function
def convert_ppt_to_pdf(ppt_path, pdf_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.WindowState = 2  # Minimize the PowerPoint application window
    
    try:
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        presentation.SaveAs(pdf_path, FileFormat=32)  # 32 is the file format for PDFs
        presentation.Close()
        print(f"Converted {ppt_path} to {pdf_path}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        powerpoint.Quit()

# Bulk PowerPoint to PDF conversion function
def bulk_convert_ppt_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for filename in os.listdir(input_folder):
        if filename.endswith(".ppt") or filename.endswith(".pptx"):
            ppt_path = os.path.join(input_folder, filename)
            pdf_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.pdf')
            convert_ppt_to_pdf(ppt_path, pdf_path)

# PDF combining function
def combine_pdfs(pdf_list, output_directory, output_filename):
    pdf_writer = PyPDF2.PdfWriter()
    
    try:
        for pdf in pdf_list:
            pdf_reader = PyPDF2.PdfReader(pdf)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)
        
        output_path = os.path.join(output_directory, output_filename)
        with open(output_path, 'wb') as out_pdf:
            pdf_writer.write(out_pdf)
        print(f"Combined PDFs saved as {output_path}")
        return True, output_path
    except Exception as e:
        print(f"Error combining PDFs: {e}")
        return False, None

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint to PDF and PDF Combiner")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(padx=10, pady=10, expand=True)

        self.create_convert_tab()
        self.create_combine_tab()

    def create_convert_tab(self):
        convert_tab = ttk.Frame(self.notebook)
        self.notebook.add(convert_tab, text="Convert PowerPoint to PDF")

        self.input_folder_convert = ""
        self.output_folder_convert = ""

        input_button = tk.Button(convert_tab, text="Select Input Folder", command=self.select_input_folder_convert)
        input_button.pack(pady=5)

        output_button = tk.Button(convert_tab, text="Select Output Folder", command=self.select_output_folder_convert)
        output_button.pack(pady=5)

        convert_button = tk.Button(convert_tab, text="Convert", command=self.convert)
        convert_button.pack(pady=20)

    def select_input_folder_convert(self):
        self.input_folder_convert = filedialog.askdirectory(title="Select Folder Containing PowerPoint Files").replace('/', '\\')
        if self.input_folder_convert:
            messagebox.showinfo("Selected Input Folder", f"Input Folder: {self.input_folder_convert}")

    def select_output_folder_convert(self):
        self.output_folder_convert = filedialog.askdirectory(title="Select Folder to Save PDF Files").replace('/', '\\')
        if self.output_folder_convert:
            messagebox.showinfo("Selected Output Folder", f"Output Folder: {self.output_folder_convert}")

    def convert(self):
        if not self.input_folder_convert or not self.output_folder_convert:
            messagebox.showwarning("Missing Folders", "Please select both input and output folders.")
            return
        
        bulk_convert_ppt_to_pdf(self.input_folder_convert, self.output_folder_convert)
        messagebox.showinfo("Conversion Complete", "All PowerPoint files have been converted to PDF.")

    def create_combine_tab(self):
        combine_tab = ttk.Frame(self.notebook)
        self.notebook.add(combine_tab, text="Combine PDFs")

        self.pdf_list = []
        self.output_directory_combine = ""
        self.output_filename = "combined.pdf"

        select_button = tk.Button(combine_tab, text="Select PDFs", command=self.select_pdfs)
        select_button.pack(pady=5)

        output_button = tk.Button(combine_tab, text="Select Output Directory", command=self.select_output_directory_combine)
        output_button.pack(pady=5)

        self.output_label = tk.Label(combine_tab, text="No output directory selected")
        self.output_label.pack(pady=5)

        combine_button = tk.Button(combine_tab, text="Combine PDFs", command=self.combine_pdfs)
        combine_button.pack(pady=20)

    def select_pdfs(self):
        self.pdf_list = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])
        if self.pdf_list:
            pdf_names = "\n".join(self.pdf_list)
            messagebox.showinfo("Selected PDFs", f"Selected Files:\n{pdf_names}")

    def select_output_directory_combine(self):
        self.output_directory_combine = filedialog.askdirectory(title="Select Output Directory")
        if self.output_directory_combine:
            self.output_label.config(text=f"Output Directory: {self.output_directory_combine}")

    def combine_pdfs(self):
        if not self.pdf_list:
            messagebox.showwarning("No PDFs Selected", "Please select PDF files to combine.")
            return

        if not self.output_directory_combine:
            messagebox.showwarning("No Output Directory Selected", "Please select an output directory.")
            return

        output_filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Combined PDF As", initialdir=self.output_directory_combine, initialfile=self.output_filename)
        if not output_filename:
            return

        success, output_path = combine_pdfs(self.pdf_list, self.output_directory_combine, os.path.basename(output_filename))
        if success:
            messagebox.showinfo("PDF Combine Complete", f"Combined PDF saved as {output_path}")
            self.output_label.config(text=f"Output File: {output_path}")
        else:
            messagebox.showerror("Error", "Failed to combine PDF files.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
