import tkinter as tk
from tkinter import filedialog, messagebox
import os
from docx import Document
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.styles import Font
from pptx import Presentation

# --- Core Logic for Font Changing ---

def change_word_font(path, new_font_name):
    """Changes the font for all text in a .docx file."""
    doc = Document(path)
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = new_font_name
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.name = new_font_name
    
    return doc

def change_excel_font(path, new_font_name):
    """Changes the font for all cells in a .xlsx file."""
    workbook = load_workbook(path)
    new_font = Font(name=new_font_name)
    
    for sheetname in workbook.sheetnames:
        worksheet = workbook[sheetname]
        for row in worksheet.iter_rows():
            for cell in row:
                # We apply the font to all cells, even empty ones, to ensure consistency
                # when data is added later.
                cell.font = new_font
                
    return workbook

def change_ppt_font(path, new_font_name):
    """Changes the font for all text in a .pptx file."""
    prs = Presentation(path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = new_font_name
    return prs

# --- GUI Application ---

class FontUnifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Font Unifier")
        self.root.geometry("500x250")

        self.file_path = tk.StringVar()
        self.font_name = tk.StringVar(value="Meiryo UI")

        # File Selection Frame
        file_frame = tk.Frame(self.root, pady=10)
        file_frame.pack(fill="x", padx=10)
        
        tk.Label(file_frame, text="File:", width=10).pack(side="left")
        self.file_entry = tk.Entry(file_frame, textvariable=self.file_path, state="readonly")
        self.file_entry.pack(side="left", expand=True, fill="x")
        tk.Button(file_frame, text="Browse...", command=self.browse_file).pack(side="right", padx=5)

        # Font Selection Frame
        font_frame = tk.Frame(self.root, pady=10)
        font_frame.pack(fill="x", padx=10)
        
        tk.Label(font_frame, text="Target Font:", width=10).pack(side="left")
        self.font_entry = tk.Entry(font_frame, textvariable=self.font_name)
        self.font_entry.pack(side="left", expand=True, fill="x")

        # Action Frame
        action_frame = tk.Frame(self.root, pady=20)
        action_frame.pack()
        
        self.start_button = tk.Button(action_frame, text="Start Processing", command=self.process_file, width=20, height=2)
        self.start_button.pack()
        
        # Status Label
        self.status_label = tk.Label(self.root, text="", fg="green")
        self.status_label.pack(pady=5)

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Office Files", "*.docx *.xlsx *.pptx"),
                ("Word Documents", "*.docx"),
                ("Excel Workbooks", "*.xlsx"),
                ("PowerPoint Presentations", "*.pptx"),
                ("All files", "*.*")
            ]
        )
        if path:
            self.file_path.set(path)
            self.status_label.config(text="")

    def process_file(self):
        path = self.file_path.get()
        font = self.font_name.get()

        if not path:
            messagebox.showerror("Error", "Please select a file first.")
            return
        if not font:
            messagebox.showerror("Error", "Please enter a target font name.")
            return

        self.status_label.config(text="Processing...", fg="blue")
        self.root.update_idletasks()

        try:
            file_dir, file_name = os.path.split(path)
            name, ext = os.path.splitext(file_name)
            output_path = os.path.join(file_dir, f"{name}_modified{ext}")

            if ext == ".docx":
                modified_doc = change_word_font(path, font)
                modified_doc.save(output_path)
            elif ext == ".xlsx":
                modified_workbook = change_excel_font(path, font)
                modified_workbook.save(output_path)
            elif ext == ".pptx":
                modified_prs = change_ppt_font(path, font)
                modified_prs.save(output_path)
            else:
                messagebox.showerror("Error", f"Unsupported file type: {ext}")
                self.status_label.config(text="", fg="red")
                return
            
            self.status_label.config(text=f"Success! Saved to {output_path}", fg="green")
            messagebox.showinfo("Success", "File processed successfully and saved as: " + output_path)

        except Exception as e:
            self.status_label.config(text="An error occurred.", fg="red")
            messagebox.showerror("Error", "An error occurred during processing: " + str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FontUnifierApp(root)
    root.mainloop()
