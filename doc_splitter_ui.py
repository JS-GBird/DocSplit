from doc_splitter import split_document
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os

class DocumentSplitterUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Document Splitter")
        self.geometry("600x400")
        self.configure(fg_color="#f5f5f5")

        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Variables
        self.input_file_path = ctk.StringVar()
        self.output_dir_path = ctk.StringVar(value=str(Path.home() / "Documents" / "split_documents"))

        self.create_widgets()

    def create_widgets(self):
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame, 
            text="Document Splitter",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="#1a1a1a"
        )
        title_label.pack(pady=10)

        # Input file section
        input_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=10)
        input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        input_label = ctk.CTkLabel(
            input_frame,
            text="Input Document:",
            font=ctk.CTkFont(size=14),
            text_color="#333333"
        )
        input_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        input_entry = ctk.CTkEntry(
            input_frame,
            textvariable=self.input_file_path,
            width=400,
            placeholder_text="Select your input .docx file"
        )
        input_entry.grid(row=1, column=0, padx=10, pady=5)

        browse_button = ctk.CTkButton(
            input_frame,
            text="Browse",
            command=self.browse_input_file,
            width=100,
            fg_color="#2d7ae5",
            hover_color="#1c5bb0"
        )
        browse_button.grid(row=1, column=1, padx=10, pady=5)

        # Output directory section
        output_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=10)
        output_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        output_label = ctk.CTkLabel(
            output_frame,
            text="Output Directory:",
            font=ctk.CTkFont(size=14),
            text_color="#333333"
        )
        output_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        output_entry = ctk.CTkEntry(
            output_frame,
            textvariable=self.output_dir_path,
            width=400,
            placeholder_text="Select output directory"
        )
        output_entry.grid(row=1, column=0, padx=10, pady=5)

        output_button = ctk.CTkButton(
            output_frame,
            text="Browse",
            command=self.browse_output_dir,
            width=100,
            fg_color="#2d7ae5",
            hover_color="#1c5bb0"
        )
        output_button.grid(row=1, column=1, padx=10, pady=5)

        # Process button
        process_button = ctk.CTkButton(
            self,
            text="Split Document",
            command=self.process_document,
            width=200,
            height=40,
            fg_color="#28a745",
            hover_color="#218838",
            font=ctk.CTkFont(size=15, weight="bold")
        )
        process_button.grid(row=3, column=0, pady=20)

        # Status frame
        self.status_label = ctk.CTkLabel(
            self,
            text="Ready to process",
            font=ctk.CTkFont(size=12),
            text_color="#666666"
        )
        self.status_label.grid(row=4, column=0, pady=(0, 20))

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)

    def browse_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir_path.set(dir_path)

    def process_document(self):
        input_path = self.input_file_path.get()
        output_dir = self.output_dir_path.get()

        if not input_path or not output_dir:
            messagebox.showerror("Error", "Please select both input file and output directory")
            return

        try:
            self.status_label.configure(text="Processing...", text_color="#ffa500")
            self.update()
            
            # Process the document
            split_document(input_path, output_dir)
            
            # Count number of generated files
            num_students = len([f for f in os.listdir(output_dir) 
                              if f.startswith('Student_') and f.endswith('.docx')])
            
            # Update UI
            self.status_label.configure(text="Documents successfully split!", text_color="#28a745")
            messagebox.showinfo("Success", 
                f"Documents have been successfully split!\n{num_students} student files created.")
                
        except Exception as e:
            self.status_label.configure(text="Error occurred", text_color="#dc3545")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    # Set appearance mode and default color theme
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    app = DocumentSplitterUI()
    app.mainloop()

if __name__ == "__main__":
    main() 