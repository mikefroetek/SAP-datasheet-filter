"""
Interactive Excel Level Processor
User-friendly interface for processing Aptive files to BOM templates.
"""

import os
import pandas as pd
from openpyxl import load_workbook
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import simpledialog

class ExcelProcessorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Level Processor")
        self.root.geometry("600x500")
        
        self.aptive_file = ""
        self.bom_template = ""
        self.output_path = ""
        
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the GUI interface"""
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Level Processor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Aptive file selection
        ttk.Label(files_frame, text="Aptive BOOM File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.aptive_label = ttk.Label(files_frame, text="No file selected", 
                                     foreground="gray")
        self.aptive_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(files_frame, text="Browse", 
                  command=self.select_aptive_file).grid(row=0, column=2, pady=2)
        
        # BOM template selection  
        ttk.Label(files_frame, text="BOM Template:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.bom_label = ttk.Label(files_frame, text="No file selected", 
                                  foreground="gray")
        self.bom_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(files_frame, text="Browse", 
                  command=self.select_bom_template).grid(row=1, column=2, pady=2)
        
        # Output path selection
        ttk.Label(files_frame, text="Output Folder:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.output_label = ttk.Label(files_frame, text="Auto-generate", 
                                     foreground="gray")
        self.output_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        ttk.Button(files_frame, text="Browse", 
                  command=self.select_output_folder).grid(row=2, column=2, pady=2)
        
        # Configure column weights
        files_frame.columnconfigure(1, weight=1)
        
        # Settings section
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Column mapping info
        mapping_text = """Column Mapping:
Aptive Column A (Level/Szint) → BOM Column E
Aptive Column B → BOM Column N  
Aptive Column D → BOM Column O
Aptive Column E → BOM Column P"""
        
        ttk.Label(settings_frame, text=mapping_text, 
                 font=("Courier", 9)).grid(row=0, column=0, sticky=tk.W)
        
        # Process button
        process_btn = ttk.Button(main_frame, text="Process Files", 
                               command=self.process_files, style="Accent.TButton")
        process_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status text
        self.status_text = tk.Text(main_frame, height=10, width=70)
        self.status_text.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def log_message(self, message):
        """Add message to status text"""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def select_aptive_file(self):
        """Select Aptive BOOM file"""
        filetypes = [
            ("Excel files", "*.xlsx *.xlsb *.xls"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Aptive BOOM File",
            filetypes=filetypes,
            initialdir=r"d:\SAP"
        )
        
        if filename:
            self.aptive_file = filename
            self.aptive_label.config(text=os.path.basename(filename), foreground="black")
            
    def select_bom_template(self):
        """Select BOM Template file"""
        filetypes = [
            ("Excel files", "*.xlsx *.xlsb *.xls"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select BOM Template File",
            filetypes=filetypes,
            initialdir=r"d:\SAP"
        )
        
        if filename:
            self.bom_template = filename  
            self.bom_label.config(text=os.path.basename(filename), foreground="black")
            
    def select_output_folder(self):
        """Select output folder"""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=r"d:\SAP"
        )
        
        if folder:
            self.output_path = folder
            self.output_label.config(text=folder, foreground="black")
            
    def process_files(self):
        """Process the selected files"""
        
        # Validate selections
        if not self.aptive_file:
            messagebox.showerror("Error", "Please select an Aptive BOOM file")
            return
            
        if not self.bom_template:
            messagebox.showerror("Error", "Please select a BOM Template file")
            return
            
        # Clear status
        self.status_text.delete(1.0, tk.END)
        
        # Start progress bar
        self.progress.start(10)
        
        try:
            self.log_message("🚀 Starting Excel processing...")
            self.log_message(f"📖 Aptive file: {os.path.basename(self.aptive_file)}")
            self.log_message(f"📋 Template: {os.path.basename(self.bom_template)}")
            
            # Generate output path
            if not self.output_path:
                output_dir = os.path.dirname(self.bom_template)
            else:
                output_dir = self.output_path
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"BOM_Processed_{timestamp}.xlsx")
            
            # Process files
            self.process_excel_files(self.aptive_file, self.bom_template, output_file)
            
            # Success message
            self.progress.stop()
            self.log_message(f"✅ Processing completed successfully!")
            self.log_message(f"📁 Output: {output_file}")
            
            messagebox.showinfo("Success", 
                              f"Processing completed!\n\nOutput file:\n{os.path.basename(output_file)}")
            
        except Exception as e:
            self.progress.stop()
            self.log_message(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            
    def process_excel_files(self, aptive_path, template_path, output_path):
        """Core processing logic"""
        
        # Read Aptive file
        self.log_message("📖 Reading Aptive data...")
        
        if aptive_path.endswith('.xlsb'):
            df = pd.read_excel(aptive_path, engine='pyxlsb')
        else:
            df = pd.read_excel(aptive_path)
            
        self.log_message(f"✓ Loaded {len(df)} rows")
        
        # Process level data
        self.log_message("🔄 Processing level data...")
        level_data = {}
        
        for index in range(1, len(df)):  # Skip header
            row = df.iloc[index]
            
            # Get level from column A
            level_value = str(row.iloc[0]).strip()
            
            # Skip empty rows
            if pd.isna(row.iloc[0]) or level_value == "" or level_value == "nan":
                continue
                
            # Extract numeric level
            numbers = re.findall(r'\d+', level_value)
            if not numbers:
                continue
                
            level = int(numbers[0])
            
            # Get data from columns B, D, E
            col_b = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            col_d = str(row.iloc[3]) if len(row) > 3 and not pd.isna(row.iloc[3]) else ""
            col_e = str(row.iloc[4]) if len(row) > 4 and not pd.isna(row.iloc[4]) else ""
            
            # Store by level
            if level not in level_data:
                level_data[level] = []
                
            level_data[level].append({
                'level_text': level_value,
                'col_b': col_b,
                'col_d': col_d,
                'col_e': col_e
            })
        
        # Log level summary
        for level in sorted(level_data.keys()):
            self.log_message(f"  Level {level}: {len(level_data[level])} items")
            
        # Write to BOM template
        self.log_message("✏️ Writing to BOM template...")
        
        workbook = load_workbook(template_path)
        worksheet = workbook.active
        
        current_row = 2  # Start from row 2
        
        for level in sorted(level_data.keys()):
            self.log_message(f"  Writing Level {level}...")
            
            for item in level_data[level]:
                # Map columns: Aptive A,B,D,E -> BOM E,N,O,P
                worksheet[f'E{current_row}'] = item['level_text']
                worksheet[f'N{current_row}'] = item['col_b']
                worksheet[f'O{current_row}'] = item['col_d']
                worksheet[f'P{current_row}'] = item['col_e']
                
                current_row += 1
                
            # Blank row between levels
            current_row += 1
            
        # Save output
        workbook.save(output_path)
        self.log_message(f"💾 Saved to: {os.path.basename(output_path)}")
        
    def run(self):
        """Start the GUI"""
        self.root.mainloop()

def main():
    """Main function"""
    try:
        app = ExcelProcessorGUI()
        app.run()
    except Exception as e:
        print(f"Error starting application: {str(e)}")

if __name__ == "__main__":
    main()