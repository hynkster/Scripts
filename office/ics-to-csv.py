import tkinter as tk
from tkinter import filedialog, messagebox
import icalendar
import csv
import os
from datetime import datetime
import shutil
import pandas as pd

class ICSConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Calendar Converter")
        self.window.geometry("500x400")
        self.setup_ui()
        
    def setup_ui(self):
        # Welcome message
        welcome_text = """Welcome to Calendar Converter!
        
This tool helps you convert calendar files (ICS) to spreadsheet format (CSV/Excel).
Click 'Select File' to choose your calendar file."""
        
        tk.Label(self.window, text=welcome_text, wraplength=450, pady=20).pack()
        
        # File selection button
        select_button = tk.Button(self.window, text="Select File", command=self.select_file,
                                width=20, height=2)
        select_button.pack(pady=20)
        
        # Options frame
        options_frame = tk.LabelFrame(self.window, text="Export Options", padx=10, pady=10)
        options_frame.pack(pady=20, padx=20, fill="x")
        
        # CSV delimiter option
        self.delimiter_var = tk.StringVar(value=",")
        tk.Label(options_frame, text="CSV Delimiter:").pack()
        tk.Radiobutton(options_frame, text="Comma (,)", variable=self.delimiter_var, value=",").pack()
        tk.Radiobutton(options_frame, text="Semicolon (;)", variable=self.delimiter_var, value=";").pack()
        tk.Radiobutton(options_frame, text="Tab", variable=self.delimiter_var, value="\t").pack()
        
        # Excel option
        self.excel_var = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="Create Excel file instead of CSV", 
                      variable=self.excel_var).pack(pady=10)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_label = tk.Label(self.window, textvariable=self.status_var, wraplength=450)
        status_label.pack(side="bottom", pady=10)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Calendar File",
            filetypes=[("Calendar files", "*.ics"), ("All files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)
            
    def validate_ics_file(self, file_path):
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError("The selected file does not exist.")
                
            # Check file size
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                raise ValueError("The selected file is empty.")
                
            # Check if it's a valid ICS file
            with open(file_path, 'rb') as file:
                calendar = icalendar.Calendar.from_ical(file.read())
                
            # Verify it has events
            events = list(calendar.walk('VEVENT'))
            if not events:
                raise ValueError("No calendar events found in the file.")
                
            return True, len(events)
            
        except icalendar.parser.ParseError:
            raise ValueError("Invalid ICS file format. Please ensure this is a valid calendar file.")
        except Exception as e:
            raise Exception(f"Error validating file: {str(e)}")
            
    def create_backup(self, file_path):
        backup_path = f"{file_path}.backup"
        shutil.copy2(file_path, backup_path)
        return backup_path
        
    def process_file(self, file_path):
        try:
            self.status_var.set("Validating file...")
            is_valid, event_count = self.validate_ics_file(file_path)
            
            # Create backup
            self.status_var.set("Creating backup...")
            backup_path = self.create_backup(file_path)
            
            # Process events
            self.status_var.set("Processing events...")
            output_data = []
            event_counter = 0
            
            with open(file_path, 'rb') as file:
                calendar = icalendar.Calendar.from_ical(file.read())
                
                for component in calendar.walk('VEVENT'):
                    event = {
                        'Subject': str(component.get('summary', 'No Title')),
                        'Start Date': component.get('dtstart').dt.strftime('%Y-%m-%d %H:%M:%S'),
                        'End Date': component.get('dtend').dt.strftime('%Y-%m-%d %H:%M:%S'),
                        'Description': str(component.get('description', '')),
                        'Location': str(component.get('location', '')),
                        'Status': str(component.get('status', '')),
                    }
                    output_data.append(event)
                    event_counter += 1
                    self.status_var.set(f"Processing event {event_counter} of {event_count}...")
            
            # Save output file
            output_dir = os.path.dirname(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            if self.excel_var.get():
                output_path = os.path.join(output_dir, f"{base_name}.xlsx")
                df = pd.DataFrame(output_data)
                df.to_excel(output_path, index=False)
            else:
                output_path = os.path.join(output_dir, f"{base_name}.csv")
                with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=output_data[0].keys(), 
                                         delimiter=self.delimiter_var.get())
                    writer.writeheader()
                    writer.writerows(output_data)
            
            # Verify conversion
            self.status_var.set("Verifying conversion...")
            if self.verify_conversion(output_path, event_counter):
                success_message = f"""Conversion completed successfully!

- Events processed: {event_counter}
- Output file: {output_path}
- Backup created: {backup_path}

Would you like to open the converted file?"""
                if messagebox.askyesno("Success", success_message):
                    os.startfile(output_path) if os.name == 'nt' else os.system(f'open "{output_path}"')
                    
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Ready")
            
    def verify_conversion(self, output_path, expected_count):
        try:
            if output_path.endswith('.xlsx'):
                df = pd.read_excel(output_path)
                actual_count = len(df)
            else:
                with open(output_path, 'r', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    actual_count = sum(1 for row in reader) - 1  # Subtract header row
                    
            if actual_count != expected_count:
                raise ValueError(f"Verification failed: Expected {expected_count} events, "
                              f"but found {actual_count} in output file.")
                              
            return True
            
        except Exception as e:
            messagebox.showerror("Verification Error", str(e))
            return False
            
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    try:
        converter = ICSConverter()
        converter.run()
    except Exception as e:
        print(f"Error starting the application: {str(e)}")
        print("\nPlease ensure you have all required packages installed:")
        print("python3 -m pip install --user icalendar pandas")