import os
import pandas as pd

def process_excel_file(file_path, label):
    output_lines = []
    file_name = os.path.basename(file_path)
    output_lines.append(f"{label}:")
    
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        output_lines.append(f"Error reading file: {e}")
        return "\n".join(output_lines)
    
    if "Form Number" not in df.columns or "Form Title" not in df.columns:
        output_lines.append("Error: 'Form Number' or 'Form Title' column not found in this file.")
        return "\n".join(output_lines)
    
    for idx, row in df.iterrows():
        form_number = str(row["Form Number"]).strip()
        form_title = str(row["Form Title"]).strip()
        output_lines.append(f"{form_number} {form_title}")
    
    output_lines.append("")
    return "\n".join(output_lines)

def process_forms(self):
    folder_entry = self.entry1() if callable(self.entry1) else self.entry1.get()
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    os.makedirs(report_directory, exist_ok=True)
    files_with_labels = {

        "matching_RevisedForms.xlsx": "Revised Forms",
        "matching_AddedForms.xlsx": "New Forms",
        "matching_DeletedForms.xlsx": "Deleted Forms"        
    }
    
    all_extracted_data = []
    for file_name, label in files_with_labels.items():
        file_path = os.path.join(report_directory, file_name)
        if os.path.exists(file_path):
            extracted_data = process_excel_file(file_path, label)
            all_extracted_data.append(extracted_data)

        else:
            all_extracted_data.append(f"{label}:\nNA\n")
    
    final_output = "\n".join(all_extracted_data)
    output_text_file = os.path.join(report_directory, "CollectedForms.txt")
    with open(output_text_file, "w", encoding="utf-8") as f:
        f.write(final_output)
    return final_output
