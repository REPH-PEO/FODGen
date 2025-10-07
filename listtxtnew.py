import os

def listrevision_txt_files(self):
    entry1_value = self.entry1.get()    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    directory_path = os.path.join(initial_dir, entry1_value,"FOD", 'Report')
    print(f"Directory path: {directory_path}")
    if not os.path.isdir(directory_path):
        print(f"Directory does not exist: {directory_path}")
        return
    txt_file_names = [
        filename for filename in os.listdir(directory_path)
        if filename.lower().endswith(".txt") and os.path.getsize(os.path.join(directory_path, filename)) > 0
    ]
    print(f"Found non-empty .txt files: {txt_file_names}")
    os.makedirs(directory_path, exist_ok=True)
    output_file = os.path.join(directory_path, 'RevisedForms.txt')
    report_lines = ["List of non-empty .txt files found:"]
    report_lines.extend(txt_file_names)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"Report saved to {output_file}")
