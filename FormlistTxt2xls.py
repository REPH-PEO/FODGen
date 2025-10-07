import re
import pandas as pd
import os

def extracttxt_to_excelcurrent(self):
    txtfile = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    # base_dir = os.path.join(initial_dir, txtfile)
    base_dir = os.path.join(initial_dir, txtfile, "FOD")    
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    pattern = re.compile(r"^Current_FORMLIST\.txt$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(report_dir):
        if pattern.match(filename):
            input_file = os.path.join(report_dir, filename)
            break

    if input_file is None:
        print("No matching file found in current excel.")
        return

    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = []
    line_pattern = re.compile(
        r'([\w\-.]+\.rtf)\|\s*(?:(Form\s+[\w\-.]+)|(§\s+[\w\-.]+)|([\w\-.]+))\|\s*(.+)')
    for line in lines:
        match = line_pattern.match(line.strip())
        if match:
            file_name = match.group(1)
            # Select the first non-None value among groups 2, 3, and 4 for the form number.
            form_number = match.group(2) or match.group(3) or match.group(4)
            form_title = match.group(5)
            data.append([file_name, form_number, form_title])
    
    df = pd.DataFrame(data, columns=['File Name', 'Form Number', 'Form Title'])
    excel_output = os.path.join(report_dir, "Current_FORMLIST.xlsx")
    df.to_excel(excel_output, index=False)
    print(f"Extracted {len(df)} entries to {excel_output}")

def extracttxt_to_excelprev(self):
    txtfile = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    
    # base_dir = os.path.join(initial_dir, txtfile)
    base_dir = os.path.join(initial_dir, txtfile, "FOD") 
    
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    pattern = re.compile(r"^Previous_FORMLIST\.txt$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(report_dir):
        if pattern.match(filename):
            input_file = os.path.join(report_dir, filename)
            break

    if input_file is None:
        print("No matching file found previous excel.")

    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = []
    line_pattern = re.compile(
        r'([\w\-.]+\.rtf)\|\s*(?:(Form\s+[\w\-.]+)|(§\s+[\w\-.]+)|([\w\-.]+))\|\s*(.+)')
    for line in lines:
        match = line_pattern.match(line.strip())
        if match:
            file_name = match.group(1)
            form_number = match.group(2) or match.group(3) or match.group(4)
            form_title = match.group(5)
            data.append([file_name, form_number, form_title])
    
    df = pd.DataFrame(data, columns=['File Name', 'Form Number', 'Form Title'])
    excel_output = os.path.join(report_dir, "Previous_FORMLIST.xlsx")
    df.to_excel(excel_output, index=False)
    print(f"Extracted {len(df)} entries to {excel_output}")

def extracttxt_to_exceladd(self):
    txtfile = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    # base_dir = os.path.join(initial_dir, txtfile)
    base_dir = os.path.join(initial_dir, txtfile, "FOD") 
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    pattern = re.compile(r"^AddedForms.txt$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(report_dir):
        if pattern.match(filename):
            input_file = os.path.join(report_dir, filename)
            break

    if input_file is None:
        print("No matching file found in added forms.")
        return

    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = []
    for line in lines:
        match = re.match(r'^(?!.*[\s_])[A-Za-z0-9-]+\.rtf$', line.strip())
        if match:
            file_name = match.groups(0)
            data.append([file_name])
    
    df = pd.DataFrame(data, columns=['File Name'])
    excel_output = os.path.join(report_dir, "AddedForms.xlsx")
    df.to_excel(excel_output, index=False)
    print(f"Extracted {len(df)} entries to {excel_output}")    


def extracttxt_to_exceldeleted(self):
    txtfile = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    # base_dir = os.path.join(initial_dir, txtfile)
    base_dir = os.path.join(initial_dir, txtfile, "FOD") 
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)
    pattern = re.compile(r"^DeletedForms\.txt$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(report_dir):
        if pattern.match(filename):
            input_file = os.path.join(report_dir, filename)
            break

    if input_file is None:
        print("No matching file found in deleted forms.")
        return

    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = []
    file_pattern = re.compile(r'^(?!.*[\s_])[A-Za-z0-9-]+\.rtf$', re.IGNORECASE)
    for line in lines:
        line = line.strip()
        match = file_pattern.match(line)
        if match:
            file_name = match.group(0)
            data.append([file_name])
    
    df = pd.DataFrame(data, columns=['File Name'])
    excel_output = os.path.join(report_dir, "DeletedForms.xlsx")
    df.to_excel(excel_output, index=False)
    print(f"Extracted {len(df)} entries to {excel_output}")

def extracttxt_to_excelrevised(self):
    txtfile = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    # base_dir = os.path.join(initial_dir, txtfile)
    base_dir = os.path.join(initial_dir, txtfile, "FOD") 
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)
    pattern = re.compile(r"^RevisedForms\.txt$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(report_dir):
        if pattern.match(filename):
            input_file = os.path.join(report_dir, filename)
            break

    if input_file is None:
        print("No matching file found in revised forms.")
        return

    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = []
    file_pattern = re.compile(r'^(?!.*[\s_])[A-Za-z0-9-]+\.txt$', re.IGNORECASE)
    for line in lines:
        line = line.strip()
        match = file_pattern.match(line)
        if match:
            file_name = match.group(0)
            data.append([file_name])
    
    df = pd.DataFrame(data, columns=['File Name'])
    excel_output = os.path.join(report_dir, "RevisedForms.xlsx")
    df.to_excel(excel_output, index=False)
    print(f"Extracted {len(df)} entries to {excel_output}")