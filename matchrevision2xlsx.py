import pandas as pd
import os

def strip_extension(filename):
    return os.path.splitext(str(filename).strip())[0]

def extract_form_infoadd(self):
    folder_entry = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)   
    
    files_extracted = os.path.join(report_dir, "Current_FORMLIST.xlsx")
    file_list = os.path.join(report_dir, "AddedForms.xlsx")
    
    if not os.path.exists(files_extracted):
        print(f"Error: {files_extracted} not found.")
        return pd.DataFrame()
    if not os.path.exists(file_list):
        print(f"Error: {file_list} not found.")
        return pd.DataFrame()
    
    try:
        df_extracted_forms = pd.read_excel(files_extracted)
        df_txt_file_list = pd.read_excel(file_list)

        extracted_forms_filename_col = 'File Name'  
        form_number_col = 'Form Number'              
        form_title_col = 'Form Title'               
        txt_file_list_filename_col = 'File Name'     

        missing_cols_extracted = []
        for col in [extracted_forms_filename_col, form_number_col, form_title_col]:
            if col not in df_extracted_forms.columns:
                missing_cols_extracted.append(col)
        if missing_cols_extracted:
            print(f"Error: Missing expected columns in 'Current_FORMLIST.xlsx': {', '.join(missing_cols_extracted)}")
            return pd.DataFrame()

        if txt_file_list_filename_col not in df_txt_file_list.columns:
            print(f"Error: Missing expected column in 'AddedForms.xlsx': {txt_file_list_filename_col}")
            return pd.DataFrame()

        df_extracted_forms['Stripped File Name'] = df_extracted_forms[extracted_forms_filename_col].apply(strip_extension)
        df_txt_file_list['Stripped File Name'] = df_txt_file_list[txt_file_list_filename_col].apply(strip_extension)

        filenames_extracted = set(df_extracted_forms['Stripped File Name'])
        filenames_txt_list = set(df_txt_file_list['Stripped File Name'])
        matching_filenames = list(filenames_extracted.intersection(filenames_txt_list))
        if not matching_filenames:
            print("No matching filenames found between the two files.")
            return pd.DataFrame()

        matched_data = df_extracted_forms[
            df_extracted_forms['Stripped File Name'].isin(matching_filenames)
        ]

        result_df = matched_data[[form_number_col, form_title_col]].copy()
        result_df.rename(columns={
            form_number_col: 'Form Number',
            form_title_col: 'Form Title'
        }, inplace=True)

        output_path = os.path.join(report_dir, 'matching_AddedForms.xlsx')
        result_df.to_excel(output_path, index=False)
        print(f"\n✅ Matching data saved to: {output_path}")        
        return result_df

    except FileNotFoundError as e:
        print(f"Error: File not found. {e}")
        return pd.DataFrame()
    except pd.errors.EmptyDataError:
        print("Error: One of the Excel files is empty.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return pd.DataFrame()
    
def extract_form_infodeleted(self):
    folder_entry = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)   
    files_extracted = os.path.join(report_dir, "Previous_FORMLIST.xlsx")
    file_list = os.path.join(report_dir, "DeletedForms.xlsx")
    if not os.path.exists(files_extracted):
        print(f"Error: {files_extracted} not found.")
        return pd.DataFrame()
    if not os.path.exists(file_list):
        print(f"Error: {file_list} not found.")
        return pd.DataFrame()
    
    try:
        df_extracted_forms = pd.read_excel(files_extracted)
        df_txt_file_list = pd.read_excel(file_list)

        extracted_forms_filename_col = 'File Name'  
        form_number_col = 'Form Number'              
        form_title_col = 'Form Title'               
        txt_file_list_filename_col = 'File Name'     

        missing_cols_extracted = []
        for col in [extracted_forms_filename_col, form_number_col, form_title_col]:
            if col not in df_extracted_forms.columns:
                missing_cols_extracted.append(col)
        if missing_cols_extracted:
            print(f"Error: Missing expected columns in 'Previous_FORMLIST.xlsx': {', '.join(missing_cols_extracted)}")
            return pd.DataFrame()

        if txt_file_list_filename_col not in df_txt_file_list.columns:
            print(f"Error: Missing expected column in 'DeletedForms.xlsx': {txt_file_list_filename_col}")
            return pd.DataFrame()

        df_extracted_forms['Stripped File Name'] = df_extracted_forms[extracted_forms_filename_col].apply(strip_extension)
        df_txt_file_list['Stripped File Name'] = df_txt_file_list[txt_file_list_filename_col].apply(strip_extension)

        filenames_extracted = set(df_extracted_forms['Stripped File Name'])
        filenames_txt_list = set(df_txt_file_list['Stripped File Name'])
        matching_filenames = list(filenames_extracted.intersection(filenames_txt_list))
        if not matching_filenames:
            print("No matching filenames found between the two files.")
            return pd.DataFrame()

        matched_data = df_extracted_forms[
            df_extracted_forms['Stripped File Name'].isin(matching_filenames)
        ]

        result_df = matched_data[[form_number_col, form_title_col]].copy()
        result_df.rename(columns={
            form_number_col: 'Form Number',
            form_title_col: 'Form Title'
        }, inplace=True)

        output_path = os.path.join(report_dir, 'matching_DeletedForms.xlsx')
        result_df.to_excel(output_path, index=False)
        print(f"\n✅ Matching data saved to: {output_path}")        
        return result_df

    except FileNotFoundError as e:
        print(f"Error: File not found. {e}")
        return pd.DataFrame()
    except pd.errors.EmptyDataError:
        print("Error: One of the Excel files is empty.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return pd.DataFrame()    

def extract_form_inforevised(self):
    folder_entry = self.entry1.get()  
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)   
    
    files_extracted = os.path.join(report_dir, "Current_FORMLIST.xlsx")
    file_list = os.path.join(report_dir, "RevisedForms.xlsx")
    
    if not os.path.exists(files_extracted):
        print(f"Error: {files_extracted} not found.")
        return pd.DataFrame()
    if not os.path.exists(file_list):
        print(f"Error: {file_list} not found.")
        return pd.DataFrame()
    
    try:
        df_extracted_forms = pd.read_excel(files_extracted)
        df_txt_file_list = pd.read_excel(file_list)

        extracted_forms_filename_col = 'File Name'  
        form_number_col = 'Form Number'              
        form_title_col = 'Form Title'               
        txt_file_list_filename_col = 'File Name'     

        missing_cols_extracted = []
        for col in [extracted_forms_filename_col, form_number_col, form_title_col]:
            if col not in df_extracted_forms.columns:
                missing_cols_extracted.append(col)
        if missing_cols_extracted:
            print(f"Error: Missing expected columns in 'Current_FORMLIST.xlsx': {', '.join(missing_cols_extracted)}")
            return pd.DataFrame()

        if txt_file_list_filename_col not in df_txt_file_list.columns:
            print(f"Error: Missing expected column in 'RevisedForms.xlsx': {txt_file_list_filename_col}")
            return pd.DataFrame()

        df_extracted_forms['Stripped File Name'] = df_extracted_forms[extracted_forms_filename_col].apply(strip_extension)
        df_txt_file_list['Stripped File Name'] = df_txt_file_list[txt_file_list_filename_col].apply(strip_extension)

        filenames_extracted = set(df_extracted_forms['Stripped File Name'])
        filenames_txt_list = set(df_txt_file_list['Stripped File Name'])
        matching_filenames = list(filenames_extracted.intersection(filenames_txt_list))
        if not matching_filenames:
            print("No matching filenames found between the two files.")
            return pd.DataFrame()

        matched_data = df_extracted_forms[
            df_extracted_forms['Stripped File Name'].isin(matching_filenames)
        ]

        result_df = matched_data[[form_number_col, form_title_col]].copy()
        result_df.rename(columns={
            form_number_col: 'Form Number',
            form_title_col: 'Form Title'
        }, inplace=True)

        output_path = os.path.join(report_dir, 'matching_RevisedForms.xlsx')
        result_df.to_excel(output_path, index=False)
        print(f"\n✅ Matching data saved to: {output_path}")        
        return result_df

    except FileNotFoundError as e:
        print(f"Error: File not found. {e}")
        return pd.DataFrame()
    except pd.errors.EmptyDataError:
        print("Error: One of the Excel files is empty.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return pd.DataFrame()    
