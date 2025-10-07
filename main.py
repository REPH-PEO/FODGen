import tkinter.ttk as ttk
import tkinter as tk
from tkinter import messagebox, PhotoImage, Label, filedialog
from PIL import Image, ImageTk
import pandas as pd
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import time

from tkinter import filedialog
import extract_zipfile
import revisionsforms
import listtxtnew
import rtf2txt
import FormlistTxt2xls
import matchrevision2xlsx
import updatertfnew
import CollectedForms
import ExtractTitle
import releasenotes
import userguideupdate
import copyguide
import move_rtf
import zip_rtf
import copy_ft_output
import deleteMODIFIED
import copyFinalguide
import clean_up
import browse


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FOD User Guide Generator®")
        self.geometry("500x480")
        self.resizable(False, False)  
        self.configure(bg="#2a0b05", highlightbackground="black", borderwidth=1, relief="solid", highlightthickness=1, highlightcolor="white")
        # self.iconbitmap(r'C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\logo\logo.ico')
        self.iconbitmap(r'\\fabwebd5.net\neptune\DataConversion\Tools\FOD\reference\logo.ico')        

        # Create a menu bar
        main_menu = tk.Menu(self)
        self.config(menu=main_menu)
        main_menu.add_cascade(label=" " * 146, state="disabled")  # Spacer
        def show_about():
            messagebox.showinfo(
            "About",
            "This tool will generate FOD User Guides for New/Deleted/Revised Forms.\nVersion: 1.0\nDeveloped by: BRi\nFor issues and concerns, please contact: brian.labrador@reedelsevier.com"
            )
        def show_guide():
            messagebox.showinfo(
            "Guide",
            "1. Enter 5-digit Pub Number.\n2. Enter Release Number.\n3. Enter/Select Revenue Month and Year.\n4. Browse and select last release zip file.\n5. Browse and select userguide reference.\n6. Enter path for [FT_OUTPUT_PATH] from job summary. \n7. Click Generate to proceed.\n8.Click ZipFile to Zip.\n9. Click Exit to close the application."
            )

        help_menu = tk.Menu(main_menu, tearoff=0)
        help_menu.add_command(label="About", command=show_about)
        help_menu.add_command(label="Guide", command=show_guide)
        main_menu.add_cascade(label="Help", menu=help_menu)
        # Main frame   
        self.main_frame = tk.Frame(self, bg="#223556", width=150 , height=50)
        self.main_frame.pack(side="right", expand=True, fill="both", padx=10, pady=10)
        self.main_frame1 = tk.Frame(self.main_frame, bg="#223556", width=300, height=50)
        self.main_frame1.pack(side="top", expand=True, fill="both", padx=5, pady=5)
        # self.image_file = r'C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\logo\FOD.png'
        self.image_file = r'\\fabwebd5.net\neptune\DataConversion\Tools\FOD\reference\FOD.png'
        self.image = Image.open(self.image_file)
        self.image = self.image.resize((80, 80))  # Resize to desired size
        self.photo = ImageTk.PhotoImage(self.image)
        self.image_label = Label(self.main_frame1, image=self.photo, bg="#223556")
        self.image_label.pack(side="left", padx=10)
        self.main_label = Label(self.main_frame1, text="FOD User Guide Generator", font=("Corporate", 18), bg="#223556", fg="white")
        self.main_label.pack(side="left", padx=10, pady=10)
        self.sub_frame1 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame1.pack(side="top", expand=False, fill="both", padx=5, pady=5)
        self.sub_frame2 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame2.pack(side="top", expand=False, fill="both", padx=3, pady=3)        

        self.entry_label1 = tk.Label(self.sub_frame1, text="Input Pub Information!", font=("Arial", 9), bg="#223556", fg="white")
        self.entry_label1.pack(pady=(5, 0), side="top", anchor="center")

        entry_fields_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields_frame.pack(pady=2, side="top", anchor="center", expand=False)
        entry_fields1_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields1_frame.pack(pady=2, side="top", anchor="center", expand=False)
        entry_fields2_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields2_frame.pack(pady=2, side="top", anchor="center", expand=False)        
        entry_fields3_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields3_frame.pack(pady=2, side="top", anchor="center", expand=False)          

        entry1_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry1_frame.grid(row=0, column=0, padx=10, pady=2, sticky="w")
        self.entry1_label = tk.Label(entry1_frame, text="Pub Number:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry1_label.pack(side="left", padx=(5, 0))
        self.entry1 = tk.Entry(entry1_frame, width=15, font=("Arial", 8))
        self.entry1.pack(side="left", fill="x", expand=False)

        entry2_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry2_frame.grid(row=0, column=1, padx=10, pady=2, sticky="w")
        self.entry2_label = tk.Label(entry2_frame, text="Rel Number:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry2_label.pack(side="left", padx=(5, 0))
        self.entry2 = tk.Entry(entry2_frame, width=15, font=("Arial", 8))
        self.entry2.pack(side="left", fill="x", expand=False)        

        entry3_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry3_frame.grid(row=1, column=0, padx=10, pady=2, sticky="w")
        self.entry3_label = tk.Label(entry3_frame, text="Rev Month:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry3_label.pack(side="left", padx=(5, 0))
        # Dropdown for months
        months = ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"]
        self.entry3_var = tk.StringVar()
        self.entry3_var.set(months[0])  # Default to current month
        current_month_index = datetime.now().month - 1  # 0-based index
        self.entry3 = ttk.Combobox(entry3_frame, textvariable=self.entry3_var, values=months, state="readonly", width=12, font=("Arial", 8))
        self.entry3.current(current_month_index)
        self.entry3_var.set(months[current_month_index])
        self.entry3.pack(side="right", fill="x", expand=False)

        entry4_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry4_frame.grid(row=1, column=1, padx=10, pady=2, sticky="w")
        self.entry4_label = tk.Label(entry4_frame, text="Year:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry4_label.pack(side="left", padx=(5, 0))
        self.entry4 = tk.Entry(entry4_frame, width=15, font=("Arial", 8))
        self.entry4.pack(side="left", fill="x", expand=True)
        self.entry4.insert(0, str(datetime.now().year))

        entry5_frame = tk.Frame(entry_fields2_frame, bg="#F7F7F8")
        entry5_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry5_label = tk.Label(entry5_frame, text="Previous Release ZipFile:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=20)
        self.entry5_label.pack(side="left", padx=(5, 0))
        self.entry5 = tk.Entry(entry5_frame, width=37, font=("Arial", 8))
        self.entry5.pack(side="left", padx=5)
        def browse_zip_file():
            #C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive
            initial_dir = os.path.join(
                r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary",
                self.entry1.get().strip(),"FOD"
            )
            if not os.path.isdir(initial_dir):
                initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
            file_path = filedialog.askopenfilename(
                initialdir=initial_dir,
                title="Select Zip File",
                filetypes=[("Zip Files", "*.zip"), ("All Files", "*.*")]
            )
            if file_path:
                self.entry5.delete(0, tk.END)
                self.entry5.insert(0, file_path)
        self.entry5_button = tk.Button(entry5_frame, text="Browse", font=("Arial", 8), command=browse_zip_file, border=1, relief="solid", bg="#F7F7F8", fg="black")
        self.entry5_button.pack(side="right", padx=1)
       
        entry6_frame = tk.Frame(entry_fields2_frame, bg="#F7F7F8")
        entry6_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry6_label = tk.Label(entry6_frame, text="Current User Guide:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=20)
        self.entry6_label.pack(side="left", padx=(5, 0))
        self.entry6 = tk.Entry(entry6_frame, width=37, font=("Arial", 8))
        self.entry6.pack(side="left", padx=5)
        def browse_userguide():
            #C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive
            initial_dir = os.path.join(
                r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary",
                self.entry1.get().strip(), "FOD"
            )   
            
            if not os.path.isdir(initial_dir):
                initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
            file_path = filedialog.askopenfilename(
                initialdir=initial_dir,
                title="Select User Guide",
                filetypes=[("RTF Files", "*.rtf"), ("All Files", "*.*")]
            )
            if file_path:
                self.entry6.delete(0, tk.END)
                self.entry6.insert(0, file_path)
        self.entry6_button = tk.Button(entry6_frame, text="Browse", font=("Arial", 8), command=browse_userguide, border=1, relief="solid", bg="#F7F7F8", fg="black")
        self.entry6_button.pack(side="right", padx=1)

        entry7_frame = tk.Frame(entry_fields3_frame, bg="#F7F7F8")
        entry7_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry7_label = tk.Label(entry7_frame, text="FT_OUTPUT_PATH:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry7_label.pack(side="left", padx=(5, 0))
        self.entry7 = tk.Entry(entry7_frame, width=50, font=("Arial", 8))
        self.entry7.pack(side="left", padx=5)  

        self.submit_button = tk.Button(self.sub_frame2, text="Generate", font=("Gagalin", 10), command=self.check_path, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.submit_button.pack(pady=15, padx=10, side="left", anchor="w")
        self.Revoke_button = tk.Button(self.sub_frame2, text="Zip File", command=self.check_pathrevoke, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.Revoke_button.pack(padx=15, side="left", anchor="w")
        # self.rerun_button = tk.Button(self.sub_frame2, text="Rerun", command=self.refresh_run, width=13, height=2)
        # self.rerun_button.pack(padx=10, side="left", anchor="w")
        self.exit_app_button = tk.Button(self.sub_frame2, text="Exit", command=self.exit_app, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.exit_app_button.pack(padx=15, side="left", anchor="w")

        self.sub_frame4 = tk.Frame(self.main_frame, bg="gray", width=150, height=100)
        self.sub_frame4.pack(side="top", expand=False, fill="both", padx=5, pady=5) 
        self.sub_frame3 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame3.pack(side="top", expand=True, fill="both", padx=5, pady=5)               
        # self.sub_frame2 = tk.Frame(self.main_frame, bg="#13F031")
        # self.sub_frame2.pack(side="top", expand=False, fill="both", padx=5, pady=5)
        self.percent_label = Label(self.sub_frame3, text="0%", font=("Arial", 5), bg="#223556", fg="white", height=1)
        self.percent_label.pack(pady=1, side="top", anchor="n")         
        self.progress_bar = tk.Scale(self.sub_frame4, orient="horizontal", length=500, sliderlength=5, 
                                     bg="#ECEEF3", highlightbackground="#223556" )
        
        self.progress_bar.set(0)  
        self.progress_bar.pack(side="top", anchor="w", expand=True)

        self.start_time = 0

    def exit_app(self):
        self.destroy()

    def update_timer(self, timer_label, popup):
        elapsed_time = time.time() - self.start_time
        minutes, seconds = divmod(int(elapsed_time), 60)
        timer_label.configure(text=f"Time Elapsed: {minutes:02}:{seconds:02}")
        if popup.winfo_exists():
            popup.after(1000, self.update_timer, timer_label, popup)        

    def check_path(self):
        # Check if all entry fields are not empty
        pub_number = self.entry1.get().strip()
        rel_number = self.entry2.get().strip()
        rev_month = self.entry3_var.get().strip()
        year = self.entry4.get().strip()
        prev_zip = self.entry5.get().strip()
        userguide = self.entry6.get().strip()
        ft_output = self.entry7.get().strip()


        if not all([pub_number, rel_number, rev_month, year, prev_zip, userguide, ft_output]):
            messagebox.showerror("Error", "All fields must be filled in.")
            return

        if messagebox.askyesno("Confirm", "All fields are filled in."):

            self.submit_button.configure(state="disabled", text="Processing...")
            self.Revoke_button.configure(state="disabled", text="Zip File")
            self.exit_app_button.configure(state="disabled", text="Exit")
            # --- Create and show the processing pop-up ---
            popup = tk.Toplevel(self)
            popup.title("Processing...")
            popup.transient(self)
            popup.grab_set()

            # Center the pop-up
            main_window_x = self.winfo_x()
            main_window_y = self.winfo_y()
            main_window_width = self.winfo_width()
            main_window_height = self.winfo_height()

            popup_width = 300
            popup_height = 100
            popup_x = main_window_x + (main_window_width // 2) - (popup_width // 2)
            popup_y = main_window_y + (main_window_height // 2) - (popup_height // 2)
            popup.geometry(f"{popup_width}x{popup_height}+{popup_x}+{popup_y}")
            popup.resizable(False, False)

            self.label = tk.Label(popup, text="Process is complete!")
            self.label.pack(pady=10)
            self.timer_label = tk.Label(popup, text="Time Elapsed: 00:00")
            self.timer_label.pack(pady=5)

            self.start_time = time.time()
            self.after(1000, self.update_timer, self.timer_label, popup)
            self.update_idletasks()            


            # --- Begin long-running processing ---
            try:
                move_rtf.move_rtf_files(self)
                copy_ft_output.copy_ftoutput_rtf_files(self)
                extract_zipfile.extract_zip_action(self)
                revisionsforms.comparertf(self)
                copyguide.copy_file(self)
                listtxtnew.listrevision_txt_files(self)
                rtf2txt.convert_rtf_to_txt_current(self)
                rtf2txt.convert_rtf_to_txt_prev(self)
                FormlistTxt2xls.extracttxt_to_excelcurrent(self)
                FormlistTxt2xls.extracttxt_to_excelprev(self)
                FormlistTxt2xls.extracttxt_to_exceladd(self)
                FormlistTxt2xls.extracttxt_to_exceldeleted(self)
                FormlistTxt2xls.extracttxt_to_excelrevised(self)
                matchrevision2xlsx.extract_form_infodeleted(self)
                matchrevision2xlsx.extract_form_infoadd(self)
                matchrevision2xlsx.extract_form_inforevised(self)

                # # updatertfnewtest.modify_rtf_with_word(self) #not needed

                CollectedForms.process_forms(self)
                ExtractTitle.extract_forms_title_from_rtf(self)
                ExtractTitle.extract_forms_sample_from_rtf(self)    
                releasenotes.generate_report(self)
                deleteMODIFIED.delete_modified_userguide_rtf(self)
                userguideupdate.userguide_update(self)
            
            #Include after testing so user can now update the userguide                
                # copyFinalguide.copy_finalfile(self)
                # zip_rtf.zip_rtf_files(self)
                # clean_up.clean_directories(self)

                print("Process Complete!")
                messagebox.showinfo("Complete", "Process Complete!")                
                if popup.winfo_exists():
                        popup.destroy() 
                # Simulate progress bar
                for i in range(101):
                    self.progress_bar.set(i)
                    self.percent_label.configure(text=f"{i}% Process Complete!")
                    self.update_idletasks()
                    time.sleep(0.0001)

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred:\n{e}")
            finally:
                self.submit_button.configure(state="disabled", text="Done")
                self.Revoke_button.configure(state="normal", text="Zip File")
                self.exit_app_button.configure(state="normal", text="Exit")
                self.progress_bar.set(100)
                self.percent_label.configure(text="100% Process Complete!")
        else:
            messagebox.showerror("Error", "Process cancelled by user.")

    def check_pathrevoke(self):
        # Check if all entry fields are not empty
        pub_number = self.entry1.get().strip()
        rel_number = self.entry2.get().strip()

        if not all([pub_number, rel_number]):
            messagebox.showerror("Error", "Please enter Pub Number and Rel Number.")
            return

        if messagebox.askyesno("Confirm", "Zip rtf files?"):

            self.submit_button.configure(state="disabled", text="Processing...")
            self.Revoke_button.configure(state="disabled", text="Zip File")
            self.exit_app_button.configure(state="disabled", text="Exit")
            # --- Create and show the processing pop-up ---
            popup = tk.Toplevel(self)
            popup.title("Processing...")
            popup.transient(self)
            popup.grab_set()

            # Center the pop-up
            main_window_x = self.winfo_x()
            main_window_y = self.winfo_y()
            main_window_width = self.winfo_width()
            main_window_height = self.winfo_height()

            popup_width = 300
            popup_height = 100
            popup_x = main_window_x + (main_window_width // 2) - (popup_width // 2)
            popup_y = main_window_y + (main_window_height // 2) - (popup_height // 2)
            popup.geometry(f"{popup_width}x{popup_height}+{popup_x}+{popup_y}")
            popup.resizable(False, False)

            self.label = tk.Label(popup, text="Process is complete!")
            self.label.pack(pady=10)
            self.timer_label = tk.Label(popup, text="Time Elapsed: 00:00")
            self.timer_label.pack(pady=5)

            self.start_time = time.time()
            self.after(1000, self.update_timer, self.timer_label, popup)
            self.update_idletasks()            
            try:
                copyFinalguide.copy_finalfile(self)
                zip_rtf.zip_rtf_files(self)
                clean_up.clean_directories(self)

                print("Process Complete!")
                messagebox.showinfo("Complete", "Process Complete!")                
                if popup.winfo_exists():
                        popup.destroy() 
                # Simulate progress bar
                for i in range(101):
                    self.progress_bar.set(i)
                    self.percent_label.configure(text=f"{i}% Process Complete!")
                    self.update_idletasks()
                    time.sleep(0.0001)

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred:\n{e}")
            finally:
                self.submit_button.configure(state="normal", text="Regenerate")
                self.Revoke_button.configure(state="normal", text="Zip File")
                self.exit_app_button.configure(state="normal", text="Exit")
                self.progress_bar.set(100)
                self.percent_label.configure(text="100% Process Complete!")
        else:
            messagebox.showerror("Error", "Process cancelled by user.")             

    def refresh_run(self):
        self.submit_button.configure(state="normal", text="Submit")    
        self.entry.delete(0, tk.END)  
        messagebox.showinfo("Re-run", "You can do another Clean up.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
