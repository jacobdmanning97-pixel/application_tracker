import os
import re
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import shutil

class Applications:
    def __init__(self, root):
        self.root = root
        self.root.title("Application Tracker")
        self.root.state('zoomed')
        self.root.configure(bg='#f0f0f0')

        self.statuses = ['To Apply', 'Applied', 'Followed Up', 'Interview Scheduled', 'Waiting', 'Ghosted', 'Rejected']
        self.application_progression = {progress: i for i, progress in enumerate(self.statuses)}

        self.resumes = ["Main", "Data"]
        self.cover_letters = ["Main", "Data"]

        self.applications = []

        self.setup_gui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_gui(self):
        # Create main frames
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left frame for input and controls
        left_frame = ttk.LabelFrame(main_frame, text="Input", padding="10")
        left_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W), padx=(0, 10))
        
        # Right frame for display
        right_frame = ttk.LabelFrame(main_frame, text="Applications", padding="10")
        right_frame.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)

        self.folder_var = tk.StringVar(value="")

        ttk.Label(left_frame, text="Job Title:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.job_title_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.job_title_var, width=20).grid(row=0, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Company Name:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.company_name_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.company_name_var, width=20).grid(row=1, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Location:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.location_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.location_var, width=20).grid(row=2, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Last Status:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.last_status_var = tk.StringVar()
        last_status_combo = ttk.Combobox(left_frame, textvariable=self.last_status_var, 
                                     values=self.statuses, width=18)
        last_status_combo.grid(row=3, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Last Updated:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.last_updated_var = tk.StringVar(value=datetime.now().strftime("%m/%d/%Y"))
        ttk.Entry(left_frame, textvariable=self.last_updated_var, width=20).grid(row=4, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Link:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.link_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.link_var, width=20).grid(row=5, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Salary Range:").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.salary_range_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.salary_range_var, width=20).grid(row=6, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Contact Info:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.contact_info_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.contact_info_var, width=20).grid(row=7, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Notes:").grid(row=8, column=0, sticky=tk.W, pady=2)
        self.notes_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.notes_var, width=20).grid(row=8, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Resume:").grid(row=9, column=0, sticky=tk.W, pady=2)
        self.resume_var = tk.StringVar()
        resume_combo = ttk.Combobox(left_frame, textvariable=self.resume_var, 
                                     values=self.resumes, width=18)
        resume_combo.grid(row=9, column=1, pady=2, padx=(5, 0))

        ttk.Label(left_frame, text="Cover Letter:").grid(row=10, column=0, sticky=tk.W, pady=2)
        self.cover_letter_var = tk.StringVar()
        cover_letter_combo = ttk.Combobox(left_frame, textvariable=self.cover_letter_var, 
                                     values=self.cover_letters, width=18)
        cover_letter_combo.grid(row=10, column=1, pady=2, padx=(5, 0))
        
        ttk.Button(left_frame, text="Add Application", command=self.add_application, width=20).grid(row=11, column=0, columnspan=2, pady=10)
        
        ttk.Separator(left_frame, orient='horizontal').grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(left_frame, text="Update Last Status:").grid(row=13, column=0, sticky=tk.W, pady=2)
        self.update_last_status_var = tk.StringVar()
        update_last_status_combo = ttk.Combobox(left_frame, textvariable=self.update_last_status_var, 
                                     values=self.statuses, width=18)
        update_last_status_combo.grid(row=13, column=1, pady=2, padx=(5, 0))

        ttk.Button(left_frame, text="Update Selected Status", command=self.update_application, width=20).grid(row=14, column=0, columnspan=2, pady=10)

        # Delete button
        ttk.Separator(left_frame, orient='horizontal').grid(row=15, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Button(left_frame, text="Delete Selected", command=self.delete_application, width=20).grid(row=16, column=0, columnspan=2, pady=10)

        ttk.Separator(left_frame, orient='horizontal').grid(row=17, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(left_frame, text="Sort By:").grid(row=18, column=0, sticky=tk.W, pady=2)
        self.sort_var = tk.StringVar(value="Last Updated")

        last_updated_rb = ttk.Radiobutton(left_frame, text="Last Updated", variable=self.sort_var, value="Last Updated", command=self.update_display)
        last_updated_rb.grid(row=18, column=1, pady=2, sticky=tk.W)

        last_status_rb = ttk.Radiobutton(left_frame, text="Last Status", variable=self.sort_var, value="Last Status", command=self.update_display)
        last_status_rb.grid(row=19, column=1, pady=2, sticky=tk.W)

        ttk.Separator(left_frame, orient='horizontal').grid(row=20, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(left_frame, text="Upload Method:").grid(row=21, column=0, sticky=tk.W, pady=2)
        self.upload_var = tk.StringVar()

        previously_rb = ttk.Radiobutton(left_frame, text="Previously Uploaded (CSV)", variable=self.upload_var, value="Previously Uploaded", command=self.upload_csv)
        previously_rb.grid(row=21, column=1, pady=2, sticky=tk.W)

        individual_upload_rb = ttk.Radiobutton(left_frame, text="Previously Uploaded (Individual)", variable=self.upload_var, value="Individual Upload", command=self.upload_individual)
        individual_upload_rb.grid(row=22, column=1, pady=2, sticky=tk.W)   

        mass_upload_rb = ttk.Radiobutton(left_frame, text="New CSV Upload", variable=self.upload_var, value="New Mass Upload", command=self.upload_csv)
        mass_upload_rb.grid(row=23, column=1, pady=2, sticky=tk.W)

        ttk.Separator(left_frame, orient='horizontal').grid(row=24, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        columns = ("Job Title", "Company Name", "Location", "Last Status", "Last Updated", "Link", "Salary Range", "Contact Info", "Notes", "Resume", "Cover Letter")
        self.tree = ttk.Treeview(right_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add scrollbar to treeview
        scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Summary section
        summary_frame = ttk.Frame(right_frame)
        summary_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)        

    def upload_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        self.upload_var.set('')

        if not file_path:
            return
        
        try:
            # Read CSV file
            df = pd.read_csv(file_path)
            
            # Check required columns
            required_columns = ["Job Title", "Company Name", "Location", "Last Status", "Last Updated", "Link", "Salary Range", "Contact Info", "Notes", "Resume", "Cover Letter"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "CSV file must contain all columns")
                return
            
            # Add transactions from CSV
            for _, row in df.iterrows():
                app = {
                    "Job Title": row["Job Title"],
                    "Company Name": row["Company Name"],
                    "Location": row["Location"],
                    "Last Status": row["Last Status"],
                    "Last Updated": row["Last Updated"],
                    "Link": row["Link"],
                    "Salary Range": row["Salary Range"],
                    "Contact Info": row["Contact Info"],
                    "Notes": row["Notes"],
                    "Resume": row["Resume"],
                    "Cover Letter": row["Cover Letter"]
                }
                self.applications.append(app)
            
            self.update_display()
            if self.upload_var.get() == "Previously Uploaded":
                messagebox.showinfo("Success", f"Successfully imported {len(df)} applications")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV file: {str(e)}")

        if self.upload_var.get() == "New Mass Upload":
            try:
                for app in self.applications:
                    self.create_application_folder(app)
                    self.create_summary_files(app)
                messagebox.showinfo("Folder Info", "Folders were created and files moved.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create folders and move files: {str(e)}")

    def upload_individual(self):
        file_path = filedialog.askopenfilename(
            title="Select JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        self.upload_var.set('')

        if not file_path:
            return
        
        with open(file_path, "r") as f:
            data = json.load(f)

        try:
            resume = re.search(r'_([^_.]+)\.pdf$', data["Documents"]["Resume"])
            cover_letter = re.search(r'_([^_.]+)\.pdf$', data["Documents"]["Cover Letter"])

            app = {
                "Job Title": data["Application Info"]["Job Title"],
                "Company Name": data["Application Info"]["Company Name"],
                "Location": data["Application Info"]["Location"],
                "Last Status": data["Last Status"],
                "Last Updated": data["Timing"]["Last Updated"],
                "Link": data["Contact Info"]["Link"],
                "Salary Range": data["Application Info"]["Salary Range"],
                "Contact Info": data["Contact Info"]["Contact Info"],
                "Notes": data["Notes"],
                "Resume": resume.group(1),
                "Cover Letter": cover_letter.group(1)
            }

            self.applications.append(app)
            self.update_display()
            messagebox.showinfo("Success", f"Successfully imported {data["Application Info"]["Company Name"]}'s application")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read JSON file: {str(e)}")

    def add_application(self):
        job_title = self.job_title_var.get().replace("\n","")
        company_name = self.company_name_var.get().replace("\n","")
        location = self.location_var.get().replace("\n","")
        status = self.last_status_var.get().replace("\n","")
        last_updated = self.last_updated_var.get().replace("\n","")
        link = self.link_var.get().replace("\n","")
        salary_range = self.salary_range_var.get().replace("\n","")
        contact_info = self.contact_info_var.get().replace("\n","")
        notes = self.notes_var.get().replace("\n","")
        resume = self.resume_var.get().replace("\n","")
        cover_letter = self.cover_letter_var.get().replace("\n","")

        if not all([job_title, company_name, location, status, last_updated, link, salary_range, contact_info, notes, resume, cover_letter]):
            messagebox.showerror("Error", "Please fill all fields with valid values")
            return

        app = {
            "Job Title": job_title,
            "Company Name": company_name,
            "Location": location,
            "Last Status": status,
            "Last Updated": last_updated,
            "Link": link,
            "Salary Range": salary_range,
            "Contact Info": contact_info,
            "Notes": notes,
            "Resume": resume,
            "Cover Letter": cover_letter
        }

        self.applications.append(app)
        self.update_display()
        try:
            self.create_application_folder(app)
            self.create_summary_files(app)
            messagebox.showinfo("Folder Info", "Folders and summaries were created and files moved.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create folders, summaries, and move files: {str(e)}")

        # Clear input fields
        self.job_title_var.set('')
        self.company_name_var.set('')
        self.location_var.set('')
        self.last_status_var.set('')
        self.last_updated_var.set(datetime.now().strftime("%m/%d/%Y"))
        self.link_var.set('')
        self.salary_range_var.set('')
        self.contact_info_var.set('')
        self.notes_var.set('')
        self.resume_var.set('')
        self.cover_letter_var.set('')

    def update_application(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an application to change update")
            return
        
        update_last_status = self.update_last_status_var.get()

        if not all([update_last_status]):
                messagebox.showerror("Error", "Please fill all fields with valid values")
                return

        index = int(selected_item[0].lstrip('I')) - 1
        self.applications[index]["Last Status"] = update_last_status
        self.applications[index]["Last Updated"] = datetime.now().strftime("%m/%d/%Y")
        self.update_display()
        self.update_last_status_var.set('')

        self.create_summary_files(self.applications[index])

    def delete_application(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a application to delete")
            return
        
        # Confirm deletion
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this application?"):
            index = int(selected_item[0].lstrip('I')) - 1
            if 0 <= index < len(self.applications):
                del self.applications[index]
                self.update_display()

    def update_display(self):
        sort = self.sort_var.get()
        if sort == "Last Updated":
            self.applications = sorted([item for item in self.applications], key=lambda x: x[sort])
        elif sort == "Last Status":
            self.applications = sorted([item for item in self.applications], key=lambda x: self.application_progression.get(x.get("Last Status")))

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for i, app in enumerate(self.applications):
            self.tree.insert("", "end", iid=f"I{i+1:03d}", values=(
                app["Job Title"],
                app["Company Name"],
                app["Location"],
                app["Last Status"],
                pd.to_datetime(app["Last Updated"]).date().strftime("%m/%d/%Y"),
                app["Link"],
                app["Salary Range"],
                app["Contact Info"],
                app["Notes"],
                app["Resume"],
                app["Cover Letter"]
            ))

    def create_application_folder(self, app):
        company_name = app["Company Name"]
        resume_file = f"Jacob_Manning_Resume_{app["Resume"]}.pdf"
        cover_file = f"Jacob_Manning_Cover_Letter_{app["Cover Letter"]}.pdf"

        master_folder = "C:\\Users\\jacob\\OneDrive\\Desktop\\Applications\\Main Folder"
        # Create full path for new folder
        company_name_path = os.path.join("C:\\Users\\jacob\\OneDrive\\Desktop\\Applications", company_name)

        if os.path.exists(company_name_path):
            company_name_path = os.path.join(company_name_path, app["Job Title"])
            self.folder_var.set("Company Applied Already")
        
        # Check if master folder exists
        if not os.path.exists(master_folder):
            messagebox.showerror("Error", f"Error: Master folder '{master_folder}' not found.")
        
        # Create new folder
        os.makedirs(company_name_path)
        
        # Copy resume and cover letter
        resume_source = os.path.join(master_folder, resume_file)
        cover_source = os.path.join(master_folder, cover_file)
        shutil.copy2(resume_source, company_name_path)
        shutil.copy2(cover_source, company_name_path)

    def create_summary_files(self, app):
        metadata = {
            "Application Info": {
                "Company Name": app["Company Name"],
                "Job Title": app["Job Title"],
                "Location": app["Location"],
                "Salary Range": app["Salary Range"]
            },
            "Last Status": app["Last Status"],
            "Documents": {
                "Resume": f"Jacob_Manning_Resume_{app["Resume"]}.pdf",
                "Cover Letter": f"Jacob_Manning_Cover_Letter_{app["Cover Letter"]}.pdf"
            },
            "Contact Info": {
                "Contact Info": app["Contact Info"],
                "Link": app["Link"]
            },
            "Notes": app["Notes"],
            "Timing": {
                "Last Updated": app["Last Updated"],        
                "Metadata Created": datetime.now().strftime("%m/%d/%Y")
            }
        }
        
        if self.folder_var.get() == "":
            base_filename = f"C:\\Users\\jacob\\OneDrive\\Desktop\\Applications\\{app["Company Name"]}"
        elif self.folder_var.get() == "Company Applied Already":
            base_filename = f"C:\\Users\\jacob\\OneDrive\\Desktop\\Applications\\{app["Company Name"]}\\{app["Job Title"]}"

        self.folder_var.set("")

        #Save as json
        metadata_file = f"{base_filename}\\ReadMe_Json.json"
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)
        
        # Also create a simple text version for quick viewing
        txt_metadata_file = f"{base_filename}\\ReadMe_Text.txt"
        with open(txt_metadata_file, 'w', encoding='utf-8') as f:
            f.write("=" * 50 + "\n")
            f.write("JOB APPLICATION DOCUMENTS\n")
            f.write("=" * 50 + "\n\n")
            
            f.write("APPLICATION INFORMATION:\n")
            f.write("-" * 30 + "\n")
            f.write(f"Company: {metadata["Application Info"]["Company Name"]}\n")
            f.write(f"Job Title: {metadata["Application Info"]["Job Title"]}\n")
            f.write(f"Location: {metadata["Application Info"]["Location"]}\n")
            f.write(f"Salary Range: {metadata["Application Info"]["Salary Range"]}\n\n")

            f.write("LAST STATUS:\n")
            f.write("-" * 30 + "\n")
            f.write(f"Last Status: {metadata["Last Status"]}\n")
            f.write(f"Last Updated: {metadata["Timing"]["Last Updated"]}\n\n")

            f.write("Contact Info:\n")
            f.write("-" * 30 + "\n")
            f.write(f"Contact Info: {metadata["Contact Info"]["Contact Info"]}\n")
            f.write(f"Link: {metadata["Contact Info"]["Link"]}\n\n")
            
            f.write("DOCUMENTS:\n")
            f.write("-" * 30 + "\n")
            f.write(f"Resume: {metadata["Documents"]["Resume"]}\n")
            f.write(f"Cover Letter: {metadata["Documents"]["Cover Letter"]}\n\n")

            f.write("NOTES:\n")
            f.write("-" * 30 + "\n")
            f.write(f"{metadata["Notes"]}\n\n")
            
            f.write("=" * 50 + "\n")
            f.write(f"Metadata created: {metadata["Timing"]["Metadata Created"]}\n")
            f.write("=" * 50 + "\n")

    def on_closing(self):
        df = pd.DataFrame(self.applications)
        df.to_csv("C:\\Users\\jacob\\OneDrive\\Desktop\\Applications\\Main Folder\\apps.csv", index = False)
        # Perform any cleanup needed
        self.root.destroy()
        self.root.quit()  # This helps terminate mainloop properly

root = tk.Tk()
app = Applications(root)
root.mainloop()