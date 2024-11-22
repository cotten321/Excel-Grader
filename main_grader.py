import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, PatternFill
from openpyxl.utils.exceptions import InvalidFileException

# Import the grading algorithms from grading_algorithms.py
from grading_algorithms import *

# Set the appearance mode and color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def get_grading_function(challenge_number):
    grading_functions = {
        "Project 1: Cafe Bloom": grade_project_1,
        "1.1: Import data into workbooks": grade_challenge_1_1,
        "2: Navigate within workbooks": grade_challenge_2,
        "3.1: Format worksheets and workbooks": grade_challenge_3_1
    }
    return grading_functions.get(challenge_number), grading_functions

def process_submissions(folder_path, challenge_number, output_path):
    grading_function, _ = get_grading_function(challenge_number)
    if not grading_function:
        messagebox.showerror("Grading Error", f"No grading function available for challenge or project {challenge_number}.")
        return

    grades = []

    # Iterate through student folders
    for student_folder in os.listdir(folder_path):
        student_folder_path = os.path.join(folder_path, student_folder)
        
        if os.path.isdir(student_folder_path):
            # Look for an Excel file inside the student's folder
            for file in os.listdir(student_folder_path):
                if file.endswith(".xlsx"):
                    student_file_path = os.path.join(student_folder_path, file)
                    print(f"Grading {student_file_path}")

                    # Try to load the workbook and process
                    try:
                        # Use the selected grading function
                        score, total_points, feedback = grading_function(student_file_path)
                        percentage = round((score / total_points) * 100, 2) if total_points > 0 else 0

                        # Add the result to the grades list
                        grades.append({
                            "Student": student_folder,
                            "Score": score,
                            "Total Points": total_points,
                            "Percentage": percentage,
                            "": "",  # Empty column for easier viewing
                            "Feedback": "; ".join(feedback)  # Join feedback items as a single string.
                        })

                    except InvalidFileException as e:
                        feedback = f"Error loading file: {e}"
                        grades.append({
                            "Student": student_folder,
                            "Score": 0,
                            "Total Points": 100,
                            "Percentage": 0,
                            "": "",
                            "Feedback": feedback
                        })
                    except Exception as e:
                        grades.append({
                            "Student": student_folder,
                            "Score": 0,
                            "Total Points": 100,
                            "Percentage": 0,
                            "": "",
                            "Feedback": f"Error: {str(e)}"
                        })

    # Convert grades list to a DataFrame and export to Excel
    df = pd.DataFrame(grades)
    
    output_file = os.path.join(output_path, "grades_report.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Grading Report"
    
    # Create named styles (keep your existing style creation code)
    outstanding_style = NamedStyle(name='Outstanding')
    outstanding_style.fill = PatternFill(start_color='C099E8', end_color='C099E8', fill_type='solid')
    
    good_style = NamedStyle(name='Good')
    good_style.fill = PatternFill(start_color='41DF45', end_color='41DF45', fill_type='solid')
    
    neutral_style = NamedStyle(name='Neutral')
    neutral_style.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
    
    bad_style = NamedStyle(name='Bad')
    bad_style.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    
    # Add styles to workbook
    if 'Outstanding' not in wb.named_styles:
        wb.add_named_style(outstanding_style)
    if 'Good' not in wb.named_styles:
        wb.add_named_style(good_style)
    if 'Neutral' not in wb.named_styles:
        wb.add_named_style(neutral_style)
    if 'Bad' not in wb.named_styles:
        wb.add_named_style(bad_style)
    
    # Append the header
    ws.append(["Student", "Score", "Total Points", "Percentage", "", "Feedback"])
    
    # Add student data and apply cell styles
    for index, row in df.iterrows():
        ws.append(row.tolist())
        
        # Apply styles based on the score
        cell = ws[f"A{index + 2}"]
        if row["Percentage"] > 100:
            cell.style = "Outstanding"
        elif row["Percentage"] > 85:
            cell.style = "Good"
        elif 65 <= row["Percentage"] <= 85:
            cell.style = "Neutral"
        else:
            cell.style = "Bad"
            
    # Save the report
    wb.save(output_file)        
    messagebox.showinfo("Success", f"Grading complete! Report saved to: {output_file}")

class ExcelGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Excel Grader")
        self.geometry("500x650")
        self.configure(fg_color="#F0F0F0")  # Light gray background

        # Main container with soft rounded corners
        self.main_frame = ctk.CTkFrame(
            self, 
            corner_radius=20, 
            fg_color="white", 
            bg_color="#F0F0F0"
        )
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Title with modern typography
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Excel Grader", 
            font=("San Francisco", 32, "bold"),
            text_color="#333333"
        )
        self.title_label.pack(pady=(30, 20))

        # Submission Folder Section
        self.create_folder_section(
            "Student Submissions", 
            self.select_submissions_folder
        )

        # Challenge Selection Section
        self.challenge_label = ctk.CTkLabel(
            self.main_frame, 
            text="Select Challenge", 
            font=("San Francisco", 16),
            text_color="#666666"
        )
        self.challenge_label.pack(anchor="w", padx=40, pady=(20, 5))

        self.challenges = [
            "Project 1: Cafe Bloom", 
            "1.1: Import data into workbooks", 
            "2: Navigate within workbooks", 
            "3.1: Format worksheets and workbooks"
        ]

        self.challenge_combobox = ctk.CTkComboBox(
            self.main_frame, 
            values=self.challenges,
            width=400,
            height=40,
            border_width=1,
            border_color="#CCCCCC",
            dropdown_hover_color="#E0E0E0",
            button_hover_color="#E0E0E0",
            font=("San Francisco", 14)
        )
        self.challenge_combobox.pack(pady=10)

        # Output Folder Section
        self.create_folder_section(
            "Output Location", 
            self.select_output_folder
        )

        # Start Grading Button with modern styling
        self.start_button = ctk.CTkButton(
            self.main_frame, 
            text="Start Grading", 
            command=self.start_grading,
            width=400,
            height=50,
            corner_radius=25,
            font=("San Francisco", 16, "bold"),
            fg_color="#007AFF",  # Apple's blue
            hover_color="#0056b3"
        )
        self.start_button.pack(pady=(30, 20))

        # State variables
        self.submissions_folder = None
        self.output_folder = None

    def create_folder_section(self, label_text, browse_command):
        # Label
        label = ctk.CTkLabel(
            self.main_frame, 
            text=label_text, 
            font=("San Francisco", 16),
            text_color="#666666"
        )
        label.pack(anchor="w", padx=40, pady=(20, 5))

        # Container for entry and button
        container = ctk.CTkFrame(
            self.main_frame, 
            fg_color="transparent"
        )
        container.pack(pady=10)

        # Entry field
        entry = ctk.CTkEntry(
            container, 
            width=330,
            height=40,
            placeholder_text=f"Select {label_text.lower()}",
            border_width=1,
            border_color="#CCCCCC",
            font=("San Francisco", 14)
        )
        entry.pack(side="left", padx=(0, 10))

        # Browse button
        browse_btn = ctk.CTkButton(
            container, 
            text="Browse", 
            command=lambda: self.browse_folder(entry, browse_command),
            width=60,
            height=40,
            corner_radius=10,
            fg_color="#F2F2F7",  # Very light gray
            text_color="#007AFF",  # Apple's blue
            hover_color="#E0E0E5"
        )
        browse_btn.pack(side="right")

        # Store references for later use
        if label_text == "Student Submissions":
            self.submissions_entry = entry
        else:
            self.output_entry = entry

    def browse_folder(self, entry_widget, selection_method):
        selection_method()
        entry_widget.configure(state="normal")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, self.submissions_folder if "submissions" in entry_widget.cget("placeholder_text").lower() else self.output_folder)
        entry_widget.configure(state="disabled")

    def select_submissions_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.submissions_folder = folder
            self.submissions_entry.configure(state="normal")
            self.submissions_entry.delete(0, "end")
            self.submissions_entry.insert(0, folder)
            self.submissions_entry.configure(state="readonly")

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.output_entry.configure(state="normal")
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.output_entry.configure(state="readonly")

    def start_grading(self):
        if not self.submissions_folder or not self.output_folder or not self.challenge_combobox.get():
            messagebox.showwarning("Input Error", "Please select all required inputs.")
            return

        process_submissions(
            self.submissions_folder, 
            self.challenge_combobox.get(), 
            self.output_folder
        )

def main():
    app = ExcelGraderApp()
    app.mainloop()

if __name__ == "__main__":
    main()