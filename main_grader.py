import os
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, messagebox
from tkinter import ttk
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, PatternFill
from openpyxl.utils.exceptions import InvalidFileException

# Import the grading algorithms from grading_algorithms.py
from grading_algorithms import *

def get_grading_function(challenge_number):
    grading_functions = {
        "Project 1: Cafe Bloom": grade_project_1,
        "Project 2: Sports Event Management": grade_project_2,
        "1.1: Import data into workbooks": grade_challenge_1_1,
        "2: Navigate within workbooks": grade_challenge_2,
        "3.1: Format worksheets and workbooks": grade_challenge_3_1
    }
    return grading_functions.get(challenge_number), grading_functions

# Function to process student submissions
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
    
    # Create named styles
    outstanding_style = NamedStyle(name='Outstanding')
    outstanding_style.fill = PatternFill(start_color='C099E8', end_color='C099E8', fill_type='solid')  # Purple
    
    good_style = NamedStyle(name='Good')
    good_style.fill = PatternFill(start_color='41DF45', end_color='41DF45', fill_type='solid')  # Green
    
    neutral_style = NamedStyle(name='Neutral')
    neutral_style.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')  #  Yellow
    
    bad_style = NamedStyle(name='Bad')
    bad_style.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')  # Red
    
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

# Function to set up the GUI
def setup_gui():
    root = Tk()
    root.title("Excel Grader")
    root.geometry("300x400")

    # Get grading function labels for the combobox
    _, grading_functions = get_grading_function(None)
    challenge_labels = list(grading_functions.keys())

    # Variables to store the full paths
    folder_full_path = None
    output_full_path = None

    def select_folder():
        nonlocal folder_full_path
        folder = filedialog.askdirectory()
        if folder:
            folder_full_path = folder
            parent_folder = os.path.basename(os.path.normpath(folder))
            folder_label.config(text=parent_folder)

    def select_output():
        nonlocal output_full_path
        output = filedialog.askdirectory()
        if output:
            output_full_path = output
            parent_folder = os.path.basename(os.path.normpath(output))
            output_label.config(text=parent_folder)

    def start_grading():
        nonlocal folder_full_path, output_full_path

        folder_path = folder_full_path
        challenge_number = challenge_combobox.get()
        output_path = output_full_path

        if not folder_path or not challenge_number or not output_path:
            messagebox.showwarning("Input Error", "Please select all required inputs.")
            return

        process_submissions(folder_path, challenge_number, output_path)

    # Create GUI components
    Label(root, text="Select Student Submissions Folder:").pack(pady=5)
    folder_label = Label(root, text="", wraplength=350)
    folder_label.pack(pady=5)
    Button(root, text="Browse", command=select_folder).pack(pady=5)

    Label(root, text="Select Challenge to Grade:").pack(pady=5)
    challenge_combobox = ttk.Combobox(root, values=challenge_labels)
    challenge_combobox.pack(pady=5)

    Label(root, text="Select Output Location:").pack(pady=5)
    output_label = Label(root, text="", wraplength=350)
    output_label.pack(pady=5)
    Button(root, text="Browse", command=select_output).pack(pady=5)

    Button(root, text="Start Grading", command=start_grading).pack(pady=20)

    root.mainloop()

# Entry point for the program
if __name__ == "__main__":
    setup_gui()