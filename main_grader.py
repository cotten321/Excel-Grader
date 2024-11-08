#import openpyxl
#import csv
import os
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, messagebox


# Import the grading algorithms from grading_algorithms.py
from grading_algorithms import *

# Function to determine the correct grading function based on challenge number
def get_grading_function(challenge_number):
    grading_functions = {
        "1.1": grade_challenge_1_1,
        #"1.2": grade_challenge_1_2,
        "3.1": grade_challenge_3_1,
        # Add more mappings like "1.2": grade_challenge_1_2, etc.
    }
    return grading_functions.get(challenge_number)

# Function to process student submissions
def process_submissions(folder_path, solution_path, output_path):
    grades = []

    # Extract challenge number from the solution file name
    challenge_number = os.path.basename(solution_path).split('_')[0]
    grading_function = get_grading_function(challenge_number)

    if not grading_function:
        messagebox.showerror("Grading Error", f"No grading function available for challenge {challenge_number}.")
        return

    # Iterate through student folders
    for student_folder in os.listdir(folder_path):
        student_folder_path = os.path.join(folder_path, student_folder)
        
        if os.path.isdir(student_folder_path):
            # Look for an Excel file inside the student's folder
            for file in os.listdir(student_folder_path):
                if file.endswith(".xlsx"):
                    student_file_path = os.path.join(student_folder_path, file)
                    print(f"Grading {student_file_path}")

                    # Use the selected grading function
                    score, total_points, feedback = grading_function(solution_path, student_file_path)
                    percentage = (score / total_points) * 100 if total_points > 0 else 0

                    # Add the result to the grades list
                    grades.append({
                        "Student": student_folder,
                        "Score": score,
                        "Total Points": total_points,
                        "Percentage": percentage,
                        "Feedback": "; ".join(feedback)  # Join feedback items as a single string.
                    })
                    break

    # Convert grades list to a DataFrame and export to CSV
    df = pd.DataFrame(grades)
    output_file = os.path.join(output_path, "grades_report.csv")
    df.to_csv(output_file, index=False)
    messagebox.showinfo("Success", f"Grading complete! Report saved to: {output_file}")

# Function to set up the GUI
def setup_gui():
    root = Tk()
    root.title("Excel Grader")
    root.geometry("300x450")

    def select_folder():
        folder = filedialog.askdirectory()
        folder_label.config(text=folder)

    def select_solution():
        solution = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        solution_label.config(text=solution)

    def select_output():
        output = filedialog.askdirectory()
        output_label.config(text=output)

    def start_grading():
        folder_path = folder_label.cget("text")
        solution_path = solution_label.cget("text")
        output_path = output_label.cget("text")

        if not folder_path or not solution_path or not output_path:
            messagebox.showwarning("Input Error", "Please select all required paths.")
            return

        process_submissions(folder_path, solution_path, output_path)

    # Create GUI components
    Label(root, text="Select Student Submissions Folder:").pack(pady=5)
    folder_label = Label(root, text="", wraplength=350)
    folder_label.pack(pady=5)
    Button(root, text="Browse", command=select_folder).pack(pady=5)

    Label(root, text="Select Solution File:").pack(pady=5)
    solution_label = Label(root, text="", wraplength=350)
    solution_label.pack(pady=5)
    Button(root, text="Browse", command=select_solution).pack(pady=5)

    Label(root, text="Select Output Folder:").pack(pady=5)
    output_label = Label(root, text="", wraplength=350)
    output_label.pack(pady=5)
    Button(root, text="Browse", command=select_output).pack(pady=5)

    Button(root, text="Start Grading", command=start_grading).pack(pady=20)

    root.mainloop()

# Entry point for the program
if __name__ == "__main__":
    setup_gui()
