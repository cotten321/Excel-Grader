import openpyxl
import traceback
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.header_footer import _HeaderFooterPart

def grade_challenge_1_1(solution_path, student_path):
    try:
        # Load workbooks and select active sheets
        solution_wb = openpyxl.load_workbook(solution_path)
        student_wb = openpyxl.load_workbook(student_path)
        student_ws = student_wb.active

        # Initialize scoring variables
        score = 0
        total_points = 10
        feedback = []

        # Define expected headers and row count
        expected_headers = ["CustomerID", "FirstName", "LastName", "Email"]
        expected_row_count = 5  # 5 rows of data (not including the header row)

        # Check headers in the first row (A1:D1)
        student_headers = [student_ws.cell(row=1, column=col).value for col in range(1, 5)]
        if student_headers == expected_headers:
            score += 2  # Award 2 points for having the correct headers
        else:
            feedback.append("Headers are incorrect.")

        # Check for 5 rows of data below the headers (A2:D6)
        data_rows = [
            [student_ws.cell(row=row, column=col).value for col in range(1, 5)]
            for row in range(2, 7)  # Rows 2 through 6 (A2:D6)
        ]

        # Ensure that each row has data (not None)
        non_empty_rows = [row for row in data_rows if all(cell is not None for cell in row)]
        if len(non_empty_rows) == expected_row_count:
            score += 2  # Award 2 points for having the correct number of data rows
        else:
            feedback.append("Incorrect number of data rows")

        # Now compare the content of the student's data with the expected solution data
        solution_data = [
            ["CustomerID", "FirstName", "LastName", "Email"],
            [101, "John", "Doe", "johndoe@example.com"],
            [102, "Jane", "Smith", "janesmith@example.com"],
            [103, "Michael", "Johnson", "mjohnson@example.com"],
            [104, "Peter", "Parker", "pparker@dailybugle.com"],
            [105, "Tony", "Stark", "tstark@starkindustries.com"]
        ]

        # Compare the content row by row (data only, ignoring the headers)
        matching_cells = 0
        total_cells = 20  # 5 rows * 4 columns

        for r in range(5):
            for c in range(4):
                solution_value = solution_data[r + 1][c]  # Offset by 1 to skip header
                student_value = data_rows[r][c]

                if student_value == solution_value:
                    matching_cells += 1
                else:
                    feedback.append("Imported data is incorrect in 1(or more) rows.")

        # Award points based on the number of correctly matched cells
        content_points = (matching_cells / total_cells) * 6  # 6 points for content accuracy
        score += content_points

        return score, total_points, feedback
    except Exception as e:
        print(f"Error comparing workbooks: {e}")
        return 0, total_points  # In case of an error, give 0 score but consider total points
    

def grade_challenge_3_1(solution_path, student_path):
    try:
        # Load workbooks and select active sheets
        ##solution_wb = openpyxl.load_workbook(solution_path)
        student_wb = openpyxl.load_workbook(student_path)
        student_ws = student_wb.active

        # Initialize scoring variables
        score = 0
        total_points = 15  # Adjust based on grading
        feedback = []  
        
        # 1. Page Setup (4 points)
        # Check page orientation
        if student_ws.page_setup.orientation == "landscape":
            score += 1
        else:
            feedback.append("Incorrect page orientation")
            
            
        #DEBUGGING
        print("Debug - Page Setup Scaling Settings:")
        print("Orientation:", student_ws.page_setup.orientation)
        print("Fit to Width:", student_ws.page_setup.fitToWidth)
        print("Margins - Left:", student_ws.page_margins.left, "Right:", student_ws.page_margins.right)
        print("Margins - Top:", student_ws.page_margins.top, "Bottom:", student_ws.page_margins.bottom)
        print("Row Height for Row 1:", student_ws.row_dimensions[1].height)
        print("Column Width for Column A:", student_ws.column_dimensions["A"].width)

            
        # Check fit to width/height with defaults
        fit_to_width = student_ws.page_setup.fitToWidth or 1
        fit_to_height = student_ws.page_setup.fitToHeight or 1
        if fit_to_width == 1 and fit_to_height == 1:
            score += 1
        else:
            feedback.append("Incorrect page orientation")
            
            
        # Scale to Fit to 1 page is ungradeable at this time

        # Check for narrow margins
        if (round(float(student_ws.page_margins.left), 2) == 0.25 and 
            round(float(student_ws.page_margins.right), 2) == 0.25 and
            round(float(student_ws.page_margins.top), 2) == 0.75 and 
            round(float(student_ws.page_margins.bottom), 2) == 0.75):
            score += 2

        # Scale to Fit to 1 page is ungradeable at this time

        # Check for narrow margins
        if (student_ws.page_margins.left == 0.25 and 
            student_ws.page_margins.right == 0.25 and
            student_ws.page_margins.top == 0.75 and 
            student_ws.page_margins.bottom == 0.75):
            score += 1.5

        else:
            feedback.append("Margins are not set to Narrow.")

        # 2. Row Height and Column Width (2 points)
        # Check row height
        if student_ws.row_dimensions[1].height == 30:
            score += 2
        else:
            feedback.append("Row height for the header row (Row 1) is incorrect; Expected 30 points.")
            
        # Check column widths for column A (Allows for a small tolerance to mitigate Excels float points)
        if 19 <= student_ws.column_dimensions["A"].width <=23:
            score += 1.5
        else:
            feedback.append("Incorrect column width for column A; Expected 20.")
        
        # Check column widths for column B-J (Allows for a small tolerance)    
        for col in range(2, 10):
            col_letter = get_column_letter(col)
            if not (13 <= student_ws.column_dimensions[col_letter].width <= 16.5):
                feedback.append(f"Incorrect width for column {col_letter}; Expected 15.")
                break
        else:
            score += 2

        # 3. Headers and Footers (2 points)
        # Check if there's text in the left part of the header
        if student_ws.oddHeader.left and hasattr(student_ws.oddHeader.left, 'text') and student_ws.oddHeader.left.text.strip():
            score += 2
        else:
            feedback.append("No text found in the left side of the header.")

        # Check for Date in the center part of the header
        if student_ws.oddHeader.center and hasattr(student_ws.oddHeader.center, 'text') and "&D" in student_ws.oddHeader.center.text:
            score += 1
        else:
            feedback.append("Header does not contain date in center.")

        # Check if the file name is in the right side of the header
        if student_ws.oddHeader.right and hasattr(student_ws.oddHeader.right, 'text') and "&F" in student_ws.oddHeader.right.text:
            score += 1
        else:
            feedback.append("Header does not contain file name on the right.")

        # Footer checks for page numbering
        if student_ws.oddFooter.left and hasattr(student_ws.oddFooter.left, 'text') and "&P" in student_ws.oddFooter.left.text:
            score += 1
        else:
            feedback.append("Footer does not contain page number on the left.")

        if student_ws.oddFooter.right and hasattr(student_ws.oddFooter.right, 'text') and "&N" in student_ws.oddFooter.right.text:
            score += 1
        else:
            feedback.append("Footer does not contain total number of pages on the right.")

        if student_ws.oddFooter.right and hasattr(student_ws.oddFooter.right, 'text') and "&N" in student_ws.oddFooter.right.text:
            score += 1
        else:
            feedback.append("Footer does not contain total number of pages on the right.")

        # 4. Options and Views (1 point)
        if not student_ws.sheet_view.showGridLines and not student_ws.sheet_view.showRowColHeaders:
            score += 2
        else:
            feedback.append("Gridlines or headings are not hidden.")

        return score, total_points, feedback

    except Exception as e:
        print(f"Error comparing workbooks for Assignment 3.1: {e}")
        traceback.print_exc()
        return 0, total_points, ["An error occurred during grading."]