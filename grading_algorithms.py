import openpyxl
import traceback
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.header_footer import _HeaderFooterPart

def grade_challenge_1_1(student_path):
    try:
        # Load workbooks and select active sheets
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
    
def grade_challenge_2(student_path):
    try:
        # Load the student workbook and select the active sheet
        student_wb = openpyxl.load_workbook(student_path)
        student_ws = student_wb.active

        # Initialize scoring variables
        score = 0
        total_points = 15  # Adjust based on grading
        feedback = []

        # 1. Searching for data (4 points)
        # Check if the cell with the employee making $75,000 is highlighted in yellow
        target_cell = "B144"  # The cell containing the employee's name making $75,000
        if student_ws[target_cell].fill.start_color.rgb == "FFFFFF00":  # Yellow color in Excel
            score += 2
        else:
            feedback.append("Cell B144 (name of employee with $75,000 salary) is not highlighted in yellow.")

        # 2. Navigating to named cells/ranges (5 points)
        # Check if "EmployeeInfo" named range exists
        if "EmployeeInfo" in student_wb.defined_names:
            score += 2
        else:
            feedback.append("Named range 'EmployeeInfo' not found.")

        # Check if the named range "EmployeeInfo" has the correct range A1:B201
        if "EmployeeInfo" in student_wb.defined_names:
            defined_range = student_wb.defined_names["EmployeeInfo"].attr_text
            if defined_range.endswith("!$A$1:$B$201"):
                score += 2
            else:
                feedback.append("Named range 'EmployeeInfo' does not refer to cells A1:B201.")

        # Check if the font for the range is Times New Roman
        cells_in_range = student_ws["A1:B201"]
        font_correct = all(cell.font.name == "Times New Roman" for row in cells_in_range for cell in row)
        if font_correct:
            score += 2
        else:
            feedback.append("Font for 'EmployeeInfo' named range is not set to Times New Roman.")

        # 3. Hyperlinks (6 points)
        # Check if the hyperlink in cell B204 has been removed
        if not student_ws["B204"].hyperlink:
            score += 3
        else:
            feedback.append("Hyperlink in cell B204 has not been removed.")

        # Check if the hyperlink was added to cell F4 with the correct URL and display text
        cell_f4 = student_ws["F4"]
        if cell_f4.hyperlink and cell_f4.hyperlink.target == "https://www.examplecompany.com" and cell_f4.value == "Example Company":
            score += 3
        else:
            feedback.append("Cell F4 does not have the correct hyperlink and display text.")

        return score, total_points, feedback

    except Exception as e:
        print(f"Error comparing workbooks for Assignment 2: {e}")
        traceback.print_exc()
        return 0, total_points, ["An error occurred during grading."]

def grade_challenge_3_1(student_path):
    try:
        # Load workbooks and select active sheets
        student_wb = openpyxl.load_workbook(student_path)
        student_ws = student_wb.active

        # Initialize scoring variables
        score = 0
        total_points = 20  # Adjust based on grading
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
        if 19 <= student_ws.column_dimensions["A"].width <=21:
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

def grade_project_1(student_path):
    try:
        # First pass: Check formulas (data_only=False)
        wb_formulas = openpyxl.load_workbook(student_path, data_only=False)
        # Second pass: Check values (data_only=True)
        wb_values = openpyxl.load_workbook(student_path, data_only=True)

        # Initialize scoring variables
        score = 0
        total_points = 50  # Base points
        feedback = []

        # 1. Sheet Structure Check (2 points)
        sheet_names = wb_formulas.sheetnames
        if len(sheet_names) == 2 and "CoffeeData" in sheet_names and "Analysis" in sheet_names:
            score += 2
        else:
            feedback.append("Incorrect number of sheets or sheet names.")

        # 2. Named Range Check
        coffee_data_sheet = wb_formulas["CoffeeData"]
        analysis_sheet = wb_formulas["Analysis"]

        named_ranges = wb_formulas.defined_names
        reviews_range_found = False
        try:
            for name in named_ranges.values():
                if "Reviews" in str(name.name) and ("simplified_coffee[review]" in str(name.attr_text) or 
                                                    "simplified_coffee[Review]" in str(name.attr_text)):
                    reviews_range_found = True
                    break
        except Exception as e:
            print(f"Error checking named ranges: {e}")
            feedback.append(f"Error checking named ranges: {e}")

        if reviews_range_found:
            score += 2
        else:
            feedback.append("Named range 'Reviews' not found or incorrect.")

        # Individual Calculations Grading
        calc_checks = [
            # Cell B1: Average
            {
                'cell': 'B1', 
                'valid_formulas': [
                    '=AVERAGE(CoffeeData!F:F)', 
                    '=AVERAGE(simplified_coffee[rating])'
                ],
                'expected_values': 10.48,
                'points': {
                    'formula': 2,
                    'value': 2,
                    'decimal_reduction': 1
                }
            },
            # Cell B2: Max Rating
            {
                'cell': 'B2', 
                'valid_formulas': [
                    '=MAX(CoffeeData!G:G)', 
                    '=MAX(simplified_coffee[rating])'
                ],
                'expected_value': 97,
                'points': {
                    'formula': 2,
                    'value': 2
                }
            },
            # Cell B3: Min Rating
            {
                'cell': 'B3', 
                'valid_formulas': [
                    '=MIN(CoffeeData!G:G)', 
                    '=MIN(simplified_coffee[rating])'
                ],
                'expected_value': 84,
                'points': {
                    'formula': 2,
                    'value': 2
                }
            },
            # Cell B4: Count of Reviews
            {
                'cell': 'B4', 
                'valid_formulas': [
                    '=COUNTA(Reviews)',
                    '=COUNTA(reviews)',
                    '=COUNTA(simplified_coffee[review])'
                ],
                'expected_value': 1246,
                'points': {
                    'formula_standard': 2,
                    'formula_alternative': 1,
                    'value': 2
                }
            },
            # Cell B5: Sum of 100g USD
            {
                'cell': 'B5', 
                'valid_formulas': [
                    '=SUM(CoffeeData!F:F)', 
                    '=SUM(simplified_coffee[100g_USD])'
                ],
                'points': {
                    'formula': 2,
                    'value': 2
                }
            },
            # Cell E1: Unique Roasters Count
            {
                'cell': 'E1', 
                'valid_formulas': ['=COUNTA(UNIQUE(Roasters))'],
                'expected_value': 296,
                'points': {
                    'formula': 2,
                    'value': 2
                }
            },
            # Cell E2: Max Review Length
            {
                'cell': 'E2', 
                'valid_formulas': ['=MAX(LEN(Reviews))'],
                'expected_value': 509,
                'points': {
                    'formula': 3,
                    'value': 2
                }
            },
            # Cell E3: Average Rating
            {
                'cell': 'E3', 
                'valid_formulas': [
                    '=AVERAGE(CoffeeData!G:G)', 
                    '=AVERAGE(simplified_coffee[Rating])'
                ],
                'expected_value': 93.3,
                'points': {
                    'formula': 2,
                    'value': 2
                }
            }
        ]

        # Perform detailed checks for each calculation
        for calc in calc_checks:
            cell = calc['cell']
            try:
                cell_obj = analysis_sheet[cell]
                cell_value = wb_values['Analysis'][cell].value

                # Formula check
                if 'valid_formulas' in calc:
                    formula_points = 0
                    if cell_obj.data_type == 'f':
                        # Convert ArrayFormula to string safely
                        formula = str(cell_obj.value).strip().replace('_xlfn.', '')
                        
                        print(f"Debugging cell {cell}:")
                        print(f"Received formula: {formula}")
                        print(f"Expected formulaas: {calc['valid_formulas']}")
                        
                        for valid_formula in calc['valid_formulas']:
                            if formula == valid_formula:
                                score += calc.get('points', {}).get('formula', 2)
                                formula_points = calc.get('points', {}).get('formula', 2)
                                break
                        
                        if formula_points == 0:
                            feedback.append(f"Incorrect formula in cell {cell}")

                # Value check
                if 'expected_value' in calc:
                    try:
                        # Safely handle potential None or non-numeric values
                        if cell_value is not None:
                            rounded_value = round(float(cell_value), 2)
                            if rounded_value == calc['expected_value']:
                                score += calc.get('points', {}).get('value', 2)
                            else:
                                feedback.append(f"Incorrect value in cell {cell}. Expected {calc['expected_value']}, got {rounded_value}")
                        else:
                            feedback.append(f"Cell {cell} contains no value")
                    except (TypeError, ValueError) as ve:
                        feedback.append(f"Error processing value in cell {cell}: {ve}")

            except Exception as e:
                print(f"Error processing cell {cell}: {e}")
                feedback.append(f"Error processing cell {cell}: {e}")

        return min(score, total_points), total_points, feedback

    except Exception as e:
        print(f"Error grading Project 1: {e}")
        traceback.print_exc()
        return 0, total_points, [f"An error occurred during grading: {str(e)}"]