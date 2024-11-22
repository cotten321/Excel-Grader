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
                    'formula': 1,
                    'value': 2
                }
            },
            # Cell B4: Count of Reviews
            {
                'cell': 'B4', 
                'expected_value': 1246,
                'points': {
                    'value': 4
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
                'expected_value': 509,
                'points': {
                    'value': 5
                }
            },
            # Cell E3: Average Rating
            {
                'cell': 'E3', 
                'valid_formulas': [
                    '=AVERAGE(CoffeeData!G:G)', 
                    '=AVERAGE(simplified_coffee[Rating])'
                ],
                'expected_value': 93.31,
                'points': {
                    'formula': 1,
                    'value': 2
                }
            },
            # TABLE GRADING: H4 (Top of the table)
            {
                'cell': 'H4', 
                'valid_formulas': ['=CoffeeData!A2'],
                'expected_value': "Ethiopia Shakiso Mormora",
                'points': {
                    'formula': 1,
                    'value': 1
                }
            },
            # Cell H13
            {
                'cell': 'H13', 
                'valid_formulas': ['=CoffeeData!A11'],
                'expected_value': "Ethiopia Yirgacheffe Washed G1",
                'points': {
                    'formula': 1,
                    'value': 1
                }
            },
            # Cell I4 (Top of the table)
            {
                'cell': 'I4', 
                'valid_formulas': ['=CoffeeData!F2*$I$1'],
                'expected_value': 235.00,
                'points': {
                    'formula': 2,
                    'value': 1
                }
            },
            # Cell I13 (Bottom of the table)
            {
                'cell': 'I13', 
                'valid_formulas': ['=CoffeeData!F11*$I$1'],
                'expected_value': 343.50,
                'points': {
                    'formula': 2,
                    'value': 1
                }
            },
            # Cell J4 (Affordable Validation TOP)
            {
                'cell': 'J4', 
                'valid_formulas': ['=IF(Table2[[#This Row],[USD per \'\'x\'\' Units]] <= 250, "Yes", "No")'],
                'expected_value': "Yes",
                'points': {
                    'formula': 2,
                    'value': 1
                }
            },
            # Cell J9 (Affordable Validation MIDDLE)
            {
                'cell': 'J9', 
                'valid_formulas': ['=IF(Table2[[#This Row],[USD per \'\'x\'\' Units]] <= 250, "Yes", "No")'],
                'expected_value': "No",
                'points': {
                    'formula': 2,
                    'value': 1
                }
            },
            # Cell J13 (Affordable Validation BOTTOM)
            {
                'cell': 'J13', 
                'valid_formulas': ['=IF(Table2[[#This Row],[USD per \'\'x\'\' Units]] <= 250, "Yes", "No")'],
                'expected_value': "No",
                'points': {
                    'formula': 2,
                    'value': 1
                }
            },
            
            
            #--------BONUS QUESTIONS---------------------
            # Cell C10: Ethiopian Light Roast
            {
                'cell': 'C10', 
                'valid_formulas': [
                    '=AVERAGEIFS(CoffeeData!G:G, CoffeeData!C:C, "Light", CoffeeData!E:E, "Ethiopia")',
                    '=ROUND(AVERAGEIFS(CoffeeData!G:G, CoffeeData!C:C, "Light", CoffeeData!E:E, "Ethiopia"), 1)',
                    '=ROUND(AVERAGE(IF((CoffeeData!C:C="Light")*(CoffeeData!E:E="Ethiopia"), CoffeeData!G:G)), 1)',
                    '=ROUND(AVERAGE(FILTER(CoffeeData!G:G, (CoffeeData!C:C="Light")*(CoffeeData!E:E="Ethiopia"))), 1)',
                    '=ROUND(SUMIFS(CoffeeData!G:G, CoffeeData!C:C, "Light", CoffeeData!E:E, "Ethiopia") / COUNTIFS(CoffeeData!C:C, "Light", CoffeeData!E:E, "Ethiopia"), 1)'
                ],
                'expected_value': 93.65,
                'points': {
                    'formula': 2.5,
                    'value': 5
                }
            },
            # Cell C11: Ethiopian Light Roast
            {
                'cell': 'C11', 
                'expected_value': 92,
                'points': {
                    'value': 7.5
                }
            },
            
        ]

        # Helper function to compare values with type checking
        def compare_values(actual_value, expected_value):
            if isinstance(expected_value, (int, float)):
                try:
                    if actual_value is None:
                        return False
                    actual_float = float(actual_value)
                    return round(actual_float, 2) == round(float(expected_value), 2)
                except (TypeError, ValueError):
                    return False
            else:
                # For text comparisons, convert both to strings and compare
                return str(actual_value).strip() == str(expected_value).strip()

        # Perform detailed checks for each calculation
        for calc in calc_checks:
            cell = calc['cell']
            try:
                cell_obj = analysis_sheet[cell]
                cell_value = wb_values['Analysis'][cell].value

                # Formula check
                if 'valid_formulas' in calc:
                    formula_points = 0
                    formula_feedback = []
                    if cell_obj.data_type == 'f':
                        formula = str(cell_obj.value).strip().replace('_xlfn.', '')
                        
                        print(f"Debugging cell {cell}:")
                        print(f"Received formula: {formula}")
                        print(f"Expected formulas: {calc['valid_formulas']}")
                        
                        for valid_formula in calc['valid_formulas']:
                            if formula == valid_formula:
                                points_to_add = calc.get('points', {}).get('formula', 0)
                                score += points_to_add
                                formula_points = points_to_add
                                break
                        
                        if formula_points == 0:
                            formula_feedback = [
                                f"Cell {cell} Formula Check:",
                                f"  - Received: {formula}",
                                f"  - Expected one of: {', '.join(calc['valid_formulas'])}"
                            ]
                            feedback.extend(formula_feedback)
                    else:
                        feedback.append(f"Cell {cell}: No formula found (cell contains static value)")

                # Value check with type handling
                if 'expected_value' in calc:
                    value_match = compare_values(cell_value, calc['expected_value'])
                    if value_match:
                        score += calc.get('points', {}).get('value', 0)
                    else:
                        value_feedback = [
                            f"Cell {cell} Value Check:",
                            f"  - Received: {cell_value}",
                            f"  - Expected: {calc['expected_value']}"
                        ]
                        feedback.extend(value_feedback)

            except Exception as e:
                print(f"Error processing cell {cell}: {e}")
                feedback.append(f"Error processing cell {cell}: {e}")

        return score, total_points, feedback

    except Exception as e:
        print(f"Error grading Project 1: {e}")
        traceback.print_exc()
        return 0, total_points, [f"An error occurred during grading: {str(e)}"]
      
#def grade_project_3(student_path):
    