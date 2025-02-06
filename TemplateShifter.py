import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.datavalidation import DataValidation
import os
import uuid

def apply_formulas_to_range(file_path, col_range, row_range, review_col,Rev):
    # Load workbook
    wb = load_workbook(file_path)
    sheet = wb.active

    # Parse column range and row range
    start_col, end_col = col_range.split('-')
    start_row, end_row = map(int, row_range.split('-'))

    start_col_index = column_index_from_string(start_col)
    end_col_index = column_index_from_string(end_col)

    import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.datavalidation import DataValidation
import os
import uuid

def apply_formulas_to_range(file_path, col_range, row_range, review_col, Rev):
    # Load workbook
    wb = load_workbook(file_path)
    sheet = wb.active

    # Parse column range and row range
    start_col, end_col = col_range.split('-')
    start_row, end_row = map(int, row_range.split('-'))

    start_col_index = column_index_from_string(start_col)
    end_col_index = column_index_from_string(end_col)

    # Delete data in the specified criteria column range from row 2 to end_row
    for col_idx in range(start_col_index, end_col_index + 1):
        col_letter = get_column_letter(col_idx)
        for row_idx in range(2, end_row + 1):  # Deleting data from row 2 to end_row
            sheet[f"{col_letter}{row_idx}"].value = None

    # Define new headers to be written in row 2
    headers = [
        "Question Accuracy", "Question Distribution", "Answer Accuracy", "Answer Explanation",
        "Tagging bloom level", "Tagging complexity level", "Distractors", "Learning Outcome",
        "No Repetition of PR Questions", "Topic Tagging", "Language and Grammar",
        "Copy Editing"
    ]
    
    # Write headers in row 2
    for col_idx, header in zip(range(start_col_index, end_col_index + 1), headers):
        sheet[f"{get_column_letter(col_idx)}2"].value = header
    
    saved_values = {"B": {}, "C": {},"H":{},"I":{}}  # Dictionary to store values for columns B and C
    if Rev in [ "R2", "R3"]:
        repowb = load_workbook(file_path,data_only=True)
        rub = repowb['AQR Rubrics']
        repo = repowb['Report Format']
        for col_idx in range(start_col_index, end_col_index + 1):
            col_letter = get_column_letter(col_idx)
            for row_idx in range(start_row, end_row + 1):
                cell = sheet[f"{col_letter}{row_idx}"]
                cell.value = None  # Clear the cell's value, removing formulas or text

        if Rev in ["R2", "R3"]:
            # Copy existing plain values from Column B
            for row_idx in range(start_row, end_row + 1):
                cell = repo[f"B{row_idx}"]
                saved_values["B"][row_idx] = cell.value
            for row_idx in range(start_row, end_row + 1):
                cell = rub[f"H{row_idx}"]
                saved_values["H"][row_idx] = cell.value  # Extract the plain value

        if Rev == "R3":
            # Copy existing plain values from Column C
            for row_idx in range(start_row, end_row + 1):
                cell = repo[f"B{row_idx}"]
                saved_values["B"][row_idx] = cell.value
            for row_idx in range(start_row, end_row + 1):
                cell = repo[f"C{row_idx}"]
                saved_values["C"][row_idx] = cell.value
            for row_idx in range(start_row, end_row + 1):
                cell = rub[f"H{row_idx}"]
                saved_values["H"][row_idx] = cell.value
            for row_idx in range(start_row, end_row + 1):
                cell = rub[f"I{row_idx}"]
                saved_values["I"][row_idx] = cell.value
        

    
    # Define formulas
    formulas = [
    f'=IF(SUM(ISNUMBER(SEARCH("Qs:", {review_col}ROW)) + ISNUMBER(SEARCH("QA:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("QD:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("AA:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("AE:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Bloom:", {review_col}ROW)) + ISNUMBER(SEARCH("BT:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Comp:", {review_col}ROW)) + ISNUMBER(SEARCH("CT:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Dis:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("LO:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Repq:", {review_col}ROW)) + ISNUMBER(SEARCH("QR:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("TopicT:", {review_col}ROW)) + ISNUMBER(SEARCH("TT:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Lang:", {review_col}ROW)) + ISNUMBER(SEARCH("LG:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")',
    f'=IF(SUM(ISNUMBER(SEARCH("Ced:", {review_col}ROW)) + ISNUMBER(SEARCH("CE:", {review_col}ROW)) + ISNUMBER(SEARCH("Reject:", {review_col}ROW)))>0, "No", "Yes")'
]

    # Apply formulas to the specified range
    # Apply formulas to the specified range
    formula_index = 0
    for col_idx in range(start_col_index, end_col_index + 1):
        col_letter = get_column_letter(col_idx)
        for row_idx in range(start_row, end_row + 1):
            # Replace "ROW" with the exact row number
            formula = formulas[formula_index % len(formulas)]
            formula_with_row = formula.replace("ROW", str(row_idx))
            formula_with_review_col = formula_with_row.replace("{review_col}", review_col)
            sheet[f"{col_letter}{row_idx}"] = formula_with_review_col
        formula_index += 1


    # Add data validation for the specified range
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv.prompt = "Please select Yes or No"
    dv.promptTitle = "Valid Options"

    # Add validation ranges explicitly
    for col_idx in range(start_col_index, end_col_index + 1):
        col_letter = get_column_letter(col_idx)
        dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")

    sheet.add_data_validation(dv)

    # Add COUNTIF formula in row 39 and percentage formula in row 40
    # Add COUNTIF and percentage formula one row after the end_row
    count_row = end_row + 1
    percentage_row = end_row + 2

    for col_idx in range(start_col_index, end_col_index + 1):
        col_letter = get_column_letter(col_idx)
        countif_formula = f'=COUNTIF({col_letter}{start_row}:{col_letter}{end_row}, "Yes")'
        percentage_formula = f'=({col_letter}{count_row}/{end_row - start_row + 1})*100'
        sheet[f"{col_letter}{count_row}"] = countif_formula
        sheet[f"{col_letter}{percentage_row}"] = percentage_formula


    # Save workbook
    # Copy additional sheets from AQR file
    aqr_wb = load_workbook('AMT_AQR.xlsx')

    for sheet_name in ["AQR Rubrics", "Report Format"]:
        if sheet_name in aqr_wb.sheetnames:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            source_sheet = aqr_wb[sheet_name]
            target_sheet = wb.create_sheet(title=sheet_name)

            # Copy data and formatting manually
            for row in source_sheet.iter_rows():
                for cell in row:
                    target_cell = target_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        target_cell._style = cell._style
    def beautify_sheet(sheet, title_row=1):
        # Apply header formatting
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4F81BD")
        header_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for cell in sheet[title_row]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

        # Adjust column widths
        for column_cells in sheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            sheet.column_dimensions[column_letter].width = max_length + 2
    # Dynamically link percentages to "AQR Rubrics" and "Report Format" sheets
    # Dynamically link percentages to "AQR Rubrics" and "Report Format" sheets
        # Dynamically link percentages to "AQR Rubrics" and "Report Format" sheets
    beautify_sheet(wb["AQR Rubrics"])
    beautify_sheet(wb["Report Format"])
    rubrics_sheet = wb["AQR Rubrics"]
    report_sheet = wb["Report Format"]
    if Rev == "R1":
        RubSheetCol = 'H'
        RepSheetCol = 'B'
    elif Rev == "R2":
        RubSheetCol = 'I'
        RepSheetCol = 'C'
    else:
        RubSheetCol = 'J'
        RepSheetCol = 'D'

    if Rev == "R2":
        # Restore Column B values
        for row_idx, value in saved_values["B"].items():
            if value is not None:
                report_sheet[f"B{row_idx}"].value = value
        for row_idx, value in saved_values["H"].items():
            if value is not None:
                rubrics_sheet[f"H{row_idx}"].value = value

    if Rev == "R3":
        # Restore Column C values
        for row_idx, value in saved_values["B"].items():
            if value is not None:
                report_sheet[f"B{row_idx}"].value = value
        for row_idx, value in saved_values["C"].items():
            if value is not None:
                report_sheet[f"C{row_idx}"].value = value
        for row_idx, value in saved_values["H"].items():
            if value is not None:
                rubrics_sheet[f"H{row_idx}"].value = value
        for row_idx, value in saved_values["I"].items():
            if value is not None:
                rubrics_sheet[f"I{row_idx}"].value = value
    current_row = 4  # Start row for the target sheets
    for col_idx in range(start_col_index, end_col_index + 1):
        col_letter = get_column_letter(col_idx)
        
        # Skip row 13 but adjust the mapping
        if current_row == 11:
            current_row += 1
        if current_row == 16:
            current_row += 1
        
        # Add formula linking percentage to the target sheets
        rubrics_sheet[f"{RubSheetCol}{current_row}"] = (
                f"=IFERROR("
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 30, 1, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 50, 2, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 70, 3, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 90, 4, 5)))), \"\")"
        )
        report_sheet[f"{RepSheetCol}{current_row}"] = (
                f"=IFERROR("
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 30, 1, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 50, 2, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 70, 3, "
                f"IF({sheet.title}!{col_letter}{percentage_row} <= 90, 4, 5)))), \"\")"
            )
        
        current_row += 1  # Move to the next row in the target sheets



    # Save workbook
    output_path = file_path.replace(".xlsx", "_processed.xlsx")
    wb.save(output_path)
    return output_path

# Streamlit app
st.title("Formula Application Tool with NEW AQR RUBRIC")
st.write("Instructions for using the tool:")
st.write("1. Upload an Excel file.")
st.write("2. Check if the provided AMT file have all the necessary columns like Question Accuracy ,Learning Outcome Accuracy ,No Repetition of PR Question ,Question Distribution ,Answer Accuracy ,Answer Explanation Accuracy ,Tagging bloom level ,Tagging complexity level ,Distractors ,Topic Tagging ,Language and Grammar ,No Plagiarism ,Copy Editing")
st.write("3. If any column is missing add the column in the AMT file before uploading and it should be in the same order as mentioned above.")
st.write("4. Enter the column range and row range where you want to apply the formulas. (for e.g : Question Accuracy (AQ) to Copy Editing (BC) )")
st.write("5. Enter the row range where you want to apply the formulas. (for e.g : 3 to 38)")
st.write("6. Enter the column where the review specific comment is present. (for e.g : AK)")
st.write("7. For R2, R3 Add the R2 and R3 Comment Column before applying formula.")
st.write("8. If the formula is not getting applied or some issue happens. Delete the Yes or No from the criteria column and then do the process ")
st.write("9. Click on 'Apply Formula' button to apply the formulas.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    col_range = st.text_input("Enter Criteria Column Range (e.g., A-Z):")
    row_range = st.text_input("Enter Row Range (e.g., 3-38):")
    review_col = st.text_input("Enter Review Specific Comment Column (e.g., AK):")
    target_column_selection = st.selectbox("Select R1, R2, or R3:", ["R1", "R2", "R3"])
    if st.button("Apply Formula"):
        if col_range and row_range and review_col:
            # Save uploaded file to a temporary unique path
            unique_id = str(uuid.uuid4())
            temp_file_path = f"temp_{unique_id}.xlsx"
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            try:
                # Apply formulas and generate the output file
                output_path = apply_formulas_to_range(temp_file_path, col_range, row_range, review_col,target_column_selection)

                # Generate output file name using uploaded file name
                output_file_name = uploaded_file.name.replace(".xlsx", "_processed.xlsx")

                # Streamlit success message and download button
                st.success(f"Formulas applied successfully! Download the file below.")
                st.download_button(
                    label="Download Processed File",
                    data=open(output_path, "rb"),
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Clean up temporary file
                os.remove(temp_file_path)

            except Exception as e:
                st.error(f"An error occurred: {e}")

        else:
            st.error("Please provide all inputs.")
