# JWFCSS Report Card Generator - User Manual

## Table of Contents

1. [Overview](#overview)
2. [Getting Started](#getting-started)
3. [How to Use](#how-to-use)
4. [Preparing Your Files](#preparing-your-files)
5. [Troubleshooting](#troubleshooting)
6. [FAQ](#faq)

---

## Overview

The **JWFCSS Report Card Generator** is a desktop application that automatically generates Word document report cards for students using data from an Excel gradesheet. Instead of manually creating individual reports, this tool processes all students at once and creates personalized report cards with:

- Student information (name, grades for each subject)
- Automatic comments based on average scores
- Academic statistics (total subjects, subjects passed)
- Behavioral assessments (deportment, responsibility, cooperation, etc.)
- Form teacher and principal comments

**Time Saved:** Generate 200+ report cards in minutes instead of hours!

---

## Getting Started

### Running the Application

#### Option 1: Windows

1. Double-click `Report_Card_Generator.exe`
2. The application window will open

#### Option 2: Mac/Linux with Python

1. Open Terminal/Command Prompt
2. Navigate to the application folder:
   ```bash
   cd /path/to/report_card_generator
   ```
3. Run:
   ```bash
   python3 report_card_gui.py
   ```

---

## How to Use

### Step 1: Select Your Excel Gradesheet

1. Click **"Browse Excel Gradesheet"** button
2. Navigate to your Excel file (must be `.xlsx` format)
3. Select the file and click "Open"
4. The application will automatically load all available sheet names

### Step 2: Choose a Sheet

1. Click the **"Sheet Name"** dropdown
2. Select the class/form you want to generate reports for (e.g., F1J, F2W)
3. The dropdown will show all sheets in your Excel workbook

### Step 3: Select Your Word Template

1. Click **"Browse Word Report Template"** button
2. Navigate to your Word template file (must be `.docx` format)
3. Select the template and click "Open"

**Important:** Your Word template must contain placeholder text in the format:

- `{{NAME}}` - Student name
- `{{MATH}}` - Mathematics grade
- `{{SAVG}}` - Student average
- See [Supported Placeholders](#supported-placeholders) for complete list

### Step 4: Choose Output Folder

1. Click **"Browse Output Folder"** button
2. Navigate to where you want the report cards saved
3. Select a folder and click "Open"

### Step 5: Generate Report Cards

1. Click the green **"Generate Report Cards"** button
2. The application will process all students and create individual report cards
3. A success message will appear showing how many reports were generated

---

## Preparing Your Files

### Excel Gradesheet Format

Your Excel file should have:

**Row 1 (Headers):** Student ID, First Name, Last Name, HFL, Literacy, Religion, English Language, English Literature, French, Spanish, Social Studies, History, Mathematics, Agricultural Science, Integrated Science, PE, IT, Visual Art, Home Economics, Total, Average, Total Subjects, Subjects Passed, Form Teacher Comments, Principal Comments, Attitude to Authority, Responsibility, Cooperation, Attitude to Corrections, Leadership, Deportment, Sociability, Initiative, Conflict Management, Application

**Rows 2+:** Student data

**Example:**

```
Student ID | First Name | Last Name | HFL | Literacy | ... | Average | ...
1          | John       | Smith     | 85  | 90       | ... | 87      | ...
2          | Mary       | Johnson   | 92  | 88       | ... | 89      | ...
```

**Requirements:**

- Use `.xlsx` format (Excel 2007+)
- First column must contain student numbers
- All grades must be numeric values
- Student names in columns 2 and 3

### Word Template Format

Create a Word document with your school letterhead and report card layout. Include placeholders for:

```
Student Name: {{NAME}}
Mathematics: {{MATH}}
Average: {{SAVG}}
```

**Important Placeholders:**

| Placeholder                 | Meaning                      |
| --------------------------- | ---------------------------- |
| `{{NAME}}`                  | Full student name            |
| `{{MATH}}`                  | Mathematics grade            |
| `{{ELANG}}`                 | English Language grade       |
| `{{ELIT}}`                  | English Literature grade     |
| `{{FR}}`                    | French grade                 |
| `{{SPN}}`                   | Spanish grade                |
| `{{SSTUD}}`                 | Social Studies grade         |
| `{{HIST}}`                  | History grade                |
| `{{HFL}}`                   | Health, Family & Life grade  |
| `{{LIT}}`                   | Literacy grade               |
| `{{REL}}`                   | Religion grade               |
| `{{AGR}}`                   | Agricultural Science grade   |
| `{{INTSCI}}`                | Integrated Science grade     |
| `{{PE}}`                    | Physical Education grade     |
| `{{IT}}`                    | Information Technology grade |
| `{{VA}}`                    | Visual Art grade             |
| `{{HE}}`                    | Home Economics grade         |
| `{{TOTAL}}`                 | Total score                  |
| `{{SAVG}}`                  | Student average              |
| `{{TOTAL_SUBJECTS}}`        | Total number of subjects     |
| `{{PASSED}}`                | Number of subjects passed    |
| `{{FORM_TEACHER_COMMENTS}}` | Form teacher's comments      |
| `{{PRINCIPAL_COMMENTS}}`    | Principal's comments         |
| `{{ATA}}`                   | Attitude to Authority        |
| `{{RESP}}`                  | Responsibility               |
| `{{CO_OP}}`                 | Cooperation                  |
| `{{ATC}}`                   | Attitude to Corrections      |
| `{{LEAD}}`                  | Leadership                   |
| `{{DEP}}`                   | Deportment                   |
| `{{SOC}}`                   | Sociability                  |
| `{{INIT}}`                  | Initiative                   |
| `{{CON_MGT}}`               | Conflict Management          |
| `{{APP}}`                   | Application                  |

### Automatic Comments

The application generates comments based on student average:

- **75+** - "Excellent performance. The student demonstrates strong understanding and consistent effort."
- **60-74** - "Good performance. Continued effort and focus will yield further improvement."
- **50-59** - "Satisfactory performance. Greater consistency and engagement are encouraged."
- **Below 50** - "Performance needs improvement. Increased effort and academic support are recommended."

---

## Troubleshooting

### "Could not find student data in Excel file"

**Problem:** The Excel file format is incorrect.
**Solution:**

- Ensure the first column contains student numbers (numeric values)
- Check that data starts in the correct row
- Verify the file is `.xlsx` format

### No sheet names appear in dropdown

**Problem:** Excel file wasn't loaded correctly.
**Solution:**

- Click "Browse Excel Gradesheet" again
- Make sure the file is `.xlsx` format
- Close the application and reopen it

### Report cards aren't generated

**Problem:** Required fields are missing.
**Solution:**

- Ensure all four sections are filled:
  - ✓ Excel file selected
  - ✓ Sheet name chosen
  - ✓ Word template selected
  - ✓ Output folder selected
- Check that students have names in the Excel file

### Missing or blank fields in report cards

**Problem:** Placeholders don't match or data is missing.
**Solution:**

- Check that all placeholders in your Word template use correct names
- Verify Excel file has data in all columns
- Ensure no cells are empty for required information

### Application crashes when generating

**Problem:** File is too large or permissions issue.
**Solution:**

- Check that output folder has write permissions
- Ensure Excel file is not open in another program
- Try with a smaller subset of students first

---

## FAQ

**Q: How many report cards can I generate at once?**
A: There's no limit! Generate 10, 100, or 1000+ at once. Processing time depends on your computer speed.

**Q: Can I edit the generated report cards?**
A: Yes! Each generated document is a standard Word file that can be edited, printed, or shared.

**Q: What if I have multiple classes in one Excel file?**
A: Each sheet in your Excel file represents a class. Use the dropdown to select which class you want to process.

**Q: Can I customize the comments?**
A: Yes! You can either:

1. Add `{{FORM_TEACHER_COMMENTS}}` and `{{PRINCIPAL_COMMENTS}}` placeholders to your template
2. Edit each document individually after generation

**Q: What file formats are supported?**
A:

- Excel: `.xlsx` only (not `.xls`)
- Word: `.docx` only (not `.doc`)

**Q: Can I run this on a Mac?**
A: Yes! The Python version works on Mac, Linux, and Windows.

**Q: Is my data secure?**
A: All processing happens on your computer. No data is sent anywhere.

**Q: How do I update the application?**
A: Visit the GitHub repository and download the latest version.

---

## Getting Help

If you encounter issues not covered in this manual:

1. Check the [Troubleshooting](#troubleshooting) section
2. Review the [Preparing Your Files](#preparing-your-files) section
3. Visit the GitHub repository: https://github.com/CAnuli-Ja/report_card_generator

---

**Version:** 1.0  
**Last Updated:** February 2026  
**School:** JWFCSS
