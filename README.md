# Automated-Excel-Processing-and-Chart-Creation-with-Python
**When compared to manual processing, this code automates the process of updating data in an Excel sheet and generating a bar chart, saving time and boosting repeatability, flexibility, and accuracy.**

***********************************************************************************************************************************************************************

# Requirements
- The code has been tested on **Python 3.7**, however it should work with other Python 3 versions as well.

- **Microsoft Excel** (or a compatible spreadsheet program)

- The code uses the following packages:
   - **openpyxl**
   
  You can install these packages by running the following command in your terminal :
  **'pip install openpyxl'**

***********************************************************************************************************************************************************************

**This repository provides a Python script that processes an Excel worksheet using the openpyxl package. The script does the following:**

1. Loads an Excel workbook using the openpyxl library.
2. Accesses the first sheet in the workbook and loops through each row starting from the second row to the last row.
3. For each row, the script multiplies the value in column 3 by 0.9 and writes the result to column 4.
4. Creates a bar chart using the data in column 4 and adds it to the first sheet in the workbook.
5. Saves the processed workbook with the same filename.

 **The script assumes that the workbook being processed has the following characteristics:**
 
 - Sheet1 is the name of the first sheet in the workbook.
 
 - The data that must be processed begins with the second row and continues to the last row.
 
 - The value in column 3 must be multiplied by 0.9 and reported in column 4.
 
 - Starting with cell "e2," add the bar chart to the first sheet.
 
 **Advantages of this code include:**
 
 - **Automation:** The code automates the process of updating data in an Excel sheet and making a bar chart, which would normally be done manually.
 
 - **Time-saving:** When compared to manually processing the worksheet, automating the procedure saves time.
 
 - **Reproducibility:** Because the code may be run several times on the same or various workbooks with comparable structures, the procedure is extremely repeatable.
 
 - **Flexibility:** By altering the values in the formulae and tweaking the chart settings, the code may be readily adjusted to accommodate new circumstances.

 - **Consistency:** The code guarantees that the data are handled consistently and precisely, decreasing the possibility of mistakes that can occur when manual processing is used.
 
 ***********************************************************************************************************************************************************************
 
 *I welcome any and all contributions to this project, large or little. Every contribution, whether it's resolving an issue, adding a feature, or enhancing the description, helps to improve this project.*
