## Report Generation Application

This Python application generates a report in Microsoft Word format based on the information provided in an Excel spreadsheet. The report includes the longest, middle, and shortest paragraphs from the spreadsheet, along with appropriate headings and an image of a python at the bottom.

### Instructions

#### Excel Spreadsheet Preparation:

- Create an Excel spreadsheet with two columns.
- The first column contains paragraphs, and the second column contains corresponding paragraph titles.
- Ensure that each paragraph title contains one of the following words: "fish", "cheese", or "car".
- Include at least three paragraphs and three titles in no particular order.

#### Python Application Setup:

- Ensure you have Python 3.x installed on your system.
- Install the required Python packages by running the following command:
  ```bash
  pip install openpyxl python-docx
Running the Application:
Place the Excel spreadsheet named data.xlsx in the project directory.
Execute the Python script generate_report.py by running the following command:

python generate_report.py

After execution, the generated report will be saved as report.docx in the project directory.
