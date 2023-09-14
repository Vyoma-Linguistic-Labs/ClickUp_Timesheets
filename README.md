# ClickUp Timesheets Generation
A script to generate timesheets in the required format, at the click of a button. Internal tool for the employees.

## Installation and Setup (to be done once)

1. Install Anaconda software from this [link](www.anaconda.com).
2. Click on the downloaded `.exe` and go through the steps to complete the installation.
3. Press the 'Start' button (in Windows) and type 'Anaconda Powershell Prompt'.
4. Click to open. You will be met with a black terminal.
5. Copy paste this command in the terminal: `pip install pandas datetime requests tkcalendar` (without quotes).
6. Press ENTER.
7. Store the Python program `Generate_Timesheet.py` in any folder.

## How to Run it every time for TS?

1. Press the 'Start' button (in Windows) and type 'Anaconda Powershell Prompt'.
2. Click to open. You will be met with a black terminal.
3. Use the `cd` command to traverse to the folder that contains the Python program. For example, `cd Downloads` and press ENTER.
4. Then type `python Generate_Timesheet.py` and press ENTER.
5. You will get a window where you input the:
   - Start Date
   - End Date
   - Employee ID
6. Your timesheet will be output as an Excel in the same folder as the program. The file name will be displayed on the window once it is created.
7. You can close the window and exit the program.
8. Do not forget to copy the data into your Google Sheets and update the submission status.

## Special notes

- The program will output the total number of hours for that week on the terminal. This might be useful to generate the number of hours worked for any period, not only weekly.
- The second output is the amount of time in seconds the program took to run 🙂.
- The Google Sheets for status submission will automatically open at the end.
- Copy from Excel and Paste only values into the Google Sheets (use Ctrl + Shift + V).
- Watch the project columns and copy-paste the values accordingly.

## What is required in the ClickUp tasks for Timesheets

For each task you work on any day, add the hours and minutes in the 'Time Tracked' field for that date (in 'When:' field).
