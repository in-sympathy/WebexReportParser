Webex Attendance Processing Office Script

This repository contains an Office Script for Excel designed to automate the processing of a Webex attendance report CSV. The script performs a series of actions to clean, consolidate, and format your attendance data so that it’s easy to review, copy, or further analyze. In particular, it:

Freezes the header row.
Adds new columns:
Training Duration (filled with a default value of 90)
Percentage (calculated dynamically as =[@[Attendance Duration]]/[@[Training Duration]])
Converts the range into an Excel table.
Cleans the "Attendance Duration" column by removing the " mins" text.
Consolidates duplicate rows by summing durations based on the "Attendee Email" column.
Sorts the table by "Attendance Duration" in ascending order.
Applies conditional formatting:
Rows with a percentage below 80% are filled red.
Rows with a percentage of 80% or above are filled green.
Filters out rows that do not contain any of the specified keywords in column D (e.g., "Comfy", "comfy", "comfi", "комфі").
Auto-fits all columns to content.
Selects the final data range (excluding the header row) so that you can easily copy the results.
All formulas remain dynamic. In particular, column O (Percentage) keeps the live structured formula =[@[Attendance Duration]]/[@[Training Duration]] so that if you change any values in the Training Duration column later, the percentages update automatically.

Features

Dynamic Percentage Calculation:
Uses a structured table formula that stays in place and recalculates automatically when Training Duration values change.
Data Cleaning & Consolidation:
Automatically removes unwanted text from the Attendance Duration column and consolidates multiple entries for the same attendee by summing their durations.
Conditional Formatting:
Highlights rows based on the Percentage value (red for below 80%, green for 80% or above).
Keyword Filtering:
Removes rows that do not include any of the specified keywords in a particular column (column D).
Auto-fit and Selection:
Adjusts column widths based on content and selects the final dataset (excluding the header row) for easy manual copying.
Prerequisites

Microsoft Excel for the web with Office Scripts enabled.
A CSV attendance report from Webex containing, at a minimum, the following columns:
Attendee Email
Attendance Duration
(And at least one column in which column D will be scanned for keywords.)
How to Use

Upload Your CSV File:
Open your CSV file in Excel for the web.
Open Office Scripts:
Click on the Automate tab and then New Script.
Copy and Paste the Script:
Copy the contents of the script from this repository and paste it into the Office Scripts editor.
Run the Script:
Save and run the script. The script will process your data as follows:
Create new columns and convert the range into a table.
Clean, consolidate, and sort your attendance data.
Apply conditional formatting and filter out rows that do not contain the specified keywords in column D.
Auto-fit the columns.
Select the final dataset (without headers) for easy manual copying.
Adjust If Needed:
The Percentage column will update automatically if you change any values in the Training Duration column. You can modify the default Training Duration value or update the keywords by editing the script.
Customization

Keywords:
The script filters rows by checking for keywords in column D. Modify the keywords array in the script if you need to change these values.
Training Duration Default Value:
By default, the script fills the "Training Duration" column with 90. You can change this value in the script.
Conditional Formatting Colors:
The red (#FFC7CE) and green (#C6EFCE) fill colors can be updated in the script to your preference.
Contributing

Contributions are welcome! If you find a bug or have ideas for improvements, please open an issue or submit a pull request.

License

This project is licensed under the MIT License.
