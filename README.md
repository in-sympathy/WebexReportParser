# Attendance Report Processor

This repository includes two Office Script variants for processing attendance data in Excel. Both scripts consolidate data by "Attendee Email", calculate a "Percentage" of training attended (with a default training duration of 90 minutes), remove unwanted columns, and format the output.

## Variants

- **With Filtering:**  
  Filters rows based on keywords (e.g., "comfy", "комфі") found in the "Display Name" column.

- **Without Filtering:**  
  Performs all data consolidation and formatting without any filtering by keywords.

## Usage

1. **Open Excel on the Web:** Go to the **Automate** tab and create a new script.
2. **Copy & Paste:** Choose the desired variant (`AttendanceReportWithFiltering.ts` or `AttendanceReportWithoutFiltering.ts`) from this repository and paste it into the script editor.
3. **Run the Script:** Ensure your worksheet includes at least the columns "Attendee Email" and "Attendance Duration" (and "Display Name" for the filtering variant), then run the script.

## Customization

- **Training Duration:** Default is 90 minutes – change as needed.
- **Columns to Remove:** Modify the `columnsToRemove` array in the script.
