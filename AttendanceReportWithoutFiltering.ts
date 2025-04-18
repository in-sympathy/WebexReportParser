// Define a type alias for acceptable cell values.
type CellValue = string | number | boolean;

function main(workbook: ExcelScript.Workbook): void {
  // Get the active worksheet.
  const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

  // Freeze header row.
  sheet.getFreezePanes().freezeRows(1);

  // Get the used range and its values.
  const usedRange: ExcelScript.Range = sheet.getUsedRange();
  const originalValues: CellValue[][] = usedRange.getValues() as CellValue[][];
  if (originalValues.length === 0) {
    throw new Error("No data found.");
  }

  // The first row is assumed to be the header row.
  let headers: string[] = originalValues[0] as string[];
  const originalRowCount: number = originalValues.length;
  const originalColCount: number = headers.length;

  // Validate that required columns exist.
  const emailColumn: string = "Attendee Email";
  const durationColumn: string = "Attendance Duration";
  const emailIdx: number = headers.indexOf(emailColumn);
  const durationIdx: number = headers.indexOf(durationColumn);
  if (emailIdx === -1 || durationIdx === -1) {
    throw new Error(`Required columns "${emailColumn}" or "${durationColumn}" not found.`);
  }

  // Append new header columns: "Training Duration" and "Percentage".
  headers.push("Training Duration", "Percentage");
  const totalCols: number = headers.length;
  
  // Write the updated header row back to the sheet.
  const headerRange: ExcelScript.Range = sheet.getRangeByIndexes(0, 0, 1, totalCols);
  headerRange.setValues([headers]);

  // Get the existing data rows (excluding the header) and add two extra columns.
  let dataRows: CellValue[][] = originalValues.slice(1);
  for (let i = 0; i < dataRows.length; i++) {
    // Ensure every data row has the original number of columns.
    while (dataRows[i].length < originalColCount) {
      dataRows[i].push("");
    }
    // Append placeholders for "Training Duration" and "Percentage".
    dataRows[i].push("");
    dataRows[i].push("");
  }

  // Clean the "Attendance Duration" column by removing " mins" and converting to a number.
  for (let i = 0; i < dataRows.length; i++) {
    const cellVal: CellValue = dataRows[i][durationIdx];
    let durationNum: number = 0;
    if (typeof cellVal === "string") {
      const cleaned: string = cellVal.replace(/ mins/gi, "").trim();
      durationNum = parseFloat(cleaned);
    } else if (typeof cellVal === "number") {
      durationNum = cellVal;
    }
    // Overwrite with the numeric value.
    dataRows[i][durationIdx] = durationNum;
    // Set "Training Duration" (the extra column at index originalColCount) to 90.
    dataRows[i][originalColCount] = 90;
  }

  // Consolidate rows by email; sum the "Attendance Duration" for duplicate emails.
  const grouped: Map<string, CellValue[]> = new Map();
  for (let i = 0; i < dataRows.length; i++) {
    const row: CellValue[] = dataRows[i];
    const email: string = row[emailIdx] as string;
    const durationValue: number = row[durationIdx] as number;
    if (grouped.has(email)) {
      const existingRow: CellValue[] = grouped.get(email)!;
      existingRow[durationIdx] = (existingRow[durationIdx] as number) + durationValue;
    } else {
      grouped.set(email, row.slice());
    }
  }
  const consolidatedRows: CellValue[][] = Array.from(grouped.values());
  
  // Sort the consolidated rows by "Attendance Duration" (ascending).
  consolidatedRows.sort((a, b) => (a[durationIdx] as number) - (b[durationIdx] as number));

  // Clear the original data rows (below the header).
  const currentUsedRows: number = sheet.getUsedRange().getRowCount();
  if (currentUsedRows > 1) {
    const clearRange: ExcelScript.Range = sheet.getRangeByIndexes(
      1,
      0,
      currentUsedRows - 1,
      sheet.getUsedRange().getColumnCount()
    );
    clearRange.clear(ExcelScript.ClearApplyTo.contents);
  }
  // Write the consolidated rows back into the sheet (starting at row 2).
  if (consolidatedRows.length > 0) {
    const outputRange: ExcelScript.Range = sheet.getRangeByIndexes(1, 0, consolidatedRows.length, totalCols);
    outputRange.setValues(consolidatedRows);
  }

  // Helper function to convert a zero-based column index to its corresponding letter.
  function colLetter(n: number): string {
    let letter: string = "";
    while (n >= 0) {
      letter = String.fromCharCode((n % 26) + 65) + letter;
      n = Math.floor(n / 26) - 1;
    }
    return letter;
  }
  
  // Determine column letters.
  const attendanceColLetter: string = colLetter(durationIdx);
  const trainingColLetter: string = colLetter(originalColCount);
  const percentageColIndex: number = totalCols - 1;
  const percentageColLetter: string = colLetter(percentageColIndex);
  
  // Set the "Percentage" formula for each consolidated row.
  for (let i = 0; i < consolidatedRows.length; i++) {
    const excelRow: number = i + 2; // Excel rows are 1-indexed; header is row 1.
    const formula: string = `=${attendanceColLetter}${excelRow}/${trainingColLetter}${excelRow}`;
    const cell: ExcelScript.Range = sheet.getCell(excelRow - 1, percentageColIndex);
    cell.setFormula(formula);
    cell.setNumberFormat("0%");
  }
  
  // (Note: Filtering step has been removed in this version.)

  // Remove unwanted columns by title.
  const columnsToRemove: string[] = ["First Name", "Last Name", "Role", "Attendee Email", "Connection Type"];
  // Retrieve the header row using getRangeByIndexes.
  const headerVals: CellValue[][] = sheet.getRangeByIndexes(0, 0, 1, sheet.getUsedRange().getColumnCount()).getValues();
  if (headerVals.length === 0) {
    throw new Error("Header row is empty.");
  }
  const latestHeader: string[] = headerVals[0] as string[];

  let indicesToRemove: number[] = [];
  columnsToRemove.forEach((colName: string) => {
    const idx: number = latestHeader.indexOf(colName);
    if (idx !== -1) {
      indicesToRemove.push(idx);
    }
  });
  // Delete columns in descending order (to avoid index shifting).
  indicesToRemove.sort((a, b) => b - a);
  indicesToRemove.forEach((idx: number) => {
    const colToDelete: string = colLetter(idx);
    sheet.getRange(`${colToDelete}:${colToDelete}`).delete(ExcelScript.DeleteShiftDirection.left);
  });
  
  // Set time formats for columns B, C, E, and F.
  sheet.getRange("B:B").setNumberFormat("hh:mm:ss");
  sheet.getRange("C:C").setNumberFormat("hh:mm:ss");
  sheet.getRange("E:E").setNumberFormat("hh:mm:ss");
  sheet.getRange("F:F").setNumberFormat("hh:mm:ss");
  
  // Auto-fit all columns.
  sheet.getUsedRange().getFormat().autofitColumns();
  
  // Select the final data range (excluding the header row).
  const finalUsed: ExcelScript.Range = sheet.getUsedRange();
  const finalRowCount: number = finalUsed.getRowCount();
  const finalColCount: number = finalUsed.getColumnCount();
  if (finalRowCount > 1) {
    const finalDataRange: ExcelScript.Range = sheet.getRangeByIndexes(1, 0, finalRowCount - 1, finalColCount);
    finalDataRange.select();
  }
}
