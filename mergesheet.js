const ExcelJS = require('exceljs');

(async () => {
  try {
    // Load the sheets from the two Excel files
    const headingsWorkbook = new ExcelJS.Workbook();
    await headingsWorkbook.xlsx.readFile('heading.xlsx'); // Change to your headings Excel file

    const dataWorkbook = new ExcelJS.Workbook();
    await dataWorkbook.xlsx.readFile('output.xlsx'); // Change to your data Excel file

    const headingsSheet = headingsWorkbook.getWorksheet(1); // Assuming headings are in the first sheet
    const dataSheet = dataWorkbook.getWorksheet(1); // Assuming data is in the first sheet

    // Create a new workbook for the merged data
    const mergedWorkbook = new ExcelJS.Workbook();
    const mergedSheet = mergedWorkbook.addWorksheet('Merged Data');

    // Copy first two rows from the headings sheet
    for (let rowNum = 1; rowNum <= 2; rowNum++) {
      const headingRow = headingsSheet.getRow(rowNum);
      const headingRowValues = headingRow.values;
      mergedSheet.addRow(headingRowValues);
    }

    // Copy rows 3 to 8 from the data sheet
    for (let rowNum = 3; rowNum <= 8; rowNum++) {
      const dataRow = dataSheet.getRow(rowNum);
      const dataRowValues = dataRow.values;
      mergedSheet.addRow(dataRowValues);
    }

    // Copy row 3 from the headings sheet again
    const headingRow3 = headingsSheet.getRow(3);
    const headingRow3Values = headingRow3.values;
    mergedSheet.addRow(headingRow3Values);

    // Copy row 11 from the data sheet
    const dataRow11 = dataSheet.getRow(11);
    const dataRow11Values = dataRow11.values;
    mergedSheet.addRow(dataRow11Values);

    // Save the merged workbook
    await mergedWorkbook.xlsx.writeFile('merged_output.xlsx');

    console.log('Merged data saved to merged_output.xlsx');
  } catch (error) {
    console.error('Error:', error);
  }
})();
