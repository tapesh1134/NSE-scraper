const puppeteer = require('puppeteer');
const Tesseract = require('tesseract.js');
const fs = require('fs');
const { createCanvas, loadImage } = require('canvas');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

(async () => {
  try {
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1366, height: 768 });

    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.55 Safari/537.36'
    );

    const url = 'https://www.nseindia.com/get-quotes/equity?symbol=TATAMOTORS';
    await page.goto(url);

    // Wait for a short time to allow any dynamic content to load (adjust as needed)
    await page.waitForTimeout(3000);

    // Capture a screenshot of the data
    const screenshotPath = 'screenshot.png';
    await page.screenshot({ path: screenshotPath });

    // Close the browser
    await browser.close();

    // Load the screenshot and create a canvas
    const image = await loadImage(screenshotPath);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(image, 0, 0, image.width, image.height);

    // Define cropping regions for the boxes
    const boxCoordinates = [
      { x: 277, y: 263, width: 855, height: 35 },  // Box 1
      { x: 116, y: 407, width: 257, height: 141 }, // Box 2
      { x: 568, y: 408, width: 98, height: 79 }  // Box 3
    ];


    const extractedData = [];

    for (const coordinates of boxCoordinates) {
      // Crop the canvas to the desired region
      const croppedCanvas = createCanvas(coordinates.width, coordinates.height);
      const croppedCtx = croppedCanvas.getContext('2d');
      croppedCtx.drawImage(canvas, coordinates.x, coordinates.y, coordinates.width, coordinates.height, 0, 0, coordinates.width, coordinates.height);

      // Use Tesseract.js to extract text from the cropped canvas
      const croppedScreenshotPath = `cropped_screenshot_${coordinates.x}_${coordinates.y}.png`;
      const croppedImageData = croppedCanvas.toBuffer('image/png');
      fs.writeFileSync(croppedScreenshotPath, croppedImageData);

      const { data: { text } } = await Tesseract.recognize(croppedScreenshotPath, 'eng', { logger: m => console.log(m) });

      // Log the extracted text for debugging
      console.log(`Extracted Text from (${coordinates.x},${coordinates.y}):\n`, text);

      if (text) {
        // Use new lines when required
        const formattedText = text.replace(/\n+/g, '\n');
        extractedData.push(formattedText);
      } else {
        console.log(`No text extracted from (${coordinates.x},${coordinates.y}).`);
      }
    }

    // Store extracted data in a JSON file
    const jsonPath = 'extracted_data.json';
    const jsonData = JSON.stringify({ extractedData }, null, 2);
    fs.writeFileSync(jsonPath, jsonData, 'utf-8');

    console.log('Extracted data saved to', jsonPath);

    // Store extracted data in an Excel table
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');
    worksheet.columns = [{ header: 'Extracted Data', key: 'data' }];

    extractedData.forEach(text => {
      worksheet.addRow({ data: text });
    });

    const excelPath = 'extracted_data.xlsx';
    await workbook.xlsx.writeFile(excelPath);

    console.log('Extracted data saved to', excelPath);

    // Read JSON data
    const jsonDataXLSX = fs.readFileSync('extracted_data.json', 'utf-8');
    const dataXLSX = JSON.parse(jsonDataXLSX);

    // Create a new workbook and worksheet
    const workbookXLSX = XLSX.utils.book_new();
    const worksheetXLSX = XLSX.utils.aoa_to_sheet([]);

    // Process extractedData
    const extractedDataXLSX = dataXLSX.extractedData;

    // Insert data from extractedData into worksheet
    const rowsXLSX = extractedDataXLSX[0].split(/\s+/).filter(Boolean);
    rowsXLSX.unshift();
    XLSX.utils.sheet_add_aoa(worksheetXLSX, [rowsXLSX], { origin: 'A11' });

    const secondDataRowsXLSX = extractedDataXLSX[1].split('\n').map(row => row.trim().split(/\s+/));
    const secondDataProcessedXLSX = secondDataRowsXLSX.map(row => row.map(cell => (cell === '-' || cell === '') ? '-' : cell));
    XLSX.utils.sheet_add_aoa(worksheetXLSX, secondDataProcessedXLSX, { origin: { r: 2, c: 1 } });

    const thirdDataXLSX = extractedDataXLSX[2].split(/\s+/).filter(Boolean);
    thirdDataXLSX.unshift();
    XLSX.utils.sheet_add_aoa(worksheetXLSX, [thirdDataXLSX], { origin: 'H11' });

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookXLSX, worksheetXLSX, 'Sheet1');

    // Write the workbook to a file
    XLSX.writeFile(workbookXLSX, 'output.xlsx');

  } catch (error) {
    console.error('Error:', error);
  }
})();
