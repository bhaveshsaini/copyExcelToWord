const XLSX = require('xlsx');
const PizZip = require("pizzip");
const Docxtemplater = require('docxtemplater');
const fs = require('fs')
const path = require('path');

function copyExcelToWord() {
  
  // Load the Excel file
  const workbook = XLSX.readFile("excelfile.xlsx");
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  // Extract the data from the worksheet
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  //writing to word file

  data.map(element => {
    const content = fs.readFileSync('template.docx', 'binary');
    const zip = new PizZip(content);
    let doc = new Docxtemplater(zip)

    doc.setData({
      Serial: element[0],
      item: element[1],
      Qty: element[2],
      Price: '$' + element[3],
      Netsales: 'Net Sales'
    });
  
    // Perform the template rendering
    doc.render();
    // Get the rendered document as a buffer
    let outputBuffer = doc.getZip().generate({ type: 'nodebuffer' })
    fs.writeFileSync(path.join(__dirname, 'output.docx'), outputBuffer);
  })
}

// Usage example
copyExcelToWord();


