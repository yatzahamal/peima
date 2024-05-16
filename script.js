const XLSX = require('xlsx');
const fs = require('fs');

const ID = "id";
const NAME = "name";
const LICENSE = "license";
const TYPE = "type";
const RELEASE_DATE = "release date";
const FROM_DATE = "from date";
const UNTIL_DATE = "until date";
const SUM = "sum";

const excelDateToString = (excelDate) => {
  if (excelDate === undefined) return undefined;
  // Convert Excel date number to JavaScript timestamp
  let timestamp = (excelDate - 25569) * 86400 * 1000;
  
  // Create a Date object from the timestamp
  let date = new Date(timestamp);
  
  // Format the date as a string
  let dateString = date.toLocaleDateString();
  
  return dateString;
}

// Load Excel file
const workbook = XLSX.readFile('input.xlsx');

// Get the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert Excel data to JSON
const data = XLSX.utils.sheet_to_json(worksheet);

// Group data by the values in the first column
const groupedData = data.reduce((acc, row) => {
    const key = row[Object.keys(row)[0]]; // Assuming the first column is the key
    if (!acc[key]) {
        acc[key] = [];
    }
    acc[key].push(row);
    return acc;
}, {});


const toolMsg = (toolJson) => {
  const license = toolJson[LICENSE];
  const type = toolJson[TYPE];
  const releaseDate = excelDateToString(toolJson[RELEASE_DATE]);
  const fromDate = excelDateToString(toolJson[FROM_DATE]);
  const untilDate = excelDateToString(toolJson[UNTIL_DATE]);
  const sum = toolJson[SUM].toLocaleString();

  let row1 = `חשבונית מס קבלה עבור כלי מספר רישוי: ${license} - ${type}\n`;
  let row2 = `${releaseDate === undefined ? "" : `תאריך שחרור: ${releaseDate}\n`}`;
  let row3 = `תשלום עבור תאריכים: ${fromDate} - ${untilDate} (נדרש לציין על גבי החשבונית)\n`;
  let row4 = `סכום לתשלום לא כולל מע"מ: ${sum} ש"ח`;
  const msg = row1 + row2 + row3 + row4;
  return msg;
}

const generateOwnerMsg = (id) => {
  const json = groupedData[id];
  const row1 = "####################\n";
  const row2 = `בעלים: ${json[0][NAME]}\n`;
  const row3 = "####################\n\n";
  const row4 = "שלום רב, אנו מקדמים תשלום נוסף עבור הכלים המגויסים. לצורך כך אודה לקבלת חשבונית בקובץ PDF עבור הכלים על פי הפרטים הבאים:\n\n";
  let msg = row1 + row2 + row3 + row4;
  for (let i = 0; i < Object.keys(json).length; i++) {
    const tool = json[i];
    msg += `${i + 1}. ${toolMsg(tool)}\n\n`;
  }
  msg += "\n\n";
  return msg;
}

const generateAll = () => {
  let msg = "";
  for (let i = 0; i < Object.keys(groupedData).length; i++) {
    const id = Object.keys(groupedData)[i];
    msg += generateOwnerMsg(id);
  }
  return msg;
}

fs.writeFileSync('output.txt', generateAll());

