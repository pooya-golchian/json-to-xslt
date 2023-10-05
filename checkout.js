const fs = require('fs');
const XLSX = require('xlsx');

// Function to convert JSON to XLSX with separate rows for each key-value pair
function convertJSONtoXLSX(jsonData) {
  const wsData = [];
  for (const item of jsonData) {
    for (const key of Object.keys(item)) {
      const row = {
        String: key,
        Value: item[key],
      };
      wsData.push(row);
    }
  }

  const ws = XLSX.utils.json_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
  const xlsxBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return xlsxBuffer;
}

// Read JSON files and convert to XLSX
function processJSONFiles() {
  fs.readdirSync('./checkoutJsonFiles').forEach((file) => {
    if (file.endsWith('.json')) {
      const jsonData = JSON.parse(
        fs.readFileSync(`./checkoutJsonFiles/${file}`, 'utf-8')
      );
      const xlsxBuffer = convertJSONtoXLSX(jsonData);
      fs.writeFileSync(
        `./checkoutXlsxFiles/${file.replace('.json', '.xlsx')}`,
        xlsxBuffer
      );
      console.log(`Converted ${file} to XLSX.`);
    }
  });
}

// Create output directory if it doesn't exist
if (!fs.existsSync('./checkoutXlsxFiles')) {
  fs.mkdirSync('./checkoutXlsxFiles');
}

// Start the conversion process
processJSONFiles();
