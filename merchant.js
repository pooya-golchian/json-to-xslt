const fs = require('fs');
const XLSX = require('xlsx');

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

function processJSONFiles() {
  fs.readdirSync('./merchantJsonFiles').forEach((file) => {
    if (file.endsWith('.json')) {
      const jsonData = JSON.parse(
        fs.readFileSync(`./merchantJsonFiles/${file}`, 'utf-8')
      );
      const xlsxBuffer = convertJSONtoXLSX(jsonData);
      fs.writeFileSync(
        `./merchantXlsxFiles/${file.replace('.json', '.xlsx')}`,
        xlsxBuffer
      );
      console.log(`Converted ${file} to XLSX.`);
    }
  });
}

if (!fs.existsSync('./merchantXlsxFiles')) {
  fs.mkdirSync('./merchantXlsxFiles');
}

processJSONFiles();
