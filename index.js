const fs = require("fs");
const XLSX = require("xlsx");
const jsonData = fs.readFileSync("data.json", "utf-8");
const data = JSON.parse(jsonData);
const keyValuePairs = [];
for (const key in data) {
  if (data.hasOwnProperty(key)) {
    const value = data[key];
    const keyValueString = `${key}: ${value}`;
    keyValuePairs.push(keyValueString);
  }
}
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(keyValuePairs.map((pair) => [pair]));
XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
XLSX.writeFile(wb, "output.xlsx");
