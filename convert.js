const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

// Path to your Excel file
const inputFile = path.join(__dirname, "data.xlsx");
const outputFile = path.join(__dirname, "output.json");

// Read the workbook
const workbook = XLSX.readFile(inputFile);

const result = {};

// Get all sheet
workbook.SheetNames.forEach((sheetName) => {
  if (sheetName == "Setting") {
    return;
  }
  const name = toSlug(sheetName);
  const worksheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(worksheet, { range: 2, defval: null })
    .filter(row => row['Tên Địa Điểm']);

  const mapped = rawData.map(row => ({
    name: row['Tên Địa Điểm'],
    category: row['Phân Loại'] ? row['Phân Loại'].split(',').map(s => s.trim()) : [],
    priceRange: parsePriceRange(row['Khoảng Giá']),
    address: row['Địa Chỉ'],
    lat: row['Latitude'],
    lng: row['Longitude'],
    openHours: row['Thời Gian Mở Cửa'],
    note: row['Ghi Chú'],
    isAllWeek: row['Cả Tuần'],
    isBestChoice: row['Best Choice'],
  }));

  result[name] = mapped;
});

// Write JSON to file
fs.writeFileSync(outputFile, JSON.stringify(result, null, 2), "utf-8");

console.log(
  `✅ Converted ${workbook.SheetNames.length} sheets to ${outputFile}`,
);

function toSlug(str) {
  return str
    .normalize('NFD')                   // separate letters from diacritics
    .replace(/[\u0300-\u036f]/g, '')    // remove diacritic marks
    .replace(/đ/g, 'd').replace(/Đ/g, 'D') // đ/Đ isn't covered by NFD normalization
    .toLowerCase()
    .replace(/\s+/g, '')                // remove all spaces
    .trim();
}

function parsePriceRange(str) {
  if (!str) return { min: null, max: null, raw: str };
  const cleaned = String(str).replace(/đ/gi, '').trim();

  // Case: "1tr+" -> open-ended minimum (e.g. 1,000,000+)
  if (/tr\+?/i.test(cleaned)) {
    const num = parseFloat(cleaned.replace(/[^\d.,]/g, '').replace(',', '.'));
    return { min: num * 1000000, max: null, raw: str };
  }

  // Case: range, e.g. "700.000 – 1.500.000"
  if (cleaned.includes('–') || cleaned.includes('-')) {
    const parts = cleaned.split(/–|-/).map(p => p.trim());
    if (parts.length === 2) {
      const min = Number(parts[0].replace(/\./g, '').replace(/,/g, ''));
      const max = Number(parts[1].replace(/\./g, '').replace(/,/g, ''));
      return { min, max, raw: str };
    }
  }

  // Case: single value, e.g. "100,000 đ" -> means "<= 100,000"
  const single = Number(cleaned.replace(/\./g, '').replace(/,/g, ''));
  if (!isNaN(single)) {
    return { min: null, max: single, raw: str };
  }

  return { min: null, max: null, raw: str };
}