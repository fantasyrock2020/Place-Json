const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

// Path to your Excel file
const inputFile = path.join(__dirname, "BẢN ĐỒ ĂN CHƠI.xlsx");
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

  // Helper: convert 1 "token" như "100k", "1.5tr", "700.000" thành number
  function toNumber(token) {
    token = token.trim();
    if (/tr\+?$/i.test(token)) {
      const num = parseFloat(token.replace(/[^\d.,]/g, '').replace(',', '.'));
      return num * 1_000_000;
    }
    if (/k$/i.test(token)) {
      const num = parseFloat(token.replace(/[^\d.,]/g, '').replace(',', '.'));
      return num * 1_000;
    }
    // số thường, dấu . hoặc , là phân cách nghìn
    return Number(token.replace(/\./g, '').replace(/,/g, ''));
  }

  // Case: mở, ví dụ "1tr+"
  if (/tr\+$/i.test(cleaned) && !cleaned.includes('-') && !cleaned.includes('–')) {
    return { min: toNumber(cleaned), max: null, raw: str };
  }

  // Case: range, ví dụ "100k - 200k", "700.000 – 1.500.000"
  if (cleaned.includes('–') || cleaned.includes('-')) {
    const parts = cleaned.split(/–|-/).map(p => p.trim());
    if (parts.length === 2) {
      const min = toNumber(parts[0]);
      const max = toNumber(parts[1]);
      return { min, max, raw: str };
    }
  }

  // Case: giá trị đơn
  const single = toNumber(cleaned);
  if (!isNaN(single)) {
    return { min: null, max: single, raw: str };
  }

  return { min: null, max: null, raw: str };
}