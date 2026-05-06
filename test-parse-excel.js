const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const uploadedPath = path.join(__dirname, 'uploads', 'uploaded.xlsx');

if (fs.existsSync(uploadedPath)) {
  console.log('找到上传的Excel文件');
  
  const wb = XLSX.readFile(uploadedPath);
  console.log('Sheet名称:', wb.SheetNames);
  
  wb.SheetNames.forEach((sheetName, idx) => {
    console.log('\n=== Sheet:', sheetName);
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log('总行数:', data.length);
    console.log('前30行数据:');
    data.slice(0, 30).forEach((row, ri) => {
      console.log(`行${ri}:`, row);
    });
  });
} else {
  console.log('未找到uploaded.xlsx文件');
}
