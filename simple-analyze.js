
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const uploadedPath = path.join(__dirname, 'uploads', 'uploaded.xlsx');

if (fs.existsSync(uploadedPath)) {
  console.log('='.repeat(80));
  console.log('分析回字形布局Excel');
  console.log('='.repeat(80));
  
  const wb = XLSX.readFile(uploadedPath);
  
  wb.SheetNames.forEach(function(sheetName, sheetIdx) {
    console.log('\n【Sheet】', sheetName);
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    console.log('总行数:', data.length);
    
    // 打印所有有内容的行
    for (var i = 0; i &lt; data.length; i++) {
      var row = data[i];
      var hasContent = false;
      var rowStr = '';
      for (var j = 0; j &lt; row.length; j++) {
        if (row[j]) {
          hasContent = true;
          rowStr += ' [' + j + ']' + row[j];
        }
      }
      if (hasContent) {
        console.log('行', i, ':', rowStr);
      }
    }
  });
}
