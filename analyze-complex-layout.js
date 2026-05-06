
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const uploadedPath = path.join(__dirname, 'uploads', 'uploaded.xlsx');

if (fs.existsSync(uploadedPath)) {
  console.log('='.repeat(80));
  console.log('详细分析回字形布局Excel文件');
  console.log('='.repeat(80));
  
  const wb = XLSX.readFile(uploadedPath);
  wb.SheetNames.forEach((sheetName, sheetIdx) =&gt; {
    console.log(`\n【Sheet ${sheetIdx + 1}】${sheetName}`);
    console.log('-'.repeat(80));
    
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    console.log(`总行数: ${data.length}`);
    
    // 统计每一行的有效单元格
    data.forEach((row, rowIdx) =&gt; {
      const validCells = row.filter(cell =&gt; cell !== null &amp;&amp; cell !== undefined &amp;&amp; cell !== '');
      if (validCells.length &gt; 0) {
        console.log(`行 ${rowIdx.toString().padStart(2, ' ')} (${validCells.length}个):`, row.map((cell, colIdx) =&gt; {
          if (cell === null || cell === undefined || cell === '') return '';
          return `[${colIdx.toString().padStart(2, ' ')}]${String(cell).padEnd(12, ' ')}`;
        }).filter(x =&gt; x).join(' '));
      }
    });
    
    console.log('\n结构化分析...');
    
    // 分析布局结构
    const rowLabels = []; // 排/列标签
    const seatNumbers = []; // 座位号
    const allCells = []; // 所有有效单元格位置
    
    data.forEach((row, rowIdx) =&gt; {
      row.forEach((cell, colIdx) =&gt; {
        if (cell &amp;&amp; String(cell).trim()) {
          const cellValue = String(cell).trim();
          allCells.push({
            row: rowIdx,
            col: colIdx,
            value: cellValue,
            isLabel: /^(外|内)(左|右|前|后).*(排|列)$/.test(cellValue) || /^(第\d+排|沙发第\d+排|\d+排|\d+列)$/.test(cellValue),
            isSeat: /^\d+$/.test(cellValue),
            isStage: /^(舞台|主席台)$/.test(cellValue)
          });
        }
      });
    });
    
    console.log(`\n找到 ${allCells.length} 个有效单元格`);
    
    // 分类
    const labelCells = allCells.filter(c =&gt; c.isLabel);
    const seatCells = allCells.filter(c =&gt; c.isSeat);
    
    console.log(`\n- 排/列标签 (${labelCells.length}个):`);
    labelCells.forEach(c =&gt; console.log(`  [${c.row},${c.col}] ${c.value}`));
    
    console.log(`\n- 座位号 (${seatCells.length}个):`);
    // 按区域分组显示座位
    const leftSeats = seatCells.filter(c =&gt; c.col &lt; 10);
    const midSeats = seatCells.filter(c =&gt; c.col &gt;= 10 &amp;&amp; c.col &lt; 20);
    const rightSeats = seatCells.filter(c =&gt; c.col &gt;= 20);
    if (leftSeats.length) console.log('  左区座位:', leftSeats.map(c =&gt; c.value).join(','));
    if (midSeats.length) console.log('  中区座位:', midSeats.map(c =&gt; c.value).join(','));
    if (rightSeats.length) console.log('  右区座位:', rightSeats.map(c =&gt; c.value).join(','));
    
    // 尝试构建回字形结构
    console.log('\n'.repeat(2));
    console.log('='.repeat(80));
    console.log('构建回字形会场布局配置');
    console.log('='.repeat(80));
    
    // 根据之前的分析，这个会场包含：
    // 外左1-23排（竖），外右1-26排（竖），内前1排（横），内左1-17排（竖），内右1-16排（竖），内后1排（横）
    
    console.log('\n创建自定义布局配置：');
    
    // 让我们构建一个回字形布局的坐标配置
    const canvasW = 1200;
    const canvasH = 1000;
    const seatW = 40;
    const seatH = 30;
    const gap = 10;
    
    const customRows = [];
    let rowNum = 1;
    
    // ===== 外左部分 (竖排) =====
    console.log('\n// 外左 (竖排)');
    for (let i = 1; i &lt;= 23; i++) {
      customRows.push({
        label: `外左第${i}排`,
        rowNum: rowNum++,
        x: 50,
        y: 100 + (i - 1) * (seatH + gap),
        width: seatW,
        height: seatH,
        seatCount: 1,
        direction: 'vertical'
      });
      console.log(`{ label: "外左第${i}排", rowNum: ${customRows.length}, x: 50, y: ${100 + (i - 1) * (seatH + gap)}, width: ${seatW}, height: ${seatH}, seatCount: 1, direction: "vertical" },`);
    }
    
    // ===== 外右部分 (竖排) =====
    console.log('\n// 外右 (竖排)');
    for (let i = 1; i &lt;= 26; i++) {
      customRows.push({
        label: `外右第${i}排`,
        rowNum: rowNum++,
        x: canvasW - 50 - seatW,
        y: 100 + (i - 1) * (seatH + gap),
        width: seatW,
        height: seatH,
        seatCount: 1,
        direction: 'vertical'
      });
      console.log(`{ label: "外右第${i}排", rowNum: ${customRows.length}, x: ${canvasW - 50 - seatW}, y: ${100 + (i - 1) * (seatH + gap)}, width: ${seatW}, height: ${seatH}, seatCount: 1, direction: "vertical" },`);
    }
    
    // ===== 内前部分 (横排) =====
    console.log('\n// 内前 (横排)');
    customRows.push({
      label: '内前第一排',
      rowNum: rowNum++,
      x: 200,
      y: 100,
      width: 800,
      height: seatH,
      seatCount: 16,
      direction: 'horizontal'
    });
    console.log(`{ label: "内前第一排", rowNum: ${customRows.length}, x: 200, y: 100, width: 800, height: ${seatH}, seatCount: 16, direction: "horizontal" },`);
    
    // ===== 内左部分 (竖排) =====
    console.log('\n// 内左 (竖排)');
    for (let i = 1; i &lt;= 17; i++) {
      customRows.push({
        label: `内左第${i}排`,
        rowNum: rowNum++,
        x: 200,
        y: 150 + (i - 1) * (seatH + gap),
        width: seatW,
        height: seatH,
        seatCount: 1,
        direction: 'vertical'
      });
      console.log(`{ label: "内左第${i}排", rowNum: ${customRows.length}, x: 200, y: ${150 + (i - 1) * (seatH + gap)}, width: ${seatW}, height: ${seatH}, seatCount: 1, direction: "vertical" },`);
    }
    
    // ===== 内右部分 (竖排) =====
    console.log('\n// 内右 (竖排)');
    for (let i = 1; i &lt;= 16; i++) {
      customRows.push({
        label: `内右第${i}排`,
        rowNum: rowNum++,
        x: canvasW - 200 - seatW,
        y: 150 + (i - 1) * (seatH + gap),
        width: seatW,
        height: seatH,
        seatCount: 1,
        direction: 'vertical'
      });
      console.log(`{ label: "内右第${i}排", rowNum: ${customRows.length}, x: ${canvasW - 200 - seatW}, y: ${150 + (i - 1) * (seatH + gap)}, width: ${seatW}, height: ${seatH}, seatCount: 1, direction: "vertical" },`);
    }
    
    // ===== 内后部分 (横排) =====
    console.log('\n// 内后 (横排)');
    const innerBottomY = 150 + Math.max(17, 16) * (seatH + gap) + 50;
    customRows.push({
      label: '内后第一排',
      rowNum: rowNum++,
      x: 200,
      y: innerBottomY,
      width: 800,
      height: seatH,
      seatCount: 16,
      direction: 'horizontal'
    });
    console.log(`{ label: "内后第一排", rowNum: ${customRows.length}, x: 200, y: ${innerBottomY}, width: 800, height: ${seatH}, seatCount: 16, direction: "horizontal" },`);
    
    console.log('\n'.repeat(2));
    console.log('='.repeat(80));
    console.log(`总计 ${customRows.length} 排/列`);
    console.log('='.repeat(80));
    
    // 保存配置文件
    const configPath = path.join(__dirname, 'uploads', 'complex-layout-config.json');
    fs.writeFileSync(configPath, JSON.stringify({
      venueName: '回字形会场',
      canvasWidth: canvasW,
      canvasHeight: canvasH,
      customRows: customRows
    }, null, 2));
    console.log(`\n配置文件已保存到: ${configPath}`);
  });
} else {
  console.log('未找到 uploaded.xlsx 文件');
}
