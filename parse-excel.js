const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ============================================================
// 灵活的中英文数字转换
// ============================================================
const cnNums = { '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,
  '十一':11,'十二':12,'十三':13,'十四':14,'十五':15,'十六':16,'十七':17,'十八':18,'十九':19,'二十':20,'二十一':21,
  '1':1,'2':2,'3':3,'4':4,'5':5,'6':6,'7':7,'8':8,'9':9,'10':10,'11':11,'12':12,'13':13,'14':14,'15':15,'16':16,'17':17,'18':18,'19':19,'20':20,'21':21 };

const cnToNum = {
  '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
  '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
  '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
  '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20,
  '二十一': 21, '二十二': 22, '二十三': 23, '二十四': 24,
  '二十五': 25, '二十六': 27, '二十八': 28, '二十九': 29, '三十': 30
};

function parseChineseNumber(str) {
  if (!str) return null;
  if (/^\d+$/.test(str)) return parseInt(str);
  if (cnToNum[str] !== undefined) return cnToNum[str];
  return null;
}

// ============================================================
// 灵活的座位号识别
// ============================================================

// 检测值是否为座位号（数字或可能的座位标识）
function isSeatNumber(value) {
  if (typeof value === 'number') return true;
  if (typeof value !== 'string') return false;
  const trimmed = value.trim();
  if (!trimmed) return false;
  
  // 纯数字
  if (/^\d+$/.test(trimmed)) return true;
  // 数字+号 (1号, 2号)
  if (/^\d+号$/.test(trimmed)) return true;
  // 字母+数字 (A1, B2, AA10) - 但排除 R1/R2 等常见排标签
  if (/^[A-Za-z]+\d+$/.test(trimmed)) {
    // 排除 R1, R2 等常见排标签
    if (/^R\d+$/i.test(trimmed)) return false;
    // 排除 F1, F2, Floor1 等楼层标签
    if (/^F\d+$/i.test(trimmed)) return false;
    if (/^Floor\d+$/i.test(trimmed)) return false;
    return true;
  }
  // 字母-数字 (A-1, B-2)
  if (/^[A-Za-z]+-\d+$/.test(trimmed)) return true;
  
  return false;
}

// 提取座位号的数字部分（用于排序）
function extractSeatNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return 0;
  const trimmed = value.trim();
  
  const match = trimmed.match(/(\d+)/);
  return match ? parseInt(match[1]) : 0;
}

// 将座位号规范化为统一格式
function normalizeSeatNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return value;
  const trimmed = value.trim();
  
  if (/^\d+$/.test(trimmed)) return parseInt(trimmed);
  if (/^\d+号$/.test(trimmed)) return parseInt(trimmed.replace('号', ''));
  
  return trimmed;
}

// ============================================================
// 灵活的排/列标签识别
// ============================================================

// 匹配各种排标签格式
const ROW_PATTERNS = [
  /^第(.+)[排]$/,              // 第X排
  /^第(.+)[列]$/,              // 第X列
  /^(?:Row|ROW|row)\s*(\d+)$/i, // Row 1, ROW 1
  /^R(\d+)$/i,                 // R1, R2
  /^(\d+)[排]$/,              // 1排
  /^(\d+)[列]$/,              // 1列
  // 组合格式：一楼第X排、二楼第X排
  /^(.+)[楼]第(.+)[排]$/,     // 一楼第一排、二楼第1排
  /^(.+)[楼](\d+)[排]$/,      // 一楼1排、二楼2排
];

function parseRowLabel(text) {
  if (!text || typeof text !== 'string') return null;
  const trimmed = text.trim();
  
  for (const pattern of ROW_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      // 组合格式：一楼第X排
      if (pattern.source.includes('[楼]第')) {
        return parseChineseNumber(match[2]);
      }
      if (pattern.source.includes('[楼]')) {
        return parseChineseNumber(match[2]);
      }
      const numPart = match[1];
      return parseChineseNumber(numPart);
    }
  }
  
  return null;
}

// 从组合格式标签中提取楼层信息（如"一楼第1排" -> "一楼"）
function extractFloorFromCombinedLabel(text) {
  if (!text || typeof text !== 'string') return null;
  const trimmed = text.trim();
  
  // 匹配 一楼第X排、二楼第X排 格式
  const match = trimmed.match(/^(.+)[楼]第/);
  if (match) return match[1] + '楼';
  
  const match2 = trimmed.match(/^(.+)[楼]\d+[排]/);
  if (match2) return match2[1] + '楼';
  
  return null;
}

// ============================================================
// 楼层标签识别（剧院布局）
// ============================================================

// 匹配各种楼层标签格式
const FLOOR_PATTERNS = [
  /^(.+)[楼]$/,                // 一楼、二楼、三楼
  /^第(.+)[楼]$/,             // 第一楼、第二楼
  /^(\d+)F$/i,                // 1F、2F、3F
  // /^F(\d+)$/i,                // F1、F2、F3  -- 注释掉，避免与 R1/R2 冲突
  /^(?:Floor|floor|FLOOR)\s*(\d+)$/i, // Floor 1, Floor 2
  /^(.+)[层]$/,               // 一层、二层、三层
  /^第(.+)[层]$/,             // 第一层、第二层
];

// 常见楼层关键字
const FLOOR_KEYWORDS = ['一楼', '二楼', '三楼', '四楼', '五楼', '六楼', '七楼', '八楼', '九楼', '十楼',
                        '一层', '二层', '三层', '四层', '五层', '六层', '七层', '八层', '九层', '十层',
                        '一楼前排', '一楼后排', '二楼前排', '二楼后排', '二楼左侧', '二楼右侧'];

function isFloorLabel(text) {
  if (!text || typeof text !== 'string') return false;
  const trimmed = text.trim();
  
  for (const pattern of FLOOR_PATTERNS) {
    if (pattern.test(trimmed)) return true;
  }
  
  for (const keyword of FLOOR_KEYWORDS) {
    if (trimmed === keyword) return true;
  }
  
  return false;
}

function parseFloorLabel(text) {
  if (!text || typeof text !== 'string') return null;
  const trimmed = text.trim();
  
  for (const pattern of FLOOR_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      const numPart = match[1];
      const num = parseChineseNumber(numPart);
      if (num !== null) {
        return { floorNum: num, label: trimmed };
      }
    }
  }
  
  // 只匹配精确关键字，不匹配"一楼第X排"等组合格式
  for (const keyword of FLOOR_KEYWORDS) {
    if (trimmed === keyword) {
      const floorNum = (FLOOR_KEYWORDS.indexOf(keyword) % 10) + 1;
      return { floorNum, label: keyword };
    }
  }
  
  return null;
}

// 判断是否为布局关键字（不是人名或座位号）
const KEYWORDS = /^(舞台|舞台区域|投影仪|演讲台|中心线|沙发|过道|横向过道|纵向过道)$/;
const LAYOUT_PHRASES = /^(课桌式|宴会厅|共计|现报名|需要增加|一排.+人|共.+排)$/;

function isLayoutKeyword(s) {
  if (s === undefined || s === null) return true;
  if (typeof s !== 'string') return false;
  s = s.trim();
  if (s.length === 0) return true;
  if (KEYWORDS.test(s)) return true;
  if (LAYOUT_PHRASES.test(s)) return true;
  if (parseRowLabel(s) !== null) return true;
  return false;
}

// 检测是否为沙发排
function isSofaRow(text) {
  if (!text || typeof text !== 'string') return false;
  const trimmed = text.trim();
  return /^沙发/.test(trimmed);
}

function parseSofaLabel(text) {
  if (!text || typeof text !== 'string') return null;
  const trimmed = text.trim();
  
  const match = trimmed.match(/^沙发\s*第(.+)[排]$/);
  if (match) {
    return { rowNum: parseChineseNumber(match[1]), label: trimmed };
  }
  
  const match2 = trimmed.match(/^沙发\s*(?:Row|ROW|row)\s*(\d+)$/i);
  if (match2) {
    return { rowNum: parseInt(match2[1]), label: trimmed };
  }
  
  return null;
}

// 按列间隔分组（间隔>1列视为过道）
function groupByGap(positions) {
  if (positions.length === 0) return [];
  positions.sort((a, b) => a.col - b.col);
  const groups = [[positions[0]]];
  for (let i = 1; i < positions.length; i++) {
    if (positions[i].col - positions[i - 1].col > 1) {
      groups.push([]);
    }
    groups[groups.length - 1].push(positions[i]);
  }
  return groups;
}

// 自动补全缺失的座位号（从1到最大座位号的连续序列）
function fillMissingSeatNumbers(seatNums) {
  if (seatNums.length === 0) return seatNums;
  
  const maxNum = Math.max(...seatNums);
  const existingSet = new Set(seatNums);
  const filled = [];
  
  for (let i = 1; i <= maxNum; i++) {
    filled.push(i);
  }
  
  return filled;
}

/**
 * 智能布局检测：根据表格结构自动推断布局类型
 * 支持更多样的表格格式
 */
function detectLayoutType(data, venueId) {
  const rows = [];
  const attendees = [];
  
  // 方法1：检测剧院楼层布局（一楼/二楼/三楼 + 排标签 + 座位号）
  // 适用于剧院、礼堂等多楼层场馆
  let currentFloor = null;
  let floorRowCounter = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    const floorLabel = row.find(c => isFloorLabel(c));
    if (floorLabel) {
      const floorInfo = parseFloorLabel(floorLabel);
      if (floorInfo) {
        currentFloor = floorInfo.label;
        if (!floorRowCounter[currentFloor]) floorRowCounter[currentFloor] = 0;
      }
    }
    
    const rowLabel = row.find(c => parseRowLabel(c) !== null);
    if (rowLabel && !isFloorLabel(rowLabel)) {
      const rn = parseRowLabel(rowLabel);
      if (rn !== null) {
        floorRowCounter[currentFloor || 'default'] = (floorRowCounter[currentFloor || 'default'] || 0) + 1;
        const rowNum = floorRowCounter[currentFloor || 'default'];
        
        const seatPositions = [];
        row.forEach((c, ci) => {
          if (isSeatNumber(c)) {
            seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
          }
        });
        
        if (seatPositions.length > 0) {
          const groups = groupByGap(seatPositions);
          const fullLabel = currentFloor ? `${currentFloor}${rowLabel}` : rowLabel;
          rows.push({
            label: fullLabel,
            rowNum: rn,
            floor: currentFloor,
            seatGroups: groups.map(g => g.map(s => s.num))
          });
          
          row.forEach((cell, ci) => {
            if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell) && !isFloorLabel(cell)) {
              const seatCol = seatPositions.find(sp => sp.col === ci);
              if (seatCol) {
                attendees.push({
                  name: cell.trim(),
                  row: fullLabel,
                  seat: seatCol.num,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            }
          });
        }
      }
    } else if (!rowLabel) {
      const seatPositions = [];
      row.forEach((c, ci) => {
        if (isSeatNumber(c)) {
          seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
        }
      });
      
      if (seatPositions.length >= 3 && !row.some(c => isFloorLabel(c))) {
        const isNameRow = row.some(c => typeof c === 'string' && !isLayoutKeyword(c) && !isSeatNumber(c) && !isFloorLabel(c));
        if (isNameRow) {
          floorRowCounter[currentFloor || 'default'] = (floorRowCounter[currentFloor || 'default'] || 0) + 1;
          const rowNum = floorRowCounter[currentFloor || 'default'];
          const groups = groupByGap(seatPositions);
          const fullLabel = currentFloor ? `${currentFloor}第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排` : `第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排`;
          
          rows.push({
            label: fullLabel,
            rowNum: rowNum,
            floor: currentFloor,
            seatGroups: groups.map(g => g.map(s => s.num))
          });
          
          row.forEach((cell, ci) => {
            if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell) && !isFloorLabel(cell)) {
              const seatCol = seatPositions.find(sp => sp.col === ci);
              if (seatCol) {
                attendees.push({
                  name: cell.trim(),
                  row: fullLabel,
                  seat: seatCol.num,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            }
          });
        }
      }
    }
  }
  
  if (rows.length > 0) {
    rows.sort((a, b) => {
      if (a.floor && b.floor) {
        const aFloorNum = parseFloorLabel(a.floor)?.floorNum || 0;
        const bFloorNum = parseFloorLabel(b.floor)?.floorNum || 0;
        if (aFloorNum !== bFloorNum) return aFloorNum - bFloorNum;
      }
      return (a.rowNum || 0) - (b.rowNum || 0);
    });
    return { rows, layout: 'theater', attendees };
  }
  
  // 方法2：检测标准矩阵布局（每行有排标签，每列有座位号）
  // 适用于：Row 1/Row 2, R1/R2, 第1排/第2排 等格式
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowLabel = row.find(c => parseRowLabel(c) !== null);
    if (rowLabel) {
      const rn = parseRowLabel(rowLabel);
      if (rn !== null) {
        const seatPositions = [];
        row.forEach((c, ci) => {
          if (isSeatNumber(c)) {
            seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
          }
        });
        
        if (seatPositions.length > 0) {
          const groups = groupByGap(seatPositions);
          rows.push({
            label: rowLabel,
            rowNum: rn,
            seatGroups: groups.map(g => g.map(s => s.num))
          });
          
          // 提取人名（非座位号、非布局关键字的字符串）
          row.forEach((cell, ci) => {
            if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell)) {
              const seatCol = seatPositions.find(sp => sp.col === ci);
              if (seatCol) {
                attendees.push({
                  name: cell.trim(),
                  row: rowLabel,
                  seat: seatCol.num,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            }
          });
        }
      }
    }
  }
  
  if (rows.length > 0) {
    rows.sort((a, b) => (a.rowNum || 0) - (b.rowNum || 0));
    return { rows, layout: 'standard', attendees };
  }
  
  // 方法2：检测只有座位号的布局（无明确排标签）
  // 查找包含座位号的行，按行顺序分配排号
  let seatNumberRows = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const seatPositions = [];
    row.forEach((c, ci) => {
      if (isSeatNumber(c)) {
        seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
      }
    });
    
    if (seatPositions.length >= 3) {
      seatNumberRows.push({
        excelRow: i,
        seatPositions: seatPositions,
        groups: groupByGap(seatPositions)
      });
    }
  }
  
  if (seatNumberRows.length > 0) {
    seatNumberRows.forEach((sr, idx) => {
      const rowNum = idx + 1;
      const label = `第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排`;
      rows.push({
        label: label,
        rowNum: rowNum,
        seatGroups: sr.groups.map(g => g.map(s => s.num))
      });
    });
    
    return { rows, layout: 'seat-numbers-only' };
  }
  
  // 方法3：检测人员姓名矩阵布局（按列间距分组）
  // 适用于：只有人名，没有座位号的表格
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const personNamesInRow = [];
    row.forEach((cell, ci) => {
      if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell)) {
        personNamesInRow.push({ col: ci, name: cell.trim() });
      }
    });
    
    if (personNamesInRow.length >= 2) {
      // 按列间距分组
      personNamesInRow.sort((a, b) => a.col - b.col);
      const positionGroups = [[personNamesInRow[0]]];
      for (let j = 1; j < personNamesInRow.length; j++) {
        const prevCol = personNamesInRow[j - 1].col;
        const curCol = personNamesInRow[j].col;
        if (curCol - prevCol > 1) {
          positionGroups.push([]);
        }
        positionGroups[positionGroups.length - 1].push(personNamesInRow[j]);
      }
      
      let seatNum = 1;
      const seatGroups = [];
      positionGroups.forEach(group => {
        const groupSeats = [];
        group.forEach((person, personIdx) => {
          attendees.push({
            name: person.name,
            row: `第${Object.keys(cnNums).find(k => cnNums[k] === i + 1) || (i + 1)}排`,
            seat: seatNum + personIdx,
            company: '',
            title: '',
            venueId: venueId,
            source: 'excel'
          });
          groupSeats.push(seatNum + personIdx);
        });
        seatGroups.push(groupSeats);
        seatNum += group.length;
      });
      
      const rowNum = i + 1;
      rows.push({
        label: `第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排`,
        rowNum: rowNum,
        seatGroups: seatGroups
      });
    }
  }
  
  if (rows.length > 0) {
    return { rows, layout: 'name-matrix', attendees };
  }
  
  return null;
}

/**
 * 自动检测Sheet的布局类型
 * @param {Array} data - 二维数组，sheet_to_json 的输出
 * @returns {string} 'u-shape' | 'theater' | 'standard'
 */
function detectSheetMode(data) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const colLabels = row.filter(c => typeof c === 'string' && /^第.+列$/.test(c.trim()));
    if (colLabels.length >= 2) {
      // 检查标题行上方是否有连续数字 → 回字型
      if (i > 0) {
        const numCount = data[i - 1].filter(c => typeof c === 'number').length;
        if (numCount >= 3) return 'hui-shape';
      }
      return 'u-shape';
    }
  }
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const hasFloor = row.some(c => isFloorLabel(c));
    if (hasFloor) return 'theater';
  }
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const hasRowLabel = row.some(c => parseRowLabel(c) !== null);
    const hasSeatNumbers = row.filter(c => isSeatNumber(c)).length >= 3;
    if (hasRowLabel && hasSeatNumbers) return 'theater';
  }
  return 'standard';
}

/**
 * 解析 Excel 工作簿，返回 { venues, attendees } 数据
 * @param {object} wb - XLSX 工作簿对象
 * @param {Array} manualAttendees - 需要保留的手动添加的参会者
 * @param {string} mode - 解析模式：'theater'(剧院/默认), 'standard'(普通), 'u-shape'(U型), 'auto'(自动检测)
 * @param {object} sheetModes - 每Sheet独立模式: { sheetName: 'theater'|'standard'|'u-shape' }
 * @returns {{ venues: Array, attendees: Array }}
 */
function parseWorkbook(wb, manualAttendees, mode, sheetModes) {
  manualAttendees = manualAttendees || [];
  mode = mode || 'theater';
  sheetModes = sheetModes || {};
  const excelAttendees = [];
  const result = { venues: [], attendees: [] };

  wb.SheetNames.forEach((sheetName, vi) => {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    const descRow = data[0] || [];
    const description = descRow.find(c => typeof c === 'string' && c.length > 3) || sheetName;

    // 每Sheet独立的布局模式
    let sheetMode = sheetModes[sheetName] || mode;
    if (sheetMode === 'auto') {
      sheetMode = detectSheetMode(data) || 'theater';
    }

    // 投屏区可能在第0行或第1行
    const stageRow0 = data[0] || [];
    const stageRow1 = data[1] || [];
    const stageName = stageRow0.find(c => typeof c === 'string' && c.trim()) || stageRow1.find(c => typeof c === 'string' && c.trim()) || '舞台';

    const venueId = 'venue-' + vi;
    const venue = {
      id: venueId,
      name: sheetName,
      description: description,
      stageName: stageName,
      rows: [],
      mode: sheetMode
    };

    // ================================================================
    // U型会场模式：直接执行U型检测，跳过所有其他逻辑
    // ================================================================
    if (sheetMode === 'u-shape' || sheetMode === 'hui-shape') {
      const uResult = detectUShapeLayout(data, venueId, venue, excelAttendees);
      if (uResult && uResult.rows.length > 0) {
        venue.rows = uResult.rows;
        venue.layout = sheetMode;  // 'u-shape' 或 'hui-shape'
        if (uResult.attendees) {
          uResult.attendees.forEach(a => {
            if (!excelAttendees.find(e => e.name === a.name && e.row === a.row && e.seat === a.seat)) {
              excelAttendees.push(a);
            }
          });
        }
      }
      // 计算总座位数
      let totalSeats = 0;
      venue.rows.forEach(r => {
        r.seatGroups.forEach(g => { totalSeats += g.length; });
      });
      venue.totalSeats = totalSeats;
      result.venues.push(venue);
      return;
    }

    const rowInfos = [];
    const seatHeaders = [];
    const sofaRows = [];
    let currentFloor = null;
    let floorRowCounter = {};

    // 普通会场模式：跳过楼层检测和沙发排检测
    const isTheaterMode = sheetMode === 'theater';

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      // 剧院模式：检测楼层标签（独立行或同行）
      if (isTheaterMode) {
        const floorLabels = row.filter(c => isFloorLabel(c));
        const hasFloor = floorLabels.length > 0;
        if (hasFloor && floorLabels[0]) {
          const floorInfo = parseFloorLabel(floorLabels[0]);
          if (floorInfo) {
            currentFloor = floorInfo.label;
            if (!floorRowCounter[currentFloor]) floorRowCounter[currentFloor] = 0;
          }
        }
      }

      // 检测沙发排（剧院和普通模式都支持）
      let hasSofa = false;
      let sofaLabelCells = [];
      if (sheetMode !== 'u-shape') {
        hasSofa = row.some(c => isSofaRow(c));
        sofaLabelCells = row.filter(c => parseSofaLabel(c) !== null);
      }

      const rowLabelCells = row.filter(c => parseRowLabel(c) !== null);
      const rowLabelCell = rowLabelCells.length > 0 ? rowLabelCells[0] : null;

      // 剧院模式：检查是否为组合格式标签（如"一楼第X排"）
      if (isTheaterMode) {
        const combinedLabel = row.find(c => {
          if (!c || typeof c !== 'string') return false;
          const trimmed = c.trim();
          return /^.+[楼]第.+[排]$/.test(trimmed) || /^.+[楼]\d+[排]$/.test(trimmed);
        });

        if (combinedLabel) {
          const extractedFloor = extractFloorFromCombinedLabel(combinedLabel);
          if (extractedFloor) {
            currentFloor = extractedFloor;
            if (!floorRowCounter[currentFloor]) floorRowCounter[currentFloor] = 0;
          }
        }
      }

      // 格式1 - 沙发排（同一行有沙发标签 + 排标签 + 座位号）（剧院和普通模式）
      if (sheetMode !== 'u-shape' && hasSofa && rowLabelCell) {
        const seatPositions = [];
        row.forEach((c, ci) => {
          if (isSeatNumber(c)) seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
        });
        const groups = groupByGap(seatPositions);
        sofaRows.push({
          excelRow: i,
          label: '沙发' + rowLabelCell,
          rowNum: parseRowLabel(rowLabelCell),
          groups: groups
        });
        continue;
      }

      // 格式2 - 行中有沙发排标签（座位号在上方行）（剧院和普通模式）
      if (sheetMode !== 'u-shape' && sofaLabelCells.length > 0) {
        const sofaLabel = sofaLabelCells[0].trim();
        const sofaInfo = parseSofaLabel(sofaLabel);
        if (sofaInfo) {
          let nearestSeatHeader = null;
          for (let k = seatHeaders.length - 1; k >= 0; k--) {
            if (seatHeaders[k].excelRow < i) {
              nearestSeatHeader = seatHeaders[k];
              break;
            }
          }
          sofaRows.push({
            excelRow: i,
            label: sofaLabel,
            rowNum: sofaInfo.rowNum,
            groups: nearestSeatHeader ? nearestSeatHeader.groups : []
          });
        }
        continue;
      }

      // 检测座位和排标签（排除楼层标签）
      const seatPositions = [];
      row.forEach((c, ci) => {
        if (isSeatNumber(c)) seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
      });

      // 剧院模式：检查是否有"楼层+排"组合标签（如"一楼第一排"）
      const combinedFloorRow = isTheaterMode ? row.find(c => {
        if (!c || typeof c !== 'string') return false;
        const trimmed = c.trim();
        return isFloorLabel(c) && parseRowLabel(c) !== null;
      }) : null;

      // 同行既有排标签又有座位号（传统格式）
      if ((rowLabelCell || combinedFloorRow) && seatPositions.length >= 3) {
        const groups = groupByGap(seatPositions);
        seatHeaders.push({ excelRow: i, groups: groups });
        
        if (combinedFloorRow && isTheaterMode) {
          const extractedFloor = extractFloorFromCombinedLabel(combinedFloorRow);
          const rn = parseRowLabel(combinedFloorRow);
          if (extractedFloor) {
            currentFloor = extractedFloor;
            if (!floorRowCounter[currentFloor]) floorRowCounter[currentFloor] = 0;
            floorRowCounter[currentFloor]++;
            rowInfos.push({ excelRow: i, label: combinedFloorRow, rowNum: rn, floor: currentFloor });
          }
        } else if (rowLabelCell) {
          const rn = parseRowLabel(rowLabelCell);
          if (rn !== null) {
            const containsFloor = isTheaterMode && currentFloor && rowLabelCell.startsWith(currentFloor);
            const fullLabel = !containsFloor && isTheaterMode && currentFloor ? `${currentFloor}${rowLabelCell}` : rowLabelCell;
            rowInfos.push({ excelRow: i, label: fullLabel, rowNum: rn, floor: isTheaterMode ? currentFloor : undefined });
          }
        }
        continue;
      }

      // 纯排标签行（可能含人名，座位号在另一行）
      if (rowLabelCell) {
        const rn = parseRowLabel(rowLabelCell);
        if (rn !== null) {
          const uniqueLabels = [...new Set(rowLabelCells.map(c => c.trim()))];
          if (uniqueLabels.length > 1) {
            // 多标签行，忽略
          } else {
            const fullLabel = isTheaterMode && currentFloor && !rowLabelCell.startsWith(currentFloor) ? `${currentFloor}${rowLabelCell}` : rowLabelCell;
            rowInfos.push({ excelRow: i, label: fullLabel, rowNum: rn, floor: isTheaterMode ? currentFloor : undefined });
          }
        }
        continue;
      }

      // 纯座位编号头行（只有座位号，没有排标签）
      if (seatPositions.length >= 3 && !rowLabelCell) {
        const groups = groupByGap(seatPositions);
        seatHeaders.push({ excelRow: i, groups: groups });
      }
    }

    // 添加沙发排（剧院和普通模式）
    if (sheetMode !== 'u-shape') {
      sofaRows.sort((a, b) => a.rowNum - b.rowNum);
      sofaRows.forEach((sr, srIdx) => {
        venue.rows.push({
          label: sr.label,
          seatGroups: sr.groups.map(g => g.map(s => s.num))
        });
        const colMap = {};
        sr.groups.forEach(g => g.forEach(s => { colMap[s.col] = s.num; }));
        const endRow = srIdx < sofaRows.length - 1
          ? sofaRows[srIdx + 1].excelRow
          : Math.min(sr.excelRow + 3, data.length);
        for (let ri = sr.excelRow; ri < endRow; ri++) {
          const row = data[ri];
          row.forEach((cell, ci) => {
            if (typeof cell === 'string' && !isLayoutKeyword(cell) && colMap[ci] !== undefined) {
              excelAttendees.push({
                name: cell.trim(),
                row: sr.label,
                seat: colMap[ci],
                company: '',
                title: '',
                venueId: venueId,
                source: 'excel'
              });
            }
          });
        }
      });
    }

    // 为每个座位头计算过道列集合
    seatHeaders.forEach(sh => {
      const allSeatCols = new Set();
      let minCol = Infinity, maxCol = -Infinity;
      sh.groups.forEach(g => {
        g.forEach(s => {
          allSeatCols.add(s.col);
          if (s.col < minCol) minCol = s.col;
          if (s.col > maxCol) maxCol = s.col;
        });
      });
      sh.aisleCols = new Set();
      for (let c = minCol; c <= maxCol; c++) {
        if (!allSeatCols.has(c)) {
          sh.aisleCols.add(c);
        }
      }
    });
    
    // 查找适用于某行的座位头（最近的前置座位头）
    function findSeatHeader(excelRowIdx) {
      let best = null;
      seatHeaders.forEach(sh => {
        if (sh.excelRow <= excelRowIdx) {
          if (!best || sh.excelRow > best.excelRow) {
            best = sh;
          }
        }
      });
      return best;
    }

    // 识别横向过道行（第一列为"过道"的行）
    const horizontalAisleRows = new Map(); // excelRow -> label
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // 只检查第一列是否为"过道"或"横向过道"（独立过道行，不是行内标记）
      const firstCell = row[0];
      if (typeof firstCell === 'string' && /^(过道|横向过道)$/.test(firstCell.trim())) {
        horizontalAisleRows.set(i, '横向过道');
      }
    }

    // 关联普通排与座位头 - 使用座位头的过道信息精确分组
    rowInfos.sort((a, b) => a.rowNum - b.rowNum);

    rowInfos.forEach((ri, idx) => {
      const excelRow = data[ri.excelRow];
      
      // 检查该排前后是否有横向过道行（检查该排到下一个排标签行之间）
      let hasHorizontalAisleAfter = false;
      const nextRi = idx < rowInfos.length - 1 ? rowInfos[idx + 1] : null;
      const checkEnd = nextRi ? nextRi.excelRow : Math.min(ri.excelRow + 3, data.length);
      for (let r = ri.excelRow + 1; r < checkEnd; r++) {
        if (horizontalAisleRows.has(r)) {
          hasHorizontalAisleAfter = true;
          break;
        }
      }
      if (hasHorizontalAisleAfter) {
      }
      
      // 提取该排中的人名及其列位置
      const personNamesInRow = [];
      excelRow.forEach((cell, ci) => {
        if (typeof cell === 'string' && !isLayoutKeyword(cell)) {
          personNamesInRow.push({ col: ci, name: cell.trim() });
        }
      });

      if (personNamesInRow.length === 0) {
        // 该排没有人名（纯布局模板），从座位头或已有排推导布局
        const sh = findSeatHeader(ri.excelRow);
        if (sh) {
          // 使用座位头行的座位号分组
          const seatGroups = sh.groups.map(g => g.map(s => s.num));
          const seatCols = sh.groups.flat().map(s => s.col);
          venue.rows.push({ 
            label: ri.label, 
            rowNum: ri.rowNum,
            floor: ri.floor,
            seatGroups: seatGroups,
            seatCols: seatCols,
            isAisle: false,
            hasAisleAfter: hasHorizontalAisleAfter
          });
        } else {
          // 没有座位头，使用最近已添加排的布局
          let nearestRow = null;
          for (let k = venue.rows.length - 1; k >= 0; k--) {
            if (venue.rows[k].seatGroups && venue.rows[k].seatGroups.length > 0) {
              nearestRow = venue.rows[k];
              break;
            }
          }
          if (nearestRow) {
            let seatNum = 1;
            const seatGroups = nearestRow.seatGroups.map(g => g.map(() => seatNum++));
            const seatCols = nearestRow.seatCols || [];
            venue.rows.push({ 
              label: ri.label, 
              rowNum: ri.rowNum,
              floor: ri.floor,
              seatGroups: seatGroups,
              seatCols: seatCols,
              isAisle: false,
              hasAisleAfter: hasHorizontalAisleAfter
            });
          }
        }
        return;
      }

      personNamesInRow.sort((a, b) => a.col - b.col);

      // 查找适用的座位头
      let sh = findSeatHeader(ri.excelRow);

      // 标准模式：验证座位头是否与当前行人名匹配（防止错行误匹配）
      if (sh && !isTheaterMode) {
        const shCols = new Set(sh.groups.flat().map(s => s.col));
        const matchCount = personNamesInRow.filter(p => shCols.has(p.col)).length;
        if (personNamesInRow.length > 0 && matchCount / personNamesInRow.length < 0.7) {
          sh = null;
        }
      }

      if (sh) {
        const seatGroups = [];
        const seatCols = [];

        if (isTheaterMode) {
          // 剧院模式：使用座位头精确列号匹配人名
          sh.groups.forEach(headerGroup => {
            const groupSeats = [];
            headerGroup.forEach(seatPos => {
              const person = personNamesInRow.find(p => p.col === seatPos.col);
              if (person) {
                excelAttendees.push({
                  name: person.name,
                  row: ri.label,
                  seat: seatPos.num,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
              groupSeats.push(seatPos.num);
              seatCols.push(seatPos.col);
            });
            seatGroups.push(groupSeats);
          });
        } else {
          // 标准模式：保留座位头的分组结构，组间过道在渲染时处理
          let seatNum = 1;
          sh.groups.forEach((group) => {
            const groupSeats = [];
            group.forEach(seatPos => {
              const person = personNamesInRow.find(p => p.col === seatPos.col);
              if (person) {
                excelAttendees.push({
                  name: person.name,
                  row: ri.label,
                  seat: seatNum,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
              groupSeats.push(seatNum);
              seatCols.push(seatPos.col);
              seatNum++;
            });
            seatGroups.push(groupSeats);
          });
        }

        venue.rows.push({ 
          label: ri.label, 
          rowNum: ri.rowNum,
          floor: ri.floor,
          seatGroups: seatGroups,
          seatCols: seatCols,
          isAisle: false,
          hasAisleAfter: hasHorizontalAisleAfter
        });
      } else {
        // 无座位头信息
        if (!isTheaterMode) {
          // 标准模式：按列间隔分组（间隔>1列视为过道），保持多组结构
          const positionGroups = [[personNamesInRow[0]]];
          for (let i = 1; i < personNamesInRow.length; i++) {
            if (personNamesInRow[i].col - personNamesInRow[i - 1].col > 1) {
              positionGroups.push([]);
            }
            positionGroups[positionGroups.length - 1].push(personNamesInRow[i]);
          }
          let seatNum = 1;
          const seatGroups = [];
          const seatCols = [];
          positionGroups.forEach(group => {
            const groupSeats = [];
            group.forEach(person => {
              excelAttendees.push({
                name: person.name, row: ri.label, seat: seatNum,
                company: '', title: '', venueId: venueId, source: 'excel'
              });
              groupSeats.push(seatNum);
              seatCols.push(person.col);
              seatNum++;
            });
            seatGroups.push(groupSeats);
          });
          venue.rows.push({
            label: ri.label, rowNum: ri.rowNum, floor: ri.floor,
            seatGroups: seatGroups, seatCols: seatCols,
            isAisle: false, hasAisleAfter: hasHorizontalAisleAfter
          });
        } else {
          // 剧院模式：按人名列位置分组
          const positionGroups = [[personNamesInRow[0]]];
          for (let i = 1; i < personNamesInRow.length; i++) {
            const prevCol = personNamesInRow[i - 1].col;
            const curCol = personNamesInRow[i].col;
            if (curCol - prevCol > 1) {
              positionGroups.push([]);
            }
            positionGroups[positionGroups.length - 1].push(personNamesInRow[i]);
          }

          let seatNum = 1;
          const seatGroups = [];
          positionGroups.forEach(group => {
            const groupSeats = [];
            const minCol = Math.min(...group.map(p => p.col));
            const maxCol = Math.max(...group.map(p => p.col));
            for (let col = minCol; col <= maxCol; col++) {
              const person = group.find(p => p.col === col);
              if (person) {
                excelAttendees.push({
                  name: person.name,
                  row: ri.label,
                  seat: seatNum,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
              groupSeats.push(seatNum++);
            }
            seatGroups.push(groupSeats);
          });

          venue.rows.push({ 
            label: ri.label, 
            rowNum: ri.rowNum,
            floor: ri.floor,
            seatGroups: seatGroups,
            seatCols: positionGroups.flat().map(p => p.col),
            isAisle: false,
            hasAisleAfter: hasHorizontalAisleAfter
          });
        }
      }
    });

    // 根据解析结构设置会场布局类型（用于SVG/preview渲染时的模式判断）
    if (venue.rows.length > 0) {
      if (seatHeaders.length > 0) {
        venue.layout = 'standard';
      } else if (sofaRows.length > 0) {
        venue.layout = 'theater';
      }
    }

    // === 智能布局检测 ===
    // 如果常规方法未能解析出排数，尝试检测更多布局类型
    if (venue.rows.length === 0) {
      // 标准模式：只使用standard/name-matrix检测
      // 剧院模式：使用完整检测（theater → standard → name-matrix）
      const layoutResult = isTheaterMode ? detectLayoutType(data, venueId) : detectStandardLayout(data, venueId);
      if (layoutResult) {
        venue.rows = layoutResult.rows;
        venue.layout = layoutResult.layout || sheetMode;
        if (layoutResult.attendees) {
          layoutResult.attendees.forEach(a => {
            if (!excelAttendees.find(e => e.name === a.name && e.row === a.row && e.seat === a.seat)) {
              excelAttendees.push(a);
            }
          });
        }
      }
    }
    
    // === U型会议室检测（仅剧院模式）：方法1 - 列式U型布局 ===
    if (venue.rows.length === 0 && isTheaterMode) {
      let uColumnHeaderRow = -1;
      let uColumnHeaders = []; // [{col, label}]
      let uSeatNumCols = [];   // "座位号"所在的列索引

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const colLabels = [];
        const seatLabelCols = [];
        row.forEach((c, ci) => {
          if (typeof c === 'string') {
            const trimmed = c.trim();
            if (/^第.+列$/.test(trimmed)) {
              colLabels.push({ col: ci, label: trimmed });
            }
            if (trimmed === '座位号') {
              seatLabelCols.push(ci);
            }
          }
        });
        if (colLabels.length >= 2) {
          uColumnHeaderRow = i;
          uColumnHeaders = colLabels.sort((a, b) => a.col - b.col);
          uSeatNumCols = seatLabelCols.sort((a, b) => a - b);
          break;
        }
      }

      if (uColumnHeaders.length >= 2) {
        // ===== 数据驱动：当无"座位号"标签时，通过数字密度推断座位列 =====
        if (uSeatNumCols.length < 2) {
          const hMid2 = Math.floor(uColumnHeaders.length / 2);
          const _leftHeaders = uColumnHeaders.slice(0, hMid2);
          const _rightHeaders = uColumnHeaders.slice(hMid2);
          const _headerColSet = new Set(uColumnHeaders.map(h => h.col));
          const _scanRows = Math.min(uColumnHeaderRow + 30, data.length);
          const _colNums = {};
          for (let ri2 = uColumnHeaderRow + 1; ri2 < _scanRows; ri2++) {
            const _row = data[ri2]; if (!_row) continue;
            for (let ci2 = 0; ci2 < _row.length; ci2++) {
              if (typeof _row[ci2] === 'number') _colNums[ci2] = (_colNums[ci2] || 0) + 1;
            }
          }
          let _bestL = -1, _bestLC = 0;
          const _lMin = Math.max(0, _leftHeaders[0].col - 2);
          const _lMax = Math.min(_leftHeaders[_leftHeaders.length - 1].col + 1, (data[0] || []).length - 1);
          for (let ci2 = _lMin; ci2 <= _lMax; ci2++) {
            if (_headerColSet.has(ci2)) continue;
            if ((_colNums[ci2] || 0) > _bestLC && (_colNums[ci2] || 0) >= 3) { _bestLC = _colNums[ci2] || 0; _bestL = ci2; }
          }
          let _bestR = -1, _bestRC = 0;
          const _rMin = Math.max(0, _rightHeaders[0].col - 2);
          const _rMax = Math.min(_rightHeaders[_rightHeaders.length - 1].col + 1, (data[0] || []).length - 1);
          for (let ci2 = _rMin; ci2 <= _rMax; ci2++) {
            if (_headerColSet.has(ci2)) continue;
            if ((_colNums[ci2] || 0) > _bestRC && (_colNums[ci2] || 0) >= 3) { _bestRC = _colNums[ci2] || 0; _bestR = ci2; }
          }
          if (_bestL >= 0 && _bestR >= 0 && _bestL !== _bestR) {
            uSeatNumCols = [Math.min(_bestL, _bestR), Math.max(_bestL, _bestR)];
          }
        }

        // 调试日志
        console.error('[U型检测] 列头:', uColumnHeaders.map(c => `${c.label}(列${c.col})`).join(', '));
        console.error('[U型检测] 座位号列:', uSeatNumCols);
        
        // 列式U型布局确认
        const leftSeatCol = uSeatNumCols.length >= 2 ? uSeatNumCols[0] : -1;
        const rightSeatCol = uSeatNumCols.length >= 2 ? uSeatNumCols[uSeatNumCols.length - 1] : -1;

        // 将列头按中心位置分成左右两组
        // 中心位置在中间两个列头之间
        const midIdx = Math.floor(uColumnHeaders.length / 2);
        const leftCols = uColumnHeaders.slice(0, midIdx);   // 左侧所有列
        const rightCols = uColumnHeaders.slice(midIdx);     // 右侧所有列
        
        // 最里面的两列（用于检测底部座位）
        const innerLeftCol = leftCols[leftCols.length - 1];  // 左侧最里面的列
        const innerRightCol = rightCols[0];                   // 右侧最里面的列
        
        console.log('[U型检测] 左侧列:', leftCols.map(c => c.label).join(', '));
        console.log('[U型检测] 右侧列:', rightCols.map(c => c.label).join(', '));
        console.log('[U型检测] 最左内列:', innerLeftCol.label, '列索引:', innerLeftCol.col);
        console.log('[U型检测] 最右内列:', innerRightCol.label, '列索引:', innerRightCol.col);

        // 先查找底部行（在所有数据行中查找，在"第一列"和"第三列"之间）
        let bottomNums = [];
        let bottomRowIdx = -1;
        const bottomColMap = {};
        
        // 方法1：在第一列和第三列之间查找底部行
        for (let i = uColumnHeaderRow + 1; i < data.length; i++) {
          const row = data[i];
          if (!row) continue;
          const nums = [];
          row.forEach((c, ci) => {
            if (typeof c === 'number') nums.push({ col: ci, num: c });
          });
          // 底部行特征：在最里面两列之间有>=2个数字（排除座位号列，防止混合行臂座号混入）
          const innerBottomNums = nums.filter(n => 
            n.col > innerLeftCol.col && n.col < innerRightCol.col &&
            n.col !== leftSeatCol && n.col !== rightSeatCol
          );
          if (innerBottomNums.length >= 2) {
            // 按列位置排序（从左到右），而不是按座位号排序
            innerBottomNums.sort((a, b) => a.col - b.col);
            bottomNums = innerBottomNums.map(n => n.num);
            bottomRowIdx = i;
            innerBottomNums.forEach(n => { bottomColMap[n.col] = n.num; });
            console.log('[U型检测] 找到底部行（方法1-内列之间）:', i, '座位号:', bottomNums);
            break;
          }
        }
        
        // 方法2：如果在内列之间没找到，尝试在第二列和第四列下方查找延伸的座位号
        if (bottomNums.length === 0 && leftSeatCol >= 0 && rightSeatCol >= 0) {
          // 查找第二列和第四列下方的行，看是否有座位号
          for (let i = uColumnHeaderRow + 1; i < data.length; i++) {
            const row = data[i];
            if (!row) continue;
            
            // 检查这一行在第二列或第四列位置是否有座位号
            const leftSeatVal = row[leftSeatCol];
            const rightSeatVal = row[rightSeatCol];
            
            // 如果这一行在第二列或第四列位置有数字，且在中间列位置也有数字，这可能是底部行
            const middleNums = [];
            row.forEach((c, ci) => {
              if (typeof c === 'number' && ci > leftCols[leftCols.length - 1].col && ci < rightCols[0].col &&
                  ci !== leftSeatCol && ci !== rightSeatCol) {
                middleNums.push({ col: ci, num: c });
              }
            });
            
            if (middleNums.length >= 2) {
              // 按列位置排序（从左到右），而不是按座位号排序
              middleNums.sort((a, b) => a.col - b.col);
              bottomNums = middleNums.map(n => n.num);
              bottomRowIdx = i;
              middleNums.forEach(n => { bottomColMap[n.col] = n.num; });
              console.log('[U型检测] 找到底部行（方法2-外列下方）:', i, '座位号:', bottomNums);
              break;
            }
          }
        }
        
        if (bottomNums.length === 0) {
          console.log('[U型检测] 未找到底部行');
        }

        // 再扫描配对行（包含底部行所在的混合行——该行可能同时含有臂座号和底部座号）
        const pairedRows = [];
        const stopRow = bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length;
        for (let i = uColumnHeaderRow + 1; i < stopRow; i++) {
          const row = data[i];
          const leftVal = leftSeatCol >= 0 ? row[leftSeatCol] : null;
          const rightVal = rightSeatCol >= 0 ? row[rightSeatCol] : null;
          if (typeof leftVal === 'number' && typeof rightVal === 'number') {
            pairedRows.push({ excelRow: i, leftNum: leftVal, rightNum: rightVal });
          }
        }
        console.error('[U型检测] 配对行数:', pairedRows.length, '最后一个配对行:', pairedRows.length > 0 ? pairedRows[pairedRows.length - 1].excelRow : '无');


        // 创建会场行 - 左侧所有列
        // 策略：内侧列使用配对行的座位号，最外侧列额外包含底部行之后的延伸座位
        leftCols.forEach((leftCol, colIdx) => {
          // 基础座位号：来自配对行
          let leftSeatNums = pairedRows.map(r => r.leftNum);
          console.error('[DEBUG] leftSeatNums before extension:', leftSeatNums.length, 'last:', leftSeatNums[leftSeatNums.length-1]);
          
          // 如果是最外侧的列（colIdx === 0 表示左侧第一列，即"第二列"），检查底部行之后是否有延伸座位
          if (colIdx === 0 && bottomRowIdx >= 0) {
            for (let i = bottomRowIdx + 1; i < data.length; i++) {
              const row = data[i];
              if (!row) continue;
              const val = row[leftSeatCol];
              if (typeof val === 'number') {
                leftSeatNums.push(val);
              }
            }
          }
          
          // 自动补全缺失的座位号
          if (leftSeatNums.length > 0) {
            leftSeatNums = fillMissingSeatNumbers(leftSeatNums);
            
            venue.rows.push({
              label: leftCol.label,
              seatGroups: [leftSeatNums]
            });
          }
        });

        // 创建会场行 - 右侧所有列
        // 策略：内侧列使用配对行的座位号，最外侧列额外包含底部行之后的延伸座位
        rightCols.forEach((rightCol, colIdx) => {
          // 基础座位号：来自配对行
          let rightSeatNums = pairedRows.map(r => r.rightNum);
          
          // 如果是最外侧的列（最后一个），检查底部行之后是否有延伸座位
          if (colIdx === rightCols.length - 1 && bottomRowIdx >= 0) {
            for (let i = bottomRowIdx + 1; i < data.length; i++) {
              const row = data[i];
              if (!row) continue;
              const val = row[rightSeatCol];
              if (typeof val === 'number') {
                rightSeatNums.push(val);
              }
            }
          }
          
          // 自动补全缺失的座位号
          if (rightSeatNums.length > 0) {
            rightSeatNums = fillMissingSeatNumbers(rightSeatNums);
            
            venue.rows.push({
              label: rightCol.label,
              seatGroups: [rightSeatNums]
            });
          }
        });
        
        console.log('[U型检测] 左侧列座位数:', leftCols.map((c, i) => {
          const row = venue.rows.find(r => r.label === c.label);
          return `${c.label}:${row ? row.seatGroups[0].length : 0}`;
        }).join(', '));
        console.log('[U型检测] 右侧列座位数:', rightCols.map((c, i) => {
          const row = venue.rows.find(r => r.label === c.label);
          return `${c.label}:${row ? row.seatGroups[0].length : 0}`;
        }).join(', '));

        if (bottomNums.length > 0) {
          venue.rows.push({
            label: '底部',
            seatGroups: [bottomNums]
          });
        }

        venue.layout = 'u-shape';

        // 提取人名 - 左侧所有列
        leftCols.forEach(leftCol => {
          for (let i = uColumnHeaderRow + 1; i < (bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length); i++) {
            const row = data[i];
            const seatVal = row[leftSeatCol];
            const nameVal = row[leftCol.col];
            if (typeof seatVal === 'number' && typeof nameVal === 'string' && nameVal.trim() && !isLayoutKeyword(nameVal)) {
              excelAttendees.push({
                name: nameVal.trim(),
                row: leftCol.label,
                seat: seatVal,
                company: '', title: '',
                venueId: venueId,
                source: 'excel'
              });
            }
          }
        });

        // 提取人名 - 右侧所有列
        rightCols.forEach(rightCol => {
          for (let i = uColumnHeaderRow + 1; i < (bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length); i++) {
            const row = data[i];
            const seatVal = row[rightSeatCol];
            const nameVal = row[rightCol.col];
            if (typeof seatVal === 'number' && typeof nameVal === 'string' && nameVal.trim() && !isLayoutKeyword(nameVal)) {
              excelAttendees.push({
                name: nameVal.trim(),
                row: rightCol.label,
                seat: seatVal,
                company: '', title: '',
                venueId: venueId,
                source: 'excel'
              });
            }
          }
        });

        // 提取人名 - 底部（名字可能在座位号上方或下方的行中）
        if (bottomRowIdx >= 0) {
          // 检查底部行上方1行和下方2行
          for (let ri = bottomRowIdx - 1; ri < Math.min(bottomRowIdx + 3, data.length); ri++) {
            if (ri < 0 || ri === bottomRowIdx) continue; // 跳过底部行本身
            const nameRow = data[ri];
            if (!nameRow) continue;
            nameRow.forEach((cell, ci) => {
              if (typeof cell === 'string' && cell.trim() && !isLayoutKeyword(cell) && bottomColMap[ci] !== undefined) {
                excelAttendees.push({
                  name: cell.trim(),
                  row: '底部',
                  seat: bottomColMap[ci],
                  company: '', title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            });
          }
        }
      }
    }

    // === U型会议室检测: 方法2 - 描述文字U型检测（仅剧院模式） ===
    if (venue.rows.length === 0 && isTheaterMode && (description.includes('u型') || description.includes('U型'))) {
      // Step 1: 找到分组标题行（一行内有多个"第X排"标签的行，如 第一排 第二排 第三排 第四排）
      let groupHeaderRow = -1;
      let groupCols = []; // [{col, label}] 列位置→分组名

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const labelCells = [];
        row.forEach((c, ci) => {
          if (typeof c === 'string' && /^第.+排$/.test(c.trim())) {
            labelCells.push({ col: ci, label: c.trim() });
          }
        });
        // 非首列的排标签>=2个，说明是分组标题行
        const nonLeftLabels = labelCells.filter(lc => lc.col > 0);
        if (nonLeftLabels.length >= 2) {
          groupHeaderRow = i;
          groupCols = nonLeftLabels.sort((a, b) => a.col - b.col);
          break;
        }
      }

      if (groupCols.length > 0) {
        // Step 2: 找所有左侧座位标签行（第一排~第九排，代表座位编号1-9）
        const seatRows = [];
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          if (typeof row[0] === 'string' && /^第.+排$/.test(row[0].trim())) {
            const seatNum = parseRowLabel(row[0].trim());
            if (seatNum !== null) {
              seatRows.push({ excelRow: i, seatNum: seatNum });
            }
          }
        }

        // Step 3: 为每个分组创建排，每排包含所有座位编号
        const seatNums = seatRows.map(sr => sr.seatNum).sort((a, b) => a - b);
        groupCols.forEach(g => {
          venue.rows.push({
            label: g.label,
            seatGroups: [seatNums.slice()]
          });
        });

        // Step 4: 构建分组列范围（用中点分界法）
        const groupRanges = groupCols.map((g, idx) => {
          let startCol = idx === 0
            ? 1
            : Math.ceil((groupCols[idx - 1].col + g.col) / 2);
          let endCol = idx === groupCols.length - 1
            ? 9999
            : Math.ceil((g.col + groupCols[idx + 1].col) / 2) - 1;
          return { label: g.label, startCol, endCol };
        });

        // Step 5: 提取人名 —— 通过列范围判断属于哪个分组，通过行位置判断座位号
        seatRows.forEach(sr => {
          const row = data[sr.excelRow];
          row.forEach((cell, ci) => {
            if (ci === 0) return; // 跳过左侧标签列
            if (typeof cell === 'string' && !isLayoutKeyword(cell)) {
              // 通过列范围找到所属分组
              const group = groupRanges.find(gr => ci >= gr.startCol && ci <= gr.endCol);
              if (group) {
                excelAttendees.push({
                  name: cell.trim(),
                  row: group.label,
                  seat: sr.seatNum,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            }
          });
        });
      } else {
        // 兜底：从描述文字生成
        const totalMatch = description.match(/共计(\d+)人/);
        const perRowMatch = description.match(/一排(\d+)人/);
        const total = totalMatch ? parseInt(totalMatch[1]) : 36;
        const perRow = perRowMatch ? parseInt(perRowMatch[1]) : 9;
        const numRows = Math.ceil(total / perRow);

        for (let r = 1; r <= numRows; r++) {
          const seats = [];
          for (let s = 1; s <= perRow; s++) seats.push(s);
          venue.rows.push({
            label: '第' + Object.keys(cnNums).find(k => cnNums[k] === r) + '排',
            seatGroups: [seats]
          });
        }
      }
    }

    // 按排号排序（考虑楼层）
    venue.rows.sort((a, b) => {
      const aIsSofa = a.label.includes('沙发');
      const bIsSofa = b.label.includes('沙发');
      if (aIsSofa && !bIsSofa) return -1;
      if (!aIsSofa && bIsSofa) return 1;
      
      // 解析楼层号
      const aFloorNum = a.floor ? (parseFloorLabel(a.floor)?.floorNum || 0) : 0;
      const bFloorNum = b.floor ? (parseFloorLabel(b.floor)?.floorNum || 0) : 0;
      
      // 先按楼层排序，再按排号排序
      if (aFloorNum !== bFloorNum) return aFloorNum - bFloorNum;
      
      return (a.rowNum || 0) - (b.rowNum || 0);
    });

    // 去重（防止同一排被重复添加）
    const seen = new Set();
    venue.rows = venue.rows.filter(r => {
      const key = `${r.label || 'unnamed'}-${r.rowNum}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    // === 统一过道位置：确保所有排的座位组结构一致（仅剧院模式）===
    if (isTheaterMode && venue.rows.length > 1) {
      // 统计所有排的分组数
      const groupCountSig = {};
      venue.rows.forEach(r => {
        if (r.seatGroups && r.seatGroups.length > 0) {
          const sig = r.seatGroups.length.toString();
          groupCountSig[sig] = (groupCountSig[sig] || 0) + 1;
        }
      });
      
      // 找到最常见的分组数
      let maxCount = 0;
      let commonGroupCount = 0;
      for (const [sig, count] of Object.entries(groupCountSig)) {
        if (count > maxCount) {
          maxCount = count;
          commonGroupCount = parseInt(sig);
        }
      }
      
      // 如果存在最常见的分组数（且超过30%的排使用它），统一所有排使用该分组数
      if (commonGroupCount > 0 && maxCount > venue.rows.length * 0.3) {
        venue.rows.forEach(r => {
          if (r.seatGroups && r.seatGroups.length !== commonGroupCount) {
            // 将该排的所有座位号按标准分组数重新分配
            const allSeats = r.seatGroups.flat();
            const baseCount = Math.floor(allSeats.length / commonGroupCount);
            const remainder = allSeats.length % commonGroupCount;
            
            const newGroups = [];
            let seatIdx = 0;
            for (let g = 0; g < commonGroupCount; g++) {
              // 将余数分配到前面的组
              const groupSize = baseCount + (g < remainder ? 1 : 0);
              const group = [];
              for (let i = 0; i < groupSize && seatIdx < allSeats.length; i++) {
                group.push(allSeats[seatIdx++]);
              }
              newGroups.push(group);
            }
            r.seatGroups = newGroups;
          }
        });
      }
    }

    // 剧院模式：重新检测横向过道（楼层变化/座位数变化触发）
    if (isTheaterMode) {
      for (let i = 0; i < venue.rows.length - 1; i++) {
        const currentRow = venue.rows[i];
        const nextRow = venue.rows[i + 1];
        
        if (currentRow.seatGroups && nextRow.seatGroups) {
          if (currentRow.floor && nextRow.floor && currentRow.floor !== nextRow.floor) {
            venue.rows[i].hasAisleAfter = true;
            continue;
          }
          
          if (i > 0 && venue.rows[i - 1].hasAisleAfter) {
            continue;
          }
          
          const currentTotal = currentRow.seatGroups.reduce((sum, g) => sum + g.length, 0);
          const nextTotal = nextRow.seatGroups.reduce((sum, g) => sum + g.length, 0);
          const totalDiff = Math.abs(currentTotal - nextTotal);
          const maxTotal = Math.max(currentTotal, nextTotal);
          
          if (maxTotal > 0 && totalDiff / maxTotal > 0.25) {
            venue.rows[i].hasAisleAfter = true;
          }
        }
      }
    }

    // 计算总座位数
    let totalSeats = 0;
    venue.rows.forEach(r => {
      r.seatGroups.forEach(g => { totalSeats += g.length; });
    });
    venue.totalSeats = totalSeats;

    result.venues.push(venue);
  });

  // 合并：Excel提取的 + 手动添加的
  result.attendees = excelAttendees.concat(manualAttendees);

  return result;
}

/**
 * U型会场布局检测（独立函数，供U型模式直接调用）
 */
function detectUShapeLayout(data, venueId, venue, excelAttendees) {
  const result = { rows: [], attendees: [] };

  let uColumnHeaderRow = -1;
  let uColumnHeaders = [];
  let uSeatNumCols = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const colLabels = [];
    const seatLabelCols = [];
    row.forEach((c, ci) => {
      if (typeof c === 'string') {
        const trimmed = c.trim();
        if (/^第.+列$/.test(trimmed)) {
          colLabels.push({ col: ci, label: trimmed });
        }
        if (trimmed === '座位号') {
          seatLabelCols.push(ci);
        }
      }
    });
    if (colLabels.length >= 2) {
      uColumnHeaderRow = i;
      uColumnHeaders = colLabels.sort((a, b) => a.col - b.col);
      if (seatLabelCols.length > 0) {
        uSeatNumCols = seatLabelCols.sort((a, b) => a - b);
      }
      break;
    }
  }

  // 如果在列头行未找到座位号标签，在后续行中查找
  if (uColumnHeaders.length >= 2 && uSeatNumCols.length < 2) {
    const allSeatCols = new Set();
    for (let i = uColumnHeaderRow + 1; i < Math.min(uColumnHeaderRow + 4, data.length); i++) {
      const row = data[i];
      if (!row) continue;
      row.forEach((c, ci) => {
        if (typeof c === 'string' && c.trim() === '座位号') {
          allSeatCols.add(ci);
        }
      });
    }
    if (allSeatCols.size >= 2) {
      uSeatNumCols = Array.from(allSeatCols).sort((a, b) => a - b);
    }
  }

  // ===== 数据驱动：当无"座位号"标签时，通过数字密度推断座位列 =====
  // 核心思想：人名不会是全数字，一排/一列全是数字的就是座位号
  if (uColumnHeaders.length >= 2 && uSeatNumCols.length < 2) {
    const hMid = Math.floor(uColumnHeaders.length / 2);
    const leftHeaders = uColumnHeaders.slice(0, hMid);
    const rightHeaders = uColumnHeaders.slice(hMid);
    const headerColSet = new Set(uColumnHeaders.map(h => h.col));

    // 扫描标题行下方30行，统计每列的数字个数
    const scanRows = Math.min(uColumnHeaderRow + 30, data.length);
    const colNums = {};
    for (let ri = uColumnHeaderRow + 1; ri < scanRows; ri++) {
      const row = data[ri]; if (!row) continue;
      for (let ci = 0; ci < row.length; ci++) {
        if (typeof row[ci] === 'number') {
          colNums[ci] = (colNums[ci] || 0) + 1;
        }
      }
    }

    // 在左组标题附近找数字密度最高的列
    let bestLeft = -1, bestLeftCnt = 0;
    const leftMin = Math.max(0, leftHeaders[0].col - 2);
    const leftMax = Math.min(leftHeaders[leftHeaders.length - 1].col + 1, (data[0] || []).length - 1);
    for (let ci = leftMin; ci <= leftMax; ci++) {
      if (headerColSet.has(ci)) continue;
      const cnt = colNums[ci] || 0;
      if (cnt > bestLeftCnt && cnt >= 3) { bestLeftCnt = cnt; bestLeft = ci; }
    }

    // 在右组标题附近找数字密度最高的列
    let bestRight = -1, bestRightCnt = 0;
    const rightMin = Math.max(0, rightHeaders[0].col - 2);
    const rightMax = Math.min(rightHeaders[rightHeaders.length - 1].col + 1, (data[0] || []).length - 1);
    for (let ci = rightMin; ci <= rightMax; ci++) {
      if (headerColSet.has(ci)) continue;
      const cnt = colNums[ci] || 0;
      if (cnt > bestRightCnt && cnt >= 3) { bestRightCnt = cnt; bestRight = ci; }
    }

    if (bestLeft >= 0 && bestRight >= 0 && bestLeft !== bestRight) {
      uSeatNumCols = [Math.min(bestLeft, bestRight), Math.max(bestLeft, bestRight)];
    }
  }

  if (uColumnHeaders.length >= 2) {
    // 如果没有找到座位号列，使用列头所在列作为座位列
    const leftSeatCol = uSeatNumCols.length >= 2 ? uSeatNumCols[0] : -1;
    const rightSeatCol = uSeatNumCols.length >= 2 ? uSeatNumCols[uSeatNumCols.length - 1] : -1;

    const midIdx = Math.floor(uColumnHeaders.length / 2);
    const leftCols = uColumnHeaders.slice(0, midIdx);
    const rightCols = uColumnHeaders.slice(midIdx);

    const innerLeftCol = leftCols[leftCols.length - 1];
    const innerRightCol = rightCols[0];

    let bottomNums = [];
    let bottomRowIdx = -1;
    const bottomColMap = {};

    for (let i = uColumnHeaderRow + 1; i < data.length; i++) {
      const row = data[i];
      if (!row) continue;
      const nums = [];
      row.forEach((c, ci) => {
        if (typeof c === 'number') nums.push({ col: ci, num: c });
      });
      const innerBottomNums = nums.filter(n =>
        n.col > innerLeftCol.col && n.col < innerRightCol.col &&
        n.col !== leftSeatCol && n.col !== rightSeatCol
      );
      if (innerBottomNums.length >= 2) {
        innerBottomNums.sort((a, b) => a.col - b.col);
        bottomNums = innerBottomNums.map(n => n.num);
        bottomRowIdx = i;
        innerBottomNums.forEach(n => { bottomColMap[n.col] = n.num; });
        break;
      }
    }

    if (bottomNums.length === 0 && leftSeatCol >= 0 && rightSeatCol >= 0) {
      for (let i = uColumnHeaderRow + 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const middleNums = [];
        row.forEach((c, ci) => {
          if (typeof c === 'number' && ci > leftCols[leftCols.length - 1].col && ci < rightCols[0].col &&
              ci !== leftSeatCol && ci !== rightSeatCol) {
            middleNums.push({ col: ci, num: c });
          }
        });
        if (middleNums.length >= 2) {
          middleNums.sort((a, b) => a.col - b.col);
          bottomNums = middleNums.map(n => n.num);
          bottomRowIdx = i;
          middleNums.forEach(n => { bottomColMap[n.col] = n.num; });
          break;
        }
      }
    }

    const pairedRows = [];
    const stopRow = bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length;
    for (let i = uColumnHeaderRow + 1; i < stopRow; i++) {
      const row = data[i];
      const leftVal = leftSeatCol >= 0 ? row[leftSeatCol] : null;
      const rightVal = rightSeatCol >= 0 ? row[rightSeatCol] : null;
      if (typeof leftVal === 'number' && typeof rightVal === 'number') {
        pairedRows.push({ excelRow: i, leftNum: leftVal, rightNum: rightVal });
      }
    }

    leftCols.forEach((leftCol, colIdx) => {
      let leftSeatNums = pairedRows.map(r => r.leftNum);
      if (colIdx === 0 && bottomRowIdx >= 0) {
        for (let i = bottomRowIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row) continue;
          const val = row[leftSeatCol];
          if (typeof val === 'number') {
            leftSeatNums.push(val);
          }
        }
      }
      if (leftSeatNums.length > 0) {
        leftSeatNums = fillMissingSeatNumbers(leftSeatNums);
        result.rows.push({
          label: leftCol.label,
          seatGroups: [leftSeatNums]
        });
      }
    });

    rightCols.forEach((rightCol, colIdx) => {
      let rightSeatNums = pairedRows.map(r => r.rightNum);
      if (colIdx === rightCols.length - 1 && bottomRowIdx >= 0) {
        for (let i = bottomRowIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row) continue;
          const val = row[rightSeatCol];
          if (typeof val === 'number') {
            rightSeatNums.push(val);
          }
        }
      }
      if (rightSeatNums.length > 0) {
        rightSeatNums = fillMissingSeatNumbers(rightSeatNums);
        result.rows.push({
          label: rightCol.label,
          seatGroups: [rightSeatNums]
        });
      }
    });

    // ===== 回字型：检测顶部行（列标题行上方的高数字密度行）=====
    // 核心思想：标题上方凡是数字密集的行就是顶部座位行
    const topRowData = [];
    for (let i = 0; i < uColumnHeaderRow; i++) {
      const row = data[i];
      if (!row) continue;
      const nums = [];
      row.forEach((c, ci) => {
        if (typeof c === 'number' && ci > innerLeftCol.col && ci < innerRightCol.col &&
            ci !== leftSeatCol && ci !== rightSeatCol) {
          nums.push(c);
        }
      });
      if (nums.length >= 2) {
        // 查找下一行的标签文本（标签通常在座位行下方）
        let label = '顶部';
        if (i + 1 < data.length) {
          const nextRow = data[i + 1];
          if (nextRow) {
            for (let ci = 0; ci < nextRow.length; ci++) {
              const c = nextRow[ci];
              if (typeof c === 'string' && c.trim() && !isLayoutKeyword(c)) {
                const t = c.trim();
                // 排除列标题模式
                if (!/^第.+列$/.test(t)) { label = t; break; }
              }
            }
          }
        }
        const sortedNums = nums.sort((a, b) => a - b);
        topRowData.push({
          label: label,
          seatGroups: [fillMissingSeatNumbers(sortedNums)]
        });
      }
    }
    // 顶部行插入到 result.rows 开头
    for (let ti = topRowData.length - 1; ti >= 0; ti--) {
      result.rows.unshift(topRowData[ti]);
    }

    // ===== 多底部行检测：扫描第一个底部行之后的后续底部座位行 =====
    const allBottomRows = [];
    if (bottomNums.length > 0) {
      // 从混合行相邻行查找首行底部标签（不再硬编码"底部"）
      let firstBotLabel = '底部';
      if (bottomRowIdx + 1 < data.length) {
        const nextRow = data[bottomRowIdx + 1];
        if (nextRow) {
          for (let ci = 0; ci < nextRow.length; ci++) {
            const c = nextRow[ci];
            if (typeof c === 'string' && c.trim() && !isLayoutKeyword(c)) {
              const t = c.trim();
              if (!/^第.+列$/.test(t)) { firstBotLabel = t; break; }
            }
          }
        }
      }
      allBottomRows.push({ nums: bottomNums, rowIdx: bottomRowIdx, label: firstBotLabel });
    }
    // 继续向后扫描更多底部行
    if (bottomRowIdx >= 0) {
      const innerLCol = innerLeftCol.col;
      const innerRCol = innerRightCol.col;
      for (let i = bottomRowIdx + 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const nums = [];
        row.forEach((c, ci) => {
          if (typeof c === 'number' && ci > innerLCol && ci < innerRCol &&
              ci !== leftSeatCol && ci !== rightSeatCol) {
            nums.push(c);
          }
        });
        if (nums.length >= 2) {
          let label = '底部';
          if (i + 1 < data.length) {
            const nextRow = data[i + 1];
            if (nextRow) {
              for (let ci = 0; ci < nextRow.length; ci++) {
                const c = nextRow[ci];
                if (typeof c === 'string' && c.trim() && !isLayoutKeyword(c)) {
                  const t = c.trim();
                  if (!/^第.+列$/.test(t)) { label = t; break; }
                }
              }
            }
          }
          const sortedNums = nums.sort((a, b) => a - b);
          allBottomRows.push({ nums: sortedNums, rowIdx: i, label: label });
          // 跳过标签行继续查找
          i++;
        }
      }
    }
    // 推入所有底部行
    allBottomRows.forEach(bRow => {
      if (bRow.nums.length > 0) {
        result.rows.push({
          label: bRow.label,
          seatGroups: [bRow.nums]
        });
      }
    });

    leftCols.forEach(leftCol => {
      for (let i = uColumnHeaderRow + 1; i < (bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length); i++) {
        const row = data[i];
        const seatVal = row[leftSeatCol];
        const nameVal = row[leftCol.col];
        if (typeof seatVal === 'number' && typeof nameVal === 'string' && nameVal.trim() && !isLayoutKeyword(nameVal)) {
          result.attendees.push({
            name: nameVal.trim(),
            row: leftCol.label,
            seat: seatVal,
            company: '', title: '',
            venueId: venueId,
            source: 'excel'
          });
        }
      }
    });

    rightCols.forEach(rightCol => {
      for (let i = uColumnHeaderRow + 1; i < (bottomRowIdx >= 0 ? bottomRowIdx + 1 : data.length); i++) {
        const row = data[i];
        const seatVal = row[rightSeatCol];
        const nameVal = row[rightCol.col];
        if (typeof seatVal === 'number' && typeof nameVal === 'string' && nameVal.trim() && !isLayoutKeyword(nameVal)) {
          result.attendees.push({
            name: nameVal.trim(),
            row: rightCol.label,
            seat: seatVal,
            company: '', title: '',
            venueId: venueId,
            source: 'excel'
          });
        }
      }
    });

    // ===== 提取底部行参会者姓名（支持多底部行）=====
    allBottomRows.forEach(bRow => {
      for (let ri = bRow.rowIdx - 1; ri < Math.min(bRow.rowIdx + 3, data.length); ri++) {
        if (ri < 0 || ri === bRow.rowIdx) continue;
        const nameRow = data[ri];
        if (!nameRow) continue;
        // 构建该底部行的列→座位号映射
        const bColMap = {};
        bRow.nums.forEach((num, idx) => {
          // 需要知道数字在行中的列位置
          const numRow = data[bRow.rowIdx];
          let foundCol = -1;
          if (numRow) {
            for (let ci = innerLeftCol.col + 1; ci < innerRightCol.col; ci++) {
              if (numRow[ci] === num) { foundCol = ci; break; }
            }
          }
          if (foundCol >= 0) bColMap[foundCol] = num;
        });
        nameRow.forEach((cell, ci) => {
          if (typeof cell === 'string' && cell.trim() && !isLayoutKeyword(cell) && bColMap[ci] !== undefined) {
            result.attendees.push({
              name: cell.trim(),
              row: bRow.label,
              seat: bColMap[ci],
              company: '', title: '',
              venueId: venueId,
              source: 'excel'
            });
          }
        });
      }
    });
  }

  return result;
}

/**
 * 标准会场布局检测（仅标准矩阵+姓名矩阵，不含剧院/U型）
 */
function detectStandardLayout(data, venueId) {
  const rows = [];
  const attendees = [];

  // 方法1：检测标准矩阵布局（每行有排标签，每列有座位号）
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowLabel = row.find(c => parseRowLabel(c) !== null);
    if (rowLabel) {
      const rn = parseRowLabel(rowLabel);
      if (rn !== null) {
        const seatPositions = [];
        row.forEach((c, ci) => {
          if (isSeatNumber(c)) {
            seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
          }
        });

        if (seatPositions.length > 0) {
          const groups = groupByGap(seatPositions);
          rows.push({
            label: rowLabel,
            rowNum: rn,
            seatGroups: groups.map(g => g.map(s => s.num))
          });

          row.forEach((cell, ci) => {
            if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell)) {
              const seatCol = seatPositions.find(sp => sp.col === ci);
              if (seatCol) {
                attendees.push({
                  name: cell.trim(),
                  row: rowLabel,
                  seat: seatCol.num,
                  company: '',
                  title: '',
                  venueId: venueId,
                  source: 'excel'
                });
              }
            }
          });
        }
      }
    }
  }

  if (rows.length > 0) {
    rows.sort((a, b) => (a.rowNum || 0) - (b.rowNum || 0));
    return { rows, layout: 'standard', attendees };
  }

  // 方法2：检测只有座位号的布局
  let seatNumberRows = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const seatPositions = [];
    row.forEach((c, ci) => {
      if (isSeatNumber(c)) {
        seatPositions.push({ col: ci, num: normalizeSeatNumber(c) });
      }
    });

    if (seatPositions.length >= 3) {
      seatNumberRows.push({
        excelRow: i,
        seatPositions: seatPositions,
        groups: groupByGap(seatPositions)
      });
    }
  }

  if (seatNumberRows.length > 0) {
    seatNumberRows.forEach((sr, idx) => {
      const rowNum = idx + 1;
      const label = `第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排`;
      rows.push({
        label: label,
        rowNum: rowNum,
        seatGroups: sr.groups.map(g => g.map(s => s.num))
      });
    });

    return { rows, layout: 'seat-numbers-only', attendees };
  }

  // 方法3：检测人员姓名矩阵布局
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const personNamesInRow = [];
    row.forEach((cell, ci) => {
      if (typeof cell === 'string' && !isLayoutKeyword(cell) && !isSeatNumber(cell)) {
        personNamesInRow.push({ col: ci, name: cell.trim() });
      }
    });

    if (personNamesInRow.length >= 2) {
      personNamesInRow.sort((a, b) => a.col - b.col);
      const positionGroups = [[personNamesInRow[0]]];
      for (let j = 1; j < personNamesInRow.length; j++) {
        const prevCol = personNamesInRow[j - 1].col;
        const curCol = personNamesInRow[j].col;
        if (curCol - prevCol > 1) {
          positionGroups.push([]);
        }
        positionGroups[positionGroups.length - 1].push(personNamesInRow[j]);
      }

      let seatNum = 1;
      const seatGroups = [];
      positionGroups.forEach(group => {
        const groupSeats = [];
        group.forEach((person, personIdx) => {
          attendees.push({
            name: person.name,
            row: `第${Object.keys(cnNums).find(k => cnNums[k] === i + 1) || (i + 1)}排`,
            seat: seatNum + personIdx,
            company: '',
            title: '',
            venueId: venueId,
            source: 'excel'
          });
          groupSeats.push(seatNum + personIdx);
        });
        seatGroups.push(groupSeats);
        seatNum += group.length;
      });

      const rowNum = i + 1;
      rows.push({
        label: `第${Object.keys(cnNums).find(k => cnNums[k] === rowNum) || rowNum}排`,
        rowNum: rowNum,
        seatGroups: seatGroups
      });
    }
  }

  if (rows.length > 0) {
    return { rows, layout: 'name-matrix', attendees };
  }

  return null;
}

// 导出模块
module.exports = { parseWorkbook, detectSheetMode };

// 如果直接运行（node parse-excel.js），执行原有的CLI逻辑
if (require.main === module) {
  // 支持命令行参数传入文件路径，方便Python调用
  const filePath = process.argv[2] || path.join(__dirname, 'uploads', 'uploaded.xlsx');
  console.log('[parse-excel] 解析文件:', filePath);
  
  const wb = XLSX.readFile(filePath);

  // 保留已有参会者数据（仅保留手动添加的，Excel提取的会重新生成）
  let manualAttendees = [];
  try {
    const dataDir = path.join(__dirname, 'data');
    const dataFile = path.join(dataDir, 'data.json');
    if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
    const existing = JSON.parse(fs.readFileSync(dataFile, 'utf-8'));
    manualAttendees = (existing.attendees || []).filter(a => a.source !== 'excel');
  } catch {}

  const result = parseWorkbook(wb, manualAttendees);

  // 确保 data 目录存在
  const dataOutputDir = path.join(__dirname, 'data');
  if (!fs.existsSync(dataOutputDir)) fs.mkdirSync(dataOutputDir, { recursive: true });
  // 写入 data.json
  fs.writeFileSync(path.join(dataOutputDir, 'data.json'), JSON.stringify(result, null, 2), 'utf-8');

  console.log('解析完成！');
  result.venues.forEach(v => {
    const va = result.attendees.filter(a => a.venueId === v.id);
    console.log(`\n【${v.name}】${v.description}`);
    console.log(`  共 ${v.rows.length} 排, ${v.totalSeats} 个座位, 已安排 ${va.length} 人`);
    v.rows.forEach(r => {
      const seats = r.seatGroups.map(g => g.join(',')).join(' | ');
      console.log(`  ${r.label}: [${seats}]`);
    });
    if (va.length > 0) {
      console.log('  --- 已安排人员 ---');
      va.forEach(a => console.log(`  ${a.name} → ${a.row} ${a.seat}座`));
    }
  });
}
