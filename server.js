const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const QRCode = require('qrcode');
const bcrypt = require('bcryptjs');
const helmet = require('helmet');
const { parseWorkbook, detectSheetMode } = require('./parse-excel');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'data.json');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json'); // 配置文件也放在 data 目录
const BACKUP_DIR = path.join(DATA_DIR, 'backups');

// ==================== 安全中间件 ====================

// ==================== 日志系统 ====================

// 访问日志（JSON 结构化格式）
const accessLogFile = path.join(__dirname, 'access.log');
const accessLogStream = fs.createWriteStream(accessLogFile, { flags: 'a' });

// 错误日志
const errorLogFile = path.join(__dirname, 'error.log');
const errorLogStream = fs.createWriteStream(errorLogFile, { flags: 'a' });

// 审计日志文件
const auditLogFile = path.join(__dirname, 'audit.log');
const auditLogStream = fs.createWriteStream(auditLogFile, { flags: 'a' });

// 结构化日志函数
function structuredLog(type, data) {
  const entry = {
    timestamp: new Date().toISOString(),
    type,
    ...data
  };
  
  if (type === 'access') {
    accessLogStream.write(JSON.stringify(entry) + '\n');
  } else if (type === 'error') {
    errorLogStream.write(JSON.stringify(entry) + '\n');
    // 错误同时输出到控制台
    console.error(`[ERROR] ${data.message}`, data.stack || '');
  } else if (type === 'system') {
    console.log(`[SYSTEM] ${data.message}`);
  }
}

// 审计日志函数
function auditLog(action, ip, details = {}) {
  const entry = {
    timestamp: new Date().toISOString(),
    type: 'audit',
    action,
    ip,
    ...details
  };
  auditLogStream.write(JSON.stringify(entry) + '\n');
}

// 请求日志中间件
app.use((req, res, next) => {
  const startTime = Date.now();
  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress || '-';
  const method = req.method;
  const url = req.originalUrl || req.url;

  res.on('finish', () => {
    const duration = Date.now() - startTime;
    const status = res.statusCode;
    
    structuredLog('access', {
      ip,
      method,
      url,
      status,
      duration_ms: duration,
      user_agent: req.headers['user-agent'] || '-',
      content_length: parseInt(req.headers['content-length']) || 0
    });
  });

  next();
});

// Helmet: 设置安全 HTTP 头
app.use(helmet({
  contentSecurityPolicy: false, // 允许内联脚本（现有页面需要）
  crossOriginEmbedderPolicy: false
}));

// ==================== 配置管理 ====================

const SALT_ROUNDS = 10;
const TOKEN_TTL = 24 * 60 * 60 * 1000; // Token 有效期 24 小时
const MAX_BACKUPS = 10; // 最多保留备份数

// 读取/初始化配置
function readConfig() {
  try {
    const config = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
    let needsSave = false;
    // 如果是明文密码，自动转换为 bcrypt 哈希
    if (config.adminPassword && !config.adminPassword.startsWith('$2')) {
      const hashed = bcrypt.hashSync(config.adminPassword, SALT_ROUNDS);
      config.adminPassword = hashed;
      needsSave = true;
      structuredLog('system', { message: '密码已自动加密存储' });
    }
    // 如果主办方密码不存在，添加默认的
    if (!config.organizerPassword) {
      config.organizerPassword = bcrypt.hashSync('organizer888', SALT_ROUNDS);
      needsSave = true;
      structuredLog('system', { message: '主办方密码已添加' });
    }
    // 如果主办方密码是明文，自动加密
    else if (config.organizerPassword && !config.organizerPassword.startsWith('$2')) {
      const hashed = bcrypt.hashSync(config.organizerPassword, SALT_ROUNDS);
      config.organizerPassword = hashed;
      needsSave = true;
      structuredLog('system', { message: '主办方密码已自动加密存储' });
    }
    if (needsSave) {
      writeConfig(config);
    }
    return config;
  } catch {
    // 默认密码 admin888，自动哈希
    const defaultConfig = {
      adminPassword: bcrypt.hashSync('admin888', SALT_ROUNDS),
      organizerPassword: bcrypt.hashSync('organizer888', SALT_ROUNDS),
      tokenTtl: TOKEN_TTL
    };
    writeConfig(defaultConfig);
    return defaultConfig;
  }
}

function writeConfig(config) {
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2), 'utf-8');
}

// 确保数据目录和备份目录存在
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
if (!fs.existsSync(BACKUP_DIR)) {
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
}

// ==================== Token 管理（带过期和角色） ====================

// 存储格式：Map<token, { createdAt: number, role: 'admin' | 'organizer' }>
const validTokens = new Map();

// 定期清理过期 token（每小时执行一次）
setInterval(() => {
  const config = readConfig();
  const ttl = config.tokenTtl || TOKEN_TTL;
  const now = Date.now();
  let cleaned = 0;
  for (const [token, meta] of validTokens.entries()) {
    if (now - meta.createdAt > ttl) {
      validTokens.delete(token);
      cleaned++;
    }
  }
  if (cleaned > 0) {
    structuredLog('system', { message: `清理了 ${cleaned} 个过期 token` });
  }
}, 60 * 60 * 1000);

// ==================== 数据管理（带文件锁、备份和内存缓存） ====================

// 内存缓存：数据加载到内存，避免每次请求都读磁盘
let dataCache = null;
let dataCacheVersion = 0;

// 简单的文件锁实现
let fileLock = null;
const LOCK_TIMEOUT = 5000; // 锁超时时间 5 秒

function acquireLock() {
  const now = Date.now();
  if (fileLock && now - fileLock.timestamp > LOCK_TIMEOUT) {
    console.warn('文件锁超时，自动释放');
    fileLock = null;
  }
  if (fileLock) {
    throw new Error('文件正在被其他操作占用，请稍后重试');
  }
  fileLock = { timestamp: now };
}

function releaseLock() {
  fileLock = null;
}

// 从磁盘读取数据并更新缓存
function loadDataFromFile() {
  try {
    dataCache = JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
    
    // ==================== 数据自动清理 ====================
    let needsCleanup = false;
    const venueIds = new Set((dataCache.venues || []).map(v => v.id));
    const now = Date.now();
    
    // 1. 清理和修复参会者数据
    const cleanedAttendees = (dataCache.attendees || []).filter(a => {
      let keep = true;
      
      // 1.1 如果没有 id，生成一个
      if (!a.id) {
        a.id = 'att-' + now + '-' + Math.random().toString(36).substr(2, 9);
        needsCleanup = true;
        console.log(`[数据清理] 为参会者 ${a.name || '未命名'} 生成了 id: ${a.id}`);
      }
      
      // 1.2 如果有 venueId，但这个会场不存在了
      if (a.venueId && !venueIds.has(a.venueId)) {
        // 检查是否同时有 row 和 seat
        if (a.row && a.seat) {
          // 这是一个指向不存在会场的参会者，把它变成未分配状态
          console.log(`[数据清理] 参会者 ${a.name || a.id} 的 venueId=${a.venueId} 不存在，已设为未分配状态`);
          delete a.venueId;
          delete a.row;
          delete a.seat;
          needsCleanup = true;
        }
      }
      
      return keep;
    });
    
    if (dataCache.attendees.length !== cleanedAttendees.length) {
      dataCache.attendees = cleanedAttendees;
      needsCleanup = true;
    }
    
    // 2. 确保必要字段存在
    if (!dataCache.venues) dataCache.venues = [];
    if (!dataCache.attendees) dataCache.attendees = [];
    if (!dataCache.deletedVenueNames) dataCache.deletedVenueNames = [];
    if (!dataCache.deletedVenueNamesNormalized) dataCache.deletedVenueNamesNormalized = [];
    if (!dataCache.deletedAttendees) dataCache.deletedAttendees = [];
    
    // 如果有数据被清理，写回文件
    if (needsCleanup) {
      console.log('[数据清理] 检测到数据问题，已自动修复并保存');
      writeData(dataCache, true); // 跳过备份，因为这是自动清理
    }
    
    dataCacheVersion++;
    return dataCache;
  } catch {
    dataCache = { 
      venues: [], 
      attendees: [], 
      deletedVenueNames: [], 
      deletedVenueNamesNormalized: [],
      deletedAttendees: [] 
    };
    dataCacheVersion++;
    return dataCache;
  }
}

// 获取数据（优先使用缓存）
function readData() {
  if (!dataCache) {
    return loadDataFromFile();
  }
  return dataCache;
}

// 强制刷新缓存（从磁盘重新加载）
function refreshCache() {
  return loadDataFromFile();
}

// 数据备份
function backupData() {
  if (!fs.existsSync(DATA_FILE)) return;
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupPath = path.join(BACKUP_DIR, `data-${timestamp}.json`);
    fs.copyFileSync(DATA_FILE, backupPath);
    structuredLog('system', { message: `数据已备份至 ${backupPath}` });

    // 清理旧备份（保留最新的 MAX_BACKUPS 个）
    const backups = fs.readdirSync(BACKUP_DIR)
      .filter(f => f.startsWith('data-') && f.endsWith('.json'))
      .sort()
      .reverse();
    if (backups.length > MAX_BACKUPS) {
      backups.slice(MAX_BACKUPS).forEach(f => {
        fs.unlinkSync(path.join(BACKUP_DIR, f));
      });
    }
  } catch (err) {
    console.error('备份失败:', err.message);
  }
}

function writeData(data, skipBackup = false) {
  acquireLock();
  try {
    // 确保所有参会者都有唯一 ID（防止 Excel 导入等场景缺少 id）
    (data.attendees || []).forEach(a => {
      if (!a.id) {
        a.id = 'att-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
      }
    });

    // 写入前先备份
    if (!skipBackup) {
      backupData();
    }
    // 原子写入：先写临时文件，再重命名
    const tmpFile = DATA_FILE + '.tmp';
    fs.writeFileSync(tmpFile, JSON.stringify(data, null, 2), 'utf-8');
    fs.renameSync(tmpFile, DATA_FILE);
    // 同步更新内存缓存
    dataCache = JSON.parse(JSON.stringify(data)); // 深拷贝
    dataCacheVersion++;
  } finally {
    releaseLock();
  }
}

// ==================== 数据验证 ====================

// 中文数字映射
const numToCn = { 1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',7:'七',8:'八',9:'九',10:'十',
  11:'十一',12:'十二',13:'十三',14:'十四',15:'十五',16:'十六',17:'十七',18:'十八',19:'十九',20:'二十',21:'二十一' };

// 规范化排名
function normalizeRowLabel(raw) {
  if (!raw) return raw;
  let s = raw.trim().replace(/\s+/g, '');

  // 方向前缀标准化: 前/后/左/右/内前/内后 + 第X排/列
  var dirMatch = s.match(/^(前|后|左|右|内前|内后)[第]?(\d+|[一二三四五六七八九十]+)([排列])$/);
  if (dirMatch) {
    let num = dirMatch[2];
    if (/^\d+$/.test(num)) num = numToCn[parseInt(num)] || num;
    return dirMatch[1] + '第' + num + dirMatch[3];
  }

  const sofaMatch = s.match(/^沙发[第]?(\d+|[一二三四五六七八九十]+)排?$/);
  if (sofaMatch) {
    let num = sofaMatch[1];
    if (/^\d+$/.test(num)) num = numToCn[parseInt(num)] || num;
    return '沙发第' + num + '排';
  }

  const match = s.match(/^[第]?(\d+|[一二三四五六七八九十]+)排$/);
  if (match) {
    let num = match[1];
    if (/^\d+$/.test(num)) num = numToCn[parseInt(num)] || num;
    return '第' + num + '排';
  }

  const colMatch = s.match(/^[第]?(\d+|[一二三四五六七八九十]+)列$/);
  if (colMatch) {
    let num = colMatch[1];
    if (/^\d+$/.test(num)) num = numToCn[parseInt(num)] || num;
    return '第' + num + '列';
  }

  return s;
}

// 清理参会者数据
function cleanAttendee(a) {
  return {
    ...a,
    name: (a.name || '').trim(),
    row: normalizeRowLabel(a.row),
    seat: typeof a.seat === 'string' ? parseInt(a.seat) : a.seat,
    company: (a.company || '').trim(),
    title: (a.title || '').trim()
  };
}

// 检查参会者是否已存在（同会场同姓名）
function findDuplicateAttendee(data, attendee) {
  return data.attendees.find(a =>
    a.venueId === attendee.venueId &&
    a.name === attendee.name
  );
}

// 检查座位冲突（同会场同排同座位号）
function findSeatConflict(data, attendee) {
  return data.attendees.find(a =>
    a.venueId === attendee.venueId &&
    a.row === attendee.row &&
    a.seat === attendee.seat
  );
}

// 记录已删除参会者（防止自动分析时重新创建）
function recordDeletedAttendee(data, a) {
  if (!data.deletedAttendees) data.deletedAttendees = [];
  const base = `${a.name || ''}|${a.venueId || ''}|${a.row || ''}|${a.seat || ''}`;
  if (!data.deletedAttendees.includes(base)) data.deletedAttendees.push(base);
  const normalizedName = (a.name || '').trim().toLowerCase().replace(/\s+/g, '');
  if (normalizedName) {
    const nk = `${normalizedName}|${a.venueId || ''}|${a.row || ''}|${a.seat || ''}`;
    if (!data.deletedAttendees.includes(nk)) data.deletedAttendees.push(nk);
  }
}

// XML 转义
function escXml(str) {
  return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ==================== Express 配置 ====================

app.use(express.json({ limit: '50mb' }));

// ==================== 系统配置 API ====================

// 获取系统标题（公开）
app.get('/api/site-config', (req, res) => {
  const config = readConfig();
  res.json({
    siteTitle: config.siteTitle || '数据创新发展大会',
    siteSubtitle: config.siteSubtitle || '座位查询系统',
    siteTitleEn: config.siteTitleEn || '',
    siteSubtitleEn: config.siteSubtitleEn || ''
  });
});

// 修改系统标题（需认证）
app.put('/api/site-config', requireAdmin, (req, res) => {
  const config = readConfig();
  const { siteTitle, siteSubtitle, siteTitleEn, siteSubtitleEn } = req.body;
  if (siteTitle !== undefined) config.siteTitle = siteTitle.trim();
  if (siteSubtitle !== undefined) config.siteSubtitle = siteSubtitle.trim();
  if (siteTitleEn !== undefined) config.siteTitleEn = siteTitleEn.trim();
  if (siteSubtitleEn !== undefined) config.siteSubtitleEn = siteSubtitleEn.trim();
  writeConfig(config);
  res.json({ ok: true, siteTitle: config.siteTitle, siteSubtitle: config.siteSubtitle, siteTitleEn: config.siteTitleEn, siteSubtitleEn: config.siteSubtitleEn });
});

// 修改主办方密码（管理员专属）
app.put('/api/organizer-password', requireAdmin, (req, res) => {
  const { newPassword } = req.body;
  if (!newPassword || newPassword.length < 4) {
    return res.status(400).json({ error: '密码长度至少为4位' });
  }
  
  const config = readConfig();
  config.organizerPassword = bcrypt.hashSync(newPassword, SALT_ROUNDS);
  writeConfig(config);
  
  res.json({ ok: true, message: '主办方密码已更新' });
});

// ==================== 认证 API ====================

// 登录 API（带速率限制和密码哈希验证）
app.post('/api/login', (req, res) => {
  const { password, role } = req.body;
  if (!password) {
    return res.status(400).json({ error: '缺少密码参数' });
  }
  if (!role) {
    return res.status(400).json({ error: '请选择登录角色' });
  }

  const config = readConfig();
  let valid = false;
  let loginRole = role;
  
  // 验证密码
  if (role === 'admin' && bcrypt.compareSync(password, config.adminPassword)) {
    valid = true;
  } else if (role === 'organizer' && config.organizerPassword && bcrypt.compareSync(password, config.organizerPassword)) {
    valid = true;
  }
  
  if (valid) {
    const token = crypto.randomBytes(32).toString('hex');
    validTokens.set(token, { createdAt: Date.now(), role: loginRole });
    res.json({ ok: true, token, role: loginRole });
  } else {
    res.status(401).json({ error: '密码错误' });
  }
});

// 验证中间件（带 token 过期检查）
function requireAuth(req, res, next) {
  const token = req.headers['x-auth-token'];
  if (!token) {
    return res.status(401).json({ error: '未登录，请先登录管理后台' });
  }

  const meta = validTokens.get(token);
  if (!meta) {
    return res.status(401).json({ error: '登录已过期，请重新登录' });
  }

  // 检查 token 是否过期
  const config = readConfig();
  const ttl = config.tokenTtl || TOKEN_TTL;
  if (Date.now() - meta.createdAt > ttl) {
    validTokens.delete(token);
    return res.status(401).json({ error: '登录已过期，请重新登录' });
  }

  // 保存角色信息到 request 对象供后续使用
  req.userRole = meta.role;
  return next();
}

// 管理员专属权限中间件
function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.userRole !== 'admin') {
      return res.status(403).json({ error: '权限不足，此功能仅管理员可用' });
    }
    next();
  });
}

// 获取当前用户角色的辅助函数
function getCurrentRole(req) {
  return req.userRole || 'unknown';
}

// ==================== 公开 API ====================

// 服务端生成二维码
app.get('/api/qrcode', async (req, res) => {
  const url = req.query.url;
  if (!url) return res.status(400).json({ error: '缺少 url 参数' });
  try {
    const dataUrl = await QRCode.toDataURL(url, {
      width: 300, margin: 2,
      color: { dark: '#1a56db', light: '#ffffff' }
    });
    res.json({ ok: true, dataUrl });
  } catch (err) {
    res.status(500).json({ error: '生成失败: ' + err.message });
  }
});

// 静态文件（index.html 公开访问）
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html') || filePath.endsWith('.js') || filePath.endsWith('.css')) {
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    }
  }
}));
// 上传文件静态访问
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// 获取所有数据
app.get('/api/data', (req, res) => {
  res.json(readData());
});

// 获取统计数据（轻量，不含详细数据）
app.get('/api/stats', (req, res) => {
  const data = readData();
  const totalSeats = data.venues.reduce((s, v) => s + (v.totalSeats || 0), 0);
  const assignedAttendees = data.attendees.filter(a => a.venueId && a.row && a.seat).length;
  res.json({
    venueCount: data.venues.length,
    totalSeats: totalSeats,
    totalAttendees: assignedAttendees,
    remainingSeats: totalSeats - assignedAttendees
  });
});

// 获取场馆列表（轻量数据，含参会者数量）
app.get('/api/venues', (req, res) => {
  const data = readData();
  res.json(data.venues.map(v => ({
    id: v.id, name: v.name, description: v.description,
    totalSeats: v.totalSeats, rowCount: v.rows.length,
    attendeeCount: data.attendees.filter(a => a.venueId === v.id).length,
    layout: v.layout || v.mode || 'theater'
  })));
});

// 获取单个场馆详情
app.get('/api/venues/:id', (req, res) => {
  const data = readData();
  const venue = data.venues.find(v => v.id === req.params.id);
  if (!venue) return res.status(404).json({ error: '场馆不存在' });
  const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
  res.json({ venue, attendees: venueAttendees });
});

// ==================== 搜索缓存 ====================

// 搜索缓存：key=name, value={ results, version, timestamp }
const searchCache = new Map();
const SEARCH_CACHE_TTL = 30 * 1000; // 30 秒缓存
const SEARCH_CACHE_MAX = 500; // 最多缓存 500 个搜索结果

// 查询统计：记录每次搜索
const searchStats = {
  totalQueries: 0,          // 总查询次数
  uniqueUsers: new Set(),   // 唯一用户（按 IP）
  queryLog: []              // 查询日志（最近 1000 条）
};
const MAX_QUERY_LOG = 1000;

// 定期清理过期缓存
setInterval(() => {
  const now = Date.now();
  for (const [key, entry] of searchCache.entries()) {
    if (now - entry.timestamp > SEARCH_CACHE_TTL) {
      searchCache.delete(key);
    }
  }
  if (searchCache.size > SEARCH_CACHE_MAX) {
    const entries = Array.from(searchCache.entries());
    entries.sort((a, b) => a[1].timestamp - b[1].timestamp);
    for (let i = 0; i < entries.length - SEARCH_CACHE_MAX; i++) {
      searchCache.delete(entries[i][0]);
    }
  }
  // 清理过旧的查询日志（保留最近 1 小时）
  const oneHourAgo = Date.now() - 60 * 60 * 1000;
  searchStats.queryLog = searchStats.queryLog.filter(q => q.timestamp > oneHourAgo);
}, 60 * 1000);

// 搜索座位（跨所有场馆，带缓存和统计）
app.get('/api/search', (req, res) => {
  const rawName = (req.query.name || '').trim().replace(/\s+/g, '');
  if (!rawName) return res.json({ results: [] });

  // 检查缓存
  const cached = searchCache.get(rawName);
  if (cached && cached.version === dataCacheVersion && Date.now() - cached.timestamp < SEARCH_CACHE_TTL) {
    // 缓存命中时，如果需要记录统计也要记录
    const trackStats = req.query.track === '1';
    if (trackStats) {
      const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress || '-';
      searchStats.totalQueries++;
      searchStats.uniqueUsers.add(ip);
      searchStats.queryLog.push({
        name: rawName,
        ip: ip,
        timestamp: Date.now(),
        found: cached.results.length > 0
      });
      if (searchStats.queryLog.length > MAX_QUERY_LOG) {
        searchStats.queryLog = searchStats.queryLog.slice(-MAX_QUERY_LOG);
      }
    }
    return res.json({ results: cached.results, cached: true });
  }

  const data = readData();
  const config = readConfig();
  const results = data.attendees
    .filter(a => {
      const n = (a.name || '').replace(/\s+/g, '');
      const nEn = ((a.nameEn || '').replace(/\s+/g, '')).toLowerCase();
      const q = rawName.toLowerCase();
      return n === rawName || n.includes(rawName) || nEn === q || nEn.includes(q);
    })
    .map(a => {
      const venue = data.venues.find(v => v.id === a.venueId);
      return { ...a, venueName: venue ? venue.name : '未知', venue, siteTitleEn: config.siteTitleEn || '' };
    });

  // 写入缓存
  searchCache.set(rawName, {
    results,
    version: dataCacheVersion,
    timestamp: Date.now()
  });

  // 只在 track=1 时记录查询统计（用户完成输入后的最终查询）
  const trackStats = req.query.track === '1';
  if (trackStats) {
    const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress || '-';
    searchStats.totalQueries++;
    searchStats.uniqueUsers.add(ip);
    searchStats.queryLog.push({
      name: rawName,
      ip: ip,
      timestamp: Date.now(),
      found: results.length > 0
    });
    // 限制日志大小
    if (searchStats.queryLog.length > MAX_QUERY_LOG) {
      searchStats.queryLog = searchStats.queryLog.slice(-MAX_QUERY_LOG);
    }
  }

  res.json({ results, cached: false });
});

// 获取查询统计
app.get('/api/search-stats', requireAdmin, (req, res) => {
  const now = Date.now();
  const oneHourAgo = now - 60 * 60 * 1000;
  const recentQueries = searchStats.queryLog.filter(q => q.timestamp > oneHourAgo);

  // 统计热门搜索词
  const nameCount = {};
  recentQueries.forEach(q => {
    nameCount[q.name] = (nameCount[q.name] || 0) + 1;
  });
  const hotSearches = Object.entries(nameCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name, count]) => ({ name, count }));

  res.json({
    totalQueries: searchStats.totalQueries,
    uniqueUsers: searchStats.uniqueUsers.size,
    lastHourQueries: recentQueries.length,
    hotSearches: hotSearches,
    recentQueries: searchStats.queryLog.slice(-50).reverse()
  });
});

// 数据统计 API
app.get('/api/statistics', requireAdmin, (req, res) => {
  const data = readData();

  // 按单位统计
  const companyStats = {};
  data.attendees.forEach(a => {
    const company = a.company || '未填写';
    if (!companyStats[company]) {
      companyStats[company] = { count: 0, venues: new Set() };
    }
    companyStats[company].count++;
    companyStats[company].venues.add(a.venueId);
  });

  // 转换为数组并排序
  const companyList = Object.entries(companyStats)
    .map(([name, stats]) => ({
      name,
      count: stats.count,
      venueCount: stats.venues.size
    }))
    .sort((a, b) => b.count - a.count);

  // 座位利用率
  const venueUtilization = data.venues.map(v => {
    const attendeeCount = data.attendees.filter(a => a.venueId === v.id).length;
    return {
      venueId: v.id,
      venueName: v.name,
      totalSeats: v.totalSeats,
      attendeeCount: attendeeCount,
      utilizationRate: v.totalSeats > 0 ? ((attendeeCount / v.totalSeats) * 100).toFixed(1) : 0
    };
  });

  res.json({
    companyStats: companyList,
    venueUtilization: venueUtilization,
    summary: {
      totalVenues: data.venues.length,
      totalSeats: data.venues.reduce((s, v) => s + (v.totalSeats || 0), 0),
      totalAttendees: data.attendees.length,
      companies: companyList.length
    }
  });
});

// 导出座位安排表（Excel 格式 - 可视化布局图）
app.get('/api/export-seating', requireAdmin, (req, res) => {
  try {
    const data = readData();
    const wb = XLSX.utils.book_new();

    // 为每个会场创建一个 sheet
    data.venues.forEach(venue => {
      const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
      
      const attendeeMap = {};
      venueAttendees.forEach(a => {
        const key = `${a.row}_${a.seat}`;
        attendeeMap[key] = a.name;
      });

      const sheetData = [];
      
      // 标题行
      sheetData.push([venue.name]);
      sheetData.push([venue.description || '', '', '', '', '', '共 ' + venueAttendees.length + ' 人']);
      sheetData.push([]);

      // 渲染每一排
      venue.rows.forEach(row => {
        const rowLabel = row.label;
        const rowData = [rowLabel];

        row.seatGroups.forEach((group, groupIdx) => {
          group.forEach(seatNum => {
            const key = `${rowLabel}_${seatNum}`;
            const name = attendeeMap[key] || '';
            rowData.push(name);
          });
          if (groupIdx < row.seatGroups.length - 1) {
            rowData.push('');
          }
        });

        sheetData.push(rowData);
      });

      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // 设置列宽
      const maxCols = Math.max(...sheetData.map(r => r.length));
      ws['!cols'] = [];
      for (let i = 0; i < maxCols; i++) {
        ws['!cols'].push({ wch: i === 0 ? 12 : 14 });
      }

      // 设置行高
      ws['!rows'] = [];
      ws['!rows'][0] = { hpt: 40 };
      ws['!rows'][1] = { hpt: 25 };

      XLSX.utils.book_append_sheet(wb, ws, venue.name.substring(0, 31));
    });

    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const filename = '座位安排表.xlsx';
    const encodedFilename = encodeURIComponent(filename).replace(/['()]/g, escape).replace(/\*/g, '%2A');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="seating-layout.xlsx"; filename*=UTF-8''${encodedFilename}`);
    res.send(excelBuffer);
  } catch (err) {
    structuredLog('error', { message: 'Excel 导出失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '导出失败: ' + err.message });
  }
});

// 中文标签转英文（SVG 导出用）
function translateRowLabelSVG(label) {
  if (!label) return '';
  const s = label.trim().replace(/\s+/g, '');
  const cnNumMap = { '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10 };
  function parseCn(s) {
    if (/^\d+$/.test(s)) return parseInt(s);
    if (cnNumMap[s]) return cnNumMap[s];
    let m = s.match(/^十([一二三四五六七八九])$/);
    if (m) return 10 + cnNumMap[m[1]];
    m = s.match(/^([二三四五六七八九])十$/);
    if (m) return cnNumMap[m[1]] * 10;
    m = s.match(/^([二三四五六七八九])十([一二三四五六七八九])$/);
    if (m) return cnNumMap[m[1]] * 10 + cnNumMap[m[2]];
    return s;
  }
  const dirs = { '前':'Front','后':'Back','左':'Left','右':'Right','内前':'Inner Front','内后':'Inner Back' };
  const floors = { '一楼':'1F','二楼':'2F','三楼':'3F','四楼':'4F','五楼':'5F','六楼':'6F','七楼':'7F','八楼':'8F','九楼':'9F','十楼':'10F' };
  // 方向+第X排/列
  let m = s.match(/^(前|后|左|右|内前|内后)第([一二三四五六七八九十]+|\d+)([排列])$/);
  if (m) return dirs[m[1]] + ' ' + (m[3] === '列' ? 'Col' : 'Row') + ' ' + parseCn(m[2]);
  m = s.match(/^(前|后|左|右|内前|内后)(\d+|[一二三四五六七八九十]+)([排列])$/);
  if (m) return dirs[m[1]] + ' ' + (m[3] === '列' ? 'Col' : 'Row') + ' ' + parseCn(m[2]);
  // 楼层+沙发+第X排
  m = s.match(/^(一楼|二楼|三楼|四楼|五楼|六楼|七楼|八楼|九楼|十楼)沙发[第]?([一二三四五六七八九十]+|\d+)[排]$/);
  if (m) return floors[m[1]] + ' Row ' + parseCn(m[2]) + ' (Sofa)';
  // 楼层+第X排
  m = s.match(/^(一楼|二楼|三楼|四楼|五楼|六楼|七楼|八楼|九楼|十楼)第([一二三四五六七八九十]+|\d+)[排]$/);
  if (m) return floors[m[1]] + ' Row ' + parseCn(m[2]);
  // 沙发第X排
  m = s.match(/^沙发[第]?([一二三四五六七八九十]+|\d+)[排]$/);
  if (m) return 'Row ' + parseCn(m[1]) + ' (Sofa)';
  // 第X排
  m = s.match(/^第([一二三四五六七八九十]+|\d+)[排]$/);
  if (m) return 'Row ' + parseCn(m[1]);
  // 第X列
  m = s.match(/^第([一二三四五六七八九十]+|\d+)[列]$/);
  if (m) return 'Column ' + parseCn(m[1]);
  // X排/列（无第）
  m = s.match(/^(\d+|[一二三四五六七八九十]+)[排]$/);
  if (m) return 'Row ' + parseCn(m[1]);
  m = s.match(/^(\d+|[一二三四五六七八九十]+)[列]$/);
  if (m) return 'Column ' + parseCn(m[1]);
  // 桌X → Table X
  m = s.match(/^桌(\d+)$/);
  if (m) return 'Table ' + m[1];
  return label;
}

// 导出座位安排表（SVG 矢量图，适合喷绘）
app.get('/api/export-seating-svg', requireAdmin, (req, res) => {
  try {
    const data = readData();
    const config = readConfig();
    const siteTitle = config.siteTitle || '会议';
    const showEn = !!(config.siteTitleEn && config.siteTitleEn.trim());

    // 双语标签辅助函数：中文 + 英文（等大字体，第二行）
    function bilingualLabel(cnText, x, y, anchor, fontSize, color) {
      let html = `<text x="${x}" y="${y}" text-anchor="${anchor}" fill="${color || '#ef4444'}" font-size="${fontSize}" font-weight="bold">${escXml(cnText)}</text>`;
      if (showEn) {
        const enText = translateRowLabelSVG(cnText);
        const enY = y + fontSize + 6;
        html += `\n  <text x="${x}" y="${enY}" text-anchor="${anchor}" fill="#64748b" font-size="${fontSize}">${escXml(enText)}</text>`;
      }
      return html;
    }

    // 座位参数
    const seatWidth = 200;
    const seatHeight = 80;
    const seatGap = 40;
    const groupGap = 120;
    const rowGap = 50;
    const labelWidth = 120;
    const margin = 40;
    const numBoxWidth = 60;
    const numBoxHeight = 32;

    // 区域颜色（与前端一致）
    const svgRegionColors = [
      { border: '#f59e0b' }, { border: '#22c55e' }, { border: '#ec4899' },
      { border: '#06b6d4' }, { border: '#6366f1' }, { border: '#ef4444' },
      { border: '#84cc16' }, { border: '#f97316' }
    ];

    // 渲染区域标注（单位预留区覆盖层）
    function renderRegionOverlays(venue, seatPositions) {
      let html = '';
      if (!venue.regions || venue.regions.length === 0) return html;
      venue.regions.forEach(function(region, idx) {
        if (!region.company) return;
        const color = svgRegionColors[idx % svgRegionColors.length];
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        let found = false;
        region.seats.forEach(function(seatKey) {
          const pos = seatPositions[seatKey];
          if (pos) {
            found = true;
            if (pos.x < minX) minX = pos.x;
            if (pos.y < minY) minY = pos.y;
            if (pos.x + pos.w > maxX) maxX = pos.x + pos.w;
            if (pos.y + pos.h > maxY) maxY = pos.y + pos.h;
          }
        });
        if (!found) return;
        const padding = 10;
        minX -= padding;
        minY -= padding;
        maxX += padding;
        maxY += padding;
        const overlayW = maxX - minX;
        const overlayH = maxY - minY;
        // 半透明覆盖层
        html += `  <rect x="${minX}" y="${minY}" width="${overlayW}" height="${overlayH}" fill="${color.border}" opacity="0.12" rx="8" stroke="${color.border}" stroke-width="2" stroke-dasharray="6,3"/>\n`;
        // 标签背景
        const labelText = escXml(region.company) + ' - ' + escXml(region.name);
        const labelY = Math.max(0, minY - 8);
        const labelW = Math.max(labelText.length * 13 + 24, 60);
        html += `  <rect x="${minX}" y="${labelY - 28}" width="${labelW}" height="28" fill="${color.border}" rx="4" opacity="0.9"/>\n`;
        html += `  <text x="${minX + 12}" y="${labelY - 10}" fill="#ffffff" font-size="16" font-weight="bold">${labelText}</text>\n`;
        // 区域中心文字
        const centerX = (minX + maxX) / 2;
        const centerY = (minY + maxY) / 2;
        html += `  <text x="${centerX}" y="${centerY}" fill="${color.border}" font-size="22" font-weight="bold" text-anchor="middle" opacity="0.5">${escXml(region.company)} 预留区</text>\n`;
      });
      return html;
    }

    // 辅助函数：计算排宽度
    function getRowWidth(row) {
      let w = 0;
      row.seatGroups.forEach((g, gIdx) => {
        w += g.length * (seatWidth + seatGap);
        if (gIdx < row.seatGroups.length - 1) w += groupGap - seatGap;
      });
      return w;
    }

    // 辅助函数：座位组结构签名
    function getGroupSignature(row) {
      return row.seatGroups.map(g => g.length).join(',');
    }

    // 计算总高度（模拟渲染流程精确计算）
    const stageHeight = 80;
    const aisleHeight = 30;
    let totalHeight = margin; // 顶部边距
    data.venues.forEach(venue => {
      // 舞台 + 标题区域
      totalHeight += stageHeight + 30 + 120 + seatHeight + rowGap;
      if ((venue.layout === 'u-shape' || venue.layout === 'hui-shape') && venue.rows && venue.rows.length > 0) {
        // 位置结构法：以/第.+列/为锚点分拆
        var firstArm = -1, lastArm = -1;
        venue.rows.forEach(function(r, idx) {
          var l = (r.label || '').replace(/\s+/g, '');
          if (/^第.+列$/.test(l)) { if (firstArm < 0) firstArm = idx; lastArm = idx; }
        });
        const uTopRows = firstArm > 0 ? venue.rows.slice(0, firstArm) : [];
        const uArmRows = venue.rows.slice(Math.max(0, firstArm), lastArm + 1);
        const uBottomRows = lastArm >= 0 && lastArm < venue.rows.length - 1 ? venue.rows.slice(lastArm + 1) : [];
        const maxArmLen = Math.max(...uArmRows.map(r => r.seatGroups[0] ? r.seatGroups[0].length : 0), 0);
        // 顶部行高度
        if (uTopRows.length > 0) {
          totalHeight += uTopRows.length * (seatHeight + 30) + 30;
        }
        totalHeight += 50; // 列标签行
        totalHeight += maxArmLen * (seatHeight + rowGap); // 两臂座位
        if (uBottomRows.length > 0) {
          totalHeight += 30 + uBottomRows.length * (seatHeight + 30); // 底部行
        }
        totalHeight += 60; // 底部间距
      } else if (venue.layout === 'banquet' && venue.rows && venue.rows.length > 0) {
        // 宴会桌高度：上座位 + 间距 + 桌子 + 间距 + 下座位 + 行间距
        const bTableH = seatHeight + 12 + 40 + 12 + seatHeight + rowGap;
        venue.rows.forEach(() => { totalHeight += bTableH; });
      } else if (venue.rows && venue.rows.length > 0) {
        // 普通渲染方式的高度计算：每行 = seatHeight + rowGap + 可能的过道
        venue.rows.forEach(row => {
          totalHeight += seatHeight + rowGap;
          if (row.hasAisleAfter) {
            totalHeight += 20 + 30 + 30; // 间隙 + 过道高 + 间隙
          }
        });
      }
      totalHeight += 60; // 会场间距
    });
    totalHeight += margin; // 底部边距

    // 计算最大宽度
    let maxWidth = 0;
    data.venues.forEach(venue => {
      if ((venue.layout === 'u-shape' || venue.layout === 'hui-shape') && venue.rows) {
         // 位置结构法：以/第.+列/为锚点分拆
         var firstArmW = -1, lastArmW = -1;
         venue.rows.forEach(function(r, idx) {
           var l = (r.label || '').replace(/\s+/g, '');
           if (/^第.+列$/.test(l)) { if (firstArmW < 0) firstArmW = idx; lastArmW = idx; }
         });
         const uTopW = firstArmW > 0 ? venue.rows.slice(0, firstArmW) : [];
         const uArmW = venue.rows.slice(Math.max(0, firstArmW), lastArmW + 1);
         const uBottomW = lastArmW >= 0 && lastArmW < venue.rows.length - 1 ? venue.rows.slice(lastArmW + 1) : [];
         const colGap = 30;
         const armBottomGap = 60;
         const _midIdx = Math.floor(uArmW.length / 2);
         const leftColCount = _midIdx;
         const rightColCount = uArmW.length - _midIdx;
         const lWidth = leftColCount * seatWidth + Math.max(0, leftColCount - 1) * colGap;
         const rWidth = rightColCount * seatWidth + Math.max(0, rightColCount - 1) * colGap;
         const tSeats = uTopW.length > 0 ? Math.max(...uTopW.map(r => (r.seatGroups[0] || []).length)) : 0;
         const tWidth = tSeats * (seatWidth + seatGap) - (tSeats > 0 ? seatGap : 0);
         const bSeatsW = uBottomW.length > 0 ? Math.max(...uBottomW.map(r => (r.seatGroups[0] || []).length)) : 0;
         const bWidth = bSeatsW * (seatWidth + seatGap) - (bSeatsW > 0 ? seatGap : 0);
         const hWidth = Math.max(tWidth, bWidth);
         const uTotal = lWidth + (hWidth > 0 ? armBottomGap + hWidth + armBottomGap : 0) + rWidth;
         maxWidth = Math.max(maxWidth, uTotal);
      } else if (venue.rows) {
        // 新方式：计算所有排的最大组数和每组最大座位数
        let vMaxGroupCount = 0;
        const vMaxSeatsPerGroup = {};
        venue.rows.forEach(row => {
          if (row.seatGroups) {
            vMaxGroupCount = Math.max(vMaxGroupCount, row.seatGroups.length);
            row.seatGroups.forEach((g, gi) => {
              vMaxSeatsPerGroup[gi] = Math.max(vMaxSeatsPerGroup[gi] || 0, g.length);
            });
          }
        });
        let vTotalSeatWidth = 0;
        for (let gi = 0; gi < vMaxGroupCount; gi++) {
          const gs = vMaxSeatsPerGroup[gi] || 0;
          vTotalSeatWidth += gs * (seatWidth + seatGap);
          if (gi < vMaxGroupCount - 1) vTotalSeatWidth += groupGap - seatGap;
        }
        maxWidth = Math.max(maxWidth, vTotalSeatWidth);
      }
    });

    const labelSpace = 160;
    const expandSpace = 80;
    let svgWidth = maxWidth + labelSpace * 2 + expandSpace * 2 + margin * 2;

    let svgContent = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${totalHeight}" viewBox="0 0 ${svgWidth} ${totalHeight}">
  <defs>
    <style>
      .main-title { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 64px; font-weight: bold; fill: #1e293b; }
      .subtitle { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 32px; fill: #64748b; }
      .row-label { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 30px; fill: #475569; }
      .seat-name { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; fill: #1e293b; font-weight: 600; }
      .seat-num { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 24px; fill: #ffffff; font-weight: bold; }
      .stage-label { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 48px; font-weight: bold; fill: #ffffff; }
      .aisle-label { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 28px; fill: #94a3b8; font-weight: 600; }
    </style>
  </defs>
  <rect width="100%" height="100%" fill="#ffffff"/>
`;

    let y = margin;

    data.venues.forEach((venue) => {
      const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
      const attendeeMap = {};
      venueAttendees.forEach(a => {
        attendeeMap[a.row + '_' + a.seat] = a.name;
      });
      const regionSeatPositions = {}; // 记录座位坐标用于区域标注

      // 标题（会场名称在最上方）
      svgContent += `  <text x="${svgWidth / 2}" y="${y + 64}" class="main-title" text-anchor="middle">${escXml(venue.name)}</text>\n`;
      const totalSeats = (venue.rows || []).reduce((sum, r) => sum + (r.seatGroups || []).reduce((s, g) => s + g.length, 0), 0);
      svgContent += `  <text x="${svgWidth / 2}" y="${y + 108}" class="subtitle" text-anchor="middle">共 ${totalSeats} 个座位</text>\n`;
      y += 120 + 30;

      // 舞台（在标题下方）- 自定义布局会场不显示默认舞台
      if (!(venue.layout === 'custom' && venue.customRows && venue.customRows.length > 0)) {
        const stageHeight = 80;
        svgContent += `  <rect x="${margin}" y="${y}" width="${svgWidth - margin * 2}" height="${stageHeight}" fill="#1a56db" rx="10"/>\n`;
        svgContent += `  <text x="${svgWidth / 2}" y="${y + 56}" class="stage-label" text-anchor="middle">${escXml(venue.stageName || '舞台区域')}</text>\n`;
        y += stageHeight + seatHeight + rowGap;
      }

      if (!venue.rows || venue.rows.length === 0) { y += 50; return; }

      // 自定义布局SVG渲染：使用实际位置
      if (venue.layout === 'custom' && venue.customRows && venue.customRows.length > 0) {
        const cr = venue.customRows;
        const sorted = [...cr].sort((a, b) => a.rowNum - b.rowNum);
        
        // 计算画布边界（包含舞台和大门）
        const cw = venue.canvasWidth || 800;
        const ch = venue.canvasHeight || 600;
        
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        sorted.forEach(row => {
          // 为行标签留出空间：横排标签在左侧，竖排标签在上方
          const extraLeft = (row.direction === 'vertical') ? 120 : 60;
          const extraTop = (row.direction === 'vertical') ? 60 : 0;
          minX = Math.min(minX, (row.x || 0) - extraLeft);
          minY = Math.min(minY, (row.y || 0) - extraTop);
          maxX = Math.max(maxX, (row.x || 0) + (row.width || 100));
          maxY = Math.max(maxY, (row.y || 0) + (row.height || 50));
        });
        
        // 加入舞台和大门到边界计算
        if (venue.customStage && Object.keys(venue.customStage).length > 0) {
          const stage = venue.customStage;
          minX = Math.min(minX, stage.x || 0);
          minY = Math.min(minY, stage.y || 0);
          maxX = Math.max(maxX, (stage.x || 0) + (stage.width || 200));
          maxY = Math.max(maxY, (stage.y || 0) + (stage.height || 80));
        }
        if (venue.customGates && venue.customGates.length > 0) {
          venue.customGates.forEach(gate => {
            minX = Math.min(minX, gate.x || 0);
            minY = Math.min(minY, gate.y || 0);
            maxX = Math.max(maxX, (gate.x || 0) + (gate.width || 80));
            maxY = Math.max(maxY, (gate.y || 0) + (gate.height || 120));
          });
        }
        if (venue.customAisles && venue.customAisles.length > 0) {
          venue.customAisles.forEach(aisle => {
            minX = Math.min(minX, aisle.x || 0);
            minY = Math.min(minY, aisle.y || 0);
            maxX = Math.max(maxX, (aisle.x || 0) + (aisle.width || 60));
            maxY = Math.max(maxY, (aisle.y || 0) + (aisle.height || 40));
          });
        }
        
        const contentW = maxX - minX;
        const contentH = maxY - minY;
        
        // 计算缩放：综合考虑画布尺寸和座位密度
        const svgMargin = 200;
        const canvasScale = Math.min(2500 / Math.max(contentW, 1), 2000 / Math.max(contentH, 1));
        
        // 基于座位密度计算最小缩放：确保每个座位至少 120px 宽/高
        let minOrigSeatPx = Infinity;
        sorted.forEach(row => {
          const size = row.direction === 'horizontal' 
            ? (row.width || 100) / Math.max(row.seatCount, 1)
            : (row.height || 50) / Math.max(row.seatCount, 1);
          minOrigSeatPx = Math.min(minOrigSeatPx, size);
        });
        const MIN_SEAT_PX = 120;
        const seatScale = Math.min(MIN_SEAT_PX / Math.max(minOrigSeatPx, 1), 15); // 最大15倍
        const scale = Math.max(canvasScale, seatScale);
        
        const baseX = svgMargin - minX * scale;
        const baseY = y - minY * scale;
        
        // === 1. 先渲染舞台和大门 ===
        if (venue.customStage && Object.keys(venue.customStage).length > 0) {
          const stage = venue.customStage;
          const stageX = baseX + (stage.x || 0) * scale;
          const stageY = baseY + (stage.y || 0) * scale;
          const stageW = (stage.width || 200) * scale;
          const stageH = (stage.height || 80) * scale;
          svgContent += `  <rect x="${stageX}" y="${stageY}" width="${stageW}" height="${stageH}" fill="#1a56db" rx="8"/>\n`;
          svgContent += `  <text x="${stageX + stageW/2}" y="${stageY + stageH/2 + 10}" text-anchor="middle" fill="#ffffff" font-size="${Math.min(48, stageH/2.2)}" font-weight="bold">${escXml(stage.label || '舞台')}</text>\n`;
        }
        
        if (venue.customGates && venue.customGates.length > 0) {
          venue.customGates.forEach(gate => {
            const gateX = baseX + (gate.x || 0) * scale;
            const gateY = baseY + (gate.y || 0) * scale;
            const gateW = (gate.width || 80) * scale;
            const gateH = (gate.height || 120) * scale;
            svgContent += `  <rect x="${gateX}" y="${gateY}" width="${gateW}" height="${gateH}" fill="#10b981" stroke="#059669" stroke-width="3" rx="6"/>\n`;
            svgContent += `  <text x="${gateX + gateW/2}" y="${gateY + gateH/2 + 10}" text-anchor="middle" fill="#ffffff" font-size="${Math.min(36, gateH/3)}" font-weight="bold">${escXml(gate.label || '门')}</text>\n`;
          });
        }
        
        if (venue.customAisles && venue.customAisles.length > 0) {
          venue.customAisles.forEach(aisle => {
            const aisleX = baseX + (aisle.x || 0) * scale;
            const aisleY = baseY + (aisle.y || 0) * scale;
            const aisleW = (aisle.width || 60) * scale;
            const aisleH = (aisle.height || 40) * scale;
            svgContent += `  <rect x="${aisleX}" y="${aisleY}" width="${aisleW}" height="${Math.min(aisleH, 24)}" fill="#dbeafe" rx="3"/>\n`;
            const aisleDisplayLabel = (aisle.label && aisle.label !== 'null') ? aisle.label : '过道';
            svgContent += `  <text x="${aisleX + aisleW/2}" y="${aisleY + Math.min(aisleH, 24)/2 + 6}" text-anchor="middle" fill="#94a3b8" font-size="${Math.min(20, Math.min(aisleH, 24)/1.6)}" font-weight="bold">${escXml(aisleDisplayLabel)}</text>\n`;
          });
        }
        
        // === 2. 再渲染座位 ===
        // 构建分组：连续且seatCount相同的排为一组
        const rowGroups = [];
        let curGroup = [sorted[0]];
        for (let g = 1; g < sorted.length; g++) {
          if (sorted[g].seatCount === sorted[g-1].seatCount) {
            curGroup.push(sorted[g]);
          } else {
            rowGroups.push(curGroup);
            curGroup = [sorted[g]];
          }
        }
        if (curGroup.length > 0) rowGroups.push(curGroup);
        
        rowGroups.forEach(group => {
          group.forEach((row, rowIdxInGroup) => {
            const isFirstInGroup = (rowIdxInGroup === 0);
            const rowLabel = row.label || '';
            const x = baseX + (row.x || 0) * scale;
            const rowY = baseY + (row.y || 0) * scale;
            const w = (row.width || 100) * scale;
            const h = (row.height || 50) * scale;
            const getSn = (idx) => (row.seatNumbers && row.seatNumbers.length > idx) ? row.seatNumbers[idx] : (row.startSeat || 1) + idx;
            const aisles = (row.aislePositions || []).slice().sort((a, b) => a - b);
            const aisleGapPx = Math.max(30, (row.direction === 'horizontal' ? w : h) * 0.05);
            const totalAisleGap = aisles.length * aisleGapPx;

            if (row.direction === 'horizontal') {
              const seatCount = row.seatCount;
              const availW = w - totalAisleGap;
              const seatW = availW / seatCount;
              const seatH = Math.max(seatHeight, h * 0.9);

              let cursorX = x;
              for (let i = 0; i < seatCount; i++) {
                for (let a = 0; a < aisles.length; a++) {
                  if (aisles[a] === i) {
                    svgContent += `  <rect x="${cursorX}" y="${rowY + h/2 - 12}" width="${aisleGapPx}" height="24" fill="#dbeafe" rx="3"/>\n`;
                    svgContent += `  <text x="${cursorX + aisleGapPx/2}" y="${rowY + h/2 + 6}" text-anchor="middle" fill="#94a3b8" font-size="${Math.min(16, 24/1.6)}" font-weight="bold">过道</text>\n`;
                    cursorX += aisleGapPx;
                  }
                }
                const sx = cursorX;
                const sy = rowY + (h - seatH) / 2;
                const sn = getSn(i);
                const name = attendeeMap[rowLabel + '_' + sn];

                if (i === 0) {
                  const lastSx = x + w - seatW;
                  svgContent += `  ` + bilingualLabel(rowLabel, sx - 15, sy + seatH/2 + 8, 'end', Math.max(28, Math.min(40, h/1.8)), '#ef4444') + `\n`;
                  svgContent += `  ` + bilingualLabel(rowLabel, lastSx + seatW + 15, sy + seatH/2 + 8, 'start', Math.max(28, Math.min(40, h/1.8)), '#ef4444') + `\n`;
                }

                // 座位号蓝色方块（仅组内第一排显示）
                if (isFirstInGroup) {
                  const numBoxW = Math.min(80, seatW * 0.85);
                  const numBoxH = 40;
                  const nbx = sx + (seatW - numBoxW) / 2;
                  const nby = sy - numBoxH - 10;
                  svgContent += `  <rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="6"/>\n`;
                  svgContent += `  <text x="${nbx + numBoxW/2}" y="${nby + 28}" class="seat-num" text-anchor="middle" font-size="${Math.min(28, numBoxH/1.6)}">${sn}</text>\n`;
                }

                svgContent += `  <rect x="${sx + 3}" y="${sy}" width="${seatW - 6}" height="${seatH}" fill="#ffffff" stroke="#cbd5e1" stroke-width="3" rx="6"/>\n`;
                if (rowLabel && sn) regionSeatPositions[rowLabel + '_' + sn] = { x: sx + 3, y: sy, w: seatW - 6, h: seatH };

                if (name) {
                  const dn = name.replace(/\n/g, ' ');
                  const nameLen = Math.max(dn.length, 1);
                  const maxFit = Math.min(seatW * 0.85 / nameLen, seatH * 0.5);
                  const fs = Math.max(12, Math.min(36, maxFit));
                  svgContent += `  <text x="${sx + seatW/2}" y="${sy + seatH/2 + fs/3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
                }
                cursorX += seatW;
              }
            } else {
              const seatCount = row.seatCount;
              const availH = h - totalAisleGap;
              const seatH = availH / seatCount;
              const seatW = Math.max(seatWidth, w * 0.9);

              let cursorY = rowY;
              for (let i = 0; i < seatCount; i++) {
                for (let a = 0; a < aisles.length; a++) {
                  if (aisles[a] === i) {
                    svgContent += `  <rect x="${x + w/2 - 12}" y="${cursorY}" width="24" height="${aisleGapPx}" fill="#dbeafe" rx="3"/>\n`;
                    svgContent += `  <text x="${x + w/2}" y="${cursorY + aisleGapPx/2 + 6}" text-anchor="middle" fill="#94a3b8" font-size="${Math.min(16, 24/1.6)}" font-weight="bold">过道</text>\n`;
                    cursorY += aisleGapPx;
                  }
                }
                const sx = x + (w - seatW) / 2;
                const sy = cursorY;
                const sn = getSn(i);
                const name = attendeeMap[rowLabel + '_' + sn];

                if (i === 0) {
                  svgContent += `  ` + bilingualLabel(rowLabel, sx + seatW/2, sy - 15, 'middle', Math.max(28, Math.min(40, h/1.8)), '#ef4444') + `\n`;
                }

                // 座位号蓝色方块（仅组内第一排显示）
                if (isFirstInGroup) {
                  const numBoxW = 60;
                  const numBoxH = Math.min(45, seatH * 0.85);
                  const nbx = sx - numBoxW - 10;
                  const nby = sy + (seatH - numBoxH) / 2;
                  svgContent += `  <rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="6"/>\n`;
                  svgContent += `  <text x="${nbx + numBoxW/2}" y="${nby + numBoxH/2 + 8}" class="seat-num" text-anchor="middle" font-size="${Math.min(28, numBoxH/1.6)}">${sn}</text>\n`;
                }

                svgContent += `  <rect x="${sx}" y="${sy + 3}" width="${seatW}" height="${seatH - 6}" fill="#ffffff" stroke="#cbd5e1" stroke-width="3" rx="6"/>\n`;
                if (rowLabel && sn) regionSeatPositions[rowLabel + '_' + sn] = { x: sx, y: sy + 3, w: seatW, h: seatH - 6 };

                if (name) {
                  const dn = name.replace(/\n/g, ' ');
                  const nameLen = Math.max(dn.length, 1);
                  const maxFit = Math.min(seatW * 0.85 / nameLen, seatH * 0.5);
                  const fs = Math.max(12, Math.min(36, maxFit));
                  svgContent += `  <text x="${sx + seatW/2}" y="${sy + seatH/2 + fs/3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
                }
                cursorY += seatH;
              }
            }
          });
        });
        
        y = baseY + maxY * scale + 150;
        
        // 确保SVG画布足够大以容纳自定义布局的缩放内容
        const customMaxW = svgMargin + maxX * scale + svgMargin;
        if (customMaxW > svgWidth) svgWidth = Math.ceil(customMaxW);
        if (y > totalHeight) totalHeight = Math.ceil(y);
        // 更新SVG头部的尺寸和viewBox
        svgContent = svgContent.replace(
          /^<svg [^>]*>/m,
          `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${totalHeight}" viewBox="0 0 ${svgWidth} ${totalHeight}">`
        );
        svgContent += renderRegionOverlays(venue, regionSeatPositions);
        return;
      }

      // 宴会桌式SVG渲染：每桌独立显示，上3下3座位，不显示座位号
      if (venue.layout === 'banquet') {
        const tSeatW = seatWidth;       // 200
        const tSeatH = seatHeight;      // 80
        const tSeatGap = seatGap;       // 40
        const tableBodyW = tSeatW * 3 + tSeatGap * 2;  // 680
        const tableBodyH = 40;
        const tableGap = 100;           // 桌间水平间距
        const seatTableGap = 12;        // 座位与桌子间距
        const tableUnitH = tSeatH + seatTableGap + tableBodyH + seatTableGap + tSeatH;

        let maxRowWidth = 0;
        venue.rows.forEach(row => {
          const groups = row.seatGroups || [];
          if (groups.length === 0) return;
          const rw = groups.length * tableBodyW + (groups.length - 1) * tableGap;
          maxRowWidth = Math.max(maxRowWidth, rw);
        });
        svgWidth = Math.max(svgWidth, maxRowWidth + 500, 1300);

        venue.rows.forEach(row => {
          const groups = row.seatGroups || [];
          if (groups.length === 0) { y += 30; return; }

          const tablesPerRow = groups.length;
          const rowWidth = tablesPerRow * tableBodyW + (tablesPerRow - 1) * tableGap;
          const rowStartX = Math.max(200, (svgWidth - rowWidth) / 2);
          const rowLabel = row.label || '';

          // 排标签（左右两侧）
          const labelY = y + tableUnitH / 2 + 6;
          svgContent += `  ` + bilingualLabel(rowLabel, rowStartX - 16, labelY, 'end', 36, '#ef4444') + `\n`;
          svgContent += `  ` + bilingualLabel(rowLabel, rowStartX + rowWidth + 16, labelY, 'start', 36, '#ef4444') + `\n`;

          groups.forEach((group, gi) => {
            const tx = rowStartX + gi * (tableBodyW + tableGap);

            // 上排3个座位（座位1,2,3），不显示座位号
            for (let si = 0; si < 3 && si < group.length; si++) {
              const seatNum = group[si];
              const sx = tx + si * (tSeatW + tSeatGap);
              const topSeatY = y;
              const name = attendeeMap[rowLabel + '_' + seatNum];
              const fill = name ? '#dbeafe' : '#ffffff';

              svgContent += `  <rect x="${sx}" y="${topSeatY}" width="${tSeatW}" height="${tSeatH}" fill="${fill}" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
              regionSeatPositions[rowLabel + '_' + seatNum] = { x: sx, y: topSeatY, w: tSeatW, h: tSeatH };
              if (name) {
                const dn = name.replace(/\n/g, ' ');
                const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                svgContent += `  <text x="${sx + tSeatW / 2}" y="${topSeatY + tSeatH / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
              }
            }

            // 桌子本体
            const tby = y + tSeatH + seatTableGap;
            svgContent += `  <rect x="${tx}" y="${tby}" width="${tableBodyW}" height="${tableBodyH}" fill="#fef3c7" stroke="#f59e0b" stroke-width="3" rx="8"/>\n`;
            const tableNum = (row.tableNums || [])[gi] || 0;
            svgContent += `  <text x="${tx + tableBodyW / 2}" y="${tby + tableBodyH / 2 + 8}" text-anchor="middle" fill="#92400e" font-size="28" font-weight="bold">桌${tableNum}` + (showEn ? '  Table ' + tableNum : '') + `</text>\n`;

            // 下排3个座位（座位4,5,6），不显示座位号
            const bottomSeatY = tby + tableBodyH + seatTableGap;
            for (let si = 3; si < 6 && si < group.length; si++) {
              const seatNum = group[si];
              const sx = tx + (si - 3) * (tSeatW + tSeatGap);
              const name = attendeeMap[rowLabel + '_' + seatNum];
              const fill = name ? '#dbeafe' : '#ffffff';

              svgContent += `  <rect x="${sx}" y="${bottomSeatY}" width="${tSeatW}" height="${tSeatH}" fill="${fill}" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
              regionSeatPositions[rowLabel + '_' + seatNum] = { x: sx, y: bottomSeatY, w: tSeatW, h: tSeatH };
              if (name) {
                const dn = name.replace(/\n/g, ' ');
                const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                svgContent += `  <text x="${sx + tSeatW / 2}" y="${bottomSeatY + tSeatH / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
              }
            }
          });

          y += tableUnitH + rowGap;
        });
        svgWidth = Math.max(svgWidth, maxRowWidth + 500, 1300);
        svgContent += renderRegionOverlays(venue, regionSeatPositions);
        return;
      }

      // U型/回字型会场SVG渲染
      if (venue.layout === 'u-shape' || venue.layout === 'hui-shape') {
        // 位置结构法：以/^第.+列$/为锚点分拆行（顶部行 | 臂列 | 底部行）
        var firstArmSv = -1, lastArmSv = -1;
        venue.rows.forEach(function(r, idx) {
          var l = (r.label || '').replace(/\s+/g, '');
          if (/^第.+列$/.test(l)) { if (firstArmSv < 0) firstArmSv = idx; lastArmSv = idx; }
        });
        const uTopRows = firstArmSv > 0 ? venue.rows.slice(0, firstArmSv) : [];
        const uArmRows = venue.rows.slice(Math.max(0, firstArmSv), lastArmSv + 1);
        const uBottomRows = lastArmSv >= 0 && lastArmSv < venue.rows.length - 1 ? venue.rows.slice(lastArmSv + 1) : [];
        const midIdx = Math.floor(uArmRows.length / 2);
        const leftArmRows = uArmRows.slice(0, midIdx);
        const rightArmRows = uArmRows.slice(midIdx);

        const maxTopSeats = uTopRows.length > 0 ? Math.max(...uTopRows.map(r => (r.seatGroups[0] || []).length)) : 0;
        const maxBottomSeats = uBottomRows.length > 0 ? Math.max(...uBottomRows.map(r => (r.seatGroups[0] || []).length)) : 0;
        const hSeatCount = Math.max(maxTopSeats, maxBottomSeats);
        const maxArmLen = Math.max(...uArmRows.map(r => r.seatGroups[0] ? r.seatGroups[0].length : 0), 0);

        // 列间距需容纳座位号牌（numBoxWidth 60 + 左侧16px间隙 + 余量）
        const colGap = Math.max(30, numBoxWidth + 16 + 24);  // 100
        const labelRowHeight = 50;

        // 计算左臂宽度（各列宽度累加）
        const leftArmWidth = leftArmRows.reduce((sum, col) => sum + seatWidth + colGap, 0) - (leftArmRows.length > 0 ? colGap : 0);
        // 计算右臂宽度
        const rightArmWidth = rightArmRows.reduce((sum, col) => sum + seatWidth + colGap, 0) - (rightArmRows.length > 0 ? colGap : 0);
        // 水平行（顶部/底部）宽度（取最大值）
        const hWidth = hSeatCount > 0 ? hSeatCount * (seatWidth + seatGap) - seatGap : 0;
        // 水平行和臂之间的间距（需容纳右臂内侧座位号牌）
        const armBottomGap = Math.max(60, numBoxWidth + 16 + 24);  // 100
        // 总宽度
        const totalWidth = leftArmWidth + (hWidth > 0 ? armBottomGap + hWidth + armBottomGap : 0) + rightArmWidth;
        
        // 整体居中偏移
        const baseX = Math.max(0, (svgWidth - totalWidth) / 2);
        
        // 各区域起始X位置
        const leftStartX = baseX;
        const bottomStartX = leftStartX + leftArmWidth + (leftArmWidth > 0 && hWidth > 0 ? armBottomGap : 0);
        const rightStartX = (hWidth > 0 ? bottomStartX + hWidth + armBottomGap : leftStartX + leftArmWidth + armBottomGap);

        // 计算各列在每个垂直位置的X坐标
        const leftColXs = [];
        let lx = leftStartX;
        leftArmRows.forEach(() => {
          leftColXs.push(lx);
          lx += seatWidth + colGap;
        });

        const rightColXs = [];
        let rx = rightStartX;
        rightArmRows.forEach(() => {
          rightColXs.push(rx);
          rx += seatWidth + colGap;
        });

       // ---- 0. 顶部行（回字型顶部封闭，水平排列于左右臂间）----
       // 按座位数分组——连续相同座位数的行为一组，只在每组首行显示座位号
        if (uTopRows.length > 0) {
          const topRowGroups = [];
          uTopRows.forEach((row) => {
            const sc = (row.seatGroups[0] || []).filter(x => x !== null && x !== undefined).length;
            const prev = topRowGroups[topRowGroups.length - 1];
            if (prev && prev.seatCount === sc) { prev.rows.push(row); }
            else { topRowGroups.push({ seatCount: sc, rows: [row] }); }
          });
          topRowGroups.forEach(group => {
            group.rows.forEach((topRow, ri) => {
              const topSeats = topRow.seatGroups[0] || [];
              if (topSeats.length > 0) {
                const topLabel = topRow.label || '';
                const topRowWidth = topSeats.length * (seatWidth + seatGap) - seatGap;
                const hAreaCenter = bottomStartX + hWidth / 2;
                let tx = hAreaCenter - topRowWidth / 2;
                svgContent += `  ` + bilingualLabel(topLabel, tx - 16, y + 24, 'end', 36, '#ef4444') + `\n`;
                const isFirstInGroup = ri === 0;
                topSeats.forEach((sn) => {
                  const name = attendeeMap[topRow.label + '_' + sn];
                  if (isFirstInGroup) {
                    const numBoxX = tx + (seatWidth - numBoxWidth) / 2;
                    const numBoxY = y - numBoxHeight - 8;
                    svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                    svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${sn}</text>\n`;
                  }
                  svgContent += `  <rect x="${tx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
                  if (topRow.label && sn) regionSeatPositions[topRow.label + '_' + sn] = { x: tx, y, w: seatWidth, h: seatHeight };
                  if (name) {
                    const dn = name.replace(/\n/g, ' ');
                    const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                    svgContent += `  <text x="${tx + seatWidth / 2}" y="${y + seatHeight / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
                  }
                  tx += seatWidth + seatGap;
                });
              }
              y += seatHeight + 30;
            });
          });
          y += 30;
        }

        // ---- 1. 列标签行 ----
        // 按座位数列分组——座位数相同的相邻列归一组，组内只在第一列显示座位号
        const armGroupFirstCols = (armRows) => {
          const groups = [];
          armRows.forEach((row) => {
            const sc = (row.seatGroups[0] || []).filter(x => x !== null && x !== undefined).length;
            const prev = groups[groups.length - 1];
            if (prev && prev.seatCount === sc) { prev.rows.push(row); }
            else { groups.push({ seatCount: sc, rows: [row] }); }
          });
          const firstSet = new Set();
          let ci = 0;
          groups.forEach(g => { firstSet.add(ci); ci += g.rows.length; });
          return firstSet;
        };
        const leftFirstInGroupCols = armGroupFirstCols(leftArmRows);
        const rightFirstInGroupCols = armGroupFirstCols(rightArmRows);

        const labelY = y;
        leftArmRows.forEach((col, idx) => {
          const cx = leftColXs[idx] + seatWidth / 2;
          svgContent += `  ` + bilingualLabel(col.label, cx, labelY, 'middle', 36, '#ef4444') + `\n`;
        });
        rightArmRows.forEach((col, idx) => {
          const cx = rightColXs[idx] + seatWidth / 2;
          svgContent += `  ` + bilingualLabel(col.label, cx, labelY, 'middle', 36, '#ef4444') + `\n`;
        });
        y += labelRowHeight;

        // ---- 2. 渲染两臂座位（垂直排列）----
        for (let si = 0; si < maxArmLen; si++) {
          // 左臂：每列的第si个座位
          leftArmRows.forEach((col, idx) => {
            const seats = col.seatGroups[0] || [];
            if (si < seats.length) {
              const sn = seats[si];
              const name = attendeeMap[col.label + '_' + sn];
              const sx = leftColXs[idx];

              if (leftFirstInGroupCols.has(idx)) {
                // 座位号放在座位框左侧16px
                const numBoxX = sx - numBoxWidth - 16;
                const numBoxY = y + (seatHeight - numBoxHeight) / 2;
                svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${sn}</text>\n`;
              }

              // 姓名框
              svgContent += `  <rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
              if (col.label && sn) regionSeatPositions[col.label + '_' + sn] = { x: sx, y, w: seatWidth, h: seatHeight };
              if (name) {
                const dn = name.replace(/\n/g, ' ');
                const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                svgContent += `  <text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
              }
            }
          });

          // 右臂：每列的第si个座位
          rightArmRows.forEach((col, idx) => {
            const seats = col.seatGroups[0] || [];
            if (si < seats.length) {
              const sn = seats[si];
              const name = attendeeMap[col.label + '_' + sn];
              const sx = rightColXs[idx];

              if (rightFirstInGroupCols.has(idx)) {
                // 右臂座位号放座位框右侧（外侧），避免与中间区域重叠
                const numBoxX = sx + seatWidth + 16;
                const numBoxY = y + (seatHeight - numBoxHeight) / 2;
                svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${sn}</text>\n`;
              }

              // 姓名框
              svgContent += `  <rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
              if (col.label && sn) regionSeatPositions[col.label + '_' + sn] = { x: sx, y, w: seatWidth, h: seatHeight };
              if (name) {
                const dn = name.replace(/\n/g, ' ');
                const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                svgContent += `  <text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
              }
            }
          });

          y += seatHeight + rowGap;
        }

        // ---- 3. 底部行（支持多行，水平排列于左右臂间）----
       // 按座位数分组——连续相同座位数的行为一组，只在每组首行显示座位号
        if (uBottomRows.length > 0) {
          y += 30;
          const bottomRowGroups = [];
          uBottomRows.forEach((row) => {
            const sc = (row.seatGroups[0] || []).filter(x => x !== null && x !== undefined).length;
            const prev = bottomRowGroups[bottomRowGroups.length - 1];
            if (prev && prev.seatCount === sc) { prev.rows.push(row); }
            else { bottomRowGroups.push({ seatCount: sc, rows: [row] }); }
          });
          bottomRowGroups.forEach(group => {
            group.rows.forEach((botRow, ri) => {
              const botSeats = botRow.seatGroups[0] || [];
              if (botSeats.length > 0) {
                const botLabel = botRow.label || '';
                const botRowWidth = botSeats.length * (seatWidth + seatGap) - seatGap;
                const hAreaCenter = bottomStartX + hWidth / 2;
                let bx = hAreaCenter - botRowWidth / 2;
                svgContent += `  ` + bilingualLabel(botLabel, bx - 16, y - 10, 'end', 36, '#ef4444') + `\n`;
                const isFirstInGroup = ri === 0;
                botSeats.forEach((sn) => {
                  const name = attendeeMap[botRow.label + '_' + sn];
                  if (isFirstInGroup) {
                    const numBoxX = bx + (seatWidth - numBoxWidth) / 2;
                    const numBoxY = y - numBoxHeight - 8;
                    svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                    svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${sn}</text>\n`;
                  }
                  svgContent += `  <rect x="${bx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
                  if (botRow.label && sn) regionSeatPositions[botRow.label + '_' + sn] = { x: bx, y, w: seatWidth, h: seatHeight };
                  if (name) {
                    const dn = name.replace(/\n/g, ' ');
                    const fs = dn.length > 6 ? 28 : dn.length > 4 ? 34 : 40;
                    svgContent += `  <text x="${bx + seatWidth / 2}" y="${y + seatHeight / 2 + fs / 3}" class="seat-name" text-anchor="middle" font-size="${fs}">${escXml(dn)}</text>\n`;
                  }
                  bx += seatWidth + seatGap;
                });
              }
              y += seatHeight + 30;
            });
          });
        }

        y += 60;
        // 右臂座位号在右侧溢出，确保 SVG 足够宽
        const rightOverflow = rightArmRows.length > 0 ? numBoxWidth + 16 : 0;
        if (totalWidth + rightOverflow > svgWidth) {
          svgWidth = Math.ceil(totalWidth + rightOverflow + 80);
          svgContent = svgContent.replace(
            /^<svg [^>]*>/m,
            `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${totalHeight}" viewBox="0 0 ${svgWidth} ${totalHeight}">`
          );
        }
        svgContent += renderRegionOverlays(venue, regionSeatPositions);
        return;
      }

      // 按分组渲染座位图
      const isTheaterMode = venue.layout === 'theater' || venue.mode === 'theater';
      const isStandardMode = !isTheaterMode && (venue.layout === 'standard' || venue.mode === 'standard');

      if (isStandardMode) {
        // 标准模式：每排独立居中渲染
        // 按座位数分组——连续相同座位数的排为一组，只在每组首排显示座位号
        const numBoxHeightWithGap = numBoxHeight + 8;

        const rowGroups = [];
        venue.rows.forEach((row, rowIdx) => {
          if (row.seatGroups) {
            const seatCount = row.seatGroups.reduce((s, g) => s + g.filter(x => x !== null && x !== undefined).length, 0);
            const totalSlots = row.seatGroups.reduce((s, g) => s + g.length, 0);
            const prev = rowGroups[rowGroups.length - 1];
            if (prev && prev.seatCount === seatCount && prev.totalSlots === totalSlots) {
              prev.rows.push(row);
            } else {
              rowGroups.push({ seatCount, totalSlots, rows: [row], isFirst: true });
            }
          }
        });

        let maxRowSeats = 0;
        venue.rows.forEach(row => {
          if (row.seatGroups) {
            const count = row.seatGroups.reduce((s, g) => s + g.length, 0);
            if (count > maxRowSeats) maxRowSeats = count;
          }
        });
        totalSeatWidth = maxRowSeats * (seatWidth + seatGap);
        svgWidth = Math.max(svgWidth, totalSeatWidth + 500, 1300);

        rowGroups.forEach(group => {
          group.rows.forEach((row, ri) => {
            if (row.isAisle) { y += aisleHeight; return; }

            const rowLabel = row.label || '';
            const isFirstRowInGroup = ri === 0;

            if (row.seatGroups) {
              let rowWidth = 0;
              row.seatGroups.forEach((group, gi) => {
                rowWidth += group.length * (seatWidth + seatGap);
                if (gi < row.seatGroups.length - 1) rowWidth += groupGap;
              });
              let sx = Math.max(160, (svgWidth - rowWidth) / 2);

              svgContent += `  ` + bilingualLabel(rowLabel, sx - 10, y + seatHeight / 2 + 10, 'end', 32, '#ef4444') + `\n`;
              svgContent += `  ` + bilingualLabel(rowLabel, sx + rowWidth + 10, y + seatHeight / 2 + 10, 'start', 32, '#ef4444') + `\n`;

              row.seatGroups.forEach((group, gi) => {
                if (gi > 0) {
                  const aisleX = sx - seatGap / 2;
                  const aisleH = seatHeight - 10;
                  svgContent += `  <rect x="${aisleX}" y="${y + 5}" width="${groupGap}" height="${aisleH}" fill="#dbeafe" rx="3"/>\n`;
                  svgContent += `  <text x="${aisleX + groupGap / 2}" y="${y + seatHeight / 2 + 6}" class="aisle-label" text-anchor="middle" font-size="16">过道</text>\n`;
                  sx += groupGap;
                }
                group.forEach((seatNum) => {
                  if (seatNum === null || seatNum === undefined) {
                    const aisleW = seatWidth + seatGap;
                    svgContent += `  <rect x="${sx}" y="${y + 10}" width="${aisleW}" height="${seatHeight - 20}" fill="#dbeafe" rx="4"/>\n`;
                    svgContent += `  <text x="${sx + aisleW / 2}" y="${y + seatHeight / 2 + 6}" class="aisle-label" text-anchor="middle" font-size="18">过道</text>\n`;
                    sx += aisleW;
                    return;
                  }

                  const name = attendeeMap[rowLabel + '_' + seatNum];

                  if (isFirstRowInGroup) {
                    const numBoxX = sx + (seatWidth - numBoxWidth) / 2;
                    const numBoxY = y - numBoxHeight - 8;
                    svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                    svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${seatNum}</text>\n`;
                  }

                  const fill = name ? '#dbeafe' : '#ffffff';
                  svgContent += `  <rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="${fill}" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
                  if (seatNum) regionSeatPositions[rowLabel + '_' + seatNum] = { x: sx, y, w: seatWidth, h: seatHeight };
                  if (name) {
                    const displayName = name.replace(/\n/g, ' ');
                    const fontSize = displayName.length > 6 ? 28 : displayName.length > 4 ? 34 : 40;
                    svgContent += `  <text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + fontSize / 3}" class="seat-name" text-anchor="middle" font-size="${fontSize}">${escXml(displayName)}</text>\n`;
                  }
                  sx += seatWidth + seatGap;
                });
              });
            }

            y += seatHeight + rowGap;
            if (row.hasAisleAfter) {
              y += 20;
              const aisleY = y;
              const aisleHeightVal = 30;
              const aisleWidth = totalSeatWidth + 80;
              const aisleX = (svgWidth - aisleWidth) / 2;
              svgContent += `  <rect x="${aisleX}" y="${aisleY}" width="${aisleWidth}" height="${aisleHeightVal}" fill="#dbeafe" rx="4"/>\n`;
              svgContent += `  <text x="${svgWidth / 2}" y="${aisleY + 22}" class="aisle-label" text-anchor="middle">横向过道</text>\n`;
              y += aisleHeightVal + 56;
            }
          });
        });
      } else {
      // 剧院模式：使用分组对齐渲染
      let maxGroupCount = 0;
      const maxSeatsPerGroup = {};
      venue.rows.forEach(row => {
        if (row.seatGroups) {
          maxGroupCount = Math.max(maxGroupCount, row.seatGroups.length);
          row.seatGroups.forEach((g, gi) => {
            maxSeatsPerGroup[gi] = Math.max(maxSeatsPerGroup[gi] || 0, g.length);
          });
        }
      });
      if (maxGroupCount === 0) maxGroupCount = 1;
      
      // 计算SVG宽度
      let totalSeatWidth = 0;
      for (let gi = 0; gi < maxGroupCount; gi++) {
        const groupSeats = maxSeatsPerGroup[gi] || 0;
        totalSeatWidth += groupSeats * (seatWidth + seatGap);
        if (gi < maxGroupCount - 1) totalSeatWidth += groupGap - seatGap;
      }
      svgWidth = Math.max(svgWidth, totalSeatWidth + 400, 1200);

      // 剧院模式：连续相同座位模式的行为一组，仅组首行显示座位号
      const startX = 180;
      const theaterGroups = [];
      venue.rows.forEach((row) => {
        if (row.isAisle) {
          theaterGroups.push({ type: 'aisle' });
          return;
        }
        const pattern = row.seatGroups ? row.seatGroups.map(g => (g || []).filter(x => x !== null && x !== undefined).length).join(',') : '';
        const last = theaterGroups[theaterGroups.length - 1];
        if (last && last.type === 'seats' && last.pattern === pattern) {
          last.rows.push(row);
        } else {
          theaterGroups.push({ type: 'seats', pattern, rows: [row] });
        }
      });

      theaterGroups.forEach((group) => {
        if (group.type === 'aisle') {
          y += aisleHeight;
          return;
        }
        group.rows.forEach((row, ri) => {
          const isFirstInGroup = ri === 0;
          const rowLabel = row.label || '';
          const labelY = y + seatHeight / 2 + 10;

          // 排号标签（左侧左对齐，右侧右对齐）
          svgContent += `  ` + bilingualLabel(rowLabel, 16, labelY, 'start', 36, '#ef4444') + `\n`;
          svgContent += `  ` + bilingualLabel(rowLabel, svgWidth - 16, labelY, 'end', 36, '#ef4444') + `\n`;

          if (row.seatGroups) {
            let x = startX;
            for (let gi = 0; gi < maxGroupCount; gi++) {
              const groupSeats = row.seatGroups[gi] || [];
              const groupMax = maxSeatsPerGroup[gi] || 0;
              const groupActualWidth = groupSeats.length * (seatWidth + seatGap);
              const groupMaxWidth = groupMax * (seatWidth + seatGap);
              const offset = groupMax > groupSeats.length ? (groupMaxWidth - groupActualWidth) / 2 : 0;
              let sx = x + offset;

              groupSeats.forEach((seatNum) => {
                const name = attendeeMap[rowLabel + '_' + seatNum];

                // 座位号仅组首行显示
                if (isFirstInGroup && seatNum) {
                  const numBoxX = sx + (seatWidth - numBoxWidth) / 2;
                  const numBoxY = y - numBoxHeight - 8;
                  svgContent += `  <rect x="${numBoxX}" y="${numBoxY}" width="${numBoxWidth}" height="${numBoxHeight}" fill="#1a56db" rx="6"/>\n`;
                  svgContent += `  <text x="${numBoxX + numBoxWidth / 2}" y="${numBoxY + 24}" class="seat-num" text-anchor="middle">${seatNum}</text>\n`;
                }

                svgContent += `  <rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" rx="6"/>\n`;
                if (rowLabel && seatNum) regionSeatPositions[rowLabel + '_' + seatNum] = { x: sx, y, w: seatWidth, h: seatHeight };
                if (name) {
                  const displayName = name.replace(/\n/g, ' ');
                  const fontSize = displayName.length > 6 ? 28 : displayName.length > 4 ? 34 : 40;
                  svgContent += `  <text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + fontSize / 3}" class="seat-name" text-anchor="middle" font-size="${fontSize}">${escXml(displayName)}</text>\n`;
                }
                sx += seatWidth + seatGap;
              });

              x += groupMaxWidth;
              if (gi < maxGroupCount - 1) x += groupGap - seatGap;
            }
          }

          y += seatHeight + rowGap;
          if (row.hasAisleAfter) {
            y += 20;
            const aisleY = y;
            const aisleHeightVal = 30;
            const aisleWidth = totalSeatWidth + 80;
            const aisleX = (svgWidth - aisleWidth) / 2;
            svgContent += `  <rect x="${aisleX}" y="${aisleY}" width="${aisleWidth}" height="${aisleHeightVal}" fill="#dbeafe" rx="4"/>\n`;
            svgContent += `  <text x="${svgWidth / 2}" y="${aisleY + 22}" class="aisle-label" text-anchor="middle">横向过道</text>\n`;
            y += aisleHeightVal + 56;
          }
        });
      });
      }
      svgContent += renderRegionOverlays(venue, regionSeatPositions);
      y += 60;
    });

    svgContent += '</svg>';

    const filename = `${siteTitle}_座位布局图.svg`;
    const encodedFilename = encodeURIComponent(filename).replace(/['()]/g, escape).replace(/\*/g, '%2A');
    res.setHeader('Content-Type', 'image/svg+xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="seating-layout.svg"; filename*=UTF-8''${encodedFilename}`);
    res.setHeader('Content-Length', Buffer.byteLength(svgContent, 'utf8'));
    res.send(svgContent);
  } catch (err) {
    structuredLog('error', { message: 'SVG 导出失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '导出失败: ' + err.message });
  }
});

// 自动分析已上传表格的布局（从 uploaded.xlsx 重新解析）
app.post('/api/analyze-layout', requireAdmin, (req, res) => {
  try {
    const uploadedPath = path.join(__dirname, 'uploads', 'uploaded.xlsx');
    if (!fs.existsSync(uploadedPath)) {
      return res.status(404).json({ error: '未找到已上传的表格文件，请先上传 Excel 文件' });
    }
    
    const mode = req.body.mode || 'auto';
    const sheetModes = req.body.sheetModes || {};
    const wb = XLSX.readFile(uploadedPath);
    const existing = readData();
    const manualAttendees = (existing.attendees || []).filter(a => a.source !== 'excel');
    const result = parseWorkbook(wb, manualAttendees, mode, sheetModes);
    
    // ===== 保护自定义会场：优先保留所有自定义布局 =====
    const customVenues = (existing.venues || []).filter(v => v.layout === 'custom');
    // 只保留新解析的非自定义会场，同时过滤掉已被用户删除的会场
    const nonCustomNewVenues = result.venues.filter(v => {
      if (v.layout === 'custom') return false;
      
      // 检查是否已删除（同时检查原名称和规范化名称）
      if (existing.deletedVenueNames && existing.deletedVenueNames.includes(v.name)) {
        return false;
      }
      const normalizedName = v.name.trim().toLowerCase().replace(/\s+/g, '');
      if (existing.deletedVenueNamesNormalized && existing.deletedVenueNamesNormalized.includes(normalizedName)) {
        return false;
      }
      
      return true;
    });
    // 合并：自定义会场 + 新解析的普通会场（过滤已删除）
    result.venues = [...customVenues, ...nonCustomNewVenues];
    
    // ===== 过滤掉孤儿参会者（所属会场已被删除或不存在的）=====
    const finalVenueIds = new Set(result.venues.map(v => v.id));
    result.attendees = result.attendees.filter(a => finalVenueIds.has(a.venueId));
    
    // ===== 过滤掉已被用户删除的参会者 =====
    if (existing.deletedAttendees && existing.deletedAttendees.length > 0) {
      result.attendees = result.attendees.filter(a => {
        // 检查该参会者是否在已删除列表中
        const key = `${a.name || ''}|${a.venueId || ''}|${a.row || ''}|${a.seat || ''}`;
        if (existing.deletedAttendees.includes(key)) return false;
        
        // 同时检查规范化名称的匹配
        const normalizedName = (a.name || '').trim().toLowerCase().replace(/\s+/g, '');
        if (normalizedName) {
          const normalizedKey = `${normalizedName}|${a.venueId || ''}|${a.row || ''}|${a.seat || ''}`;
          if (existing.deletedAttendees.includes(normalizedKey)) return false;
        }
        
        return true;
      });
    }
    
    // ===== 保留已删除记录，防止下次自动分析时重新出现 =====
    if (existing.deletedVenueNames) result.deletedVenueNames = existing.deletedVenueNames;
    if (existing.deletedVenueNamesNormalized) result.deletedVenueNamesNormalized = existing.deletedVenueNamesNormalized;
    if (existing.deletedAttendees) result.deletedAttendees = existing.deletedAttendees;
    
    writeData(result);
    
    res.json({ 
      ok: true, 
      data: { venues: result.venues },
      mode: mode,
      sheetModes: sheetModes,
      message: '布局分析完成（已保留自定义会场）'
    });
  } catch (err) {
    structuredLog('error', { message: '布局分析失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '分析失败: ' + err.message });
  }
});

// 生成布局预览图
app.post('/api/generate-preview', requireAdmin, (req, res) => {
  try {
    // 自定义布局临时预览（未保存到 data.json），使用实际位置信息
    if (req.body.customPreview) {
      const { venueName, customRows, canvasWidth, canvasHeight, customStage, customGates, customAisles } = req.body.customPreview;
      if (!customRows || !customRows.length) {
        return res.json({ ok: true, data: { venues: [] } });
      }
      // 计算缩放比例——基于座位密度确保每个座位至少 60px 宽/高
      const cw = canvasWidth || 800;
      const ch = canvasHeight || 600;
      const margin = 60;
      const canvasScale = Math.min(1800 / Math.max(cw + margin * 2, 1), 2);
      
      let minOrigSeatPx1 = Infinity;
      customRows.forEach(row => {
        const size = row.direction === 'horizontal' 
          ? (row.width || 100) / Math.max(row.seatCount, 1)
          : (row.height || 50) / Math.max(row.seatCount, 1);
        minOrigSeatPx1 = Math.min(minOrigSeatPx1, size);
      });
      const MIN_PREVIEW_SEAT_PX = 60;
      const seatScale = Math.min(MIN_PREVIEW_SEAT_PX / Math.max(minOrigSeatPx1, 0.5), 10);
      const scale = Math.max(canvasScale, seatScale);
      
      // 为座位号和排标签留出额外空间
      const extraMargin = 60;
      const svgW = cw * scale + (margin + extraMargin) * 2;
      const svgH = ch * scale + (margin + extraMargin) * 2 + 60;
      
      let sc = `<rect width="100%" height="100%" fill="#f8fafc"/>`;
      sc += `<text x="${svgW/2}" y="35" text-anchor="middle" font-size="18" font-family="Microsoft YaHei, sans-serif" font-weight="bold" fill="#1e293b">${escXml(venueName || '自定义会场')}</text>`;
      
      const sorted = [...customRows].sort((a, b) => a.rowNum - b.rowNum);
      // 构建分组：连续且seatCount相同的排为一组
      const groups = [];
      let currentGroup = [sorted[0]];
      for (let g = 1; g < sorted.length; g++) {
        if (sorted[g].seatCount === sorted[g-1].seatCount) {
          currentGroup.push(sorted[g]);
        } else {
          groups.push(currentGroup);
          currentGroup = [sorted[g]];
        }
      }
      if (currentGroup.length > 0) groups.push(currentGroup);

      groups.forEach(group => {
        group.forEach((row, rowIdxInGroup) => {
          const isFirstInGroup = (rowIdxInGroup === 0);
          const rl = row.label || '';
          const x = margin + (row.x || 0) * scale;
          const y = margin + 60 + (row.y || 0) * scale;
          const w = (row.width || 100) * scale;
          const h = (row.height || 50) * scale;
          const getSn = (idx) => (row.seatNumbers && row.seatNumbers.length > idx) ? row.seatNumbers[idx] : (row.startSeat || 1) + idx;
          const aisles = (row.aislePositions || []).slice().sort((a, b) => a - b);
          const aisleGapPx = Math.max(40, (row.direction === 'horizontal' ? w : h) * 0.05);
          const totalAisleGap = aisles.length * aisleGapPx;

          if (row.direction === 'horizontal') {
            const availW = w - totalAisleGap;
            const seatW = availW / row.seatCount;
            const seatH = Math.min(80, h * 0.95);
            let cursorX = x;
            for (let i = 0; i < row.seatCount; i++) {
              for (let a = 0; a < aisles.length; a++) {
                if (aisles[a] === i) {
                  sc += `<rect x="${cursorX}" y="${y + h/2 - 12}" width="${aisleGapPx}" height="24" fill="#dbeafe" rx="3"/>`;
                  sc += `<text x="${cursorX + aisleGapPx/2}" y="${y + h/2 + 6}" text-anchor="middle" font-size="${Math.min(14, 24/1.6)}" font-family="Microsoft YaHei, sans-serif" fill="#94a3b8" font-weight="bold">过道</text>`;
                  cursorX += aisleGapPx;
                }
              }
              const seatX = cursorX;
              const seatY = y + (h - seatH) / 2;
              // 座位框（始终渲染）
              sc += `<rect x="${seatX + 2}" y="${seatY}" width="${seatW - 4}" height="${seatH}" fill="#fff" stroke="#cbd5e1" stroke-width="2" rx="4"/>`;
              // 座位号（仅组内第一排显示，蓝色方块在座位上方）
              if (isFirstInGroup) {
                const numBoxW = Math.min(50, seatW * 0.8);
                const numBoxH = Math.max(18, Math.min(28, seatH * 0.45));
                const numFS = Math.max(10, Math.min(16, numBoxH * 0.55));
                const nbx = seatX + (seatW - numBoxW) / 2;
                const nby = seatY - numBoxH - 6;
                sc += `<rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="4"/>`;
                sc += `<text x="${nbx + numBoxW/2}" y="${nby + numBoxH - 6}" text-anchor="middle" font-size="${numFS}" font-family="Microsoft YaHei, sans-serif" fill="#ffffff">${getSn(i)}</text>`;
              }
              cursorX += seatW;
            }
          } else {
            const availH = h - totalAisleGap;
            const seatH = availH / row.seatCount;
            const seatW = Math.min(100, w * 0.95);
            let cursorY = y;
            for (let i = 0; i < row.seatCount; i++) {
              for (let a = 0; a < aisles.length; a++) {
                if (aisles[a] === i) {
                  sc += `<rect x="${x + w/2 - 12}" y="${cursorY}" width="24" height="${aisleGapPx}" fill="#dbeafe" rx="3"/>`;
                  sc += `<text x="${x + w/2}" y="${cursorY + aisleGapPx/2 + 6}" text-anchor="middle" font-size="${Math.min(14, 24/1.6)}" font-family="Microsoft YaHei, sans-serif" fill="#94a3b8" font-weight="bold">过道</text>`;
                  cursorY += aisleGapPx;
                }
              }
              const seatX = x + (w - seatW) / 2;
              const seatY = cursorY;
              // 座位框（始终渲染）
              sc += `<rect x="${seatX}" y="${seatY + 2}" width="${seatW}" height="${seatH - 4}" fill="#fff" stroke="#cbd5e1" stroke-width="2" rx="4"/>`;
              // 座位号（仅组内第一排显示，蓝色方块在座位左侧）
              if (isFirstInGroup) {
                const numBoxW = Math.min(42, seatW * 0.7);
                const numBoxH = Math.max(18, Math.min(30, seatH * 0.75));
                const numFS = Math.max(10, Math.min(16, numBoxH * 0.5));
                const nbx = seatX - numBoxW - 8;
                const nby = seatY + (seatH - numBoxH) / 2;
                sc += `<rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="4"/>`;
                sc += `<text x="${nbx + numBoxW/2}" y="${nby + numBoxH/2 + 5}" text-anchor="middle" font-size="${numFS}" font-family="Microsoft YaHei, sans-serif" fill="#ffffff">${getSn(i)}</text>`;
              }
              cursorY += seatH;
            }
          }
          sc += `<text x="${x - 12}" y="${y + h/2 + 6}" text-anchor="end" font-size="${Math.max(12, Math.min(18, h / 2.8))}" font-family="Microsoft YaHei, sans-serif" fill="#ef4444">${escXml(rl)}</text>`;
        });
      });
      
      // 渲染自定义舞台（临时预览）
      if (customStage && Object.keys(customStage).length > 0) {
        const stageX = margin + (customStage.x || 0) * scale;
        const stageY = margin + 60 + (customStage.y || 0) * scale;
        const stageW = (customStage.width || 200) * scale;
        const stageH = (customStage.height || 80) * scale;
        sc += `<rect x="${stageX}" y="${stageY}" width="${stageW}" height="${stageH}" fill="#1a56db" rx="8"/>`;
        sc += `<text x="${stageX + stageW/2}" y="${stageY + stageH/2 + 5}" text-anchor="middle" fill="#ffffff" font-size="20" font-weight="bold">${escXml(customStage.label || '舞台')}</text>`;
      }
      
      // 渲染自定义大门（临时预览）
      if (customGates && customGates.length > 0) {
        customGates.forEach(gate => {
          const gateX = margin + (gate.x || 0) * scale;
          const gateY = margin + 60 + (gate.y || 0) * scale;
          const gateW = (gate.width || 80) * scale;
          const gateH = (gate.height || 120) * scale;
          sc += `<rect x="${gateX}" y="${gateY}" width="${gateW}" height="${gateH}" fill="#10b981" stroke="#059669" stroke-width="3" rx="6"/>`;
          sc += `<text x="${gateX + gateW/2}" y="${gateY + gateH/2 + 5}" text-anchor="middle" fill="#ffffff" font-size="16" font-weight="bold">${escXml(gate.label || '门')}</text>`;
        });
      }
      
      // 渲染自定义过道（临时预览）
      if (customAisles && customAisles.length > 0) {
        customAisles.forEach(aisle => {
          const aisleX = margin + (aisle.x || 0) * scale;
          const aisleY = margin + 60 + (aisle.y || 0) * scale;
          const aisleW = (aisle.width || 60) * scale;
          const aisleH = (aisle.height || 40) * scale;
          sc += `<rect x="${aisleX}" y="${aisleY}" width="${aisleW}" height="${Math.min(aisleH, 24)}" fill="#dbeafe" rx="3"/>`;
          const aisleDisplayLabel2 = (aisle.label && aisle.label !== 'null') ? aisle.label : '过道';
          sc += `<text x="${aisleX + aisleW/2}" y="${aisleY + Math.min(aisleH, 24)/2 + 6}" text-anchor="middle" fill="#94a3b8" font-size="${Math.min(16, Math.min(aisleH, 24)/1.6)}" font-weight="bold">${escXml(aisleDisplayLabel2)}</text>`;
        });
      }
      
      return res.json({ ok: true, data: { venues: [{ name: venueName || '自定义会场', width: svgW, height: svgH, svgContent: sc }] } });
    }

    const data = readData();
    const config = readConfig();
    const siteTitle = config.siteTitle || '会议';
    
    // 座位参数
    const seatWidth = 60;
    const seatHeight = 30;
    const seatGap = 8;
    const groupGap = 30;
    const rowGap = 12;
    const margin = 20;
    
    // 辅助函数：计算排宽度
    function getRowWidth(row) {
      let w = 0;
      if (!row.seatGroups) return w;
      row.seatGroups.forEach((g, gIdx) => {
        w += g.length * (seatWidth + seatGap);
        if (gIdx < row.seatGroups.length - 1) w += groupGap - seatGap;
      });
      return w;
    }
    
    // 辅助函数：座位组结构签名（基于分组数，而非精确座位数）
    function getGroupSignature(row) {
      if (!row.seatGroups) return '';
      // 使用分组数作为签名，确保相同过道数的排在同一组
      return `groups:${row.seatGroups.length}`;
    }
    
    const previewVenues = [];
    
    data.venues.forEach(venue => {
      const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
      const attendeeMap = {};
      venueAttendees.forEach(a => {
        attendeeMap[a.row + '_' + a.seat] = a.name;
      });
      
      // 自定义布局：根据 customRows 渲染干净矢量预览（使用实际位置）
      if (venue.layout === 'custom' && venue.customRows && venue.customRows.length > 0) {
        const cs = venue.customRows;
        // 计算边界（包含座位号的空间）
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        cs.forEach(row => {
          // 为座位号留出空间：横排座位上方，竖排座位左侧
          const extraLeft = (row.direction === 'vertical') ? 80 : 50;
          const extraTop = (row.direction === 'horizontal') ? 50 : 0;
          
          minX = Math.min(minX, (row.x || 0) - extraLeft / 2);
          minY = Math.min(minY, (row.y || 0) - extraTop);
          maxX = Math.max(maxX, (row.x || 0) + (row.width || 100) + 20);
          maxY = Math.max(maxY, (row.y || 0) + (row.height || 50) + 20);
        });
        // 只在明确有自定义舞台时加入到边界
        if (venue.customStage && Object.keys(venue.customStage).length > 0) {
          const stage = venue.customStage;
          minX = Math.min(minX, stage.x || 0);
          minY = Math.min(minY, stage.y || 0);
          maxX = Math.max(maxX, (stage.x || 0) + (stage.width || 200));
          maxY = Math.max(maxY, (stage.y || 0) + (stage.height || 80));
        }
        if (venue.customGates && venue.customGates.length > 0) {
          venue.customGates.forEach(gate => {
            minX = Math.min(minX, gate.x || 0);
            minY = Math.min(minY, gate.y || 0);
            maxX = Math.max(maxX, (gate.x || 0) + (gate.width || 80));
            maxY = Math.max(maxY, (gate.y || 0) + (gate.height || 120));
          });
        }
        const contentW = maxX - minX;
        const contentH = maxY - minY;
        
        // 计算缩放比例 - 综合考虑画布和座位密度
        const margin = 80;
        const maxPreviewWidth = 1600;
        const canvasScale = Math.min(maxPreviewWidth / Math.max(contentW, 1), 1200 / Math.max(contentH, 1));
        
        // 基于座位密度计算最小缩放：确保每个座位至少 85px 宽/高
        let minOrigSeatPx2 = Infinity;
        cs.forEach(row => {
          const size = row.direction === 'horizontal' 
            ? (row.width || 100) / Math.max(row.seatCount, 1)
            : (row.height || 50) / Math.max(row.seatCount, 1);
          minOrigSeatPx2 = Math.min(minOrigSeatPx2, size);
        });
        const MIN_SEAT_PX2 = 85;
        const seatScale = Math.min(MIN_SEAT_PX2 / Math.max(minOrigSeatPx2, 1), 12);
        const scale = Math.max(canvasScale, seatScale);
        
        const baseX = margin - minX * scale;
        const baseY = margin;
        const totalSW = (maxX - minX) * scale + margin * 2;
        const svgH = (maxY - minY) * scale + margin * 2 + 60;
        
        let sc = `<rect width="100%" height="100%" fill="#f8fafc"/>`;
        
        const sorted = [...cs].sort((a, b) => a.rowNum - b.rowNum);
        
        // === 1. 先渲染舞台和大门 ===
        if (venue.customStage && Object.keys(venue.customStage).length > 0) {
          const stage = venue.customStage;
          const stageX = baseX + (stage.x || 0) * scale;
          const stageY = baseY + (stage.y || 0) * scale;
          const stageW = (stage.width || 200) * scale;
          const stageH = (stage.height || 80) * scale;
          sc += `<rect x="${stageX}" y="${stageY}" width="${stageW}" height="${stageH}" fill="#1a56db" rx="8"/>`;
          sc += `<text x="${stageX + stageW/2}" y="${stageY + stageH/2 + 8}" text-anchor="middle" fill="#ffffff" font-size="${Math.min(32, stageH/2.2)}" font-weight="bold">${escXml(stage.label || '舞台')}</text>`;
        }
        
        if (venue.customGates && venue.customGates.length > 0) {
          venue.customGates.forEach(gate => {
            const gateX = baseX + (gate.x || 0) * scale;
            const gateY = baseY + (gate.y || 0) * scale;
            const gateW = (gate.width || 80) * scale;
            const gateH = (gate.height || 120) * scale;
            sc += `<rect x="${gateX}" y="${gateY}" width="${gateW}" height="${gateH}" fill="#10b981" stroke="#059669" stroke-width="3" rx="6"/>`;
            sc += `<text x="${gateX + gateW/2}" y="${gateY + gateH/2 + 8}" text-anchor="middle" fill="#ffffff" font-size="${Math.min(26, gateH/3)}" font-weight="bold">${escXml(gate.label || '门')}</text>`;
          });
        }
        
        if (venue.customAisles && venue.customAisles.length > 0) {
          venue.customAisles.forEach(aisle => {
            const aisleX = baseX + (aisle.x || 0) * scale;
            const aisleY = baseY + (aisle.y || 0) * scale;
            const aisleW = (aisle.width || 60) * scale;
            const aisleH = (aisle.height || 40) * scale;
            sc += `<rect x="${aisleX}" y="${aisleY}" width="${aisleW}" height="${Math.min(aisleH, 24)}" fill="#dbeafe" rx="3"/>`;
            const aisleDisplayLabel3 = (aisle.label && aisle.label !== 'null') ? aisle.label : '过道';
            sc += `<text x="${aisleX + aisleW/2}" y="${aisleY + Math.min(aisleH, 24)/2 + 6}" text-anchor="middle" fill="#94a3b8" font-size="${Math.min(16, Math.min(aisleH, 24)/1.6)}" font-weight="bold">${escXml(aisleDisplayLabel3)}</text>`;
          });
        }
        
        // === 2. 再渲染座位 ===
        // 构建分组：连续且seatCount相同的排为一组
        const rowGroups = [];
        let curGroup = [sorted[0]];
        for (let g = 1; g < sorted.length; g++) {
          if (sorted[g].seatCount === sorted[g-1].seatCount) {
            curGroup.push(sorted[g]);
          } else {
            rowGroups.push(curGroup);
            curGroup = [sorted[g]];
          }
        }
        if (curGroup.length > 0) rowGroups.push(curGroup);

        rowGroups.forEach(group => {
          group.forEach((row, rowIdxInGroup) => {
            const isFirstInGroup = (rowIdxInGroup === 0);
            const rl = row.label || '';
            const x = baseX + (row.x || 0) * scale;
            const y = baseY + (row.y || 0) * scale;
            const w = (row.width || 100) * scale;
            const h = (row.height || 50) * scale;
            const getSn = (idx) => (row.seatNumbers && row.seatNumbers.length > idx) ? row.seatNumbers[idx] : (row.startSeat || 1) + idx;
            const aisles = (row.aislePositions || []).slice().sort((a, b) => a - b);
            const aisleGapPx = Math.max(40, (row.direction === 'horizontal' ? w : h) * 0.05);
            const totalAisleGap = aisles.length * aisleGapPx;

            if (row.direction === 'horizontal') {
              const seatCount = row.seatCount;
              const availW = w - totalAisleGap;
              const seatW = availW / seatCount;
              const seatH = Math.max(50, h * 0.9);

              let cursorX = x;
              for (let i = 0; i < seatCount; i++) {
                for (let a = 0; a < aisles.length; a++) {
                  if (aisles[a] === i) {
                    sc += `<rect x="${cursorX}" y="${y + h/2 - 12}" width="${aisleGapPx}" height="24" fill="#dbeafe" rx="3"/>`;
                    sc += `<text x="${cursorX + aisleGapPx/2}" y="${y + h/2 + 6}" text-anchor="middle" font-size="${Math.min(14, 24/1.6)}" font-family="Microsoft YaHei, sans-serif" fill="#94a3b8" font-weight="bold">过道</text>`;
                    cursorX += aisleGapPx;
                  }
                }
                const sx = cursorX;
                const sy = y + (h - seatH) / 2;
                const sn = getSn(i);
                const name = attendeeMap[rl + '_' + sn];

                if (i === 0) {
                  sc += `<text x="${sx - 8}" y="${sy + seatH/2 + 6}" text-anchor="end" font-size="${Math.max(18, Math.min(28, h/1.8))}" font-family="Microsoft YaHei, sans-serif" fill="#ef4444" font-weight="bold">${escXml(rl)}</text>`;
                }

                // 座位号蓝色方块（仅组内第一排显示）
                if (isFirstInGroup) {
                  const numBoxW = Math.min(50, seatW * 0.8);
                  const numBoxH = 28;
                  const nbx = sx + (seatW - numBoxW) / 2;
                  const nby = sy - numBoxH - 8;
                  sc += `<rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="6"/>`;
                  sc += `<text x="${nbx + numBoxW/2}" y="${nby + 20}" text-anchor="middle" font-size="18" font-family="Microsoft YaHei, sans-serif" fill="#ffffff">${sn}</text>`;
                }

                sc += `<rect x="${sx + 2}" y="${sy}" width="${seatW - 4}" height="${seatH}" fill="${name ? '#dbeafe' : '#fff'}" stroke="#cbd5e1" stroke-width="2" rx="6"/>`;
                if (name) {
                  const dn = name.replace(/\n/g, ' ');
                  const nameLen = Math.max(dn.length, 1);
                  const maxFit = Math.min(seatW * 0.85 / nameLen, seatH * 0.45);
                  const fs = Math.max(11, Math.min(28, maxFit));
                  sc += `<text x="${sx + seatW/2}" y="${sy + seatH/2 + fs/3}" text-anchor="middle" font-size="${fs}" font-family="Microsoft YaHei, sans-serif" fill="#1e293b">${escXml(dn)}</text>`;
                }
                cursorX += seatW;
              }
            } else {
              const seatCount = row.seatCount;
              const availH = h - totalAisleGap;
              const seatH = availH / seatCount;
              const seatW = Math.max(60, w * 0.9);

              let cursorY = y;
              for (let i = 0; i < seatCount; i++) {
                for (let a = 0; a < aisles.length; a++) {
                  if (aisles[a] === i) {
                    sc += `<rect x="${x + w/2 - 12}" y="${cursorY}" width="24" height="${aisleGapPx}" fill="#dbeafe" rx="3"/>`;
                    sc += `<text x="${x + w/2}" y="${cursorY + aisleGapPx/2 + 6}" text-anchor="middle" font-size="${Math.min(14, 24/1.6)}" font-family="Microsoft YaHei, sans-serif" fill="#94a3b8" font-weight="bold">过道</text>`;
                    cursorY += aisleGapPx;
                  }
                }
                const sx = x + (w - seatW) / 2;
                const sy = cursorY;
                const sn = getSn(i);
                const name = attendeeMap[rl + '_' + sn];

                if (i === 0) {
                  sc += `<text x="${sx + seatW/2}" y="${sy - 10}" text-anchor="middle" font-size="${Math.max(18, Math.min(28, h/1.8))}" font-family="Microsoft YaHei, sans-serif" fill="#ef4444" font-weight="bold">${escXml(rl)}</text>`;
                }

                // 座位号蓝色方块（仅组内第一排显示）
                if (isFirstInGroup) {
                  const numBoxW = 42;
                  const numBoxH = Math.min(36, seatH * 0.85);
                  const nbx = sx - numBoxW - 8;
                  const nby = sy + (seatH - numBoxH) / 2;
                  sc += `<rect x="${nbx}" y="${nby}" width="${numBoxW}" height="${numBoxH}" fill="#1a56db" rx="6"/>`;
                  sc += `<text x="${nbx + numBoxW/2}" y="${nby + numBoxH/2 + 6}" text-anchor="middle" font-size="18" font-family="Microsoft YaHei, sans-serif" fill="#ffffff">${sn}</text>`;
                }

                sc += `<rect x="${sx}" y="${sy + 2}" width="${seatW}" height="${seatH - 4}" fill="${name ? '#dbeafe' : '#fff'}" stroke="#cbd5e1" stroke-width="2" rx="6"/>`;
                if (name) {
                  const dn = name.replace(/\n/g, ' ');
                  const nameLen = Math.max(dn.length, 1);
                  const maxFit = Math.min(seatW * 0.85 / nameLen, seatH * 0.45);
                  const fs = Math.max(11, Math.min(28, maxFit));
                  sc += `<text x="${sx + seatW/2}" y="${sy + seatH/2 + fs/3}" text-anchor="middle" font-size="${fs}" font-family="Microsoft YaHei, sans-serif" fill="#1e293b">${escXml(dn)}</text>`;
                }
                cursorY += seatH;
              }
            }
          });
        });
        
        previewVenues.push({ name: venue.name, width: totalSW, height: svgH, svgContent: sc });
        return;
      }
      
      // 计算所有排的最大组数和每组最大座位数
      let maxGroupCount = 0;
      const maxSeatsPerGroup = {};
      venue.rows.forEach(row => {
        if (row.seatGroups) {
          maxGroupCount = Math.max(maxGroupCount, row.seatGroups.length);
          row.seatGroups.forEach((g, gi) => {
            maxSeatsPerGroup[gi] = Math.max(maxSeatsPerGroup[gi] || 0, g.length);
          });
        }
      });
      if (maxGroupCount === 0) maxGroupCount = 1;
      
      // 计算SVG宽度：各组最大宽度 + 过道间距
      let totalSeatWidth = 0;
      for (let gi = 0; gi < maxGroupCount; gi++) {
        const groupSeats = maxSeatsPerGroup[gi] || 0;
        totalSeatWidth += groupSeats * (seatWidth + seatGap);
        if (gi < maxGroupCount - 1) totalSeatWidth += groupGap - seatGap;
      }
      const svgWidth = totalSeatWidth + margin * 2 + 160;
      
      // 计算高度
      let venueHeight = margin + 50 + margin;
      if (venue.rows) {
        venue.rows.forEach(row => {
          venueHeight += seatHeight + rowGap;
          if (row.hasAisleAfter) {
            venueHeight += 10 + 20 + 10; // 间隙 + 过道高 + 间隙
          }
        });
      }
      
      // 生成SVG内容
      let svgContent = `<rect width="100%" height="100%" fill="#f8fafc"/>`;
      let y = margin;

      const isTheaterMode = venue.layout === 'theater' || venue.mode === 'theater';
      const isStandardMode = !isTheaterMode && (venue.layout === 'standard' || venue.mode === 'standard');

      if (isStandardMode) {
        // 标准模式：每排独立居中，按座位数分组，只在每组首排显示座位号
        const rowGroups = [];
        venue.rows.forEach((row, rowIdx) => {
          if (row.seatGroups) {
            const seatCount = row.seatGroups.reduce((s, g) => s + g.filter(x => x !== null && x !== undefined).length, 0);
            const totalSlots = row.seatGroups.reduce((s, g) => s + g.length, 0);
            const prev = rowGroups[rowGroups.length - 1];
            if (prev && prev.seatCount === seatCount && prev.totalSlots === totalSlots) {
              prev.rows.push(row);
            } else {
              rowGroups.push({ seatCount, totalSlots, rows: [row] });
            }
          }
        });

        let maxRowSeats = 0;
        venue.rows.forEach(row => {
          if (row.seatGroups) {
            const count = row.seatGroups.reduce((s, g) => s + g.length, 0);
            if (count > maxRowSeats) maxRowSeats = count;
          }
        });
        totalSeatWidth = maxRowSeats * (seatWidth + seatGap);
        const svgWidth = totalSeatWidth + margin * 2 + 200;

        venueHeight = margin + 50 + margin;
        venue.rows.forEach(row => { venueHeight += seatHeight + rowGap; if (row.hasAisleAfter) venueHeight += 40; });

        svgContent = `<rect width="100%" height="100%" fill="#f8fafc"/>`;
        y = margin;
        svgContent += `<text x="${svgWidth / 2}" y="${y + 30}" font-family="Microsoft YaHei, sans-serif" font-size="16" font-weight="bold" fill="#1e293b" text-anchor="middle">${escXml(venue.name)}</text>`;
        y += 50;

        rowGroups.forEach(group => {
          group.rows.forEach((row, ri) => {
            const rowLabel = row.label || '';
            const isFirstRowInGroup = ri === 0;

            if (row.seatGroups) {
              let rowWidth = 0;
              row.seatGroups.forEach((group, gi) => {
                rowWidth += group.length * (seatWidth + seatGap);
                if (gi < row.seatGroups.length - 1) rowWidth += groupGap;
              });
              let sx = Math.max(margin + 100, (svgWidth - rowWidth) / 2);
              svgContent += `<text x="${sx - 8}" y="${y + seatHeight / 2 + 5}" font-family="Microsoft YaHei, sans-serif" font-size="12" fill="#64748b" text-anchor="end">${escXml(rowLabel)}</text>`;
              svgContent += `<text x="${sx + rowWidth + 8}" y="${y + seatHeight / 2 + 5}" font-family="Microsoft YaHei, sans-serif" font-size="12" fill="#64748b" text-anchor="start">${escXml(rowLabel)}</text>`;

              row.seatGroups.forEach((group, gi) => {
                if (gi > 0) {
                  const aisleX = sx - seatGap / 2;
                  const aisleH = seatHeight - 4;
                  svgContent += `<rect x="${aisleX}" y="${y + 2}" width="${groupGap}" height="${aisleH}" fill="#dbeafe" rx="2"/>`;
                  svgContent += `<text x="${aisleX + groupGap / 2}" y="${y + seatHeight / 2 + 4}" font-family="Microsoft YaHei, sans-serif" font-size="8" fill="#94a3b8" text-anchor="middle" font-weight="bold">过道</text>`;
                  sx += groupGap;
                }
                group.forEach((seatNum) => {
                  if (seatNum === null || seatNum === undefined) {
                    const aisleW = seatWidth + seatGap;
                    svgContent += `<rect x="${sx}" y="${y + 4}" width="${aisleW}" height="${seatHeight - 8}" fill="#dbeafe" rx="3"/>`;
                    svgContent += `<text x="${sx + aisleW / 2}" y="${y + seatHeight / 2 + 4}" font-family="Microsoft YaHei, sans-serif" font-size="8" fill="#94a3b8" text-anchor="middle" font-weight="bold">过道</text>`;
                    sx += aisleW;
                    return;
                  }
                  const name = attendeeMap[rowLabel + '_' + seatNum];
                  const fill = name ? '#dbeafe' : '#ffffff';
                  svgContent += `<rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="${fill}" stroke="#cbd5e1" stroke-width="1" rx="3"/>`;
                  if (isFirstRowInGroup) {
                    svgContent += `<text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + 4}" font-family="Microsoft YaHei, sans-serif" font-size="9" fill="#64748b" text-anchor="middle">${seatNum}</text>`;
                  }
                  if (name) {
                    const displayName = name.length > 4 ? name.substring(0, 3) + '...' : name;
                    svgContent += `<text x="${sx + seatWidth / 2}" y="${y + seatHeight - 6}" font-family="Microsoft YaHei, sans-serif" font-size="9" fill="#1e293b" text-anchor="middle">${escXml(displayName)}</text>`;
                  }
                  sx += seatWidth + seatGap;
                });
              });
            }

            y += seatHeight + rowGap;
            if (row.hasAisleAfter) {
              y += 10;
              const aisleY = y;
              const aisleHeightVal = 20;
              const aisleWidth = totalSeatWidth + 80;
              const aisleX = (svgWidth - aisleWidth) / 2;
              svgContent += `<rect x="${aisleX}" y="${aisleY}" width="${aisleWidth}" height="${aisleHeightVal}" fill="#dbeafe" rx="4"/>`;
              svgContent += `<text x="${svgWidth / 2}" y="${aisleY + 14}" font-family="Microsoft YaHei, sans-serif" font-size="10" fill="#94a3b8" text-anchor="middle">横向过道</text>`;
              y += aisleHeightVal + 24;
            }
          });
        });

        previewVenues.push({ name: venue.name, width: svgWidth, height: venueHeight, svgContent: svgContent });
      } else {

      // 标题
      svgContent += `<text x="${svgWidth / 2}" y="${y + 30}" font-family="Microsoft YaHei, sans-serif" font-size="16" font-weight="bold" fill="#1e293b" text-anchor="middle">${escXml(venue.name)}</text>`;
      y += 50;
      
      if (venue.rows && venue.rows.length > 0) {
        const startX = margin + 100;

        // 连续相同座位模式的行为一组，仅组首行显示座位号
        const previewGroups = [];
        venue.rows.forEach((row) => {
          const pattern = row.seatGroups ? row.seatGroups.map(g => (g || []).filter(x => x !== null && x !== undefined).length).join(',') : '';
          const last = previewGroups[previewGroups.length - 1];
          if (last && last.pattern === pattern) {
            last.rows.push(row);
          } else {
            previewGroups.push({ pattern, rows: [row] });
          }
        });

        previewGroups.forEach((group) => {
          group.rows.forEach((row, ri) => {
            const isFirstInGroup = ri === 0;
            const rowLabel = row.label || '';

            // 排号标签（左侧左对齐，右侧右对齐）
            const labelY = y + seatHeight / 2 + 5;
            svgContent += `<text x="${margin}" y="${labelY}" font-family="Microsoft YaHei, sans-serif" font-size="12" fill="#64748b" text-anchor="start">${escXml(rowLabel)}</text>`;
            svgContent += `<text x="${svgWidth - margin}" y="${labelY}" font-family="Microsoft YaHei, sans-serif" font-size="12" fill="#64748b" text-anchor="end">${escXml(rowLabel)}</text>`;

            if (row.seatGroups) {
              // 逐组渲染，每组使用最大宽度，较少的座位居中
              let x = startX;
              for (let gi = 0; gi < maxGroupCount; gi++) {
                const groupSeats = row.seatGroups[gi] || [];
                const groupMax = maxSeatsPerGroup[gi] || groupSeats.length;

                // 计算该组占用的实际宽度
                const groupActualWidth = groupSeats.length * (seatWidth + seatGap);
                const groupMaxWidth = groupMax * (seatWidth + seatGap);

                // 偏移量：居中对齐
                const offset = (groupMaxWidth - groupActualWidth) / 2;
                let sx = x + offset;

                groupSeats.forEach((seatNum) => {
                  const name = attendeeMap[rowLabel + '_' + seatNum];
                  svgContent += `<rect x="${sx}" y="${y}" width="${seatWidth}" height="${seatHeight}" fill="${name ? '#dbeafe' : '#ffffff'}" stroke="#cbd5e1" stroke-width="1" rx="3"/>`;
                  if (isFirstInGroup && seatNum) {
                    svgContent += `<text x="${sx + seatWidth / 2}" y="${y + seatHeight / 2 + 4}" font-family="Microsoft YaHei, sans-serif" font-size="9" fill="#64748b" text-anchor="middle">${seatNum}</text>`;
                  }
                  if (name) {
                    const displayName = name.length > 4 ? name.substring(0, 3) + '...' : name;
                    svgContent += `<text x="${sx + seatWidth / 2}" y="${y + seatHeight - 6}" font-family="Microsoft YaHei, sans-serif" font-size="9" fill="#1e293b" text-anchor="middle">${escXml(displayName)}</text>`;
                  }
                  sx += seatWidth + seatGap;
                });

                x += groupMaxWidth;
                if (gi < maxGroupCount - 1) x += groupGap - seatGap;
              }
            }

            y += seatHeight + rowGap;
            if (row.hasAisleAfter) {
              y += 10;
              const aisleY = y;
              const aisleHeightVal = 20;
              const aisleWidth = totalSeatWidth + 80;
              const aisleX = (svgWidth - aisleWidth) / 2;
              svgContent += `<rect x="${aisleX}" y="${aisleY}" width="${aisleWidth}" height="${aisleHeightVal}" fill="#dbeafe" rx="4"/>`;
              svgContent += `<text x="${svgWidth / 2}" y="${aisleY + 14}" font-family="Microsoft YaHei, sans-serif" font-size="10" fill="#94a3b8" text-anchor="middle">横向过道</text>`;
              y += aisleHeightVal + 24;
            }
          });
        });
      }

      previewVenues.push({
        name: venue.name,
        width: svgWidth,
        height: venueHeight,
        svgContent: svgContent
      });
      }
    });
    
    res.json({
      ok: true,
      data: { venues: previewVenues }
    });
  } catch (err) {
    structuredLog('error', { message: '预览图生成失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '生成失败: ' + err.message });
  }
});

// 健康检查接口
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    uptime: process.uptime(),
    timestamp: new Date().toISOString(),
    venues: readData().venues.length,
    attendees: readData().attendees.length
  });
});

// ==================== 管理 API（需认证） ====================

// 批量导入参会者（指定场馆）
app.post('/api/attendees/import', requireAuth, (req, res) => {
  const data = readData();
  const { attendees, mode, venueId } = req.body;

  if (!venueId) {
    return res.status(400).json({ error: '缺少 venueId 参数' });
  }

  // 验证导入数据
  if (!Array.isArray(attendees) || attendees.length === 0) {
    return res.status(400).json({ error: '无效的参会者数据' });
  }
  if (attendees.length > MAX_ATTENDEES_IMPORT) {
    return res.status(413).json({ 
      error: `导入数据过多（${attendees.length} 人），不超过 ${MAX_ATTENDEES_IMPORT} 人` 
    });
  }

  const cleaned = attendees.map(a => {
    const attendee = cleanAttendee({ ...a, venueId });
    // 生成ID
    if (!attendee.id) {
      attendee.id = 'att-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
    }
    return attendee;
  });
  const duplicates = [];
  const conflicts = [];
  const valid = [];

  cleaned.forEach(a => {
    // 检查重复
    if (mode === 'append' && findDuplicateAttendee(data, a)) {
      duplicates.push(a.name);
      return;
    }
    // 检查座位冲突
    const conflict = findSeatConflict(data, a);
    if (conflict) {
      conflicts.push({ name: a.name, row: a.row, seat: a.seat, occupiedBy: conflict.name });
      return;
    }
    valid.push(a);
  });

  if (mode === 'replace') {
    data.attendees = data.attendees.filter(a => a.venueId !== venueId);
  }

  data.attendees = data.attendees.concat(valid);
  writeData(data);

  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
  auditLog('IMPORT_ATTENDEES', ip, {
    venueId,
    mode,
    imported: valid.length,
    duplicates: duplicates.length,
    conflicts: conflicts.length
  });

  res.json({
    ok: true,
    count: data.attendees.filter(a => a.venueId === venueId).length,
    imported: valid.length,
    duplicates: duplicates.length,
    conflicts: conflicts.length,
    duplicateNames: duplicates,
    conflictDetails: conflicts
  });
});

// 添加单个参会者（带重复和冲突检测）
app.post('/api/attendees', requireAuth, (req, res) => {
  const data = readData();
  const attendee = cleanAttendee(req.body);

  // 生成ID
  if (!attendee.id) {
    attendee.id = 'att-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
  }

  // 检查重复
  const duplicate = findDuplicateAttendee(data, attendee);
  if (duplicate) {
    return res.status(409).json({
      error: `参会者「${attendee.name}」在该会场已存在`,
      existing: duplicate
    });
  }

  // 检查座位冲突
  const conflict = findSeatConflict(data, attendee);
  if (conflict) {
    return res.status(409).json({
      error: `座位「${attendee.row} ${attendee.seat}号」已被「${conflict.name}」占用`,
      existing: conflict
    });
  }

  data.attendees.push(attendee);
  writeData(data);
  res.json({ ok: true });
});

// 删除会场
app.delete('/api/venues/:id', requireAdmin, (req, res) => {
  const data = readData();
  const venue = data.venues.find(v => v.id === req.params.id);
  if (!venue) return res.status(404).json({ error: '场馆不存在' });

  const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
  const attendeeCount = venueAttendees.length;

  // 记录被删除会场参会者信息到 deletedAttendees，防止自动分析时重新创建
  if (venueAttendees.length > 0) {
    venueAttendees.forEach(a => recordDeletedAttendee(data, a));
  }

  data.attendees = data.attendees.map(a => {
    if (a.venueId === venue.id) {
      return { ...a, venueId: '', row: '', seat: '' };
    }
    return a;
  });

  data.venues = data.venues.filter(v => v.id !== req.params.id);
  
  // 记录被删除的会场名称，防止自动分析时重新创建（同时保存原名称和规范化后的名称）
  if (!data.deletedVenueNames) data.deletedVenueNames = [];
  if (!data.deletedVenueNames.includes(venue.name)) {
    data.deletedVenueNames.push(venue.name);
  }
  // 同时保存规范化名称（去掉空格、统一大小写）用于匹配
  if (!data.deletedVenueNamesNormalized) data.deletedVenueNamesNormalized = [];
  const normalizedName = venue.name.trim().toLowerCase().replace(/\s+/g, '');
  if (!data.deletedVenueNamesNormalized.includes(normalizedName)) {
    data.deletedVenueNamesNormalized.push(normalizedName);
  }

  writeData(data);
  auditLog('DELETE_VENUE', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
    venueId: venue.id,
    venueName: venue.name,
    clearedAttendees: attendeeCount
  });

  res.json({ ok: true, clearedAttendees: attendeeCount });
});

// 修改会场名称/描述
app.put('/api/venues/:id', requireAdmin, (req, res) => {
  const data = readData();
  const venue = data.venues.find(v => v.id === req.params.id);
  if (!venue) return res.status(404).json({ error: '场馆不存在' });
  const { name, description } = req.body;
  if (name !== undefined) venue.name = name.trim();
  if (description !== undefined) venue.description = description.trim();
  writeData(data);
  res.json({ ok: true, venue });
});

// 应用增量变动
app.post('/api/apply-diffs', requireAdmin, (req, res) => {
  const { diffs } = req.body;
  if (!diffs || !Array.isArray(diffs)) {
    return res.status(400).json({ error: '缺少变动数据' });
  }

  const data = readData();
  let applied = 0;

  diffs.forEach(d => {
    // 查找对应的会场
    const venue = data.venues.find(v => v.name === d.venue);
    if (!venue) return;

    if (d.type === 'added') {
      // 新增：添加参会者
      const exists = data.attendees.find(a => a.venueId === venue.id && a.name === d.name);
      if (!exists) {
        data.attendees.push({
          name: d.name,
          row: d.newRow,
          seat: d.newSeat,
          company: '',
          title: '',
          venueId: venue.id,
          source: 'excel'
        });
        applied++;
      }
    } else if (d.type === 'removed') {
      // 移除：删除参会者
      const idx = data.attendees.findIndex(a => a.venueId === venue.id && a.name === d.name);
      if (idx !== -1) {
        data.attendees.splice(idx, 1);
        applied++;
      }
    } else if (d.type === 'moved') {
      // 换座：更新位置
      const attendee = data.attendees.find(a => a.venueId === venue.id && a.name === d.name);
      if (attendee) {
        attendee.row = d.newRow;
        attendee.seat = d.newSeat;
        applied++;
      }
    }
  });

  writeData(data);
  
  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
  auditLog('APPLY_DIFFS', ip, {
    totalDiffs: diffs.length,
    applied,
    summary: {
      added: diffs.filter(d => d.type === 'added').length,
      removed: diffs.filter(d => d.type === 'removed').length,
      moved: diffs.filter(d => d.type === 'moved').length
    }
  });
  
  res.json({ ok: true, applied: applied });
});

// 删除参会者
app.delete('/api/attendees', requireAuth, (req, res) => {
  const { name, venueId, id } = req.query;
  const data = readData();
  const beforeCount = data.attendees.length;

  // 先保存要删除的参会者信息，用于记录已删除的名单
  const toBeDeleted = [];
  if (id) {
    toBeDeleted.push(...data.attendees.filter(a => a.id === id));
    data.attendees = data.attendees.filter(a => a.id !== id);
  } else if (name && venueId) {
    toBeDeleted.push(...data.attendees.filter(a => a.name === name && a.venueId === venueId));
    data.attendees = data.attendees.filter(a => !(a.name === name && a.venueId === venueId));
  } else if (venueId) {
    toBeDeleted.push(...data.attendees.filter(a => a.venueId === venueId));
    data.attendees = data.attendees.filter(a => a.venueId !== venueId);
  } else {
    toBeDeleted.push(...data.attendees);
    data.attendees = [];
  }

  // 记录被删除的参会者信息，防止自动分析时重新创建
  if (!data.deletedAttendees) data.deletedAttendees = [];
  toBeDeleted.forEach(a => recordDeletedAttendee(data, a));

  const deletedCount = beforeCount - data.attendees.length;
  writeData(data);

  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
  auditLog('DELETE_ATTENDEES', ip, {
    id: id || null,
    name: name || null,
    venueId: venueId || null,
    deletedCount,
    scope: id ? 'single_id' : name && venueId ? 'single' : venueId ? 'venue' : 'all'
  });
  
  res.json({ ok: true, deleted: deletedCount });
});

// 批量删除参会者
app.post('/api/attendees/batch-delete', requireAuth, (req, res) => {
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: '请提供要删除的参会者ID列表' });
  }
  const data = readData();
  const beforeCount = data.attendees.length;
  const idsSet = new Set(ids);
  
  // 先保存要删除的参会者信息
  const toBeDeleted = data.attendees.filter(a => idsSet.has(a.id));
  data.attendees = data.attendees.filter(a => !idsSet.has(a.id));

  toBeDeleted.forEach(a => recordDeletedAttendee(data, a));

  const deletedCount = beforeCount - data.attendees.length;
  writeData(data);
  
  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
  auditLog('BATCH_DELETE_ATTENDEES', ip, {
    ids, deletedCount
  });
  
  res.json({ ok: true, deleted: deletedCount });
});

// 更新参会者信息
app.put('/api/attendees/:id', requireAuth, (req, res) => {
  const data = readData();
  const attendee = data.attendees.find(a => a.id === req.params.id);
  
  if (!attendee) {
    return res.status(404).json({ error: '参会者不存在' });
  }
  
  const { name, nameEn, company, title } = req.body;
  if (name) attendee.name = name.trim();
  if (nameEn !== undefined) attendee.nameEn = nameEn.trim();
  if (company !== undefined) attendee.company = company.trim();
  if (title !== undefined) attendee.title = title.trim();
  
  writeData(data);
  
  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
  auditLog('UPDATE_ATTENDEE', ip, {
    id: req.params.id,
    oldName: req.body.oldName || null,
    newName: name || null
  });
  
  res.json({ ok: true, attendee });
});

// 重新从服务端 Excel 导入布局
app.post('/api/import-excel', requireAdmin, (req, res) => {
  const excelPath = path.join(__dirname, '..', '座位排表.xlsx');
  if (!fs.existsSync(excelPath)) {
    return res.status(404).json({ error: '未找到 座位排表.xlsx' });
  }
  try {
    const mode = req.body.mode || 'auto';
    const sheetModes = req.body.sheetModes || {};
    const wb = XLSX.readFile(excelPath);
    const existing = readData();
    const manualAttendees = (existing.attendees || []).filter(a => a.source !== 'excel');
    const result = parseWorkbook(wb, manualAttendees, mode, sheetModes);
    
    // 新导入Excel：不继承旧删除记录，保留自定义会场
    const customVenues2 = (existing.venues || []).filter(v => v.layout === 'custom');
    result.venues = [...customVenues2, ...result.venues.filter(v => v.layout !== 'custom')];
    
    // ===== 过滤掉孤儿参会者（所属会场已被删除或不存在的）=====
    const finalVenueIds2b = new Set(result.venues.map(v => v.id));
    result.attendees = result.attendees.filter(a => finalVenueIds2b.has(a.venueId));
    
    // 新导入Excel：清空旧删除记录
    result.deletedVenueNames = [];
    result.deletedVenueNamesNormalized = [];
    result.deletedAttendees = [];
    
    writeData(result);
    res.json({ ok: true, data: result, mode: mode, sheetModes: sheetModes });
  } catch (err) {
    res.status(500).json({ error: '解析失败: ' + err.message });
  }
});

// 上传 Excel 文件导入（从浏览器上传）
// 文件上传配置
const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
const MAX_ATTENDEES_IMPORT = 5000; // 最多导入 5000 人

// 验证 Excel 文件
function validateExcelFile(fileData) {
  if (!fileData || typeof fileData !== 'string') {
    throw new Error('未提供文件数据');
  }
  
  // 检查 base64 大小
  const estimatedSize = fileData.length * 0.75;
  if (estimatedSize > MAX_FILE_SIZE) {
    throw new Error(`文件过大（${(estimatedSize / 1024 / 1024).toFixed(1)}MB），不超过 10MB`);
  }
  
  const buffer = Buffer.from(fileData, 'base64');
  
  // 检查实际大小
  if (buffer.length > MAX_FILE_SIZE) {
    throw new Error(`文件过大（${(buffer.length / 1024 / 1024).toFixed(1)}MB），不超过 10MB`);
  }
  
  // 检查 Excel 文件签名（ZIP 格式：PK）
  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    throw new Error('不是有效的 Excel 文件');
  }
  
  return buffer;
}

// 自动检测已上传 Excel 文件每个 Sheet 的布局类型
app.post('/api/detect-sheet-modes', requireAdmin, (req, res) => {
  const { fileData } = req.body;
  
  try {
    const buffer = validateExcelFile(fileData);
    const wb = XLSX.read(buffer, { type: 'buffer' });
    
    const sheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const detectedMode = detectSheetMode(data);
      return { name, detectedMode };
    });
    
    res.json({ ok: true, sheets });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.post('/api/upload-excel', requireAdmin, (req, res) => {
  const { fileData, keepManual, mode, sheetModes } = req.body;
  
  try {
    const buffer = validateExcelFile(fileData);
    const parseMode = mode || 'auto';
    const sheetModesMap = sheetModes || {};
    
    // 保存上传的文件到 uploaded.xlsx
    const uploadedPath = path.join(__dirname, 'uploads', 'uploaded.xlsx');
    fs.writeFileSync(uploadedPath, buffer);
    
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const existing = readData();
    const manualAttendees = keepManual
      ? (existing.attendees || []).filter(a => a.source !== 'excel')
      : [];
    const result = parseWorkbook(wb, manualAttendees, parseMode, sheetModesMap);
    
    // 新上传Excel：不继承旧删除记录，保留自定义会场
    const customVenues = (existing.venues || []).filter(v => v.layout === 'custom');
    result.venues = [...customVenues, ...result.venues.filter(v => v.layout !== 'custom')];
    
    // ===== 过滤掉孤儿参会者（所属会场已被删除或不存在的）=====
    const finalVenueIds3 = new Set(result.venues.map(v => v.id));
    result.attendees = result.attendees.filter(a => finalVenueIds3.has(a.venueId));
    
    // 验证导入数量
    if (result.attendees.length > MAX_ATTENDEES_IMPORT) {
      return res.status(413).json({ 
        error: `导入数据过大（${result.attendees.length} 人），不超过 ${MAX_ATTENDEES_IMPORT} 人` 
      });
    }
    
    // 新上传Excel：清空旧删除记录，开启新会话
    result.deletedVenueNames = [];
    result.deletedVenueNamesNormalized = [];
    result.deletedAttendees = [];
    
    writeData(result);
    auditLog('UPLOAD_EXCEL', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      fileName: 'uploaded.xlsx',
      venueCount: result.venues.length,
      attendeeCount: result.attendees.length,
      keepManual: keepManual,
      mode: parseMode,
      sheetModes: sheetModesMap
    });
    res.json({ ok: true, data: result });
  } catch (err) {
    structuredLog('error', { message: 'Excel 上传失败', error: err.message });
    res.status(400).json({ error: err.message });
  }
});

// ========== 自定义会场布局 API ==========

// 图片文件头魔数验证
const IMAGE_SIGNATURES = {
  png:  [0x89, 0x50, 0x4E, 0x47],
  jpg:  [0xFF, 0xD8, 0xFF],
  gif:  [0x47, 0x49, 0x46, 0x38],
  webp: [0x52, 0x49, 0x46, 0x46]
};

function detectImageType(buffer) {
  const head = buffer.slice(0, 4);
  for (const [type, sig] of Object.entries(IMAGE_SIGNATURES)) {
    if (sig.every((b, i) => head[i] === b)) return type;
  }
  if (buffer.length >= 12 &&
      head[0] === 0x52 && head[1] === 0x49 && head[2] === 0x46 && head[3] === 0x46 &&
      buffer[8] === 0x57 && buffer[9] === 0x45 && buffer[10] === 0x42 && buffer[11] === 0x50) {
    return 'webp';
  }
  return null;
}

// 上传自定义布局背景图
app.post('/api/custom-layout/upload-image', requireAdmin, (req, res) => {
  try {
    const { imageData } = req.body;
    if (!imageData) {
      return res.status(400).json({ error: '缺少图片数据' });
    }
    let base64Str = imageData;
    let ext = 'png';
    if (base64Str.startsWith('data:')) {
      const match = base64Str.match(/^data:image\/(\w+);base64,(.+)$/);
      if (match) {
        ext = match[1] === 'jpeg' ? 'jpg' : match[1];
        base64Str = match[2];
      }
    }
    const buffer = Buffer.from(base64Str, 'base64');
    if (buffer.length > 10 * 1024 * 1024) {
      return res.status(413).json({ error: '图片文件过大，最大支持 10MB' });
    }
    const detectedType = detectImageType(buffer);
    if (!detectedType || !['png', 'jpg', 'jpeg', 'gif', 'webp'].includes(detectedType)) {
      return res.status(400).json({ error: '不支持的图片格式，仅支持 JPG/PNG/GIF/WebP' });
    }
    const imagesDir = path.join(__dirname, 'uploads', 'images');
    if (!fs.existsSync(imagesDir)) {
      fs.mkdirSync(imagesDir, { recursive: true });
    }
    const filename = 'custom_' + Date.now() + '.' + detectedType;
    const filePath = path.join(imagesDir, filename);
    fs.writeFileSync(filePath, buffer);
    const imagePath = 'uploads/images/' + filename;
    auditLog('UPLOAD_CUSTOM_IMAGE', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      filename: filename, size: buffer.length
    });
    res.json({ success: true, imagePath });
  } catch (err) {
    structuredLog('error', { message: '自定义布局图片上传失败', error: err.message });
    res.status(400).json({ error: err.message });
  }
});

// 保存自定义布局
app.post('/api/custom-layout/save', requireAdmin, (req, res) => {
  try {
    const { venueName, description, stageName, backgroundImage, gridConfig, customRows, canvasWidth, canvasHeight, customStage, customGates, customAisles } = req.body;
    if (!venueName) return res.status(400).json({ error: '请输入会场名称' });
    if (!customRows || !customRows.length) return res.status(400).json({ error: '请至少定义一个排或列' });
    const data = readData();
    // 使用时间戳生成唯一ID，避免冲突
    const venueId = 'venue-custom-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
    const standardRows = customRows.map((cr, idx) => {
      const seats = [];
      const getSn = (i) => (cr.seatNumbers && cr.seatNumbers.length > i) ? cr.seatNumbers[i] : (cr.startSeat || 1) + i;
      const aislePositions = cr.aislePositions || [];
      // 根据aislePositions将座位分成多个group
      // aislePositions[i] = k 表示在第k个座位（0-based index）前插入过道
      let seatGroups = [];
      let currentGroup = [];
      for (let i = 0; i < cr.seatCount; i++) {
        if (aislePositions.indexOf(i) > -1 && currentGroup.length > 0) {
          seatGroups.push(currentGroup);
          currentGroup = [];
        }
        currentGroup.push(getSn(i));
      }
      if (currentGroup.length > 0) seatGroups.push(currentGroup);
      // 如果没有seatGroups（seatCount=0的情况），创建一个空group
      if (seatGroups.length === 0) {
        for (let i = 0; i < cr.seatCount; i++) seats.push(getSn(i));
        seatGroups = [seats];
      }
      return {
        label: cr.label || ('第' + (idx + 1) + '排'),
        rowNum: cr.rowNum || (idx + 1),
        seatGroups: seatGroups,
        isAisle: false,
        hasAisleAfter: false
      };
    });
    const totalSeats = standardRows.reduce((sum, r) => sum + (r.seatGroups[0] || []).length, 0);
    const venue = {
      id: venueId,
      name: venueName,
      description: description || venueName,
      stageName: '',
      mode: 'custom',
      layout: 'custom',
      totalSeats,
      rows: standardRows,
      backgroundImage: backgroundImage || '',
      gridConfig: gridConfig || { cellWidth: 50, cellHeight: 50, offsetX: 0, offsetY: 0, seatSize: 40, seatGap: 8 },
      customRows: customRows,
      canvasWidth: canvasWidth || 800,
      canvasHeight: canvasHeight || 600,
      customStage: customStage,
      customGates: customGates || [],
      customAisles: customAisles || []
    };
    data.venues.push(venue);
    writeData(data);
    auditLog('SAVE_CUSTOM_LAYOUT', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      venueId, venueName, rowCount: customRows.length, totalSeats
    });
    res.json({ success: true, venue });
  } catch (err) {
    structuredLog('error', { message: '自定义布局保存失败', error: err.message });
    res.status(400).json({ error: err.message });
  }
});

// 对比 Excel：上传新 Excel 与当前系统数据对比，返回差异
app.post('/api/compare-excel', requireAdmin, (req, res) => {
  const { fileData, mode, sheetModes } = req.body;
  
  try {
    const buffer = validateExcelFile(fileData);
    const parseMode = mode || 'auto';
    const sheetModesMap = sheetModes || {};
    
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const newData = parseWorkbook(wb, [], parseMode, sheetModesMap);
    const oldData = readData();

    const diffs = [];

    // 遍历每个会场做对比
    newData.venues.forEach(newVenue => {
      const oldVenue = oldData.venues.find(v => v.name === newVenue.name);
      const venueName = newVenue.name;
      const newAttendees = newData.attendees.filter(a => a.venueId === newVenue.id);

      if (!oldVenue) {
        // 整个会场是新增的
        newAttendees.forEach(a => {
          diffs.push({ type: 'added', venue: venueName, name: a.name, newRow: a.row, newSeat: a.seat, oldRow: '', oldSeat: '' });
        });
        return;
      }

      const oldAttendees = oldData.attendees.filter(a => a.venueId === oldVenue.id);

      // 建立 name -> attendee 的映射
      const oldMap = {};
      oldAttendees.forEach(a => { oldMap[a.name] = a; });
      const newMap = {};
      newAttendees.forEach(a => { newMap[a.name] = a; });

      // 新增的人（新表有、旧表没有）
      newAttendees.forEach(a => {
        if (!oldMap[a.name]) {
          diffs.push({ type: 'added', venue: venueName, name: a.name, newRow: a.row, newSeat: a.seat, oldRow: '', oldSeat: '' });
        }
      });

      // 删除的人（旧表有、新表没有）
      oldAttendees.forEach(a => {
        if (!newMap[a.name]) {
          diffs.push({ type: 'removed', venue: venueName, name: a.name, newRow: '', newSeat: '', oldRow: a.row, oldSeat: a.seat });
        }
      });

      // 位置变动的人
      newAttendees.forEach(a => {
        const old = oldMap[a.name];
        if (old && (old.row !== a.row || old.seat !== a.seat)) {
          diffs.push({ type: 'moved', venue: venueName, name: a.name, oldRow: old.row, oldSeat: old.seat, newRow: a.row, newSeat: a.seat });
        }
      });
    });

    // 检查旧数据中有但新数据里整个会场都没有的
    oldData.venues.forEach(oldVenue => {
      const newVenue = newData.venues.find(v => v.name === oldVenue.name);
      if (!newVenue) {
        const oldAttendees = oldData.attendees.filter(a => a.venueId === oldVenue.id);
        oldAttendees.forEach(a => {
          diffs.push({ type: 'removed', venue: oldVenue.name, name: a.name, newRow: '', newSeat: '', oldRow: a.row, oldSeat: a.seat });
        });
      }
    });

    res.json({ ok: true, diffs, summary: {
      total: diffs.length,
      added: diffs.filter(d => d.type === 'added').length,
      removed: diffs.filter(d => d.type === 'removed').length,
      moved: diffs.filter(d => d.type === 'moved').length
    }});
  } catch (err) {
    res.status(500).json({ error: '对比失败: ' + err.message });
  }
});

// ==================== 座位安排可视化 API ====================

// 获取会场座位状态（返回已占座位和空座位）
app.get('/api/venues/:id/seating', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    const venueAttendees = data.attendees.filter(a => a.venueId === venue.id);
    const attendeeMap = {};
    venueAttendees.forEach(a => {
      const key = `${a.row}_${String(a.seat)}`;
      attendeeMap[key] = a;
    });

    // 构建座位状态矩阵
    const seating = {
      venue,
      attendees: venueAttendees,
      seatOccupancy: {}
    };

    if (venue.rows) {
      venue.rows.forEach(row => {
        if (row.seatGroups) {
          row.seatGroups.forEach(group => {
            group.forEach(seatNum => {
              const key = `${row.label}_${String(seatNum)}`;
              seating.seatOccupancy[key] = attendeeMap[key] || null;
            });
          });
        }
      });
    }

    // 自定义布局：也从 customRows 填充 seatOccupancy
    if (venue.customRows) {
      venue.customRows.forEach(row => {
        for (let i = 0; i < (row.seatCount || 0); i++) {
          const seatNum = (row.seatNumbers && row.seatNumbers.length > i) ? row.seatNumbers[i] : (row.startSeat || 1) + i;
          const key = `${row.label}_${String(seatNum)}`;
          if (!(key in seating.seatOccupancy)) {
            seating.seatOccupancy[key] = attendeeMap[key] || null;
          }
        }
      });
    }

    res.json({ ok: true, seating });
  } catch (err) {
    structuredLog('error', { message: '获取座位状态失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '获取失败: ' + err.message });
  }
});

// 更新单个参会者座位（拖拽用）
app.put('/api/attendees/:id/seat', requireAuth, (req, res) => {
  try {
    const data = readData();
    const attendee = data.attendees.find(a => a.id === req.params.id);
    if (!attendee) return res.status(404).json({ error: '参会者不存在' });

    const { venueId, row, seat } = req.body;
    const venue = data.venues.find(v => v.id === venueId);
    if (!venue) return res.status(404).json({ error: '会场不存在' });

    // 检查目标座位是否已被占用
    const existing = data.attendees.find(
      a => a.venueId === venueId &&
           a.row === row &&
           a.seat === seat &&
           a.id !== attendee.id
    );

    if (existing) {
      return res.status(409).json({
        error: `座位「${row} ${seat}号」已被「${existing.name}」占用`,
        occupiedBy: existing
      });
    }

    // 更新座位
    attendee.venueId = venueId;
    attendee.row = row;
    attendee.seat = seat;

    writeData(data);
    auditLog('UPDATE_ATTENDEE_SEAT', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      attendeeId: attendee.id,
      name: attendee.name,
      newVenueId: venueId,
      newRow: row,
      newSeat: seat
    });

    res.json({ ok: true, attendee });
  } catch (err) {
    structuredLog('error', { message: '更新座位失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '更新失败: ' + err.message });
  }
});

// 将参会者安排到指定座位（如果座位被占用则交换）
app.post('/api/venues/:id/assign-seat', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    const { attendeeId, row, seat } = req.body;
    const attendee = data.attendees.find(a => a.id === attendeeId);
    if (!attendee) return res.status(404).json({ error: '参会者不存在' });

    // 检查目标座位是否被占用（统一转为字符串比较，避免数字/字符串类型不匹配）
    const existing = data.attendees.find(
      a => a.venueId === venue.id &&
           String(a.row) === String(row) &&
           String(a.seat) === String(seat)
    );

    if (existing) {
      // 交换座位
      const oldRow = attendee.row;
      const oldSeat = attendee.seat;
      const oldVenueId = attendee.venueId;

      attendee.venueId = venue.id;
      attendee.row = row;
      attendee.seat = seat;

      existing.venueId = oldVenueId;
      existing.row = oldRow;
      existing.seat = oldSeat;

      auditLog('SWAP_SEATS', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
        venueId: venue.id,
        attendee1: { id: attendee.id, name: attendee.name, fromRow: oldRow, fromSeat: oldSeat, toRow: row, toSeat: seat },
        attendee2: { id: existing.id, name: existing.name, fromRow: row, fromSeat: seat, toRow: oldRow, toSeat: oldSeat }
      });
    } else {
      // 直接安排
      attendee.venueId = venue.id;
      attendee.row = row;
      attendee.seat = seat;

      auditLog('ASSIGN_SEAT', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
        venueId: venue.id,
        attendeeId: attendee.id,
        name: attendee.name,
        row: row,
        seat: seat
      });
    }

    writeData(data);
    res.json({ ok: true, venueAttendees: data.attendees.filter(a => a.venueId === venue.id) });
  } catch (err) {
    structuredLog('error', { message: '安排座位失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '安排失败: ' + err.message });
  }
});

// 清除参会者的座位
app.put('/api/attendees/:id/clear-seat', requireAuth, (req, res) => {
  try {
    const data = readData();
    const attendee = data.attendees.find(a => a.id === req.params.id);
    if (!attendee) return res.status(404).json({ error: '参会者不存在' });

    attendee.venueId = '';
    attendee.row = '';
    attendee.seat = '';

    writeData(data);
    auditLog('CLEAR_SEAT', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      attendeeId: attendee.id,
      name: attendee.name
    });

    res.json({ ok: true, attendee });
  } catch (err) {
    structuredLog('error', { message: '清除座位失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '清除失败: ' + err.message });
  }
});

// 清空会场所有座位
app.delete('/api/venues/:id/seating', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    const count = data.attendees.filter(a => a.venueId === venue.id).length;

    data.attendees = data.attendees.map(a => {
      if (a.venueId === venue.id) {
        return { ...a, venueId: '', row: '', seat: '' };
      }
      return a;
    });

    writeData(data);
    auditLog('CLEAR_VENUE_SEATING', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      venueId: venue.id,
      venueName: venue.name,
      clearedCount: count
    });

    res.json({ ok: true, clearedCount: count });
  } catch (err) {
    structuredLog('error', { message: '清空会场座位失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '清空失败: ' + err.message });
  }
});

// 随机排座
app.post('/api/venues/:id/random-seat', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    // 获取该会场所有未安排座位的参会者
    const unassigned = data.attendees.filter(a => !a.venueId || !a.row || !a.seat);
    
    // 获取该会场所有座位
    const allSeats = [];
    if (venue.rows) {
      venue.rows.forEach(row => {
        if (row.seatGroups) {
          row.seatGroups.forEach(group => {
            group.forEach(seatNum => {
              allSeats.push({ row: row.label, seat: seatNum });
            });
          });
        }
      });
    }

    // 获取已占用的座位
    const occupied = new Set();
    data.attendees.forEach(a => {
      if (a.venueId === venue.id && a.row && a.seat) {
        occupied.add(`${a.row}_${a.seat}`);
      }
    });

    // 获取空座位
    const emptySeats = allSeats.filter(s => !occupied.has(`${s.row}_${s.seat}`));
    
    // 打乱顺序
    emptySeats.sort(() => Math.random() - 0.5);
    unassigned.sort(() => Math.random() - 0.5);

    // 分配座位
    let assigned = 0;
    const min = Math.min(unassigned.length, emptySeats.length);
    for (let i = 0; i < min; i++) {
      unassigned[i].venueId = venue.id;
      unassigned[i].row = emptySeats[i].row;
      unassigned[i].seat = emptySeats[i].seat;
      assigned++;
    }

    writeData(data);
    auditLog('RANDOM_SEATING', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      venueId: venue.id,
      venueName: venue.name,
      assignedCount: assigned,
      unassignedCount: unassigned.length - assigned
    });

    res.json({ 
      ok: true, 
      assignedCount: assigned, 
      remainingCount: Math.max(0, unassigned.length - assigned) 
    });
  } catch (err) {
    structuredLog('error', { message: '随机排座失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '随机排座失败: ' + err.message });
  }
});

// 批量导入参会者（Excel文件或纯文本名单）
app.post('/api/attendees/import-names', requireAuth, (req, res) => {
  try {
    const { names, venueId, fileData } = req.body;
    const data = readData();
    let attendeesToImport = [];

    // 如果是Excel文件，先解析
    if (fileData) {
      const buffer = Buffer.from(fileData, 'base64');
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

      let startRow = 0;
      // 寻找数据起始行
      for (let i = 0; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (row.length > 0 && row[0] && (typeof row[0] === 'string' && (row[0].includes('姓名') || row[0].includes('name')))) {
          startRow = i + 1;
          break;
        }
      }

      // 解析每行数据
      for (let i = startRow; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (!row || row.length === 0) continue;
        
        const name = row[0] ? String(row[0]).trim() : '';
        if (!name) continue;

        const company = row[1] ? String(row[1]).trim() : '';
        const title = row[2] ? String(row[2]).trim() : '';
        const rowLabel = row[3] ? String(row[3]).trim() : '';
        const seatNum = row[4] ? String(row[4]).trim() : '';

        attendeesToImport.push({
          name,
          company,
          title,
          row: rowLabel,
          seat: seatNum
        });
      }
    }
    // 如果是纯文本名单
    else if (Array.isArray(names) && names.length > 0) {
      attendeesToImport = names.map(item => {
        if (typeof item === 'string') {
          return { name: item, company: '' };
        }
        return {
          name: item.name || '',
          company: item.company || '',
          title: item.title || '',
          row: item.row || '',
          seat: item.seat || ''
        };
      });
    }

    if (attendeesToImport.length === 0) {
      return res.status(400).json({ error: '没有有效的参会者数据' });
    }

    if (attendeesToImport.length > 500) {
      return res.status(400).json({ error: '单次导入不超过 500 人' });
    }

    const imported = [];
    const duplicates = [];

    attendeesToImport.forEach(item => {
      const trimmedName = item.name.trim();
      if (!trimmedName) return;

      // 检查重复
      if (data.attendees.some(a => a.name === trimmedName)) {
        duplicates.push(trimmedName);
        return;
      }

      const newAttendee = {
        id: 'att-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9),
        name: trimmedName,
        company: item.company || '',
        title: item.title || '',
        venueId: (item.row && item.seat) ? venueId || '' : '',
        row: item.row || '',
        seat: item.seat || '',
        source: 'manual'
      };

      data.attendees.push(newAttendee);
      imported.push(newAttendee);
    });

    writeData(data);
    auditLog('IMPORT_NAMES', req.headers['x-forwarded-for'] || req.connection.remoteAddress, {
      importedCount: imported.length,
      duplicateCount: duplicates.length
    });

    res.json({ ok: true, importedCount: imported.length, duplicates, attendees: data.attendees });
  } catch (err) {
    structuredLog('error', { message: '导入名单失败', error: err.message, stack: err.stack });
    res.status(500).json({ error: '导入失败: ' + err.message });
  }
});

// ==================== 备份管理 API ====================

// 获取备份列表
app.get('/api/backups', requireAdmin, (req, res) => {
  if (!fs.existsSync(BACKUP_DIR)) {
    return res.json({ backups: [] });
  }
  const backups = fs.readdirSync(BACKUP_DIR)
    .filter(f => f.startsWith('data-') && f.endsWith('.json'))
    .sort()
    .reverse()
    .map(f => {
      const filePath = path.join(BACKUP_DIR, f);
      const stat = fs.statSync(filePath);
      return {
        filename: f,
        size: stat.size,
        createdAt: stat.birthtime.toISOString()
      };
    });
  res.json({ backups });
});

// 获取访问日志
app.get('/api/logs', requireAdmin, (req, res) => {
  const lines = parseInt(req.query.lines) || 100;
  if (!fs.existsSync(accessLogFile)) {
    return res.json({ logs: [] });
  }
  try {
    const content = fs.readFileSync(accessLogFile, 'utf-8');
    const allLines = content.trim().split('\n').filter(l => l);
    const recentLines = allLines.slice(-lines);
    res.json({ logs: recentLines, total: allLines.length });
  } catch (err) {
    res.json({ logs: [], total: 0 });
  }
});

// 获取审计日志
app.get('/api/audit-logs', requireAdmin, (req, res) => {
  const lines = parseInt(req.query.lines) || 100;
  if (!fs.existsSync(auditLogFile)) {
    return res.json({ logs: [] });
  }
  try {
    const content = fs.readFileSync(auditLogFile, 'utf-8');
    const allLines = content.trim().split('\n').filter(l => l);
    const recentLines = allLines.slice(-lines).map(l => {
      try { return JSON.parse(l); } catch { return { raw: l }; }
    });
    res.json({ logs: recentLines.reverse(), total: allLines.length });
  } catch (err) {
    res.json({ logs: [], total: 0 });
  }
});

// 清空访问日志
app.delete('/api/logs', requireAdmin, (req, res) => {
  try {
    fs.writeFileSync(accessLogFile, '', 'utf-8');
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: '清空失败: ' + err.message });
  }
});

// 恢复备份
app.post('/api/backups/restore', requireAdmin, (req, res) => {
  const { filename } = req.body;
  if (!filename) {
    return res.status(400).json({ error: '缺少 filename 参数' });
  }
  const backupPath = path.join(BACKUP_DIR, filename);
  if (!fs.existsSync(backupPath)) {
    return res.status(404).json({ error: '备份文件不存在' });
  }
  try {
    // 先备份当前数据
    backupData();
    // 恢复
    const data = JSON.parse(fs.readFileSync(backupPath, 'utf-8'));
    writeData(data, true); // 跳过备份，避免循环
    res.json({ ok: true, message: `已从 ${filename} 恢复` });
  } catch (err) {
    res.status(500).json({ error: '恢复失败: ' + err.message });
  }
});

// ==================== 区域（单位就坐区）管理 ====================

// 保存会场的区域数据
app.post('/api/venues/:id/regions', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    const { regions } = req.body;
    if (!Array.isArray(regions)) {
      return res.status(400).json({ error: 'regions 必须是数组' });
    }

    venue.regions = regions;
    writeData(data);
    res.json({ ok: true });
  } catch (err) {
    structuredLog('error', { message: '保存区域失败', error: err.message });
    res.status(500).json({ error: '保存失败: ' + err.message });
  }
});

// 获取会场的区域数据
app.get('/api/venues/:id/regions', requireAuth, (req, res) => {
  try {
    const data = readData();
    const venue = data.venues.find(v => v.id === req.params.id);
    if (!venue) return res.status(404).json({ error: '场馆不存在' });

    res.json({ ok: true, regions: venue.regions || [] });
  } catch (err) {
    structuredLog('error', { message: '获取区域失败', error: err.message });
    res.status(500).json({ error: '获取失败: ' + err.message });
  }
});

// ==================== 错误处理 ====================

// 404 处理
app.use((req, res) => {
  res.status(404).json({ error: '接口不存在' });
});

// 全局错误处理
app.use((err, req, res, next) => {
  structuredLog('error', { message: '未捕获的错误', error: err.message, stack: err.stack, url: req.url });
  res.status(500).json({ error: '服务器内部错误' });
});

// ==================== 启动服务 ====================

// 优雅关闭
function gracefulShutdown(signal) {
  structuredLog('system', { message: `收到 ${signal} 信号，正在优雅关闭...` });
  // 这里可以添加关闭数据库连接等清理操作
  process.exit(0);
}

process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
process.on('SIGINT', () => gracefulShutdown('SIGINT'));

app.listen(PORT, () => {
  structuredLog('system', { message: `座位查询系统已启动：http://localhost:${PORT}` });
  structuredLog('system', { message: `管理后台：http://localhost:${PORT}/admin.html` });
  structuredLog('system', { message: `用户查询：http://localhost:${PORT}/` });
  structuredLog('system', { message: `健康检查：http://localhost:${PORT}/api/health` });

  // 首次启动提示
  const config = readConfig();
  if (config.adminPassword === bcrypt.hashSync('admin888', SALT_ROUNDS)) {
    structuredLog('error', { message: '检测到使用默认密码，请及时修改 config.json 中的密码' });
  }
});

// 导出供 Vercel 使用
module.exports = app;
