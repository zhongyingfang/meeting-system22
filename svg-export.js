'use strict';

/**
 * 座位图导出渲染模块
 * 输入是本系统生成的 SVG（只含 rect / circle / text 三种图元），输出 SVG（可选文字转曲）、PDF、PNG。
 *
 * 设计要点：
 * 1. 转曲与 PDF 共用同一套字形轮廓提取，PDF 因此不需要嵌入字体，任何机器打开都不会变形。
 * 2. 所有坐标继续沿用 SVG 的像素坐标系（y 轴向下），PDF 里用一次 cm 变换翻转，避免逐个元素换算。
 * 3. fontkit / sharp 都是可选依赖：缺任一不影响原有 SVG 导出，只关闭对应格式并给出明确错误。
 */

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

let fontkit = null;
try { fontkit = require('fontkit'); } catch (e) { /* 可选 */ }
let sharp = null;
try { sharp = require('sharp'); } catch (e) { /* 可选 */ }

const FONT_CANDIDATES = {
  regular: [
    process.env.SEAT_OUTLINE_FONT,
    'C:/Windows/Fonts/msyh.ttc',
    'C:/Windows/Fonts/msyh.ttf',
    'C:/Windows/Fonts/simhei.ttf',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
    '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
    '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
    '/usr/share/fonts/truetype/arphic/uming.ttc',
    '/System/Library/Fonts/PingFang.ttc',
    '/System/Library/Fonts/STHeiti Medium.ttc'
  ],
  bold: [
    process.env.SEAT_OUTLINE_FONT_BOLD,
    'C:/Windows/Fonts/msyhbd.ttc',
    'C:/Windows/Fonts/msyhbd.ttf',
    'C:/Windows/Fonts/simhei.ttf',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Bold.otf',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc',
    '/System/Library/Fonts/PingFang.ttc'
  ]
};

// 项目自带字体兜底：部署包里 fonts/ 目录带了 msyh/simhei，系统没装中文字体时仍可转曲与导 PDF
const bundledFont = (name) => path.join(__dirname, 'fonts', name);
FONT_CANDIDATES.regular.push(bundledFont('msyh.ttc'), bundledFont('simhei.ttf'), bundledFont('simsun.ttc'));
FONT_CANDIDATES.bold.push(bundledFont('msyh.ttc'), bundledFont('msyhbd.ttc'), bundledFont('simhei.ttf'));

const fontCache = {};

/** 从字体集合（.ttc）中挑出合适的一个字面 */
function pickFace(collection, wantBold) {
  if (!collection || !collection.fonts) return collection;
  const list = collection.fonts;
  // 优先非 UI 变体（Microsoft YaHei UI 的字宽与桌面版略有差异）
  let pool = list.filter((f) => !/\bUI\b/i.test(f.postscriptName || ''));
  if (!pool.length) pool = list;
  const isBold = (f) => /bold/i.test(f.postscriptName || '') || /bold/i.test(f.familyName || '');
  return pool.find((f) => isBold(f) === wantBold) || pool[0];
}

/** 加载字体，失败返回 null（由调用方决定如何降级） */
function getFont(weight) {
  if (!fontkit) return null;
  const wantBold = Number(weight) >= 600;
  const key = wantBold ? 'bold' : 'regular';
  if (key in fontCache) return fontCache[key];

  const order = wantBold ? ['bold', 'regular'] : ['regular'];
  for (const kind of order) {
    for (const p of FONT_CANDIDATES[kind]) {
      if (!p) continue;
      try {
        if (!fs.existsSync(p)) continue;
        const coll = fontkit.openSync(p);
        const font = pickFace(coll, kind === 'bold');
        if (font && font.unitsPerEm) {
          fontCache[key] = { font, file: p };
          return fontCache[key];
        }
      } catch (e) { /* 试下一个 */ }
    }
  }
  fontCache[key] = null;
  return null;
}

function fontInfo() {
  const hit = getFont(400);
  return hit ? { available: true, file: hit.file, family: hit.font.familyName || '' } : { available: false };
}

/* ------------------------------------------------------------------ *
 * 基础工具
 * ------------------------------------------------------------------ */

const XML_ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

function decodeXml(s) {
  return String(s).replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, code) => {
    if (code[0] === '#') {
      const n = code[1] === 'x' || code[1] === 'X'
        ? parseInt(code.slice(2), 16)
        : parseInt(code.slice(1), 10);
      return Number.isFinite(n) ? String.fromCodePoint(n) : m;
    }
    return XML_ENTITIES[code] !== undefined ? XML_ENTITIES[code] : m;
  });
}

function parseAttrs(raw) {
  const out = {};
  const re = /([\w:-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(raw))) out[m[1]] = m[2];
  return out;
}

function parseLength(v, base) {
  if (v === undefined || v === null || v === '') return 0;
  const s = String(v).trim();
  if (s.endsWith('%')) {
    const p = parseFloat(s);
    return Number.isFinite(p) ? (base * p) / 100 : 0;
  }
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

/** 解析 <style> 块里的类选择器 */
function parseStyleBlock(svg) {
  const styles = {};
  const block = svg.match(/<style[^>]*>([\s\S]*?)<\/style>/);
  if (!block) return styles;
  const body = block[1].replace(/\/\*[\s\S]*?\*\//g, '');
  const ruleRe = /([^{}]+)\{([^}]*)\}/g;
  let m;
  while ((m = ruleRe.exec(body))) {
    const decls = {};
    for (const part of m[2].split(';')) {
      const i = part.indexOf(':');
      if (i < 0) continue;
      const k = part.slice(0, i).trim().toLowerCase();
      const v = part.slice(i + 1).trim();
      if (k) decls[k] = v;
    }
    for (const sel of m[1].split(',')) {
      const name = sel.trim().replace(/^\./, '');
      if (name) styles[name] = Object.assign(styles[name] || {}, decls);
    }
  }
  return styles;
}

const DEFAULT_PROPS = {
  fill: '#000000',
  stroke: 'none',
  'stroke-width': '1',
  'font-size': '16',
  'font-weight': '400',
  'text-anchor': 'start',
  'dominant-baseline': 'auto'
};

/** 合并类样式与行内属性，得到元素的最终表现属性 */
function resolveProps(attrs, styles) {
  const props = Object.assign({}, DEFAULT_PROPS);
  for (const cls of String(attrs.class || '').split(/\s+/)) {
    if (cls && styles[cls]) Object.assign(props, styles[cls]);
  }
  for (const k of Object.keys(DEFAULT_PROPS)) {
    if (attrs[k] !== undefined) props[k] = attrs[k];
  }
  if (attrs.fill !== undefined) props.fill = attrs.fill;
  if (attrs.stroke !== undefined) props.stroke = attrs.stroke;
  return props;
}

/* ------------------------------------------------------------------ *
 * 文字转轮廓
 * ------------------------------------------------------------------ */

/**
 * 把一段文字转成最终坐标下的绘制命令
 * 返回 [{ op: 'M'|'L'|'Q'|'C'|'Z', pts: [{x,y}...] }]
 */
function textToCommands(text, x, y, fontSize, anchor, font, baselineMode) {
  const cmds = [];
  if (!text) return cmds;
  const em = font.unitsPerEm || 1000;
  const scale = fontSize / em;
  const run = font.layout(text);

  let totalAdvance = 0;
  for (const pos of run.positions) totalAdvance += pos.xAdvance;
  totalAdvance *= scale;

  let startX = x;
  if (anchor === 'middle') startX = x - totalAdvance / 2;
  else if (anchor === 'end' || anchor === 'end ') startX = x - totalAdvance;

  // dominant-baseline="central"/"middle" 时，传入的 y 是盒中心而非基线
  let baselineY = y;
  if (baselineMode === 'central' || baselineMode === 'middle') {
    const ascent = font.ascent || em * 0.88;
    const descent = font.descent || -em * 0.12;
    baselineY = y + ((ascent + descent) / 2) * scale;
  }

  let pen = 0;
  for (let i = 0; i < run.glyphs.length; i++) {
    const glyph = run.glyphs[i];
    const pos = run.positions[i];
    if (!glyph || !glyph.path) { pen += (pos ? pos.xAdvance : 0) * scale; continue; }
    const gx = startX + pen + (pos.xOffset || 0) * scale;
    const gy = baselineY - (pos.yOffset || 0) * scale;

    for (const cmd of glyph.path.commands) {
      const a = cmd.args;
      switch (cmd.command) {
        case 'moveTo':
          cmds.push({ op: 'M', pts: [{ x: gx + a[0] * scale, y: gy - a[1] * scale }] });
          break;
        case 'lineTo':
          cmds.push({ op: 'L', pts: [{ x: gx + a[0] * scale, y: gy - a[1] * scale }] });
          break;
        case 'quadraticCurveTo':
          cmds.push({
            op: 'Q',
            pts: [
              { x: gx + a[0] * scale, y: gy - a[1] * scale },
              { x: gx + a[2] * scale, y: gy - a[3] * scale }
            ]
          });
          break;
        case 'bezierCurveTo':
          cmds.push({
            op: 'C',
            pts: [
              { x: gx + a[0] * scale, y: gy - a[1] * scale },
              { x: gx + a[2] * scale, y: gy - a[3] * scale },
              { x: gx + a[4] * scale, y: gy - a[5] * scale }
            ]
          });
          break;
        case 'closePath':
          cmds.push({ op: 'Z', pts: [] });
          break;
        default:
          break;
      }
    }
    pen += (pos ? pos.xAdvance : 0) * scale;
  }
  return cmds;
}

const num = (n) => {
  const r = Math.round(n * 100) / 100;
  return Number.isFinite(r) ? String(r) : '0';
};

function commandsToSvgPath(cmds) {
  let d = '';
  for (const c of cmds) {
    if (c.op === 'Z') { d += 'Z'; continue; }
    d += c.op + c.pts.map((p) => num(p.x) + ' ' + num(p.y)).join(' ');
  }
  return d;
}

/* ------------------------------------------------------------------ *
 * SVG 解析
 * ------------------------------------------------------------------ */

function svgSize(svg) {
  const root = svg.match(/<svg\b([^>]*)>/);
  const attrs = root ? parseAttrs(root[1]) : {};
  let w = parseFloat((attrs.width || '').replace(/px$/, ''));
  let h = parseFloat((attrs.height || '').replace(/px$/, ''));
  const vb = (attrs.viewBox || '').split(/[\s,]+/).map(Number);
  let x0 = 0;
  let y0 = 0;
  if (vb.length === 4 && vb.every((n) => Number.isFinite(n))) {
    x0 = vb[0];
    y0 = vb[1];
    if (!Number.isFinite(w) || !Number.isFinite(h)) {
      w = vb[2]; h = vb[3];
    }
  }
  return { width: Math.ceil(w || 1000), height: Math.ceil(h || 1000), x0, y0 };
}

/**
 * 解析 SVG 为绘制图元列表（保持文档顺序）
 */
function parseItems(svg) {
  const { width, height, x0, y0 } = svgSize(svg);
  const styles = parseStyleBlock(svg);
  const items = [];

  const tagRe = /<(rect|circle|text)\b([^>]*?)\/>|<(rect|circle|text)\b([^>]*?)>([\s\S]*?)<\/\3>/g;
  let m;
  while ((m = tagRe.exec(svg))) {
    const name = m[1] || m[3];
    const attrRaw = m[2] !== undefined ? m[2] : m[4];
    const inner = m[5];
    const attrs = parseAttrs(attrRaw || '');
    const props = resolveProps(attrs, styles);
    const fill = props.fill && props.fill !== 'none' ? props.fill : null;
    const stroke = props.stroke && props.stroke !== 'none' ? props.stroke : null;
    const strokeWidth = parseFloat(props['stroke-width']) || 0;

    if (name === 'rect') {
      // 满幅底色：按当前画布铺满，不参与坐标偏移
      if (String(attrs.width).includes('%') && String(attrs.height).includes('%')) {
        items.push({ type: 'rect', background: true, x: 0, y: 0, w: width, h: height, rx: 0, fill, stroke, strokeWidth });
        continue;
      }
      items.push({
        type: 'rect',
        x: parseLength(attrs.x, width) - x0,
        y: parseLength(attrs.y, height) - y0,
        w: parseLength(attrs.width, width),
        h: parseLength(attrs.height, height),
        rx: parseLength(attrs.rx, width),
        fill, stroke, strokeWidth
      });
    } else if (name === 'circle') {
      items.push({
        type: 'circle',
        cx: (parseFloat(attrs.cx) || 0) - x0,
        cy: (parseFloat(attrs.cy) || 0) - y0,
        r: parseFloat(attrs.r) || 0,
        fill, stroke, strokeWidth
      });
    } else if (name === 'text') {
      const text = decodeXml(inner || '').replace(/\s+/g, ' ').trim();
      if (!text) continue;
      const fontSize = parseFloat(props['font-size']) || 16;
      const weight = props['font-weight'];
      const hit = getFont(weight);
      if (!hit) continue;
      const cmds = textToCommands(
        text,
        (parseFloat(attrs.x) || 0) - x0,
        (parseFloat(attrs.y) || 0) - y0,
        fontSize,
        props['text-anchor'],
        hit.font,
        props['dominant-baseline']
      );
      if (!cmds.length) continue;
      items.push({ type: 'path', cmds, fill: fill || '#000000', stroke: null, strokeWidth: 0 });
    }
  }
  return { width, height, items };
}

/* ------------------------------------------------------------------ *
 * 画布裁切（只改 viewBox，不动任何元素坐标）
 * ------------------------------------------------------------------ */

function inkBBox(parsed) {
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const add = (x, y) => {
    if (x < minX) minX = x;
    if (x > maxX) maxX = x;
    if (y < minY) minY = y;
    if (y > maxY) maxY = y;
  };
  for (const it of parsed.items) {
    if (it.background) continue;
    if (it.type === 'rect') {
      // 顶部的会场名称条是按整幅画布铺开的装饰元素，裁切时忽略它；
      // 裁完后它依然横跨整个可视宽度，视觉上不受影响
      if (it.w > parsed.width * 0.8 && it.h <= 200) continue;
      add(it.x, it.y); add(it.x + it.w, it.y + it.h);
    } else if (it.type === 'circle') { add(it.cx - it.r, it.cy - it.r); add(it.cx + it.r, it.cy + it.r); }
    else if (it.type === 'path') { for (const c of it.cmds) for (const p of c.pts) add(p.x, p.y); }
  }
  if (!Number.isFinite(minX)) return null;
  return { minX, minY, maxX, maxY };
}

/**
 * 按实际内容裁掉四周空白。只重写根节点的 width/height/viewBox，
 * 元素坐标保持原样，因此不会影响任何现有渲染逻辑。
 */
function trimSvg(svg, pad = 40) {
  const parsed = parseItems(svg);
  const bb = inkBBox(parsed);
  if (!bb) return svg;
  const x0 = Math.max(0, Math.floor(bb.minX - pad));
  const y0 = Math.max(0, Math.floor(bb.minY - pad));
  const x1 = Math.ceil(bb.maxX + pad);
  const y1 = Math.ceil(bb.maxY + pad);
  const w = Math.max(1, x1 - x0);
  const h = Math.max(1, y1 - y0);

  const trimmed = svg.replace(/<svg\b[^>]*>/, (tag) => {
    const attrs = parseAttrs(tag);
    let out = '<svg';
    for (const k of Object.keys(attrs)) {
      if (/^(width|height|viewbox)$/i.test(k)) continue;
      out += ` ${k}="${attrs[k]}"`;
    }
    return `${out} width="${w}" height="${h}" viewBox="${x0} ${y0} ${w} ${h}">`;
  });

  // 满幅底色矩形的 100% 是按视口算的，x=0 落在新 viewBox 之外，
  // 必须改写成新画布的实际矩形，否则背景会消失（导出 PNG 时变成透明）
  return trimmed.replace(
    /<rect\b([^>]*width="100%"[^>]*height="100%"[^>]*)\/>/g,
    (whole, raw) => {
      const a = parseAttrs(raw);
      let out = '<rect';
      for (const k of Object.keys(a)) {
        if (/^(x|y|width|height)$/i.test(k)) continue;
        out += ` ${k}="${a[k]}"`;
      }
      return `${out} x="${x0}" y="${y0}" width="${w}" height="${h}"/>`;
    }
  );
}

/* ------------------------------------------------------------------ *
 * SVG 文字转曲
 * ------------------------------------------------------------------ */

function outlineSvgText(svg) {
  const styles = parseStyleBlock(svg);
  let replaced = 0;
  const out = svg.replace(/<text\b([^>]*)>([\s\S]*?)<\/text>/g, (whole, attrRaw, inner) => {
    const attrs = parseAttrs(attrRaw || '');
    const props = resolveProps(attrs, styles);
    const text = decodeXml(inner || '').replace(/\s+/g, ' ').trim();
    if (!text) return '';
    const hit = getFont(props['font-weight']);
    if (!hit) return whole; // 没有可用字体时保留原文字，至少不会丢内容
    const cmds = textToCommands(
      text,
      parseFloat(attrs.x) || 0,
      parseFloat(attrs.y) || 0,
      parseFloat(props['font-size']) || 16,
      props['text-anchor'],
      hit.font,
      props['dominant-baseline']
    );
    if (!cmds.length) return '';
    const fill = props.fill && props.fill !== 'none' ? props.fill : '#000000';
    replaced++;
    return `<path d="${commandsToSvgPath(cmds)}" fill="${fill}"/>`;
  });
  return { svg: out, replaced };
}

/* ------------------------------------------------------------------ *
 * PDF 输出
 * ------------------------------------------------------------------ */

function hexToRgb(color) {
  if (!color) return null;
  const c = String(color).trim();
  if (c === 'none' || c === 'transparent') return null;
  if (c[0] === '#') {
    let h = c.slice(1);
    if (h.length === 3) h = h.split('').map((x) => x + x).join('');
    if (h.length >= 6) {
      return {
        r: parseInt(h.slice(0, 2), 16) / 255,
        g: parseInt(h.slice(2, 4), 16) / 255,
        b: parseInt(h.slice(4, 6), 16) / 255
      };
    }
  }
  const m = c.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
  if (m) return { r: m[1] / 255, g: m[2] / 255, b: m[3] / 255 };
  const named = {
    white: '#ffffff', black: '#000000', red: '#ff0000', blue: '#0000ff',
    green: '#008000', gray: '#808080', grey: '#808080'
  };
  if (named[c.toLowerCase()]) return hexToRgb(named[c.toLowerCase()]);
  return null;
}

function rectPathOps(x, y, w, h, rx) {
  if (!rx || rx <= 0) {
    return [`${num(x)} ${num(y)} ${num(w)} ${num(h)} re`];
  }
  const r = Math.min(rx, w / 2, h / 2);
  const k = 0.5522847498307936 * r;
  return [
    `${num(x + r)} ${num(y)} m`,
    `${num(x + w - r)} ${num(y)} l`,
    `${num(x + w - r + k)} ${num(y)} ${num(x + w)} ${num(y + r - k)} ${num(x + w)} ${num(y + r)} c`,
    `${num(x + w)} ${num(y + h - r)} l`,
    `${num(x + w)} ${num(y + h - r + k)} ${num(x + w - r + k)} ${num(y + h)} ${num(x + w - r)} ${num(y + h)} c`,
    `${num(x + r)} ${num(y + h)} l`,
    `${num(x + r - k)} ${num(y + h)} ${num(x)} ${num(y + h - r + k)} ${num(x)} ${num(y + h - r)} c`,
    `${num(x)} ${num(y + r)} l`,
    `${num(x)} ${num(y + r - k)} ${num(x + r - k)} ${num(y)} ${num(x + r)} ${num(y)} c`,
    'h'
  ];
}

function circlePathOps(cx, cy, r) {
  const k = 0.5522847498307936 * r;
  return [
    `${num(cx + r)} ${num(cy)} m`,
    `${num(cx + r)} ${num(cy + k)} ${num(cx + k)} ${num(cy + r)} ${num(cx)} ${num(cy + r)} c`,
    `${num(cx - k)} ${num(cy + r)} ${num(cx - r)} ${num(cy + k)} ${num(cx - r)} ${num(cy)} c`,
    `${num(cx - r)} ${num(cy - k)} ${num(cx - k)} ${num(cy - r)} ${num(cx)} ${num(cy - r)} c`,
    `${num(cx + k)} ${num(cy - r)} ${num(cx + r)} ${num(cy - k)} ${num(cx + r)} ${num(cy)} c`,
    'h'
  ];
}

function cmdsToPdfOps(cmds) {
  const ops = [];
  let cur = { x: 0, y: 0 };
  for (const c of cmds) {
    if (c.op === 'M') {
      cur = c.pts[0];
      ops.push(`${num(cur.x)} ${num(cur.y)} m`);
    } else if (c.op === 'L') {
      cur = c.pts[0];
      ops.push(`${num(cur.x)} ${num(cur.y)} l`);
    } else if (c.op === 'Q') {
      const [q, p] = c.pts;
      const c1 = { x: cur.x + (2 / 3) * (q.x - cur.x), y: cur.y + (2 / 3) * (q.y - cur.y) };
      const c2 = { x: p.x + (2 / 3) * (q.x - p.x), y: p.y + (2 / 3) * (q.y - p.y) };
      ops.push(`${num(c1.x)} ${num(c1.y)} ${num(c2.x)} ${num(c2.y)} ${num(p.x)} ${num(p.y)} c`);
      cur = p;
    } else if (c.op === 'C') {
      const [c1, c2, p] = c.pts;
      ops.push(`${num(c1.x)} ${num(c1.y)} ${num(c2.x)} ${num(c2.y)} ${num(p.x)} ${num(p.y)} c`);
      cur = p;
    } else if (c.op === 'Z') {
      ops.push('h');
    }
  }
  return ops;
}

/** 把图元列表写进 PDF 内容流 */
function buildContentStream(parsed) {
  const { width, height, items } = parsed;
  const ops = [];
  ops.push('q');
  // SVG 像素坐标 -> PDF：缩放 0.75(96dpi->72pt)，y 轴翻转
  ops.push(`0.75 0 0 -0.75 0 ${num(height * 0.75)} cm`);

  for (const it of items) {
    let pathOps = null;
    if (it.type === 'rect') {
      if (it.w <= 0 || it.h <= 0) continue;
      pathOps = rectPathOps(it.x, it.y, it.w, it.h, it.rx);
    } else if (it.type === 'circle') {
      if (it.r <= 0) continue;
      pathOps = circlePathOps(it.cx, it.cy, it.r);
    } else if (it.type === 'path') {
      pathOps = cmdsToPdfOps(it.cmds);
    }
    if (!pathOps || !pathOps.length) continue;

    const fill = hexToRgb(it.fill);
    const stroke = hexToRgb(it.stroke);
    if (!fill && !stroke) continue;

    if (fill) ops.push(`${num(fill.r)} ${num(fill.g)} ${num(fill.b)} rg`);
    if (stroke) {
      ops.push(`${num(stroke.r)} ${num(stroke.g)} ${num(stroke.b)} RG`);
      ops.push(`${num(it.strokeWidth || 1)} w`);
    }
    ops.push(...pathOps);
    if (fill && stroke) ops.push('B');
    else if (fill) ops.push('f');
    else ops.push('S');
  }

  ops.push('Q');
  return Buffer.from(ops.join('\n'), 'latin1');
}

/** 生成单页 PDF */
function buildPdf(parsed) {
  const pageW = parsed.width * 0.75;
  const pageH = parsed.height * 0.75;
  const content = zlib.deflateSync(buildContentStream(parsed), { level: 9 });

  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(pageW)} ${num(pageH)}] /Resources << /ProcSet [/PDF] >> /Contents 4 0 R >>`
  ];

  const chunks = [];
  let offset = 0;
  const push = (buf) => { chunks.push(buf); offset += buf.length; };
  const offsets = [];

  push(Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1'));

  objects.forEach((body, i) => {
    offsets.push(offset);
    push(Buffer.from(`${i + 1} 0 obj\n${body}\nendobj\n`, 'latin1'));
  });

  // 内容流（单独拼，便于写入二进制）
  offsets.push(offset);
  push(Buffer.from(`4 0 obj\n<< /Length ${content.length} /Filter /FlateDecode >>\nstream\n`, 'latin1'));
  push(content);
  push(Buffer.from('\nendstream\nendobj\n', 'latin1'));

  const xrefStart = offset;
  let xref = `xref\n0 ${objects.length + 2}\n0000000000 65535 f \n`;
  for (const off of offsets) xref += String(off).padStart(10, '0') + ' 00000 n \n';
  push(Buffer.from(xref, 'latin1'));
  push(Buffer.from(
    `trailer\n<< /Size ${objects.length + 2} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`,
    'latin1'
  ));

  return Buffer.concat(chunks);
}

function svgToPdf(svg) {
  const parsed = parseItems(svg);
  return { pdf: buildPdf(parsed), width: parsed.width, height: parsed.height };
}

/* ------------------------------------------------------------------ *
 * PNG 输出
 * ------------------------------------------------------------------ */

const MAX_PIXELS = 60 * 1000 * 1000;

async function svgToPng(svg, scale) {
  if (!sharp) throw new Error('未安装 sharp，无法导出 PNG');
  const { width, height } = svgSize(svg);
  const s = Math.max(0.05, Math.min(4, Number(scale) || 1));
  const outW = Math.round(width * s);
  const outH = Math.round(height * s);
  if (outW * outH > MAX_PIXELS) {
    const allow = Math.sqrt(MAX_PIXELS / (width * height));
    throw new Error(
      `PNG 尺寸过大（${outW}×${outH}）。请把缩放比例调小到 ${allow.toFixed(2)} 以下，或改用单会场导出。`
    );
  }
  // sharp/librsvg 默认按 72dpi 渲染，即 1 个 SVG 用户单位 = 1 像素；
  // 因此 density = 72 * scale 才能得到「scale 倍原尺寸」的位图
  const density = 72 * s;
  const buf = await sharp(Buffer.from(svg), { density })
    .png({ compressionLevel: 9 })
    .toBuffer();
  return buf;
}

module.exports = {
  fontInfo,
  outlineSvgText,
  trimSvg,
  svgToPdf,
  svgToPng,
  svgSize,
  parseItems,
  hasPdfSupport: () => !!fontkit,
  hasPngSupport: () => !!sharp
};
