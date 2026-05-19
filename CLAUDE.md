# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

会议座位管理系统 — 支持座位编排、参会者管理、区域规划和座位牌生成。

## Architecture

### Dual-Stack Layout

```
Node.js (Express, port 3000)      Python (Streamlit, port 8505)
├── server.js (100KB+)            └── app.py (90KB+)
│   ├── REST API (60+ endpoints)
│   ├── Static frontend serving
│   └── Excel parsing engine
├── parse-excel.js (86KB)
└── public/
    ├── index.html        # 座位查询首页
    ├── admin.html        # 管理后台 (216KB, 单体 SPA)
    └── login.html        # 登录页
```

### Production Stack (Docker)

```
Nginx (port 80/8506)
 ├── /api/*  ──proxy──> Node (port 3000)
 ├── /*      ──proxy──> Node (port 3000)
 └── /       ──proxy──> Streamlit (port 8505)
Supervisor manages all three processes.
```

### Data Flow

```
Excel Upload
  ├── venue layout (.xlsx) ──> parseWorkbook() in parse-excel.js
  │     ├── detectLayoutType() → standard / theater / u-shape
  │     └── returns structured venue data with seats, rows, floors
  └── attendee list (.xlsx) ──> /api/attendees/import
        └── parsed into data.json alongside venues

Seat Assignment
  ├── Drag & drop:  PUT /api/attendees/:id/seat
  ├── Region-based: POST /api/venues/:id/assign-seat (batch)
  ├── Random:       POST /api/venues/:id/random-seat
  └── SVG Export:   GET /api/export-seating-svg → inline SVG

Seat Card (Python side)
  Streamlit app.py → reportlab PDF (A4/A5)
    └── Token-based auth (from admin.html, hmac SHA-256)
```

### Data Storage

- **data/data.json** — All venues, attendees, seating assignments (single JSON file)
- **data/config.json** — Admin/organizer passwords (bcrypt hashed)
- **data/backups/** — Auto-rotated backups (last 10)
- **audit.log** — All data-modifying operations
- **access.log** — Structured request logs

## Key Modules

### server.js
Express app with 60+ API endpoints. Major API groups:
- `/api/venues` — CRUD + layout analysis + preview generation
- `/api/attendees` — CRUD + batch import + seat assignment
- `/api/search` — QR code + seat search by name
- `/api/export-*` — SVG seating chart + Excel export
- `/api/backups`, `/api/logs`, `/api/config` — Admin operations
- Auth via `requireAdmin` / `requireAuth` middleware (token-based)

### parse-excel.js
Three layout detection algorithms:
- **standard** — Grid-based, column gaps > 1 → aisles, seat numbers in header rows
- **theater** — Multi-floor, floor labels (一楼/二楼/1F/2F), floor rows with seat numbers
- **u-shape** — Requires `第X列` labels, multi-column with bottom-row seats
Parsing pipeline: `parseWorkbook()` → `detectSheetMode()` → layout-specific detection

### app.py (Streamlit)
Single-page PDF card generator:
- Sidebar: font, color, size, spacing, background image, text effects
- Main: attendee selection + preview
- Data sourced from Node API (token-auth) or direct Excel upload
- Generates PDF via reportlab with CJK font support

## Commands

```bash
# Install dependencies
npm install
pip install -r requirements.txt

# Development (two terminals)
node server.js                          # API + frontend on port 3000
streamlit run app.py --server.port 8505 # Card generator on port 8505

# Docker
docker compose up -d --build

# Build deployment package
node build-package.js                   # Creates meeting-system-v*.zip
```

## Important Details

- **Passwords**: Stored bcrypt-hashed in `data/config.json`. Default: admin/admin888, organizer/organizer888
- **Excel header detection**: Row 1 = short text (column names) → treated as header; otherwise row 2
- **Seat number detection**: Pure numbers, `数字号`, `字母+数字` (but not R1/F1/Floor1)
- **Row label patterns**: `第X排`, `第X列`, `Row N`, `RN`, `一楼第X排`
- **Sofa rows**: Cells containing `沙发`, `sofa`, `Sofa` are rendered as sofa markers in SVG
- **Auto-backup**: Every data write triggers `createBackup()`, rotating at 10 max
- **admin.html**: 216KB single-page HTML (no JS framework), uses `<script>` includes for modularity
- **file upload限**: Nginx 50MB, Express JSON body limit 50MB
- **SVG export**: Inline SVG rendered via `res.send()`, designed for large-format printing
- **Card generator auth**: HMAC SHA-256 token (60-minute expiry), shared secret across Node/Streamlit
