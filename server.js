/**
 * 재고 매칭 v4 - 로케이션별 수량 관리
 * Node.js + Express + better-sqlite3
 */
const express = require('express');
const Database = require('better-sqlite3');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';
const DB_PATH = path.join(__dirname, 'popup.db');
const XLSX_PATH = path.join(__dirname, 'popup_location.xlsx');

const UNASSIGNED = '미지정';

// ============================================================================
// DB 초기화 + 마이그레이션
// ============================================================================
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS products (
    barcode      TEXT PRIMARY KEY,
    style        TEXT,
    style_name   TEXT,
    color        TEXT,
    color_name   TEXT,
    size         TEXT,
    qty          INTEGER,
    store_name   TEXT,
    raw_json     TEXT
  );

  CREATE TABLE IF NOT EXISTS location_assignments (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    barcode      TEXT NOT NULL,
    location     TEXT NOT NULL,
    assigned_at  TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    is_current   INTEGER NOT NULL DEFAULT 1
  );
  CREATE INDEX IF NOT EXISTS idx_assign_barcode    ON location_assignments(barcode);
  CREATE INDEX IF NOT EXISTS idx_assign_location   ON location_assignments(location);
  CREATE INDEX IF NOT EXISTS idx_assign_is_current ON location_assignments(is_current);

  CREATE TABLE IF NOT EXISTS stock_by_location (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    barcode     TEXT NOT NULL,
    location    TEXT NOT NULL,
    qty         INTEGER NOT NULL DEFAULT 0,
    updated_at  TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    UNIQUE(barcode, location)
  );
  CREATE INDEX IF NOT EXISTS idx_sbl_barcode  ON stock_by_location(barcode);
  CREATE INDEX IF NOT EXISTS idx_sbl_location ON stock_by_location(location);
`);

// transactions: 기존이 있으면 type CHECK가 다를 수 있으므로 안전하게 마이그레이션
function ensureTransactionsTable() {
  const exists = db.prepare(
    "SELECT name FROM sqlite_master WHERE type='table' AND name='transactions'"
  ).get();
  if (!exists) {
    db.exec(`
      CREATE TABLE transactions (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        barcode       TEXT NOT NULL,
        type          TEXT NOT NULL CHECK (type IN ('SALE','RESTOCK','CANCEL','MOVE')),
        qty           INTEGER NOT NULL,
        before_qty    INTEGER NOT NULL,
        after_qty     INTEGER NOT NULL,
        related_tx_id INTEGER,
        memo          TEXT,
        location      TEXT,
        created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime')),
        cancelled_at  TEXT
      );
      CREATE INDEX idx_tx_barcode    ON transactions(barcode);
      CREATE INDEX idx_tx_type       ON transactions(type);
      CREATE INDEX idx_tx_created_at ON transactions(created_at);
      CREATE INDEX idx_tx_location   ON transactions(location);
    `);
    console.log('[MIGRATE] transactions 테이블 생성');
    return;
  }
  // location 컬럼이 없으면 추가
  const cols = db.prepare("PRAGMA table_info(transactions)").all();
  if (!cols.some(c => c.name === 'location')) {
    db.exec(`ALTER TABLE transactions ADD COLUMN location TEXT`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_tx_location ON transactions(location)`);
    console.log('[MIGRATE] transactions.location 컬럼 추가');
  }
  // CHECK 제약에 MOVE를 추가하려면 테이블을 재생성해야 한다 (SQLite 제약)
  // 기존 CHECK가 MOVE를 허용하지 않으면 재구축
  const sql = db.prepare(
    "SELECT sql FROM sqlite_master WHERE type='table' AND name='transactions'"
  ).get().sql || '';
  if (!sql.includes("'MOVE'")) {
    console.log('[MIGRATE] transactions 테이블 재구축 (MOVE type 허용)');
    db.exec('BEGIN');
    try {
      db.exec(`
        CREATE TABLE transactions_new (
          id            INTEGER PRIMARY KEY AUTOINCREMENT,
          barcode       TEXT NOT NULL,
          type          TEXT NOT NULL CHECK (type IN ('SALE','RESTOCK','CANCEL','MOVE')),
          qty           INTEGER NOT NULL,
          before_qty    INTEGER NOT NULL,
          after_qty     INTEGER NOT NULL,
          related_tx_id INTEGER,
          memo          TEXT,
          location      TEXT,
          created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime')),
          cancelled_at  TEXT
        );
        INSERT INTO transactions_new (id, barcode, type, qty, before_qty, after_qty,
          related_tx_id, memo, location, created_at, cancelled_at)
        SELECT id, barcode, type, qty, before_qty, after_qty,
               related_tx_id, memo, location, created_at, cancelled_at
        FROM transactions;
        DROP TABLE transactions;
        ALTER TABLE transactions_new RENAME TO transactions;
        CREATE INDEX idx_tx_barcode    ON transactions(barcode);
        CREATE INDEX idx_tx_type       ON transactions(type);
        CREATE INDEX idx_tx_created_at ON transactions(created_at);
        CREATE INDEX idx_tx_location   ON transactions(location);
      `);
      db.exec('COMMIT');
    } catch (e) {
      db.exec('ROLLBACK');
      throw e;
    }
  }
}
ensureTransactionsTable();

// products 마이그레이션
function columnExists(table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some(r => r.name === col);
}
function safeAddColumn(table, col, ddl) {
  if (!columnExists(table, col)) {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${ddl}`);
    console.log(`[MIGRATE] ${table}.${col} 추가`);
    return true;
  }
  return false;
}
const a1 = safeAddColumn('products', 'initial_qty', 'initial_qty INTEGER');
const a2 = safeAddColumn('products', 'current_qty', 'current_qty INTEGER');
if (a1 || a2) {
  const r = db.prepare(`
    UPDATE products SET initial_qty = COALESCE(initial_qty, qty),
                        current_qty = COALESCE(current_qty, qty)
  `).run();
  console.log(`[MIGRATE] products 재고 ${r.changes}건 채움`);
}

// ============================================================================
// 엑셀 마스터 적재 (이미 적재된 경우 스킵)
// ============================================================================
function loadMasterFromExcel() {
  const existing = db.prepare('SELECT COUNT(*) AS c FROM products').get().c;
  if (existing > 0) {
    console.log(`[MASTER] 이미 ${existing}건 적재됨. 스킵.`);
    return;
  }
  if (!fs.existsSync(XLSX_PATH)) {
    console.warn(`[MASTER] 엑셀 없음: ${XLSX_PATH}`);
    return;
  }
  console.log('[MASTER] 엑셀 적재 시작...');
  const wb = XLSX.readFile(XLSX_PATH);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  const insert = db.prepare(`
    INSERT OR REPLACE INTO products
      (barcode, style, style_name, color, color_name, size, qty, store_name, raw_json,
       initial_qty, current_qty)
    VALUES (@barcode, @style, @style_name, @color, @color_name, @size, @qty, @store_name, @raw_json,
            @qty, @qty)
  `);
  const items = rows.filter(r => r['바코드']).map(r => ({
    barcode:    String(r['바코드']).trim(),
    style:      r['스타일']   != null ? String(r['스타일']).trim()   : null,
    style_name: r['스타일명'] != null ? String(r['스타일명']).trim() : null,
    color:      r['컬러']     != null ? String(r['컬러']).trim()     : null,
    color_name: r['컬러명']   != null ? String(r['컬러명']).trim()   : null,
    size:       r['사이즈']   != null ? String(r['사이즈']).trim()   : null,
    qty:        r['수량합계'] != null ? (Number(r['수량합계']) || 0) : 0,
    store_name: r['매장명']   != null ? String(r['매장명']).trim()   : null,
    raw_json:   JSON.stringify(r),
  }));
  const insertMany = db.transaction((arr) => { for (const x of arr) insert.run(x); });
  insertMany(items);

  // 엑셀의 기존 로케이션도 보존
  let preset = 0;
  const insLoc = db.prepare(
    `INSERT INTO location_assignments (barcode, location, is_current) VALUES (?, ?, 1)`
  );
  const presetTx = db.transaction(() => {
    for (const r of rows) {
      if (r['바코드'] && r['로케이션'] && String(r['로케이션']).trim()) {
        insLoc.run(String(r['바코드']).trim(), String(r['로케이션']).trim());
        preset++;
      }
    }
  });
  presetTx();
  console.log(`[MASTER] ${items.length}건 적재 (로케이션 ${preset}건 보존)`);
}
loadMasterFromExcel();

// ============================================================================
// stock_by_location 초기 마이그레이션 (1회만)
// ============================================================================
function migrateStockByLocation() {
  const existing = db.prepare('SELECT COUNT(*) AS c FROM stock_by_location').get().c;
  if (existing > 0) {
    console.log(`[SBL-MIGRATE] 이미 ${existing}건 존재. 스킵.`);
    return;
  }
  // 모든 products에 대해:
  // - is_current=1 위치가 있으면 그 위치에 current_qty 전체를 배치 (기존 로케이션 등록 보존)
  // - 없으면 '미지정'에 current_qty 배치
  const products = db.prepare('SELECT barcode, current_qty FROM products').all();
  const getCurLoc = db.prepare(`
    SELECT location FROM location_assignments
    WHERE barcode = ? AND is_current = 1
    ORDER BY id DESC LIMIT 1
  `);
  const ins = db.prepare(`
    INSERT INTO stock_by_location (barcode, location, qty)
    VALUES (?, ?, ?)
  `);
  let assigned = 0, unassigned = 0;
  const tx = db.transaction(() => {
    for (const p of products) {
      const curLoc = getCurLoc.get(p.barcode);
      if (curLoc && curLoc.location) {
        ins.run(p.barcode, curLoc.location, p.current_qty || 0);
        assigned++;
      } else {
        ins.run(p.barcode, UNASSIGNED, p.current_qty || 0);
        unassigned++;
      }
    }
  });
  tx();
  console.log(`[SBL-MIGRATE] 완료: 등록 위치 ${assigned}건, 미지정 ${unassigned}건`);

  // 검증: SUM(sbl.qty) by barcode == products.current_qty
  const mismatches = db.prepare(`
    SELECT p.barcode, p.current_qty, COALESCE(s.s,0) AS sbl_sum
    FROM products p
    LEFT JOIN (SELECT barcode, SUM(qty) s FROM stock_by_location GROUP BY barcode) s
      ON s.barcode = p.barcode
    WHERE p.current_qty != COALESCE(s.s, 0)
  `).all();
  if (mismatches.length === 0) {
    console.log('[SBL-MIGRATE] 검증 통과: 모든 바코드의 합계 일치');
  } else {
    console.warn(`[SBL-MIGRATE] 검증 실패: ${mismatches.length}건 불일치`);
  }
}
migrateStockByLocation();

// ============================================================================
// Express
// ============================================================================
const app = express();
app.use(express.json({ limit: '2mb' }));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.use(express.static(path.join(__dirname, 'public')));

// 공통 prepared
const stmtProduct = db.prepare('SELECT * FROM products WHERE barcode = ?');
const stmtStockByLoc = db.prepare(`
  SELECT location, qty, updated_at FROM stock_by_location
  WHERE barcode = ? ORDER BY qty DESC, location ASC
`);
const stmtLocHistory = db.prepare(`
  SELECT id, location, assigned_at, is_current FROM location_assignments
  WHERE barcode = ? ORDER BY id DESC
`);
const stmtRecentTx = db.prepare(`
  SELECT * FROM transactions WHERE barcode = ?
  ORDER BY id DESC LIMIT ?
`);

// 헬퍼: stock_by_location UPSERT (delta 가감)
function applyStockDelta(barcode, location, delta) {
  const loc = location && String(location).trim() ? String(location).trim() : UNASSIGNED;
  const exist = db.prepare(
    'SELECT id, qty FROM stock_by_location WHERE barcode = ? AND location = ?'
  ).get(barcode, loc);
  if (exist) {
    const newQty = exist.qty + delta;
    db.prepare(
      `UPDATE stock_by_location SET qty = ?, updated_at = datetime('now','localtime') WHERE id = ?`
    ).run(newQty, exist.id);
    return { location: loc, before: exist.qty, after: newQty };
  } else {
    db.prepare(
      `INSERT INTO stock_by_location (barcode, location, qty, updated_at)
       VALUES (?, ?, ?, datetime('now','localtime'))`
    ).run(barcode, loc, delta);
    return { location: loc, before: 0, after: delta };
  }
}

// ============================================================================
// 제품 조회
// ============================================================================
app.get('/api/product/:barcode', (req, res) => {
  const bc = String(req.params.barcode).trim();
  const product = stmtProduct.get(bc);
  const stockByLocation = stmtStockByLoc.all(bc);
  const history = stmtLocHistory.all(bc);
  const recentTx = product ? stmtRecentTx.all(bc, 10) : [];
  const currentLoc = history.find(h => h.is_current) || null;
  const totalAcrossLocations = stockByLocation.reduce((s, x) => s + (x.qty || 0), 0);

  if (!product) {
    return res.status(404).json({
      ok: false, barcode: bc, message: '마스터에 없는 바코드',
      current: currentLoc, history, recentTx: [],
      stockByLocation, totalAcrossLocations,
      currentQty: null, mismatch: false,
    });
  }
  res.json({
    ok: true, barcode: bc, product,
    initial_qty: product.initial_qty,
    current_qty: product.current_qty,
    currentQty: product.current_qty,
    sold_qty: (product.initial_qty || 0) - (product.current_qty || 0),
    stockByLocation,
    totalAcrossLocations,
    mismatch: totalAcrossLocations !== product.current_qty,
    current: currentLoc ? { location: currentLoc.location, assigned_at: currentLoc.assigned_at } : null,
    history, recentTx,
  });
});

// 그 바코드의 로케이션별 수량
app.get('/api/stock-locations/:barcode', (req, res) => {
  const bc = String(req.params.barcode).trim();
  const rows = stmtStockByLoc.all(bc);
  res.json({ ok: true, barcode: bc, rows });
});

// 그 위치에 있는 모든 (바코드, 수량)
app.get('/api/location/:location/stocks', (req, res) => {
  const loc = String(req.params.location).trim();
  const rows = db.prepare(`
    SELECT s.barcode, s.location, s.qty, s.updated_at,
           p.style, p.style_name, p.color, p.color_name, p.size,
           p.initial_qty, p.current_qty
    FROM stock_by_location s
    LEFT JOIN products p ON p.barcode = s.barcode
    WHERE s.location = ?
    ORDER BY s.qty DESC, s.barcode ASC
  `).all(loc);
  res.json({ ok: true, location: loc, rows });
});

// ============================================================================
// 로케이션 등록 (변경됨)
// ============================================================================
// POST /api/assign
//   신형식: { location, items: [{barcode, qty}, ...] }
//   구형식(하위호환): { location, barcodes: [bc1, bc2, ...] } → 각 qty=1
app.post('/api/assign', (req, res) => {
  const { location } = req.body || {};
  let items = req.body && req.body.items;
  const legacyBarcodes = req.body && req.body.barcodes;
  if (!location || typeof location !== 'string' || !location.trim()) {
    return res.status(400).json({ ok: false, message: 'location 누락' });
  }
  if (!Array.isArray(items) || items.length === 0) {
    if (Array.isArray(legacyBarcodes) && legacyBarcodes.length) {
      // 구형식 → 같은 바코드는 카운트 합산 (qty=count)
      const m = new Map();
      for (const b of legacyBarcodes) {
        const k = String(b || '').trim();
        if (!k) continue;
        m.set(k, (m.get(k) || 0) + 1);
      }
      items = Array.from(m.entries()).map(([barcode, qty]) => ({ barcode, qty }));
    } else {
      return res.status(400).json({ ok: false, message: 'items 또는 barcodes 누락' });
    }
  }
  const loc = location.trim();
  const checkProd = db.prepare('SELECT 1 FROM products WHERE barcode = ?');
  const insLocAssign = db.prepare(
    `INSERT INTO location_assignments (barcode, location, is_current) VALUES (?, ?, 1)`
  );
  const clearOldAssign = db.prepare(
    `UPDATE location_assignments SET is_current = 0 WHERE barcode = ? AND is_current = 1`
  );

  const results = [];
  const notFound = [];
  const tx = db.transaction(() => {
    for (const it of items) {
      const bc = String(it.barcode || '').trim();
      const qty = Math.max(1, parseInt(it.qty, 10) || 1);
      if (!bc) continue;
      if (!checkProd.get(bc)) notFound.push(bc);
      // 정책: 다른 위치는 건드리지 않고 해당 위치에만 += qty
      const r = applyStockDelta(bc, loc, qty);
      // location_assignments 이력도 갱신 (가장 최근 위치)
      clearOldAssign.run(bc);
      insLocAssign.run(bc, loc);
      results.push({ barcode: bc, qty, locationQty: r.after });
    }
  });
  tx();
  res.json({ ok: true, success: true, saved: results.length, results, notFound, location: loc });
});

// ============================================================================
// 검색 (상품명, 바코드, 로케이션)
// ============================================================================
app.get('/api/search/product', (req, res) => {
  const qRaw = req.query.q;
  const limitRaw = Number(req.query.limit);
  if (!qRaw || !String(qRaw).trim()) {
    return res.status(400).json({ ok: false, message: '검색어(q)가 비어 있습니다' });
  }
  const q = String(qRaw).trim();
  const limit = Math.min(Math.max(Number.isFinite(limitRaw) ? limitRaw : 1000, 1), 5000);
  const like = '%' + q.replace(/[%_]/g, c => '\\' + c) + '%';
  const prefix = q.replace(/[%_]/g, c => '\\' + c) + '%';

  const rows = db.prepare(`
    SELECT p.barcode, p.style, p.style_name, p.color, p.color_name, p.size,
           p.initial_qty, p.current_qty,
           (p.initial_qty - p.current_qty) AS sold_qty,
           (SELECT COUNT(*) FROM stock_by_location s
              WHERE s.barcode = p.barcode AND s.qty > 0) AS location_count,
           (SELECT location FROM stock_by_location s
              WHERE s.barcode = p.barcode AND s.qty > 0
              ORDER BY s.qty DESC LIMIT 1) AS top_location,
           (SELECT qty FROM stock_by_location s
              WHERE s.barcode = p.barcode AND s.qty > 0
              ORDER BY s.qty DESC LIMIT 1) AS top_location_qty,
           CASE
             WHEN LOWER(COALESCE(p.style_name,'')) LIKE LOWER(?) ESCAPE '\\' THEN 0
             WHEN LOWER(COALESCE(p.style_name,'')) LIKE LOWER(?) ESCAPE '\\' THEN 1
             ELSE 2
           END AS rank_score
    FROM products p
    WHERE
      LOWER(COALESCE(p.style_name,'')) LIKE LOWER(?) ESCAPE '\\'
      OR LOWER(COALESCE(p.style,''))     LIKE LOWER(?) ESCAPE '\\'
      OR LOWER(COALESCE(p.color_name,'')) LIKE LOWER(?) ESCAPE '\\'
      OR LOWER(COALESCE(p.barcode,''))    LIKE LOWER(?) ESCAPE '\\'
    ORDER BY rank_score ASC, p.style_name ASC, p.barcode ASC
    LIMIT ?
  `).all(prefix, like, like, like, like, like, limit + 1);
  const hasMore = rows.length > limit;
  const sliced = rows.slice(0, limit);

  // 각 row에 stockByLocation 배열 부착 (재고조회 카드 렌더링용)
  const stmtLocs = db.prepare(`
    SELECT location, qty, updated_at FROM stock_by_location
    WHERE barcode = ? AND qty > 0
    ORDER BY qty DESC, location ASC
  `);
  for (const r of sliced) {
    r.stockByLocation = stmtLocs.all(r.barcode);
  }

  res.json({ ok: true, q, hasMore, limit, rows: sliced });
});

app.get('/api/search/barcode/:q', (req, res) => {
  const q = '%' + String(req.params.q).trim() + '%';
  const rows = db.prepare(`
    SELECT p.barcode, p.style_name, p.color_name, p.size,
           p.initial_qty, p.current_qty,
           (SELECT location FROM stock_by_location s
              WHERE s.barcode = p.barcode AND s.qty > 0
              ORDER BY s.qty DESC LIMIT 1) AS current_location
    FROM products p
    WHERE p.barcode LIKE ?
    ORDER BY p.barcode
    LIMIT 100
  `).all(q);
  res.json({ ok: true, rows });
});

// 로케이션 부분일치 검색 (stock_by_location 기반)
app.get('/api/search/location/:q', (req, res) => {
  const q = '%' + String(req.params.q).trim() + '%';
  const rows = db.prepare(`
    SELECT s.barcode, s.location, s.qty AS location_qty, s.updated_at,
           p.style, p.style_name, p.color, p.color_name, p.size,
           p.initial_qty, p.current_qty
    FROM stock_by_location s
    LEFT JOIN products p ON p.barcode = s.barcode
    WHERE s.location LIKE ? AND s.qty > 0
    ORDER BY s.location, s.qty DESC
    LIMIT 500
  `).all(q);
  res.json({ ok: true, rows });
});

// 로케이션 목록 (자동완성용) - stock_by_location 기반
app.get('/api/locations', (req, res) => {
  const rows = db.prepare(`
    SELECT location, SUM(qty) AS total_qty, COUNT(DISTINCT barcode) AS bc_count,
           MAX(updated_at) AS last_at
    FROM stock_by_location
    WHERE qty > 0
    GROUP BY location
    ORDER BY last_at DESC
  `).all();
  res.json({ ok: true, rows });
});

app.get('/api/history/:barcode', (req, res) => {
  res.json({ ok: true, barcode: req.params.barcode, rows: stmtLocHistory.all(req.params.barcode) });
});

// ============================================================================
// 판매 / 입고 / 취소 / 이동
// ============================================================================
function ensureLocationProvided(req, defaultLoc) {
  const { location } = req.body || {};
  if (location && String(location).trim()) return String(location).trim();
  if (defaultLoc !== undefined) return defaultLoc;
  return null;
}

app.post('/api/sale', (req, res) => {
  const { barcode, qty, memo } = req.body || {};
  const bc = String(barcode || '').trim();
  const n = Number(qty);
  const loc = ensureLocationProvided(req, null);
  if (!bc) return res.status(400).json({ ok: false, message: 'barcode 누락' });
  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    return res.status(400).json({ ok: false, message: 'qty는 1 이상의 정수여야 합니다' });
  }
  if (!loc) return res.status(400).json({ ok: false, message: '판매할 위치(location)를 선택하세요' });
  const product = stmtProduct.get(bc);
  if (!product) return res.status(404).json({ ok: false, message: `마스터에 없는 바코드: ${bc}` });

  let result;
  const trx = db.transaction(() => {
    const before = product.current_qty;
    db.prepare('UPDATE products SET current_qty = current_qty - ? WHERE barcode = ?').run(n, bc);
    const sblRes = applyStockDelta(bc, loc, -n);
    const after = before - n;
    const r = db.prepare(`
      INSERT INTO transactions (barcode, type, qty, before_qty, after_qty, memo, location)
      VALUES (?, 'SALE', ?, ?, ?, ?, ?)
    `).run(bc, n, before, after, memo ? String(memo).trim() : null, loc);
    result = { txId: r.lastInsertRowid, before, after, locBefore: sblRes.before, locAfter: sblRes.after };
  });
  trx();
  const updated = stmtProduct.get(bc);
  res.json({
    ok: true, success: true, txId: result.txId,
    beforeQty: result.before, afterQty: result.after,
    locationBefore: result.locBefore, locationAfter: result.locAfter,
    location: loc,
    warning: result.after < 0, locationWarning: result.locAfter < 0,
    product: updated,
  });
});

app.post('/api/restock', (req, res) => {
  const { barcode, qty, memo } = req.body || {};
  const bc = String(barcode || '').trim();
  const n = Number(qty);
  const loc = ensureLocationProvided(req, UNASSIGNED);
  if (!bc) return res.status(400).json({ ok: false, message: 'barcode 누락' });
  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    return res.status(400).json({ ok: false, message: 'qty는 1 이상의 정수여야 합니다' });
  }
  const product = stmtProduct.get(bc);
  if (!product) {
    return res.status(404).json({ ok: false, message: `마스터에 없는 바코드: ${bc} (입고 불가)` });
  }
  let result;
  const trx = db.transaction(() => {
    const before = product.current_qty;
    db.prepare('UPDATE products SET current_qty = current_qty + ? WHERE barcode = ?').run(n, bc);
    const sblRes = applyStockDelta(bc, loc, n);
    const after = before + n;
    const r = db.prepare(`
      INSERT INTO transactions (barcode, type, qty, before_qty, after_qty, memo, location)
      VALUES (?, 'RESTOCK', ?, ?, ?, ?, ?)
    `).run(bc, n, before, after, memo ? String(memo).trim() : null, loc);
    result = { txId: r.lastInsertRowid, before, after, locBefore: sblRes.before, locAfter: sblRes.after };
  });
  trx();
  const updated = stmtProduct.get(bc);
  res.json({
    ok: true, success: true, txId: result.txId,
    beforeQty: result.before, afterQty: result.after,
    locationBefore: result.locBefore, locationAfter: result.locAfter,
    location: loc, product: updated,
  });
});

app.post('/api/cancel/:txId', (req, res) => {
  const txId = Number(req.params.txId);
  if (!Number.isInteger(txId) || txId <= 0) {
    return res.status(400).json({ ok: false, message: '잘못된 거래 id' });
  }
  const orig = db.prepare('SELECT * FROM transactions WHERE id = ?').get(txId);
  if (!orig) return res.status(404).json({ ok: false, message: '거래를 찾을 수 없습니다' });
  if (orig.type === 'CANCEL') {
    return res.status(400).json({ ok: false, message: '취소 거래는 다시 취소할 수 없습니다' });
  }
  if (orig.cancelled_at) {
    return res.status(400).json({ ok: false, message: '이미 취소된 거래입니다' });
  }

  let result;
  const trx = db.transaction(() => {
    const p = stmtProduct.get(orig.barcode);
    if (!p) throw new Error('product gone');
    const before = p.current_qty;
    let after = before;

    if (orig.type === 'SALE') {
      db.prepare('UPDATE products SET current_qty = current_qty + ? WHERE barcode = ?')
        .run(orig.qty, orig.barcode);
      after = before + orig.qty;
      if (orig.location) applyStockDelta(orig.barcode, orig.location, orig.qty);
    } else if (orig.type === 'RESTOCK') {
      db.prepare('UPDATE products SET current_qty = current_qty - ? WHERE barcode = ?')
        .run(orig.qty, orig.barcode);
      after = before - orig.qty;
      if (orig.location) applyStockDelta(orig.barcode, orig.location, -orig.qty);
    } else if (orig.type === 'MOVE') {
      // MOVE 취소: from으로 되돌리기. memo에 from/to 보존되어 있을 것
      // (location 필드는 'from → to' 형태로 기록되어 있음)
      const arrow = (orig.location || '').split('→').map(s => s.trim());
      if (arrow.length === 2) {
        const [fromLoc, toLoc] = arrow;
        applyStockDelta(orig.barcode, toLoc, -orig.qty);
        applyStockDelta(orig.barcode, fromLoc, orig.qty);
      }
      after = before; // 전체 재고 변화 없음
    }

    db.prepare(`UPDATE transactions SET cancelled_at = datetime('now','localtime') WHERE id = ?`)
      .run(orig.id);
    const r = db.prepare(`
      INSERT INTO transactions (barcode, type, qty, before_qty, after_qty, related_tx_id, memo, location)
      VALUES (?, 'CANCEL', ?, ?, ?, ?, ?, ?)
    `).run(orig.barcode, orig.qty, before, after, orig.id, `취소: ${orig.type} 거래 #${orig.id}`, orig.location);
    result = { cancelTxId: r.lastInsertRowid, before, after };
  });
  trx();

  const updated = stmtProduct.get(orig.barcode);
  res.json({
    ok: true, success: true,
    cancelTxId: result.cancelTxId,
    originalTxId: orig.id, originalType: orig.type,
    restoredQty: orig.qty,
    beforeQty: result.before, afterQty: result.after,
    product: updated,
  });
});

// 위치 이동
app.post('/api/move', (req, res) => {
  const { barcode, qty, fromLocation, toLocation, memo } = req.body || {};
  const bc = String(barcode || '').trim();
  const n = Number(qty);
  const fromLoc = String(fromLocation || '').trim();
  const toLoc   = String(toLocation || '').trim();
  if (!bc) return res.status(400).json({ ok: false, message: 'barcode 누락' });
  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    return res.status(400).json({ ok: false, message: 'qty는 1 이상의 정수여야 합니다' });
  }
  if (!fromLoc || !toLoc) return res.status(400).json({ ok: false, message: 'fromLocation, toLocation 필수' });
  if (fromLoc === toLoc) return res.status(400).json({ ok: false, message: '같은 위치로는 이동할 수 없습니다' });
  const product = stmtProduct.get(bc);
  if (!product) return res.status(404).json({ ok: false, message: `마스터에 없는 바코드: ${bc}` });

  let result;
  const trx = db.transaction(() => {
    const fromRes = applyStockDelta(bc, fromLoc, -n);
    const toRes   = applyStockDelta(bc, toLoc, n);
    const r = db.prepare(`
      INSERT INTO transactions (barcode, type, qty, before_qty, after_qty, memo, location)
      VALUES (?, 'MOVE', ?, ?, ?, ?, ?)
    `).run(bc, n, product.current_qty, product.current_qty,
           memo ? String(memo).trim() : null, `${fromLoc} → ${toLoc}`);
    result = { txId: r.lastInsertRowid, fromRes, toRes };
  });
  trx();
  res.json({
    ok: true, success: true, txId: result.txId,
    fromLocation: fromLoc, toLocation: toLoc, qty: n,
    fromBefore: result.fromRes.before, fromAfter: result.fromRes.after,
    toBefore: result.toRes.before, toAfter: result.toRes.after,
    fromWarning: result.fromRes.after < 0,
  });
});

// 거래 내역 조회
app.get('/api/transactions', (req, res) => {
  const { type, barcode, location, from, to } = req.query;
  const limit = Math.min(Number(req.query.limit) || 100, 1000);
  const conds = [];
  const params = [];
  if (type) {
    const t = String(type).toUpperCase();
    if (['SALE', 'RESTOCK', 'CANCEL', 'MOVE'].includes(t)) {
      conds.push('t.type = ?'); params.push(t);
    }
  }
  if (barcode)  { conds.push('t.barcode = ?');  params.push(String(barcode).trim()); }
  if (location) { conds.push('t.location LIKE ?'); params.push('%' + String(location).trim() + '%'); }
  if (from)     { conds.push('t.created_at >= ?'); params.push(String(from)); }
  if (to)       { conds.push('t.created_at <= ?'); params.push(String(to)); }
  const where = conds.length ? 'WHERE ' + conds.join(' AND ') : '';
  const sql = `
    SELECT t.*, p.style_name, p.color_name, p.size,
           CASE WHEN t.cancelled_at IS NULL THEN 0 ELSE 1 END AS is_cancelled
    FROM transactions t
    LEFT JOIN products p ON p.barcode = t.barcode
    ${where}
    ORDER BY t.id DESC
    LIMIT ?
  `;
  res.json({ ok: true, rows: db.prepare(sql).all(...params, limit) });
});

// ============================================================================
// 통계
// ============================================================================
app.get('/api/stats', (req, res) => {
  const totalProducts = db.prepare('SELECT COUNT(*) AS c FROM products').get().c;
  const totalInitialStock = db.prepare('SELECT COALESCE(SUM(initial_qty),0) AS c FROM products').get().c;
  const totalCurrentStock = db.prepare('SELECT COALESCE(SUM(current_qty),0) AS c FROM products').get().c;
  const unassignedTotal = db.prepare(`
    SELECT COALESCE(SUM(qty),0) AS c FROM stock_by_location WHERE location = ?
  `).get(UNASSIGNED).c;
  const assignedTotal = db.prepare(`
    SELECT COALESCE(SUM(qty),0) AS c FROM stock_by_location WHERE location != ? AND qty > 0
  `).get(UNASSIGNED).c;

  const today = db.prepare(`
    SELECT COUNT(*) AS cnt, COALESCE(SUM(qty),0) AS qty
    FROM transactions
    WHERE type='SALE' AND cancelled_at IS NULL AND date(created_at)=date('now','localtime')
  `).get();
  const todayRestock = db.prepare(`
    SELECT COUNT(*) AS cnt, COALESCE(SUM(qty),0) AS qty
    FROM transactions
    WHERE type='RESTOCK' AND cancelled_at IS NULL AND date(created_at)=date('now','localtime')
  `).get();

  const lowStockCount = db.prepare('SELECT COUNT(*) AS c FROM products WHERE current_qty <= 0').get().c;
  const lowStockList = db.prepare(`
    SELECT barcode, style_name, color_name, size, initial_qty, current_qty
    FROM products WHERE current_qty <= 5
    ORDER BY current_qty ASC, barcode ASC LIMIT 100
  `).all();

  const locations = db.prepare(`
    SELECT location, SUM(qty) AS total_qty, COUNT(DISTINCT barcode) AS bc_count, MAX(updated_at) AS last_at
    FROM stock_by_location
    WHERE qty > 0
    GROUP BY location
    ORDER BY last_at DESC
  `).all();

  // 재고 불일치 (current_qty != SUM(sbl.qty))
  const mismatches = db.prepare(`
    SELECT p.barcode, p.style_name, p.color_name, p.size,
           p.current_qty, COALESCE(s.s, 0) AS sbl_sum
    FROM products p
    LEFT JOIN (SELECT barcode, SUM(qty) s FROM stock_by_location GROUP BY barcode) s
      ON s.barcode = p.barcode
    WHERE p.current_qty != COALESCE(s.s, 0)
    ORDER BY ABS(p.current_qty - COALESCE(s.s, 0)) DESC
    LIMIT 100
  `).all();

  res.json({
    ok: true,
    totalProducts,
    totalInitialStock, totalCurrentStock,
    unassignedTotal, assignedTotal,
    todaySales: { count: today.cnt, qty: today.qty },
    todayRestock: { count: todayRestock.cnt, qty: todayRestock.qty },
    lowStockCount, lowStockList,
    locations,
    mismatchCount: mismatches.length,
    mismatches,
  });
});

// ============================================================================
// CSV 내보내기 (위치별 행)
// ============================================================================
app.get('/api/export', (req, res) => {
  const rows = db.prepare(`
    SELECT p.barcode, p.style, p.style_name, p.color_name, p.size,
           p.initial_qty, p.current_qty,
           (p.initial_qty - p.current_qty) AS sold_qty,
           s.location, s.qty AS location_qty, s.updated_at
    FROM products p
    LEFT JOIN stock_by_location s ON s.barcode = p.barcode
    ORDER BY p.barcode, s.location
  `).all();
  const headers = [
    '바코드','스타일','스타일명','컬러명','사이즈',
    '초도재고','로케이션','로케이션수량','전체현재재고','판매수량','마지막업데이트',
  ];
  const esc = v => { if (v == null) return ''; const s = String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const lines = [headers.join(',')];
  for (const r of rows) {
    lines.push([
      r.barcode, r.style, r.style_name, r.color_name, r.size,
      r.initial_qty, r.location, r.location_qty, r.current_qty, r.sold_qty, r.updated_at,
    ].map(esc).join(','));
  }
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="stock_by_location_${Date.now()}.csv"`);
  res.send('\uFEFF' + lines.join('\n'));
});

// ============================================================================
app.get('/health', (req, res) => res.json({ ok: true, ts: Date.now() }));
app.listen(PORT, HOST, () => console.log(`[SERVER] http://${HOST}:${PORT}`));
