import fs from 'fs';
import path from 'path';
import mysql, { ConnectionOptions } from 'mysql2/promise';
import { STOCK_ITEMS_MAP, STOCK_ITEMS_LIST, StockItem, ReplenishmentRecord, DailyCountRecord, SalesRecord, DiscrepancyReport, DiscrepancyRemark } from './src/types.js';

// Setup local folder for database backup
const DATA_DIR = path.join(process.cwd(), 'data');
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
const LOCAL_DB_FILE = path.join(DATA_DIR, 'local_db.json');

// Interface for local JSON database structure
interface LocalDB {
  inventory: Record<string, number>; // itemCode -> currentQty
  replenishments: ReplenishmentRecord[];
  dailyCounts: DailyCountRecord[];
  sales: SalesRecord[];
  discrepancyRemarks?: DiscrepancyRemark[];
}

// Default stock values to make it look great out-of-the-box
const DEFAULT_STOCK_VALUES: Record<string, number> = {
  cheese: 45,
  seafood_set: 30,
  salmon: 25,
  ham: 50,
  minced_pork: 15,
  bacon: 40,
  french_fries: 35,
  parma_ham: 20,
  pork_chop: 28,
  beef_steak: 18,
  banana_samosa: 60,
  tuna: 32,
  beef_salami: 15,
  champignon_mushroom: 24,
  shrimp: 12,
  squid: 14,
  clam: 20,
};

let mysqlPool: mysql.Pool | null = null;
let isUsingMySQL = false;

// Initialize MySQL pool if credentials exist
export async function initDatabase() {
  const host = process.env.DB_HOST || process.env.MYSQL_HOST;
  const user = process.env.DB_USER || process.env.MYSQL_USER;
  const password = process.env.DB_PASSWORD || process.env.MYSQL_PASSWORD;
  const database = process.env.DB_NAME || process.env.MYSQL_DATABASE;
  const port = parseInt(process.env.DB_PORT || process.env.MYSQL_PORT || '3306', 10);

  if (host && user && database) {
    try {
      console.log(`Connecting to MySQL database at ${host}:${port}...`);
      const connectionConfig: ConnectionOptions = {
        host,
        user,
        password,
        database,
        port,
        waitForConnections: true,
        connectionLimit: 10,
        queueLimit: 0,
      };

      mysqlPool = mysql.createPool(connectionConfig);
      
      // Test the connection
      const connection = await mysqlPool.getConnection();
      console.log('MySQL connection successful!');
      connection.release();
      isUsingMySQL = true;

      // Bootstrap tables
      await bootstrapTables();
    } catch (err) {
      console.error('MySQL connection failed. Falling back to local file storage.', err);
      mysqlPool = null;
      isUsingMySQL = false;
      initLocalDB();
    }
  } else {
    console.log('No MySQL credentials found in env. Running in local JSON storage mode.');
    isUsingMySQL = false;
    initLocalDB();
  }
}

// Local File DB helper functions
function initLocalDB() {
  if (!fs.existsSync(LOCAL_DB_FILE)) {
    const initialDB: LocalDB = {
      inventory: { ...DEFAULT_STOCK_VALUES },
      replenishments: [],
      dailyCounts: [],
      sales: [],
      discrepancyRemarks: []
    };
    fs.writeFileSync(LOCAL_DB_FILE, JSON.stringify(initialDB, null, 2), 'utf-8');
  }
}

function readLocalDB(): LocalDB {
  initLocalDB();
  try {
    const data = fs.readFileSync(LOCAL_DB_FILE, 'utf-8');
    const parsed = JSON.parse(data);
    if (!parsed.discrepancyRemarks) {
      parsed.discrepancyRemarks = [];
    }
    return parsed;
  } catch (err) {
    console.error('Error reading local JSON DB. Resetting database file.', err);
    const initialDB: LocalDB = {
      inventory: { ...DEFAULT_STOCK_VALUES },
      replenishments: [],
      dailyCounts: [],
      sales: [],
      discrepancyRemarks: []
    };
    return initialDB;
  }
}

function writeLocalDB(db: LocalDB) {
  fs.writeFileSync(LOCAL_DB_FILE, JSON.stringify(db, null, 2), 'utf-8');
}

// Bootstrap MySQL Tables
async function bootstrapTables() {
  if (!mysqlPool) return;
  const connection = await mysqlPool.getConnection();
  try {
    console.log('Bootstrapping MySQL tables...');
    
    // 1. Stock Inventory
    await connection.query(`
      CREATE TABLE IF NOT EXISTS stock_inventory (
        code VARCHAR(50) PRIMARY KEY,
        currentQty DOUBLE NOT NULL
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // 2. Stock Replenishments
    await connection.query(`
      CREATE TABLE IF NOT EXISTS stock_replenishments (
        id VARCHAR(50) PRIMARY KEY,
        itemCode VARCHAR(50) NOT NULL,
        qty DOUBLE NOT NULL,
        timestamp VARCHAR(50) NOT NULL
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // 3. Stock Daily Counts (Physical counts)
    await connection.query(`
      CREATE TABLE IF NOT EXISTS stock_daily_counts (
        id VARCHAR(50) PRIMARY KEY,
        itemCode VARCHAR(50) NOT NULL,
        qty DOUBLE NOT NULL,
        date VARCHAR(20) NOT NULL,
        timestamp VARCHAR(50) NOT NULL
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // 4. Sales Records
    await connection.query(`
      CREATE TABLE IF NOT EXISTS sales_records (
        id VARCHAR(50) PRIMARY KEY,
        itemCode VARCHAR(50) NOT NULL,
        qty DOUBLE NOT NULL,
        date VARCHAR(20) NOT NULL,
        timestamp VARCHAR(50) NOT NULL
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // 5. Stock Discrepancy Remarks
    await connection.query(`
      CREATE TABLE IF NOT EXISTS stock_discrepancy_remarks (
        itemCode VARCHAR(50) NOT NULL,
        date VARCHAR(20) NOT NULL,
        remark TEXT NOT NULL,
        PRIMARY KEY (itemCode, date)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // Seed default values in MySQL if stock_inventory is empty
    const [rows]: [any[], any] = await connection.query('SELECT COUNT(*) as count FROM stock_inventory');
    if (rows[0].count === 0) {
      console.log('Seeding initial stock levels into MySQL...');
      for (const [code, qty] of Object.entries(DEFAULT_STOCK_VALUES)) {
        await connection.query('INSERT INTO stock_inventory (code, currentQty) VALUES (?, ?)', [code, qty]);
      }
    }
    console.log('MySQL bootstrapping complete!');
  } catch (err) {
    console.error('Error bootstrapping MySQL tables:', err);
    throw err;
  } finally {
    connection.release();
  }
}

// DB SERVICE METHODS

export async function getIsUsingMySQL() {
  return isUsingMySQL;
}

export async function getInventory(): Promise<StockItem[]> {
  if (isUsingMySQL && mysqlPool) {
    try {
      const [rows]: [any[], any] = await mysqlPool.query('SELECT code, currentQty FROM stock_inventory');
      const map = new Map<string, number>();
      rows.forEach(r => map.set(r.code, r.currentQty));

      return STOCK_ITEMS_LIST.map(item => ({
        ...item,
        currentQty: map.has(item.code) ? map.get(item.code)! : 0
      }));
    } catch (err) {
      console.error('Error fetching inventory from MySQL. Falling back to JSON.', err);
    }
  }

  // Fallback / Local
  const db = readLocalDB();
  return STOCK_ITEMS_LIST.map(item => ({
    ...item,
    currentQty: db.inventory[item.code] !== undefined ? db.inventory[item.code] : 0
  }));
}

export async function replenishStock(itemCode: string, qty: number, date?: string): Promise<boolean> {
  if (!STOCK_ITEMS_MAP[itemCode]) return false;
  const id = `rep_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
  const timestamp = date ? `${date}T12:00:00.000Z` : new Date().toISOString();

  if (isUsingMySQL && mysqlPool) {
    try {
      await mysqlPool.query('START TRANSACTION');
      // Record replenishment
      await mysqlPool.query(
        'INSERT INTO stock_replenishments (id, itemCode, qty, timestamp) VALUES (?, ?, ?, ?)',
        [id, itemCode, qty, timestamp]
      );
      // Update inventory (increment)
      await mysqlPool.query(
        'INSERT INTO stock_inventory (code, currentQty) VALUES (?, ?) ON DUPLICATE KEY UPDATE currentQty = currentQty + ?',
        [itemCode, qty, qty]
      );
      await mysqlPool.query('COMMIT');
      return true;
    } catch (err) {
      if (mysqlPool) await mysqlPool.query('ROLLBACK');
      console.error('MySQL replenish transaction failed. Falling back to local file.', err);
    }
  }

  // Local JSON write
  const db = readLocalDB();
  const current = db.inventory[itemCode] || 0;
  db.inventory[itemCode] = current + qty;
  db.replenishments.push({ id, itemCode, qty, timestamp });
  writeLocalDB(db);
  return true;
}

export async function recordDailyCount(itemCode: string, qty: number, date: string): Promise<boolean> {
  if (!STOCK_ITEMS_MAP[itemCode]) return false;
  const id = `cnt_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
  const timestamp = new Date().toISOString();

  if (isUsingMySQL && mysqlPool) {
    try {
      await mysqlPool.query('START TRANSACTION');
      // Remove any existing count for this item on this date to overwrite
      await mysqlPool.query(
        'DELETE FROM stock_daily_counts WHERE itemCode = ? AND date = ?',
        [itemCode, date]
      );
      // Insert new daily count
      await mysqlPool.query(
        'INSERT INTO stock_daily_counts (id, itemCode, qty, date, timestamp) VALUES (?, ?, ?, ?, ?)',
        [id, itemCode, qty, date, timestamp]
      );
      // Update inventory to reflect actual physical count (as it is the absolute physical truth)
      await mysqlPool.query(
        'INSERT INTO stock_inventory (code, currentQty) VALUES (?, ?) ON DUPLICATE KEY UPDATE currentQty = ?',
        [itemCode, qty, qty]
      );
      await mysqlPool.query('COMMIT');
      return true;
    } catch (err) {
      if (mysqlPool) await mysqlPool.query('ROLLBACK');
      console.error('MySQL physical count record failed:', err);
    }
  }

  // Local JSON write
  const db = readLocalDB();
  db.inventory[itemCode] = qty;
  // Overwrite local duplicate on same date
  db.dailyCounts = db.dailyCounts.filter(c => !(c.itemCode === itemCode && c.date === date));
  db.dailyCounts.push({ id, itemCode, qty, date, timestamp });
  writeLocalDB(db);
  return true;
}

export async function recordSales(date: string, salesMap: Record<string, number>): Promise<boolean> {
  const timestamp = new Date().toISOString();

  if (isUsingMySQL && mysqlPool) {
    try {
      await mysqlPool.query('START TRANSACTION');
      
      // Delete existing sales records for this date
      await mysqlPool.query('DELETE FROM sales_records WHERE date = ?', [date]);

      // Record new sales and deduct from inventory
      for (const [itemCode, qty] of Object.entries(salesMap)) {
        if (!STOCK_ITEMS_MAP[itemCode]) continue;
        const id = `sale_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
        
        await mysqlPool.query(
          'INSERT INTO sales_records (id, itemCode, qty, date, timestamp) VALUES (?, ?, ?, ?, ?)',
          [id, itemCode, qty, date, timestamp]
        );

        // Deduct from current inventory
        await mysqlPool.query(
          'INSERT INTO stock_inventory (code, currentQty) VALUES (?, 0) ON DUPLICATE KEY UPDATE currentQty = GREATEST(0, currentQty - ?)',
          [itemCode, qty]
        );
      }
      
      await mysqlPool.query('COMMIT');
      return true;
    } catch (err) {
      if (mysqlPool) await mysqlPool.query('ROLLBACK');
      console.error('MySQL record sales failed:', err);
    }
  }

  // Local JSON write
  const db = readLocalDB();
  // Filter out old sales for this date
  db.sales = db.sales.filter(s => s.date !== date);

  for (const [itemCode, qty] of Object.entries(salesMap)) {
    if (!STOCK_ITEMS_MAP[itemCode]) continue;
    const id = `sale_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    db.sales.push({ id, itemCode, qty, date, timestamp });

    // Deduct stock
    const current = db.inventory[itemCode] || 0;
    db.inventory[itemCode] = Math.max(0, current - qty);
  }

  writeLocalDB(db);
  return true;
}

export async function getReplenishments(): Promise<ReplenishmentRecord[]> {
  if (isUsingMySQL && mysqlPool) {
    try {
      const [rows]: [any[], any] = await mysqlPool.query('SELECT * FROM stock_replenishments ORDER BY timestamp DESC');
      return rows;
    } catch (err) {
      console.error('MySQL getReplenishments failed:', err);
    }
  }
  return readLocalDB().replenishments.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

export async function getDailyCounts(): Promise<DailyCountRecord[]> {
  if (isUsingMySQL && mysqlPool) {
    try {
      const [rows]: [any[], any] = await mysqlPool.query('SELECT * FROM stock_daily_counts ORDER BY timestamp DESC');
      return rows;
    } catch (err) {
      console.error('MySQL getDailyCounts failed:', err);
    }
  }
  return readLocalDB().dailyCounts.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

export async function getSales(): Promise<SalesRecord[]> {
  if (isUsingMySQL && mysqlPool) {
    try {
      const [rows]: [any[], any] = await mysqlPool.query('SELECT * FROM sales_records ORDER BY timestamp DESC');
      return rows;
    } catch (err) {
      console.error('MySQL getSales failed:', err);
    }
  }
  return readLocalDB().sales.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

// Generate the beautiful reconciliation report:
// Subtracts today's count from yesterday's count to find actual usage,
// and compares with recorded sales.
// expectedRemaining = yesterdayCount + replenishment - sales
export async function getDiscrepancyReport(targetDate: string): Promise<DiscrepancyReport[]> {
  // Get all data
  const inventory = await getInventory();
  const allCounts = await getDailyCounts();
  const allReplenishments = await getReplenishments();
  const allSales = await getSales();

  // Find yesterday's date string
  const d = new Date(targetDate);
  d.setDate(d.getDate() - 1);
  const yesterdayDate = d.toISOString().split('T')[0];

  // Map counts
  const todayCountsMap = new Map<string, number>();
  allCounts.filter(c => c.date === targetDate).forEach(c => todayCountsMap.set(c.itemCode, c.qty));

  const yesterdayCountsMap = new Map<string, number>();
  allCounts.filter(c => c.date === yesterdayDate).forEach(c => yesterdayCountsMap.set(c.itemCode, c.qty));

  // Map replenishments for targetDate
  const repsMap = new Map<string, number>();
  allReplenishments
    .filter(r => {
      if (!r.timestamp) return false;
      const tStr = String(r.timestamp);
      return tStr.startsWith(targetDate);
    })
    .forEach(r => repsMap.set(r.itemCode, (repsMap.get(r.itemCode) || 0) + r.qty));

  // Map sales for targetDate
  const salesMap = new Map<string, number>();
  allSales.filter(s => s.date === targetDate).forEach(s => salesMap.set(s.itemCode, s.qty));

  // Map remarks for targetDate
  const remarksMap = new Map<string, string>();
  const remarks = await getDiscrepancyRemarks(targetDate);
  remarks.forEach(r => remarksMap.set(r.itemCode, r.remark));

  return inventory.map(item => {
    // Yesterday count defaults to current stock if not recorded, or a logical mock
    const yesterdayCount = yesterdayCountsMap.has(item.code) 
      ? yesterdayCountsMap.get(item.code)! 
      : (item.currentQty + (salesMap.get(item.code) || 0) - (repsMap.get(item.code) || 0)); // estimate yesterday if not in DB
    
    const replenished = repsMap.get(item.code) || 0;
    const sold = salesMap.get(item.code) || 0;
    
    // Expected remaining based on sales ledger
    const expectedRemaining = Math.max(0, yesterdayCount + replenished - sold);
    
    // Actual remaining physical count reported today. If not recorded today, show expected or fallback to current
    const actualRemaining = todayCountsMap.has(item.code)
      ? todayCountsMap.get(item.code)!
      : item.currentQty;

    // Difference between what we EXPECTED to have vs what was ACTUALLY physically counted
    // positive = missing (e.g. expected 10, physical is 8 -> difference is 2 units lost!)
    // negative = surplus (e.g. expected 10, physical is 12 -> difference is -2 units extra!)
    const difference = expectedRemaining - actualRemaining;

    return {
      date: targetDate,
      itemCode: item.code,
      nameThai: item.nameThai,
      yesterdayCount: Math.round(yesterdayCount * 10) / 10,
      replenished: Math.round(replenished * 10) / 10,
      sold: Math.round(sold * 10) / 10,
      expectedRemaining: Math.round(expectedRemaining * 10) / 10,
      actualRemaining: Math.round(actualRemaining * 10) / 10,
      difference: Math.round(difference * 10) / 10,
      remark: remarksMap.get(item.code) || '',
    };
  });
}

export async function getDiscrepancyRemarks(targetDate: string): Promise<DiscrepancyRemark[]> {
  if (isUsingMySQL && mysqlPool) {
    try {
      const [rows]: [any[], any] = await mysqlPool.query(
        'SELECT * FROM stock_discrepancy_remarks WHERE date = ?',
        [targetDate]
      );
      return rows;
    } catch (err) {
      console.error('MySQL getDiscrepancyRemarks failed:', err);
    }
  }
  const db = readLocalDB();
  return (db.discrepancyRemarks || []).filter(r => r.date === targetDate);
}

export async function saveDiscrepancyRemark(itemCode: string, date: string, remark: string): Promise<void> {
  if (isUsingMySQL && mysqlPool) {
    try {
      await mysqlPool.query(
        'INSERT INTO stock_discrepancy_remarks (itemCode, date, remark) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE remark = ?',
        [itemCode, date, remark, remark]
      );
      return;
    } catch (err) {
      console.error('MySQL saveDiscrepancyRemark failed:', err);
    }
  }
  const db = readLocalDB();
  if (!db.discrepancyRemarks) {
    db.discrepancyRemarks = [];
  }
  // Remove existing if any
  db.discrepancyRemarks = db.discrepancyRemarks.filter(r => !(r.itemCode === itemCode && r.date === date));
  if (remark.trim()) {
    db.discrepancyRemarks.push({ itemCode, date, remark });
  }
  writeLocalDB(db);
}
