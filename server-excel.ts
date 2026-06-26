import * as XLSX from 'xlsx';
import { STOCK_ITEMS_MAP, STOCK_ITEMS_LIST } from './src/types.js';

// Auto-fit column widths helper for clean, professional sheets
function autofitColumns(ws: XLSX.WorkSheet) {
  if (!ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const cols: XLSX.ColInfo[] = [];
  
  for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
    let maxLen = 12; // Min width to keep headers readable
    
    for (let rowIdx = range.s.r; rowIdx <= range.e.r; rowIdx++) {
      const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      const cell = ws[cellRef];
      if (cell && cell.v !== undefined && cell.v !== null) {
        const strVal = String(cell.v);
        // Thai characters look better with a bit of extra padding
        const len = Math.ceil(strVal.replace(/[^\x00-\xff]/g, 'xx').length * 0.9) + 4;
        if (len > maxLen) {
          maxLen = len;
        }
      }
    }
    // Set max limit to prevent excessively wide columns
    cols.push({ wch: Math.min(maxLen, 45) });
  }
  ws['!cols'] = cols;
}

// Generate a buffer for the multi-sheet Excel report
export function generateExcelReport(
  inventory: any[],
  replenishments: any[],
  dailyCounts: any[],
  sales: any[],
  discrepancyReport: any[],
  date: string
) {
  const wb = XLSX.utils.book_new();

  // 1. Sheet: สรุปผลต่างและการเปรียบเทียบ (Discrepancy Report)
  const compData = discrepancyReport.map(row => ({
    'วันที่': row.date,
    'วัตถุดิบ': row.nameThai,
    'คงเหลือยกมา (เมื่อวาน)': row.yesterdayCount,
    'เติมสต๊อก (+)': row.replenished,
    'ยอดขายตามระบบ (-)': row.sold,
    'คงเหลือตามทฤษฎี (สูตร)': row.expectedRemaining,
    'คงเหลือนับจริง': row.actualRemaining,
    'ผลต่าง (หายไป / เกิน)': row.difference === 0 ? 'ตรงกัน (0)' : row.difference > 0 ? `หายไป -${row.difference}` : `เกินมา +${Math.abs(row.difference)}`,
    'หมายเหตุประจำวัน': row.remark || '',
  }));
  const wsComp = XLSX.utils.json_to_sheet(compData);
  autofitColumns(wsComp);
  XLSX.utils.book_append_sheet(wb, wsComp, 'เปรียบเทียบยอดขายและคงเหลือ');

  // 2. Sheet: สต๊อกคงเหลือปัจจุบัน (Current Inventory)
  const invData = inventory.map(item => ({
    'รหัสวัตถุดิบ': item.code,
    'ชื่อวัตถุดิบ': item.nameThai,
    'จำนวนคงเหลือในระบบ': item.currentQty,
    'หน่วย': item.unit,
  }));
  const wsInv = XLSX.utils.json_to_sheet(invData);
  autofitColumns(wsInv);
  XLSX.utils.book_append_sheet(wb, wsInv, 'สต๊อกคงเหลือปัจจุบัน');

  // 3. Sheet: ข้อมูลการเติมสต๊อก (Replenishments)
  const repData = replenishments.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'รหัสวัตถุดิบ': row.itemCode,
      'ชื่อวัตถุดิบ': item.nameThai,
      'จำนวนที่เติม': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่เติม': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsRep = XLSX.utils.json_to_sheet(repData);
  autofitColumns(wsRep);
  XLSX.utils.book_append_sheet(wb, wsRep, 'ประวัติการเติมสต๊อก');

  // 4. Sheet: ข้อมูลคงเหลือรายวัน (Daily Counts)
  const cntData = dailyCounts.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'วันที่ตรวจเช็ค': row.date,
      'รหัสวัตถุดิบ': row.itemCode,
      'ชื่อวัตถุดิบ': item.nameThai,
      'จำนวนคงเหลือเช็คจริง': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่บันทึกเช็ค': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsCnt = XLSX.utils.json_to_sheet(cntData);
  autofitColumns(wsCnt);
  XLSX.utils.book_append_sheet(wb, wsCnt, 'ยอดคงเหลือรายวัน');

  // 5. Sheet: ข้อมูลการขายรายวัน (Sales)
  const saleData = sales.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'วันที่ขาย': row.date,
      'รหัสวัตถุดิบ': row.itemCode,
      'ชื่อวัตถุดิบ': item.nameThai,
      'จำนวนวัตถุดิบที่ใช้': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่บันทึกขาย': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsSale = XLSX.utils.json_to_sheet(saleData);
  autofitColumns(wsSale);
  XLSX.utils.book_append_sheet(wb, wsSale, 'ยอดขายรายวัน');

  // Return a buffer
  const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return excelBuffer;
}

// Generate an empty/sample Sales Excel file for the user to download as template
export function generateSampleSalesTemplate() {
  const wb = XLSX.utils.book_new();
  const templateData = STOCK_ITEMS_LIST.map(item => ({
    'วัตถุดิบ / สินค้า': item.nameThai,
    'จำนวนที่ขาย': 0,
    'หน่วย': item.unit,
    'รหัส (ห้ามลบ)': item.code
  }));
  const ws = XLSX.utils.json_to_sheet(templateData);
  autofitColumns(ws);
  XLSX.utils.book_append_sheet(wb, ws, 'Sales Data Template');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// Parse daily sales Excel file
// Handles BOTH template uploads (direct ingredients) and POS exports (recipe dishes)
export function parseSalesExcel(fileBuffer: Buffer): Record<string, number> {
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  
  // Use header: 1 to retrieve raw matrix of values
  const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });

  const salesMap: Record<string, number> = {};

  // Seed default 0
  STOCK_ITEMS_LIST.forEach(item => {
    salesMap[item.code] = 0;
  });

  for (const row of rows) {
    if (!row || row.length === 0) continue;

    // Convert row cells to string list and look for numbers
    const cellStrings = row.map(cell => cell !== null && cell !== undefined ? String(cell).trim() : '');
    
    // Scan row for quantity values (exclude index columns)
    let bestQty = 0;
    for (let i = row.length - 1; i >= 0; i--) {
      const cellVal = row[i];
      if (cellVal === null || cellVal === undefined) continue;
      
      const parsed = parseFloat(String(cellVal).trim());
      if (!isNaN(parsed) && parsed > 0) {
        // Prevent matching sequential ID columns (index 0) if the row is wide
        if (i > 0 || row.length <= 2) {
          bestQty = parsed;
          break;
        }
      }
    }

    if (bestQty <= 0) continue;

    // A. Check for DIRECT CODE MATCH (Template Mode)
    let matchedDirectCode = '';
    for (const cellStr of cellStrings) {
      const cleanCell = cellStr.toLowerCase();
      if (STOCK_ITEMS_MAP[cleanCell]) {
        matchedDirectCode = cleanCell;
        break;
      }
    }

    if (matchedDirectCode) {
      salesMap[matchedDirectCode] += bestQty;
      continue; // Process direct code matches immediately without recipe translation
    }

    // B. Check for DISH/PRODUCT MATCH (POS Report Mode with recipe rules)
    let productName = '';
    for (const cellStr of cellStrings) {
      const cellStrLower = cellStr.toLowerCase();
      // Skip cells that match the quantity
      if (parseFloat(cellStr) === bestQty) continue;

      const isKnownDish = 
        cellStr.includes('พิซซ่า') || 
        cellStrLower.includes('pizza') ||
        cellStr.includes('สปาเก็ตตี้') ||
        cellStrLower.includes('spaghetti') ||
        cellStr.includes('เฟรนช์ฟราย') ||
        cellStr.includes('เฟรนฟราย') ||
        cellStrLower.includes('fries') ||
        cellStr.includes('พ็อกชอป') ||
        cellStrLower.includes('chop') ||
        cellStr.includes('สเต็กเนื้อ') ||
        cellStrLower.includes('steak') ||
        cellStr.includes('ซาโมซ่า') ||
        cellStrLower.includes('samosa') ||
        cellStr.includes('เบคอน') ||
        cellStrLower.includes('bacon') ||
        cellStr.includes('แฮม') ||
        cellStrLower.includes('ham') ||
        cellStr.includes('ซาลามี่') ||
        cellStrLower.includes('salami') ||
        cellStr.includes('แซลมอน') ||
        cellStrLower.includes('salmon') ||
        cellStr.includes('เห็ด') ||
        cellStrLower.includes('mushroom') ||
        cellStr.includes('กุ้ง') ||
        cellStr.includes('หมึก') ||
        cellStr.includes('หอย') ||
        cellStr.includes('ทะเล') ||
        cellStrLower.includes('seafood');

      if (isKnownDish && cellStr.length > 2) {
        productName = cellStr;
        break;
      }
    }

    if (!productName) continue;

    const matchedNameLower = productName.toLowerCase();

    // Skip headers or summary rows
    if (
      matchedNameLower.includes('รวม') || 
      matchedNameLower.includes('total') || 
      matchedNameLower.includes('sum')
    ) {
      continue;
    }

    // --- APPLY SPECIFIC RESTAURANT INGREDIENT DEDUCTION RECIPES ---

    // Rule 1: "การตัดสต๊อกนับเป็นพิซซ่าทุกหน้าตัดชีสแค่ 1 แต่ถ้าเป็นดับเบิ้ลตัด 1 ครึ่ง"
    const isPizza = matchedNameLower.includes('พิซซ่า') || matchedNameLower.includes('pizza');
    const isDouble = matchedNameLower.includes('ดับเบิ้ล') || matchedNameLower.includes('double');

    if (isPizza) {
      const cheeseCut = isDouble ? 1.5 : 1.0;
      salesMap['cheese'] += (cheeseCut * bestQty);
    }

    // Rule 2: "ส่วนชุดทะเลตัด 1 แค่กับเมนู พิซซ่าซีฟู๊ดและสปาเก็ตตี้หมึกดำพริกทะเลกระเทียมเมนูอื่นไม่ตัดชุดทะเล"
    const isSeafoodPizza = isPizza && (matchedNameLower.includes('ซีฟู้ด') || matchedNameLower.includes('ซีฟู๊ด') || matchedNameLower.includes('seafood'));
    const isSpaghettiBlack = matchedNameLower.includes('หมึกดำ') || matchedNameLower.includes('spaghetti black') || (matchedNameLower.includes('พริกกระเทียม') && matchedNameLower.includes('สปาเก็ตตี้'));
    
    if (isSeafoodPizza || isSpaghettiBlack) {
      salesMap['seafood_set'] += (1.0 * bestQty);
    }

    // Rule 3: "แซลมอนตัดแค่พิซซ่าแซลมอนอย่าอื่นไม่ตัด"
    const isSalmonPizza = isPizza && (matchedNameLower.includes('แซลมอน') || matchedNameLower.includes('salmon'));
    if (isSalmonPizza) {
      salesMap['salmon'] += (1.0 * bestQty);
    }

    // Standard items matching
    if (matchedNameLower.includes('ซาลามี่') || matchedNameLower.includes('salami')) {
      salesMap['beef_salami'] += bestQty;
    }
    if (matchedNameLower.includes('พามาแฮม') || matchedNameLower.includes('pama ham') || matchedNameLower.includes('parma')) {
      salesMap['parma_ham'] += bestQty;
    }
    if (matchedNameLower.includes('เห็ดแชมปิญอง') || matchedNameLower.includes('truffle') || matchedNameLower.includes('ทรัฟเฟิล') || matchedNameLower.includes('champignon')) {
      salesMap['champignon_mushroom'] += bestQty;
    }
    if (matchedNameLower.includes('แฮม') || matchedNameLower.includes('ham')) {
      // Exclude parma ham to avoid double counting general ham
      if (!matchedNameLower.includes('พามา') && !matchedNameLower.includes('pama') && !matchedNameLower.includes('parma')) {
        salesMap['ham'] += bestQty;
      }
    }
    if (matchedNameLower.includes('เบคอน') || matchedNameLower.includes('bacon')) {
      salesMap['bacon'] += bestQty;
    }
    if (matchedNameLower.includes('เฟรนช์ฟราย') || matchedNameLower.includes('เฟรนฟราย') || matchedNameLower.includes('french fries') || matchedNameLower.includes('fries')) {
      salesMap['french_fries'] += bestQty;
    }
    if (matchedNameLower.includes('พ็อกชอป') || matchedNameLower.includes('pork chop') || matchedNameLower.includes('porkchop')) {
      salesMap['pork_chop'] += bestQty;
    }
    if (matchedNameLower.includes('สเต็กเนื้อ') || matchedNameLower.includes('tenderloin') || matchedNameLower.includes('steak')) {
      salesMap['beef_steak'] += bestQty;
    }
    if (matchedNameLower.includes('ซาโมซ่า') || matchedNameLower.includes('samosa')) {
      salesMap['banana_samosa'] += bestQty;
    }
    if (matchedNameLower.includes('ทูน่า') || matchedNameLower.includes('tuna')) {
      salesMap['tuna'] += bestQty;
    }
    if (matchedNameLower.includes('หมูบด') || matchedNameLower.includes('minced pork')) {
      salesMap['minced_pork'] += bestQty;
    }
    if (matchedNameLower.includes('กุ้ง') || matchedNameLower.includes('shrimp') || matchedNameLower.includes('prawn')) {
      salesMap['shrimp'] += bestQty;
    }
    if (matchedNameLower.includes('หมึก') || matchedNameLower.includes('squid') || matchedNameLower.includes('octopus')) {
      salesMap['squid'] += bestQty;
    }
    if (matchedNameLower.includes('หอย') || matchedNameLower.includes('clam') || matchedNameLower.includes('mussel')) {
      salesMap['clam'] += bestQty;
    }
  }

  return salesMap;
}
