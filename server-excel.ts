import * as XLSX from 'xlsx';
import { STOCK_ITEMS_MAP, STOCK_ITEMS_LIST } from './src/types.js';

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
    'สินค้า': row.nameThai,
    'คงเหลือยกมา (เมื่อวาน)': row.yesterdayCount,
    'เติมสต๊อก (+)': row.replenished,
    'ยอดขายตามระบบ (-)': row.sold,
    'คงเหลือตามทฤษฎี (สูตร)': row.expectedRemaining,
    'คงเหลือนับจริง': row.actualRemaining,
    'ผลต่าง (หายไป / เกิน)': row.difference === 0 ? 'ตรงกัน' : row.difference > 0 ? `หายไป ${row.difference}` : `เกินมา ${Math.abs(row.difference)}`,
  }));
  const wsComp = XLSX.utils.json_to_sheet(compData);
  XLSX.utils.book_append_sheet(wb, wsComp, 'เปรียบเทียบยอดขายและคงเหลือ');

  // 2. Sheet: สต๊อกคงเหลือปัจจุบัน (Current Inventory)
  const invData = inventory.map(item => ({
    'รหัสสินค้า': item.code,
    'ชื่อสินค้า': item.nameThai,
    'จำนวนคงเหลือล่าสุด': item.currentQty,
    'หน่วย': item.unit,
  }));
  const wsInv = XLSX.utils.json_to_sheet(invData);
  XLSX.utils.book_append_sheet(wb, wsInv, 'สต๊อกคงเหลือปัจจุบัน');

  // 3. Sheet: ข้อมูลการเติมสต๊อก (Replenishments)
  const repData = replenishments.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'รหัสสินค้า': row.itemCode,
      'ชื่อสินค้า': item.nameThai,
      'จำนวนที่เติม': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่เติม': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsRep = XLSX.utils.json_to_sheet(repData);
  XLSX.utils.book_append_sheet(wb, wsRep, 'ประวัติการเติมสต๊อก');

  // 4. Sheet: ข้อมูลคงเหลือรายวัน (Daily Counts)
  const cntData = dailyCounts.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'วันที่เช็คสต๊อก': row.date,
      'รหัสสินค้า': row.itemCode,
      'ชื่อสินค้า': item.nameThai,
      'จำนวนที่นับได้': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่บันทึก': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsCnt = XLSX.utils.json_to_sheet(cntData);
  XLSX.utils.book_append_sheet(wb, wsCnt, 'ยอดคงเหลือรายวัน');

  // 5. Sheet: ข้อมูลการขายรายวัน (Sales)
  const saleData = sales.map(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: '' };
    return {
      'วันที่ขาย': row.date,
      'รหัสสินค้า': row.itemCode,
      'ชื่อสินค้า': item.nameThai,
      'จำนวนที่ขาย': row.qty,
      'หน่วย': item.unit,
      'วันเวลาที่บันทึก': new Date(row.timestamp).toLocaleString('th-TH'),
    };
  });
  const wsSale = XLSX.utils.json_to_sheet(saleData);
  XLSX.utils.book_append_sheet(wb, wsSale, 'ยอดขายรายวัน');

  // Return a buffer
  const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return excelBuffer;
}

// Generate an empty/sample Sales Excel file for the user to download as template
export function generateSampleSalesTemplate() {
  const wb = XLSX.utils.book_new();
  const templateData = STOCK_ITEMS_LIST.map(item => ({
    'สินค้า': item.nameThai,
    'จำนวนที่ขาย': 0,
    'หน่วย': item.unit,
    'รหัส (ห้ามลบ)': item.code
  }));
  const ws = XLSX.utils.json_to_sheet(templateData);
  XLSX.utils.book_append_sheet(wb, ws, 'Sales Data');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// Parse daily sales Excel file
// Handles flexible columns and only parses rows under the "PIZZA" header.
// Applies specific restaurant recipe deduction rules for cheese, seafood, and salmon.
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

  let inPizzaSection = false;

  for (const row of rows) {
    if (!row || row.length === 0) continue;

    if (!inPizzaSection) {
      // Check if any cell in this row contains "PIZZA" to identify the main PIZZA block starting point
      const containsPizzaHeader = row.some(cell => 
        cell !== null && 
        cell !== undefined && 
        typeof cell === 'string' && 
        cell.toUpperCase().includes('PIZZA')
      );

      if (containsPizzaHeader) {
        inPizzaSection = true;
        continue; // Skip the header row itself
      }
    }

    if (inPizzaSection) {
      const colA = String(row[0] || '').trim(); // Code / Index
      const colB = String(row[1] || '').trim(); // Product Name
      const colCVal = row[2]; // Quantity sold

      // If both colA and colB are empty, we might have hit a blank separator or the end of the data block
      if (!colA && !colB) {
        continue;
      }

      const matchedNameLower = colB.toLowerCase();
      const colALower = colA.toLowerCase();

      // Skip summary, subtotal or total rows to prevent double counting
      if (
        matchedNameLower.includes('รวม') || 
        matchedNameLower.includes('total') || 
        matchedNameLower.includes('sum') ||
        colALower.includes('รวม') ||
        colALower.includes('total') ||
        colALower.includes('sum')
      ) {
        continue;
      }

      const qtyVal = parseFloat(String(colCVal || '0'));
      if (isNaN(qtyVal) || qtyVal <= 0) {
        continue;
      }

      // --- SPECIFIC INGREDIENT DEDUCTION RULES ---

      // Rule 1: "การตัดสต๊อกนับเป็นพิซซ่าทุกหน้าตัดชีสแค่ 1 แต่ถ้าเป็นดับเบิ้ลตัด 1 ครึ่ง"
      const isPizza = matchedNameLower.includes('พิซซ่า') || matchedNameLower.includes('pizza');
      const isDouble = matchedNameLower.includes('ดับเบิ้ล') || matchedNameLower.includes('double');

      if (isPizza) {
        const cheeseCut = isDouble ? 1.5 : 1.0;
        salesMap['cheese'] += (cheeseCut * qtyVal);
      }

      // Rule 2: "ส่วนชุดทะเลตัด 1 แค่กับเมนู พิซซ่าซีฟู๊ดและสปาเก็ตตี้หมึกดำพริกทะเลกระเทียมเมนูอื่นไม่ตัดชุดทะเล"
      // Must be Pizza and contain "ซีฟู้ด", "ซีฟู๊ด", or "seafood"
      const isSeafoodPizza = isPizza && (matchedNameLower.includes('ซีฟู้ด') || matchedNameLower.includes('ซีฟู๊ด') || matchedNameLower.includes('seafood'));
      // Matches "สปาเก็ตตี้หมึกดำพริกกระเทียม", "สปาเก็ตตี้หมึกดำพริกทะเลกระเทียม", or spaghetti black
      const isSpaghettiBlack = matchedNameLower.includes('หมึกดำ') || matchedNameLower.includes('spaghetti black') || (matchedNameLower.includes('พริกกระเทียม') && matchedNameLower.includes('สปาเก็ตตี้'));
      
      if (isSeafoodPizza || isSpaghettiBlack) {
        salesMap['seafood_set'] += (1.0 * qtyVal);
      }

      // Rule 3: "แซลมอนตัดแค่พิซซ่าแซลมอนอย่าอื่นไม่ตัด"
      const isSalmonPizza = isPizza && (matchedNameLower.includes('แซลมอน') || matchedNameLower.includes('salmon'));
      if (isSalmonPizza) {
        salesMap['salmon'] += (1.0 * qtyVal);
      }

      // Handle other standard items that aren't restricted
      if (matchedNameLower.includes('ซาลามี่') || matchedNameLower.includes('salami')) {
        salesMap['beef_salami'] += qtyVal;
      }
      if (matchedNameLower.includes('พามาแฮม') || matchedNameLower.includes('pama ham')) {
        salesMap['parma_ham'] += qtyVal;
      }
      if (matchedNameLower.includes('เห็ดแชมปิญอง') || matchedNameLower.includes('truffle') || matchedNameLower.includes('ทรัฟเฟิล')) {
        salesMap['champignon_mushroom'] += qtyVal;
      }
      if (matchedNameLower.includes('แฮม') || matchedNameLower.includes('ham')) {
        // Exclude parma ham or salmon pizza from double counting general ham
        if (!matchedNameLower.includes('พามา') && !matchedNameLower.includes('pama')) {
          salesMap['ham'] += qtyVal;
        }
      }
      if (matchedNameLower.includes('เบคอน') || matchedNameLower.includes('bacon')) {
        salesMap['bacon'] += qtyVal;
      }
      if (matchedNameLower.includes('เฟรนช์ฟราย') || matchedNameLower.includes('เฟรนฟราย') || matchedNameLower.includes('french fries')) {
        salesMap['french_fries'] += qtyVal;
      }
      if (matchedNameLower.includes('พ็อกชอป') || matchedNameLower.includes('pork chop')) {
        salesMap['pork_chop'] += qtyVal;
      }
      if (matchedNameLower.includes('สเต็กเนื้อ') || matchedNameLower.includes('tenderloin')) {
        salesMap['beef_steak'] += qtyVal;
      }
      if (matchedNameLower.includes('ซาโมซ่า') || matchedNameLower.includes('samosa')) {
        salesMap['banana_samosa'] += qtyVal;
      }
    }
  }

  return salesMap;
}
