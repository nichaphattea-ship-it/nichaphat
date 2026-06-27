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

// Helper to style a worksheet with high visual appeal (Titles, Merges, Heights, Formats)
function styleWorksheet(
  ws: XLSX.WorkSheet, 
  title: string, 
  subtitle: string, 
  dataStartRow: number, 
  numCols: number,
  numericColIndices: number[]
) {
  // 1. Setup Merges for Title and Subtitle
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: numCols - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: numCols - 1 } }
  ];

  // 2. Set Row Heights (spacious and modern padding)
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A5');
  const rowHeights: XLSX.RowInfo[] = [];
  rowHeights.push({ hpt: 35 }); // Title row
  rowHeights.push({ hpt: 20 }); // Subtitle row
  rowHeights.push({ hpt: 12 }); // Empty spacer row
  rowHeights.push({ hpt: 28 }); // Table headers (tall & robust)
  
  for (let r = 4; r <= range.e.r; r++) {
    // Check if it's the last row (Total row)
    if (r === range.e.r) {
      rowHeights.push({ hpt: 28 }); // Total row (tall & robust)
    } else {
      rowHeights.push({ hpt: 22 }); // Standard data row (comfortable padding)
    }
  }
  ws['!rows'] = rowHeights;

  // 3. Format Cell Numbers & Formulas
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (!cell) continue;

      // Format numeric data columns starting from row idx 4 (5th Excel row)
      if (r >= 4 && numericColIndices.includes(c)) {
        if (typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = '#,##0.00'; // Formatted with commas and 2 decimals (e.g. 1,234.50)
        } else if (cell.f) {
          cell.t = 'n';
          cell.z = '#,##0.00'; // Keep formulas as numeric formatted as well
        }
      }
    }
  }

  // 4. Auto-fit column widths
  autofitColumns(ws);
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

  // -------------------------------------------------------------------------
  // 1. Sheet: เปรียบเทียบยอดขายและคงเหลือ (Discrepancy Report)
  // -------------------------------------------------------------------------
  const compRows: any[][] = [
    [`📊 รายงานเปรียบเทียบยอดขายและสต๊อกคงเหลือร้านอาหาร (Discrepancy Report)`],
    [`📅 วันที่รายงาน: ${date}  |  🕒 ออกรายงานเมื่อ: ${new Date().toLocaleString('th-TH')}  |  📍 ระบบสต๊อกอัตโนมัติ`],
    [], // Spacer
    [
      'วันที่',
      'วัตถุดิบ / ส่วนผสม',
      'คงเหลือยกมา (เมื่อวาน)',
      'เติมสต๊อก (+)',
      'ยอดขายตามระบบ (-)',
      'คงเหลือตามทฤษฎี (สูตร)',
      'คงเหลือนับจริง',
      'ผลต่าง (หายไป / เกิน)',
      'หมายเหตุประจำวัน / หน้าครึ่ง'
    ]
  ];

  discrepancyReport.forEach(row => {
    const diffText = row.difference === 0 
      ? '🟢 ตรงกัน (0)' 
      : row.difference > 0 
        ? `❌ หายไป -${row.difference}` 
        : `⚠️ เกินมา +${Math.abs(row.difference)}`;

    compRows.push([
      row.date,
      row.nameThai,
      row.yesterdayCount,
      row.replenished,
      row.sold,
      row.expectedRemaining,
      row.actualRemaining,
      diffText,
      row.remark || ''
    ]);
  });

  // Totals Row with Excel formulas!
  const compDataStart = 5; // row index 4, which is row 5 in Excel (1-indexed)
  const compDataEnd = compDataStart + discrepancyReport.length - 1;
  compRows.push([
    'รวมยอดสต๊อกทั้งหมด (Total)',
    '',
    { f: `SUM(C${compDataStart}:C${compDataEnd})` },
    { f: `SUM(D${compDataStart}:D${compDataEnd})` },
    { f: `SUM(E${compDataStart}:E${compDataEnd})` },
    { f: `SUM(F${compDataStart}:F${compDataEnd})` },
    { f: `SUM(G${compDataStart}:G${compDataEnd})` },
    '', // Text status
    ''
  ]);

  const wsComp = XLSX.utils.aoa_to_sheet(compRows);
  // Numeric columns are C (2), D (3), E (4), F (5), G (6)
  styleWorksheet(wsComp, 'Discrepancy Report', date, compDataStart, 9, [2, 3, 4, 5, 6]);
  XLSX.utils.book_append_sheet(wb, wsComp, 'เปรียบเทียบยอดขายและคงเหลือ');


  // -------------------------------------------------------------------------
  // 2. Sheet: สต๊อกคงเหลือปัจจุบัน (Current Inventory)
  // -------------------------------------------------------------------------
  const invRows: any[][] = [
    [`📋 บัญชีรายการสต๊อกคงเหลือปัจจุบันในระบบ (Current Inventory)`],
    [`📅 ประจำวันที่: ${date}  |  🕒 ข้อมูลอัปเดต ณ ปัจจุบัน`],
    [], // Spacer
    [
      'รหัสวัตถุดิบ',
      'ชื่อวัตถุดิบ / ส่วนผสม',
      'จำนวนคงเหลือในระบบ',
      'หน่วยสินค้า'
    ]
  ];

  inventory.forEach(item => {
    invRows.push([
      item.code,
      item.nameThai,
      item.currentQty,
      item.unit
    ]);
  });

  const invDataStart = 5;
  const invDataEnd = invDataStart + inventory.length - 1;
  invRows.push([
    'รวมคงเหลือวัตถุดิบทั้งหมด',
    '',
    { f: `SUM(C${invDataStart}:C${invDataEnd})` },
    ''
  ]);

  const wsInv = XLSX.utils.aoa_to_sheet(invRows);
  // Numeric column is C (2)
  styleWorksheet(wsInv, 'Current Inventory', date, invDataStart, 4, [2]);
  XLSX.utils.book_append_sheet(wb, wsInv, 'สต๊อกคงเหลือปัจจุบัน');


  // -------------------------------------------------------------------------
  // 3. Sheet: ข้อมูลการเติมสต๊อก (Replenishments)
  // -------------------------------------------------------------------------
  const repRows: any[][] = [
    [`📥 ประวัติรายการทำรายการเติมสต๊อกวัตถุดิบ (Replenishments History)`],
    [`📅 ข้อมูลบันทึกประวัติสะสมทั้งหมดในระบบ`],
    [], // Spacer
    [
      'รหัสวัตถุดิบ',
      'ชื่อวัตถุดิบ',
      'จำนวนที่เติมวัตถุดิบ',
      'หน่วยสินค้า',
      'วันเวลาที่ทำการเติมสต๊อก'
    ]
  ];

  replenishments.forEach(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: 'ยูนิต' };
    repRows.push([
      row.itemCode,
      item.nameThai,
      row.qty,
      item.unit,
      new Date(row.timestamp).toLocaleString('th-TH')
    ]);
  });

  const repDataStart = 5;
  const repDataEnd = repDataStart + replenishments.length - 1;
  if (replenishments.length > 0) {
    repRows.push([
      'รวมการเติมวัตถุดิบทั้งหมด',
      '',
      { f: `SUM(C${repDataStart}:C${repDataEnd})` },
      '',
      ''
    ]);
  }

  const wsRep = XLSX.utils.aoa_to_sheet(repRows);
  styleWorksheet(wsRep, 'Replenishments History', date, repDataStart, 5, [2]);
  XLSX.utils.book_append_sheet(wb, wsRep, 'ประวัติการเติมสต๊อก');


  // -------------------------------------------------------------------------
  // 4. Sheet: ข้อมูลคงเหลือรายวัน (Daily Counts)
  // -------------------------------------------------------------------------
  const cntRows: any[][] = [
    [`📝 ยอดสต๊อกคงเหลือนับจริงรายวันหน้าเครื่อง (Daily Physical Counts)`],
    [`📅 ข้อมูลรายงานตรวจสอบนับจริงสะสม`],
    [], // Spacer
    [
      'วันที่เช็คสต๊อก',
      'รหัสวัตถุดิบ',
      'ชื่อวัตถุดิบ',
      'จำนวนนับจริง',
      'หน่วยสินค้า',
      'วันเวลาที่ลงบันทึกนับสต๊อก'
    ]
  ];

  dailyCounts.forEach(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: 'ยูนิต' };
    cntRows.push([
      row.date,
      row.itemCode,
      item.nameThai,
      row.qty,
      item.unit,
      new Date(row.timestamp).toLocaleString('th-TH')
    ]);
  });

  const cntDataStart = 5;
  const cntDataEnd = cntDataStart + dailyCounts.length - 1;
  if (dailyCounts.length > 0) {
    cntRows.push([
      'รวมสต๊อกนับจริงทั้งหมด',
      '',
      '',
      { f: `SUM(D${cntDataStart}:D${cntDataEnd})` },
      '',
      ''
    ]);
  }

  const wsCnt = XLSX.utils.aoa_to_sheet(cntRows);
  styleWorksheet(wsCnt, 'Daily Physical Counts', date, cntDataStart, 6, [3]);
  XLSX.utils.book_append_sheet(wb, wsCnt, 'ยอดคงเหลือรายวัน');


  // -------------------------------------------------------------------------
  // 5. Sheet: ข้อมูลการขายรายวัน (Sales)
  // -------------------------------------------------------------------------
  const saleRows: any[][] = [
    [`📈 สรุปรายการหักตัดยอดวัตถุดิบจากยอดขายรายวัน (Sales Deduction)`],
    [`📅 รายงานบันทึกการหักตัดจำหน่ายสะสม`],
    [], // Spacer
    [
      'วันที่ขายสินค้า',
      'รหัสวัตถุดิบ',
      'ชื่อวัตถุดิบ',
      'จำนวนที่หักสต๊อกตามยอดขาย',
      'หน่วยสินค้า',
      'วันเวลาประมวลผลตัดสต๊อก'
    ]
  ];

  sales.forEach(row => {
    const item = STOCK_ITEMS_MAP[row.itemCode] || { nameThai: row.itemCode, unit: 'ยูนิต' };
    saleRows.push([
      row.date,
      row.itemCode,
      item.nameThai,
      row.qty,
      item.unit,
      new Date(row.timestamp).toLocaleString('th-TH')
    ]);
  });

  const saleDataStart = 5;
  const saleDataEnd = saleDataStart + sales.length - 1;
  if (sales.length > 0) {
    saleRows.push([
      'รวมตัดขายสะสมทั้งหมด',
      '',
      '',
      { f: `SUM(D${saleDataStart}:D${saleDataEnd})` },
      '',
      ''
    ]);
  }

  const wsSale = XLSX.utils.aoa_to_sheet(saleRows);
  styleWorksheet(wsSale, 'Sales Deduction', date, saleDataStart, 6, [3]);
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

  // 1. Find all row indices containing "PIZZA V20" to parse these specific sections
  const pizzaRowIndices: number[] = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const hasPizzaV20 = row.some(cell => {
      if (cell === null || cell === undefined) return false;
      const str = String(cell).toUpperCase().replace(/\s/g, '');
      return str.includes('PIZZAV20');
    });
    if (hasPizzaV20) {
      pizzaRowIndices.push(r);
    }
  }

  if (pizzaRowIndices.length > 0) {
    for (const pizzaRowIdx of pizzaRowIndices) {
      // Found a "PIZZA V20" section! Parse the rows under this section.
      let targetNameCol = 1; // Default: column B
      let targetQtyCol = 2;  // Default: column C
      let targetCodeCol = 0; // Default: column A
      
      let startIdx = pizzaRowIdx + 1;
      
      // Check if the row immediately after is a header row
      if (startIdx < rows.length) {
        const nextRow = rows[startIdx];
        const isHeader = nextRow && nextRow.some(cell => {
          if (cell === null || cell === undefined) return false;
          const str = String(cell).trim().toLowerCase();
          return str.includes('สินค้า') || str.includes('รายการ') || str.includes('จำนวน') || str.includes('รหัส');
        });
        if (isHeader) {
          const cells = nextRow.map(c => c !== null && c !== undefined ? String(c).trim().toLowerCase() : '');
          const nameIdx = cells.findIndex(c => c.includes('สินค้า') || c.includes('รายการ') || c === 'name');
          const qtyIdx = cells.findIndex(c => c.includes('จำนวน') || c === 'qty' || c === 'quantity' || c === 'sold');
          const codeIdx = cells.findIndex(c => c.includes('รหัส') || c === 'code' || c === 'id');
          if (nameIdx !== -1) targetNameCol = nameIdx;
          if (qtyIdx !== -1) targetQtyCol = qtyIdx;
          if (codeIdx !== -1) targetCodeCol = codeIdx;
          startIdx++; // Skip the header row
        }
      }

      for (let r = startIdx; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;

        // Stop condition: any cell containing "TOTAL" or "รวม" (case-insensitive)
        const isStopRow = row.some(cell => {
          if (cell === null || cell === undefined) return false;
          const str = String(cell).trim().toUpperCase();
          return str === 'TOTAL' || str === 'รวม' || str.startsWith('TOTAL ') || str.startsWith('รวม ');
        });
        if (isStopRow) {
          break; // Reached the end of this "PIZZA V20" section, continue to any other sections if any
        }

        const nameVal = row[targetNameCol] !== null && row[targetNameCol] !== undefined ? String(row[targetNameCol]).trim() : '';
        const qtyVal = row[targetQtyCol];
        const codeVal = targetCodeCol !== -1 && row[targetCodeCol] !== null && row[targetCodeCol] !== undefined ? String(row[targetCodeCol]).trim() : '';

        if (!nameVal) continue;

        const parsedQty = parseFloat(String(qtyVal).trim());
        if (isNaN(parsedQty) || parsedQty <= 0) continue;

        // Identify stock code
        let matchedCode = '';
        if (codeVal && STOCK_ITEMS_MAP[codeVal.toLowerCase()]) {
          matchedCode = codeVal.toLowerCase();
        } else if (STOCK_ITEMS_MAP[nameVal.toLowerCase()]) {
          matchedCode = nameVal.toLowerCase();
        }

        if (!matchedCode) {
          // Direct Thai name match
          const cleanName = nameVal.toLowerCase();
          const matched = STOCK_ITEMS_LIST.find(item => 
            item.nameThai.trim().toLowerCase() === cleanName ||
            (item.code === 'french_fries' && (cleanName.includes('เฟรนฟราย') || cleanName.includes('เฟรนช์ฟราย'))) ||
            (item.code === 'banana_samosa' && cleanName.includes('ซาโมซ่า')) ||
            (item.code === 'beef_salami' && cleanName.includes('ซาลามี่')) ||
            (item.code === 'beef_steak' && cleanName.includes('สเต็กเนื้อ')) ||
            (item.code === 'pork_chop' && (cleanName.includes('พ็อกชอป') || cleanName.includes('พอร์คชอป') || cleanName.includes('พอร์คช็อป'))) ||
            (item.code === 'seafood_set' && (cleanName === 'ทะเล' || cleanName === 'ชุดทะเล'))
          );
          if (matched) {
            matchedCode = matched.code;
          }
        }

        if (matchedCode) {
          salesMap[matchedCode] += parsedQty;
          continue;
        }

        // Apply recipe mappings based on dish name
        const matchedNameLower = nameVal.toLowerCase();

        // Check if it's a known recipe dish
        const isKnownDish = 
          matchedNameLower.includes('พิซซ่า') || 
          matchedNameLower.includes('pizza') ||
          matchedNameLower.includes('สปาเก็ตตี้') ||
          matchedNameLower.includes('spaghetti') ||
          matchedNameLower.includes('เฟรนช์ฟราย') ||
          matchedNameLower.includes('เฟรนฟราย') ||
          matchedNameLower.includes('fries') ||
          matchedNameLower.includes('พ็อกชอป') ||
          matchedNameLower.includes('chop') ||
          matchedNameLower.includes('สเต็กเนื้อ') ||
          matchedNameLower.includes('steak') ||
          matchedNameLower.includes('ซาโมซ่า') ||
          matchedNameLower.includes('samosa') ||
          matchedNameLower.includes('เบคอน') ||
          matchedNameLower.includes('bacon') ||
          matchedNameLower.includes('แฮม') ||
          matchedNameLower.includes('ham') ||
          matchedNameLower.includes('ซาลามี่') ||
          matchedNameLower.includes('salami') ||
          matchedNameLower.includes('แซลมอน') ||
          matchedNameLower.includes('salmon') ||
          matchedNameLower.includes('เห็ด') ||
          matchedNameLower.includes('mushroom') ||
          matchedNameLower.includes('กุ้ง') ||
          matchedNameLower.includes('หมึก') ||
          matchedNameLower.includes('หอย') ||
          matchedNameLower.includes('ทะเล') ||
          matchedNameLower.includes('seafood');

        if (!isKnownDish) continue;

        // Apply recipes
        const isPizza = matchedNameLower.includes('พิซซ่า') || matchedNameLower.includes('pizza');
        const isDouble = matchedNameLower.includes('ดับเบิ้ล') || matchedNameLower.includes('double');

        if (isPizza) {
          const cheeseCut = isDouble ? 1.5 : 1.0;
          salesMap['cheese'] += (cheeseCut * parsedQty);
        }

        const isSeafoodPizza = isPizza && (matchedNameLower.includes('ซีฟู้ด') || matchedNameLower.includes('ซีฟู๊ด') || matchedNameLower.includes('seafood'));
        const isSpaghettiBlack = matchedNameLower.includes('หมึกดำ') || matchedNameLower.includes('spaghetti black') || (matchedNameLower.includes('พริกกระเทียม') && matchedNameLower.includes('สปาเก็ตตี้'));
        
        if (isSeafoodPizza || isSpaghettiBlack) {
          salesMap['seafood_set'] += (1.0 * parsedQty);
        }

        const isSalmonPizza = isPizza && (matchedNameLower.includes('แซลมอน') || matchedNameLower.includes('salmon'));
        if (isSalmonPizza) {
          salesMap['salmon'] += (1.0 * parsedQty);
        }

        if (matchedNameLower.includes('ซาลามี่') || matchedNameLower.includes('salami')) {
          salesMap['beef_salami'] += parsedQty;
        }
        if (matchedNameLower.includes('พามาแฮม') || matchedNameLower.includes('pama ham') || matchedNameLower.includes('parma')) {
          salesMap['parma_ham'] += parsedQty;
        }
        if (matchedNameLower.includes('เห็ดแชมปิญอง') || matchedNameLower.includes('truffle') || matchedNameLower.includes('ทรัฟเฟิล') || matchedNameLower.includes('champignon')) {
          salesMap['champignon_mushroom'] += parsedQty;
        }
        if (matchedNameLower.includes('แฮม') || matchedNameLower.includes('ham')) {
          if (!matchedNameLower.includes('พามา') && !matchedNameLower.includes('pama') && !matchedNameLower.includes('parma')) {
            salesMap['ham'] += parsedQty;
          }
        }
        if (matchedNameLower.includes('เบคอน') || matchedNameLower.includes('bacon')) {
          salesMap['bacon'] += parsedQty;
        }
        if (matchedNameLower.includes('เฟรนช์ฟราย') || matchedNameLower.includes('เฟรนฟราย') || matchedNameLower.includes('french fries') || matchedNameLower.includes('fries')) {
          salesMap['french_fries'] += parsedQty;
        }
        if (matchedNameLower.includes('พ็อกชอป') || matchedNameLower.includes('pork chop') || matchedNameLower.includes('porkchop')) {
          salesMap['pork_chop'] += parsedQty;
        }
        if (matchedNameLower.includes('สเต็กเนื้อ') || matchedNameLower.includes('tenderloin') || matchedNameLower.includes('steak')) {
          salesMap['beef_steak'] += parsedQty;
        }
        if (matchedNameLower.includes('ซาโมซ่า') || matchedNameLower.includes('samosa')) {
          salesMap['banana_samosa'] += parsedQty;
        }
        if (matchedNameLower.includes('ทูน่า') || matchedNameLower.includes('tuna')) {
          salesMap['tuna'] += parsedQty;
        }
        if (matchedNameLower.includes('หมูบด') || matchedNameLower.includes('minced pork')) {
          salesMap['minced_pork'] += parsedQty;
        }
        if (matchedNameLower.includes('กุ้ง') || matchedNameLower.includes('shrimp') || matchedNameLower.includes('prawn')) {
          salesMap['shrimp'] += parsedQty;
        }
        if (matchedNameLower.includes('หมึก') || matchedNameLower.includes('squid') || matchedNameLower.includes('octopus')) {
          salesMap['squid'] += parsedQty;
        }
        if (matchedNameLower.includes('หอย') || matchedNameLower.includes('clam') || matchedNameLower.includes('mussel')) {
          salesMap['clam'] += parsedQty;
        }
      }
    }
    return salesMap;
  }

  // Try to find header column indices to prevent selecting other numeric columns (like physical counts or theory remaining)
  let nameColIdx = -1;
  let qtyColIdx = -1;
  let codeColIdx = -1;

  for (let r = 0; r < Math.min(15, rows.length); r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    
    const cells = row.map(c => c !== null && c !== undefined ? String(c).trim().toLowerCase() : '');
    
    // Look for product/item identifier column and quantity column
    const hasProductHeader = cells.some(c => 
      c.includes('สินค้า') || c.includes('วัตถุดิบ') || c.includes('รายการ') || 
      c === 'item' || c === 'product' || c === 'name' || c.includes('ดิบ')
    );
    
    if (hasProductHeader) {
      nameColIdx = cells.findIndex(c => 
        c.includes('สินค้า') || c.includes('วัตถุดิบ') || c.includes('รายการ') || 
        c === 'item' || c === 'product' || c === 'name' || c.includes('ดิบ')
      );
      
      // Look for sales quantity column specifically (prioritize sales-related names over physical count or diff columns)
      qtyColIdx = cells.findIndex(c => 
        c.includes('จำนวนที่ขาย') || c.includes('ยอดขายตามระบบ') || 
        c.includes('ยอดขาย') || c.includes('จำนวนขาย') || c.includes('ขาย') || 
        (c.includes('จำนวน') && !c.includes('เหลือ') && !c.includes('นับ') && !c.includes('จริง') && !c.includes('เติม')) || 
        c === 'qty' || c === 'quantity' || c === 'sold' || c === 'sales'
      );
      
      codeColIdx = cells.findIndex(c => c.includes('รหัส') || c === 'code' || c === 'id');
      
      if (qtyColIdx !== -1 && nameColIdx !== -1) {
        // Headers identified successfully!
        break;
      }
    }
  }

  // If columns are identified, parse using specific column indices
  if (nameColIdx !== -1 && qtyColIdx !== -1) {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const nameVal = row[nameColIdx] !== null && row[nameColIdx] !== undefined ? String(row[nameColIdx]).trim() : '';
      const qtyVal = row[qtyColIdx];
      const codeVal = codeColIdx !== -1 && row[codeColIdx] !== null && row[codeColIdx] !== undefined ? String(row[codeColIdx]).trim() : '';

      if (!nameVal) continue;
      
      const parsedQty = parseFloat(String(qtyVal).trim());
      if (isNaN(parsedQty) || parsedQty <= 0) continue;

      // Skip header rows if accidentally parsed in data loop
      if (
        nameVal.includes('สินค้า') || nameVal.includes('วัตถุดิบ') || nameVal.includes('รายการ') ||
        nameVal === 'item' || nameVal === 'product' || nameVal === 'name' ||
        nameVal.includes('รวม') || nameVal.includes('total') || nameVal.includes('sum')
      ) {
        continue;
      }

      // Identify stock code
      let matchedCode = '';
      if (codeVal && STOCK_ITEMS_MAP[codeVal.toLowerCase()]) {
        matchedCode = codeVal.toLowerCase();
      } else if (STOCK_ITEMS_MAP[nameVal.toLowerCase()]) {
        matchedCode = nameVal.toLowerCase();
      }

      if (!matchedCode) {
        // Direct Thai name match
        const cleanName = nameVal.toLowerCase();
        const matched = STOCK_ITEMS_LIST.find(item => 
          item.nameThai.trim().toLowerCase() === cleanName ||
          (item.code === 'french_fries' && (cleanName.includes('เฟรนฟราย') || cleanName.includes('เฟรนช์ฟราย'))) ||
          (item.code === 'banana_samosa' && cleanName.includes('ซาโมซ่า')) ||
          (item.code === 'beef_salami' && cleanName.includes('ซาลามี่')) ||
          (item.code === 'beef_steak' && cleanName.includes('สเต็กเนื้อ')) ||
          (item.code === 'pork_chop' && (cleanName.includes('พ็อกชอป') || cleanName.includes('พอร์คชอป') || cleanName.includes('พอร์คช็อป'))) ||
          (item.code === 'seafood_set' && (cleanName === 'ทะเล' || cleanName === 'ชุดทะเล'))
        );
        if (matched) {
          matchedCode = matched.code;
        }
      }

      if (matchedCode) {
        salesMap[matchedCode] += parsedQty;
        continue;
      }

      // Apply recipe mappings based on dish name
      const matchedNameLower = nameVal.toLowerCase();

      // Check if it's a known recipe dish
      const isKnownDish = 
        matchedNameLower.includes('พิซซ่า') || 
        matchedNameLower.includes('pizza') ||
        matchedNameLower.includes('สปาเก็ตตี้') ||
        matchedNameLower.includes('spaghetti') ||
        matchedNameLower.includes('เฟรนช์ฟราย') ||
        matchedNameLower.includes('เฟรนฟราย') ||
        matchedNameLower.includes('fries') ||
        matchedNameLower.includes('พ็อกชอป') ||
        matchedNameLower.includes('chop') ||
        matchedNameLower.includes('สเต็กเนื้อ') ||
        matchedNameLower.includes('steak') ||
        matchedNameLower.includes('ซาโมซ่า') ||
        matchedNameLower.includes('samosa') ||
        matchedNameLower.includes('เบคอน') ||
        matchedNameLower.includes('bacon') ||
        matchedNameLower.includes('แฮม') ||
        matchedNameLower.includes('ham') ||
        matchedNameLower.includes('ซาลามี่') ||
        matchedNameLower.includes('salami') ||
        matchedNameLower.includes('แซลมอน') ||
        matchedNameLower.includes('salmon') ||
        matchedNameLower.includes('เห็ด') ||
        matchedNameLower.includes('mushroom') ||
        matchedNameLower.includes('กุ้ง') ||
        matchedNameLower.includes('หมึก') ||
        matchedNameLower.includes('หอย') ||
        matchedNameLower.includes('ทะเล') ||
        matchedNameLower.includes('seafood');

      if (!isKnownDish) continue;

      // Apply recipes
      const isPizza = matchedNameLower.includes('พิซซ่า') || matchedNameLower.includes('pizza');
      const isDouble = matchedNameLower.includes('ดับเบิ้ล') || matchedNameLower.includes('double');

      if (isPizza) {
        const cheeseCut = isDouble ? 1.5 : 1.0;
        salesMap['cheese'] += (cheeseCut * parsedQty);
      }

      const isSeafoodPizza = isPizza && (matchedNameLower.includes('ซีฟู้ด') || matchedNameLower.includes('ซีฟู๊ด') || matchedNameLower.includes('seafood'));
      const isSpaghettiBlack = matchedNameLower.includes('หมึกดำ') || matchedNameLower.includes('spaghetti black') || (matchedNameLower.includes('พริกกระเทียม') && matchedNameLower.includes('สปาเก็ตตี้'));
      
      if (isSeafoodPizza || isSpaghettiBlack) {
        salesMap['seafood_set'] += (1.0 * parsedQty);
      }

      const isSalmonPizza = isPizza && (matchedNameLower.includes('แซลมอน') || matchedNameLower.includes('salmon'));
      if (isSalmonPizza) {
        salesMap['salmon'] += (1.0 * parsedQty);
      }

      if (matchedNameLower.includes('ซาลามี่') || matchedNameLower.includes('salami')) {
        salesMap['beef_salami'] += parsedQty;
      }
      if (matchedNameLower.includes('พามาแฮม') || matchedNameLower.includes('pama ham') || matchedNameLower.includes('parma')) {
        salesMap['parma_ham'] += parsedQty;
      }
      if (matchedNameLower.includes('เห็ดแชมปิญอง') || matchedNameLower.includes('truffle') || matchedNameLower.includes('ทรัฟเฟิล') || matchedNameLower.includes('champignon')) {
        salesMap['champignon_mushroom'] += parsedQty;
      }
      if (matchedNameLower.includes('แฮม') || matchedNameLower.includes('ham')) {
        if (!matchedNameLower.includes('พามา') && !matchedNameLower.includes('pama') && !matchedNameLower.includes('parma')) {
          salesMap['ham'] += parsedQty;
        }
      }
      if (matchedNameLower.includes('เบคอน') || matchedNameLower.includes('bacon')) {
        salesMap['bacon'] += parsedQty;
      }
      if (matchedNameLower.includes('เฟรนช์ฟราย') || matchedNameLower.includes('เฟรนฟราย') || matchedNameLower.includes('french fries') || matchedNameLower.includes('fries')) {
        salesMap['french_fries'] += parsedQty;
      }
      if (matchedNameLower.includes('พ็อกชอป') || matchedNameLower.includes('pork chop') || matchedNameLower.includes('porkchop')) {
        salesMap['pork_chop'] += parsedQty;
      }
      if (matchedNameLower.includes('สเต็กเนื้อ') || matchedNameLower.includes('tenderloin') || matchedNameLower.includes('steak')) {
        salesMap['beef_steak'] += parsedQty;
      }
      if (matchedNameLower.includes('ซาโมซ่า') || matchedNameLower.includes('samosa')) {
        salesMap['banana_samosa'] += parsedQty;
      }
      if (matchedNameLower.includes('ทูน่า') || matchedNameLower.includes('tuna')) {
        salesMap['tuna'] += parsedQty;
      }
      if (matchedNameLower.includes('หมูบด') || matchedNameLower.includes('minced pork')) {
        salesMap['minced_pork'] += parsedQty;
      }
      if (matchedNameLower.includes('กุ้ง') || matchedNameLower.includes('shrimp') || matchedNameLower.includes('prawn')) {
        salesMap['shrimp'] += parsedQty;
      }
      if (matchedNameLower.includes('หมึก') || matchedNameLower.includes('squid') || matchedNameLower.includes('octopus')) {
        salesMap['squid'] += parsedQty;
      }
      if (matchedNameLower.includes('หอย') || matchedNameLower.includes('clam') || matchedNameLower.includes('mussel')) {
        salesMap['clam'] += parsedQty;
      }
    }
    return salesMap;
  }

  // Fallback to older row-scanning logic if no clear columns are found
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

    // A.1. DIRECT THAI NAME MATCH (If the user uploads a direct list of Thai ingredient names like "ชีส", "แฮม")
    let matchedThaiCode = '';
    for (const cellStr of cellStrings) {
      const cleanCell = cellStr.trim().toLowerCase();
      if (!cleanCell) continue;

      const matched = STOCK_ITEMS_LIST.find(item => 
        item.nameThai.trim().toLowerCase() === cleanCell ||
        (item.code === 'french_fries' && (cleanCell.includes('เฟรนฟราย') || cleanCell.includes('เฟรนช์ฟราย'))) ||
        (item.code === 'banana_samosa' && cleanCell.includes('ซาโมซ่า')) ||
        (item.code === 'beef_salami' && cleanCell.includes('ซาลามี่')) ||
        (item.code === 'beef_steak' && cleanCell.includes('สเต็กเนื้อ')) ||
        (item.code === 'pork_chop' && (cleanCell.includes('พ็อกชอป') || cleanCell.includes('พอร์คชอป') || cleanCell.includes('พอร์คช็อป'))) ||
        (item.code === 'seafood_set' && (cleanCell === 'ทะเล' || cleanCell === 'ชุดทะเล'))
      );
      if (matched) {
        matchedThaiCode = matched.code;
        break;
      }
    }

    if (matchedThaiCode) {
      salesMap[matchedThaiCode] += bestQty;
      continue;
    }

    // A.2. DIRECT ENGLISH CODE MATCH (Template Mode)
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
