import express from 'express';
import path from 'path';
import multer from 'multer';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';
import { 
  initDatabase, 
  getInventory, 
  replenishStock, 
  recordDailyCount, 
  recordSales, 
  getReplenishments, 
  getDailyCounts, 
  getSales, 
  getDiscrepancyReport,
  getIsUsingMySQL,
  saveDiscrepancyRemark
} from './server-db.js';
import { generateExcelReport, generateSampleSalesTemplate, parseSalesExcel } from './server-excel.js';
import { STOCK_ITEMS_LIST, STOCK_ITEMS_MAP, ChatMessage } from './src/types.js';

// Setup environment variables
import dotenv from 'dotenv';
dotenv.config();

const app = express();
const PORT = Number(process.env.PORT) || 3000;

// Body parsers with limits for excel uploads
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Multer setup for Excel uploads in memory
const upload = multer({ storage: multer.memoryStorage() });

// Lazy-initialize Gemini SDK
let aiClient: GoogleGenAI | null = null;
function getGemini(): GoogleGenAI {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error('GEMINI_API_KEY is not defined. Please add it in Settings > Secrets.');
    }
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        },
      },
    });
  }
  return aiClient;
}

// REST API ENDPOINTS

// 1. Connection status
app.get('/api/db-status', async (req, res) => {
  try {
    const isMySQL = await getIsUsingMySQL();
    res.json({
      status: 'ok',
      isUsingMySQL: isMySQL,
      databaseType: isMySQL ? 'MySQL Database' : 'Local JSON Storage (Fallback)',
      lineWebhookUrl: process.env.APP_URL ? `${process.env.APP_URL}/api/line-webhook` : 'Not configured'
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 2. Current Stock Levels
app.get('/api/inventory', async (req, res) => {
  try {
    const items = await getInventory();
    res.json(items);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 3. Record Replenishment
app.post('/api/replenish', async (req, res) => {
  try {
    const { itemCode, qty } = req.body;
    const amount = parseFloat(qty);
    if (!itemCode || isNaN(amount) || amount <= 0) {
      return res.status(400).json({ error: 'รหัสสินค้าหรือจำนวนไม่ถูกต้อง' });
    }
    const success = await replenishStock(itemCode, amount);
    if (success) {
      res.json({ success: true, message: `เพิ่มสต๊อกเรียบร้อยแล้ว` });
    } else {
      res.status(404).json({ error: 'ไม่พบสินค้าที่ระบุ' });
    }
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 4. Record Physical Count (คงเหลือรายวัน)
app.post('/api/daily-count', async (req, res) => {
  try {
    const { itemCode, qty, date } = req.body;
    const amount = parseFloat(qty);
    const targetDate = date || new Date().toISOString().split('T')[0];
    
    if (!itemCode || isNaN(amount) || amount < 0) {
      return res.status(400).json({ error: 'ข้อมูลไม่ถูกต้อง' });
    }
    
    const success = await recordDailyCount(itemCode, amount, targetDate);
    if (success) {
      res.json({ success: true, message: `บันทึกคงเหลือประจำวันที่ ${targetDate} เรียบร้อยแล้ว` });
    } else {
      res.status(404).json({ error: 'ไม่พบสินค้าที่ระบุ' });
    }
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 5. Upload Sales Excel directly from Web Dashboard
app.post('/api/excel/upload-sales', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ Excel' });
    }
    const targetDate = req.body.date || new Date().toISOString().split('T')[0];
    const salesMap = parseSalesExcel(req.file.buffer);
    
    // Save to DB
    const success = await recordSales(targetDate, salesMap);
    if (success) {
      res.json({ 
        success: true, 
        message: `อัปโหลดและประมวลผลยอดขายสำหรับวันที่ ${targetDate} สำเร็จ! หักสต๊อกตามยอดขายเรียบร้อยแล้ว`,
        salesMap 
      });
    } else {
      res.status(500).json({ error: 'เกิดข้อผิดพลาดในการบันทึกข้อมูลยอดขาย' });
    }
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 6. Get Discrepancy Reconciliation Report
app.get('/api/discrepancy', async (req, res) => {
  try {
    const date = (req.query.date as string) || new Date().toISOString().split('T')[0];
    const report = await getDiscrepancyReport(date);
    res.json(report);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Save discrepancy remark
app.post('/api/discrepancy/remark', async (req, res) => {
  try {
    const { itemCode, date, remark } = req.body;
    if (!itemCode || !date) {
      return res.status(400).json({ error: 'itemCode and date are required' });
    }
    await saveDiscrepancyRemark(itemCode, date, remark || '');
    res.json({ success: true, message: 'บันทึกหมายเหตุเรียบร้อยแล้ว' });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 7. Get All Logs
app.get('/api/logs', async (req, res) => {
  try {
    const [replenishments, dailyCounts, sales] = await Promise.all([
      getReplenishments(),
      getDailyCounts(),
      getSales()
    ]);
    res.json({ replenishments, dailyCounts, sales });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// 8. Download Generated Excel Report
app.get('/api/excel/download-report', async (req, res) => {
  try {
    const date = (req.query.date as string) || new Date().toISOString().split('T')[0];
    const [inventory, replenishments, dailyCounts, sales, discrepancy] = await Promise.all([
      getInventory(),
      getReplenishments(),
      getDailyCounts(),
      getSales(),
      getDiscrepancyReport(date)
    ]);

    const buffer = generateExcelReport(inventory, replenishments, dailyCounts, sales, discrepancy, date);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Stock_Report_${date}.xlsx`);
    res.send(buffer);
  } catch (err: any) {
    res.status(500).send(`Error generating Excel file: ${err.message}`);
  }
});

// 9. Download Sales template excel
app.get('/api/excel/sales-template', (req, res) => {
  try {
    const buffer = generateSampleSalesTemplate();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Sales_Template_Daily.xlsx');
    res.send(buffer);
  } catch (err: any) {
    res.status(500).send(`Error generating template: ${err.message}`);
  }
});

// 10. Smart AI Insights from Gemini API
app.get('/api/ai-insights', async (req, res) => {
  try {
    const date = (req.query.date as string) || new Date().toISOString().split('T')[0];
    const discrepancy = await getDiscrepancyReport(date);
    
    // Check if key is configured
    if (!process.env.GEMINI_API_KEY) {
      return res.json({ 
        insights: "กรุณากำหนด GEMINI_API_KEY ในแถบ Secrets ด้านขวาเพื่อใช้งาน AI วิเคราะห์ความต่างของสต๊อกวัตถุดิบรายวัน" 
      });
    }

    const ai = getGemini();
    const prompt = `
      คุณคือผู้ช่วยบริหารจัดการสต๊อกครัวร้านอาหารระดับมืออาชีพ
      นี่คือข้อมูลรายงานความคลาดเคลื่อนสต๊อกของวัตถุดิบประจำวันที่ ${date}:
      ${JSON.stringify(discrepancy.map(d => ({
        name: d.nameThai,
        yesterday: d.yesterdayCount,
        added: d.replenished,
        sold: d.sold,
        expected: d.expectedRemaining,
        actual: d.actualRemaining,
        lost: d.difference
      })), null, 2)}

      ในข้อมูล lost:
      - เป็นบวก (+) แปลว่า วัตถุดิบจริงนับแล้วหายไปน้อยกว่าทฤษฎี (เช่น สูญเสียจากขยะ ตักขนาดจานใหญ่เกินไป หรือลืมบันทึกยอดขาย)
      - เป็นลบ (-) แปลว่า วัตถุดิบนับจริงเหลือเยอะกว่าในระบบ (เช่น เติมของแล้วไม่ลงบันทึก หรือยอดขายรายงานคลาดเคลื่อน)

      กรุณาวิเคราะห์สั้นๆ กระชับและระบุปัญหาหลัก 3 รายการเป็นภาษาไทย โดยให้คำแนะนำที่เป็นรูปธรรมในการดูแลวัตถุดิบและป้องกันสต๊อกหาย
    `;

    const response = await ai.models.generateContent({
      model: 'gemini-3.5-flash',
      contents: prompt,
    });

    res.json({ insights: response.text });
  } catch (err: any) {
    res.json({ insights: `วิเคราะห์โดย AI ขัดข้องชั่วคราว: ${err.message}` });
  }
});


// -----------------------------------------------------------------------------
// LINE Messaging API Bot Webhook & Chat Simulator Core
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
// LINE Messaging API Bot Webhook & Chat Simulator Core
// -----------------------------------------------------------------------------

// Helper to extract dates from user commands to support retrospective (backdated) entries
function extractDateFromText(text: string, defaultDate: string): { date: string; cleanText: string } {
  if (!text) return { date: defaultDate, cleanText: '' };
  let date = defaultDate;
  let remainingText = text;

  // 1. Match YYYY-MM-DD
  const yyyymmddMatch = text.match(/\b(20\d{2})[-/](0?[1-9]|1[0-2])[-/](0?[1-9]|[12]\d|3[01])\b/);
  if (yyyymmddMatch) {
    const year = yyyymmddMatch[1];
    const month = yyyymmddMatch[2].padStart(2, '0');
    const day = yyyymmddMatch[3].padStart(2, '0');
    date = `${year}-${month}-${day}`;
    remainingText = text.replace(yyyymmddMatch[0], '').replace(/วันที่/g, '').trim();
    return { date, cleanText: remainingText };
  }

  // 2. Match DD-MM-YYYY or DD/MM/YYYY
  const ddmmyyyyMatch = text.match(/\b(0?[1-9]|[12]\d|3[01])[-/](0?[1-9]|1[0-2])[-/](20\d{2})\b/);
  if (ddmmyyyyMatch) {
    const day = ddmmyyyyMatch[1].padStart(2, '0');
    const month = ddmmyyyyMatch[2].padStart(2, '0');
    const year = ddmmyyyyMatch[3];
    date = `${year}-${month}-${day}`;
    remainingText = text.replace(ddmmyyyyMatch[0], '').replace(/วันที่/g, '').trim();
    return { date, cleanText: remainingText };
  }

  // 3. Match "วันที่ [1-31]"
  const thDayMatch = text.match(/วันที่\s*([1-9]|[12]\d|3[01])\b/);
  if (thDayMatch) {
    const dayVal = thDayMatch[1].padStart(2, '0');
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    date = `${year}-${month}-${dayVal}`;
    remainingText = text.replace(thDayMatch[0], '').trim();
    return { date, cleanText: remainingText };
  }

  return { date, cleanText: remainingText };
}

// User state to remember selection context for interactive chat flow
interface UserState {
  pendingAction: 'replenish' | 'count' | 'note_add' | 'note_sub';
  itemCode: string;
  itemName: string;
  date: string;
}

const userStates = new Map<string, UserState>();
const userWorkingDates = new Map<string, string>(); // User persistent working dates (userId -> YYYY-MM-DD)

function getThaiDateLabel(daysAgo: number): { label: string; dateStr: string } {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);
  const dateStr = date.toISOString().split('T')[0];
  const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const day = date.getDate();
  const month = thaiMonths[date.getMonth()];
  
  let label = '';
  if (daysAgo === 0) {
    label = `วันนี้ (${day} ${month})`;
  } else if (daysAgo === 1) {
    label = `เมื่อวาน (${day} ${month})`;
  } else {
    label = `${day} ${month}`;
  }
  return { label, dateStr };
}

// Helper to handle bot commands dynamically (used by both REAL Webhook and Simulator!)
async function processBotMessage(messageText: string, fileBuffer?: Buffer, fileName?: string, userId: string = 'default'): Promise<any> {
  const todayStr = new Date().toISOString().split('T')[0];
  const activeWorkingDate = userWorkingDates.get(userId) || todayStr;
  
  const trimmedMsg = messageText.trim();

  // Handle "วันนี้" / "today" to reset working date back to the current day
  if (trimmedMsg === 'วันนี้' || trimmedMsg.toLowerCase() === 'today' || trimmedMsg === 'วันปัจจุบัน' || trimmedMsg === 'ปัจจุบัน') {
    userWorkingDates.delete(userId);
    return {
      type: 'text',
      text: `📅 รีเซ็ตวันทำงานของบอทกลับสู่: "วันปัจจุบัน" (${todayStr}) เรียบร้อยแล้วค่ะ!\n\nหลังจากนี้ รายการเติมของ, บันทึกตรวจนับ, หรือไฟล์ Excel ยอดขายที่คุณส่งมา จะถูกลงบันทึกเป็นวันปัจจุบันค่ะ`
    };
  }

  // Handle setting working date manually if they type ONLY a date (e.g., "วันที่ 2026-06-25" or "2026-06-25")
  const rawDateCheck = extractDateFromText(trimmedMsg, '');
  if (rawDateCheck.date && (rawDateCheck.cleanText === '' || rawDateCheck.cleanText === 'ตั้ง' || rawDateCheck.cleanText === 'เลือก')) {
    userWorkingDates.set(userId, rawDateCheck.date);
    return {
      type: 'text',
      text: `📅 ล็อควันทำงานย้อนหลังสำเร็จ!\n\nระบบล็อควันทำรายการเป็นวันที่: ${rawDateCheck.date} 🔒\n\nหลังจากนี้ ทุกปุ่มการใช้งาน, รายการหักสต๊อก, หรือไฟล์ Excel ที่อัปโหลดทาง LINE จะลงข้อมูลในวันที่นี้โดยอัตโนมัติค่ะ\n\n(หากต้องการล้างค่านี้กลับมาใช้วันปัจจุบัน พิมพ์ตอบกลับมาว่า "วันนี้")`
    };
  }

  const { date: targetDate, cleanText: extractedCleanText } = extractDateFromText(messageText, activeWorkingDate);
  
  let cleanText = extractedCleanText.trim();
  if (!cleanText && messageText.trim()) {
    cleanText = messageText.trim();
  }

  // 1.5. Intercept pure numbers or numbers with trailing units if user has a pending action
  const numberOnlyRegex = /^\s*(\d+(\.\d+)?)\s*(ยูนิต|ถุง|ฟอง|กิโล|กรัม|ชิ้น|กล่อง|แพ็ค|แพค|อัน|ขวด|ซอง|กระป๋อง|g|kg|pcs|units|unit)?\s*$/i;
  const numMatch = cleanText.match(numberOnlyRegex);
  if (numMatch) {
    const qty = parseFloat(numMatch[1]);
    const state = userStates.get(userId);
    if (state) {
      const itemCode = state.itemCode;
      const itemName = state.itemName;
      const action = state.pendingAction;
      const actionDate = state.date || targetDate;

      if (action === 'replenish' || action === 'count') {
        // Clear the pending state once successfully processed
        userStates.delete(userId);

        if (action === 'replenish') {
          await replenishStock(itemCode, qty, actionDate);
          const updatedInv = await getInventory();
          const current = updatedInv.find(i => i.code === itemCode)?.currentQty || 0;
          const matchedItem = STOCK_ITEMS_LIST.find(i => i.code === itemCode);
          const unit = matchedItem ? matchedItem.unit : 'ยูนิต';

          return {
            type: 'text',
            text: `✅ เติมสต๊อกสำเร็จสำหรับวันที่ ${actionDate}!\nวัตถุดิบ: ${itemName}\nจำนวนที่เพิ่ม: +${qty} ${unit}\nยอดคงเหลือรวมล่าสุด: ${current} ${unit}`
          };
        } else if (action === 'count') {
          await recordDailyCount(itemCode, qty, actionDate);
          const matchedItem = STOCK_ITEMS_LIST.find(i => i.code === itemCode);
          const unit = matchedItem ? matchedItem.unit : 'ยูนิต';

          return {
            type: 'text',
            text: `✅ บันทึกยอดคงเหลือจริงสำเร็จสำหรับวันที่ ${actionDate}!\nวัตถุดิบ: ${itemName}\nจำนวนที่นับได้จริง: ${qty} ${unit}\nระบบนำไปบันทึกเปรียบเทียบในประวัติสต๊อกเรียบร้อยแล้วค่ะ`
          };
        }
      }
    }
  }

  // 1.6. Intercept pizza name if user has a pending note action (like note_add or note_sub)
  const pendingState = userStates.get(userId);
  if (pendingState && (pendingState.pendingAction === 'note_add' || pendingState.pendingAction === 'note_sub')) {
    const pizzaName = cleanText;
    const isAdd = pendingState.pendingAction === 'note_add';
    const cheeseAdjust = 0.25;
    const actionDate = pendingState.date || targetDate;

    // Clear the pending state once successfully processed
    userStates.delete(userId);

    if (isAdd) {
      // Increase cheese count
      await replenishStock('cheese', cheeseAdjust, actionDate);
    } else {
      // Decrease cheese count (recorded as sales of cheese)
      const salesMap: Record<string, number> = {};
      STOCK_ITEMS_LIST.forEach(item => {
        salesMap[item.code] = 0;
      });
      salesMap['cheese'] = cheeseAdjust;
      await recordSales(actionDate, salesMap);
    }

    const updatedInv = await getInventory();
    const current = updatedInv.find(i => i.code === 'cheese')?.currentQty || 0;

    return {
      type: 'text',
      text: `📝 บันทึกหมายเหตุ "${pizzaName}" เรียบร้อยแล้วสำหรับวันที่ ${actionDate}!\nวัตถุดิบ: ชีส (Cheese)\nการปรับสต๊อก: ${isAdd ? 'เพิ่มชีส (+)' : 'ลดชีส (-)'} ${cheeseAdjust} ยูนิต\nยอดคงเหลือชีสล่าสุด: ${current} ยูนิต`
    };
  }

  // 1. Handle File Upload (Sales Excel uploaded via LINE)
  if (fileBuffer) {
    try {
      const salesMap = parseSalesExcel(fileBuffer);
      await recordSales(targetDate, salesMap);
      
      let itemSummary = '';
      let totalItems = 0;
      Object.entries(salesMap).forEach(([code, qty]) => {
        if (qty > 0) {
          totalItems++;
          itemSummary += `\n- ${STOCK_ITEMS_MAP[code]?.nameThai}: ${qty} ${STOCK_ITEMS_MAP[code]?.unit}`;
        }
      });

      return {
        type: 'text',
        text: `📊 อัปโหลดไฟล์ "${fileName || 'sales.xlsx'}" สำเร็จสำหรับวันที่ ${targetDate}!\nระบบได้นำยอดขายบันทึกลงฐานข้อมูลและทำการหักยอดตัดสต๊อกเรียบร้อยแล้วค่ะ (${totalItems} รายการ):${itemSummary}\n\nคุณสามารถพิมพ์ "รายงาน ${targetDate}" เพื่อดูเปรียบเทียบยอดได้ทันทีค่ะ`
      };
    } catch (err: any) {
      return {
        type: 'text',
        text: `❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: ${err.message}`
      };
    }
  }

  // 2. Command: "เติมสต๊อก" (Show replenishment instructions / options)
  if (cleanText === 'เติมสต๊อก') {
    return {
      type: 'flex',
      text: '📋 รายการเติมสต๊อกวัตถุดิบ',
      flexContent: {
        title: `📋 เมนูเติมสต๊อก (${targetDate})`,
        description: `เลือกวัตถุดิบด้านล่างเพื่อเติมสต๊อกสำหรับวันที่ ${targetDate}\n(หรือพิมพ์: เติม [ชื่อสินค้า] [จำนวน])`,
        items: STOCK_ITEMS_LIST.map(item => ({
          code: item.code,
          name: item.nameThai,
          unit: item.unit,
          actionText: `เติม ${item.nameThai} วันที่ ${targetDate}`
        }))
      }
    };
  }

  // 3. Command: "คงเหลือ" (Show physical stock remaining instructions / options)
  if (cleanText === 'คงเหลือ') {
    return {
      type: 'flex',
      text: '📊 บันทึกสต๊อกคงเหลือประจำวัน',
      flexContent: {
        title: `📊 บันทึกคงเหลือรายวัน (${targetDate})`,
        description: `เลือกวัตถุดิบเพื่อบันทึกของที่นับได้จริงประจำวันที่ ${targetDate}\n(หรือพิมพ์: คงเหลือ [ชื่อสินค้า] [จำนวน])`,
        items: STOCK_ITEMS_LIST.map(item => ({
          code: item.code,
          name: item.nameThai,
          unit: item.unit,
          actionText: `คงเหลือ ${item.nameThai} วันที่ ${targetDate}`
        }))
      }
    };
  }

  // 3.5. Command: "เลือกวันที่" / "วันที่" (Show Date-Selection menu)
  if (cleanText === 'เลือกวันที่' || cleanText === 'ตั้งวันที่' || cleanText === 'เปลี่ยนวันที่' || cleanText === 'ย้อนหลัง' || cleanText === 'วันที่') {
    const d0 = getThaiDateLabel(0);
    const d1 = getThaiDateLabel(1);
    const d2 = getThaiDateLabel(2);
    const d3 = getThaiDateLabel(3);
    const d4 = getThaiDateLabel(4);

    return {
      type: 'flex',
      text: '📅 เลือกวันทำงานของระบบ',
      flexContent: {
        title: '📅 เลือกวันทำงาน (Working Date)',
        description: `วันทำงานของบอทขณะนี้: ${targetDate}\n\nคุณสามารถคลิกปุ่ม "เลือกจากปฏิทิน" ด้านล่างเพื่อระบุวันที่ต้องการ หรือเลือกวันย้อนหลังด่วนได้เลยค่ะ:`,
        items: [
          {
            name: '📅 เลือกจากปฏิทิน (Calendar)',
            actionType: 'datetimepicker',
            actionData: 'action=set_working_date'
          },
          {
            name: d0.label,
            actionText: `วันที่ ${d0.dateStr}`
          },
          {
            name: d1.label,
            actionText: `วันที่ ${d1.dateStr}`
          },
          {
            name: d2.label,
            actionText: `วันที่ ${d2.dateStr}`
          },
          {
            name: d3.label,
            actionText: `วันที่ ${d3.dateStr}`
          },
          {
            name: d4.label,
            actionText: `วันที่ ${d4.dateStr}`
          }
        ]
      }
    };
  }

  // 4. Command: "รายงาน" (Send summary report & comparison link)
  if (cleanText === 'รายงาน' || cleanText.startsWith('รายงาน')) {
    const discrepancy = await getDiscrepancyReport(targetDate);
    const lostItems = discrepancy.filter(d => d.difference !== 0);
    
    let summaryText = `📈 รายงานความคลาดเคลื่อนสต๊อก (${targetDate})\n`;
    if (lostItems.length === 0) {
      summaryText += `\n✅ ยอดสต๊อกคงเหลือจริงตรงกับยอดขายระบบครบถ้วนทุกรายการค่ะ! Perfect!`;
    } else {
      summaryText += `\n⚠️ พบรายการไม่ตรงกัน ${lostItems.length} วัตถุดิบ:`;
      lostItems.slice(0, 7).forEach(d => {
        const sign = d.difference > 0 ? '❌ ขาดหายไป' : '➕ เกินมา';
        summaryText += `\n- ${d.nameThai}: ${sign} ${Math.abs(d.difference)} ${STOCK_ITEMS_MAP[d.itemCode]?.unit} (ควรมี ${d.expectedRemaining} นับได้ ${d.actualRemaining})`;
      });
      if (lostItems.length > 7) {
        summaryText += `\n...และรายการอื่นๆ รวม ${lostItems.length} รายการ`;
      }
    }

    summaryText += `\n\n📋 ยอดวัตถุดิบคงเหลือจริง:`;
    discrepancy.forEach(d => {
      summaryText += `\n- ${d.nameThai}: ${d.actualRemaining} ${STOCK_ITEMS_MAP[d.itemCode]?.unit || 'ยูนิต'}`;
    });

    summaryText += `\n\n📥 ดาวน์โหลดไฟล์ Excel สำหรับสต๊อกทั้งหมดได้ที่นี่:\n${process.env.APP_URL || 'http://localhost:3000'}/api/excel/download-report?date=${targetDate}`;

    return {
      type: 'text',
      text: summaryText
    };
  }

  // 5. Pattern Parsing for: เติม [ชื่อสินค้า] [จำนวน]
  const replenishRegex = /^(เติม|add)\s+([ก-๙a-zA-Z0-9\s_]+)\s+(\d+(\.\d+)?)$/i;
  const repMatch = cleanText.match(replenishRegex);
  if (repMatch) {
    const queryName = repMatch[2].trim().toLowerCase();
    const qty = parseFloat(repMatch[3]);

    // Find closest item match
    const matchedItem = STOCK_ITEMS_LIST.find(item => 
      item.nameThai === queryName || 
      item.code.toLowerCase() === queryName ||
      item.nameThai.includes(queryName)
    );

    if (!matchedItem) {
      return {
        type: 'text',
        text: `🔍 ไม่พบสินค้าชื่อ "${queryName}" ในระบบ กรุณาเลือกพิมพ์ชื่อสินค้าที่ถูกต้อง เช่น ชีส, แซลมอน, แฮม, หมูบด`
      };
    }

    await replenishStock(matchedItem.code, qty, targetDate);
    const updatedInv = await getInventory();
    const current = updatedInv.find(i => i.code === matchedItem.code)?.currentQty || 0;

    return {
      type: 'text',
      text: `✅ เติมสต๊อกสำเร็จสำหรับวันที่ ${targetDate}!\nวัตถุดิบ: ${matchedItem.nameThai}\nจำนวนที่เพิ่ม: +${qty} ${matchedItem.unit}\nยอดคงเหลือรวมล่าสุด: ${current} ${matchedItem.unit}`
    };
  }

  // 5.5. Partial Match for: เติม [ชื่อสินค้า] (without quantity specified)
  const partialReplenishRegex = /^(เติม|add)\s+([ก-๙a-zA-Z0-9\s_]+)$/i;
  const partRepMatch = cleanText.match(partialReplenishRegex);
  if (partRepMatch) {
    const queryName = partRepMatch[2].trim().toLowerCase();
    const matchedItem = STOCK_ITEMS_LIST.find(item => 
      item.nameThai === queryName || 
      item.code.toLowerCase() === queryName ||
      item.nameThai.includes(queryName)
    );

    if (matchedItem) {
      // Save pending state
      userStates.set(userId, {
        pendingAction: 'replenish',
        itemCode: matchedItem.code,
        itemName: matchedItem.nameThai,
        date: targetDate
      });

      return {
        type: 'text',
        text: `✍️ บันทึกเลือกเติมวัตถุดิบ: "${matchedItem.nameThai}"\n\n👉 กรุณาพิมพ์ระบุเฉพาะ "จำนวน" ที่ต้องการเติมตอบกลับมาได้เลยค่ะ (เช่น "10" หรือ "15.5")`
      };
    }
  }

  // 6. Pattern Parsing for: คงเหลือ [ชื่อสินค้า] [จำนวน]
  const dailyCountRegex = /^(คงเหลือ|เหลือ|นับได้|count)\s+([ก-๙a-zA-Z0-9\s_]+)\s+(\d+(\.\d+)?)$/i;
  const countMatch = cleanText.match(dailyCountRegex);
  if (countMatch) {
    const queryName = countMatch[2].trim().toLowerCase();
    const qty = parseFloat(countMatch[3]);

    const matchedItem = STOCK_ITEMS_LIST.find(item => 
      item.nameThai === queryName || 
      item.code.toLowerCase() === queryName ||
      item.nameThai.includes(queryName)
    );

    if (!matchedItem) {
      return {
        type: 'text',
        text: `🔍 ไม่พบสินค้าชื่อ "${queryName}" ในระบบ กรุณาเลือกพิมพ์ชื่อวัตถุดิบที่ระบุในเมนู`
      };
    }

    await recordDailyCount(matchedItem.code, qty, targetDate);
    
    return {
      type: 'text',
      text: `✅ บันทึกยอดคงเหลือจริงสำเร็จสำหรับวันที่ ${targetDate}!\nวัตถุดิบ: ${matchedItem.nameThai}\nจำนวนที่นับได้จริง: ${qty} ${matchedItem.unit}\nระบบนำไปบันทึกเปรียบเทียบในประวัติสต๊อกเรียบร้อยแล้วค่ะ`
    };
  }

  // 6.2. Partial Match for: คงเหลือ [ชื่อสินค้า] (without quantity specified)
  const partialCountRegex = /^(คงเหลือ|เหลือ|นับได้|count)\s+([ก-๙a-zA-Z0-9\s_]+)$/i;
  const partCountMatch = cleanText.match(partialCountRegex);
  if (partCountMatch) {
    const queryName = partCountMatch[2].trim().toLowerCase();
    const matchedItem = STOCK_ITEMS_LIST.find(item => 
      item.nameThai === queryName || 
      item.code.toLowerCase() === queryName ||
      item.nameThai.includes(queryName)
    );

    if (matchedItem) {
      // Save pending state
      userStates.set(userId, {
        pendingAction: 'count',
        itemCode: matchedItem.code,
        itemName: matchedItem.nameThai,
        date: targetDate
      });

      return {
        type: 'text',
        text: `✍️ บันทึกเลือกคงเหลือจริง: "${matchedItem.nameThai}"\n\n👉 กรุณาพิมพ์ระบุเฉพาะ "จำนวนคงเหลือจริง" ที่นับได้วันนี้ส่งกลับมาได้เลยค่ะ (เช่น "45" หรือ "8")`
      };
    }
  }

  // 6.5. Pattern Parsing for: หมายเหตุ [เพิ่ม|ลด] [ชื่อพิซซ่า]
  const noteRegex = /^(หมายเหตุ|remark|note)\s+(เพิ่ม|ลด|add|sub|remove)\s+(.+)$/i;
  const noteMatch = cleanText.match(noteRegex);
  if (noteMatch) {
    const direction = noteMatch[2].trim().toLowerCase();
    const pizzaName = noteMatch[3].trim();
    const isAdd = direction === 'เพิ่ม' || direction === 'add';
    const cheeseAdjust = 0.25;

    if (isAdd) {
      // Increase cheese count (recorded as replenishment of cheese)
      await replenishStock('cheese', cheeseAdjust, targetDate);
    } else {
      // Decrease cheese count (recorded as sales of cheese)
      const salesMap: Record<string, number> = {};
      STOCK_ITEMS_LIST.forEach(item => {
        salesMap[item.code] = 0;
      });
      salesMap['cheese'] = cheeseAdjust;
      await recordSales(targetDate, salesMap);
    }

    const updatedInv = await getInventory();
    const current = updatedInv.find(i => i.code === 'cheese')?.currentQty || 0;

    return {
      type: 'text',
      text: `📝 บันทึกหมายเหตุ "${pizzaName}" เรียบร้อยแล้วสำหรับวันที่ ${targetDate}!\nวัตถุดิบ: ชีส (Cheese)\nการปรับสต๊อก: ${isAdd ? 'เพิ่มชีส (+)' : 'ลดชีส (-)'} ${cheeseAdjust} ยูนิต\nยอดคงเหลือชีสล่าสุด: ${current} ยูนิต`
    };
  }

  // 6.6. Partial Match for: หมายเหตุ [เพิ่ม|ลด] (without pizza name specified)
  const partialNoteRegex = /^(หมายเหตุ|remark|note)\s+(เพิ่ม|ลด|add|sub|remove)\s*$/i;
  const partNoteMatch = cleanText.match(partialNoteRegex);
  if (partNoteMatch) {
    const direction = partNoteMatch[2].trim().toLowerCase();
    const isAdd = direction === 'เพิ่ม' || direction === 'add';

    // Save pending state
    userStates.set(userId, {
      pendingAction: isAdd ? 'note_add' : 'note_sub',
      itemCode: 'cheese',
      itemName: 'ชีส',
      date: targetDate
    });

    return {
      type: 'text',
      text: `✍️ บันทึกเลือก ${isAdd ? 'เพิ่มชีส (+0.25)' : 'ลดชีส (-0.25)'} สำหรับหน้าครึ่ง\n\n👉 กรุณาพิมพ์ "ชื่อพิซซ่า" ที่ต้องการระบุเข้ามาได้เลยค่ะ (เช่น "ฮาวายเอี้ยน" หรือ "ซีฟู้ด")`
    };
  }

  // Support sending just "หมายเหตุ"
  if (cleanText === 'หมายเหตุ') {
    return {
      type: 'flex',
      text: '📝 เมนูหมายเหตุพิเศษ (พิซซ่าหน้าครึ่ง)',
      flexContent: {
        title: '📝 หมายเหตุพิเศษ (พิซซ่าหน้าครึ่ง)',
        description: 'เลือกประเภทการปรับสต๊อกชีสทีละ 0.25 ยูนิต หรือพิมพ์: หมายเหตุ [เพิ่ม/ลด] [ชื่อพิซซ่า]',
        items: [
          {
            name: '➕ เพิ่มชีส (+0.25)',
            actionText: 'หมายเหตุ เพิ่ม'
          },
          {
            name: '➖ ลดชีส (-0.25)',
            actionText: 'หมายเหตุ ลด'
          }
        ]
      }
    };
  }

  // 7. Standard Friendly Help Guide as a beautiful responsive Flex Menu!
  return {
    type: 'flex',
    text: '🍽️ ระบบจัดการสต๊อกห้องอาหาร ยินดีต้อนรับค่ะ',
    flexContent: {
      title: '🍽️ ระบบจัดการสต๊อกห้องอาหาร',
      description: `วันทำงานบอทขณะนี้: ${activeWorkingDate}\n\nแตะเลือกเมนูด่วนด้านล่าง หรือพิมพ์สั่งงานตรงๆ ได้เลยค่ะ:`,
      items: [
        {
          name: '🟢 บันทึกเติมสต๊อก',
          actionText: 'เติมสต๊อก'
        },
        {
          name: '🔴 ตรวจนับคงเหลือวันนี้',
          actionText: 'คงเหลือ'
        },
        {
          name: '📅 ตั้งวันทำงานย้อนหลัง',
          actionType: 'datetimepicker',
          actionData: 'action=set_working_date'
        },
        {
          name: '🔵 รายงานส่วนต่างสต๊อก',
          actionText: `รายงาน วันที่ ${activeWorkingDate}`
        },
        {
          name: '🔄 รีเซ็ตกลับสู่วันปัจจุบัน',
          actionText: 'วันนี้'
        }
      ]
    }
  };
}


// Helper function to build actual LINE-compatible responses (supporting text and beautiful Flex messages!)
function buildLineMessage(result: any): any {
  if (result.type === 'flex' && result.flexContent) {
    const flexContent = result.flexContent;
    const rows: any[] = [];
    const items = flexContent.items || [];
    
    const buildAction = (item: any) => {
      if (item.actionType === 'datetimepicker') {
        return {
          type: "datetimepicker",
          label: item.name,
          data: item.actionData || "action=set_working_date",
          mode: "date"
        };
      }
      return {
        type: "message",
        label: item.name,
        text: item.actionText
      };
    };

    // Group items into rows of 2 buttons each to look super-polished on mobile screens
    for (let i = 0; i < items.length; i += 2) {
      const item1 = items[i];
      const item2 = items[i + 1];
      
      const rowContents: any[] = [
        {
          type: "button",
          style: "secondary",
          height: "sm",
          action: buildAction(item1)
        }
      ];
      
      if (item2) {
        rowContents.push({
          type: "button",
          style: "secondary",
          height: "sm",
          action: buildAction(item2),
          margin: "sm"
        });
      } else {
        rowContents.push({
          type: "spacer",
          size: "sm"
        });
      }
      
      rows.push({
        type: "box",
        layout: "horizontal",
        margin: "xs",
        contents: rowContents
      });
    }

    return {
      type: "flex",
      altText: flexContent.title,
      contents: {
        type: "bubble",
        size: "mega",
        header: {
          type: "box",
          layout: "vertical",
          backgroundColor: "#1DB446",
          contents: [
            {
              type: "text",
              text: flexContent.title,
              weight: "bold",
              color: "#FFFFFF",
              size: "md"
            }
          ]
        },
        body: {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: flexContent.description,
              wrap: true,
              size: "xs",
              color: "#555555"
            },
            {
              type: "box",
              layout: "vertical",
              margin: "md",
              spacing: "xs",
              contents: rows
            }
          ]
        }
      }
    };
  }

  // Fallback to text message
  return {
    type: "text",
    text: result.text || 'รับทราบคำสั่งค่ะ!'
  };
}


// A. Real LINE Webhook Endpoint (For production messaging integration!)
app.post('/api/line-webhook', async (req, res) => {
  try {
    const events = req.body.events;
    if (!events || events.length === 0) {
      return res.status(200).send('No events');
    }

    // Since we handle actual LINE SDK, parse the text or document and reply
    // To ensure zero failures even if webhook is called with incomplete headers from test tools:
    for (const event of events) {
      if (event.type === 'message') {
        const replyToken = event.replyToken;
        let lineMessagePayload: any = null;
        let originalResult: any = null;

        if (event.message.type === 'text') {
          const userId = event.source?.userId || 'default';
          originalResult = await processBotMessage(event.message.text, undefined, undefined, userId);
          lineMessagePayload = buildLineMessage(originalResult);
        } else if (event.message.type === 'file') {
          const userId = event.source?.userId || 'default';
          const messageId = event.message.id;
          const fileName = event.message.fileName || 'sales.xlsx';
          
          if (process.env.LINE_CHANNEL_ACCESS_TOKEN) {
            try {
              const fileRes = await fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
                headers: {
                  'Authorization': `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`
                }
              });
              if (fileRes.ok) {
                const arrayBuf = await fileRes.arrayBuffer();
                const fileBuf = Buffer.from(arrayBuf);
                originalResult = await processBotMessage('', fileBuf, fileName, userId);
                lineMessagePayload = buildLineMessage(originalResult);
              } else {
                const errText = await fileRes.text();
                console.error('Error downloading from LINE API:', errText);
                lineMessagePayload = {
                  type: 'text',
                  text: '❌ ไม่สามารถดาวน์โหลดไฟล์การขายจาก LINE ได้สำเร็จ กรุณาลองอัปโหลดใหม่อีกครั้งค่ะ'
                };
              }
            } catch (err: any) {
              console.error('Error in LINE file download:', err);
              lineMessagePayload = {
                type: 'text',
                text: `❌ เกิดข้อผิดพลาดขณะดาวน์โหลดและบันทึกไฟล์: ${err.message}`
              };
            }
          } else {
            lineMessagePayload = {
              type: 'text',
              text: '⚠️ บอทยังไม่ได้กำหนด LINE_CHANNEL_ACCESS_TOKEN ในแถบ Secret ของระบบ แอดมินกรุณาตั้งค่าก่อนการประมวลผลไฟล์ค่ะ'
            };
          }
        } else {
          lineMessagePayload = {
            type: 'text',
            text: 'ขออภัยค่ะ ระบบสามารถตรวจสอบคำสั่งแบบข้อความตัวอักษรและไฟล์ Excel ยอดขายเท่านั้นค่ะ'
          };
        }

        // Send actual LINE response back using fetch to LINE API if configured
        if (process.env.LINE_CHANNEL_ACCESS_TOKEN && replyToken && lineMessagePayload) {
          const response = await fetch('https://api.line.me/v2/bot/message/reply', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`
            },
            body: JSON.stringify({
              replyToken,
              messages: [lineMessagePayload]
            })
          });

          if (!response.ok) {
            const errText = await response.text();
            console.error('LINE Reply API error:', errText, 'Payload was:', JSON.stringify(lineMessagePayload));

            // FALLBACK: If the Flex Message payload had any validation issue, immediately send a clean text response so user gets their choices!
            if (lineMessagePayload.type === 'flex' && originalResult) {
              const fallbackText = `📌 ${originalResult.flexContent.title}\n${originalResult.flexContent.description}\n\nพิมพ์สั่งงาน เช่น:\n- เติม ชีส 10\n- คงเหลือ แซลมอน 15`;
              
              await fetch('https://api.line.me/v2/bot/message/reply', {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json',
                  'Authorization': `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`
                },
                body: JSON.stringify({
                  replyToken,
                  messages: [{
                    type: 'text',
                    text: fallbackText
                  }]
                })
              });
            }
          }
        }
      } else if (event.type === 'postback') {
        const replyToken = event.replyToken;
        const userId = event.source?.userId || 'default';
        const postbackData = event.postback?.data;
        const selectedDate = event.postback?.params?.date;

        if (postbackData === 'action=set_working_date' && selectedDate) {
          userWorkingDates.set(userId, selectedDate);
          
          const replyText = `📅 เปลี่ยนวันทำงานของบอทเป็นวันย้อนหลัง: ${selectedDate} เรียบร้อยแล้วค่ะ!\n\nหลังจากนี้ รายการเติมของ, ตรวจนับสต๊อก หรือไฟล์ Excel ยอดขายที่คุณส่งเข้ามา จะถูกบันทึกในวันที่ ${selectedDate} โดยอัตโนมัติค่ะ\n\n(พิมพ์ว่า "วันนี้" เพื่อยกเลิกล็อควันที่ และกลับสู่วันปัจจุบันค่ะ)`;
          
          if (process.env.LINE_CHANNEL_ACCESS_TOKEN && replyToken) {
            await fetch('https://api.line.me/v2/bot/message/reply', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`
              },
              body: JSON.stringify({
                replyToken,
                messages: [{
                  type: 'text',
                  text: replyText
                }]
              })
            });
          }
        }
      }
    }
    res.status(200).send('OK');
  } catch (err: any) {
    console.error('Webhook processing error:', err);
    res.status(200).send('Error but captured'); // return 200 to LINE to avoid endless retries
  }
});


// B. Chat Simulator API for rich browser-based interactive demonstration
app.post('/api/simulator/message', upload.single('file'), async (req, res) => {
  try {
    const text = req.body.text || '';
    let response;

    if (req.file) {
      // User uploaded a sales Excel file through the chat simulator!
      response = await processBotMessage(text, req.file.buffer, req.file.originalname, 'simulator');
    } else {
      response = await processBotMessage(text, undefined, undefined, 'simulator');
    }

    res.json({
      reply: response
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});


// -----------------------------------------------------------------------------
// MAIN SERVER START & ROUTING
// -----------------------------------------------------------------------------

async function startServer() {
  // Initialize Database (Bootstrap or File Fallback)
  await initDatabase();

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running at http://localhost:${PORT}`);
  });
}

startServer();
