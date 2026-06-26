export interface StockItem {
  code: string;
  nameThai: string;
  currentQty: number;
  unit: string;
}

export interface ReplenishmentRecord {
  id: string;
  itemCode: string;
  qty: number;
  timestamp: string; // ISO format
}

export interface DailyCountRecord {
  id: string;
  itemCode: string;
  qty: number; // The reported physical remaining stock
  date: string; // YYYY-MM-DD
  timestamp: string;
}

export interface SalesRecord {
  id: string;
  itemCode: string;
  qty: number; // The sales quantity from Excel
  date: string; // YYYY-MM-DD
  timestamp: string;
}

export interface DiscrepancyReport {
  date: string;
  itemCode: string;
  nameThai: string;
  yesterdayCount: number;
  replenished: number;
  sold: number;
  expectedRemaining: number;
  actualRemaining: number;
  difference: number; // expectedRemaining - actualRemaining (positive means missing, negative means extra)
  remark?: string;
}

export interface DiscrepancyRemark {
  itemCode: string;
  date: string;
  remark: string;
}

export interface ChatMessage {
  id: string;
  sender: 'user' | 'bot';
  text?: string;
  timestamp: string;
  type: 'text' | 'flex' | 'file';
  fileUrl?: string;
  fileName?: string;
  flexContent?: any; // Simulating LINE Flex Messages
}

export const STOCK_ITEMS_MAP: Record<string, { nameThai: string; unit: string }> = {
  cheese: { nameThai: "ชีส", unit: "ยูนิต" },
  seafood_set: { nameThai: "ชุดทะเล", unit: "ยูนิต" },
  salmon: { nameThai: "แซลมอน", unit: "ยูนิต" },
  ham: { nameThai: "แฮม", unit: "ยูนิต" },
  minced_pork: { nameThai: "หมูบด", unit: "ยูนิต" },
  bacon: { nameThai: "เบคอน", unit: "ยูนิต" },
  french_fries: { nameThai: "เฟรนฟราย", unit: "ยูนิต" },
  parma_ham: { nameThai: "พามาแฮม", unit: "ยูนิต" },
  pork_chop: { nameThai: "พ็อกชอป", unit: "ยูนิต" },
  beef_steak: { nameThai: "สเต็กเนื้อ", unit: "ยูนิต" },
  banana_samosa: { nameThai: "ซาโมซ่ากล้วย", unit: "ยูนิต" },
  tuna: { nameThai: "ทูน่า", unit: "ยูนิต" },
  beef_salami: { nameThai: "ซาลามี่เนื้อ", unit: "ยูนิต" },
  champignon_mushroom: { nameThai: "เห็ดแชมปิญอง", unit: "ยูนิต" },
  shrimp: { nameThai: "กุ้ง", unit: "ยูนิต" },
  squid: { nameThai: "หมึก", unit: "ยูนิต" },
  clam: { nameThai: "หอย", unit: "ยูนิต" },
};

export const STOCK_ITEMS_LIST = Object.entries(STOCK_ITEMS_MAP).map(([code, value]) => ({
  code,
  nameThai: value.nameThai,
  unit: value.unit,
}));
