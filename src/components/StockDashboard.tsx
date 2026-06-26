import React, { useState, useEffect } from 'react';
import { 
  BarChart3, RefreshCw, FileSpreadsheet, Plus, CheckCircle, 
  AlertCircle, Database, Sparkles, Download, Upload, Info, HelpCircle
} from 'lucide-react';
import { StockItem, DiscrepancyReport, STOCK_ITEMS_LIST } from '../types.js';

interface StockDashboardProps {
  refreshTrigger: number;
  onRefresh: () => void;
}

export default function StockDashboard({ refreshTrigger, onRefresh }: StockDashboardProps) {
  const [inventory, setInventory] = useState<StockItem[]>([]);
  const [discrepancies, setDiscrepancies] = useState<DiscrepancyReport[]>([]);
  const [dbStatus, setDbStatus] = useState<any>(null);
  const [targetDate, setTargetDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [loading, setLoading] = useState(false);
  
  // Sales upload state
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadSuccess, setUploadSuccess] = useState<string | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);

  // AI State
  const [aiInsights, setAiInsights] = useState<string>('');
  const [loadingAI, setLoadingAI] = useState(false);

  // Remarks state
  const [tempRemarks, setTempRemarks] = useState<Record<string, string>>({});
  const [savingRemark, setSavingRemark] = useState<string | null>(null);

  const handleRemarkChange = (itemCode: string, value: string) => {
    setTempRemarks(prev => ({
      ...prev,
      [itemCode]: value
    }));
  };

  const handleSaveRemark = async (itemCode: string) => {
    const remarkValue = tempRemarks[itemCode];
    if (remarkValue === undefined) return;

    setSavingRemark(itemCode);
    try {
      const res = await fetch('/api/discrepancy/remark', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          itemCode,
          date: targetDate,
          remark: remarkValue,
        }),
      });
      if (res.ok) {
        // Update the discrepancy row locally to stay in sync
        setDiscrepancies(prev => 
          prev.map(row => 
            row.itemCode === itemCode ? { ...row, remark: remarkValue } : row
          )
        );
      }
    } catch (err) {
      console.error('Failed to save remark:', err);
    } finally {
      setSavingRemark(null);
    }
  };

  // Fetch data
  const fetchData = async () => {
    setLoading(true);
    try {
      const [invRes, dbRes, discRes] = await Promise.all([
        fetch('/api/inventory'),
        fetch('/api/db-status'),
        fetch(`/api/discrepancy?date=${targetDate}`)
      ]);

      if (invRes.ok) setInventory(await invRes.json());
      if (dbRes.ok) setDbStatus(await dbRes.json());
      if (discRes.ok) setDiscrepancies(await discRes.json());
    } catch (err) {
      console.error('Error fetching dashboard data:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
    setTempRemarks({});
  }, [refreshTrigger, targetDate]);

  // Request AI Insights
  const generateAIInsights = async () => {
    setLoadingAI(true);
    setAiInsights('');
    try {
      const res = await fetch(`/api/ai-insights?date=${targetDate}`);
      const data = await res.json();
      setAiInsights(data.insights || 'ไม่สามารถดึงข้อมูลวิเคราะห์ได้');
    } catch (err: any) {
      setAiInsights(`เกิดข้อผิดพลาด: ${err.message}`);
    } finally {
      setLoadingAI(false);
    }
  };

  // Upload Excel Sales
  const handleExcelUpload = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!uploadFile) return;

    setUploading(true);
    setUploadSuccess(null);
    setUploadError(null);

    const formData = new FormData();
    formData.append('file', uploadFile);
    formData.append('date', targetDate);

    try {
      const res = await fetch('/api/excel/upload-sales', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();

      if (res.ok) {
        setUploadSuccess(data.message);
        setUploadFile(null);
        // Reset file input element
        const fileInput = document.getElementById('sales-file-input') as HTMLInputElement;
        if (fileInput) fileInput.value = '';
        
        onRefresh(); // Refresh inventory and report
      } else {
        setUploadError(data.error || 'เกิดข้อผิดพลาดในการอัปโหลด');
      }
    } catch (err: any) {
      setUploadError(`เกิดข้อผิดพลาด: ${err.message}`);
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="flex-1 bg-gray-50 p-6 overflow-y-auto h-full space-y-6 font-sans">
      {/* Header Block */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-gray-200 pb-5">
        <div>
          <span className="text-[10px] font-mono tracking-widest uppercase bg-emerald-100 text-emerald-800 px-2.5 py-1 rounded-full font-bold">
            Restaurant Backend Console
          </span>
          <h1 className="text-2xl font-black text-gray-900 mt-2 tracking-tight">
            ระบบตรวจสอบและประเปรียบเทียบสต๊อกวัตถุดิบรายวัน
          </h1>
          <p className="text-sm text-gray-500 mt-1 font-sans">
            สรุปข้อมูลการตัดสต๊อกตามยอดขายประจำวัน เติมเพิ่มสต๊อก และตรวจนับของจริงพร้อมออกไฟล์ Excel
          </p>
        </div>

        {/* Action Controls & Date Picker */}
        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          <div className="flex items-center gap-2 bg-white px-3.5 py-2 rounded-xl shadow-xs border border-gray-200">
            <span className="text-xs font-semibold text-gray-500">เลือกวันที่:</span>
            <input 
              type="date" 
              value={targetDate} 
              onChange={(e) => setTargetDate(e.target.value)}
              className="text-xs font-bold text-gray-800 outline-none cursor-pointer bg-transparent"
            />
          </div>

          <button 
            onClick={() => { fetchData(); onRefresh(); }} 
            disabled={loading}
            className="p-2.5 bg-white hover:bg-gray-100 disabled:opacity-50 text-gray-700 rounded-xl shadow-xs border border-gray-200 transition-colors"
            title="รีเฟรชข้อมูล"
          >
            <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
          </button>
        </div>
      </div>

      {/* Database Status & Platform Details */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {/* Status Card 1 */}
        <div className="bg-white p-4 rounded-2xl shadow-xs border border-gray-200 flex items-center gap-4">
          <div className="w-12 h-12 bg-emerald-50 rounded-xl flex items-center justify-center text-emerald-600 shrink-0">
            <Database className="w-6 h-6" />
          </div>
          <div>
            <p className="text-[11px] text-gray-400 font-mono font-bold uppercase">DATABASE CONNECTION</p>
            <h3 className="text-sm font-bold text-gray-800 mt-0.5">
              {dbStatus?.isUsingMySQL ? 'MySQL Connection' : 'SQLite Local File'}
            </h3>
            <span className="inline-flex items-center gap-1 text-[10px] text-green-600 bg-green-50 px-2 py-0.5 rounded-full font-sans mt-1">
              <span className="w-1.5 h-1.5 rounded-full bg-green-500"></span>
              {dbStatus?.isUsingMySQL ? 'Active Online' : 'Active Offline Fallback'}
            </span>
          </div>
        </div>

        {/* Status Card 2 */}
        <div className="bg-white p-4 rounded-2xl shadow-xs border border-gray-200 flex items-center gap-4">
          <div className="w-12 h-12 bg-blue-50 rounded-xl flex items-center justify-center text-blue-600 shrink-0">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <p className="text-[11px] text-gray-400 font-mono font-bold uppercase">LINE BOT WEBHOOK</p>
            <h3 className="text-sm font-bold text-gray-800 mt-0.5 truncate max-w-[180px]">
              {dbStatus?.lineWebhookUrl === 'Not configured' ? 'Local Sandbox URL' : dbStatus?.lineWebhookUrl}
            </h3>
            <p className="text-[10px] text-gray-400 font-sans mt-0.5">เชื่อมโยง Webhook ทันทีเมื่ออัปขึ้นคลาวด์</p>
          </div>
        </div>

        {/* Status Card 3 */}
        <div className="bg-white p-4 rounded-2xl shadow-xs border border-gray-200 flex items-center gap-4">
          <div className="w-12 h-12 bg-purple-50 rounded-xl flex items-center justify-center text-purple-600 shrink-0">
            <Sparkles className="w-6 h-6" />
          </div>
          <div>
            <p className="text-[11px] text-gray-400 font-mono font-bold uppercase">STOCK ITEMS ON GUARD</p>
            <h3 className="text-sm font-bold text-gray-800 mt-0.5">17 รายการวัตถุดิบสำคัญ</h3>
            <p className="text-[10px] text-gray-400 font-sans mt-0.5">ชีส, แซลมอน, แฮม, หมูบด, ซาโมซ่า, เบคอน, ทะเล</p>
          </div>
        </div>
      </div>

      {/* Main Reconciliation comparison table (โจทย์ต้องการ: นำข้อมูลคงเหลือวันนี้ไปลบเมื่อวาน เช็คว่าหายเท่าไหร่ เทียบกับยอดอัปขาย) */}
      <div className="bg-white rounded-2xl shadow-xs border border-gray-200 overflow-hidden">
        <div className="px-5 py-4 border-b border-gray-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-gray-50/50">
          <div className="flex items-center gap-2">
            <BarChart3 className="w-5 h-5 text-gray-800" />
            <div>
              <h2 className="font-bold text-base text-gray-900 leading-tight">ตารางเปรียบเทียบข้อมูลยอดขายและยอดคงเหลือรายวัตถุดิบ</h2>
              <p className="text-xs text-gray-400 mt-0.5">Reconciliation Ledger ประจําวันที่ {targetDate}</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <a 
              href={`/api/excel/download-report?date=${targetDate}`}
              download
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl transition-all shadow-sm"
            >
              <Download className="w-3.5 h-3.5" />
              ดาวน์โหลดรายงานสต๊อก Excel
            </a>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-gray-100/70 border-b border-gray-200 text-gray-500 font-mono uppercase text-[10px]">
                <th className="py-3 px-4 font-bold text-gray-600">ชื่อสินค้า (Thai)</th>
                <th className="py-3 px-4 text-right font-bold text-gray-600">คงเหลือยกมา (เมื่อวาน)</th>
                <th className="py-3 px-4 text-right text-emerald-600 font-bold">เติมสต๊อก (+)</th>
                <th className="py-3 px-4 text-right text-red-600 font-bold">ตัดยอดขาย (-)</th>
                <th className="py-3 px-4 text-right font-bold text-gray-600 bg-blue-50/40">สต๊อกตามทฤษฎี</th>
                <th className="py-3 px-4 text-right font-bold text-gray-800 bg-emerald-50/20">คงเหลือนับจริง</th>
                <th className="py-3 px-4 text-center font-bold text-gray-600">ผลต่าง (สูญหาย)</th>
                <th className="py-3 px-4 text-center font-bold text-gray-600">สถานะ</th>
                <th className="py-3 px-4 text-left font-bold text-gray-600 min-w-[200px]">หมายเหตุผลต่าง</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {discrepancies.length === 0 ? (
                <tr>
                  <td colSpan={9} className="py-12 text-center text-gray-400 font-sans">
                    ไม่มีข้อมูลสำหรับวันที่เลือก กรุณากดกรอกยอดคงเหลือผ่านบอทไลน์จำลอง หรือ อัปโหลดไฟล์ Excel ยอดขายเพื่อตัดสต๊อก
                  </td>
                </tr>
              ) : (
                discrepancies.map((row) => {
                  const unit = STOCK_ITEMS_LIST.find(item => item.code === row.itemCode)?.unit || '';
                  const hasDisc = row.difference !== 0;
                  const currentRemark = tempRemarks[row.itemCode] !== undefined ? tempRemarks[row.itemCode] : (row.remark || '');
                  const needsRemark = hasDisc && !currentRemark.trim();
                  
                  return (
                    <tr key={row.itemCode} className={`hover:bg-gray-50/80 transition-colors ${hasDisc ? 'bg-red-50/10' : ''}`}>
                      <td className="py-3 px-4 font-bold text-gray-800 flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-gray-300"></span>
                        {row.nameThai}
                      </td>
                      <td className="py-3 px-4 text-right font-semibold text-gray-500 font-mono">
                        {row.yesterdayCount} {unit}
                      </td>
                      <td className="py-3 px-4 text-right font-bold text-emerald-600 font-mono">
                        {row.replenished > 0 ? `+${row.replenished}` : '0'} {unit}
                      </td>
                      <td className="py-3 px-4 text-right font-bold text-red-500 font-mono">
                        {row.sold > 0 ? `-${row.sold}` : '0'} {unit}
                      </td>
                      <td className="py-3 px-4 text-right font-bold text-blue-700 font-mono bg-blue-50/20">
                        {row.expectedRemaining} {unit}
                      </td>
                      <td className="py-3 px-4 text-right font-bold text-emerald-800 font-mono bg-emerald-50/10">
                        {row.actualRemaining} {unit}
                      </td>
                      <td className={`py-3 px-4 text-center font-bold font-mono text-[11px] ${
                        row.difference === 0 
                        ? 'text-gray-400' 
                        : row.difference > 0 
                          ? 'text-red-600 bg-red-50/50 rounded-lg px-2' 
                          : 'text-amber-600 bg-amber-50/50 rounded-lg px-2'
                      }`}>
                        {row.difference === 0 ? '0' : row.difference > 0 ? `หาย -${row.difference}` : `เกิน +${Math.abs(row.difference)}`}
                      </td>
                      <td className="py-3 px-4 text-center">
                        {row.difference === 0 ? (
                          <span className="inline-flex items-center gap-1 text-[10px] text-green-700 bg-green-50 px-2.5 py-0.5 rounded-full font-bold">
                            <CheckCircle className="w-3 h-3 shrink-0" /> ตรงกัน
                          </span>
                        ) : (
                          <span className={`inline-flex items-center gap-1 text-[10px] px-2.5 py-0.5 rounded-full font-bold ${
                            row.difference > 0 
                            ? 'text-red-700 bg-red-50' 
                            : 'text-amber-700 bg-amber-50'
                          }`}>
                            <AlertCircle className="w-3 h-3 shrink-0" /> คลาดเคลื่อน
                          </span>
                        )}
                      </td>
                      <td className="py-2 px-4 text-left">
                        <div className="flex items-center gap-1.5 w-full">
                          <input
                            type="text"
                            placeholder={hasDisc ? "โปรดระบุเหตุผลผลต่าง..." : "บันทึกข้อความเพิ่มเติม..."}
                            value={currentRemark}
                            onChange={(e) => handleRemarkChange(row.itemCode, e.target.value)}
                            onBlur={() => handleSaveRemark(row.itemCode)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                handleSaveRemark(row.itemCode);
                              }
                            }}
                            className={`w-full min-w-[160px] max-w-[260px] px-2.5 py-1 text-xs rounded-lg border font-sans focus:outline-none transition-all focus:ring-1 focus:ring-blue-500 focus:border-blue-500 ${
                              needsRemark 
                                ? 'border-amber-300 bg-amber-50/30 placeholder-amber-500 text-amber-900 focus:ring-amber-500 focus:border-amber-500 animate-pulse' 
                                : 'border-gray-200 focus:ring-blue-500 focus:border-blue-500'
                            }`}
                          />
                          {savingRemark === row.itemCode && (
                            <span className="text-[10px] text-blue-500 font-mono animate-pulse shrink-0">กำลังบันทึก...</span>
                          )}
                          {!savingRemark && currentRemark && (
                            <span className="text-[10px] text-emerald-600 font-bold shrink-0">✓ บันทึกแล้ว</span>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
        {/* Left Side: Upload excel and download templates (4 cols) */}
        <div className="xl:col-span-5 space-y-6">
          {/* Quick Manual Upload for testing */}
          <div className="bg-white p-5 rounded-2xl shadow-xs border border-gray-200">
            <h3 className="font-bold text-sm text-gray-800 mb-1 flex items-center gap-1.5">
              <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
              นำเข้าข้อมูลการขายรายวัน (Excel)
            </h3>
            <p className="text-xs text-gray-400 mb-4">
              อัปเดตไฟล์สรุปยอดขายจาก LINE หรืออัปที่นี่โดยตรงเพื่อหักสต๊อกตามจริง
            </p>

            <form onSubmit={handleExcelUpload} className="space-y-3.5">
              <div className="border-2 border-dashed border-gray-200 hover:border-emerald-500 rounded-xl p-4 flex flex-col items-center justify-center text-center cursor-pointer bg-gray-50/50 transition-colors">
                <input 
                  type="file" 
                  id="sales-file-input"
                  accept=".xlsx,.xls" 
                  onChange={(e) => setUploadFile(e.target.files?.[0] || null)}
                  className="hidden"
                />
                <label htmlFor="sales-file-input" className="cursor-pointer flex flex-col items-center justify-center">
                  <Upload className="w-8 h-8 text-emerald-500 mb-2" />
                  <span className="text-xs font-semibold text-gray-700 block">
                    {uploadFile ? uploadFile.name : 'เลือกไฟล์ Excel ยอดขาย (.xlsx)'}
                  </span>
                  <span className="text-[10px] text-gray-400 mt-1 block">ลากไฟล์มาวางหรือกดตรงนี้เพื่อเปิดเครื่อง</span>
                </label>
              </div>

              {uploadSuccess && (
                <div className="p-3 bg-green-50 text-green-800 text-xs rounded-xl border border-green-100 font-semibold flex items-start gap-2">
                  <CheckCircle className="w-4 h-4 text-green-600 shrink-0 mt-0.5" />
                  <span>{uploadSuccess}</span>
                </div>
              )}

              {uploadError && (
                <div className="p-3 bg-red-50 text-red-800 text-xs rounded-xl border border-red-100 font-semibold flex items-start gap-2">
                  <AlertCircle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
                  <span>{uploadError}</span>
                </div>
              )}

              <div className="flex gap-2">
                <button
                  type="submit"
                  disabled={!uploadFile || uploading}
                  className="flex-1 py-2 bg-[#06c755] hover:bg-green-600 disabled:opacity-50 text-white font-bold text-xs rounded-xl transition-all shadow-sm"
                >
                  {uploading ? 'กำลังนำเข้า...' : 'อัปโหลดและประมวลผลสต๊อก'}
                </button>
                <a
                  href="/api/excel/sales-template"
                  download
                  className="px-3.5 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl transition-colors text-xs font-bold flex items-center justify-center shrink-0"
                  title="ดาวน์โหลดเทมเพลต Excel"
                >
                  <Download className="w-4 h-4" />
                </a>
              </div>
            </form>
          </div>

          {/* Quick Guide */}
          <div className="bg-[#1e1e24] text-white p-5 rounded-2xl shadow-xs border border-gray-800">
            <h4 className="font-bold text-xs uppercase text-emerald-400 tracking-wider mb-2 flex items-center gap-1.5">
              <Info className="w-4 h-4" />
              สรุปวิธีการคำนวณและตัดยอดสต๊อก
            </h4>
            <ul className="text-xs text-gray-300 space-y-2.5 font-sans leading-relaxed">
              <li className="flex items-start gap-2">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 mt-1.5 shrink-0"></span>
                <span><strong>เติมสต๊อก (Top Up)</strong>: ยอดสินค้าคงเหลือปัจจุบันในฐานข้อมูลจะเพิ่มขึ้นทันทีตามปริมาณที่เติม</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 mt-1.5 shrink-0"></span>
                <span><strong>ตัดสต๊อกยอดขาย (Deduction)</strong>: เมื่อได้รับไฟล์ Excel ข้อมูลยอดขาย ระบบจะหักลบจำนวนที่ขายออกไปจากสต๊อกปัจจุบันทันที</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 mt-1.5 shrink-0"></span>
                <span><strong>นับสต๊อกจริง (Physical Count)</strong>: การบันทึกคงเหลือในตอนท้ายวันจะเป็นตัวลบยอดสต๊อกทฤษฎี (เมื่อวาน + เติม - ขาย) และบันทึกผลต่างที่หายไป (Discrepancy)</span>
              </li>
            </ul>
          </div>
        </div>

        {/* Right Side: Smart AI discrepancy analysis using Gemini (7 cols) */}
        <div className="xl:col-span-7">
          <div className="bg-white p-5 rounded-2xl shadow-xs border border-gray-200 h-full flex flex-col">
            <div className="flex items-center justify-between border-b pb-3.5 mb-4">
              <div className="flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-purple-600 animate-pulse" />
                <div>
                  <h3 className="font-bold text-sm text-gray-900">AI ช่วยตรวจเช็คและวิเคราะห์สต๊อกรั่วไหล</h3>
                  <p className="text-xs text-gray-400 mt-0.5 font-sans">Gemini 3.5 Flash Smart Analyzer</p>
                </div>
              </div>
              
              <button
                onClick={generateAIInsights}
                disabled={loadingAI}
                className="inline-flex items-center gap-1 px-3 py-1.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs rounded-xl transition-all shadow-md disabled:opacity-50 shrink-0"
              >
                <Sparkles className="w-3.5 h-3.5" />
                {loadingAI ? 'กำลังวิเคราะห์...' : 'เริ่มให้ AI วิเคราะห์'}
              </button>
            </div>

            <div className="flex-1 bg-gray-50/50 rounded-xl p-4 min-h-[160px] border border-gray-100 flex flex-col justify-between">
              {aiInsights ? (
                <div className="text-xs text-gray-700 font-sans leading-relaxed whitespace-pre-wrap">
                  {aiInsights}
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center text-center h-full text-gray-400 py-8 my-auto">
                  <Sparkles className="w-8 h-8 text-purple-300 mb-2.5" />
                  <p className="text-xs font-semibold">ยังไม่ได้เริ่มวิเคราะห์</p>
                  <p className="text-[10px] text-gray-400 mt-1 max-w-[280px]">กดปุ่มด้านขวาเพื่อส่งรายงานสต๊อกวันนี้ให้ AI วิเคราะห์หาสินค้าสูญเสียหรือปัญหาคลาดเคลื่อนทันที</p>
                </div>
              )}

              {aiInsights && (
                <div className="text-[9px] text-gray-400 font-mono mt-4 pt-2.5 border-t border-gray-200/60 flex items-center gap-1 justify-end">
                  <span>Powered by Gemini 3.5 Flash</span>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Grid of 17 Stock Item Cards for quick overview (โจทย์ต้องการ: โชว์ข้อมูลสต๊อกคงเหลือรายวัน) */}
      <div className="space-y-3">
        <h3 className="font-black text-sm text-gray-900 flex items-center gap-2">
          <span className="w-3 h-3 bg-emerald-500 rounded-full"></span>
          ระดับคงเหลือสต๊อกปัจจุบัน (Real-time Physical Stock Levels)
        </h3>
        
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-6 gap-3.5">
          {inventory.map((item) => {
            const isLow = item.currentQty < 15;
            return (
              <div 
                key={item.code} 
                className={`p-3.5 rounded-2xl bg-white border shadow-xs transition-all hover:scale-[1.02] flex flex-col justify-between h-[96px] ${
                  isLow ? 'border-red-200 bg-red-50/10' : 'border-gray-200 hover:border-gray-300'
                }`}
              >
                <div>
                  <span className={`w-2 h-2 rounded-full inline-block mr-1.5 ${isLow ? 'bg-red-500' : 'bg-emerald-500'}`}></span>
                  <span className="text-[11px] font-mono text-gray-400 font-bold uppercase">{item.code}</span>
                  <h4 className="font-black text-sm text-gray-800 tracking-tight mt-0.5 truncate">{item.nameThai}</h4>
                </div>
                
                <div className="flex items-baseline justify-between mt-2">
                  <span className={`text-lg font-black font-mono leading-none ${isLow ? 'text-red-600' : 'text-gray-900'}`}>
                    {item.currentQty}
                  </span>
                  <span className="text-[10px] text-gray-400 font-sans font-semibold">{item.unit}</span>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}
