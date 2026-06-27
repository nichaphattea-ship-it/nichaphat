import React, { useState, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Send, Upload, FileSpreadsheet, Bot, User, CheckCircle2, AlertTriangle, Download, ChevronRight } from 'lucide-react';
import { ChatMessage, STOCK_ITEMS_LIST, STOCK_ITEMS_MAP } from '../types.js';

interface LineSimulatorProps {
  onDatabaseUpdate: () => void;
}

export default function LineSimulator({ onDatabaseUpdate }: LineSimulatorProps) {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      sender: 'bot',
      text: `สวัสดีค่ะ บอทจัดการสต๊อกห้องอาหารยินดีให้บริการค่ะ! 🍽️\n\nท่านสามารถใช้งานด่วนผ่าน Rich Menu ด้านล่าง:\n🟢 ปุ่ม "เติมสต๊อก" -> เพื่อระบุการเติมสินค้า\n🔴 ปุ่ม "คงเหลือ" -> เพื่อบันทึกตรวจนับคงเหลือวันนี้\n🔵 ปุ่ม "รายงาน" -> เพื่อตรวจสอบส่วนต่างและรับไฟล์ Excel\n\nหรือพิมพ์ข้อความได้ทันที เช่น:\n- เติม ชีส 10\n- คงเหลือ แซลมอน 15\n- อัปโหลดไฟล์ Excel ยอดขายเพื่อหักสต๊อกรายวัน`,
      timestamp: new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }),
      type: 'text'
    }
  ]);
  const [inputText, setInputText] = useState('');
  const [activeMenuTab, setActiveMenuTab] = useState<'none' | 'replenish' | 'count' | 'note' | 'report'>(() => {
    const params = new URLSearchParams(window.location.search);
    const tabParam = params.get('tab');
    if (tabParam === 'replenish' || tabParam === 'count') {
      return tabParam;
    }
    return 'none';
  });
  const [inputQty, setInputQty] = useState('');
  const [selectedItem, setSelectedItem] = useState<any>(null);
  const [noteDirection, setNoteDirection] = useState<'เพิ่ม' | 'ลด'>('เพิ่ม');
  const [pizzaName, setPizzaName] = useState('');
  const [isTyping, setIsTyping] = useState(false);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedDate, setSelectedDate] = useState(() => {
    const params = new URLSearchParams(window.location.search);
    const dateParam = params.get('date');
    if (dateParam && /^\d{4}-\d{2}-\d{2}$/.test(dateParam)) {
      return dateParam;
    }
    return new Date().toISOString().split('T')[0];
  });
  const [quantities, setQuantities] = useState<Record<string, string>>({});
  const [flexQuantities, setFlexQuantities] = useState<Record<string, Record<string, string>>>({});
  const [submittedMessages, setSubmittedMessages] = useState<string[]>([]);

  // Auto scroll to bottom
  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isTyping]);

  const addMessage = (msg: Omit<ChatMessage, 'id' | 'timestamp'>) => {
    const newMsg: ChatMessage = {
      ...msg,
      id: `${Date.now()}-${Math.random()}`,
      timestamp: new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })
    };
    setMessages(prev => [...prev, newMsg]);
  };

  const handleSendMessage = async (textToSend: string, fileToUpload?: File) => {
    if (!textToSend.trim() && !fileToUpload) return;

    // 1. Add User Message to Chat Log
    if (fileToUpload) {
      addMessage({
        sender: 'user',
        type: 'file',
        text: `ส่งไฟล์: ${fileToUpload.name}`,
        fileName: fileToUpload.name
      });
    } else {
      addMessage({
        sender: 'user',
        type: 'text',
        text: textToSend
      });
    }

    setInputText('');
    setIsTyping(true);

    try {
      // 2. Prepare payload
      const formData = new FormData();
      formData.append('text', textToSend);
      if (fileToUpload) {
        formData.append('file', fileToUpload);
      }

      // 3. Post to Express Simulator API
      const response = await fetch('/api/simulator/message', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Network error sending message');
      }

      let data;
      try {
        data = await response.json();
      } catch (jsonErr) {
        throw new Error('ระบบเซิร์ฟเวอร์กำลังบูตหรือรีสตาร์ทเพื่ออัปเดตข้อมูล กรุณารอประมาณ 3-5 วินาทีแล้วกดส่งอีกครั้งนะคะ 🙏');
      }
      
      // Artificial delay to mimic typing
      await new Promise<void>((resolve) => {
        setTimeout(() => {
          setIsTyping(false);
          addMessage({
            sender: 'bot',
            type: data.reply.type || 'text',
            text: data.reply.text,
            flexContent: data.reply.flexContent
          });
          
          // Trigger dashboard data refresh
          onDatabaseUpdate();
          resolve();
        }, 700);
      });

    } catch (err: any) {
      setIsTyping(false);
      addMessage({
        sender: 'bot',
        type: 'text',
        text: `❌ เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์: ${err.message}`
      });
    }
  };

  // Synchronize URL query parameters with bot actions on mount
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const tabParam = params.get('tab');
    const dateParam = params.get('date');
    
    if (dateParam && /^\d{4}-\d{2}-\d{2}$/.test(dateParam)) {
      handleSendMessage(`วันที่ ${dateParam}`);
    }
    
    if (tabParam === 'replenish') {
      setTimeout(() => {
        handleSendMessage('เติมสต๊อก');
      }, 500);
    } else if (tabParam === 'count') {
      setTimeout(() => {
        handleSendMessage('คงเหลือ');
      }, 500);
    }
  }, []);

  const handleRichMenuClick = (menuType: 'replenish' | 'count' | 'note' | 'report') => {
    setQuantities({});
    // Close the web sheet panel so that chat bubbles and bot cards are fully visible in the chat interface
    setActiveMenuTab('none');
    
    if (menuType === 'report') {
      handleSendMessage(`รายงาน วันที่ ${selectedDate}`);
    } else if (menuType === 'replenish') {
      handleSendMessage('เติมสต๊อก');
    } else if (menuType === 'count') {
      handleSendMessage('คงเหลือ');
    } else if (menuType === 'note') {
      handleSendMessage('หมายเหตุ');
    }
  };

  const handleQtyInputChange = (itemCode: string, value: string) => {
    setQuantities(prev => ({
      ...prev,
      [itemCode]: value
    }));
  };

  const handleIncrement = (itemCode: string) => {
    setQuantities(prev => {
      const currentStr = prev[itemCode] ?? "0";
      let currentVal = parseFloat(currentStr);
      if (isNaN(currentVal)) currentVal = 0;
      return {
        ...prev,
        [itemCode]: String(currentVal + 1)
      };
    });
  };

  const handleDecrement = (itemCode: string) => {
    setQuantities(prev => {
      const currentStr = prev[itemCode] ?? "0";
      let currentVal = parseFloat(currentStr);
      if (isNaN(currentVal)) currentVal = 0;
      const newVal = Math.max(0, currentVal - 1);
      return {
        ...prev,
        [itemCode]: String(newVal)
      };
    });
  };

  const handleBulkSubmit = async () => {
    const activeItems = STOCK_ITEMS_LIST.filter(item => {
      const val = quantities[item.code];
      if (val === undefined || val === '') return false;
      const num = parseFloat(val);
      if (isNaN(num)) return false;
      if (activeMenuTab === 'replenish') {
        return num > 0;
      } else {
        return num >= 0;
      }
    });

    if (activeItems.length === 0) {
      alert('กรุณากรอกจำนวนอย่างน้อย 1 รายการ');
      return;
    }

    const mode = activeMenuTab;
    const quantitiesCopy = { ...quantities };
    setActiveMenuTab('none');
    setQuantities({});

    // Send commands sequentially
    for (const item of activeItems) {
      const qty = parseFloat(quantitiesCopy[item.code] || '0');
      const commandText = mode === 'replenish'
        ? `เติม ${item.nameThai} ${qty} วันที่ ${selectedDate}`
        : `คงเหลือ ${item.nameThai} ${qty} วันที่ ${selectedDate}`;
      
      await handleSendMessage(commandText);
    }
  };

  const getFlexQty = (msgId: string, itemCode: string) => {
    return flexQuantities[msgId]?.[itemCode] ?? "0";
  };

  const setFlexQty = (msgId: string, itemCode: string, value: string) => {
    setFlexQuantities(prev => ({
      ...prev,
      [msgId]: {
        ...(prev[msgId] || {}),
        [itemCode]: value
      }
    }));
  };

  const handleFlexIncrement = (msgId: string, itemCode: string) => {
    const currentStr = getFlexQty(msgId, itemCode);
    let currentVal = parseFloat(currentStr);
    if (isNaN(currentVal)) currentVal = 0;
    setFlexQty(msgId, itemCode, String(currentVal + 1));
  };

  const handleFlexDecrement = (msgId: string, itemCode: string) => {
    const currentStr = getFlexQty(msgId, itemCode);
    let currentVal = parseFloat(currentStr);
    if (isNaN(currentVal)) currentVal = 0;
    const newVal = Math.max(0, currentVal - 1);
    setFlexQty(msgId, itemCode, String(newVal));
  };

  const handleFlexBulkSubmit = async (msgId: string, isReplenish: boolean) => {
    const qMap = flexQuantities[msgId] || {};
    const activeItems = STOCK_ITEMS_LIST.filter(item => {
      const val = qMap[item.code];
      if (val === undefined || val === '') return false;
      const num = parseFloat(val);
      if (isNaN(num)) return false;
      if (isReplenish) {
        return num > 0;
      } else {
        return num >= 0;
      }
    });

    if (activeItems.length === 0) {
      alert('กรุณากรอกจำนวนอย่างน้อย 1 รายการ');
      return;
    }

    setSubmittedMessages(prev => [...prev, msgId]);

    // Build a single bulk command string!
    // Example: "เติม ชีส 10 แซลมอน 5 วันที่ 2026-06-27"
    const prefix = isReplenish ? 'เติม' : 'คงเหลือ';
    const itemsStr = activeItems.map(item => {
      const qty = parseFloat(qMap[item.code] || '0');
      return `${item.nameThai} ${qty}`;
    }).join(' ');

    const commandText = `${prefix} ${itemsStr} วันที่ ${selectedDate}`;
    await handleSendMessage(commandText);
  };

  const handleItemSelectInMenu = (item: any) => {
    setSelectedItem(item);
    setInputQty('');
  };

  const handleConfirmActionInMenu = () => {
    if (activeMenuTab === 'report') {
      const commandText = `รายงาน วันที่ ${selectedDate}`;
      setActiveMenuTab('none');
      handleSendMessage(commandText);
      return;
    }

    if (activeMenuTab === 'note') {
      if (!pizzaName.trim()) {
        alert('กรุณากรอกชื่อพิซซ่า');
        return;
      }
      const commandText = `หมายเหตุ ${noteDirection} ${pizzaName.trim()} วันที่ ${selectedDate}`;
      setActiveMenuTab('none');
      handleSendMessage(commandText);
      return;
    }

    if (!selectedItem || !inputQty.trim()) return;
    const qty = parseFloat(inputQty);
    if (isNaN(qty) || qty <= 0 && activeMenuTab === 'replenish') {
      alert('กรุณากรอกจำนวนที่ถูกต้อง (มากกว่า 0)');
      return;
    }

    const commandText = activeMenuTab === 'replenish' 
      ? `เติม ${selectedItem.nameThai} ${qty} วันที่ ${selectedDate}` 
      : `คงเหลือ ${selectedItem.nameThai} ${qty} วันที่ ${selectedDate}`;

    setSelectedItem(null);
    setInputQty('');
    handleSendMessage(commandText);
  };

  const handleFileUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      handleSendMessage(`วันที่ ${selectedDate}`, file);
    }
  };

  return (
    <div id="line-simulator-card" className="bg-[#1e1e24] rounded-3xl p-4 shadow-2xl flex flex-col h-[680px] w-full max-w-[390px] border border-gray-800 relative mx-auto overflow-hidden text-white font-sans">
      {/* Phone Notch & Speaker */}
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-40 h-5 bg-black rounded-b-2xl z-20 flex justify-center items-center">
        <div className="w-16 h-1 bg-gray-700 rounded-full mb-1"></div>
      </div>

      {/* Real-time Status Bar */}
      <div className="flex justify-between items-center px-4 pt-1 pb-3 text-[11px] text-gray-400 font-mono z-10">
        <div>AIS 5G</div>
        <div className="font-sans font-semibold">21:26</div>
        <div className="flex items-center gap-1">
          <span>98%</span>
          <div className="w-5 h-2.5 border border-gray-400 rounded-sm p-0.5 flex">
            <div className="bg-green-500 w-full h-full rounded-2xs"></div>
          </div>
        </div>
      </div>

      {/* LINE Bot Header */}
      <div className="bg-[#06c755] rounded-2xl p-3 flex items-center justify-between shadow-md mb-2 relative">
        <div className="flex items-center gap-2.5">
          <div className="w-10 h-10 rounded-full bg-white flex items-center justify-center border-2 border-green-200">
            <Bot className="w-6 h-6 text-[#06c755]" />
          </div>
          <div>
            <h3 className="font-bold text-sm tracking-wide leading-none flex items-center gap-1.5">
              LINE Bot Stock
              <span className="w-2 h-2 rounded-full bg-emerald-300 animate-pulse"></span>
            </h3>
            <p className="text-[10px] text-green-100 mt-0.5 font-sans">บอทเติม/ตัดสต๊อกร้านอาหาร</p>
          </div>
        </div>
        <div className="text-[10px] bg-green-700/40 text-green-100 px-2 py-1 rounded-full font-mono">
          LINE API LIVE
        </div>
      </div>

      {/* Simulation Date Selector Bar */}
      <div className="bg-gray-800/80 border border-gray-700/60 rounded-xl p-2 mb-2 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-1.5">
          <span className="flex h-1.5 w-1.5 relative">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-500"></span>
          </span>
          <span className="text-[10px] font-bold text-gray-300 font-sans">เลือกวันจำลองข้อมูลย้อนหลัง:</span>
        </div>
        <input
          type="date"
          value={selectedDate}
          onChange={(e) => {
            const newDate = e.target.value;
            setSelectedDate(newDate);
            handleSendMessage(`วันที่ ${newDate}`);
          }}
          className="text-[11px] font-black text-emerald-400 bg-[#121215] px-2 py-0.5 rounded border border-gray-700 outline-none cursor-pointer focus:ring-1 focus:ring-[#06c755] text-center"
        />
      </div>



      {/* Chat Messages Display */}
      <div className="flex-1 overflow-y-auto px-2 py-3 space-y-4 bg-[#8c9ca9] rounded-2xl scrollbar-thin scrollbar-thumb-gray-600/30">
        <AnimatePresence initial={false}>
          {messages.map((msg) => (
            <motion.div
              key={msg.id}
              initial={{ opacity: 0, y: 10, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              transition={{ duration: 0.2 }}
              className={`flex items-start gap-1.5 ${msg.sender === 'user' ? 'flex-row-reverse' : ''}`}
            >
              {/* Profile Avatar */}
              {msg.sender === 'bot' ? (
                <div className="w-7 h-7 rounded-full bg-[#06c755] flex items-center justify-center text-white shrink-0 shadow">
                  <Bot className="w-4 h-4" />
                </div>
              ) : (
                <div className="w-7 h-7 rounded-full bg-[#1e293b] flex items-center justify-center text-white shrink-0 shadow border border-gray-700">
                  <User className="w-4 h-4 text-emerald-400" />
                </div>
              )}

              {/* Message Bubble */}
              <div className="flex flex-col max-w-[75%]">
                <span className="text-[9px] text-gray-200 mb-0.5 ml-1 select-none font-sans">
                  {msg.sender === 'bot' ? 'บอทสต๊อก' : 'ผู้ใช้งาน'}
                </span>
                
                {/* Text Bubble */}
                {msg.type === 'text' && (
                  <div className={`p-3 rounded-2xl text-xs shadow-sm leading-relaxed whitespace-pre-line ${
                    msg.sender === 'user' 
                      ? 'bg-[#06c755] text-white rounded-tr-xs' 
                      : 'bg-white text-gray-800 rounded-tl-xs'
                  }`}>
                    {msg.text}
                  </div>
                )}

                {/* File Bubble */}
                {msg.type === 'file' && (
                  <div className={`p-3 rounded-2xl text-xs shadow-md flex items-center gap-2.5 ${
                    msg.sender === 'user' 
                      ? 'bg-[#05b04b] text-white rounded-tr-xs' 
                      : 'bg-white text-gray-800 rounded-tl-xs'
                  }`}>
                    <FileSpreadsheet className="w-7 h-7 text-green-300 shrink-0" />
                    <div className="overflow-hidden">
                      <p className="font-semibold truncate">{msg.fileName}</p>
                      <p className="text-[9px] opacity-80 font-sans">อัปโหลดไฟล์สำเร็จ</p>
                    </div>
                  </div>
                )}

                {/* Flex Message Simulation */}
                {msg.type === 'flex' && (() => {
                  const isReplenish = msg.flexContent?.title?.includes('เติมสต๊อก');
                  const isCount = msg.flexContent?.title?.includes('คงเหลือ');
                  const isInteractive = isReplenish || isCount;
                  
                  if (isInteractive) {
                    const isSubmitted = submittedMessages.includes(msg.id);
                    return (
                      <div className="bg-white rounded-2xl text-gray-800 overflow-hidden shadow-lg border border-gray-100 rounded-tl-xs flex flex-col w-[300px] max-w-full">
                        {/* Header banner */}
                        <div className={`p-2.5 text-xs font-bold flex items-center justify-center text-white shrink-0 ${
                          isReplenish ? 'bg-[#06c755]' : 'bg-rose-600'
                        }`}>
                          {msg.flexContent?.title}
                        </div>
                        {/* Description */}
                        <div className="p-3 text-[11px] text-gray-600 whitespace-pre-line border-b border-gray-100">
                          {msg.flexContent?.description}
                        </div>
                        {/* List of items in a 2-column grid */}
                        <div className="max-h-72 overflow-y-auto p-2 scrollbar-thin">
                          <div className="grid grid-cols-2 gap-2">
                            {msg.flexContent?.items?.filter((item: any) => item.actionType !== 'uri').map((item: any) => {
                              const itemCode = item.code || item.name;
                              const currentVal = getFlexQty(msg.id, itemCode);
                              return (
                                <div key={itemCode} className="flex flex-col p-1.5 bg-gray-50 border border-gray-100 rounded-xl space-y-1">
                                  <span className="text-[10px] font-bold text-gray-700 truncate">{item.name}</span>
                                  <div className="flex items-center justify-between gap-1">
                                    <button
                                      type="button"
                                      disabled={isSubmitted}
                                      onClick={() => handleFlexDecrement(msg.id, itemCode)}
                                      className="w-5 h-5 bg-white hover:bg-gray-100 text-gray-700 font-bold rounded flex items-center justify-center border border-gray-200 select-none cursor-pointer transition-all active:scale-90 disabled:opacity-50 text-xs"
                                    >
                                      -
                                    </button>
                                    <input
                                      type="text"
                                      disabled={isSubmitted}
                                      value={currentVal}
                                      onChange={(e) => setFlexQty(msg.id, itemCode, e.target.value)}
                                      className="w-8 h-5 bg-white border border-gray-200 rounded text-center font-bold text-[10px] outline-none focus:ring-1 focus:ring-emerald-500 disabled:opacity-50"
                                    />
                                    <button
                                      type="button"
                                      disabled={isSubmitted}
                                      onClick={() => handleFlexIncrement(msg.id, itemCode)}
                                      className="w-5 h-5 bg-white hover:bg-gray-100 text-gray-700 font-bold rounded flex items-center justify-center border border-gray-200 select-none cursor-pointer transition-all active:scale-90 disabled:opacity-50 text-xs"
                                    >
                                      +
                                    </button>
                                  </div>
                                </div>
                              );
                            })}
                          </div>

                          {/* Render URI buttons at the bottom of the card, full width */}
                          {msg.flexContent?.items?.filter((item: any) => item.actionType === 'uri').map((item: any) => (
                            <button
                              key={item.name}
                              onClick={() => {
                                try {
                                  const url = new URL(item.actionUri);
                                  const tab = url.searchParams.get('tab') as any;
                                  if (tab) {
                                    setActiveMenuTab(tab);
                                  }
                                } catch (e) {
                                  if (item.actionUri?.includes('tab=replenish')) {
                                    setActiveMenuTab('replenish');
                                  } else if (item.actionUri?.includes('tab=count')) {
                                    setActiveMenuTab('count');
                                  }
                                }
                              }}
                              className="mt-3 block w-full text-center px-3 py-2 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 text-white font-bold text-xs rounded-xl transition-all shadow-sm cursor-pointer"
                            >
                              {item.name}
                            </button>
                          ))}
                        </div>
                        {/* Submit Button */}
                        <div className="p-2 border-t border-gray-100 bg-gray-50">
                          <button
                            disabled={isSubmitted}
                            onClick={() => handleFlexBulkSubmit(msg.id, isReplenish)}
                            className={`w-full font-bold text-xs py-2 rounded-xl transition-all shadow-sm active:scale-[0.98] flex items-center justify-center gap-1.5 cursor-pointer text-white ${
                              isSubmitted
                                ? 'bg-gray-400 cursor-not-allowed'
                                : isReplenish 
                                ? 'bg-[#06c755] hover:bg-green-600' 
                                : 'bg-rose-600 hover:bg-rose-700'
                            }`}
                          >
                            {isSubmitted ? (
                              <>
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                ส่งข้อมูลเรียบร้อยแล้ว
                              </>
                            ) : (
                              <>
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                ส่งข้อมูล{isReplenish ? 'เติมสต๊อก' : 'ยอดคงเหลือจริง'}
                              </>
                            )}
                          </button>
                        </div>
                      </div>
                    );
                  }

                  // Default rendering for other Flex Messages
                  return (
                    <div className="bg-white rounded-2xl text-gray-800 overflow-hidden shadow-lg border border-gray-100 rounded-tl-xs">
                      <div className="bg-[#1f2937] text-emerald-400 p-2.5 text-xs font-bold flex items-center gap-1.5">
                        <Bot className="w-4 h-4" />
                        {msg.flexContent?.title}
                      </div>
                      <div className="p-3 text-[11px] text-gray-600 whitespace-pre-line border-b border-gray-100">
                        {msg.flexContent?.description}
                      </div>
                      <div className="max-h-48 overflow-y-auto p-1 bg-gray-50/50 scrollbar-thin">
                        {msg.flexContent?.items?.map((item: any) => (
                          <button
                            key={item.code || item.name}
                            onClick={() => {
                              if (msg.id === messages[messages.length - 1].id) {
                                if (item.actionText && (msg.flexContent?.title?.includes('หมายเหตุ') || !item.unit)) {
                                  handleSendMessage(item.actionText);
                                } else {
                                  handleItemSelectInMenu(item);
                                }
                              }
                            }}
                            className="w-full text-left px-2.5 py-1.5 text-xs hover:bg-emerald-50 rounded-lg flex items-center justify-between border-b border-gray-100/50 transition-colors"
                          >
                            <span className="font-medium text-gray-700">{item.name}</span>
                            {item.unit ? (
                              <span className="text-[10px] text-gray-400 bg-gray-200/60 px-1.5 py-0.5 rounded font-mono flex items-center gap-1">
                                {item.unit}
                                <ChevronRight className="w-2.5 h-2.5 text-gray-400" />
                              </span>
                            ) : (
                              <span className="text-[10px] text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded font-bold flex items-center gap-1">
                                กดด่วน
                                <ChevronRight className="w-2.5 h-2.5 text-emerald-500" />
                              </span>
                            )}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })()}

                <span className="text-[9px] text-gray-100 mt-1 select-none font-mono text-right pr-1">
                  {msg.timestamp}
                </span>
              </div>
            </motion.div>
          ))}

          {/* Typing Indicator */}
          {isTyping && (
            <div className="flex items-start gap-1.5">
              <div className="w-7 h-7 rounded-full bg-[#06c755] flex items-center justify-center text-white shrink-0">
                <Bot className="w-4 h-4" />
              </div>
              <div className="bg-white text-gray-500 p-3 rounded-2xl text-xs rounded-tl-xs flex items-center gap-1 shadow-sm">
                <span className="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce"></span>
                <span className="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce delay-100"></span>
                <span className="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce delay-200"></span>
              </div>
            </div>
          )}
        </AnimatePresence>
        <div ref={chatEndRef} />
      </div>

      {/* Mini Sheet Panel for item inputs */}
      <AnimatePresence>
        {activeMenuTab !== 'none' && (
          <motion.div
            initial={{ y: 150, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 150, opacity: 0 }}
            className={`absolute bottom-[110px] left-4 right-4 bg-white text-gray-800 rounded-2xl p-3 shadow-2xl z-30 border max-h-[440px] flex flex-col ${
              activeMenuTab === 'replenish' ? 'border-green-100' : activeMenuTab === 'count' ? 'border-red-100' : activeMenuTab === 'note' ? 'border-amber-100' : 'border-blue-100'
            }`}
          >
            <div className="flex justify-between items-center mb-2.5 border-b pb-1.5 shrink-0">
              <div className="flex flex-col">
                <span className="font-bold text-xs flex items-center gap-1.5 text-emerald-700">
                  <span className={`w-2.5 h-2.5 rounded-full ${
                    activeMenuTab === 'replenish' ? 'bg-green-500' : activeMenuTab === 'count' ? 'bg-red-500' : activeMenuTab === 'note' ? 'bg-amber-500' : 'bg-blue-500'
                  }`}></span>
                  {activeMenuTab === 'replenish' 
                    ? 'บันทึกการเติมสต๊อก' 
                    : activeMenuTab === 'count' 
                    ? 'บันทึกสต๊อกคงเหลือจริง' 
                    : activeMenuTab === 'note'
                    ? 'บันทึกหมายเหตุพิเศษ (ชีส 0.25)'
                    : 'ดึงรายงานผลต่างสต๊อก'}
                </span>
                <div className="flex items-center gap-1.5 mt-1 ml-4 bg-gray-100 hover:bg-gray-200 px-2 py-0.5 rounded border border-gray-200 w-fit transition-colors">
                  <span className="text-[10px] text-gray-500 font-bold">เลือกวันที่ทำรายการ:</span>
                  <input
                    type="date"
                    value={selectedDate}
                    onChange={(e) => setSelectedDate(e.target.value)}
                    className="text-[10px] font-bold text-emerald-700 bg-transparent outline-none cursor-pointer border-none p-0 focus:ring-0"
                  />
                </div>
              </div>
              <button onClick={() => { setActiveMenuTab('none'); setSelectedItem(null); }} className="text-[11px] text-gray-400 hover:text-gray-600 self-start">ปิด</button>
            </div>

            {activeMenuTab === 'report' ? (
              <div className="space-y-3">
                <div className="bg-blue-50/60 border border-blue-100 p-2.5 rounded-xl">
                  <p className="text-[11px] font-semibold text-blue-950 leading-relaxed">
                    ระบบจะส่งรายงานสรุปยอดขาย การเติมสต๊อก และสต๊อกคงเหลือจริง เพื่อคำนวณหักลบส่วนต่างของวันที่เลือกด้านบน ให้คุณทางแชท
                  </p>
                </div>
                <button
                  onClick={handleConfirmActionInMenu}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs py-2 rounded-xl transition-all shadow-md active:scale-[0.98] flex items-center justify-center gap-1.5"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  ดึงรายงานของวันที่นี้
                </button>
              </div>
            ) : activeMenuTab === 'note' ? (
              <div className="space-y-3">
                <div className="flex bg-gray-100 p-1 rounded-xl">
                  <button
                    onClick={() => setNoteDirection('เพิ่ม')}
                    className={`flex-1 py-1.5 text-xs font-bold rounded-lg transition-all ${
                      noteDirection === 'เพิ่ม' ? 'bg-emerald-500 text-white shadow-sm' : 'text-gray-500 hover:text-gray-800'
                    }`}
                  >
                    เพิ่มชีส (+0.25)
                  </button>
                  <button
                    onClick={() => setNoteDirection('ลด')}
                    className={`flex-1 py-1.5 text-xs font-bold rounded-lg transition-all ${
                      noteDirection === 'ลด' ? 'bg-amber-500 text-white shadow-sm' : 'text-gray-500 hover:text-gray-800'
                    }`}
                  >
                    ลดชีส (-0.25)
                  </button>
                </div>
                
                <div className="flex justify-between items-center bg-amber-50 p-2 rounded-lg">
                  <span className="text-[11px] font-semibold text-amber-900">ระบุชื่อพิซซ่าหน้าครึ่งเพื่อเพิ่มหรือลดชีส 0.25 ยูนิต</span>
                </div>

                <div className="flex gap-2">
                  <input
                    type="text"
                    value={pizzaName}
                    onChange={(e) => setPizzaName(e.target.value)}
                    placeholder="พิมพ์ชื่อพิซซ่า..."
                    className="flex-1 bg-gray-50 border border-gray-200 rounded-xl px-3 py-1.5 text-xs outline-none focus:ring-2 focus:ring-amber-500"
                    autoFocus
                  />
                  <button
                    onClick={handleConfirmActionInMenu}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 rounded-xl transition-colors shrink-0"
                  >
                    บันทึก
                  </button>
                </div>
              </div>
            ) : selectedItem ? (
              <div className="space-y-3">
                <div className="flex justify-between items-center bg-emerald-50 p-2 rounded-lg">
                  <span className="text-xs font-semibold text-emerald-900">{selectedItem.nameThai}</span>
                  <span className="text-[10px] bg-emerald-200 text-emerald-800 px-1.5 py-0.5 rounded font-mono">หน่วย: {selectedItem.unit}</span>
                </div>
                <div className="flex gap-2">
                  <input
                    type="number"
                    step="any"
                    value={inputQty}
                    onChange={(e) => setInputQty(e.target.value)}
                    placeholder="ใส่จำนวนสินค้า..."
                    className="flex-1 bg-gray-50 border border-gray-200 rounded-xl px-3 py-1.5 text-xs outline-none focus:ring-2 focus:ring-emerald-500"
                    autoFocus
                  />
                  <button
                    onClick={handleConfirmActionInMenu}
                    className="bg-[#06c755] hover:bg-green-600 text-white font-bold text-xs px-4 rounded-xl transition-colors shrink-0"
                  >
                    ส่งข้อมูล
                  </button>
                </div>
              </div>
            ) : (
              <div className="space-y-1 max-h-32 overflow-y-auto scrollbar-thin">
                <p className="text-[10px] text-gray-400 mb-1">เลือกวัตถุดิบที่ต้องการบันทึก:</p>
                <div className="grid grid-cols-2 gap-1.5">
                  {STOCK_ITEMS_LIST.map((item) => (
                    <button
                      key={item.code}
                      onClick={() => handleItemSelectInMenu(item)}
                      className="text-left px-2 py-1 bg-gray-50 hover:bg-emerald-50 border border-gray-100 rounded-lg text-[11px] font-semibold text-gray-700 truncate"
                    >
                      {item.nameThai}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </motion.div>
        )}
      </AnimatePresence>

      {/* LINE Rich Menu (แถบ Rich Menu ตามโจทย์ต้องการ) */}
      <div className="bg-[#111115] border-t border-gray-800/80 -mx-4 -mb-4 mt-2 grid grid-cols-4 h-[90px] relative">
        <button
          onClick={() => handleRichMenuClick('replenish')}
          className={`flex flex-col items-center justify-center transition-colors border-r border-gray-800/50 ${
            activeMenuTab === 'replenish' ? 'bg-[#06c755]/10 text-[#06c755]' : 'text-gray-300 hover:text-white'
          }`}
        >
          <div className="w-1.5 h-1.5 rounded-full bg-green-500 mb-1.5"></div>
          <span className="text-[11px] font-bold tracking-tight">เติมสต๊อก</span>
          <span className="text-[8px] opacity-50 mt-0.5">INFLOW</span>
        </button>

        <button
          onClick={() => handleRichMenuClick('count')}
          className={`flex flex-col items-center justify-center transition-colors border-r border-gray-800/50 ${
            activeMenuTab === 'count' ? 'bg-red-500/10 text-red-400' : 'text-gray-300 hover:text-white'
          }`}
        >
          <div className="w-1.5 h-1.5 rounded-full bg-red-500 mb-1.5"></div>
          <span className="text-[11px] font-bold tracking-tight">คงเหลือ</span>
          <span className="text-[8px] opacity-50 mt-0.5">PHYSICAL</span>
        </button>

        <button
          onClick={() => handleRichMenuClick('note')}
          className={`flex flex-col items-center justify-center transition-colors border-r border-gray-800/50 ${
            activeMenuTab === 'note' ? 'bg-amber-500/10 text-amber-400' : 'text-gray-300 hover:text-white'
          }`}
        >
          <div className="w-1.5 h-1.5 rounded-full bg-amber-500 mb-1.5"></div>
          <span className="text-[11px] font-bold tracking-tight">หมายเหตุ</span>
          <span className="text-[8px] opacity-50 mt-0.5">REMARK</span>
        </button>

        <button
          onClick={() => handleRichMenuClick('report')}
          className={`flex flex-col items-center justify-center transition-colors ${
            activeMenuTab === 'report' ? 'bg-blue-500/10 text-blue-400' : 'text-gray-300 hover:text-white'
          }`}
        >
          <div className="w-1.5 h-1.5 rounded-full bg-blue-500 mb-1.5 animate-pulse"></div>
          <span className="text-[11px] font-bold tracking-tight">รายงาน</span>
          <span className="text-[8px] opacity-50 mt-0.5">REPORT</span>
        </button>
      </div>

      {/* Chat Input Field (At the very bottom, hidden slightly below rich menu but fully expandable) */}
      <div className="pt-2 flex items-center gap-1.5 bg-[#1e1e24] -mx-4 px-4 pb-2 border-t border-gray-800/50 z-10">
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileChange}
          accept=".xlsx,.xls"
          className="hidden"
        />
        <button
          onClick={handleFileUploadClick}
          className="p-2 bg-gray-800 hover:bg-gray-700 rounded-xl transition-colors border border-gray-700/50 flex items-center justify-center text-emerald-400"
          title="อัปโหลดไฟล์ Excel ยอดขาย"
        >
          <Upload className="w-4 h-4" />
        </button>

        <input
          type="text"
          value={inputText}
          onChange={(e) => setInputText(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleSendMessage(inputText)}
          placeholder="พิมพ์ข้อความคุยกับบอท..."
          className="flex-1 bg-gray-800/60 border border-gray-700/80 rounded-xl px-3 py-1.5 text-xs text-white outline-none focus:ring-1 focus:ring-[#06c755]/50 font-sans"
        />

        <button
          onClick={() => handleSendMessage(inputText)}
          className="p-2 bg-[#06c755] hover:bg-green-600 rounded-xl transition-colors flex items-center justify-center text-white shadow-md"
        >
          <Send className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
}
