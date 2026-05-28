import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  AlertTriangle,
  CheckCircle,
  Truck,
  EyeOff,
  Search,
  RefreshCcw,
  Package,
  TrendingDown,
  Plus,
  Factory,
  Calendar,
  Activity,
  LayoutDashboard,
  AlertCircle,
  Clock,
  Sparkles,
  Loader2,
  X,
  Target,
  FileUp,
  FileSpreadsheet,
  Check
} from 'lucide-react';

// הזרקת ספריית XLSX לטובת עיבוד קבצי אקסל
const loadXLSX = () => {
  return new Promise((resolve) => {
    if (window.XLSX) return resolve(window.XLSX);
    const script = document.createElement('script');
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
    script.onload = () => resolve(window.XLSX);
    document.head.appendChild(script);
  });
};

const PRODUCTION_DATES = [
  { id: '2026-05-14', label: 'חמישי 14.5', target: 'לשישי' },
  { id: '2026-05-15', label: 'שישי 15.5', target: 'לראשון' },
  { id: '2026-05-17', label: 'ראשון 17.5', target: 'לשני' },
  { id: '2026-05-18', label: 'שני 18.5', target: 'לשלישי' },
  { id: '2026-05-19', label: 'שלישי 19.5', target: 'לרביעי' }
];

const INITIAL_DATA = [
  { sku: "37575", name: "לזניה קלאסית בינונית", stock: 74, orders: 216, category: "לזניות" },
  { sku: "12933", name: "ספנקופיטה אישי", stock: 5, orders: 70, category: "מאפים" },
  { sku: "12934", name: "ספנקופיטה", stock: 58, orders: 112, category: "מאפים" },
  { sku: "36028", name: "קיש פטריות בינוני טבעוני ללא קמח", stock: 7, orders: 38, category: "קישים" },
  { sku: "37551", name: "לזניה חצילים פארמה בינוני", stock: 92, orders: 108, category: "לזניות" },
  { sku: "37125", name: "פשטידה עגולה פטריות ללא קמח בינונית", stock: 55, orders: 72, category: "פשטידות" },
  { sku: "36300", name: "קיש בצל אישי", stock: 640, orders: 240, category: "קישים" },
  { sku: "36400", name: "קיש ברוקולי אישי", stock: 1024, orders: 264, category: "קישים" },
  { sku: "36128", name: "קיש בטטה בינוני טבעוני ללא קמח", stock: 70, orders: 38, category: "קישים" },
  { sku: "12913", name: "מאפה פילו יווני אירועים", stock: 33, orders: 7, category: "מאפים" },
  { sku: "36201", name: "קיש תרד ומנגולד בינוני", stock: 143, orders: 117, category: "קישים" },
  { sku: "36402", name: "קיש ברוקולי גדול", stock: 253, orders: 121, category: "קישים" },
  { sku: "36501", name: "קיש שרי בינוני", stock: 131, orders: 56, category: "קישים" },
  { sku: "36702", name: "קיש גבינות וזיתים גדול", stock: 83, orders: 76, category: "קישים" },
  { sku: "36851", name: "קיש זוקיני ומנצגו בינוני", stock: 320, orders: 40, category: "קישים" },
  { sku: "368511", name: "קיש זוקיני ומנצגו בינוני (מכסה)", stock: 346, orders: 20, category: "קישים" },
  { sku: "375701", name: "לזניה קלאסית (8)", stock: 18, orders: 10, category: "לזניות" },
  { sku: "37521", name: "לזניה תרד בינונית", stock: 0, orders: 42, category: "לזניות" },
  { sku: "37561", name: "לזניה שמנת פטריות בינונית", stock: 0, orders: 35, category: "לזניות" },
];

const REPORTED_SHIPMENTS = {
  "37575": 146, "12933": 225, "12934": 100, "36028": 60, "37551": 30, "37125": 60, "12913": 55, "36128": 43, "36300": 640
};

const INITIAL_PRODUCTION_PLAN = {
  "2026-05-14": {
    "36201": 60, "36402": 30, "36501": 60, "36702": 60, "36851": 60, "368511": 30, "375701": 60,
    "37551": 24, "37575": 24, "12934": 24, "12933": 24, "37521": 10, "37561": 12
  },
  "2026-05-15": { "37575": 100, "37551": 50, "12933": 65, "12934": 55, "36028": 35 },
  "2026-05-17": {}, "2026-05-18": {},
  "2026-05-19": { "37575": 60, "37551": 40 }
};

export default function App() {
  const [items, setItems] = useState(INITIAL_DATA);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [searchTerm, setSearchTerm] = useState('');
  const [manualShipments, setManualShipments] = useState(REPORTED_SHIPMENTS);
  const [productionPlan, setProductionPlan] = useState(INITIAL_PRODUCTION_PLAN);
  const [ignoredSkus, setIgnoredSkus] = useState([]);
  const [uploadStatus, setUploadStatus] = useState(null); // 'loading', 'success', 'error'

  const [aiInsights, setAiInsights] = useState(null);
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [isAiModalOpen, setIsAiModalOpen] = useState(false);
  const [aiError, setAiError] = useState(null);

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [activeItem, setActiveItem] = useState(null);
  const [activeDate, setActiveDate] = useState(null);
  const [modalType, setModalType] = useState('shipment');
  const [modalValue, setModalValue] = useState('');

  const fileInputRef = useRef(null);

  // 1. חישוב נתונים מעובדים
  const processedData = useMemo(() => {
    return items.map(item => {
      const manual = manualShipments[item.sku] || 0;
      const totalPlanned = Object.values(productionPlan).reduce((acc, dayPlan) => acc + (dayPlan[item.sku] || 0), 0);
      const realBalance = (item.stock || 0) - (item.orders || 0);
      const forecastBalance = (item.stock || 0) + manual + totalPlanned - (item.orders || 0);
      const dailyProduction = PRODUCTION_DATES.reduce((acc, date) => {
        acc[date.id] = productionPlan[date.id]?.[item.sku] || 0;
        return acc;
      }, {});
      return { ...item, manual, totalPlanned, realBalance, forecastBalance, dailyProduction };
    }).filter(item => !ignoredSkus.includes(item.sku));
  }, [items, manualShipments, productionPlan, ignoredSkus]);

  // 2. חישוב סטטיסטיקות
  const stats = useMemo(() => {
    return {
      realShortageCount: processedData.filter(i => i.realBalance < 0).length,
      realTotalMissingUnits: processedData.reduce((acc, i) => acc + (i.realBalance < 0 ? Math.abs(i.realBalance) : 0), 0),
      totalPlannedUnits: Object.values(productionPlan).reduce((acc, day) => acc + Object.values(day).reduce((dAcc, qty) => dAcc + qty, 0), 0),
      totalShippedToday: Object.values(manualShipments).reduce((acc, val) => acc + val, 0)
    };
  }, [processedData, productionPlan, manualShipments]);

  const urgentTasks = useMemo(() => {
    return [...processedData]
      .filter(i => i.realBalance < 0)
      .sort((a, b) => a.realBalance - b.realBalance)
      .slice(0, 5);
  }, [processedData]);

  const filteredData = useMemo(() => {
    return processedData.filter(item => {
      const matchesSearch = (item.name || "").includes(searchTerm) || (item.sku || "").includes(searchTerm);
      if (activeTab === 'shortages') return matchesSearch && item.realBalance < 0;
      return matchesSearch;
    });
  }, [processedData, searchTerm, activeTab]);

  // --- Excel Import Logic ---
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploadStatus('loading');
    const XLSX = await loadXLSX();

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) throw new Error("File is empty");

        setItems(prevItems => {
          const newItems = [...prevItems];

          data.forEach(row => {
            const sku = String(row["מק'ט"] || row["מק\"ט"] || "");
            if (!sku) return;

            const existingIdx = newItems.findIndex(i => i.sku === sku);

            // זיהוי אם זה קובץ מלאי או הזמנות
            const stockVal = row["מלאי מרלוג ביח'"] || row["מלאי"];
            const orderVal = row["כמות בהזמנות"] || row["הזמנות"];

            if (existingIdx > -1) {
              if (stockVal !== undefined) newItems[existingIdx].stock = parseInt(stockVal);
              if (orderVal !== undefined) newItems[existingIdx].orders = parseInt(orderVal);
            } else {
              // מוצר חדש שלא היה ברשימה
              newItems.push({
                sku: sku,
                name: row["תאור מוצר"] || row["שם מוצר"] || "מוצר חדש",
                stock: stockVal !== undefined ? parseInt(stockVal) : 0,
                orders: orderVal !== undefined ? parseInt(orderVal) : 0,
                category: "כללי"
              });
            }
          });
          return newItems;
        });

        setUploadStatus('success');
        setTimeout(() => setUploadStatus(null), 3000);
      } catch (err) {
        console.error(err);
        setUploadStatus('error');
        setTimeout(() => setUploadStatus(null), 3000);
      }
    };
    reader.readAsBinaryString(file);
  };

  // AI Logic
  const callGemini = async (prompt) => {
    const apiKey = "";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`;
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          systemInstruction: { parts: [{ text: "אתה מנהל תפעול מזון. נתח דחיפויות ייצור לגלעד בעברית בסגנון נקי ומקצועי." }] }
        })
      });
      const data = await response.json();
      return data.candidates?.[0]?.content?.parts?.[0]?.text || "לא התקבלה תשובה מהשרת.";
    } catch (err) { throw err; }
  };

  const getAiInsights = async () => {
    setIsAiLoading(true);
    setAiError(null);
    setIsAiModalOpen(true);
    const summary = processedData.map(i => ({ name: i.name, sku: i.sku, gap: i.realBalance, forecast: i.forecastBalance }));
    const prompt = `נתח את המחסן של גלעד. מה הכי דחוף לייצור מחר (שישי) לאספקות ראשון? נתונים: ${JSON.stringify(summary)}`;
    try {
      const text = await callGemini(prompt);
      setAiInsights(text);
    } catch (err) {
        setAiError("תקלה בחיבור ל-AI. נסה שוב מאוחר יותר.");
    } finally {
        setIsAiLoading(false);
    }
  };

  const handleUpdateValue = () => {
    if (activeItem) {
      const val = parseInt(modalValue || 0);
      if (modalType === 'shipment') {
        setManualShipments(prev => ({ ...prev, [activeItem.sku]: (prev[activeItem.sku] || 0) + val }));
      } else if (modalType === 'production' && activeDate) {
        setProductionPlan(prev => ({
          ...prev,
          [activeDate]: { ...(prev[activeDate] || {}), [activeItem.sku]: val }
        }));
      }
      setIsModalOpen(false);
      setModalValue('');
    }
  };

  const toggleIgnore = (sku) => {
    setIgnoredSkus(prev => prev.includes(sku) ? prev.filter(s => s !== sku) : [...prev, sku]);
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] font-sans text-right flex text-[#202124] antialiased" dir="rtl">

      {/* Sidebar - Google Style */}
      <aside className="w-64 bg-white border-e border-[#DADCE0] flex flex-col sticky top-0 h-screen z-30 transition-all">
        <div className="p-6">
          <div className="flex items-center gap-3 mb-8 px-2">
            <div className="bg-[#4285F4] p-2 rounded-lg text-white">
              <Package size={20} />
            </div>
            <span className="text-xl font-medium tracking-tight text-[#5F6368]">גלעד WMS</span>
          </div>

          <nav className="space-y-1">
            <button onClick={() => setActiveTab('dashboard')} className={`w-full flex items-center gap-4 px-4 py-3 rounded-full transition-colors ${activeTab === 'dashboard' ? 'bg-[#E8F0FE] text-[#1967D2]' : 'text-[#3C4043] hover:bg-[#F1F3F4]'}`}>
              <LayoutDashboard size={20} />
              <span className="font-medium text-sm">מרכז בקרה</span>
            </button>
            <button onClick={() => setActiveTab('production')} className={`w-full flex items-center gap-4 px-4 py-3 rounded-full transition-colors ${activeTab === 'production' ? 'bg-[#E8F0FE] text-[#1967D2]' : 'text-[#3C4043] hover:bg-[#F1F3F4]'}`}>
              <Factory size={20} />
              <span className="font-medium text-sm">תוכנית ייצור</span>
            </button>
            <button onClick={() => setActiveTab('shortages')} className={`w-full flex items-center gap-4 px-4 py-3 rounded-full transition-colors ${activeTab === 'shortages' ? 'bg-[#FCE8E6] text-[#D93025]' : 'text-[#3C4043] hover:bg-[#F1F3F4]'}`}>
              <AlertCircle size={20} />
              <span className="font-medium text-sm">חוסרים נטו</span>
            </button>

            <div className="pt-6 px-4 pb-2 text-[11px] font-bold text-[#70757A] uppercase tracking-wider">ניהול נתונים</div>

            {/* Excel Upload Button */}
            <button
              onClick={() => fileInputRef.current.click()}
              className={`w-full flex items-center gap-4 px-4 py-3 rounded-full transition-all border border-[#DADCE0] hover:bg-[#F1F3F4] text-[#5F6368] ${uploadStatus === 'loading' ? 'animate-pulse' : ''}`}
            >
              {uploadStatus === 'loading' ? <Loader2 size={18} className="animate-spin text-[#4285F4]" /> :
               uploadStatus === 'success' ? <Check size={18} className="text-[#1E8E3E]" /> :
               uploadStatus === 'error' ? <AlertTriangle size={18} className="text-[#D93025]" /> :
               <FileSpreadsheet size={18} className="text-[#1E8E3E]" />}
              <span className="font-medium text-sm">
                {uploadStatus === 'loading' ? 'מעבד...' : 'ייבוא אקסל'}
              </span>
            </button>
            <input
              type="file"
              ref={fileInputRef}
              className="hidden"
              accept=".xlsx, .xls, .csv"
              onChange={handleFileUpload}
            />

            <button onClick={getAiInsights} className="w-full mt-4 flex items-center gap-4 px-4 py-3 rounded-full text-[#1A73E8] hover:bg-[#F1F3F4] transition-colors border border-[#DADCE0]">
              <Sparkles size={18} />
              <span className="font-medium text-sm">ניתוח חכם</span>
            </button>
          </nav>
        </div>

        <div className="mt-auto p-6 text-center">
          <button onClick={() => {setManualShipments({}); setProductionPlan(INITIAL_PRODUCTION_PLAN); setItems(INITIAL_DATA);}} className="inline-flex items-center gap-2 text-[11px] text-[#70757A] hover:text-[#202124] transition-colors">
            <RefreshCcw size={12} />
            <span>איפוס למצב התחלתי</span>
          </button>
        </div>
      </aside>

      <main className="flex-1 p-10 overflow-y-auto">
        <header className="mb-10 flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-normal text-[#202124]">
              {activeTab === 'dashboard' ? 'מרכז בקרה' : activeTab === 'production' ? 'תוכנית ייצור שבועי' : 'חוסרים במחסן'}
            </h1>
            <div className="flex items-center gap-2 mt-1 text-[#70757A] text-sm">
                <Clock size={14} />
                <span>מדיניות: מוכן יום לפני | {uploadStatus === 'success' ? 'נתונים עודכנו מקובץ' : 'מידע: GAP-8'}</span>
            </div>
          </div>
          <div className="relative w-80">
            <Search className="absolute right-4 top-1/2 -translate-y-1/2 text-[#5F6368]" size={18} />
            <input type="text" placeholder='חפש מוצרים...' className="w-full pr-12 pl-4 py-2.5 bg-white border border-[#DADCE0] rounded-lg focus:border-[#4285F4] focus:ring-1 focus:ring-[#4285F4] outline-none text-sm transition-all shadow-sm" value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
          </div>
        </header>

        {/* Urgent Tasks Panel */}
        {activeTab === 'dashboard' && (
          <div className="mb-8 bg-white border border-[#DADCE0] rounded-xl overflow-hidden shadow-sm">
            <div className="px-6 py-4 border-b border-[#DADCE0] bg-[#F8F9FA] flex justify-between items-center">
                <h2 className="text-sm font-bold text-[#3C4043] flex items-center gap-2 uppercase tracking-wide">
                    <Target size={16} className="text-[#D93025]" />
                    ייצור מומלץ למחר (שישי)
                </h2>
                <span className="text-xs text-[#70757A]">מיקוד בחוסר נטו עמוק</span>
            </div>
            <div className="divide-y divide-[#DADCE0]">
                {urgentTasks.map(item => (
                    <div key={item.sku} className="px-6 py-4 flex items-center justify-between hover:bg-[#F1F3F4]/50 transition-colors">
                        <div className="flex items-center gap-4">
                            <div className="w-10 h-10 bg-[#FCE8E6] text-[#D93025] rounded-full flex items-center justify-center font-bold text-sm">
                                {Math.abs(item.realBalance)}
                            </div>
                            <div>
                                <h4 className="text-sm font-medium text-[#202124]">{item.name}</h4>
                                <p className="text-[11px] text-[#70757A]">מק"ט: {item.sku}</p>
                            </div>
                        </div>
                        <button onClick={() => {setActiveItem(item); setModalType('production'); setActiveDate('2026-05-15'); setIsModalOpen(true);}} className="text-[#1A73E8] text-xs font-bold hover:underline">
                            הוסף לתוכנית
                        </button>
                    </div>
                ))}
            </div>
          </div>
        )}

        {/* Stats Grid */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-10">
          {[
            { label: 'חוסר נטו במחסן', val: stats.realTotalMissingUnits, icon: Activity, color: '#D93025' },
            { label: 'נשלח היום', val: stats.totalShippedToday, icon: Truck, color: '#1E8E3E' },
            { label: 'סה"כ בייצור', val: stats.totalPlannedUnits, icon: Factory, color: '#1A73E8' },
            { label: 'מוצרים בחוסר', val: stats.realShortageCount, icon: Package, color: '#5F6368' }
          ].map((stat, i) => (
            <div key={i} className="bg-white p-6 rounded-xl border border-[#DADCE0] shadow-sm flex flex-col">
              <div className="flex items-center justify-between mb-4">
                 <span className="text-[11px] font-bold text-[#70757A] uppercase tracking-wider">{stat.label}</span>
                 <stat.icon size={18} style={{ color: stat.color }} />
              </div>
              <h3 className="text-3xl font-normal text-[#202124] tracking-tight">{stat.val.toLocaleString()}</h3>
            </div>
          ))}
        </div>

        {/* Main Table */}
        <div className="bg-white border border-[#DADCE0] rounded-xl overflow-hidden shadow-sm mb-20">
          <div className="overflow-x-auto">
            <table className="w-full text-right border-collapse table-auto text-sm">
              <thead className="bg-[#F8F9FA] border-b border-[#DADCE0]">
                <tr>
                  <th className="px-6 py-4 font-medium text-[#5F6368] text-[12px] sticky right-0 bg-[#F8F9FA] z-10">מוצר</th>
                  <th className="px-4 py-4 font-medium text-[#5F6368] text-[12px] text-center">WMS</th>
                  <th className="px-4 py-4 font-medium text-[#1E8E3E] text-[12px] text-center">נשלח</th>
                  <th className="px-4 py-4 font-medium text-[#D93025] text-[12px] text-center">חוסר</th>

                  {activeTab === 'production' && PRODUCTION_DATES.map(date => (
                    <th key={date.id} className="px-2 py-4 text-center border-s border-[#DADCE0]">
                      <div className="text-[11px] font-bold text-[#1A73E8]">{date.label.split(' ')[0]}</div>
                      <div className="text-[10px] text-[#70757A] font-normal">{date.target}</div>
                    </th>
                  ))}

                  <th className="px-6 py-4 font-bold text-[#202124] text-center bg-[#F1F3F4]">יתרה צפויה</th>
                  <th className="px-6 py-4 text-center"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#DADCE0]">
                {filteredData.map((item) => (
                  <tr key={item.sku} className="hover:bg-[#F8F9FA] transition-colors group">
                    <td className="px-6 py-4 sticky right-0 bg-white group-hover:bg-[#F8F9FA] z-10">
                        <div className="font-medium text-[#202124]">{item.name}</div>
                        <div className="text-[11px] text-[#70757A]">#{item.sku}</div>
                    </td>
                    <td className="px-4 py-4 text-center text-[#5F6368]">{item.stock}</td>
                    <td className="px-4 py-4 text-center text-[#1E8E3E] font-medium">{item.manual > 0 ? `+${item.manual}` : '-'}</td>
                    <td className="px-4 py-4 text-center font-medium text-[#D93025]">{item.realBalance < 0 ? item.realBalance : '0'}</td>

                    {activeTab === 'production' && PRODUCTION_DATES.map(date => (
                      <td key={date.id} className={`px-2 py-4 text-center border-s border-[#DADCE0]/50 cursor-pointer ${item.dailyProduction[date.id] > 0 ? 'bg-[#E8F0FE]/30' : ''}`}
                        onClick={() => {
                          setActiveItem(item); setActiveDate(date.id); setModalType('production');
                          setModalValue(item.dailyProduction[date.id] || ''); setIsModalOpen(true);
                        }}>
                        <span className={item.dailyProduction[date.id] > 0 ? 'font-bold text-[#1A73E8]' : 'text-[#DADCE0]'}>
                            {item.dailyProduction[date.id] || '0'}
                        </span>
                      </td>
                    ))}

                    <td className="px-6 py-4 text-center bg-[#F1F3F4]/30">
                      <span className={`px-3 py-1 rounded-full text-[11px] font-bold ${
                        item.forecastBalance < 0 ? 'bg-[#FCE8E6] text-[#D93025]' : 'bg-[#E6F4EA] text-[#1E8E3E]'
                      }`}>
                        {item.forecastBalance > 0 ? `+${item.forecastBalance}` : item.forecastBalance}
                      </span>
                    </td>

                    <td className="px-6 py-4 text-center">
                      <div className="flex justify-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button onClick={() => {setActiveItem(item); setModalType('shipment'); setIsModalOpen(true);}} className="p-1.5 text-[#5F6368] hover:bg-[#F1F3F4] rounded-md transition-colors" title="עדכן משלוח">
                          <Truck size={16} />
                        </button>
                        <button onClick={() => toggleIgnore(item.sku)} className="p-1.5 text-slate-300 hover:text-slate-600 transition-colors">
                          <EyeOff size={16} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </main>

      {/* AI Modal */}
      {isAiModalOpen && (
        <div className="fixed inset-0 bg-[#202124]/40 backdrop-blur-sm flex items-center justify-center z-[120] p-6 text-right" dir="rtl">
          <div className="bg-white rounded-2xl p-10 w-full max-w-2xl shadow-2xl scale-in">
            <div className="flex items-center gap-4 mb-8">
              <Sparkles className="text-[#4285F4]" size={32} />
              <h2 className="text-xl font-medium text-[#202124]">תובנות ניהול חכמות</h2>
            </div>
            <div className="bg-[#F8F9FA] rounded-xl p-8 min-h-[300px] max-h-[500px] overflow-auto border border-[#DADCE0] text-[#3C4043] leading-relaxed text-sm">
               {isAiLoading ? (
                 <div className="flex flex-col items-center justify-center h-full gap-4 text-[#4285F4]">
                    <Loader2 size={40} className="animate-spin" />
                    <p className="font-medium">Gemini מנתח את נתוני המחסן שלך...</p>
                 </div>
               ) : aiError ? (
                 <p className="text-[#D93025]">{aiError}</p>
               ) : (
                 <div className="whitespace-pre-wrap">{String(aiInsights)}</div>
               )}
            </div>
            <div className="mt-8 flex justify-end">
              <button onClick={() => setIsAiModalOpen(false)} className="px-6 py-2 bg-[#1A73E8] text-white rounded-md text-sm font-medium hover:bg-[#1557B0] transition-colors shadow-md">סגור</button>
            </div>
          </div>
        </div>
      )}

      {/* Update Dialog */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-[#202124]/40 backdrop-blur-xs flex items-center justify-center z-[110] p-4 text-right" dir="rtl">
          <div className="bg-white rounded-xl p-8 w-full max-md shadow-2xl border border-[#DADCE0]">
            <h2 className="text-lg font-medium text-[#202124] mb-2">
              {modalType === 'shipment' ? 'עדכון משלוח' : 'עדכון תוכנית ייצור'}
            </h2>
            <p className="text-xs text-[#70757A] mb-8">{activeItem?.name}</p>

            <div className="mb-10">
                <input
                  type="number" autoFocus
                  className="w-full p-4 bg-[#F8F9FA] border-b-2 border-[#1A73E8] rounded-t-md text-4xl font-light text-center focus:bg-[#F1F3F4] outline-none transition-all"
                  value={modalValue} placeholder="0"
                  onChange={(e) => setModalValue(e.target.value)}
                  onKeyPress={(e) => e.key === 'Enter' && handleUpdateValue()}
                />
            </div>

            <div className="flex gap-2 justify-end">
              <button onClick={() => setIsModalOpen(false)} className="px-4 py-2 text-[#5F6368] hover:bg-[#F1F3F4] rounded-md text-sm font-medium transition-colors">ביטול</button>
              <button onClick={handleUpdateValue} className="px-6 py-2 bg-[#1A73E8] text-white rounded-md text-sm font-medium hover:bg-[#1557B0] transition-colors shadow-sm">עדכן</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
