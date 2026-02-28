import React, { useState, useCallback, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  ChevronRight, 
  ShoppingCart, 
  Trash2, 
  CheckCircle2, 
  Package,
  MapPin,
  X,
  Hospital,
  Calendar,
  User,
  MessageSquare,
  UserCheck,
  Mail,
  Clock,
  History
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { RegionData, MedicalSet, OrderItem, OrderMetadata, OrderHistoryItem } from './types';

// --- DATA VOOR DROPDOWNS ---
const ZIEKENHUIS_DATA: Record<string, { chirurgen: string[] }> = {
  "Jessa Hasselt": {
    chirurgen: ["Dr. Wissels", "Dr. Achahbar", "Dr. Put", "Dr. Roosen", "Dr. Bamps", "Dr. Vanvolsem", "Dr. Plazier", "Dr. Meeus"]
  },
  "St. Franciscus Heusden": {
    chirurgen: ["Dr. Vanvolsem", "Dr. Achahbar"]
  }
};

const TARGET_EMAILS = [
  "belgiumorders@globusmedical.com",
  "spelckmans@globusmedical.com",
  "jwalravens@globusmedical.com",
  "dirk@i-conic.be"
];

export default function App() {
  const [data, setData] = useState<RegionData[]>([]);
  const [selectedRegion, setSelectedRegion] = useState<string | null>(null);
  const [selectedSets, setSelectedSets] = useState<OrderItem[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [orderPlaced, setOrderPlaced] = useState(false);
  const [showConfirmModal, setShowConfirmModal] = useState(false);
  const [view, setView] = useState<'order' | 'history'>('order');
  const [selectedEmails, setSelectedEmails] = useState<string[]>(["belgiumorders@globusmedical.com"]);
  
  const [orderHistory, setOrderHistory] = useState<OrderHistoryItem[]>(() => {
    const saved = localStorage.getItem('medset_history');
    return saved ? JSON.parse(saved) : [];
  });

  const [metadata, setMetadata] = useState<OrderMetadata>({
    hospital: '',
    surgerydate: '',
    surgeon: '',
    agentName: 'Dirk',
    remarks: '',
    infoEmails: ''
  });

  // 1. DATA VERWERKEN
  const processExcelData = useCallback((workbook: XLSX.WorkBook) => {
    const regions: RegionData[] = workbook.SheetNames.map((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<any>(ws);
      const sets: MedicalSet[] = rows.map((row, index) => {
        const keys = Object.keys(row);
        return {
          id: `${sheetName}-${index}`,
          name: String(row[keys[0]] || `Set ${index + 1}`),
          code: row[keys[1]] ? String(row[keys[1]]) : '',
          description: row[keys[2]] ? String(row[keys[2]]) : '',
          region: sheetName
        };
      });
      return { name: sheetName, sets };
    });
    setData(regions);
    if (regions.length > 0) setSelectedRegion(regions[0].name);
  }, []);

  // 2. AUTOMATISCH LADEN
  useEffect(() => {
    const loadDefaultFile = async () => {
      try {
        const response = await fetch('/data.xlsx');
        if (response.ok) {
          const arrayBuffer = await response.arrayBuffer();
          const wb = XLSX.read(arrayBuffer, { type: 'buffer' });
          processExcelData(wb);
        }
      } catch (error) { console.error("Laadfout:", error); }
    };
    loadDefaultFile();
  }, [processExcelData]);

  // 3. HANDMATIGE UPLOAD & DRAG/DROP
  const handleFileUpload = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const bstr = e.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      processExcelData(wb);
    };
    reader.readAsBinaryString(file);
  }, [processExcelData]);

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      handleFileUpload(file);
    }
  };

  const toggleSet = (set: MedicalSet) => {
    setSelectedSets(prev => {
      const exists = prev.find(s => s.setId === set.id);
      if (exists) return prev.filter(s => s.setId !== set.id);
      return [...prev, { setId: set.id, setName: set.name, setCode: set.code, region: set.region }];
    });
  };

  const toggleEmail = (email: string) => {
    setSelectedEmails(prev => prev.includes(email) ? prev.filter(e => e !== email) : [...prev, email]);
  };

  const placeOrder = () => {
    setShowConfirmModal(false);
    setOrderPlaced(true);
    
    const newOrder: OrderHistoryItem = {
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      items: selectedSets,
      metadata: { ...metadata, infoEmails: selectedEmails.join(', ') }
    };

    const updatedHistory = [newOrder, ...orderHistory];
    setOrderHistory(updatedHistory);
    localStorage.setItem('medset_history', JSON.stringify(updatedHistory));

    const subject = `Bestelling Medische Sets - ${metadata.hospital} - ${metadata.date}`;
    const body = `Beste,\n\nHierbij een nieuwe bestelling voor medische sets.\n\nDETAILS INGREEP\n------------------\nHospitaal: ${metadata.hospital}\nDatum: ${metadata.date}\nChirurg: ${metadata.surgeon}\nAgent: ${metadata.agentName}\n\nBESTELDE SETS\n------------------\n${selectedSets.map(s => `- ${s.setName} (${s.setCode}) [${s.region}]`).join('\n')}\n\nOPMERKINGEN\n------------------\n${metadata.remarks || 'Geen'}`;

    window.location.href = `mailto:${selectedEmails.join(',')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    setTimeout(() => {
      setOrderPlaced(false);
      setSelectedSets([]);
      setView('history');
    }, 2000);
  };

  // --- RENDERING ---

  if (data.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6 bg-slate-50">
        <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-xl w-full">
          <div className="text-center mb-8">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-indigo-600 text-white mb-4 shadow-lg"><Package className="animate-bounce" size={32} /></div>
            <h1 className="text-3xl font-bold text-slate-900">MedSet Loaner Order</h1>
            <p className="text-slate-500 mt-2">Sleep een Excel bestand om te beginnen.</p>
          </div>
          <div onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }} onDragLeave={() => setIsDragging(false)} onDrop={onDrop} onClick={() => document.getElementById('fileInput')?.click()} className={cn("relative cursor-pointer border-2 border-dashed rounded-3xl p-12 transition-all text-center", isDragging ? "border-indigo-500 bg-indigo-50" : "border-slate-200 bg-white hover:border-indigo-400")}>
            <input id="fileInput" type="file" className="hidden" accept=".xlsx, .xls" onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0])} />
            <Upload className="mx-auto text-slate-400 mb-4" size={48} />
            <p className="text-lg font-semibold text-slate-900">Klik of sleep bestand</p>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col lg:flex-row font-sans">
      <aside className="w-full lg:w-72 bg-white border-r p-6 shrink-0">
        <div className="flex items-center gap-3 text-indigo-600 mb-8 cursor-pointer" onClick={() => setView('order')}><Package size={24} /><span className="font-bold text-xl text-slate-900 tracking-tight">MedSet</span></div>
        <nav className="space-y-1">
          <button onClick={() => setView('history')} className={cn("w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium mb-4", view === 'history' ? "bg-indigo-50 text-indigo-700" : "text-slate-600 hover:bg-slate-50")}><History size={18}/> Geschiedenis</button>
          <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4 ml-4">Regio's</h2>
          {data.map(r => (
            <button key={r.name} onClick={() => { setSelectedRegion(r.name); setView('order'); }} className={cn("w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-medium", view === 'order' && selectedRegion === r.name ? "bg-indigo-50 text-indigo-700" : "text-slate-600 hover:bg-slate-50")}>
              <div className="flex items-center gap-3"><MapPin size={18}/> {r.name}</div>
              <span className="text-xs bg-slate-100 px-2 py-0.5 rounded-full">{r.sets.length}</span>
            </button>
          ))}
        </nav>
      </aside>

      <main className="flex-1 p-6 lg:p-10 overflow-auto">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-3xl font-bold text-slate-900 mb-8">{view === 'history' ? 'Bestelgeschiedenis' : selectedRegion}</h1>
          {view === 'order' ? (
            <div className="grid gap-3">
              {data.find(r => r.name === selectedRegion)?.sets.map(set => (
                <div key={set.id} onClick={() => toggleSet(set)} className={cn("p-4 rounded-2xl border cursor-pointer transition-all flex items-center justify-between", selectedSets.some(s => s.setId === set.id) ? "border-indigo-500 bg-white shadow-md ring-2 ring-indigo-500/10" : "bg-white border-slate-200")}>
                  <div><h3 className="font-bold text-slate-800">{set.name}</h3><p className="text-sm text-slate-500">{set.code}</p></div>
                  {selectedSets.some(s => s.setId === set.id) && <CheckCircle2 className="text-indigo-600" size={20} />}
                </div>
              ))}
            </div>
          ) : (
            <div className="space-y-4">
              {orderHistory.map(order => (
                <div key={order.id} className="p-6 bg-white rounded-2xl border border-slate-200 shadow-sm">
                  <div className="flex justify-between items-start mb-4">
                    <div><p className="font-bold text-slate-900">{new Date(order.timestamp).toLocaleString()}</p><p className="text-sm text-indigo-600">{order.metadata.hospital}</p></div>
                    <span className="px-3 py-1 bg-emerald-100 text-emerald-700 rounded-full text-xs font-bold uppercase">Verzonden</span>
                  </div>
                  <div className="text-sm text-slate-600 space-y-1">{order.items.map((it, i) => <p key={i}>• {it.setName} ({it.region})</p>)}</div>
                </div>
              ))}
            </div>
          )}
        </div>
      </main>

      <aside className="w-full lg:w-96 bg-white border-l p-6 flex flex-col shrink-0">
        <h2 className="font-bold text-lg mb-6 flex items-center gap-2"><ShoppingCart size={20} /> Bestelling</h2>
        <div className="flex-1 space-y-2 overflow-y-auto">
          {selectedSets.map(item => (
            <div key={item.setId} className="p-3 bg-slate-50 rounded-xl flex justify-between items-center text-sm border">
              <span className="truncate mr-2 font-medium">{item.setName}</span>
              <button onClick={() => setSelectedSets(prev => prev.filter(s => s.setId !== item.setId))} className="text-slate-400 hover:text-red-500"><Trash2 size={16}/></button>
            </div>
          ))}
        </div>
        <button disabled={selectedSets.length === 0 || orderPlaced} onClick={() => setShowConfirmModal(true)} className="w-full py-4 mt-6 rounded-2xl font-bold text-white bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 shadow-lg transition-transform active:scale-95">Bestelling Afronden</button>
      </aside>

      {/* MODAL MET DROPDOWNS & CHECKBOXES */}
      <AnimatePresence>
        {showConfirmModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowConfirmModal(false)} />
            <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} exit={{ opacity: 0, scale: 0.95 }} className="relative w-full max-w-lg bg-white rounded-3xl p-8 shadow-2xl max-h-[90vh] overflow-y-auto">
              <h3 className="text-2xl font-bold mb-6 text-slate-900">Bestelgegevens</h3>
              <div className="space-y-4">
                <div>
                  <label className="text-xs font-bold text-slate-400 uppercase flex items-center gap-2 mb-1"><Hospital size={14}/> Hospitaal</label>
                  <select className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500 transition-all" value={metadata.hospital} onChange={e => setMetadata(prev => ({ ...prev, hospital: e.target.value, surgeon: '' }))}>
                    <option value="">-- Kies een ziekenhuis --</option>
                    {Object.keys(ZIEKENHUIS_DATA).map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-400 uppercase flex items-center gap-2 mb-1"><User size={14}/> Chirurg</label>
                  <select disabled={!metadata.hospital} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500 disabled:opacity-50 transition-all" value={metadata.surgeon} onChange={e => setMetadata(prev => ({ ...prev, surgeon: e.target.value }))}>
                    <option value="">-- Kies een chirurg --</option>
                    {metadata.hospital && ZIEKENHUIS_DATA[metadata.hospital].chirurgen.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase flex items-center gap-2 mb-1"><Calendar size={14}/> Datum</label>
                    <input type="date" className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500" onChange={e => setMetadata(prev => ({ ...prev, date: e.target.value }))} />
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase flex items-center gap-2 mb-1"><UserCheck size={14}/> Uw Naam</label>
                    <input type="text" placeholder="Naam agent" className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500" onChange={e => setMetadata(prev => ({ ...prev, agentName: e.target.value }))} />
                  </div>
                </div>

                <div className="pt-4 border-t">
                  <label className="text-xs font-bold text-slate-400 uppercase flex items-center gap-2 mb-3"><Mail size={14} /> Kopie versturen naar:</label>
                  <div className="space-y-2">
                    {TARGET_EMAILS.map(email => (
                      <label key={email} className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl border border-slate-100 cursor-pointer hover:bg-slate-100 transition-colors">
                        <input type="checkbox" checked={selectedEmails.includes(email)} onChange={() => toggleEmail(email)} className="w-5 h-5 accent-indigo-600 rounded" />
                        <span className="text-sm font-medium text-slate-700">{email}</span>
                      </label>
                    ))}
                  </div>
                </div>
              </div>
              <div className="flex gap-4 mt-8">
                <button onClick={() => setShowConfirmModal(false)} className="flex-1 py-3 font-bold text-slate-400 hover:text-slate-600 transition-colors">Annuleren</button>
                <button disabled={!metadata.hospital || !metadata.surgeon || !metadata.date || selectedEmails.length === 0} onClick={placeOrder} className="flex-1 py-3 font-bold text-white bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:bg-slate-200 shadow-md">Verstuur Bestelling</button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
