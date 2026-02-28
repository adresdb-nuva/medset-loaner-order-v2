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

// --- CONFIGURATIE DATA ---
const ZIEKENHUIS_DATA: Record<string, { chirurgen: string[] }> = {
  "Jessa Hasselt": {
    chirurgen: ["Dr. Wissels", "Dr. Achahbar", "Dr. Put", "Dr. Roosen", "Dr. Bamps", "Dr. Vanvolsem", "Dr. Plazier", "Dr. Meeus", "andere"]
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
    date: '',
    surgeon: '',
    agentName: '',
    remarks: '',
    infoEmails: ''
  });

  // 1. EXCEL DATA VERWERKEN
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

  // 2. AUTOMATISCH LADEN VAN DATA.XLSX
  useEffect(() => {
    const loadDefaultFile = async () => {
      try {
        const response = await fetch('/data.xlsx');
        if (response.ok) {
          const arrayBuffer = await response.arrayBuffer();
          const wb = XLSX.read(arrayBuffer, { type: 'buffer' });
          processExcelData(wb);
        }
      } catch (error) { console.error("Standaardbestand niet gevonden."); }
    };
    loadDefaultFile();
  }, [processExcelData]);

  // 3. UPLOAD & DRAG/DROP LOGICA
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

  // 4. BESTELLING FINALISEREN (MAILTO)
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
    const body = `Beste,

Hierbij een nieuwe bestelling voor medische sets.

DETAILS INGREEP
------------------------------------------------
Ziekenhuis: ${metadata.hospital}
OKa-dag: ${metadata.date}
Chirurg: ${metadata.surgeon}
Agent: ${metadata.agentName}

BESTELDE SETS
------------------------------------------------
${selectedSets.map(s => `- ${s.setName} [Code: ${s.setCode}] (${s.region})`).join('\n')}

OPMERKINGEN
------------------------------------------------
${metadata.remarks || 'Geen opmerkingen'}

Met vriendelijke groet,
${metadata.agentName}`;

    window.location.href = `mailto:${selectedEmails.join(',')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    setTimeout(() => {
      setOrderPlaced(false);
      setSelectedSets([]);
      setView('history');
    }, 2000);
  };

  // --- INTERFACE RENDERING ---

  if (data.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6 bg-slate-50 font-sans">
        <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-xl w-full">
          <div className="text-center mb-8">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-indigo-600 text-white mb-4 shadow-lg shadow-indigo-100">
              <Package className="animate-bounce" size={32} />
            </div>
            <h1 className="text-3xl font-bold text-slate-900 tracking-tight">MedSet Loaner Order</h1>
            <p className="text-slate-500 mt-2">Inventaris laden...</p>
          </div>
          <div 
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }} 
            onDragLeave={() => setIsDragging(false)} 
            onDrop={onDrop} 
            onClick={() => document.getElementById('fileInput')?.click()}
            className={cn("relative group cursor-pointer border-2 border-dashed rounded-3xl p-12 transition-all flex flex-col items-center justify-center gap-4", isDragging ? "border-indigo-500 bg-indigo-50" : "border-slate-200 bg-white hover:border-indigo-400")}
          >
            <input id="fileInput" type="file" className="hidden" accept=".xlsx, .xls" onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0])} />
            <Upload className="text-slate-400" size={40} />
            <div className="text-center">
              <p className="text-lg font-semibold text-slate-900">Sleep Excel bestand hierheen</p>
              <p className="text-sm text-slate-500">of klik om te bladeren</p>
            </div>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col lg:flex-row font-sans">
      {/* Linker Sidebar */}
      <aside className="w-full lg:w-72 bg-white border-r border-slate-200 p-6 flex flex-col shrink-0">
        <div className="flex items-center gap-3 text-indigo-600 mb-8 cursor-pointer" onClick={() => setView('order')}>
          <Package size={24} />
          <span className="font-bold text-xl tracking-tight text-slate-900">MedSet</span>
        </div>
        <nav className="space-y-1">
          <button onClick={() => setView('history')} className={cn("w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all mb-4", view === 'history' ? "bg-indigo-50 text-indigo-700 shadow-sm" : "text-slate-600 hover:bg-slate-50")}>
            <History size={18} /> Geschiedenis
          </button>
          <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4 ml-4">Regio's</h2>
          {data.map((region) => (
            <button key={region.name} onClick={() => { setSelectedRegion(region.name); setView('order'); }} className={cn("w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-medium transition-all", view === 'order' && selectedRegion === region.name ? "bg-indigo-50 text-indigo-700 shadow-sm" : "text-slate-600 hover:bg-slate-50")}>
              <div className="flex items-center gap-3"><MapPin size={18} /> {region.name}</div>
              <span className="text-xs bg-slate-100 px-2 py-0.5 rounded-full text-slate-400">{region.sets.length}</span>
            </button>
          ))}
        </nav>
      </aside>

      {/* Middenstuk (Sets / Geschiedenis) */}
      <main className="flex-1 p-6 lg:p-10 overflow-auto bg-slate-50/50">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-3xl font-bold text-slate-900 mb-8">{view === 'history' ? 'Bestelgeschiedenis' : selectedRegion}</h1>
          {view === 'order' ? (
            <div className="grid gap-3">
              {data.find(r => r.name === selectedRegion)?.sets.map((set) => (
                <div key={set.id} onClick={() => toggleSet(set)} className={cn("group cursor-pointer flex items-center justify-between p-4 rounded-2xl border transition-all", selectedSets.some(s => s.setId === set.id) ? "border-indigo-500 bg-white shadow-md ring-2 ring-indigo-500/10" : "bg-white border-slate-200 hover:border-slate-300")}>
                  <div className="flex items-center gap-4">
                    <div className={cn("w-12 h-12 rounded-xl flex items-center justify-center", selectedSets.some(s => s.setId === set.id) ? "bg-indigo-100 text-indigo-600" : "bg-slate-100 text-slate-400")}>
                      <Package size={24} />
                    </div>
                    <div><h3 className="font-bold text-slate-900">{set.name}</h3><p className="text-sm text-slate-500">{set.code}</p></div>
                  </div>
                  {selectedSets.some(s => s.setId === set.id) && <CheckCircle2 className="text-indigo-600" size={20} />}
                </div>
              ))}
            </div>
          ) : (
            <div className="space-y-4">
              {orderHistory.map((order) => (
                <div key={order.id} className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                  <div className="flex justify-between items-start mb-4">
                    <div><p className="text-sm font-bold text-slate-900">{new Date(order.timestamp).toLocaleString('nl-NL')}</p><p className="text-xs text-indigo-600 font-medium">{order.metadata.hospital}</p></div>
                    <span className="px-3 py-1 bg-emerald-100 text-emerald-700 rounded-full text-[10px] font-bold uppercase">Verzonden</span>
                  </div>
                  <div className="text-sm text-slate-600 space-y-1">{order.items.map((it, i) => <p key={i}>• {it.setName} ({it.region})</p>)}</div>
                </div>
              ))}
            </div>
          )}
        </div>
      </main>

      {/* Rechter Sidebar (Winkelmand) */}
      <aside className="w-full lg:w-96 bg-white border-l border-slate-200 p-6 flex flex-col shrink-0">
        <h2 className="font-bold text-lg mb-6 flex items-center gap-2 text-slate-900"><ShoppingCart size={20} className="text-indigo-600" /> Bestelling</h2>
        <div className="flex-1 space-y-2 overflow-y-auto">
          {selectedSets.length === 0 && <p className="text-slate-400 text-sm italic text-center mt-10">Selecteer sets uit de lijst</p>}
          {selectedSets.map(item => (
            <div key={item.setId} className="p-3 bg-slate-50 rounded-xl flex justify-between items-center text-sm border border-slate-100">
              <span className="font-medium truncate mr-2 text-slate-700">{item.setName}</span>
              <button onClick={() => setSelectedSets(prev => prev.filter(s => s.setId !== item.setId))} className="text-slate-400 hover:text-red-500 transition-colors"><Trash2 size={16} /></button>
            </div>
          ))}
        </div>
        <button disabled={selectedSets.length === 0 || orderPlaced} onClick={() => setShowConfirmModal(true)} className="w-full py-4 mt-6 rounded-2xl font-bold text-white bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 shadow-lg shadow-indigo-100 transition-all active:scale-95">Bestelling Afronden</button>
      </aside>

      {/* MODAL VOOR GEGEVENS */}
      <AnimatePresence>
        {showConfirmModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowConfirmModal(false)} />
            <motion.div initial={{ opacity: 0, scale: 0.95, y: 20 }} animate={{ opacity: 1, scale: 1, y: 0 }} exit={{ opacity: 0, scale: 0.95, y: 20 }} className="relative w-full max-w-lg bg-white rounded-3xl p-8 shadow-2xl overflow-y-auto max-h-[90vh]">
              <h3 className="text-2xl font-bold mb-6 text-slate-900">Bestelgegevens</h3>
              <div className="space-y-5">
                <div>
                  <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1 block">Ziekenhuis</label>
                  <select className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500 transition-all text-sm" value={metadata.hospital} onChange={e => setMetadata(prev => ({ ...prev, hospital: e.target.value, surgeon: '' }))}>
                    <option value="">-- Kies ziekenhuis --</option>
                    {Object.keys(ZIEKENHUIS_DATA).map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1 block">Chirurg</label>
                  <select disabled={!metadata.hospital} className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-500 disabled:opacity-50 transition-all text-sm" value={metadata.surgeon} onChange={e => setMetadata(prev => ({ ...prev, surgeon: e.target.value }))}>
                    <option value="">-- Kies chirurg --</option>
                    {metadata.hospital && ZIEKENHUIS_DATA[metadata.hospital].chirurgen.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1 block">OKa-dag</label>
                    <input type="date" className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:ring-2 focus:ring-indigo-500" value={metadata.date} onChange={e => setMetadata(prev => ({ ...prev, date: e.target.value }))} />
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1 block">Agent Naam</label>
                    <input type="text" placeholder="Uw naam" className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:ring-2 focus:ring-indigo-500" value={metadata.agentName} onChange={e => setMetadata(prev => ({ ...prev, agentName: e.target.value }))} />
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

              <div className="flex gap-4 mt-10">
                <button onClick={() => setShowConfirmModal(false)} className="flex-1 py-3 font-bold text-slate-400 hover:text-slate-600">Annuleren</button>
                <button 
                  disabled={!metadata.hospital || !metadata.surgeon || !metadata.date || selectedEmails.length === 0} 
                  onClick={placeOrder} 
                  className="flex-1 py-3 font-bold text-white bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:bg-slate-200 shadow-md transition-all active:scale-95"
                >
                  Verstuur Mail
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
