import React, { useState, useCallback, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  ChevronRight, 
  ShoppingCart, 
  Trash2, 
  CheckCircle2, 
  FileSpreadsheet,
  Package,
  MapPin,
  AlertCircle,
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

export default function App() {
  const [data, setData] = useState<RegionData[]>([]);
  const [selectedRegion, setSelectedRegion] = useState<string | null>(null);
  const [selectedSets, setSelectedSets] = useState<OrderItem[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [orderPlaced, setOrderPlaced] = useState(false);
  const [showConfirmModal, setShowConfirmModal] = useState(false);
  const [view, setView] = useState<'order' | 'history'>('order');
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

  // 1. FUNCTIE OM DATA TE VERWERKEN
  const processExcelData = useCallback((workbook: XLSX.WorkBook) => {
    const regions: RegionData[] = workbook.SheetNames.map((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<any>(ws);
      
      const sets: MedicalSet[] = rows.map((row, index) => {
        const keys = Object.keys(row);
        const name = row[keys[0]] || `Set ${index + 1}`;
        const code = row[keys[1]] ? String(row[keys[1]]) : '';
        const description = row[keys[2]] ? String(row[keys[2]]) : '';
        
        return {
          id: `${sheetName}-${index}`,
          name: String(name),
          code: code,
          description: description,
          region: sheetName
        };
      });

      return {
        name: sheetName,
        sets
      };
    });

    setData(regions);
    if (regions.length > 0) {
      setSelectedRegion(regions[0].name);
    }
  }, []);

  // 2. AUTOMATISCH LADEN VAN DATA.XLSX UIT DE PUBLIC MAP
  useEffect(() => {
    const loadDefaultFile = async () => {
      try {
        const response = await fetch('/data.xlsx');
        if (response.ok) {
          const arrayBuffer = await response.arrayBuffer();
          const wb = XLSX.read(arrayBuffer, { type: 'buffer' });
          processExcelData(wb);
        } else {
          console.error("data.xlsx niet gevonden in public map");
        }
      } catch (error) {
        console.error("Fout bij laden van standaardbestand:", error);
      }
    };
    loadDefaultFile();
  }, [processExcelData]);

  // 3. HANDMATIGE UPLOAD LOGICA
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
      if (exists) {
        return prev.filter(s => s.setId !== set.id);
      }
      return [...prev, { setId: set.id, setName: set.name, setCode: set.code, region: set.region }];
    });
  };

  const removeOrderItem = (setId: string) => {
    setSelectedSets(prev => prev.filter(s => s.setId !== setId));
  };

  const currentSets = useMemo(() => {
    return data.find(r => r.name === selectedRegion)?.sets || [];
  }, [data, selectedRegion]);

  const placeOrder = () => {
    setShowConfirmModal(false);
    setOrderPlaced(true);
    
    const newOrder: OrderHistoryItem = {
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      items: selectedSets,
      metadata: { ...metadata }
    };

    const updatedHistory = [newOrder, ...orderHistory];
    setOrderHistory(updatedHistory);
    localStorage.setItem('medset_history', JSON.stringify(updatedHistory));

    const subject = `Bestelling Medische Sets - ${metadata.hospital} - ${metadata.date}`;
    const body = `
Beste,

Hierbij een nieuwe bestelling voor medische sets.

DETAILS INGREEP
------------------------------------------------
Ziekenhuis: ${metadata.hospital}
Datum: ${metadata.date}
Chirurg: ${metadata.surgeon}
Agent: ${metadata.agentName}

BESTELDE SETS
------------------------------------------------
${selectedSets.map(s => `- ${s.setName} ${s.setCode ? `[Code: ${s.setCode}]` : ''} (${s.region})`).join('\n')}

OPMERKINGEN
------------------------------------------------
${metadata.remarks || 'Geen opmerkingen'}

Met vriendelijke groet,
${metadata.agentName}
    `.trim();

    const mailtoLink = `mailto:${metadata.infoEmails}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    window.location.href = mailtoLink;

    setTimeout(() => {
      setOrderPlaced(false);
      setSelectedSets([]);
      setMetadata({
        hospital: '',
        date: '',
        surgeon: '',
        agentName: '',
        remarks: '',
        infoEmails: ''
      });
      setView('history');
    }, 2000);
  };

  // 4. WEERGAVE ALS ER NOG GEEN DATA IS (LAADSCHERM)
  if (data.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6 bg-slate-50">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-xl w-full"
        >
          <div className="text-center mb-8">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-indigo-600 text-white mb-4 shadow-lg shadow-indigo-200">
              <Package className="animate-bounce" size={32} />
            </div>
            <h1 className="text-3xl font-bold text-slate-900 tracking-tight">MedSet Loaner Order</h1>
            <p className="text-slate-500 mt-2">Inventaris laden... Geduld aub.</p>
          </div>

          <div
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
            className={cn(
              "relative group cursor-pointer border-2 border-dashed rounded-3xl p-12 transition-all duration-300 flex flex-col items-center justify-center gap-4",
              isDragging 
                ? "border-indigo-500 bg-indigo-50/50 scale-[1.02]" 
                : "border-slate-200 bg-white hover:border-indigo-400 hover:bg-slate-50/50"
            )}
            onClick={() => document.getElementById('fileInput')?.click()}
          >
            <input
              id="fileInput"
              type="file"
              className="hidden"
              accept=".xlsx, .xls"
              onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0])}
            />
            <div className="w-16 h-16 rounded-full bg-slate-100 flex items-center justify-center text-slate-400 group-hover:text-indigo-500 group-hover:bg-indigo-100 transition-colors">
              <Upload size={28} />
            </div>
            <div className="text-center">
              <p className="text-lg font-semibold text-slate-900">Sleep Excel bestand hierheen</p>
              <p className="text-sm text-slate-500 mt-1">of klik om handmatig te bladeren</p>
            </div>
          </div>
        </motion.div>
      </div>
    );
  }

  // 5. HOOFD APPLICATIE
  return (
    <div className="min-h-screen bg-slate-50 flex flex-col lg:flex-row">
      <aside className="w-full lg:w-72 bg-white border-r border-slate-200 flex flex-col shrink-0">
        <div className="p-6 border-bottom border-slate-100">
          <div className="flex items-center gap-3 text-indigo-600 mb-6 cursor-pointer" onClick={() => setView('order')}>
            <Package size={24} />
            <span className="font-bold text-xl tracking-tight text-slate-900">MedSet</span>
          </div>

          <div className="mb-6">
            <button
              onClick={() => setView('history')}
              className={cn(
                "w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all mb-2",
                view === 'history'
                  ? "bg-indigo-50 text-indigo-700 shadow-sm"
                  : "text-slate-600 hover:bg-slate-50 hover:text-slate-900"
              )}
            >
              <History size={18} className={view === 'history' ? "text-indigo-500" : "text-slate-400"} />
              Geschiedenis
            </button>
          </div>

          <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Regio's</h2>
          <nav className="space-y-1">
            {data.map((region) => (
              <button
                key={region.name}
                onClick={() => {
                  setSelectedRegion(region.name);
                  setView('order');
                }}
                className={cn(
                  "w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-medium transition-all",
                  view === 'order' && selectedRegion === region.name
                    ? "bg-indigo-50 text-indigo-700 shadow-sm"
                    : "text-slate-600 hover:bg-slate-50 hover:text-slate-900"
                )}
              >
                <div className="flex items-center gap-3">
                  <MapPin size={18} className={selectedRegion === region.name ? "text-indigo-500" : "text-slate-400"} />
                  {region.name}
                </div>
                <span className="text-xs bg-white px-2 py-0.5 rounded-full border border-slate-200 text-slate-400">
                  {region.sets.length}
                </span>
              </button>
            ))}
          </nav>
        </div>
        
        <div className="mt-auto p-6 border-t border-slate-100">
          <button 
            onClick={() => setData([])}
            className="w-full flex items-center justify-center gap-2 px-4 py-2 text-sm font-medium text-slate-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
          >
            <X size={16} />
            Ander bestand uploaden
          </button>
        </div>
      </aside>

      <main className="flex-1 p-6 lg:p-10 overflow-auto bg-slate-50/50">
        <div className="max-w-4xl mx-auto">
          {view === 'history' ? (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6"
            >
              <header className="mb-8">
                <div className="flex items-center gap-2 text-sm text-slate-400 mb-2">
                  <span>Overzicht</span>
                  <ChevronRight size={14} />
                  <span className="text-indigo-600 font-medium">Geschiedenis</span>
                </div>
                <h1 className="text-3xl font-bold text-slate-900">Bestelgeschiedenis</h1>
              </header>

              {orderHistory.length === 0 ? (
                <div className="text-center py-12 bg-white rounded-3xl border border-slate-200">
                  <div className="w-16 h-16 rounded-full bg-slate-50 flex items-center justify-center mx-auto mb-4 text-slate-400">
                    <History size={32} />
                  </div>
                  <h3 className="text-lg font-semibold text-slate-900">Nog geen bestellingen</h3>
                  <p className="text-slate-500 mt-1">Uw geplaatste bestellingen verschijnen hier.</p>
                  <button 
                    onClick={() => setView('order')}
                    className="mt-6 px-6 py-2 bg-indigo-600 text-white rounded-xl font-medium hover:bg-indigo-700 transition-colors"
                  >
                    Nieuwe bestelling
                  </button>
                </div>
              ) : (
                <div className="space-y-4">
                  {orderHistory.map((order) => (
                    <div key={order.id} className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition-shadow">
                      <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/50">
                        <div className="flex items-center gap-4">
                          <div className="w-10 h-10 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                            <Clock size={20} />
                          </div>
                          <div>
                            <p className="text-sm font-bold text-slate-900">
                              {new Date(order.timestamp).toLocaleDateString('nl-NL', { 
                                weekday: 'long', 
                                year: 'numeric', 
                                month: 'long', 
                                day: 'numeric',
                                hour: '2-digit',
                                minute: '2-digit'
                              })}
                            </p>
                            <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                              <span className="font-medium text-indigo-600">{order.items.length} sets</span>
                              <span>•</span>
                              <span>{order.metadata.hospital || 'Onbekend ziekenhuis'}</span>
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold uppercase tracking-wider">
                            Besteld
                          </span>
                        </div>
                      </div>
                      
                      <div className="p-6">
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 mb-6 text-sm">
                          <div>
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Details</p>
                            <div className="space-y-1 text-slate-600">
                              <p><span className="text-slate-400 w-24 inline-block">Chirurg:</span> {order.metadata.surgeon || '-'}</p>
                              <p><span className="text-slate-400 w-24 inline-block">Datum Ingreep:</span> {order.metadata.date || '-'}</p>
                              <p><span className="text-slate-400 w-24 inline-block">Agent:</span> {order.metadata.agentName || '-'}</p>
                            </div>
                          </div>
                          {order.metadata.remarks && (
                            <div>
                              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Opmerkingen</p>
                              <p className="text-slate-600 italic bg-slate-50 p-3 rounded-lg border border-slate-100">
                                "{order.metadata.remarks}"
                              </p>
                            </div>
                          )}
                        </div>

                        <div>
                          <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">Bestelde Sets</p>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            {order.items.map((item, idx) => (
                              <div key={`${order.id}-${idx}`} className="flex items-center gap-2 p-2 rounded-lg bg-slate-50 border border-slate-100 text-sm">
                                <Package size={14} className="text-indigo-400" />
                                <span className="font-medium text-slate-700 truncate">
                                  {item.setName}
                                  {item.setCode && <span className="ml-1 text-[10px] text-slate-400">({item.setCode})</span>}
                                </span>
                                <span className="text-xs text-slate-400 ml-auto bg-white px-1.5 py-0.5 rounded border border-slate-200">
                                  {item.region}
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </motion.div>
          ) : (
            <>
              <header className="mb-8">
                <div className="flex items-center gap-2 text-sm text-slate-400 mb-2">
                  <span>Inventaris</span>
                  <ChevronRight size={14} />
                  <span className="text-indigo-600 font-medium">{selectedRegion}</span>
                </div>
                <h1 className="text-3xl font-bold text-slate-900">Beschikbare Sets</h1>
              </header>

              <div className="grid gap-3">
                {currentSets.map((set) => {
                  const isSelected = selectedSets.some(s => s.setId === set.id);
                  return (
                    <motion.div
                      layout
                      key={set.id}
                      onClick={() => toggleSet(set)}
                      className={cn(
                        "group cursor-pointer flex items-center justify-between p-4 rounded-2xl border transition-all duration-200",
                        isSelected 
                          ? "bg-white border-indigo-500 shadow-md ring-1 ring-indigo-500" 
                          : "bg-white border-slate-200 hover:border-slate-300 hover:shadow-sm"
                      )}
                    >
                      <div className="flex items-center gap-4">
                        <div className={cn(
                          "w-12 h-12 rounded-xl flex items-center justify-center transition-colors",
                          isSelected ? "bg-indigo-100 text-indigo-600" : "bg-slate-100 text-slate-400 group-hover:bg-slate-200"
                        )}>
                          <Package size={24} />
                        </div>
                        <div>
                          <div className="flex items-center gap-2">
                            <h3 className="font-semibold text-slate-900">{set.name}</h3>
                            {set.code && (
                              <span className="text-[10px] font-bold bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded border border-slate-200">
                                {set.code}
                              </span>
                            )}
                          </div>
                          {set.description && (
                            <p className="text-sm text-slate-500">{set.description}</p>
                          )}
                        </div>
                      </div>
                      <div className={cn(
                        "w-6 h-6 rounded-full border-2 flex items-center justify-center transition-all",
                        isSelected 
                          ? "bg-indigo-600 border-indigo-600 text-white scale-110" 
                          : "border-slate-200 group-hover:border-slate-300"
                      )}>
                        {isSelected && <CheckCircle2 size={14} strokeWidth={3} />}
                      </div>
                    </motion.div>
                  );
                })}
              </div>
            </>
          )}
        </div>
      </main>

      <aside className="w-full lg:w-96 bg-white border-l border-slate-200 flex flex-col shrink-0">
        <div className="p-6 flex items-center justify-between border-b border-slate-100">
          <h2 className="font-bold text-lg text-slate-900 flex items-center gap-2">
            <ShoppingCart size={20} className="text-indigo-600" />
            Bestelling
          </h2>
          <span className="bg-indigo-100 text-indigo-700 px-2.5 py-0.5 rounded-full text-xs font-bold">
            {selectedSets.length} items
          </span>
        </div>

        <div className="flex-1 overflow-auto p-6">
          <AnimatePresence mode="popLayout">
            {selectedSets.length === 0 ? (
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="h-full flex flex-col items-center justify-center text-center text-slate-400"
              >
                <div className="w-16 h-16 rounded-full bg-slate-50 flex items-center justify-center mb-4">
                  <ShoppingCart size={24} />
                </div>
                <p className="text-sm">Geen sets geselecteerd</p>
                <p className="text-xs mt-1">Vink sets aan in de lijst om ze toe te voegen</p>
              </motion.div>
            ) : (
              <div className="space-y-3">
                {selectedSets.map((item) => (
                  <motion.div
                    key={item.setId}
                    layout
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: -20 }}
                    className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100 group"
                  >
                    <div className="min-w-0">
                      <p className="text-sm font-semibold text-slate-900 truncate">
                        {item.setName}
                        {item.setCode && <span className="ml-2 text-[10px] text-slate-400">({item.setCode})</span>}
                      </p>
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{item.region}</p>
                    </div>
                    <button 
                      onClick={() => removeOrderItem(item.setId)}
                      className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                    >
                      <Trash2 size={16} />
                    </button>
                  </motion.div>
                ))}
              </div>
            )}
          </AnimatePresence>
        </div>

        <div className="p-6 border-t border-slate-100 bg-slate-50/50">
          <button
            disabled={selectedSets.length === 0 || orderPlaced}
            onClick={() => setShowConfirmModal(true)}
            className={cn(
              "w-full py-4 rounded-2xl font-bold text-white shadow-lg transition-all flex items-center justify-center gap-2",
              selectedSets.length === 0 
                ? "bg-slate-300 cursor-not-allowed shadow-none" 
                : orderPlaced 
                  ? "bg-emerald-500 shadow-emerald-100" 
                  : "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-100 hover:scale-[1.02] active:scale-[0.98]"
            )}
          >
            {orderPlaced ? (
              <>
                <CheckCircle2 size={20} />
                Bestelling Geplaatst!
              </>
            ) : (
              <>
                Bestellen
                <ChevronRight size={20} />
              </>
            )}
          </button>
          <p className="text-[10px] text-center text-slate-400 mt-4 uppercase tracking-widest font-bold">
            MedSet Loaner Management System
          </p>
        </div>
      </aside>

      <AnimatePresence>
        {showConfirmModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowConfirmModal(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-lg bg-white rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                <h3 className="text-xl font-bold text-slate-900">Bevestig Bestelling</h3>
                <button 
                  onClick={() => setShowConfirmModal(false)}
                  className="p-2 hover:bg-slate-100 rounded-full text-slate-400 transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
              
              <div className="p-6 max-h-[70vh] overflow-auto space-y-6">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <Hospital size={12} /> Ziekenhuis
                    </label>
                    <input
                      type="text"
                      value={metadata.hospital}
                      onChange={(e) => setMetadata(prev => ({ ...prev, hospital: e.target.value }))}
                      placeholder="Naam ziekenhuis"
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <Calendar size={12} /> Datum Ingreep
                    </label>
                    <input
                      type="date"
                      value={metadata.date}
                      onChange={(e) => setMetadata(prev => ({ ...prev, date: e.target.value }))}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <User size={12} /> Chirurg
                    </label>
                    <input
                      type="text"
                      value={metadata.surgeon}
                      onChange={(e) => setMetadata(prev => ({ ...prev, surgeon: e.target.value }))}
                      placeholder="Naam chirurg"
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <UserCheck size={12} /> Agent Naam
                    </label>
                    <input
                      type="text"
                      value={metadata.agentName}
                      onChange={(e) => setMetadata(prev => ({ ...prev, agentName: e.target.value }))}
                      placeholder="Uw naam"
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <Mail size={12} /> Info aan:
                    </label>
                    <input
                      type="text"
                      value={metadata.infoEmails}
                      onChange={(e) => setMetadata(prev => ({ ...prev, infoEmails: e.target.value }))}
                      placeholder="E-mailadressen (gescheiden door komma's)"
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    />
                  </div>
                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                      <MessageSquare size={12} /> Opmerkingen
                    </label>
                    <textarea
                      value={metadata.remarks}
                      onChange={(e) => setMetadata(prev => ({ ...prev, remarks: e.target.value }))}
                      placeholder="Eventuele extra informatie..."
                      rows={2}
                      className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all resize-none"
                    />
                  </div>
                </div>

                <div className="pt-4 border-t border-slate-100">
                  <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">Geselecteerde Sets</p>
                  <div className="space-y-2">
                    {selectedSets.map((item) => (
                      <div key={item.setId} className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl border border-slate-100">
                        <div className="w-8 h-8 rounded-lg bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                          <Package size={16} />
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-semibold text-slate-900 truncate">
                            {item.setName}
                            {item.setCode && <span className="ml-2 text-[10px] text-slate-400">({item.setCode})</span>}
                          </p>
                          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{item.region}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="p-6 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row gap-3">
                <button
                  onClick={() => setShowConfirmModal(false)}
                  className="flex-1 px-6 py-3 rounded-xl font-bold text-slate-600 hover:bg-slate-200 transition-colors"
                >
                  Annuleren
                </button>
                <button
                  onClick={placeOrder}
                  className="flex-1 px-6 py-3 rounded-xl font-bold text-white bg-indigo-600 hover:bg-indigo-700 shadow-lg shadow-indigo-100 transition-all"
                >
                  Bevestig Bestelling
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
