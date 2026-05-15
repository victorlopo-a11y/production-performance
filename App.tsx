import React, { useState, useMemo, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell,
  PieChart, Pie
} from 'recharts';
import { 
  Upload, CheckCircle2, AlertCircle, Filter, 
  BarChart3, List, Calendar, User, ArrowUpRight, ArrowDownRight, FileType, Search, Trash2, X, ChevronDown
} from 'lucide-react';
import { ProductionPlanRow, ActualProductionRecord, FailureReportRecord, ComparisonResult } from './types';
import { parseProductionPlanExcel, parseActualProductionExcel, parseFailureReportExcel, fileToBase64, calculateShiftAndProductionDay } from './services/excelParser';
import { GoogleGenAI } from "@google/genai";
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';

const App: React.FC = () => {
  const apiKey =
    (import.meta as any)?.env?.VITE_API_KEY ||
    (import.meta as any)?.env?.API_KEY ||
    (globalThis as any)?.process?.env?.API_KEY ||
    '';
  const [planData, setPlanData] = useState<ProductionPlanRow[]>([]);
  const [actualData, setActualData] = useState<ActualProductionRecord[]>([]);
  const [failureData, setFailureData] = useState<FailureReportRecord[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState('Processando Arquivos...');
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard'>('upload');
  
  const [selectedShift, setSelectedShift] = useState<Array<number | 'all'>>(['all']);
  const [selectedDate, setSelectedDate] = useState<Array<string | 'all'>>(['all']);
  const [selectedLine, setSelectedLine] = useState<Array<string | 'all'>>(['all']);
  const [searchTerm, setSearchTerm] = useState('');
  const tableRef = useRef<HTMLDivElement | null>(null);
  const [isExportingPng, setIsExportingPng] = useState(false);

  const readMultiSelectValues = (e: React.ChangeEvent<HTMLSelectElement>) =>
    Array.from(e.target.selectedOptions).map(option => option.value);

  const normalizeProductCode = (code: string): string => {
    if (!code) return '';
    const cleaned = code
      .toLowerCase()
      .trim()
      .replace(/\s+/g, '')
      .replace(/[^a-z0-9]/g, '');
    const stripped = cleaned.replace(/(rw|aj|a)$/i, '');
    const unified = stripped.replace(/mme$/i, 'mm');
    const match = unified.match(/^([a-z]+)0*([0-9]+[a-z]*)$/i);
    if (match) return `${match[1]}${match[2]}`;
    return unified;
  };

  
  const normalizeLineCode = (line: string): string => {
    if (!line) return '';
    const raw = line.toLowerCase().trim();
    const match = raw.match(/([a-z]+)\s*0*([0-9]+[a-z]?)/i);
    if (match) return `${match[1]}${match[2]}`;
    const cleaned = raw.replace(/\s+/g, '');
    const fallback = cleaned.match(/^([a-z]+)0*([0-9]+[a-z]*)$/i);
    if (fallback) return `${fallback[1]}${fallback[2]}`;
    return cleaned;
  };

  const extractPlanFromPDF = async (file: File): Promise<ProductionPlanRow[]> => {
    try {
      const ai = new GoogleGenAI({ apiKey });
      const base64 = await fileToBase64(file);
      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: {
          parts: [
            { inlineData: { data: base64, mimeType: 'application/pdf' } },
            { text: "Extraia o plano de produção. LINHA, TURNO, CÓDIGO, PRODUTO e metas (PROG). Retorne JSON: Array<{ line: string, shift: number, code: string, product: string, date: string (DD/MM/YYYY), meta: number }>." }
          ]
        },
        config: { responseMimeType: "application/json" }
      });
      return JSON.parse(response.text || '[]');
    } catch (e) {
      console.error(e);
      return [];
    }
  };

  const extractActualFromPDF = async (file: File): Promise<ActualProductionRecord[]> => {
    try {
      const ai = new GoogleGenAI({ apiKey });
      const base64 = await fileToBase64(file);
      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: {
          parts: [
            { inlineData: { data: base64, mimeType: 'application/pdf' } },
            { text: "Extraia a produção real. Linha, Material, Quantidade, Data e Hora. Retorne JSON: Array<{ line: string, material: string, quantity: number, date: string (DD/MM/YYYY), time: string (HH:MM:SS) }>." }
          ]
        },
        config: { responseMimeType: "application/json" }
      });
      const raw = JSON.parse(response.text || '[]');
      return raw.map((item: any) => {
        const shiftData = calculateShiftAndProductionDay(item.date, item.time);
        return { ...item, shift: shiftData.shift, productionDay: shiftData.productionDay };
      });
    } catch (e) {
      console.error(e);
      return [];
    }
  };

  const handlePlanUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    setIsLoading(true);
    setLoadingMessage('Importando Planos de Produção...');
    try {
      const files = Array.from(e.target.files);
      let all = [];
      for (const f of files) {
        const results = f.name.endsWith('.pdf') ? await extractPlanFromPDF(f) : await parseProductionPlanExcel(f);
        all = [...all, ...results];
      }
      if (all.length === 0) {
        alert('NÃ£o encontrei metas (PROG) nesse arquivo. Confirme se a aba \"Plano\" existe e se o arquivo Ã© do modelo correto.');
      }
      setPlanData(prev => [...prev, ...all]);
    } catch (err) { alert('Erro no processamento.'); }
    finally { setIsLoading(false); e.target.value = ''; }
  };

  const handlePadokaPlanUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    setIsLoading(true);
    setLoadingMessage('Importando Plano PADOKA...');
    try {
      const files = Array.from(e.target.files);
      let all = [];
      for (const f of files) {
        const results = f.name.endsWith('.pdf') ? await extractPlanFromPDF(f) : await parseProductionPlanExcel(f);
        all = [...all, ...results];
      }
      setPlanData(prev => [...prev, ...all]);
    } catch (err) { alert('Erro no processamento.'); }
    finally { setIsLoading(false); e.target.value = ''; }
  };

  const handleFailureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    setIsLoading(true);
    setLoadingMessage('Importando Relatorio AT - AR...');
    try {
      const files = Array.from(e.target.files);
      let all: FailureReportRecord[] = [];
      for (const f of files) {
        if (f.name.endsWith('.pdf')) {
          const results = await extractActualFromPDF(f);
          all = [...all, ...results.map(r => ({ ...r, origin: '' }))];
        } else {
          const results = await parseFailureReportExcel(f);
          all = [...all, ...results];
        }
      }
      setFailureData(prev => [...prev, ...all]);
    } catch (err) { alert('Erro no processamento.'); }
    finally { setIsLoading(false); e.target.value = ''; }
  };

  const handleActualUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    setIsLoading(true);
    setLoadingMessage('Sincronizando Produção Real...');
    try {
      const files = Array.from(e.target.files);
      let all = [];
      for (const f of files) {
        const results = f.name.endsWith('.pdf') ? await extractActualFromPDF(f) : await parseActualProductionExcel(f);
        all = [...all, ...results];
      }
      setActualData(prev => [...prev, ...all]);
    } catch (err) { alert('Erro no processamento.'); }
    finally { setIsLoading(false); e.target.value = ''; }
  };

  const comparisonResults = useMemo(() => {
    const results: ComparisonResult[] = [];
    
    // 1. Agrega Produção Real: Dia | Turno | Linha | Material (Normalizado)
    const actualAggregated: Record<string, { produced: number; originalLine: string; originalMaterial: string }> = {};
    const producedByLine: Record<string, number> = {};
    actualData.forEach(item => {
      const normCode = normalizeProductCode(item.material);
      const lineKey = normalizeLineCode(item.line);
      const key = `${item.productionDay}|${item.shift}|${lineKey}|${normCode}`;
      const lineOnlyKey = `${item.productionDay}|${item.shift}|${lineKey}`;
      if (!actualAggregated[key]) {
        actualAggregated[key] = { produced: 0, originalLine: item.line, originalMaterial: item.material };
      }
      actualAggregated[key].produced += item.quantity;
      producedByLine[lineOnlyKey] = (producedByLine[lineOnlyKey] || 0) + item.quantity;
    });

    // 2. Agrega Plano de Metas: Dia | Turno | Linha | Material (Normalizado)
    const planAggregated: Record<string, { meta: number; originalCode: string; originalProduct: string; originalLine: string }> = {};
    planData.forEach(p => {
      const normCode = normalizeProductCode(p.code);
      const lineKey = normalizeLineCode(p.line);
      const key = `${p.date}|${p.shift}|${lineKey}|${normCode}`;
      
      if (!planAggregated[key]) {
        planAggregated[key] = { meta: 0, originalCode: p.code, originalProduct: p.product, originalLine: p.line };
      }
      planAggregated[key].meta += p.meta;
    });

    // 3. Agrega Falhas: Dia | Turno | Linha (Normalizado)
    const failuresByLine: Record<string, { total: number; origins: Record<string, number> }> = {};
    failureData.forEach(item => {
      const lineKey = normalizeLineCode(item.line);
      const key = `${item.productionDay}|${item.shift}|${lineKey}`;
      if (!failuresByLine[key]) {
        failuresByLine[key] = { total: 0, origins: {} };
      }
      failuresByLine[key].total += item.quantity;
      const originKey = item.origin ? item.origin : 'Sem origem';
      failuresByLine[key].origins[originKey] = (failuresByLine[key].origins[originKey] || 0) + item.quantity;
    });

    // 4. Cruzamento de Dados
    Object.keys(planAggregated).forEach(key => {
      const [date, shiftStr, lineKey, normCode] = key.split('|');
      const shift = parseInt(shiftStr);
      const p = planAggregated[key];
      
      // Busca produção REAL exata para este turno e linha
      const produced = actualAggregated[key]?.produced || 0;
      const lineOnlyKey = `${date}|${shift}|${lineKey}`;
      const failures = failuresByLine[lineOnlyKey]?.total || 0;
      const failuresByOrigin = failuresByLine[lineOnlyKey]?.origins || {};
      const producedLine = producedByLine[lineOnlyKey] || 0;
      const yieldValue = producedLine > 0 ? Math.max(0, (1 - failures / producedLine) * 100) : 0;
      const difference = produced - p.meta;
      const efficiency = p.meta > 0 ? (produced / p.meta) * 100 : 0;

      results.push({
        date,
        shift,
        line: p.originalLine,
        material: p.originalCode,
        product: p.originalProduct,
        meta: p.meta,
        produced,
        failures,
        failuresByOrigin,
        yield: yieldValue,
        difference,
        efficiency
      });
    });

    Object.keys(actualAggregated).forEach(key => {
      if (planAggregated[key]) return;
      const [date, shiftStr, lineKey] = key.split('|');
      const shift = parseInt(shiftStr);
      const a = actualAggregated[key];
      const lineOnlyKey = `${date}|${shift}|${lineKey}`;
      const failures = failuresByLine[lineOnlyKey]?.total || 0;
      const failuresByOrigin = failuresByLine[lineOnlyKey]?.origins || {};
      const producedLine = producedByLine[lineOnlyKey] || 0;
      const yieldValue = producedLine > 0 ? Math.max(0, (1 - failures / producedLine) * 100) : 0;
      results.push({
        date,
        shift,
        line: a.originalLine,
        material: a.originalMaterial,
        product: 'EXPORT',
        meta: 0,
        produced: a.produced,
        failures,
        failuresByOrigin,
        yield: yieldValue,
        difference: a.produced,
        efficiency: 0
      });
    });

    return results;
  }, [planData, actualData, failureData]);

  const uniqueDates = useMemo(() => [...new Set(comparisonResults.map(r => r.date))].sort((a,b) => {
    const [da,ma,ya] = a.split('/').map(Number);
    const [db,mb,yb] = b.split('/').map(Number);
    return new Date(ya,ma,da).getTime() - new Date(yb,mb,db).getTime();
  }), [comparisonResults]);

  const uniqueLines = useMemo(() => [...new Set(comparisonResults.map(r => r.line))].sort(), [comparisonResults]);

  const filtered = useMemo(() => comparisonResults.filter(r => {
    const mShift = selectedShift.includes('all') || selectedShift.includes(r.shift);
    const mDate = selectedDate.includes('all') || selectedDate.includes(r.date);
    const mLine = selectedLine.includes('all') || selectedLine.includes(r.line);
    const mSearch = searchTerm === '' || r.material.toLowerCase().includes(searchTerm.toLowerCase()) || r.product.toLowerCase().includes(searchTerm.toLowerCase());
    return mShift && mDate && mLine && mSearch;
  }), [comparisonResults, selectedShift, selectedDate, selectedLine, searchTerm]);

  const stats = useMemo(() => {
    const tm = filtered.reduce((a,r) => a + r.meta, 0);
    const tp = filtered.reduce((a,r) => a + r.produced, 0);
    return { tm, tp, eff: tm > 0 ? (tp/tm)*100 : 0, diff: tp - tm };
  }, [filtered]);

  const handleExportExcel = () => {
    const originKeys = Array.from(new Set(
      filtered.flatMap(r => Object.keys(r.failuresByOrigin))
    )).sort();
    const rows = filtered.map(r => {
      const base = {
        Data: r.date,
        Linha: r.line,
        Turno: r.shift,
        Material: r.material,
        Produto: r.product,
        Meta: r.meta,
        ProducaoReal: r.produced,
        Falhas: r.failures,
        Yield: Number(r.yield.toFixed(1)),
        Diferenca: r.difference,
        Eficiencia: Number(r.efficiency.toFixed(1))
      } as Record<string, string | number>;
      originKeys.forEach(k => {
        base[`Falhas - ${k}`] = r.failuresByOrigin[k] || 0;
      });
      return base;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Relatorio');
    const name = `production_performance_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, name);
  };

  const handleExportCsv = () => {
    const originKeys = Array.from(new Set(
      filtered.flatMap(r => Object.keys(r.failuresByOrigin))
    )).sort();
    const headers = ['Data', 'Linha', 'Turno', 'Material', 'Produto', 'Meta', 'ProducaoReal', 'Falhas', 'Yield', 'Diferenca', 'Eficiencia', ...originKeys.map(k => `Falhas - ${k}`)];
    const rows = filtered.map(r => ([
      r.date,
      r.line,
      String(r.shift),
      r.material,
      r.product,
      String(r.meta),
      String(r.produced),
      String(r.failures),
      r.yield.toFixed(1),
      String(r.difference),
      r.efficiency.toFixed(1),
      ...originKeys.map(k => String(r.failuresByOrigin[k] || 0))
    ]));
    const escapeCsv = (val: string) => `"${val.replace(/"/g, '""')}"`;
    const csv = [headers, ...rows]
      .map(row => row.map(escapeCsv).join(','))
      .join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `production_performance_${new Date().toISOString().slice(0,10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const escapeHtml = (val: string) =>
    val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const handleExportPdf = async () => {
    if (!tableRef.current) return;
    setIsExportingPng(true);
    await new Promise(resolve => setTimeout(resolve, 0));
    const canvas = await html2canvas(tableRef.current, { backgroundColor: '#ffffff', scale: 3 });
    setIsExportingPng(false);
    const dataUrl = canvas.toDataURL('image/png');
    const win = window.open('', '_blank', 'width=1200,height=800');
    if (!win) return;
    win.document.write(`
      <html>
        <head>
          <title>Production Performance</title>
          <style>
            @page { margin: 0; }
            html, body { margin: 0; padding: 0; background: #fff; }
            img { width: 100%; height: auto; display: block; }
          </style>
        </head>
        <body>
          <img id="report-img" />
        </body>
      </html>
    `);
    win.document.close();
    const img = win.document.getElementById('report-img') as HTMLImageElement | null;
    if (!img) return;
    img.onload = () => {
      win.focus();
      win.print();
    };
    img.src = dataUrl;
  };

  const handleExportPng = async () => {
    if (!tableRef.current) return;
    setIsExportingPng(true);
    await new Promise(resolve => setTimeout(resolve, 0));
    const canvas = await html2canvas(tableRef.current, { backgroundColor: '#ffffff', scale: 3 });
    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/png');
    link.download = `production_performance_${new Date().toISOString().slice(0,10)}.png`;
    link.click();
    setIsExportingPng(false);
  };

  return (
    <div className="min-h-screen flex flex-col">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm h-16 flex items-center justify-between px-6">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-1.5 rounded-lg text-white"><BarChart3 size={20}/></div>
          <h1 className="font-black text-slate-800 tracking-tight">PRODUCTION PERFORMANCE</h1>
        </div>
        <div className="flex items-center gap-2 bg-slate-100 p-1 rounded-xl">
          <button onClick={() => setActiveTab('upload')} className={`px-4 py-1.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'upload' ? 'bg-white shadow-sm text-blue-600' : 'text-slate-500'}`}>UPLOAD</button>
          <button onClick={() => setActiveTab('dashboard')} disabled={!isReady} className={`px-4 py-1.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'dashboard' ? 'bg-white shadow-sm text-blue-600' : 'text-slate-400'}`}>PAINEL</button>
        </div>
      </header>

      <main className="flex-grow p-6 max-w-7xl mx-auto w-full">
        {activeTab === 'upload' ? (
          <div className="max-w-4xl mx-auto space-y-8 py-10">
            <div className="text-center space-y-2">
              <h2 className="text-4xl font-black text-slate-900 tracking-tighter">CONFIGURAÇÃO DE DADOS</h2>
              <p className="text-slate-400 font-medium">Sincronize metas e produção para análise instantânea por turno.</p>
            </div>
            <div className="grid md:grid-cols-3 gap-6">
              <div className="bg-white p-10 rounded-3xl border border-slate-100 shadow-xl text-center space-y-6">
                <Calendar className="mx-auto text-blue-500" size={48}/>
                <h3 className="text-xl font-black">METAS PROGRAMADAS</h3>
                <label className="block bg-blue-600 text-white py-4 rounded-2xl font-black cursor-pointer hover:bg-blue-700 transition-all shadow-lg shadow-blue-100">
                  <input type="file" multiple onChange={handlePlanUpload} className="hidden" />
                  CARREGAR PLANOS
                </label>
                <label className="block bg-emerald-600 text-white py-3 rounded-2xl font-black cursor-pointer hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100">
                  <input type="file" multiple onChange={handlePadokaPlanUpload} className="hidden" />
                  CARREGAR PLANO PADOKA
                </label>
                {planData.length > 0 && <span className="inline-block bg-green-50 text-green-600 px-4 py-1 rounded-full text-xs font-black uppercase tracking-wider">{planData.length} registros ativos</span>}
              </div>
              <div className="bg-white p-10 rounded-3xl border border-slate-100 shadow-xl text-center space-y-6">
                <List className="mx-auto text-indigo-500" size={48}/>
                <h3 className="text-xl font-black">PRODUÇÃO REALIZADA</h3>
                <label className="block bg-slate-900 text-white py-4 rounded-2xl font-black cursor-pointer hover:bg-slate-800 transition-all shadow-lg shadow-slate-200">
                  <input type="file" multiple onChange={handleActualUpload} className="hidden" />
                  CARREGAR PRODUÇÃO
                </label>
                {actualData.length > 0 && <span className="inline-block bg-green-50 text-green-600 px-4 py-1 rounded-full text-xs font-black uppercase tracking-wider">{actualData.length} entradas reais</span>}
              </div>
              <div className="bg-white p-10 rounded-3xl border border-slate-100 shadow-xl text-center space-y-6">
                <AlertCircle className="mx-auto text-rose-500" size={48}/>
                <h3 className="text-xl font-black">RELATORIO AT - AR</h3>
                <label className="block bg-rose-600 text-white py-4 rounded-2xl font-black cursor-pointer hover:bg-rose-700 transition-all shadow-lg shadow-rose-100">
                  <input type="file" multiple onChange={handleFailureUpload} className="hidden" />
                  CARREGAR RELATORIO
                </label>
                {failureData.length > 0 && <span className="inline-block bg-green-50 text-green-600 px-4 py-1 rounded-full text-xs font-black uppercase tracking-wider">{failureData.length} falhas importadas</span>}
              </div>
            </div>
            {isReady && <div className="text-center pt-6"><button onClick={() => setActiveTab('dashboard')} className="bg-emerald-600 text-white px-16 py-6 rounded-3xl font-black text-2xl shadow-2xl hover:scale-105 transition-all">INICIAR ANÁLISE</button></div>}
          </div>
        ) : (
          <div className="space-y-6 animate-in fade-in duration-500">
            <div className="bg-white p-5 rounded-3xl shadow-sm border border-slate-100 grid md:grid-cols-4 gap-4">
               <div>
                 <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Referência</span>
                 <select multiple size={1} value={selectedDate.map(String)} onChange={e => {
                   const values = readMultiSelectValues(e);
                   setSelectedDate(prev => {
                     const hadAll = prev.includes('all');
                     if (values.includes('all')) {
                       return values.length > 1 && hadAll ? values.filter(v => v !== 'all') : ['all'];
                     }
                     return values;
                   });
                 }} className="w-full bg-slate-50 p-2 rounded-xl text-sm font-bold border-none outline-none">
                   <option value="all">Todas as Datas</option>
                   {uniqueDates.map(d => <option key={d} value={d}>{d}</option>)}
                 </select>
               </div>
               <div>
                 <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Período</span>
                 <select multiple size={1} value={selectedShift.map(String)} onChange={e => {
                   const values = readMultiSelectValues(e);
                   setSelectedShift(prev => {
                     const hadAll = prev.includes('all');
                     if (values.includes('all')) {
                       const next = values.length > 1 && hadAll ? values.filter(v => v !== 'all') : ['all'];
                       return next.map(v => (v === 'all' ? 'all' : Number(v))).filter(v => v === 'all' || !Number.isNaN(v));
                     }
                     return values.map(v => Number(v)).filter(v => !Number.isNaN(v));
                   });
                 }} className="w-full bg-slate-50 p-2 rounded-xl text-sm font-bold border-none outline-none">
                   <option value="all">Todos os Turnos</option>
                   <option value="1">1° Turno (05:00 - 17:29)</option>
                   <option value="2">2° Turno (17:30 - 04:59)</option>
                 </select>
               </div>
               <div>
                 <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Linha</span>
                 <select multiple size={1} value={selectedLine.map(String)} onChange={e => {
                   const values = readMultiSelectValues(e);
                   setSelectedLine(prev => {
                     const hadAll = prev.includes('all');
                     if (values.includes('all')) {
                       return values.length > 1 && hadAll ? values.filter(v => v !== 'all') : ['all'];
                     }
                     return values;
                   });
                 }} className="w-full bg-slate-50 p-2 rounded-xl text-sm font-bold border-none outline-none">
                   <option value="all">Todas as Linhas</option>
                   {uniqueLines.map(l => <option key={l} value={l}>{l}</option>)}
                 </select>
               </div>
               <div>
                 <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Pesquisa</span>
                 <input type="text" value={searchTerm} onChange={e => setSearchTerm(e.target.value)} placeholder="Buscar..." className="w-full bg-slate-50 p-2 rounded-xl text-sm font-bold border-none outline-none" />
               </div>
            </div>

            <div className="flex flex-wrap gap-3">
              <button onClick={handleExportExcel} className="bg-slate-900 text-white px-4 py-2 rounded-xl font-black text-xs tracking-wide hover:bg-slate-800 transition-all">EXPORTAR EXCEL</button>
              <button onClick={handleExportPdf} className="bg-white text-slate-900 px-4 py-2 rounded-xl font-black text-xs tracking-wide border border-slate-200 hover:bg-slate-50 transition-all">EXPORTAR PDF</button>
              <button onClick={handleExportCsv} className="bg-white text-slate-900 px-4 py-2 rounded-xl font-black text-xs tracking-wide border border-slate-200 hover:bg-slate-50 transition-all">EXPORTAR CSV</button>
              <button onClick={handleExportPng} className="bg-white text-slate-900 px-4 py-2 rounded-xl font-black text-xs tracking-wide border border-slate-200 hover:bg-slate-50 transition-all">EXPORTAR PNG</button>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {[
                { l: 'Meta Total', v: stats.tm, c: 'blue' },
                { l: 'Produzido', v: stats.tp, c: 'indigo', d: stats.diff },
                { l: 'Eficiência', v: `${stats.eff.toFixed(1)}%`, c: stats.eff >= 95 ? 'emerald' : 'amber' },
                { l: 'Diferença', v: stats.diff, c: stats.diff >= 0 ? 'emerald' : 'rose' }
              ].map((s, i) => (
                <div key={i} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
                  <span className="text-[10px] font-black text-slate-400 uppercase mb-2 block">{s.l}</span>
                  <div className="flex justify-between items-baseline">
                    <h3 className={`text-2xl font-black text-slate-900 tracking-tighter`}>{typeof s.v === 'number' ? s.v.toLocaleString() : s.v}</h3>
                    {s.d !== undefined && <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${s.d >= 0 ? 'bg-green-50 text-green-600' : 'bg-red-50 text-red-600'}`}>{s.d >= 0 ? '+' : ''}{s.d.toLocaleString()}</span>}
                  </div>
                </div>
              ))}
            </div>

            <div ref={tableRef} className="bg-white rounded-3xl border border-slate-100 overflow-hidden shadow-sm">
              {isExportingPng && (
                <style>
                  {`.export-table{border-collapse:collapse;} .export-table th,.export-table td{border:1px solid #e2e8f0;}`}
                </style>
              )}
              <table className={`w-full text-left ${isExportingPng ? 'export-table' : ''}`}>
                <thead>
                  <tr className="bg-slate-50/50 text-[10px] font-black text-slate-400 uppercase tracking-widest">
                    <th className="px-6 py-4">Data</th>
                    <th className="px-6 py-4">Linha</th>
                    <th className="px-6 py-4">Turno</th>
                    <th className="px-6 py-4">Material / Detalhes</th>
                    <th className="px-6 py-4 text-right">Meta</th>
                    <th className="px-6 py-4 text-right">Produção Real</th>
                    <th className="px-6 py-4 text-right">Falhas</th>
                    <th className="px-6 py-4 text-right">Yield</th>
                    <th className="px-6 py-4 text-right">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filtered.map((r, idx) => (
                    <tr key={idx} className="hover:bg-blue-50/20 transition-all group">
                      <td className="px-6 py-5 text-[11px] font-bold text-slate-400">{r.date}</td>
                      <td className="px-6 py-5">
                         <span className="bg-slate-100 text-slate-600 px-2.5 py-1 rounded-lg text-[10px] font-black uppercase inline-flex items-center justify-center text-center min-w-[56px]">{r.line}</span>
                      </td>
                      <td className="px-6 py-5">
                        <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black text-white inline-flex items-center justify-center text-center min-w-[72px] ${r.shift === 1 ? 'bg-blue-500' : 'bg-indigo-600'}`}>{r.shift}° TURNO</span>
                      </td>
                      <td className="px-6 py-5">
                        <div className="flex flex-col">
                          <span className="text-sm font-black text-slate-800 tracking-tight">{r.material}</span>
                          <span className={`text-[10px] font-bold text-slate-400 uppercase ${isExportingPng ? 'whitespace-normal break-words' : 'truncate max-w-[200px]'}`}>{r.product}</span>
                        </div>
                      </td>
                      <td className="px-6 py-5 text-right font-bold text-slate-400 text-sm">{r.meta.toLocaleString()}</td>
                      <td className="px-6 py-5 text-right font-black text-blue-600 text-sm">{r.produced.toLocaleString()}</td>
                      <td className="px-6 py-5 text-right font-bold text-slate-400 text-sm">{r.failures.toLocaleString()}</td>
                      <td className="px-6 py-5 text-right font-black text-emerald-600 text-sm">{r.yield.toFixed(1)}%</td>
                      
                      <td className="px-6 py-5 text-right">
                        <div className="flex flex-col items-end gap-1">
                          <span className={`text-xs font-black ${r.efficiency >= 100 ? 'text-green-600' : r.efficiency >= 90 ? 'text-blue-600' : 'text-rose-600'}`}>{r.efficiency.toFixed(1)}%</span>
                          <div className="w-16 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                            <div className={`h-full ${r.efficiency >= 100 ? 'bg-green-500' : r.efficiency >= 90 ? 'bg-blue-500' : 'bg-rose-500'}`} style={{ width: `${Math.min(r.efficiency, 100)}%` }} />
                          </div>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtered.length === 0 && <div className="py-20 text-center font-bold text-slate-300 uppercase tracking-widest italic">Nenhum dado consolidado</div>}
            </div>
          </div>
        )}
      </main>

      {isLoading && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center">
          <div className="bg-white p-10 rounded-[40px] shadow-2xl flex flex-col items-center gap-6">
            <div className="w-16 h-16 border-8 border-slate-100 border-t-blue-600 rounded-full animate-spin" />
            <p className="font-black text-slate-900 uppercase tracking-widest animate-pulse">{loadingMessage}</p>
          </div>
        </div>
      )}
    </div>
  );
};

const isReady = (p: any, a: any) => p.length > 0 && a.length > 0;

export default App;
















