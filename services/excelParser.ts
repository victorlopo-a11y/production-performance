import * as XLSX from 'xlsx';
import { ProductionPlanRow, ActualProductionRecord, FailureReportRecord } from '../types';

type WorkerRequest =
  | { id: string; kind: 'plan'; buffer: ArrayBuffer }
  | { id: string; kind: 'actual'; buffer: ArrayBuffer }
  | { id: string; kind: 'failure'; buffer: ArrayBuffer };
type WorkerResponse =
  | { id: string; ok: true; result: unknown }
  | { id: string; ok: false; error: string };

let workerSingleton: Worker | null = null;
let workerSeq = 0;
const workerPending = new Map<string, { resolve: (v: unknown) => void; reject: (e: Error) => void }>();

function getWorker(): Worker | null {
  if (typeof Worker === 'undefined') return null;
  if (workerSingleton) return workerSingleton;

  try {
    workerSingleton = new Worker(new URL('./excelParseWorker.ts', import.meta.url), { type: 'module' });
    workerSingleton.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const msg = event.data;
      const pending = workerPending.get(msg.id);
      if (!pending) return;
      workerPending.delete(msg.id);
      if (msg.ok) {
        pending.resolve(msg.result);
      } else {
        pending.reject(new Error(msg.error || 'Erro ao processar arquivo'));
      }
    };
    workerSingleton.onerror = (err) => {
      workerPending.forEach(({ reject }) => reject(new Error(`Worker error: ${String(err)}`)));
      workerPending.clear();
      workerSingleton?.terminate();
      workerSingleton = null;
    };
    return workerSingleton;
  } catch {
    workerSingleton = null;
    return null;
  }
}

async function runInWorker<T>(payload: Omit<WorkerRequest, 'id'>): Promise<T> {
  const worker = getWorker();
  if (!worker) throw new Error('Worker indisponível');

  const id = `excel-${Date.now()}-${workerSeq++}`;
  const message: WorkerRequest = { id, ...payload } as WorkerRequest;

  return new Promise<T>((resolve, reject) => {
    const timeout = setTimeout(() => {
      workerPending.delete(id);
      reject(new Error('Tempo excedido ao processar o arquivo. Tente novamente.'));
    }, 120_000);

    workerPending.set(id, {
      resolve: (value) => {
        clearTimeout(timeout);
        resolve(value as T);
      },
      reject: (err) => {
        clearTimeout(timeout);
        reject(err);
      },
    });
    const transfer: Transferable[] = [];
    if ('buffer' in message && message.buffer instanceof ArrayBuffer) {
      transfer.push(message.buffer);
    }
    // Transfer the ArrayBuffer to avoid a costly structured clone of large XLSM files.
    worker.postMessage(message, transfer);
  });
}

export const parseExcelDate = (serial: any): string => {
  if (!serial) return '';
  if (serial instanceof Date) {
    return serial.toLocaleDateString('pt-BR');
  }
  if (typeof serial === 'string') {
    const trimmed = serial.trim();
    const dateMatch = trimmed.match(/^(\d{2}\/\d{2}\/\d{4})/);
    if (dateMatch) return dateMatch[1];
    
    if (trimmed.match(/^\d{4}-\d{2}-\d{2}/)) {
      const d = new Date(trimmed);
      return isNaN(d.getTime()) ? trimmed : d.toLocaleDateString('pt-BR');
    }
    return trimmed;
  }
  try {
    if (typeof serial === 'number') {
      const date = XLSX.SSF.parse_date_code(serial);
      const d = new Date(date.y, date.m - 1, date.d);
      return d.toLocaleDateString('pt-BR');
    }
  } catch (e) {
    console.error("Erro ao processar data:", serial, e);
  }
  return String(serial);
};

export const parseExcelTime = (val: any): string => {
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2, '0')}:${String(val.getMinutes()).padStart(2, '0')}:${String(val.getSeconds()).padStart(2, '0')}`;
  }
  if (typeof val === 'number') {
    const timeFrac = val % 1;
    const totalSeconds = Math.round(timeFrac * 24 * 3600);
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  }
  if (typeof val === 'string') {
    const timeMatch = val.match(/(\d{1,2}:\d{2}(?::\d{2})?)/);
    if (timeMatch) {
      let t = timeMatch[1];
      if (t.length <= 5) t += ":00";
      const parts = t.split(':');
      return `${parts[0].padStart(2, '0')}:${parts[1]}:${parts[2]}`;
    }
  }
  return '00:00:00';
};

export const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const base64 = (reader.result as string).split(',')[1];
      resolve(base64);
    };
    reader.onerror = error => reject(error);
  });
};

export const parseProductionPlanExcel = async (file: File): Promise<ProductionPlanRow[]> => {
  try {
    const buffer = await file.arrayBuffer();
    return await runInWorker<ProductionPlanRow[]>({ kind: 'plan', buffer });
  } catch {
    // fallback to main thread parser
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        const results: ProductionPlanRow[] = [];
        let headerRowIndex = -1;
        let dateRowIndex = -1;

        for (let i = 0; i < Math.min(rows.length, 100); i++) {
          const rowValues = rows[i].map(v => String(v).toUpperCase().trim());
          const hasProg = rowValues.some(v => v === 'PROG' || v.includes('PROGRAMADO'));
          if (hasProg) {
            headerRowIndex = i;
            for (let j = i - 1; j >= 0; j--) {
              const upperRow = rows[j].map(v => String(v).trim());
              const dateCount = upperRow.filter(v => v.includes('/') || (typeof rows[j][upperRow.indexOf(v)] === 'number' && rows[j][upperRow.indexOf(v)] > 40000)).length;
              if (dateCount >= 1) {
                dateRowIndex = j;
                break;
              }
            }
            break;
          }
        }

        if (headerRowIndex === -1) return resolve([]);

        const headers = rows[headerRowIndex].map(v => String(v).toUpperCase().trim());
        const dateRow = dateRowIndex !== -1 ? rows[dateRowIndex] : [];

        const lineIdx = headers.findIndex(v => v.includes('LINHA') || v.includes('LINE'));
        const shiftIdx = headers.findIndex(v => v.includes('TURNO') || v.includes('SHIFT'));
        const codeIdx = headers.findIndex(v => v.includes('CÓDIGO') || v.includes('CODIGO') || v.includes('CODE'));
        const productIdx = headers.findIndex(v => v.includes('PRODUTO') || v.includes('PRODUCT'));

        const progColumns: { index: number; date: string }[] = [];
        headers.forEach((h, idx) => {
          if (h === 'PROG' || h.includes('PROGRAMADO')) {
            let dateStr = '';
            for (let j = idx; j >= 0; j--) {
              if (dateRow[j]) {
                const potentialDate = parseExcelDate(dateRow[j]);
                if (potentialDate.includes('/')) {
                  dateStr = potentialDate;
                  break;
                }
              }
            }
            if (dateStr) progColumns.push({ index: idx, date: dateStr });
          }
        });

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
          const row = rows[i];
          const code = String(row[codeIdx] || '').trim();
          if (!code || code === '' || code.toLowerCase() === 'total') continue;

          const line = String(row[lineIdx] || '').trim();
          const rawShift = String(row[shiftIdx] || '');
          const shift = parseInt(rawShift.replace(/\D/g, '')) || 1;
          const product = String(row[productIdx] || '').trim();

          progColumns.forEach(progCol => {
            const metaValue = row[progCol.index];
            if (metaValue === '' || metaValue === null || metaValue === undefined) return;
            const parsedMeta = typeof metaValue === 'number' ? metaValue : parseFloat(String(metaValue).replace(/\./g, '').replace(',', '.'));
            const meta = Math.round(parsedMeta);
            if (!isNaN(meta) && meta > 0) {
              results.push({
                line: line || 'LINHA DESCONHECIDA',
                shift,
                code,
                product: product || 'PRODUTO NÃO INFORMADO',
                date: progCol.date,
                meta
              });
            }
          });
        }
        resolve(results);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const parseActualProductionExcel = async (file: File): Promise<ActualProductionRecord[]> => {
  try {
    const buffer = await file.arrayBuffer();
    return await runInWorker<ActualProductionRecord[]>({ kind: 'actual', buffer });
  } catch {
    // fallback to main thread parser
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const exportSheetName = workbook.SheetNames.find(name => name.toLowerCase().includes('export'));
        const sheetName = exportSheetName || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        
        const normalizeHeader = (val: string) =>
          val
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();

        const results: ActualProductionRecord[] = rows.map(row => {
          const findVal = (keys: string[]) => {
            const targets = keys.map(k => normalizeHeader(k));
            const headers = Object.keys(row);
            for (const k of headers) {
              const nk = normalizeHeader(k);
              if (targets.some(target => nk === target)) return row[k];
            }
            for (const k of headers) {
              const nk = normalizeHeader(k);
              if (targets.some(target => nk.includes(target))) return row[k];
            }
            return null;
          };

          // Tenta encontrar a coluna de linha/posto de trabalho
          const line = String(findVal(['linha', 'line', 'centro', 'posto', 'recurso', 'maquina', 'deposito', 'centro de trabalho']) || '').trim();
          const material = String(findVal(['material', 'codigo', 'cod']) || '').trim();
          const quantityVal = findVal(['qtd. um registro', 'qtd um registro', 'quantidade', 'qtd', 'quant', 'produzido']);
          const parsedQuantity = typeof quantityVal === 'number' ? quantityVal : parseFloat(String(quantityVal).replace(/\./g, '').replace(',', '.')) || 0;
          const quantity = Math.round(parsedQuantity);
          
          const dateVal = findVal(['data de lancamento', 'data lancamento', 'data de entrada', 'data', 'dia']);
          const dateStr = parseExcelDate(dateVal);
          
          const timeVal = findVal(['hora do registro', 'hora', 'horario', 'time']);
          let timeStr = parseExcelTime(timeVal);
          
          // Fallback de horário se estiver embutido na data
          if (timeStr === '00:00:00' && dateVal) {
             timeStr = parseExcelTime(dateVal);
          }
          
          const shiftData = calculateShiftAndProductionDay(dateStr, timeStr);
          
          return {
            line,
            material,
            quantity,
            date: dateStr,
            time: timeStr,
            shift: shiftData.shift,
            productionDay: shiftData.productionDay
          };
        }).filter(r => {
          if (!r.material) return false;
          const material = r.material.toLowerCase();
          if (material === 'total') return false;
          if (material.startsWith('mon')) return false;
          return r.quantity > 0;
        });
        
        resolve(results);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const parseFailureReportExcel = async (file: File): Promise<FailureReportRecord[]> => {
  try {
    const buffer = await file.arrayBuffer();
    return await runInWorker<FailureReportRecord[]>({ kind: 'failure', buffer });
  } catch {
    // fallback to main thread parser
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const normalizeSheet = (val: string) =>
          val
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');
        const reportSheetName = workbook.SheetNames.find(name => normalizeSheet(name).includes('relatorio at'));
        const sheetName = reportSheetName || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rowsMatrix: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        const normalizeHeader = (val: string) =>
          val
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();

        const headerRowIndex = rowsMatrix.findIndex(row => {
          const cols = row.map(v => normalizeHeader(String(v || '')));
          const hasQty = cols.some(c => c === 'quantidade' || c.includes('quantidade') || c.startsWith('qtd'));
          const hasOrigin = cols.some(c => c === 'origem' || c.includes('origem'));
          const hasLineOrProduct = cols.some(c => c.includes('linha') || c.includes('produto') || c.includes('material') || c.includes('codigo'));
          return hasQty && (hasOrigin || hasLineOrProduct);
        });

        const headerIndex = headerRowIndex === -1 ? 0 : headerRowIndex;
        const headers = rowsMatrix[headerIndex] || [];

        const findIndex = (keys: string[]) => {
          const targets = keys.map(k => normalizeHeader(k));
          const normalized = headers.map(h => normalizeHeader(String(h || '')));
          let idx = normalized.findIndex(h => targets.some(t => h === t));
          if (idx === -1) {
            idx = normalized.findIndex(h => targets.some(t => h.includes(t)));
          }
          return idx;
        };

        const lineIdx = findIndex(['processo linha', 'linha', 'line', 'centro', 'posto', 'recurso', 'maquina', 'deposito', 'centro de trabalho']);
        const materialIdx = findIndex(['material', 'codigo', 'cod', 'produto']);
        const originIdx = findIndex(['origem', 'origem processo', 'origem da falha']);
        const qtyIdx = findIndex(['quantidade', 'qtd falha', 'qtde falha', 'falha', 'defeito', 'qtd', 'quant']);
        const dateIdx = findIndex(['data de lancamento', 'data lancamento', 'data de entrada', 'data', 'dia']);
        const timeIdx = findIndex(['hora do registro', 'hora', 'horario', 'time']);

        const dataRows = rowsMatrix.slice(headerIndex + 1);
        const results: FailureReportRecord[] = dataRows.map(row => {
          const line = String(lineIdx >= 0 ? row[lineIdx] : '').trim();
          const material = String(materialIdx >= 0 ? row[materialIdx] : '').trim();
          const origin = String(originIdx >= 0 ? row[originIdx] : '').trim();
          const quantityVal = qtyIdx >= 0 ? row[qtyIdx] : '';
          const parsedQuantity = typeof quantityVal === 'number' ? quantityVal : parseFloat(String(quantityVal).replace(/\./g, '').replace(',', '.')) || 0;
          const quantity = Math.round(parsedQuantity);

          const dateVal = dateIdx >= 0 ? row[dateIdx] : '';
          const dateStr = parseExcelDate(dateVal);

          const timeVal = timeIdx >= 0 ? row[timeIdx] : '';
          let timeStr = parseExcelTime(timeVal);
          if (timeStr === '00:00:00' && dateVal) {
            timeStr = parseExcelTime(dateVal);
          }

          const shiftData = calculateShiftAndProductionDay(dateStr, timeStr);

          return {
            line,
            material: material || 'NA',
            origin,
            quantity,
            date: dateStr,
            time: timeStr,
            shift: shiftData.shift,
            productionDay: shiftData.productionDay
          };
        }).filter(r => {
          if (r.quantity <= 0) return false;
          const material = r.material.toLowerCase();
          const line = r.line.toLowerCase();
          if (material === 'total' || line === 'total') return false;
          if (material.startsWith('mon')) return false;
          if (!r.line && !r.material) return false;
          return true;
        });

        resolve(results);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export function calculateShiftAndProductionDay(dateStr: string, timeStr: string): { shift: number; productionDay: string } {
  try {
    const timeParts = timeStr.split(':');
    const h = parseInt(timeParts[0]) || 0;
    const m = parseInt(timeParts[1]) || 0;
    const s = parseInt(timeParts[2]) || 0;
    const totalSeconds = h * 3600 + m * 60 + s;

    const s1Start = 5 * 3600; // 05:00:00
    const s1End = 17 * 3600 + 29 * 60 + 59; // 17:29:59

    // 1o Turno: 05:00:00 as 17:29:59
    if (totalSeconds >= s1Start && totalSeconds <= s1End) {
      return { shift: 1, productionDay: dateStr };
    }
    // 2o Turno: 17:30:00 as 04:59:59 (apos meia-noite volta um dia)
    if (totalSeconds < s1Start) {
      const parts = dateStr.split('/').map(Number);
      if (parts.length === 3) {
        const [d, mo, y] = parts;
        const prev = new Date(y, (mo || 1) - 1, d || 1);
        if (!isNaN(prev.getTime())) {
          prev.setDate(prev.getDate() - 1);
          const dd = String(prev.getDate()).padStart(2, '0');
          const mm = String(prev.getMonth() + 1).padStart(2, '0');
          const yyyy = String(prev.getFullYear());
          return { shift: 2, productionDay: `${dd}/${mm}/${yyyy}` };
        }
      }
    }
    return { shift: 2, productionDay: dateStr };
  } catch (e) {
    return { shift: 1, productionDay: dateStr };
  }
}
