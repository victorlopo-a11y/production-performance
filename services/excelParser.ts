import * as XLSX from 'xlsx';
import { ProductionPlanRow, ActualProductionRecord } from '../types';

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
