import * as XLSX from 'xlsx';
import type { ProductionPlanRow, ActualProductionRecord, FailureReportRecord } from '../types';

type RequestMessage =
  | { id: string; kind: 'plan'; buffer: ArrayBuffer }
  | { id: string; kind: 'actual'; buffer: ArrayBuffer }
  | { id: string; kind: 'failure'; buffer: ArrayBuffer };

type ResponseMessage =
  | { id: string; ok: true; result: unknown }
  | { id: string; ok: false; error: string };

self.onmessage = async (event: MessageEvent<RequestMessage>) => {
  const payload = event.data;
  try {
    if (payload.kind === 'plan') {
      const result = parseProductionPlanExcelCore(payload.buffer);
      postMessage({ id: payload.id, ok: true, result } satisfies ResponseMessage);
      return;
    }
    if (payload.kind === 'actual') {
      const result = parseActualProductionExcelCore(payload.buffer);
      postMessage({ id: payload.id, ok: true, result } satisfies ResponseMessage);
      return;
    }
    const result = parseFailureReportExcelCore(payload.buffer);
    postMessage({ id: payload.id, ok: true, result } satisfies ResponseMessage);
  } catch (err) {
    postMessage({
      id: payload.id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    } satisfies ResponseMessage);
  }
};

function parseExcelDate(serial: any): string {
  if (!serial) return '';
  if (serial instanceof Date) return serial.toLocaleDateString('pt-BR');
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
  if (typeof serial === 'number') {
    const date = XLSX.SSF.parse_date_code(serial);
    const d = new Date(date.y, date.m - 1, date.d);
    return d.toLocaleDateString('pt-BR');
  }
  return String(serial);
}

function parseExcelTime(val: any): string {
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
      if (t.length <= 5) t += ':00';
      const parts = t.split(':');
      return `${parts[0].padStart(2, '0')}:${parts[1]}:${parts[2]}`;
    }
  }
  return '00:00:00';
}

function normalizeHeader(value: string): string {
  return String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function calculateShiftAndProductionDay(dateStr: string, timeStr: string): { shift: number; productionDay: string } {
  try {
    const timeParts = timeStr.split(':');
    const h = parseInt(timeParts[0]) || 0;
    const m = parseInt(timeParts[1]) || 0;
    const s = parseInt(timeParts[2]) || 0;
    const totalSeconds = h * 3600 + m * 60 + s;

    const s1Start = 5 * 3600; // 05:00:00
    const s1End = 17 * 3600 + 29 * 60 + 59; // 17:29:59

    if (totalSeconds >= s1Start && totalSeconds <= s1End) return { shift: 1, productionDay: dateStr };

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
  } catch {
    return { shift: 1, productionDay: dateStr };
  }
}

function readSheetMatrix(buffer: ArrayBuffer): any[][] {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][];
}

function parseProductionPlanExcelCore(buffer: ArrayBuffer): ProductionPlanRow[] {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const results: ProductionPlanRow[] = [];
  let headerRowIndex = -1;
  let dateRowIndex = -1;

  for (let i = 0; i < Math.min(rows.length, 100); i++) {
    const rowValues = rows[i].map((v) => normalizeHeader(String(v || '')));
    const hasProg = rowValues.some((v) => v === 'PROG' || v.includes('PROGRAMADO'));
    if (hasProg) {
      headerRowIndex = i;
      for (let j = i - 1; j >= 0; j--) {
        const upperRow = rows[j].map((v) => String(v).trim());
        const dateCount = upperRow.filter(
          (v) => v.includes('/') || (typeof rows[j][upperRow.indexOf(v)] === 'number' && rows[j][upperRow.indexOf(v)] > 40000)
        ).length;
        if (dateCount >= 1) {
          dateRowIndex = j;
          break;
        }
      }
      break;
    }
  }

  if (headerRowIndex === -1) return [];

  const headersRaw = rows[headerRowIndex].map((v) => String(v || ''));
  const headers = headersRaw.map((v) => normalizeHeader(v));
  const dateRow = dateRowIndex !== -1 ? rows[dateRowIndex] : [];

  const lineIdx = headers.findIndex((v) => v.includes('LINHA') || v.includes('LINE'));
  const shiftIdx = headers.findIndex((v) => v.includes('TURNO') || v.includes('SHIFT'));
  const codeIdx = headers.findIndex((v) => v.includes('CODIGO') || v.includes('CODE'));
  const productIdx = headers.findIndex((v) => v.includes('PRODUTO') || v.includes('PRODUCT'));

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
    const code = String(codeIdx >= 0 ? row[codeIdx] : '').trim();
    if (!code || code.toLowerCase() === 'total') continue;

    const line = String(lineIdx >= 0 ? row[lineIdx] : '').trim();
    const rawShift = String(shiftIdx >= 0 ? row[shiftIdx] : '');
    const shift = parseInt(rawShift.replace(/\D/g, '')) || 1;
    const product = String(productIdx >= 0 ? row[productIdx] : '').trim();

    progColumns.forEach((progCol) => {
      const metaValue = row[progCol.index];
      if (metaValue === '' || metaValue == null) return;
      const parsedMeta = typeof metaValue === 'number' ? metaValue : parseFloat(String(metaValue).replace(/\./g, '').replace(',', '.'));
      const meta = Math.round(parsedMeta);
      if (!isNaN(meta) && meta > 0) {
        results.push({
          line: line || 'LINHA DESCONHECIDA',
          shift,
          code,
          product: product || 'PRODUTO NÃƒO INFORMADO',
          date: progCol.date,
          meta,
        });
      }
    });
  }

  return results;
}

function parseActualProductionExcelCore(buffer: ArrayBuffer): ActualProductionRecord[] {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const exportSheetName = workbook.SheetNames.find((name) => name.toLowerCase().includes('export'));
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

  const results: ActualProductionRecord[] = rows
    .map((row) => {
      const findVal = (keys: string[]) => {
        const targets = keys.map((k) => normalizeHeader(k));
        const headers = Object.keys(row);
        for (const k of headers) {
          const nk = normalizeHeader(k);
          if (targets.some((target) => nk === target)) return row[k];
        }
        for (const k of headers) {
          const nk = normalizeHeader(k);
          if (targets.some((target) => nk.includes(target))) return row[k];
        }
        return null;
      };

      const line = String(findVal(['linha', 'line', 'centro', 'posto', 'recurso', 'maquina', 'deposito', 'centro de trabalho']) || '').trim();
      const material = String(findVal(['material', 'codigo', 'cod']) || '').trim();
      const quantityVal = findVal(['qtd. um registro', 'qtd um registro', 'quantidade', 'qtd', 'quant', 'produzido']);
      const parsedQuantity = typeof quantityVal === 'number' ? quantityVal : parseFloat(String(quantityVal).replace(/\./g, '').replace(',', '.')) || 0;
      const quantity = Math.round(parsedQuantity);

      const dateVal = findVal(['data de lancamento', 'data lancamento', 'data de entrada', 'data', 'dia']);
      const dateStr = parseExcelDate(dateVal);

      const timeVal = findVal(['hora do registro', 'hora', 'horario', 'time']);
      let timeStr = parseExcelTime(timeVal);
      if (timeStr === '00:00:00' && dateVal) timeStr = parseExcelTime(dateVal);

      const shiftData = calculateShiftAndProductionDay(dateStr, timeStr);
      return { line, material, quantity, date: dateStr, time: timeStr, shift: shiftData.shift, productionDay: shiftData.productionDay };
    })
    .filter((r) => {
      if (!r.material) return false;
      const material = r.material.toLowerCase();
      if (material === 'total') return false;
      if (material.startsWith('mon')) return false;
      return r.quantity > 0;
    });

  return results;
}

function parseFailureReportExcelCore(buffer: ArrayBuffer): FailureReportRecord[] {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  const normalizeSheet = (val: string) =>
    val
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  const reportSheetName = workbook.SheetNames.find((name) => normalizeSheet(name).includes('relatorio at'));
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

  const headerRowIndex = rowsMatrix.findIndex((row) => {
    const cols = row.map((v) => normalizeHeader(String(v || '')));
    const hasQty = cols.some((c) => c === 'quantidade' || c.includes('quantidade') || c.startsWith('qtd'));
    const hasOrigin = cols.some((c) => c === 'origem' || c.includes('origem'));
    const hasLineOrProduct = cols.some((c) => c.includes('linha') || c.includes('produto') || c.includes('material') || c.includes('codigo'));
    return hasQty && (hasOrigin || hasLineOrProduct);
  });

  const headerIndex = headerRowIndex === -1 ? 0 : headerRowIndex;
  const headers = rowsMatrix[headerIndex] || [];

  const findIndex = (keys: string[]) => {
    const targets = keys.map((k) => normalizeHeader(k));
    const normalized = headers.map((h) => normalizeHeader(String(h || '')));
    let idx = normalized.findIndex((h) => targets.some((t) => h === t));
    if (idx === -1) idx = normalized.findIndex((h) => targets.some((t) => h.includes(t)));
    return idx;
  };

  const lineIdx = findIndex(['processo linha', 'linha', 'line', 'centro', 'posto', 'recurso', 'maquina', 'deposito', 'centro de trabalho']);
  const materialIdx = findIndex(['material', 'codigo', 'cod', 'produto']);
  const originIdx = findIndex(['origem', 'origem processo', 'origem da falha']);
  const qtyIdx = findIndex(['quantidade', 'qtd falha', 'qtde falha', 'falha', 'defeito', 'qtd', 'quant']);
  const dateIdx = findIndex(['data de lancamento', 'data lancamento', 'data de entrada', 'data', 'dia']);
  const timeIdx = findIndex(['hora do registro', 'hora', 'horario', 'time']);

  const dataRows = rowsMatrix.slice(headerIndex + 1);
  const results: FailureReportRecord[] = dataRows
    .map((row) => {
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
      if (timeStr === '00:00:00' && dateVal) timeStr = parseExcelTime(dateVal);

      const shiftData = calculateShiftAndProductionDay(dateStr, timeStr);
      return { line, material: material || 'NA', origin, quantity, date: dateStr, time: timeStr, shift: shiftData.shift, productionDay: shiftData.productionDay };
    })
    .filter((r) => {
      if (r.quantity <= 0) return false;
      const material = r.material.toLowerCase();
      const line = r.line.toLowerCase();
      if (material === 'total' || line === 'total') return false;
      if (material.startsWith('mon')) return false;
      if (!r.line && !r.material) return false;
      return true;
    });

  return results;
}
