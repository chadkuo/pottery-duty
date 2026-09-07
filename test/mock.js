// 極簡 Google Apps Script 環境模擬，用來驗證後端邏輯
const chain = new Proxy({}, { get: (t, p) => (p === 'then' ? undefined : () => chain) });

class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sheet._get(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getValue() { return this.sheet._get(this.r, this.c); }
  setValues(v) { v.forEach((row, i) => row.forEach((val, j) => this.sheet._set(this.r + i, this.c + j, val))); return this; }
  setValue(v) { this.sheet._set(this.r, this.c, v); return this; }
  setFontWeight() { return this; } setBackground() { return this; }
  setNumberFormat() { return this; } insertCheckboxes() { return this; } setNote() { return this; }
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  _get(r, c) { return (this.data[r - 1] || [])[c - 1] ?? ''; }
  _set(r, c, v) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() { let last = 0; this.data.forEach((row, i) => { if (row.some(v => v !== '' && v != null)) last = i + 1; }); return last; }
  getLastColumn() { return Math.max(0, ...this.data.map(r => r.length)); }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  getDataRange() { return new Range(this, 1, 1, this.getLastRow(), this.getLastColumn()); }
  appendRow(vals) { const r = this.getLastRow() + 1; vals.forEach((v, j) => this._set(r, j + 1, v)); }
  setColumnWidth() { return this; } setFrozenRows() { return this; }
}

const ss = {
  sheets: {},
  getSheetByName(n) { return this.sheets[n] || null; },
  insertSheet(n) { return (this.sheets[n] = new Sheet(n)); },
  setSpreadsheetTimeZone() {},
};

global.SpreadsheetApp = { getActiveSpreadsheet: () => ss, flush() {}, getUi: () => ({ createMenu: () => chain, showModalDialog() {} }) };
global.Logger = { log: (...a) => console.log('[log]', ...a) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => s }) };
global.HtmlService = { createHtmlOutput: () => chain };
global.MailApp = { sendEmail: o => console.log('[mail]', JSON.stringify(o).slice(0, 200)) };
global.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => chain, WeekDay: { TUESDAY: 3, SUNDAY: 1 } };
global.Utilities = {
  formatDate(d, tz, fmt) {
    const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false })
      .formatToParts(d).reduce((a, x) => (a[x.type] = x.value, a), {});
    return fmt === 'yyyy-MM-dd' ? `${p.year}-${p.month}-${p.day}` : `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
  },
};
module.exports = { ss };
