/**
 * 極簡 Google Apps Script 環境模擬。
 * 目的是在本機驗證後端邏輯，不需要真的連上 Google。
 */
const chain = new Proxy(function () {}, {
  get: (t, p) => (p === 'then' ? undefined : chain),
  apply: () => chain,
});

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
  setFontWeight() { return this; }
  setBackground() { return this; }
  setNumberFormat() { return this; }
  setNote() { return this; }
  insertCheckboxes() {
    // 真實試算表插入核取方塊時，空白會變成 false
    for (let i = 0; i < this.nr; i++) {
      for (let j = 0; j < this.nc; j++) {
        if (this.sheet._get(this.r + i, this.c + j) === '') this.sheet._set(this.r + i, this.c + j, false);
      }
    }
    return this;
  }
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  _get(r, c) { const v = (this.data[r - 1] || [])[c - 1]; return v === undefined ? '' : v; }
  _set(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() {
    let last = 0;
    this.data.forEach((row, i) => { if (row.some(v => v !== '' && v != null && v !== false)) last = i + 1; });
    return last;
  }
  getLastColumn() { return Math.max(0, ...this.data.map(r => r.length)); }
  getMaxColumns() { return Math.max(1, this.getLastColumn()); }
  deleteColumns(start, num) { this.data = this.data.map(r => r.slice(0, start - 1).concat(r.slice(start - 1 + num))); }
  clear() { this.data = []; return this; }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  copyTo(target) {
    const copy = new Sheet(this.name + ' 的副本');
    copy.data = this.data.map(r => r.slice());
    copy.setName = n => { delete target.sheets[copy.name]; copy.name = n; target.sheets[n] = copy; return copy; };
    target.sheets[copy.name] = copy;
    return copy;
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(vals) { const r = this.getLastRow() + 1; vals.forEach((v, j) => this._set(r, j + 1, v)); }
  setColumnWidth() { return this; }
  setFrozenRows() { return this; }
}

const ss = {
  sheets: {},
  getSheetByName(n) { return this.sheets[n] || null; },
  insertSheet(n) { return (this.sheets[n] = new Sheet(n)); },
  setSpreadsheetTimeZone() {},
};

const mails = [];
const dialogs = [];
// 測試可改 ui.answer 來模擬使用者在確認視窗按了「取消」
const ui = {
  answer: 'OK',
  createMenu: () => chain,
  showModalDialog() {},
  alert(title, msg, buttons) { dialogs.push({ title, msg, buttons }); return ui.answer; },
  ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
  Button: { OK: 'OK', CANCEL: 'CANCEL' },
};

global.SpreadsheetApp = { getActiveSpreadsheet: () => ss, flush() {}, getUi: () => ui };
global.Logger = { log: () => {} };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => s }) };
global.HtmlService = { createHtmlOutput: () => chain };
global.MailApp = { sendEmail: o => mails.push(o) };
global.ScriptApp = {
  getProjectTriggers: () => [],
  newTrigger: () => chain,
  WeekDay: { SUNDAY: 0, MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3, THURSDAY: 4, FRIDAY: 5, SATURDAY: 6 },
};
global.Utilities = {
  formatDate(d, tz, fmt) {
    if (fmt === 'MMdd-HHmm') {
      const p = new Intl.DateTimeFormat('en-CA', {
        timeZone: tz, month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false,
      }).formatToParts(d).reduce((a, x) => ((a[x.type] = x.value), a), {});
      return `${p.month}${p.day}-${p.hour}${p.minute}`;
    }
    const p = new Intl.DateTimeFormat('en-CA', {
      timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    }).formatToParts(d).reduce((a, x) => ((a[x.type] = x.value), a), {});
    return fmt === 'yyyy-MM-dd'
      ? `${p.year}-${p.month}-${p.day}`
      : `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
  },
};

module.exports = { ss, mails, dialogs, ui };
