// Imitação mínima dos serviços do Google (planilha em memória) para testar Code.gs fora do Google.
// Não vai para o Apps Script.
(function (g) {
  function Sheet(nome) { this.nome = nome; this.d = []; }
  Sheet.prototype = {
    getName() { return this.nome; },
    deleteRow(r) { this.d.splice(r - 1, 1); },
    getLastRow() { return this.d.length; },
    getLastColumn() { return Math.max(0, ...this.d.map(r => r.length)); },
    getRange(a, b, c, e) {
      if (typeof a !== 'string') return new Range(this, a, b, c || 1, e || 1);
      const m = a.match(/^([A-Z])(\d*)(?::([A-Z])(\d*))?$/); // 'A1', 'A:Z', 'E2:E'
      const col = m[1].charCodeAt(0) - 64, row = +m[2] || 1, col2 = m[3] ? m[3].charCodeAt(0) - 64 : col;
      const row2 = m[4] ? +m[4] : m[3] ? Math.max(this.d.length, row) : row;
      return new Range(this, row, col, row2 - row + 1, col2 - col + 1);
    },
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); },
    appendRow(v) { this.d.push(v.map(x => /^\d+$/.test(x) ? String(Number(x)) : String(x))); }, // como o Google: '0001' vira 1
    setFrozenRows() {},
  };
  function Range(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  Range.prototype = {
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const row = [];
        for (let j = 0; j < this.nc; j++) row.push((this.sh.d[this.r - 1 + i] || [])[this.c - 1 + j] ?? '');
        out.push(row);
      }
      return out;
    },
    getDisplayValues() { return this.getValues(); },
    getValue() { return this.getValues()[0][0]; },
    set(r, c, x) {
      const d = this.sh.d;
      while (d.length < r) d.push([]);
      while (d[r - 1].length < c) d[r - 1].push('');
      d[r - 1][c - 1] = String(x);
    },
    setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.set(this.r + i, this.c + j, x))); return this; },
    setValue(x) { this.set(this.r, this.c, x); return this; },
    clearContent() {
      for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) if (this.sh.d[this.r - 1 + i]) this.set(this.r + i, this.c + j, '');
      return this;
    },
    setNumberFormat() { return this; }, setFontWeight() { return this; }, setDataValidation() { return this; },
  };
  function Planilha() { this.abas = {}; }
  Planilha.prototype = {
    getSheetByName(n) { return this.abas[n] || null; },
    insertSheet(n) { return (this.abas[n] = new Sheet(n)); },
    deleteSheet(s) { delete this.abas[s.nome]; },
    getSheets() { return Object.values(this.abas); },
    getId() { return 'planilha-simulada'; },
    getName() { return 'CEPRESS'; },
    copy(nome) { (g.copias = g.copias || []).push(nome); return this; },
  };
  const ativa = new Planilha();
  const regra = { requireValueInList() { return this; }, setAllowInvalid() { return this; }, build() { return this; } };
  g.Planilha = Planilha;
  g.externas = {};
  g.SpreadsheetApp = {
    getActive: () => ativa, flush() {}, newDataValidation: () => regra,
    openById: id => g.externas[id], openByUrl: url => g.externas[url],
    getUi: () => ({ alert() {}, Button: { OK: 'OK' }, ButtonSet: { OK_CANCEL: 1 },
      prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => ' Nova Pessoa ' }) }),
  };
  g.Session = { getScriptTimeZone: () => 'America/Sao_Paulo' };
  g.chamadasAPI = 0;
  g.Sheets = { Spreadsheets: { Values: { batchGet: (id, { ranges }) => {
    g.chamadasAPI++;
    return { valueRanges: ranges.map(r => {
      const sh = ativa.getSheetByName(r.replace(/^'|'$/g, ''));
      if (!sh) throw new Error('Unable to parse range: ' + r);
      const linhas = sh.d.map(l => { const c = l.slice(); while (c.length && c[c.length - 1] === '') c.pop(); return c; });
      while (linhas.length && !linhas[linhas.length - 1].length) linhas.pop();
      return { values: linhas };
    }) };
  } } } };
  const cache = {};
  g.CacheService = { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  const gatilhos = [];
  g.gatilhos = gatilhos;
  g.ScriptApp = {
    getProjectTriggers: () => gatilhos.slice(),
    WeekDay: { SUNDAY: 'SUNDAY' },
    deleteTrigger: t => gatilhos.splice(gatilhos.indexOf(t), 1),
    newTrigger(fn) {
      const t = { getHandlerFunction: () => fn };
      const passo = { timeBased: () => passo, everyDays: () => passo, everyHours: () => passo, onWeekDay: () => passo, atHour: () => passo, create: () => { gatilhos.push(t); return t; } };
      return passo;
    },
  };
  g.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  g.Utilities = {
    getUuid: () => crypto.randomUUID(),
    formatDate(d, tz, fmt) {
      const p = {};
      new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
        .formatToParts(d).forEach(x => { p[x.type] = x.value; });
      return fmt.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute);
    },
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
