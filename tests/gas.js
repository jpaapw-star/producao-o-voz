// Imitação mínima do Google Apps Script para testar src/Codigo.gs em Node.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function criarAba(nome, linhas) {
  const dados = linhas.map((l) => l.slice());
  const formulas = {};
  let maxRows = Math.max(dados.length, 1);
  const largura = () => Math.max(1, ...dados.map((l) => l.length));
  const celula = (r, c) => (dados[r - 1] && dados[r - 1][c - 1] !== undefined ? dados[r - 1][c - 1] : '');
  const aba = {
    getName: () => nome,
    dados,
    getLastRow: () => {
      for (let r = dados.length; r >= 1; r--) if (dados[r - 1].some((v) => v !== '' && v !== null && v !== undefined)) return r;
      return 0;
    },
    getLastColumn: () => largura(),
    getMaxRows: () => Math.max(maxRows, dados.length),
    insertRowsAfter: (_, n) => { maxRows = Math.max(maxRows, dados.length) + n; },
    insertRowsAfterCount: () => maxRows,
    deleteRow: (r) => { dados.splice(r - 1, 1); },
    appendRow: (l) => { dados.splice(aba.getLastRow(), 0, l.slice()); },
    setFrozenRows: () => {},
    getRange: (r, c, nr = 1, nc = 1) => {
      const ler = (f) => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => f(r + i, c + j)));
      return {
        getValues: () => ler(celula),
        getValue: () => celula(r, c),
        getDisplayValues: () => ler((a, b) => { const v = celula(a, b); return v instanceof Date ? formatar(v, 'America/Sao_Paulo', 'dd/MM/yyyy') : String(v); }),
        getFormulas: () => ler((a, b) => formulas[a + ',' + b] || ''),
        getFormulasR1C1: () => ler((a, b) => formulas[a + ',' + b] || ''),
        setValues: (v) => {
          v.forEach((l, i) => l.forEach((x, j) => {
            while (dados.length < r + i) dados.push([]);
            dados[r + i - 1][c + j - 1] = x;
          }));
        },
        setNumberFormat: () => {},
        setValue: (x) => { while (dados.length < r) dados.push([]); dados[r - 1][c - 1] = x; },
        setFormulaR1C1: (fx) => { for (let i = 0; i < nr; i++) formulas[(r + i) + ',' + c] = fx; },
        copyTo: () => {}
      };
    },
    _formulas: formulas
  };
  return aba;
}

function criarPlanilha(abas) {
  const lista = Object.keys(abas).map((n) => criarAba(n, abas[n]));
  return {
    abas: lista,
    getSheets: () => lista,
    getSheetByName: (n) => lista.find((a) => a.getName() === n) || null,
    insertSheet: (n) => { const a = criarAba(n, []); lista.push(a); return a; },
    getSpreadsheetTimeZone: () => 'America/Sao_Paulo'
  };
}

function formatar(data, tz, formato) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  }).formatToParts(data).map((x) => [x.type, x.value]));
  const hora = p.hour === '24' ? '00' : p.hour;
  return formato.replace(/yyyy|MM|M|dd|HH|H|mm|ss/g, (k) => ({
    yyyy: p.year, MM: p.month, M: String(Number(p.month)), dd: p.day, HH: hora, H: String(Number(hora)), mm: p.minute, ss: p.second
  })[k]);
}

function carregarServidor(planilha, opcoes = {}) {
  const cache = {};
  let uuid = 0;
  const ctx = {
    SEGREDOS: { planilhaId: 'PLANILHA' },
    SpreadsheetApp: {
      openById: () => planilha,
      flush: () => {},
      CopyPasteType: { PASTE_FORMAT: 'f' }
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; },
      putAll: (o) => Object.assign(cache, o), getAll: (ks) => Object.fromEntries(ks.filter((k) => k in cache).map((k) => [k, cache[k]])) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Utilities: {
      formatDate: (d, tz, f) => formatar(opcoes.agora && d.getTime() > Date.now() - 1000 ? opcoes.agora : d, tz, f),
      getUuid: () => ('0000000' + (++uuid)).slice(-8) + '-0000',
      parseDate: (s) => {
        const m = s.match(/(\d+)\/(\d+)\/(\d+) (\d+):(\d+):(\d+)/);
        return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1], +m[4] + 3, +m[5], +m[6]));
      }
    },
    HtmlService: {},
    UrlFetchApp: { fetch: (url, o) => (opcoes.fetch ? opcoes.fetch(url, o) : { getResponseCode: () => 500, getContentText: () => '{}' }) }
  };
  Object.assign(ctx.Utilities, {
    DigestAlgorithm: { MD5: 'md5' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (_, t) => Array.from(require('crypto').createHash('md5').update(t).digest()),
    base64EncodeWebSafe: (b) => Buffer.from(b).toString('base64url'),
    base64Encode: (b) => Buffer.from(b.map((x) => x & 255)).toString('base64'),
    base64Decode: (t) => Array.from(Buffer.from(t, 'base64')).map((x) => (x > 127 ? x - 256 : x))
  });
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', 'Codigo.gs'), 'utf8'), ctx);
  return ctx;
}

module.exports = { criarPlanilha, carregarServidor };
