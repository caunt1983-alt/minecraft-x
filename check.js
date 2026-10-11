// Самопроверка Code.gs на заглушке Google Sheets: node check.js
const fs = require('fs'), vm = require('vm'), assert = require('assert');

function mockSheet() {
  const d = [];
  const cell = (r, c) => ((d[r - 1] ||= [])[c - 1] ?? '');
  const range = (r, c, nr = 1, nc = 1) => typeof r === 'string' ? range(+r.slice(1), r.charCodeAt(0) - 64) : ({
    getValue: () => cell(r, c),
    setValue: v => { (d[r - 1] ||= [])[c - 1] = v; },
    setValues: vs => vs.forEach((row, i) => row.forEach((v, j) => {
      (d[r - 1 + i] ||= [])[c - 1 + j] = typeof v === 'string' && v[0] === "'" ? v.slice(1) : v;
    })),
    clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (d[r - 1 + i]) d[r - 1 + i][c - 1 + j] = ''; },
  });
  return { d, appendRow: row => d.push(row.slice()), setFrozenRows() {}, hideSheet() {},
    getLastRow: () => d.length, getRange: range,
    getDataRange: () => ({ getValues: () => d.map(r => r.slice()) }) };
}
const sheets = {};
const ctx = vm.createContext({
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n], insertSheet: n => (sheets[n] = mockSheet()) }) },
  ContentService: { createTextOutput: t => ({ setMimeType: () => JSON.parse(t) }), MimeType: { JSON: 1 } },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: { formatDate: d => d.toISOString().slice(0, 16) },
  Session: { getScriptTimeZone: () => 'UTC' },
  Date, JSON, Math, String, Number, Object, Set,
});
vm.runInContext(fs.readFileSync(__dirname + '/Code.gs', 'utf8') + ';this.doPost=doPost;this.doGet=doGet;', ctx);
const post = ev => ctx.doPost({ postData: { contents: JSON.stringify(ev) } });

// честный ход: вылазка, старт теста, тест
const state = { rev: 5, nick: 'Тим', facts: { '7×8': { b: 3 } }, won: '', raidWon: '' };
post({ id: '1', type: 'create', nick: 'Тим', state });
post({ id: '2', type: 'session', nick: 'Тим', ms: 60000, state, answers: [
  { q: '7×8', g: '56', ms: 2600 }, { q: '7×8', g: '56', ms: 2900 }, { q: '56:8', g: '7', ms: 2500 }, { q: '56:8', g: '6', ms: 2700 }] });
assert.equal(post({ id: '2', type: 'session', state }).dup, true, 'дубль отбрасывается');
assert.equal(sheets['Ответы'].d[3][2], '56:8', 'деление пишется текстом');
assert.equal(sheets['Ответы'].d[4][4], 'нет', 'неверный ответ распознан сервером');

post({ id: '3', type: 'test_start', nick: 'Тим', state });
// подделка: 80 мгновенных одинаковых ответов, включая «7×8», который на тренировке шёл 2.6 с
const fake = Array.from({ length: 80 }, (_, i) => ({ q: '7×8', g: '56', ms: 200 }));
post({ id: '4', type: 'test', nick: 'Тим', ms: 16000, passed: true, state: { ...state, won: '2026-10-11', wonAt: 'x' }, answers: fake });
const flags = sheets['Лог'].d.at(-1)[5];
for (const f of ['слишком быстро', 'мало вылазок', 'нереально быстро', 'одинаковое время', 'скачок']) assert.ok(flags.includes(f), 'флаг: ' + f + ' в ' + flags);

// таймаут теста > 3 с — не засчитан
post({ id: '5', type: 'raid', nick: 'Тим', state: { ...state, won: '2026-10-11', wonAt: 'x' }, answers: [{ q: '7×8', g: '56', ms: 3100 }] });
assert.equal(sheets['Лог'].d.at(-1)[3], '0/1', 'медленный ответ на тесте не засчитан');

// сводка и восстановление
assert.equal(sheets['Сводка'].d[0][1], 'Тим', 'ник в сводке');
assert.ok(sheets['Сводка'].d[6][1].includes('400'), 'выплата отражена');
assert.deepEqual(sheets['Сводка'].d[10], ['56 : 8', 7], 'самый трудный — пример с ошибкой');
assert.equal(JSON.parse(ctx.doGet().state).rev, 5, 'doGet отдаёт состояние');
console.log('OK');
