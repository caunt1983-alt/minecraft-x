// Minecraft-X: приёмник результатов игры. Вставляется в Расширения → Apps Script таблицы,
// разворачивается как веб-приложение (Выполнять как: я, Доступ: все).

const SS = SpreadsheetApp.getActive();
const HEADERS = {
  'Сводка': ['Ник (можно поменять)', ''],
  'Лог': ['Время', 'Тип', 'Ник', 'Результат', 'Сек', '⚠ Флаги', 'id'],
  'Ответы': ['Время', 'Тип', 'Пример', 'Ответ', 'Верно', 'мс'],
};
const LIMIT = { test: 3000, raid: 3000 }; // остальное — 6000

function sheet(name) {
  let s = SS.getSheetByName(name);
  if (!s) {
    s = SS.insertSheet(name);
    if (HEADERS[name]) { s.appendRow(HEADERS[name]); s.setFrozenRows(1); }
    if (name === '_state') s.hideSheet();
  }
  return s;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const day = d => Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
const fmt = d => Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
function answerOf(q) {
  const m = String(q).match(/^(\d+)([×:])(\d+)$/);
  if (!m) return NaN;
  return m[2] === '×' ? m[1] * m[3] : m[1] / m[3];
}

// Игра забирает прогресс (восстановление после чистки браузера) и ник.
function doGet() {
  return out({ state: sheet('_state').getRange('A1').getValue() || null, nick: String(sheet('Сводка').getRange('B1').getValue() || '') });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ev = JSON.parse(e.postData.contents);
    const log = sheet('Лог');
    const rows = log.getDataRange().getValues().slice(1);
    if (rows.some(r => r[6] === ev.id)) return out({ ok: true, dup: true });

    const now = new Date();
    const limit = LIMIT[ev.type] || 6000;
    // Правильность пересчитываем сами, клиенту не верим.
    const ans = (ev.answers || []).map(a => ({ q: a.q, g: a.g, ms: Number(a.ms) || 0, ok: Number(a.g) === answerOf(a.q) && a.g !== '' && a.ms <= limit }));
    const good = ans.filter(a => a.ok).length;
    const flags = (ev.type === 'test' || ev.type === 'raid') ? suspicious(ev.type, ans, rows, now) : [];

    log.appendRow([now, ev.type, ev.nick, ans.length ? good + '/' + ans.length : '', Math.round((ev.ms || 0) / 1000), flags.join('; '), ev.id]);
    if (ans.length) {
      const s = sheet('Ответы');
      s.getRange(s.getLastRow() + 1, 1, ans.length, 6).setValues(ans.map(a => [now, ev.type, "'" + a.q, a.g, a.ok ? 'да' : 'нет', a.ms]));
    }
    if (ev.state) sheet('_state').getRange('A1').setValue(JSON.stringify(ev.state));
    summary(ev.state || {}, now);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Признаки подделки результата теста/рейда.
function suspicious(type, ans, rows, now) {
  const f = [];
  const start = rows.filter(r => r[1] === type + '_start').pop();
  if (!start) f.push('нет старта');
  else {
    const sec = (now - new Date(start[0])) / 1000;
    if (sec < ans.length * 0.6) f.push('слишком быстро: ' + Math.round(sec) + ' с');
    if (sec > 1800) f.push('старт ' + Math.round(sec / 60) + ' мин назад');
  }
  const days = new Set(rows.filter(r => r[1] === 'session' || r[1] === 'diag').map(r => day(r[0]))).size;
  if (type === 'test' && days < 5) f.push('мало вылазок: ' + days + ' дн.');
  const ms = ans.map(a => a.ms).sort((a, b) => a - b);
  if (ms.length) {
    const med = ms[ms.length >> 1], avg = ms.reduce((s, x) => s + x, 0) / ms.length;
    const sd = Math.sqrt(ms.reduce((s, x) => s + (x - avg) ** 2, 0) / ms.length);
    if (med < 500) f.push('нереально быстро: медиана ' + med + ' мс');
    if (sd < 120) f.push('одинаковое время ответов');
  }
  // Скачок: на тренировках пример давался долго, а на тесте вдруг молниеносно.
  const practice = {};
  sheet('Ответы').getDataRange().getValues().slice(1)
    .filter(r => r[1] === 'session' && r[4] === 'да')
    .forEach(r => { (practice[r[2]] = practice[r[2]] || []).push(Number(r[5])); });
  const jumps = ans.filter(a => {
    const p = practice[a.q]; if (!p || p.length < 2) return false;
    const avg = p.reduce((s, x) => s + x, 0) / p.length;
    return avg > 2000 && a.ms < avg * 0.35;
  }).length;
  if (jumps >= 5) f.push('скачок скорости: ' + jumps + ' примеров');
  return f;
}

function summary(st, now) {
  const s = sheet('Сводка');
  const facts = st.facts || {};
  const chest = Object.keys(facts).filter(k => facts[k].b === 3).length;
  const log = sheet('Лог').getDataRange().getValues().slice(1);
  const days = new Set(log.filter(r => r[1] === 'session' || r[1] === 'diag').map(r => day(r[0]))).size;
  if (!s.getRange('B1').getValue() && st.nick) s.getRange('B1').setValue(st.nick);

  const pay = [];
  if (st.won) pay.push('400 ₽ — дракон ' + st.wonAt, '+300 ₽ — если звонок 10/10');
  if (st.raidWon) pay.push('+100 ₽ — рейд ' + st.raidWon);
  const flagged = log.filter(r => r[5]).map(r => fmt(r[0]) + ' ' + r[1] + ': ' + r[5]);

  const rows = [
    ['В сундуке', chest + ' / 128'],
    ['Дней занятий', days],
    ['Последняя активность', fmt(now)],
    ['Дракон', st.won ? 'повержен ' + st.wonAt : (st.lastTest ? 'последняя попытка ' + st.lastTest : 'ещё не открыт')],
    ['Рейд', st.raidWon ? 'пройден ' + st.raidWon : (st.won ? 'через 7 дней после дракона' : '—')],
    ['Выплаты', pay.join('\n') || '—'],
    ['⚠ Подозрения', flagged.join('\n') || 'нет'],
    ['', ''],
    ['🎯 10 примеров для звонка', 'правильный ответ'],
  ];
  s.getRange(2, 1, rows.length, 2).setValues(rows);
  const hard = hardest(10);
  s.getRange(11, 1, 10, 2).clearContent();
  if (hard.length) s.getRange(11, 1, hard.length, 2).setValues(hard);
}

// Самые трудные для него примеры: ошибки весят как 3 секунды раздумий.
function hardest(n) {
  const st = {};
  sheet('Ответы').getDataRange().getValues().slice(1).forEach(r => {
    const x = st[r[2]] = st[r[2]] || { err: 0, sum: 0, cnt: 0 };
    if (r[4] === 'нет') x.err++; else { x.sum += Number(r[5]); x.cnt++; }
  });
  return Object.keys(st)
    .map(q => [q, st[q].err * 3000 + (st[q].cnt ? st[q].sum / st[q].cnt : 3000)])
    .sort((a, b) => b[1] - a[1]).slice(0, n)
    .map(([q]) => [q.replace('×', ' × ').replace(':', ' : '), answerOf(q)]);
}
