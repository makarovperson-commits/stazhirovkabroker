// Дневник стажёра — хранилище в Google-таблице.
// 1) Впишите свой пароль РОП ниже. 2) Развернуть → Новое развёртывание → Веб-приложение,
//    «Выполнять от имени»: Я, «Доступ»: Все. 3) Ссылку веб-приложения вставьте в config.js.
const ROP_PASSWORD = 'ЗАМЕНИТЕ-НА-СВОЙ-ПАРОЛЬ';

const H_B = ['id', 'ФИО', 'Старт', 'Обновлено', 'token', 'profile_json'];
const H_D = ['key', 'id', 'ФИО', 'Дата', 'Звонки', 'Новые лиды', 'Показов назначено', 'Показов проведено', 'Брони',
  'Сумма броней', 'Сделки', 'Комиссия', 'Не работал', 'Ролёвка', 'Материал изучен', 'Встреча с РОП', 'Отметка',
  'Главный вывод дня', 'Обновлено', 'json'];

function doGet() { return out({ ok: true, app: 'rich-dnevnik' }); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) { return out({ ok: false, error: 'bad json' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.action === 'save') return out(save(p));
    if (p.action === 'all') return out(all(p));
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet(name, header) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(header); s.setFrozenRows(1); }
  return s;
}

function save(p) {
  if (!p.id || !p.token) return { ok: false, error: 'no id' };
  const prof = p.profile || {};
  const name = String(prof.label || '');
  const now = new Date();
  const b = sheet('Брокеры', H_B);
  const bv = b.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < bv.length; i++) {
    if (bv[i][0] === p.id) {
      if (bv[i][4] !== p.token) return { ok: false, error: 'token' };
      row = i + 1;
      break;
    }
  }
  const rec = [p.id, name, String(prof.start || ''), now, p.token, JSON.stringify(prof)];
  if (row < 0) b.appendRow(rec); else b.getRange(row, 1, 1, rec.length).setValues([rec]);

  const d = sheet('Отчёты', H_D);
  const dv = d.getDataRange().getValues();
  const idx = {};
  for (let i = 1; i < dv.length; i++) idx[dv[i][0]] = i + 1;
  const del = [];
  Object.keys(p.days || {}).forEach(function (date) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
    const e = p.days[date];
    const key = p.id + '|' + date;
    if (!e) { if (idx[key]) del.push(idx[key]); return; }
    const r = [key, p.id, name, date, e.calls || 0, e.leads || 0, e.showsSet || 0, e.showsDone || 0, e.bookings || 0,
      e.bookingSum || 0, e.deals || 0, e.commission || 0, e.off ? 1 : 0, e.rp || '', e.mat ? 1 : 0, e.meet ? 1 : 0,
      e.done ? 1 : 0, e.note || '', now, JSON.stringify(e)];
    if (idx[key]) d.getRange(idx[key], 1, 1, r.length).setValues([r]);
    else { d.appendRow(r); idx[key] = d.getLastRow(); }
  });
  del.sort(function (x, y) { return y - x; }).forEach(function (r) { d.deleteRow(r); });
  return { ok: true, at: now.toISOString() };
}

function all(p) {
  if (!p.key || p.key !== ROP_PASSWORD || ROP_PASSWORD === 'ЗАМЕНИТЕ-НА-СВОЙ-ПАРОЛЬ') return { ok: false, error: 'password' };
  const res = {};
  const bv = sheet('Брокеры', H_B).getDataRange().getValues();
  for (let i = 1; i < bv.length; i++) {
    let prof = {};
    try { prof = JSON.parse(bv[i][5] || '{}'); } catch (_) {}
    const upd = bv[i][3];
    res[bv[i][0]] = { profile: prof, days: {}, updated: upd instanceof Date ? upd.toISOString() : String(upd) };
  }
  const dv = sheet('Отчёты', H_D).getDataRange().getValues();
  for (let i = 1; i < dv.length; i++) {
    const parts = String(dv[i][0]).split('|');
    if (!res[parts[0]]) continue;
    let e = {};
    try { e = JSON.parse(dv[i][19] || '{}'); } catch (_) {}
    res[parts[0]].days[parts[1]] = e;
  }
  return { ok: true, brokers: res };
}
