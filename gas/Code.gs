/**
 * タスク・予定時間軸一元管理システム (Spreadsheet + Calendar + Gemini手動コピペ)
 */

const SHEETS = {
  TASKS: 'タスク管理',
  SHIFT: 'シフト一括登録',
  SETTINGS: 'システム設定',
  APPLY: 'スケジュール反映',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️スケジュール管理')
    .addItem('1. 朝のGemini用テキスト生成', 'showMorningPromptDialog')
    .addItem('2. カレンダーへスケジュール反映', 'applyGeminiScheduleFromSheet')
    .addItem('3. シフト一括登録', 'importShiftRowsToCalendar')
    .addToUi();
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('シート未作成: ' + SHEETS.SETTINGS);
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const map = {};
  values.forEach(([k, v]) => { if (k) map[String(k).trim()] = v; });
  return {
    calendarId: String(map['カレンダーID'] || 'primary').trim(),
    workStart: String(map['稼働開始時間'] || '09:00').trim(),
    workEnd: String(map['稼働終了時間'] || '18:00').trim(),
    bufferMin: Number(map['前後バッファ（分）'] || 15),
    scanDays: Number(map['スキャン日数'] || 3),
  };
}

function showMorningPromptDialog() {
  const text = buildMorningGeminiPrompt();
  const html = HtmlService.createHtmlOutput(
    '<textarea style="width:100%;height:420px">' + text.replace(/</g, '&lt;') + '</textarea>'
  ).setWidth(800).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Geminiへ貼り付けるテキスト');
}

function buildMorningGeminiPrompt() {
  const cfg = getSettings_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';

  const tasks = getIncompleteTasks_();
  const events = getCalendarEventsForScan_(cfg, tz);

  const taskLines = tasks.length ? tasks.map((t) => t.geminiLine).join('\n') : '・未完了タスクなし';
  const eventLines = events.length
    ? events.map((e) => `・${e.date} ${e.start}-${e.end} ${e.title}`).join('\n')
    : '・固定予定なし';

  return [
    '以下の条件でスケジュールを再構築してください。',
    '',
    '【ルール】',
    `- 稼働時間: ${cfg.workStart}〜${cfg.workEnd}`,
    `- 前後バッファ: ${cfg.bufferMin}分`,
    '- 固定予定（既存イベント）は移動しない',
    '- 夜間（稼働終了後）は割り当てない',
    '- 期限があるものは期限内で逆算',
    '- いつでもタスクは直近の空きへ',
    '- 長時間は分割指示に従う',
    '',
    '【既存予定】',
    eventLines,
    '',
    '【未完了タスク】',
    taskLines,
    '',
    '【出力形式（厳守）】',
    '1行1予定、カンマ区切りで次のみ:',
    'YYYY-MM-DD,HH:mm-HH:mm,タイトル',
    '',
    '例:',
    '2026-05-21,09:30-10:10,【PC】資料作成(1/3)',
  ].join('\n');
}

function getIncompleteTasks_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TASKS);
  if (!sh) throw new Error('シート未作成: ' + SHEETS.TASKS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const vals = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  return vals
    .filter((r) => r[1] && r[0] !== true)
    .map((r, i) => ({
      row: i + 2,
      title: String(r[1]),
      attr: String(r[2] || ''),
      deadline: r[3],
      minutes: Number(r[4] || 0),
      split: Number(r[5] || 1),
      geminiLine: String(r[7] || '').trim() || fallbackGeminiLine_(r),
    }));
}

function fallbackGeminiLine_(r) {
  const title = String(r[1] || '');
  const attr = String(r[2] || '');
  const due = r[3] === 'いつでも' ? '期限：なし（空き時間おまかせ）' : '期限あり';
  const min = Number(r[4] || 0);
  return `・${due}: ${title}【属性:${attr}】➔ 「${min}分」を配置せよ`;
}

function getCalendarEventsForScan_(cfg, tz) {
  const cal = CalendarApp.getCalendarById(cfg.calendarId);
  if (!cal) throw new Error('カレンダーが見つかりません: ' + cfg.calendarId);
  const start = new Date();
  const end = new Date(start.getTime() + cfg.scanDays * 24 * 60 * 60 * 1000);
  const events = cal.getEvents(start, end);
  return events.map((e) => ({
    date: Utilities.formatDate(e.getStartTime(), tz, 'yyyy-MM-dd'),
    start: Utilities.formatDate(e.getStartTime(), tz, 'HH:mm'),
    end: Utilities.formatDate(e.getEndTime(), tz, 'HH:mm'),
    title: e.getTitle(),
  }));
}

function applyGeminiScheduleFromSheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.APPLY);
  if (!sh) throw new Error('シート未作成: ' + SHEETS.APPLY);
  const raw = String(sh.getRange('A1').getValue() || '').trim();
  if (!raw) throw new Error('スケジュール反映!A1 が空です');

  const cfg = getSettings_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
  const cal = CalendarApp.getCalendarById(cfg.calendarId);
  if (!cal) throw new Error('カレンダーが見つかりません: ' + cfg.calendarId);

  const lines = raw.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  const results = [];

  lines.forEach((line) => {
    const parsed = parseScheduleLine_(line);
    if (!parsed) {
      results.push([line, 'SKIP', 'parse_error']);
      return;
    }
    if (parsed.title.includes('【確定】')) {
      results.push([line, 'SKIP', 'fixed_event']);
      return;
    }

    const start = toDate_(parsed.date, parsed.start, tz);
    const end = toDate_(parsed.date, parsed.end, tz);
    if (!(start < end)) {
      results.push([line, 'SKIP', 'invalid_time_range']);
      return;
    }

    if (hasDuplicateEvent_(cal, parsed.title, start, end)) {
      results.push([line, 'SKIP', 'duplicate']);
      return;
    }

    cal.createEvent(parsed.title, start, end, { description: '登録元：Geminiスケジュール反映' });
    results.push([line, 'CREATE', 'ok']);
  });

  sh.getRange(1, 2, Math.max(results.length, 1), 3).clearContent();
  if (results.length) sh.getRange(1, 2, results.length, 3).setValues(results);
  SpreadsheetApp.getUi().alert('反映完了: ' + results.length + '行処理しました');
}

function parseScheduleLine_(line) {
  const m = line.match(/^(\d{4}-\d{2}-\d{2})\s*,\s*(\d{2}:\d{2})-(\d{2}:\d{2})\s*,\s*(.+)$/);
  if (!m) return null;
  return { date: m[1], start: m[2], end: m[3], title: m[4].trim() };
}

function toDate_(dateStr, timeStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const [hh, mm] = timeStr.split(':').map(Number);
  return new Date(y, m - 1, d, hh, mm, 0, 0);
}

function hasDuplicateEvent_(cal, title, start, end) {
  const events = cal.getEvents(start, end);
  return events.some((e) =>
    e.getTitle() === title &&
    e.getStartTime().getTime() === start.getTime() &&
    e.getEndTime().getTime() === end.getTime()
  );
}

function importShiftRowsToCalendar() {
  const cfg = getSettings_();
  const cal = CalendarApp.getCalendarById(cfg.calendarId);
  if (!cal) throw new Error('カレンダーが見つかりません: ' + cfg.calendarId);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SHIFT);
  if (!sh) throw new Error('シート未作成: ' + SHEETS.SHIFT);
  const last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();

  rows.forEach((r) => {
    const date = r[0], startT = r[1], endT = r[2], title = r[3];
    if (!date || !startT || !endT || !title) return;
    const start = mergeDateTime_(date, startT);
    const end = mergeDateTime_(date, endT);
    const fixedTitle = String(title).includes('【確定】') ? String(title) : `【確定】${title}`;
    if (!hasDuplicateEvent_(cal, fixedTitle, start, end)) {
      cal.createEvent(fixedTitle, start, end, { description: '登録元：シフト一括登録' });
    }
  });
}

function mergeDateTime_(d, t) {
  const date = new Date(d);
  const time = new Date(t);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), time.getHours(), time.getMinutes(), 0, 0);
}
