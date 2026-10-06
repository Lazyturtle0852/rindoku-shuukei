/**
 * 輪読事前課題 提出状況の集計
 *
 * メニュー「輪読集計」から手動で実行する。
 *   - 提出状況を更新      : フォームの回答から ◯ / 遅延 / 未 / - を書き込む
 *   - 名簿と発表者を反映  : 輪読スケジュールから名簿の追加と「発表」の書き込みをする
 *
 * 学期が変わっても、シート名・見出し名が同じならコードを触る必要はない。
 * 名簿は 集計!A2 から下へ空白セルまで、輪読回は 集計!B1 から右へ日付が続く限りを自動で検出する。
 */

const CONFIG = {
  SUMMARY_SHEET: '集計',
  RESPONSE_SHEET: 'フォームの回答 1',
  SCHEDULE_SHEET: '輪読スケジュール',

  // フォームの回答シートの見出し
  HEADER_TIMESTAMP: 'タイムスタンプ',
  HEADER_NAME: '名前',
  HEADER_ROUND: '輪読回', // 値の先頭が「10/13 …」の形

  // 期限: 研究会の日付の N 日前の HH:00 (JST)。ちょうどは期限内
  DEADLINE_DAYS_BEFORE: 1,
  DEADLINE_HOUR_JST: 18,

  // 「今」の日時。null なら実行した時刻。デモで日時を固定するときだけ '2026-11-09T12:00:00+09:00' のように書く
  NOW: null,

  MARK: {
    ON_TIME: '◯',
    LATE: '遅延',
    MISSING: '未',     // 期限を過ぎても提出がない
    NOT_YET: '-',     // まだ期限前で、提出もない
    PRESENTER: '発表',
  },
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('輪読集計')
    .addItem('提出状況を更新', 'updateSubmissionStatus')
    .addItem('名簿と発表者を反映', 'syncRosterAndPresenters')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* 提出状況を更新                                                       */
/* ------------------------------------------------------------------ */

function updateSubmissionStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const warnings = [];
  const M = CONFIG.MARK;

  const summary = getSheet_(ss, CONFIG.SUMMARY_SHEET);
  const layout = detectLayout_(summary, tz, warnings);
  if (layout.roster.length === 0 || layout.rounds.length === 0) {
    showResult_('提出状況の更新を中止しました', ['名簿（A2以下）または輪読回の日付（B1以右）が見つかりません。'].concat(warnings));
    return;
  }

  const rosterIndex = buildRosterIndex_(layout.roster, warnings);
  const roundIndex = {};
  layout.rounds.forEach((r, i) => { roundIndex[r.key] = i; });

  // 人 × 回 ごとの最初の提出時刻
  const firstSubmit = layout.roster.map(() => layout.rounds.map(() => null));

  const responses = getSheet_(ss, CONFIG.RESPONSE_SHEET);
  const data = responses.getDataRange().getValues();
  const head = data[0].map(h => String(h).trim());
  const col = {
    ts: head.indexOf(CONFIG.HEADER_TIMESTAMP),
    name: head.indexOf(CONFIG.HEADER_NAME),
    round: head.indexOf(CONFIG.HEADER_ROUND),
  };
  const missing = Object.keys(col)
    .filter(k => col[k] < 0)
    .map(k => ({ ts: CONFIG.HEADER_TIMESTAMP, name: CONFIG.HEADER_NAME, round: CONFIG.HEADER_ROUND })[k]);
  if (missing.length) {
    showResult_('提出状況の更新を中止しました', [`「${CONFIG.RESPONSE_SHEET}」に見出し「${missing.join('」「')}」がありません。`]);
    return;
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const rowNo = r + 1;
    if (row.every(v => v === '' || v === null)) continue;

    const ts = row[col.ts];
    const rawName = String(row[col.name]).trim();
    const rawRound = row[col.round];
    const where = `${CONFIG.RESPONSE_SHEET} ${rowNo}行目`;

    if (!(ts instanceof Date) || isNaN(ts)) {
      warnings.push(`${where}: タイムスタンプが読めません（${ts}）`);
      continue;
    }
    const roundKey = parseRoundKey_(rawRound, tz);
    if (!roundKey) {
      warnings.push(`${where}: 輪読回「${rawRound}」から日付を読めません（${rawName}）`);
      continue;
    }
    if (!(roundKey in roundIndex)) {
      warnings.push(`${where}: 輪読回 ${roundKey} が集計の見出しにありません（${rawName}）`);
      continue;
    }
    const p = rosterIndex[normalizeName_(rawName)];
    if (p === undefined) {
      warnings.push(`${where}: 名前「${rawName}」が名簿にありません（輪読回 ${roundKey}）`);
      continue;
    }
    const c = roundIndex[roundKey];
    if (firstSubmit[p][c] === null || ts < firstSubmit[p][c]) firstSubmit[p][c] = ts;
  }

  // 書き込み
  const range = summary.getRange(2, 2, layout.roster.length, layout.rounds.length);
  const values = range.getValues();
  const overwritable = new Set([M.ON_TIME, M.LATE, M.MISSING, M.NOT_YET, '']);
  const now = CONFIG.NOW ? new Date(CONFIG.NOW) : new Date();
  let changed = 0;

  values.forEach((rowVals, p) => {
    rowVals.forEach((cur, c) => {
      const current = String(cur).trim();
      if (current === M.PRESENTER) return;
      if (!overwritable.has(current)) {
        warnings.push(`${CONFIG.SUMMARY_SHEET}!${a1_(p + 2, c + 2)}（${layout.roster[p]}・${layout.rounds[c].key}）: 想定外の値「${current}」をそのまま残しました`);
        return;
      }
      const t = firstSubmit[p][c];
      const deadline = layout.rounds[c].deadline.getTime();
      const next = t !== null ? (t.getTime() <= deadline ? M.ON_TIME : M.LATE)
        : now.getTime() > deadline ? M.MISSING
        : M.NOT_YET;
      if (next !== current) changed++;
      rowVals[c] = next;
    });
  });
  changed -= writeValues_(range, values, warnings);

  const nowNote = CONFIG.NOW ? `・基準日時 ${Utilities.formatDate(now, 'Asia/Tokyo', 'M/d HH:mm')} で固定中` : '';
  showResult_(`提出状況を更新しました（変更 ${changed} セル${nowNote}）`, warnings);
}

/* ------------------------------------------------------------------ */
/* 名簿と発表者を反映                                                   */
/* ------------------------------------------------------------------ */

function syncRosterAndPresenters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const warnings = [];
  const M = CONFIG.MARK;

  const summary = getSheet_(ss, CONFIG.SUMMARY_SHEET);
  let layout = detectLayout_(summary, tz, warnings);
  if (layout.rounds.length === 0) {
    showResult_('反映を中止しました', ['輪読回の日付（B1以右）が見つかりません。'].concat(warnings));
    return;
  }

  // スケジュール: A列が日付、B列以降が発表者
  const schedule = getSheet_(ss, CONFIG.SCHEDULE_SHEET).getDataRange().getValues();
  const presenters = []; // { key, name }
  schedule.forEach((row, i) => {
    const key = parseRoundKey_(row[0], tz);
    if (!key) {
      if (row.some(v => v !== '')) warnings.push(`${CONFIG.SCHEDULE_SHEET} ${i + 1}行目: A列の日付が読めません（${row[0]}）`);
      return;
    }
    row.slice(1).map(v => String(v).trim()).filter(Boolean)
      .forEach(name => presenters.push({ key, name }));
  });

  // 名簿にいない人を末尾に追加（集計行の式の範囲が広がるよう、名簿の直後に行を挿入）
  const known = new Set(layout.roster.map(normalizeName_));
  const toAdd = [];
  presenters.forEach(({ name }) => {
    const n = normalizeName_(name);
    if (!known.has(n)) { known.add(n); toAdd.push(name); }
  });
  if (toAdd.length) {
    const last = 1 + layout.roster.length;
    summary.insertRowsAfter(last, toAdd.length);
    const lastCol = summary.getLastColumn();
    if (layout.roster.length > 0) {
      summary.getRange(last, 1, 1, lastCol).copyTo(summary.getRange(last + 1, 1, toAdd.length, lastCol));
    }
    summary.getRange(last + 1, 1, toAdd.length, 1).setValues(toAdd.map(n => [n]));
    summary.getRange(last + 1, 2, toAdd.length, layout.rounds.length)
      .setValues(toAdd.map(() => layout.rounds.map(() => '')));
    layout = detectLayout_(summary, tz, []);
  }

  // 「発表」を書き込む
  const rosterIndex = buildRosterIndex_(layout.roster, warnings);
  const roundIndex = {};
  layout.rounds.forEach((r, i) => { roundIndex[r.key] = i; });
  const range = summary.getRange(2, 2, layout.roster.length, layout.rounds.length);
  const values = range.getValues();
  const overwritable = new Set([M.ON_TIME, M.LATE, M.MISSING, M.NOT_YET, M.PRESENTER, '']);
  let marked = 0;

  presenters.forEach(({ key, name }) => {
    const c = roundIndex[key];
    if (c === undefined) {
      warnings.push(`${CONFIG.SCHEDULE_SHEET}: ${key} が集計の見出しにありません（${name}）`);
      return;
    }
    const p = rosterIndex[normalizeName_(name)];
    if (p === undefined) {
      warnings.push(`${CONFIG.SUMMARY_SHEET}: 「${name}」を名簿で見つけられません。A列の途中に空白行がないか確認してください`);
      return;
    }
    const current = String(values[p][c]).trim();
    if (!overwritable.has(current)) {
      warnings.push(`${CONFIG.SUMMARY_SHEET}!${a1_(p + 2, c + 2)}（${name}・${key}）: 想定外の値「${current}」があるため「発表」にしませんでした`);
      return;
    }
    if (current !== M.PRESENTER) marked++;
    values[p][c] = M.PRESENTER;
  });
  marked -= writeValues_(range, values, warnings);

  const head = toAdd.length ? `名簿に ${toAdd.length} 人追加し、` : '';
  showResult_(`${head}「発表」を ${marked} セル新たに書き込みました`, warnings);
}

/* ------------------------------------------------------------------ */
/* 共通                                                                 */
/* ------------------------------------------------------------------ */

function getSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シート「${name}」が見つかりません`);
  return sh;
}

/** 名簿（A2から下へ空白まで）と輪読回（B1から右へ日付が続く限り）を検出する */
function detectLayout_(sheet, tz, warnings) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const roster = [];
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (const [v] of names) {
      const s = String(v).trim();
      if (!s) break;
      roster.push(s);
    }
  }

  const rounds = [];
  if (lastCol >= 2) {
    const heads = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    for (const v of heads) {
      const h = toDate_(v);
      if (!h) break;
      const [y, m, d] = Utilities.formatDate(h, tz, 'yyyy-M-d').split('-').map(Number);
      rounds.push({
        key: `${m}/${d}`,
        // Date.UTC は日付の繰り下がり（1日 → 前月末）を扱える。JST = UTC+9
        deadline: new Date(Date.UTC(y, m - 1, d - CONFIG.DEADLINE_DAYS_BEFORE, CONFIG.DEADLINE_HOUR_JST - 9)),
        date: h,
      });
    }
  }
  for (let i = 1; i < rounds.length; i++) {
    if (rounds[i].date <= rounds[i - 1].date) {
      warnings.push(`${CONFIG.SUMMARY_SHEET}: 見出し ${rounds[i].key} の日付が前の回より前です。年が正しいか確認してください（期限の判定に影響します）`);
    }
  }
  return { roster, rounds };
}

function buildRosterIndex_(roster, warnings) {
  const index = {};
  roster.forEach((name, i) => {
    const n = normalizeName_(name);
    if (n in index) warnings.push(`${CONFIG.SUMMARY_SHEET}: 名簿の「${name}」が重複しています（${index[n] + 2}行目と${i + 2}行目）`);
    else index[n] = i;
  });
  return index;
}

/**
 * 値だけを書き込む。書式・入力規則・プルダウンの色には触れない。
 * 入力規則が「拒否」設定で、書こうとした記号がリストにないマスは元の値のまま残し、記号ごとに警告へまとめる。
 * 戻り値は書き込めなかったマスの数。
 */
function writeValues_(range, values, warnings) {
  const before = range.getValues();
  const rules = range.getDataValidations();
  const blocked = {};
  values.forEach((row, i) => row.forEach((v, j) => {
    if (v === before[i][j] || v === '') return;
    const rule = rules[i][j];
    if (!rule || rule.getAllowInvalid()) return;
    if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;
    if (rule.getCriteriaValues()[0].map(String).includes(String(v))) return;
    if (!blocked[v]) blocked[v] = [];
    blocked[v].push(a1_(range.getRow() + i, range.getColumn() + j));
    row[j] = before[i][j];
  }));

  try {
    range.setValues(values);
  } catch (e) {
    warnings.push(`書き込みに失敗しました: ${e.message}`);
  }

  let n = 0;
  Object.keys(blocked).forEach(mark => {
    const cells = blocked[mark];
    n += cells.length;
    const where = cells.slice(0, 5).join(', ') + (cells.length > 5 ? ' ほか' : '');
    warnings.push(`入力規則のリストに「${mark}」がないため、${cells.length} マス（${where}）を書き換えられませんでした。「データ › データの入力規則」でリストに「${mark}」を足してください`);
  });
  return n;
}

/** 全角/半角をそろえ（NFKC）、空白をすべて除く */
function normalizeName_(s) {
  return String(s).normalize('NFKC').replace(/\s+/g, '');
}

/** 「10/13 チームインテリジェンス…」や日付セルから "10/13" を取り出す */
function parseRoundKey_(v, tz) {
  const d = toDate_(v);
  if (d) return Utilities.formatDate(d, tz, 'M/d');
  const m = String(v).normalize('NFKC').match(/^\s*(\d{1,2})\s*\/\s*(\d{1,2})/);
  return m ? `${Number(m[1])}/${Number(m[2])}` : null;
}

/**
 * 日付セルを Date にする。表示形式が外れて数値（シリアル値 46308 など）になったセルも日付として読む。
 * シリアル値はシートのタイムゾーンでの 0:00 を指すので、JST の 0:00 として扱う。
 */
function toDate_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number' && v > 40000 && v < 80000) {
    return new Date(Date.UTC(1899, 11, 30) + v * 86400000 - 9 * 3600000);
  }
  return null;
}

function a1_(row, col) {
  let s = '';
  for (let c = col; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s + row;
}

function showResult_(title, warnings) {
  const ui = SpreadsheetApp.getUi();
  if (warnings.length === 0) {
    ui.alert(title, '警告はありません。', ui.ButtonSet.OK);
    return;
  }
  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;font-size:13px">
       <p>${esc(title)}</p>
       <p style="color:#b45309">警告 ${warnings.length} 件</p>
       <ul style="padding-left:1.2em">${warnings.map(w => `<li style="margin-bottom:4px">${esc(w)}</li>`).join('')}</ul>
     </div>`
  ).setWidth(640).setHeight(Math.min(160 + warnings.length * 28, 560));
  ui.showModalDialog(html, '輪読集計');
}
