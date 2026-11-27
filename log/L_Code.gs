/**
 * Cutech Log — Apps Script サーバー側
 * スプレッドシートにバインドして使います（拡張機能 > Apps Script）。
 * 管理者ID/パスワードは「プロジェクトの設定 > スクリプト プロパティ」の
 * ADMIN_ID / ADMIN_PASS に設定します（コードには書かない）。
 */
const SHEETS = {
  records:   ['id', 'date', 'event', 'name', 'json'],
  names:     ['name'],
  scrambles: ['id', 'date', 'event', 'json'],
  events:    ['json'],                               // 種目の設定（順序・名称・略称・スクランブルの種類）
  badges:    ['json'],                               // バッジの設定（基準・名称・説明文）
  changes:   ['rev', 'table', 'op', 'id', 'json'],  // 差分読み込み用の変更履歴
  meta:      ['key', 'value']                        // 閲覧用サイトが版番号を読むための控え
};
// 種目の「スクランブルの種類」に選べる初期の種目（画面側の BUILTIN_EV と同じ）
const BUILTIN_EVENTS = ['3x3x3 Cube', '2x2x2 Cube', '4x4x4 Cube', '5x5x5 Cube', '6x6x6 Cube', '7x7x7 Cube', '3x3x3 Blindfolded', '3x3x3 Fewest Moves', '3x3x3 One-Handed', 'Clock', 'Megaminx', 'Pyraminx', 'Skewb', 'Square-1', '4x4x4 Blindfolded', '5x5x5 Blindfolded', '3x3x3 Multi-Blind', 'Face Turning Octahedron'];
const PUBLIC_OPS = { addRec: 1, addName: 1, genScr: 1 };   // ログイン不要の操作（genScr＝その日の、まだ無いスクランブルの生成・保存）
const APP_TZ = 'Asia/Tokyo';                      // 「その日」の判定に使うタイムゾーン
const PUBLIC_LIMIT_PER_MIN = 300;                // 公開書き込みの連投制限（全体・毎分）
const PUBLIC_LIMIT_PER_DEV = 30;                 // 公開書き込みの連投制限（端末ごと・毎分・操作数。端末IDが不正・なしの端末は全員で共有）
const MAX_RECORDS = 50000;

// 管理者ログインの保護
const DEVICE_MAX_FAILS = 5;       // 端末ごと：5回失敗で
const DEVICE_LOCK_SEC = 3600;     //   その端末を1時間ロック
const GLOBAL_MAX_FAILS = 100;     // 全体：10分間に100回を超えて失敗している間は、
const GLOBAL_WINDOW_SEC = 600;    //   （10分ごとに集計）失敗の応答を3秒遅らせる（全員のロックはしない）
const SESSION_SEC = 21600;        // ログイン保持の上限（6時間：CacheServiceの上限）

// 差分読み込み
const LOG_KEEP = 5000;            // 変更履歴の保持件数
const LOG_PRUNE_AT = 6000;        // これを超えたら古い履歴を削除
const DIFF_MAX = 2000;            // これ以上の差分は全件読み込みにする

function doGet() {
  return HtmlService.createHtmlOutputFromFile('L_Index')
    .setTitle('Cutech Log')
    .setFaviconUrl('https://2023kado01.github.io/cutech-icon/cutechicon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, interactive-widget=resizes-content');
}

/* ---------- シート入出力 ---------- */
function sh_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(SHEETS[name]);
    s.setFrozenRows(1);
  }
  return s;
}
function readAll_(name) {
  const s = sh_(name), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, SHEETS[name].length).getValues();
}
function writeRows_(name, rows, append) {
  const s = sh_(name), w = SHEETS[name].length;
  const last = s.getLastRow();
  const start = append ? Math.max(last + 1, 2) : 2;
  if (rows.length) {
    const r = s.getRange(start, 1, rows.length, w);
    r.setNumberFormat('@');            // 日付や数字が自動変換されないよう文字列固定
    r.setValues(rows);
  }
  if (append || last < 2) return;
  // 上書きモード：先に消さず、新しい内容を2行目から上書きしてから、余った末尾の行だけ消す
  const extra = last - 1 - rows.length;
  if (extra <= 0) return;
  if (rows.length) s.deleteRows(2 + rows.length, extra);
  else s.getRange(2, 1, extra, w).clearContent();   // 全行を消すことはできない（見出し行しか残らない）ので中身だけ消す
}
// A列（id）だけを読む。行番号は index+2
function readIds_(name) {
  const s = sh_(name), n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]));
}
// 行単位の更新・削除・追記。ids＝シート上のid列（2行目から）。更新/削除が200行を超えるときだけ全行を読んで上書き方式にする
function patchRows_(name, ids, newRows, updRows, delIds) {
  if (!delIds.size && !updRows.size) { if (newRows.length) writeRows_(name, newRows, true); return; }
  const w = SHEETS[name].length, updAt = [], delAt = [];
  ids.forEach((id, i) => {
    if (delIds.has(id)) delAt.push(i + 2);
    else if (updRows.has(id)) updAt.push(i + 2);
  });
  if (delAt.length > 200 || updAt.length > 200) {
    // 件数が多いときは、新しい内容を2行目から一度に上書きして、余った末尾の行だけ消す
    const kept = readAll_(name).filter(r => !delIds.has(String(r[0]))).map(r => updRows.get(String(r[0])) || r);
    writeRows_(name, kept.concat(newRows), false);
    return;
  }
  const s = sh_(name);
  for (const n of updAt) {
    const rg = s.getRange(n, 1, 1, w);
    rg.setNumberFormat('@');
    rg.setValues([updRows.get(ids[n - 2])]);
  }
  if (newRows.length) writeRows_(name, newRows, true);
  if (!delAt.length) return;
  if (ids.length - delAt.length + newRows.length === 0) { s.getRange(2, 1, ids.length, w).clearContent(); return; }   // 全行は削除できないので中身だけ消す
  delAt.sort((a, b) => b - a);   // 下の行から消して、行番号がずれないようにする
  let i = 0;
  while (i < delAt.length) {
    let j = i;
    while (j + 1 < delAt.length && delAt[j + 1] === delAt[j] - 1) j++;
    s.deleteRows(delAt[j], j - i + 1);   // 連続する行はまとめて削除
    i = j + 1;
  }
}
function parse_(v) { try { return JSON.parse(v); } catch (e) { return null; } }
function props_() { return PropertiesService.getScriptProperties(); }
function rev_() { return Number(props_().getProperty('REV') || 0); }
// 版番号を、閲覧用サイト（別プロジェクト）が読めるように、metaシートにも書いておく
function setMeta_(rev, minRev) {
  sh_('meta').getRange(2, 1, 2, 2).setValues([['REV', rev], ['MINREV', minRev]]);
}
function ensureMeta_() {
  if (!SpreadsheetApp.getActive().getSheetByName('meta')) setMeta_(rev_(), minRev_());
}
function minRev_() { return Number(props_().getProperty('MINREV') || 0); }

/* ---------- 読み込み ---------- */
function getAll() {
  try { ensureMeta_(); } catch (e) { /* 閲覧用の控えが作れなくても入力側には影響させない */ }
  const rev = rev_();   // データより先に読む（途中で更新されても差分で補正される）
  return {
    rev: rev,
    records:   readAll_('records').map(r => parse_(r[4])).filter(Boolean).reverse(),
    names:     readAll_('names').map(r => String(r[0])),
    scrambles: readAll_('scrambles').map(r => parse_(r[3])).filter(Boolean).reverse(),
    events:    readEvents_(),
    badges:    readBadges_()
  };
}
// 種目の設定。未設定のときは空の配列（画面側が初期の18種目を使う）
function readEvents_() {
  const r = readAll_('events');
  const v = r.length ? parse_(r[0][0]) : null;
  return Array.isArray(v) ? v : [];
}

// バッジの設定。未設定のときは空のオブジェクト（画面側が初期設定を使う）
function readBadges_() {
  const r = readAll_('badges');
  const v = r.length ? parse_(r[0][0]) : null;
  return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
}

// 前回の版（since）以降の変更だけを返す。履歴が足りない場合は全件を返す。
function getChanges(since) {
  const rev = rev_();
  since = Number(since);
  if (!(since >= 0) || since > rev || since < minRev_()) return fullResult_();
  if (since === rev) return { rev: rev, changes: [] };
  const rows = changesSince_(since);
  if (rows === null) return fullResult_();
  const changes = [];
  for (const r of rows) {
    const v = parse_(r[4]);
    if (r[2] === 'del') changes.push({ t: String(r[1]), op: 'del', id: String(r[3]) });
    else if (v !== null) changes.push({ t: String(r[1]), op: r[2] === 'set' ? 'set' : 'put', id: String(r[3]), v: v });
  }
  return { rev: rev, changes: changes };
}
// changes は rev の昇順。A列だけ読んで二分探索し、since より新しい行だけ5列読む（多すぎるときは null）
function changesSince_(since) {
  const s = sh_('changes'), n = s.getLastRow() - 1;
  if (n < 1) return [];
  const revs = s.getRange(2, 1, n, 1).getValues();
  let lo = 0, hi = n;
  while (lo < hi) { const m = (lo + hi) >> 1; if (Number(revs[m][0]) > since) hi = m; else lo = m + 1; }
  if (n - lo > DIFF_MAX) return null;
  if (lo === n) return [];
  return s.getRange(2 + lo, 1, n - lo, 5).getValues().filter(r => Number(r[0]) > since);
}
function fullResult_() {
  const d = getAll();
  return { full: true, rev: d.rev, data: d };
}

/* ---------- ログイン ---------- */
function login(id, pw, dev) {
  const c = CacheService.getScriptCache();
  const d = devKey_(dev);
  if (c.get('ld_' + d)) return { ok: false, msg: 'この端末は、ログインに5回失敗したため、1時間ログインできません。' };
  if (credOk_(id, pw)) {
    c.remove('fd_' + d);
    const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    c.put('tk_' + token, '1', SESSION_SEC);
    return { ok: true, token: token };
  }
  if (recordFail_(d) > GLOBAL_MAX_FAILS) Utilities.sleep(3000);   // 全体の失敗が多いときは応答を遅らせる（総当たり対策）
  if (c.get('ld_' + d)) return { ok: false, msg: 'ログインに5回失敗したため、この端末は1時間ログインできません。' };
  const left = DEVICE_MAX_FAILS - Number(c.get('fd_' + d) || 0);
  return { ok: false, msg: 'IDまたはパスワードが違います。（あと' + left + '回失敗すると、この端末は1時間ログインできません）' };
}
function checkSession(token) { return { ok: sessionOk_({ token: token }) }; }
function logout(token) {
  if (typeof token === 'string' && /^\w{32,128}$/.test(token)) CacheService.getScriptCache().remove('tk_' + token);
  return { ok: true };
}

/* ---------- 書き込み ---------- */
function apply(ops, auth, dev) {
  if (!Array.isArray(ops) || ops.length > 20000) return { error: '不正なリクエストです' };
  let tag = null;   // 端末タグ（addRec があるときだけ作る）
  const needAdmin = ops.some(o => !o || !PUBLIC_OPS[o.t]);
  if (needAdmin) {
    if (!sessionOk_(auth)) return { error: '管理者のログインが切れました。ログインし直してください。', authError: true };
  } else if (!rateOk_(ops.length, dev)) {
    return { error: '短時間に操作が集中しています。少し待ってからやり直してください。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const recRowIds = readIds_('records');   // idの列だけ（重複確認と行番号の対応に使う）
    const recIds = new Set(recRowIds);
    const rejected = [];   // 保存しなかった操作（画面に知らせる）
    const rej = (o, id, code, msg) => rejected.push({ t: o.t, id: String(id === undefined || id === null ? '' : id).slice(0, 80), code: code, msg: msg });
    const savedNames = new Set(), pendNames = [];   // ログインなしのとき、名前の追加は記録が保存された名前だけにする
    const newRows = [], updRows = new Map(), delIds = new Set();
    let names = readAll_('names').map(r => String(r[0])), namesDirty = false;
    // スクランブルは、バッチに操作があるときだけ読む（id列だけ。genScrの重複確認は日付が同じ行だけ中身を読む）
    const hasScr = ops.some(o => o.t === 'addScr' || o.t === 'updScr' || o.t === 'genScr' || o.t === 'delScr');
    const scrRowIds = hasScr ? readIds_('scrambles') : [];
    const scrSet = new Set(scrRowIds);
    const scrNew = [], scrUpd = new Map(), scrDel = new Set();
    let scrMeta = null;   // genScrの重複確認用：[id, date] の一覧（A〜C列）
    const scrPut = (c, add) => {   // 追加は末尾に追記、更新は（このバッチで追加した行なら差し替え、なければ）該当行を上書き
      const row = [c.id, c.date, c.event, JSON.stringify(c)];
      if (add) scrNew.push(row);
      else { const idx = scrNew.findIndex(x => x[0] === c.id); if (idx >= 0) scrNew[idx] = row; else scrUpd.set(c.id, row); }
      scrSet.add(c.id);
      log.push(['scrambles', 'put', c.id, row[3]]);
    };
    // 同じ日・種目・形式のスクランブルが既にあるか。日付が同じ行だけ中身を読んで調べる
    const scrDup = c => {
      const f = c.format || {}, same = x => x && x.date === c.date && x.event === c.event && x.format && x.format.type === f.type && x.format.n === f.n;
      if (scrNew.some(r => same(parse_(r[3])))) return true;
      if (!scrMeta) scrMeta = scrRowIds.length ? sh_('scrambles').getRange(2, 1, scrRowIds.length, 3).getValues() : [];
      for (let i = 0; i < scrMeta.length; i++) {
        const id = String(scrMeta[i][0]);
        if (scrDel.has(id)) continue;
        const d = scrMeta[i][1] instanceof Date ? Utilities.formatDate(scrMeta[i][1], APP_TZ, 'yyyy-MM-dd') : String(scrMeta[i][1]);
        if (d !== c.date) continue;
        const row = scrUpd.get(id);
        if (row) { if (same(parse_(row[3]))) return true; continue; }
        if (same(parse_(sh_('scrambles').getRange(i + 2, 4).getValue()))) return true;
      }
      return false;
    };
    let evNew = null;   // 種目の設定を更新するときだけ入る
    let bdNew = null;   // バッジの設定を更新するときだけ入る
    const log = [];   // [table, op, id, json]

    for (const o of ops) {
      switch (o.t) {
        case 'addRec': {
          const r = cleanRec_(o.rec);
          if (!r) { rej(o, o.rec && o.rec.id, 'invalid', '記録の内容が正しくありません'); break; }
          if (recIds.has(r.id)) { savedNames.add(r.name.trim()); break; }   // 再送で既に保存済み（エラーにしない）
          if (recIds.size >= MAX_RECORDS) { rej(o, r.id, 'limit', '保存できる記録数の上限に達しています'); break; }
          savedNames.add(r.name.trim());
          r.by = tag || (tag = devTag_(dev));   // 送信元の端末（画面から来た by は無視して上書き）
          newRows.push([r.id, r.date, r.event, r.name, JSON.stringify(r)]);
          recIds.add(r.id); log.push(['records', 'put', r.id, JSON.stringify(r)]);
          break;
        }
        case 'updRec': {
          const r = cleanRec_(o.rec);
          if (!r) { rej(o, o.rec && o.rec.id, 'invalid', '記録の内容が正しくありません'); break; }
          if (!recIds.has(r.id)) { rej(o, r.id, 'notfound', '更新する記録が見つかりません'); break; }
          const idx = newRows.findIndex(x => x[0] === r.id);
          const row = [r.id, r.date, r.event, r.name, JSON.stringify(r)];
          if (idx >= 0) newRows[idx] = row; else updRows.set(r.id, row);
          log.push(['records', 'put', r.id, JSON.stringify(r)]);
          break;
        }
        case 'delRec': {
          const id = String(o.id);
          if (!recIds.has(id)) { rej(o, id, 'notfound', '削除する記録が見つかりません'); break; }
          recIds.delete(id); updRows.delete(id);
          const idx = newRows.findIndex(x => x[0] === id);
          if (idx >= 0) newRows.splice(idx, 1); else delIds.add(id);
          log.push(['records', 'del', id, '']);
          break;
        }
        case 'addName': {
          const n = String(o.name || '').trim();
          if (!n || n.length > 80) { rej(o, n, 'invalid', '名前が正しくありません'); break; }
          if (needAdmin) { names = [n].concat(names.filter(x => x !== n)).slice(0, 500); namesDirty = true; }
          else pendNames.push([o, n]);   // ログインなし：記録の処理がすべて終わってから判断する
          break;
        }
        case 'delName':
          names = names.filter(x => x !== o.name); namesDirty = true;
          break;
        case 'addScr': case 'updScr': {
          const c = cleanScr_(o.scr);
          if (!c) { rej(o, o.scr && o.scr.id, 'invalid', 'スクランブルの内容が正しくありません'); break; }
          if (o.t === 'addScr' && scrSet.has(c.id)) break;
          if (o.t === 'updScr' && !scrSet.has(c.id)) { rej(o, c.id, 'notfound', '更新するスクランブルが見つかりません'); break; }
          scrPut(c, o.t === 'addScr');
          break;
        }
        case 'genScr': {
          // ログインなしで保存できるのは、「今日の日付」で「まだ保存されていない」種目・形式のスクランブルだけ
          const c = cleanScr_(o.scr);
          if (!c) { rej(o, o.scr && o.scr.id, 'invalid', 'スクランブルの内容が正しくありません'); break; }
          if (scrSet.has(c.id)) break;   // 再送で既に保存済み
          if (c.date !== Utilities.formatDate(new Date(), APP_TZ, 'yyyy-MM-dd')) { rej(o, c.id, 'date', '日付が今日（日本時間）ではないため、保存できませんでした'); break; }
          const f = c.format || {};
          const dup = scrDup(c);
          if (dup) { rej(o, c.id, 'dup', '同じ日・種目・形式のスクランブルが既に保存されています'); break; }
          scrPut(c, true);
          break;
        }
        case 'setEvents': {
          const c = cleanEvents_(o.events);
          if (c === null) { rej(o, '', 'invalid', '種目の設定が正しくありません'); break; }
          evNew = c; log.push(['events', 'set', '', JSON.stringify(c)]);
          break;
        }
        case 'setBadges': {
          const c = cleanBadges_(o.badges);
          if (c === null) { rej(o, '', 'invalid', 'バッジの設定が正しくありません'); break; }
          bdNew = c; log.push(['badges', 'set', '', JSON.stringify(c)]);
          break;
        }
        case 'delScr': {
          const id = String(o.id);
          if (!scrSet.delete(id)) { rej(o, id, 'notfound', '削除するスクランブルが見つかりません'); break; }
          scrUpd.delete(id);
          const idx = scrNew.findIndex(x => x[0] === id);
          if (idx >= 0) scrNew.splice(idx, 1); else scrDel.add(id);
          log.push(['scrambles', 'del', id, '']);
          break;
        }
      }
    }
    // ログインなしの名前の追加：既にある名前は先頭へ移すだけ。新しい名前は、同じ送信で記録が保存された名前だけ、一覧が500件未満のときに限り追加する
    for (const [o, n] of pendNames) {
      if (names.indexOf(n) >= 0) names = [n].concat(names.filter(x => x !== n));
      else if (savedNames.has(n) && names.length < 500) names = [n].concat(names);
      else { rej(o, n, 'invalid', '記録が保存されなかった、または名前の一覧がいっぱいのため、名前は追加しませんでした'); continue; }
      namesDirty = true;
      log.push(['name', 'put', n, JSON.stringify(n)]);   // ログインなしは名前ごとに記録する（一覧全体は記録しない）
    }
    if (namesDirty && needAdmin) log.push(['names', 'set', '', JSON.stringify(names)]);

    patchRows_('records', recRowIds, newRows, updRows, delIds);
    if (hasScr) patchRows_('scrambles', scrRowIds, scrNew, scrUpd, scrDel);
    if (namesDirty) writeRows_('names', names.map(n => [n]), false);
    if (evNew !== null) writeRows_('events', [[JSON.stringify(evNew)]], false);
    if (bdNew !== null) writeRows_('badges', [[JSON.stringify(bdNew)]], false);

    let rev = rev_();
    if (log.length) {
      rev += 1;
      writeRows_('changes', log.map(l => [rev, l[0], l[1], l[2], l[3]]), true);
      pruneLog_();
      props_().setProperty('REV', String(rev));   // 最後に更新（読む側は完成した履歴だけを見る）
      setMeta_(rev, minRev_());
      resumeBackupIfPaused_();
    }
    SpreadsheetApp.flush();
    return { ok: true, rev: rev, rejected: rejected };
  } finally {
    lock.releaseLock();
  }
}

function pruneLog_() {
  const s = sh_('changes'), n = s.getLastRow() - 1;
  if (n <= LOG_PRUNE_AT) return;
  const del = n - LOG_KEEP;
  const firstKeptRev = Number(s.getRange(2 + del, 1).getValue());
  s.deleteRows(2, del);
  props_().setProperty('MINREV', String(firstKeptRev));
}

/* ---------- 検証 ---------- */
// 実在する日付か（yyyy-MM-dd）
function dateOk_(d) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(d));
  if (!m) return false;
  const t = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  return t.getUTCFullYear() === Number(m[1]) && t.getUTCMonth() === Number(m[2]) - 1 && t.getUTCDate() === Number(m[3]);
}
// 記録は、画面が実際に使うフィールドだけを許可して作り直す（未知のキーは捨てる）。
// 画面側（L_Index.html）が記録・試技に設定するフィールドを増やしたら、ここにも足すこと。
function intIn_(v, min, max) { return typeof v === 'number' && Number.isInteger(v) && v >= min && v <= max; }
function cleanAtt_(a) {
  if (!a || typeof a !== 'object' || Array.isArray(a)) return null;
  if (a.kind !== 'OK' && a.kind !== 'DNF' && a.kind !== 'DNS') return null;
  const o = { kind: a.kind };
  if (a.time === null) o.time = null;
  else if (a.time !== undefined) {
    if (typeof a.time !== 'number' || !isFinite(a.time) || a.time < 0 || a.time > 10000000) return null;
    o.time = a.time;
  }
  if (a.plus2 !== undefined) { if (typeof a.plus2 !== 'boolean') return null; o.plus2 = a.plus2; }
  // MBLD
  if (a.points !== undefined) { if (!intIn_(a.points, -10000, 10000)) return null; o.points = a.points; }
  if (a.solved !== undefined) { if (!intIn_(a.solved, 0, 10000)) return null; o.solved = a.solved; }
  if (a.attempted !== undefined) { if (!intIn_(a.attempted, 0, 10000)) return null; o.attempted = a.attempted; }
  // Exスクランブルの試技（1か2）
  if (a.ex !== undefined) { if (!intIn_(a.ex, 1, 9)) return null; o.ex = a.ex; }
  // FMCの解答とコメント
  if (a.sol !== undefined) { if (typeof a.sol !== 'string' || a.sol.length > 1000) return null; o.sol = a.sol; }
  if (a.cm !== undefined) { if (typeof a.cm !== 'string' || a.cm.length > 1000) return null; o.cm = a.cm; }
  return o;
}
function cleanRec_(r) {
  if (!r || typeof r !== 'object' || Array.isArray(r)) return null;
  if (JSON.stringify(r).length > 4000) return null;
  if (!/^[\w-]{1,40}$/.test(String(r.id))) return null;
  if (!dateOk_(r.date)) return null;
  if (typeof r.event !== 'string' || !r.event || r.event.length > 80) return null;
  if (typeof r.name !== 'string' || !r.name.trim() || r.name.length > 80) return null;
  if (!Array.isArray(r.attempts) || r.attempts.length < 1 || r.attempts.length > 5) return null;
  const o = { id: String(r.id), date: String(r.date), event: r.event, name: r.name };
  if (r.format !== undefined && r.format !== null) {
    const f = r.format;
    if (typeof f !== 'object' || Array.isArray(f) || ['ao', 'mo', 'bo'].indexOf(f.type) < 0 || !intIn_(f.n, 1, 5)) return null;
    o.format = { type: f.type, n: f.n };
  }
  o.attempts = [];
  for (const a of r.attempts) {
    const c = cleanAtt_(a);
    if (!c) return null;
    o.attempts.push(c);
  }
  // best / worst / average は画面で再計算される。保存値は文字列・数値・null だけ
  for (const k of ['best', 'worst', 'average']) {
    const v = r[k];
    if (v === undefined) continue;
    if (v === null || (typeof v === 'number' && isFinite(v)) || (typeof v === 'string' && v.length <= 40)) o[k] = v;
    else return null;
  }
  for (const k of ['createdAt', 'updatedAt']) {
    const v = r[k];
    if (v === undefined) continue;
    if (typeof v !== 'string' || v.length > 40) return null;
    o[k] = v;
  }
  if (r.by !== undefined) { if (typeof r.by !== 'string' || !/^[0-9a-f]{12}$/.test(r.by)) return null; o.by = r.by; }   // 送信元の端末タグ（既存の記録にはない）
  return o;
}
function cleanScr_(s) {
  if (!s || typeof s !== 'object') return null;
  if (JSON.stringify(s).length > 15000) return null;
  if (!/^[\w-]{1,40}$/.test(String(s.id))) return null;
  if (!dateOk_(s.date)) return null;
  if (typeof s.event !== 'string' || !s.event || s.event.length > 80) return null;
  if (!Array.isArray(s.scrambles) || s.scrambles.length > 5) return null;
  s.id = String(s.id);
  return s;
}

// 種目の設定の検証。空（初期設定に戻す）は [] を返し、不正なら null を返す。
function cleanEvents_(a) {
  if (a === null || a === undefined || (Array.isArray(a) && !a.length)) return [];
  if (!Array.isArray(a) || a.length > 200) return null;
  const seen = {}, out = [];
  for (const e of a) {
    if (!e || typeof e !== 'object') return null;
    const name = typeof e.name === 'string' ? e.name.trim() : '';
    const short = typeof e.short === 'string' ? e.short.trim() : '';
    if (!name || name.length > 80 || name === '__other__' || seen[name]) return null;
    if (short.length > 20) return null;
    if (BUILTIN_EVENTS.indexOf(e.scr) < 0) return null;
    seen[name] = true;
    out.push({ name: name, short: short, scr: e.scr, official: e.official !== false });   // official が未設定の種目は公式とみなす
  }
  return out;
}

// バッジの設定の検証。空（初期設定に戻す）は {} を返し、不正なら null を返す。
const BDG_CATS_ = ['cr', 'gc', 'pbq', 'single', 'kyoyo', 'sub', 'streak', 'nodnf', 'rep'];
const BDG_STAGED_ = ['gc', 'single', 'kyoyo', 'streak', 'nodnf', 'rep'];
const BDG_KYO_IDS_ = ['333', '222', '444', '555', '666', '777', '333bf', '333fm', '333oh', 'clock', 'minx', 'pyram', 'skewb', 'sq1', '444bf', '555bf', '333mbf', 'fto'];
const BDG_PBQ_IDS_ = ['bld1', 'bld2', 'bld3', 'fmc', 'fmcbld', 'all'];
const BDG_SUB_ALL_ = [60, 45, 30, 20, 15, 12, 11, 10, 9, 8, 7.5, 7, 6.5, 6, 5.5, 5, 4.5, 4, 3.5, 3, 2.5, 2, 1.5, 1, 0.5];   // 画像のある基準だけ
function cleanBadges_(c) {
  if (c === null || c === undefined) return {};
  if (typeof c !== 'object' || Array.isArray(c)) return null;
  const keys = Object.keys(c);
  if (!keys.length) return {};
  if (JSON.stringify(c).length > 30000) return null;
  if (keys.some(k => BDG_CATS_.indexOf(k) < 0)) return null;
  const str = (s, max, req) => {
    if (typeof s !== 'string') return null;
    s = s.trim();
    if ((req && !s) || s.length > max) return null;
    return s;
  };
  const tierOk = t => t === 'bronze' || t === 'silver' || t === 'gold';
  const out = {};
  for (const k of BDG_CATS_) {
    const v = c[k];
    if (v === undefined) continue;
    if (!v || typeof v !== 'object' || Array.isArray(v)) return null;
    const o = { on: v.on !== false };
    const d = str(v.desc === undefined ? '' : v.desc, 200, false);
    if (d === null) return null;
    o.desc = d;
    if (k === 'cr') {
      const n = str(v.name, 60, true); if (n === null) return null;
      o.name = n;
    } else if (BDG_STAGED_.indexOf(k) >= 0) {
      if (!Array.isArray(v.stages) || v.stages.length > 30) return null;
      const seen = {}, st = [];
      for (const s of v.stages) {
        if (!s || typeof s !== 'object') return null;
        const min = Number(s.min), name = str(s.name, 60, true);
        if (!(min >= 1 && min <= 100000) || Math.floor(min) !== min || seen[min]) return null;
        if (name === null || !tierOk(s.tier)) return null;
        seen[min] = true; st.push({ min: min, name: name, tier: s.tier });
      }
      st.sort((a, b) => a.min - b.min);
      o.stages = st;
      if (k === 'kyoyo') {
        // 種目ごとの基準。value が null の種目は「基準未設定（対象外）」
        if (!Array.isArray(v.events) || v.events.length > BDG_KYO_IDS_.length) return null;
        const seenE = {}, ev = [];
        for (const e of v.events) {
          if (!e || typeof e !== 'object' || BDG_KYO_IDS_.indexOf(e.id) < 0 || seenE[e.id]) return null;
          if (e.metric !== 'single' && e.metric !== 'average' && e.metric !== 'points') return null;
          if ((e.metric === 'points') !== (e.id === '333mbf')) return null;   // 点数はMBLDだけ、MBLDは点数だけ
          let val = e.value;
          if (val !== null) {
            val = Number(val);
            if (!(val > 0 && val <= 1000000)) return null;
            val = Math.round(val * 100) / 100;
          }
          seenE[e.id] = true; ev.push({ id: e.id, metric: e.metric, value: val });
        }
        o.events = ev;
      }
    } else if (k === 'pbq') {
      if (!Array.isArray(v.items) || v.items.length !== BDG_PBQ_IDS_.length) return null;
      const seen = {}, it = [];
      for (const s of v.items) {
        if (!s || typeof s !== 'object' || BDG_PBQ_IDS_.indexOf(s.id) < 0 || seen[s.id]) return null;
        const name = str(s.name, 60, true);
        if (name === null || !tierOk(s.tier)) return null;
        seen[s.id] = true; it.push({ id: s.id, name: name, tier: s.tier });
      }
      o.items = it;
    } else if (k === 'sub') {
      const n = str(v.name, 80, true); if (n === null) return null;
      o.name = n;
      if (!Array.isArray(v.values) || v.values.length > BDG_SUB_ALL_.length) return null;
      const seen = {}, vals = [];
      for (const x of v.values) {
        if (BDG_SUB_ALL_.indexOf(x) < 0 || seen[x]) return null;
        seen[x] = true; vals.push(x);
      }
      vals.sort((a, b) => b - a);
      o.values = vals;
    }
    out[k] = o;
  }
  return out;
}

/* ---------- 認証・ロック・連投制限 ---------- */
function devKey_(dev) {
  return (typeof dev === 'string' && /^[\w-]{8,64}$/.test(dev)) ? dev : 'anon';
}
function credOk_(id, pw) {
  if (typeof id !== 'string' || typeof pw !== 'string') return false;
  const p = props_();
  const aid = p.getProperty('ADMIN_ID'), apw = p.getProperty('ADMIN_PASS');
  if (!aid || !apw) return false;
  return (eq_(id, aid) & eq_(pw, apw)) === 1;
}
function recordFail_(d) {
  const c = CacheService.getScriptCache();
  // 端末ごと：5回失敗で1時間ロック
  const n = Number(c.get('fd_' + d) || 0) + 1;
  if (n >= DEVICE_MAX_FAILS) { c.put('ld_' + d, '1', DEVICE_LOCK_SEC); c.put('fd_' + d, String(n), DEVICE_LOCK_SEC); }
  else c.put('fd_' + d, String(n), DEVICE_LOCK_SEC);
  // 全体：10分間（10分ごとの集計）の失敗数を数えて返す
  const k = 'gf' + Math.floor(Date.now() / (GLOBAL_WINDOW_SEC * 1000));
  const g = Number(c.get(k) || 0) + 1;
  c.put(k, String(g), GLOBAL_WINDOW_SEC * 2);
  return g;
}
function sessionOk_(a) {
  if (!a || typeof a.token !== 'string' || !/^\w{32,128}$/.test(a.token)) return false;
  return CacheService.getScriptCache().get('tk_' + a.token) === '1';
}
function eq_(a, b) {
  if (a.length !== b.length) return 0;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0 ? 1 : 0;
}
function rateOk_(n, dev) {
  const c = CacheService.getScriptCache();
  const m = Math.floor(Date.now() / 60000);
  const kd = 'rd_' + devKey_(dev) + '_' + m, kg = 'rl' + m;   // 端末ごと・全体
  const vd = Number(c.get(kd) || 0) + n, vg = Number(c.get(kg) || 0) + n;
  if (vd > PUBLIC_LIMIT_PER_DEV || vg > PUBLIC_LIMIT_PER_MIN) return false;   // 超えたら、どちらも加算しない
  c.put(kd, String(vd), 120); c.put(kg, String(vg), 120);
  return true;
}
// 送信元の端末タグ（端末IDのハッシュの先頭12文字）。端末ID自体は記録に残さない（Rから誰でも記録を読めるため）
function devTag_(dev) {
  const p = props_();
  let salt = p.getProperty('DEV_SALT');
  if (!salt) { salt = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('DEV_SALT', salt); }
  const h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + devKey_(dev));
  return h.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('').slice(0, 12);
}

/* ---------- バックアップ ---------- */
const BACKUP_FOLDER_NAME = 'Cutech Log バックアップ';
const BACKUP_KEEP = 30;           // 新しい順にこの件数だけ残し、古いものはゴミ箱へ
const BACKUP_IDLE_WEEKS = 15;     // これだけ変更がなければ自動バックアップを一時停止する

// 1回だけ実行する。毎週月曜の午前3時台に backup を動かすトリガーを作り直す
// （以前に作ったバックアップ用のトリガーは、ここで置き換えられる）。
function setupBackupTrigger() {
  createBackupTrigger_();
  props_().deleteProperty('BACKUP_PAUSED');
}

// スプレッドシートのコピーをドライブに作る。週1回、トリガーで自動実行される。
// 手動で実行したときは、変更がなくても必ずバックアップを作る。
function backup(e) {
  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const auto = !!(e && e.triggerUid);
  if (auto && Date.now() - file.getLastUpdated().getTime() >= BACKUP_IDLE_WEEKS * 7 * 24 * 3600 * 1000) {
    pauseBackup_();   // 長い間変更がないので、自動バックアップを一時停止する
    return;
  }
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  file.makeCopy('Cutech Log ' + stamp, folder);
  const files = [], fi = folder.getFiles();
  while (fi.hasNext()) files.push(fi.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(BACKUP_KEEP).forEach(f => f.setTrashed(true));
}

function createBackupTrigger_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backup').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backup').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
}
function pauseBackup_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backup').forEach(t => ScriptApp.deleteTrigger(t));
  props_().setProperty('BACKUP_PAUSED', '1');
}
// 一時停止中にデータが変更されたら、自動バックアップを再開する（apply から呼ばれる）。
function resumeBackupIfPaused_() {
  try {
    if (props_().getProperty('BACKUP_PAUSED') !== '1') return;
    createBackupTrigger_();
    props_().deleteProperty('BACKUP_PAUSED');
  } catch (err) { /* 権限がない場合などは無視（データの保存には影響させない） */ }
}

// バックアップや版の履歴からデータを戻したあとに1回だけ実行する。
// 開いている画面が古い状態のまま残らないよう、全員が全件を読み直すようにする。
function resetAfterRestore() {
  const s = sh_('changes');
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  const rev = rev_() + 1;
  props_().setProperty('REV', String(rev));
  props_().setProperty('MINREV', String(rev));
  setMeta_(rev, rev);
}
