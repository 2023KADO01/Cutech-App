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
const PUBLIC_LIMIT_PER_MIN = 120;                // 公開書き込みの連投制限（全体）
const MAX_RECORDS = 50000;

// 管理者ログインの保護
const DEVICE_MAX_FAILS = 5;       // 端末ごと：5回失敗で
const DEVICE_LOCK_SEC = 3600;     //   その端末を1時間ロック
const GLOBAL_MAX_FAILS = 100;     // 全体：10分間に100回失敗で
const GLOBAL_WINDOW_SEC = 600;    //   （10分ごとに集計）
const GLOBAL_LOCK_SEC = 600;      //   全員を10分ロック
const SESSION_SEC = 21600;        // ログイン保持の上限（6時間：CacheServiceの上限）

// 差分読み込み
const LOG_KEEP = 5000;            // 変更履歴の保持件数
const LOG_PRUNE_AT = 6000;        // これを超えたら古い履歴を削除
const DIFF_MAX = 2000;            // これ以上の差分は全件読み込みにする

function doGet() {
  return HtmlService.createHtmlOutputFromFile('L_Index')
    .setTitle('Cutech Log')
    .setFaviconUrl('https://2023kado01.github.io/cutech-icon/cutechicon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, interactive-widget=resizes-content')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  let start = 2;
  if (append) start = Math.max(s.getLastRow() + 1, 2);
  else if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, w).clearContent();
  if (!rows.length) return;
  const r = s.getRange(start, 1, rows.length, w);
  r.setNumberFormat('@');            // 日付や数字が自動変換されないよう文字列固定
  r.setValues(rows);
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
  const rows = readAll_('changes').filter(r => Number(r[0]) > since);
  if (rows.length > DIFF_MAX) return fullResult_();
  const changes = [];
  for (const r of rows) {
    const v = parse_(r[4]);
    if (r[2] === 'del') changes.push({ t: String(r[1]), op: 'del', id: String(r[3]) });
    else if (v !== null) changes.push({ t: String(r[1]), op: r[2] === 'set' ? 'set' : 'put', id: String(r[3]), v: v });
  }
  return { rev: rev, changes: changes };
}
function fullResult_() {
  const d = getAll();
  return { full: true, rev: d.rev, data: d };
}

/* ---------- ログイン ---------- */
function login(id, pw, dev) {
  const c = CacheService.getScriptCache();
  if (c.get('gl')) return { ok: false, msg: 'ログインの失敗が多発したため、全員のログインを一時停止しています。10分ほど待ってください。' };
  const d = devKey_(dev);
  if (c.get('ld_' + d)) return { ok: false, msg: 'この端末は、ログインに5回失敗したため、1時間ログインできません。' };
  if (credOk_(id, pw)) {
    c.remove('fd_' + d);
    const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    c.put('tk_' + token, '1', SESSION_SEC);
    return { ok: true, token: token };
  }
  recordFail_(d);
  if (c.get('gl')) return { ok: false, msg: 'ログインの失敗が多発したため、全員のログインを一時停止しています。10分ほど待ってください。' };
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
  const needAdmin = ops.some(o => !o || !PUBLIC_OPS[o.t]);
  if (needAdmin) {
    if (!sessionOk_(auth)) return { error: '管理者のログインが切れました。ログインし直してください。', authError: true };
  } else if (!rateOk_(ops.length)) {
    return { error: '短時間に操作が集中しています。少し待ってからやり直してください。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const recs = readAll_('records');
    const recIds = new Set(recs.map(r => String(r[0])));
    const newRows = [], updRows = new Map(), delIds = new Set();
    let names = readAll_('names').map(r => String(r[0])), namesDirty = false;
    const scrMap = new Map(readAll_('scrambles').map(r => [String(r[0]), r]));
    let scrDirty = false;
    let evNew = null;   // 種目の設定を更新するときだけ入る
    let bdNew = null;   // バッジの設定を更新するときだけ入る
    const log = [];   // [table, op, id, json]

    for (const o of ops) {
      switch (o.t) {
        case 'addRec': {
          const r = cleanRec_(o.rec);
          if (!r || recIds.has(r.id) || recIds.size >= MAX_RECORDS) break;
          newRows.push([r.id, r.date, r.event, r.name, JSON.stringify(r)]);
          recIds.add(r.id); log.push(['records', 'put', r.id, JSON.stringify(r)]);
          break;
        }
        case 'updRec': {
          const r = cleanRec_(o.rec);
          if (!r || !recIds.has(r.id)) break;
          const idx = newRows.findIndex(x => x[0] === r.id);
          const row = [r.id, r.date, r.event, r.name, JSON.stringify(r)];
          if (idx >= 0) newRows[idx] = row; else updRows.set(r.id, row);
          log.push(['records', 'put', r.id, JSON.stringify(r)]);
          break;
        }
        case 'delRec': {
          const id = String(o.id);
          if (!recIds.has(id)) break;
          recIds.delete(id); updRows.delete(id);
          const idx = newRows.findIndex(x => x[0] === id);
          if (idx >= 0) newRows.splice(idx, 1); else delIds.add(id);
          log.push(['records', 'del', id, '']);
          break;
        }
        case 'addName': {
          const n = String(o.name || '').trim();
          if (!n || n.length > 80) break;
          names = [n].concat(names.filter(x => x !== n)).slice(0, 500); namesDirty = true;
          break;
        }
        case 'delName':
          names = names.filter(x => x !== o.name); namesDirty = true;
          break;
        case 'addScr': case 'updScr': {
          const c = cleanScr_(o.scr); if (!c) break;
          if (o.t === 'addScr' && scrMap.has(c.id)) break;
          if (o.t === 'updScr' && !scrMap.has(c.id)) break;
          scrMap.set(c.id, [c.id, c.date, c.event, JSON.stringify(c)]); scrDirty = true;
          log.push(['scrambles', 'put', c.id, JSON.stringify(c)]);
          break;
        }
        case 'genScr': {
          // ログインなしで保存できるのは、「今日の日付」で「まだ保存されていない」種目・形式のスクランブルだけ
          const c = cleanScr_(o.scr); if (!c) break;
          if (scrMap.has(c.id)) break;
          if (c.date !== Utilities.formatDate(new Date(), APP_TZ, 'yyyy-MM-dd')) break;
          const f = c.format || {};
          let dup = false;
          scrMap.forEach(r => {
            const x = parse_(r[3]);
            if (x && x.date === c.date && x.event === c.event && x.format && x.format.type === f.type && x.format.n === f.n) dup = true;
          });
          if (dup) break;
          scrMap.set(c.id, [c.id, c.date, c.event, JSON.stringify(c)]); scrDirty = true;
          log.push(['scrambles', 'put', c.id, JSON.stringify(c)]);
          break;
        }
        case 'setEvents': {
          const c = cleanEvents_(o.events); if (c === null) break;
          evNew = c; log.push(['events', 'set', '', JSON.stringify(c)]);
          break;
        }
        case 'setBadges': {
          const c = cleanBadges_(o.badges); if (c === null) break;
          bdNew = c; log.push(['badges', 'set', '', JSON.stringify(c)]);
          break;
        }
        case 'delScr': {
          const id = String(o.id);
          if (!scrMap.delete(id)) break;
          scrDirty = true; log.push(['scrambles', 'del', id, '']);
          break;
        }
      }
    }
    if (namesDirty) log.push(['names', 'set', '', JSON.stringify(names)]);

    if (delIds.size || updRows.size) {
      const kept = recs.filter(r => !delIds.has(String(r[0]))).map(r => updRows.get(String(r[0])) || r);
      writeRows_('records', kept.concat(newRows), false);
    } else if (newRows.length) {
      writeRows_('records', newRows, true);
    }
    if (namesDirty) writeRows_('names', names.map(n => [n]), false);
    if (scrDirty) writeRows_('scrambles', Array.from(scrMap.values()), false);
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
    return { ok: true, rev: rev };
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
function cleanRec_(r) {
  if (!r || typeof r !== 'object') return null;
  if (JSON.stringify(r).length > 4000) return null;
  if (!/^[\w-]{1,40}$/.test(String(r.id))) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r.date))) return null;
  if (typeof r.event !== 'string' || !r.event || r.event.length > 80) return null;
  if (typeof r.name !== 'string' || !r.name.trim() || r.name.length > 80) return null;
  if (!Array.isArray(r.attempts) || r.attempts.length < 1 || r.attempts.length > 5) return null;
  r.id = String(r.id);
  return r;
}
function cleanScr_(s) {
  if (!s || typeof s !== 'object') return null;
  if (JSON.stringify(s).length > 15000) return null;
  if (!/^[\w-]{1,40}$/.test(String(s.id))) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s.date))) return null;
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
  // 全体：10分間（10分ごとの集計）に100回失敗で全員を10分ロック
  const k = 'gf' + Math.floor(Date.now() / (GLOBAL_WINDOW_SEC * 1000));
  const g = Number(c.get(k) || 0) + 1;
  c.put(k, String(g), GLOBAL_WINDOW_SEC * 2);
  if (g >= GLOBAL_MAX_FAILS) c.put('gl', '1', GLOBAL_LOCK_SEC);
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
function rateOk_(n) {
  const c = CacheService.getScriptCache();
  const k = 'rl' + Math.floor(Date.now() / 60000);
  const v = Number(c.get(k) || 0) + n;
  if (v > PUBLIC_LIMIT_PER_MIN) return false;
  c.put(k, String(v), 120);
  return true;
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
