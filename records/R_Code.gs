/**
 * Cutech Records — 記録閲覧用サイト（読み取り専用）
 *
 * 記録入力用の Cutech Log とは別の Apps Script プロジェクトです。
 * このプロジェクトには、記録を書き込む関数も、管理者ログインの関数もありません。
 * 入力用のスプレッドシートを「読む」ことだけができます。
 *
 * 使い方：
 *  1. 「プロジェクトの設定 > スクリプト プロパティ」に、プロパティ SHEET_ID を追加し、
 *     値に入力用スプレッドシートのIDを入れる（最初の1回だけ。コードを貼り替えても消えません）。
 *     （スプレッドシートのURL  https://docs.google.com/spreadsheets/d/ ←ここの部分→ /edit  がID）
 *  2. HTMLファイルの名前は R_Index にする（この名前で読み込みます）。
 *  3. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 */
// スプレッドシートのIDはスクリプト プロパティ SHEET_ID から読む（コードの更新のたびに貼り直さなくてよいように）
function sheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('スクリプト プロパティ SHEET_ID が設定されていません');
  return id.trim();
}

const WIDTH = { records: 5, names: 1, scrambles: 4, changes: 5, events: 1, badges: 1 };
const DIFF_MAX = 2000;   // これ以上の差分は全件読み込みにする

function doGet() {
  return HtmlService.createHtmlOutputFromFile('R_Index')
    .setTitle('Cutech Records')
    .setFaviconUrl('https://raw.githubusercontent.com/2023KADO01/Cutech-App/main/assets/cutechicon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, interactive-widget=resizes-content');
}

/* ---------- 読み取り ---------- */
let ss = null;   // 1回の実行の中では開いた結果を使い回す
function ss_() { return ss || (ss = SpreadsheetApp.openById(sheetId_())); }

function readAll_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) return [];
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, WIDTH[name]).getValues();
}
function parse_(v) { try { return JSON.parse(v); } catch (e) { return null; } }

// 入力用サイトが metaシートに書いた版番号を読む（なければ変更履歴から推定する）
// 結果は15秒だけキャッシュする（fresh が true のときは必ずシートを読む）
function meta_(fresh) {
  const c = CacheService.getScriptCache();
  if (!fresh) { try { const v = c.get('meta'); if (v) return JSON.parse(v); } catch (e) { /* キャッシュが読めなければシートを読む */ } }
  let m;
  const s = ss_().getSheetByName('meta');
  if (s && s.getLastRow() >= 3) {
    const v = s.getRange(2, 2, 2, 1).getValues();
    m = { rev: Number(v[0][0]) || 0, min: Number(v[1][0]) || 0 };
  } else {
    const ch = readAll_('changes');
    m = { rev: ch.length ? Number(ch[ch.length - 1][0]) : 0, min: ch.length ? Number(ch[0][0]) : 0 };
  }
  try { c.put('meta', JSON.stringify(m), 15); } catch (e) { /* 保存できなくても読み取りには影響させない */ }
  return m;
}

function getAll() {
  const rev = meta_(true).rev;   // データより先に読む（途中で更新されても差分で補正される）
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

// バッジの設定（管理者が編集したもの）。未設定のときは空のオブジェクト（画面側の初期設定を使う）
function readBadges_() {
  const r = readAll_('badges');
  const v = r.length ? parse_(r[0][0]) : null;
  return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
}

// 前回の版（since）以降の変更だけを返す。履歴が足りない場合は全件を返す。
function getChanges(since) {
  let m = meta_();   // キャッシュ（15秒）
  since = Number(since);
  if (since > m.rev) m = meta_(true);   // 手元の方が新しいときは、キャッシュが古いだけかもしれないので読み直す
  if (!(since >= 0) || since > m.rev || since < m.min) return fullResult_();
  if (since === m.rev) return { rev: m.rev, changes: [] };
  const rows = changesSince_(since);
  if (rows === null) return fullResult_();
  const changes = [];
  for (const r of rows) {
    const v = parse_(r[4]);
    if (r[2] === 'del') changes.push({ t: String(r[1]), op: 'del', id: String(r[3]) });
    else if (v !== null) changes.push({ t: String(r[1]), op: r[2] === 'set' ? 'set' : 'put', id: String(r[3]), v: v });
  }
  return { rev: m.rev, changes: changes };
}
// changes は rev の昇順。A列だけ読んで二分探索し、since より新しい行だけ5列読む（多すぎるときは null）
function changesSince_(since) {
  const s = ss_().getSheetByName('changes');
  const n = s ? s.getLastRow() - 1 : 0;
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
