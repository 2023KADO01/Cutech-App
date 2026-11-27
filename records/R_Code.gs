/**
 * Cutech Records — 記録閲覧用サイト（読み取り専用）
 *
 * 記録入力用の Cutech Log とは別の Apps Script プロジェクトです。
 * このプロジェクトには、記録を書き込む関数も、管理者ログインの関数もありません。
 * 入力用のスプレッドシートを「読む」ことだけができます。
 *
 * 使い方：
 *  1. 下の SHEET_ID に、入力用スプレッドシートのIDを貼り付ける。
 *     （スプレッドシートのURL  https://docs.google.com/spreadsheets/d/ ←ここの部分→ /edit  がID）
 *  2. HTMLファイルの名前は R_Index にする（この名前で読み込みます）。
 *  3. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 */
const SHEET_ID = 'ここにスプレッドシートのIDを貼り付けてください';

const WIDTH = { records: 5, names: 1, scrambles: 4, changes: 5, events: 1, badges: 1 };
const DIFF_MAX = 2000;   // これ以上の差分は全件読み込みにする

function doGet() {
  return HtmlService.createHtmlOutputFromFile('R_Index')
    .setTitle('Cutech Records')
    .setFaviconUrl('https://2023kado01.github.io/cutech-icon/cutechicon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, interactive-widget=resizes-content')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- 読み取り ---------- */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function readAll_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) return [];
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, WIDTH[name]).getValues();
}
function parse_(v) { try { return JSON.parse(v); } catch (e) { return null; } }

// 入力用サイトが metaシートに書いた版番号を読む（なければ変更履歴から推定する）
function meta_() {
  const s = ss_().getSheetByName('meta');
  if (s && s.getLastRow() >= 3) {
    const v = s.getRange(2, 2, 2, 1).getValues();
    return { rev: Number(v[0][0]) || 0, min: Number(v[1][0]) || 0 };
  }
  const ch = readAll_('changes');
  return { rev: ch.length ? Number(ch[ch.length - 1][0]) : 0, min: ch.length ? Number(ch[0][0]) : 0 };
}

function getAll() {
  const rev = meta_().rev;   // データより先に読む（途中で更新されても差分で補正される）
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
  const m = meta_();
  since = Number(since);
  if (!(since >= 0) || since > m.rev || since < m.min) return fullResult_();
  if (since === m.rev) return { rev: m.rev, changes: [] };
  const rows = readAll_('changes').filter(r => Number(r[0]) > since);
  if (rows.length > DIFF_MAX) return fullResult_();
  const changes = [];
  for (const r of rows) {
    const v = parse_(r[4]);
    if (r[2] === 'del') changes.push({ t: String(r[1]), op: 'del', id: String(r[3]) });
    else if (v !== null) changes.push({ t: String(r[1]), op: r[2] === 'set' ? 'set' : 'put', id: String(r[3]), v: v });
  }
  return { rev: m.rev, changes: changes };
}
function fullResult_() {
  const d = getAll();
  return { full: true, rev: d.rev, data: d };
}
