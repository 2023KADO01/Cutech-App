// cubing.js（スクランブル生成ライブラリ）の配布元ビルドを取得して、リポジトリに固定コピーする。
// 使い方: node .github/scripts/fetch-cubing.mjs   （リポジトリのルートで実行）
// 環境変数（テスト用）: CUBING_BASE … 配布元（既定 https://cdn.cubing.net/v0/）、REPO_ROOT … リポジトリのルート（既定 カレントディレクトリ）
import { createHash } from 'node:crypto';
import { appendFileSync } from 'node:fs';
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const BASE = process.env.CUBING_BASE || 'https://cdn.cubing.net/v0/';
const ROOT = path.resolve(process.env.REPO_ROOT || process.cwd());
const ENTRY_URL = new URL('js/cubing/scramble', BASE).href;   // 配布元でのエントリ（拡張子なし）
const ENTRY_PATH = 'js/cubing/scramble.js';                   // 保存時の名前
const OUT_URL_PREFIX = 'https://cdn.jsdelivr.net/gh/2023KADO01/Cutech-App@main/vendor/cubing/';
const CONCURRENCY = 8;

// 相対参照（./ か ../ で始まる .js）を探す。引用符は " ' どちらでも、空白は有っても無くてもよい。
//   静的 import（from"./x.js" / import"./x.js"）と動的 import（import("./x.js")）→ 必須
//   import.meta.resolve("./x.js") / new URL("./x.js",import.meta.url) → 予備の経路。ハッシュなしの名前なら任意
const REF_RE = /(\bfrom|\bimport\s*\(|\bimport|\bresolve\s*\(|\bnew\s+URL\s*\()\s*(["'])(\.{1,2}\/[^"'\s]+?\.js)\2/g;
const HASHED_RE = /-(?=[A-Za-z0-9_]*[A-Z0-9])[A-Za-z0-9_]{6,}\.js$/;   // 末尾が -英数字6字以上（大文字か数字を含む）＝ハッシュ付き

export function findRefs(code) {
  const refs = [];
  for (const m of code.matchAll(REF_RE)) {
    const viaFallback = /^(\bresolve|\bnew)/.test(m[1]);
    const spec = m[3];
    refs.push({ spec, optional: viaFallback && !HASHED_RE.test(spec) });
  }
  return refs;
}

export function toRelPath(urlStr) {
  if (urlStr === ENTRY_URL) return ENTRY_PATH;
  const base = new URL(BASE);
  const u = new URL(urlStr);
  if (u.origin !== base.origin || !u.pathname.startsWith(base.pathname)) throw new Error(`配布元の外です: ${urlStr}`);
  return decodeURIComponent(u.pathname.slice(base.pathname.length));
}

async function download(url) {
  const res = await fetch(url, { redirect: 'follow' });
  if (new URL(res.url).origin !== new URL(BASE).origin) throw new Error(`配布元の外にリダイレクトされました: ${res.url}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

async function crawl() {
  const files = new Map();                     // url -> Buffer
  const required = new Set([ENTRY_URL]);       // 必須のurl
  const seen = new Set([ENTRY_URL]);
  const failed = new Map();                    // url -> エラー
  let queue = [ENTRY_URL];
  while (queue.length) {
    const next = [];
    for (let i = 0; i < queue.length; i += CONCURRENCY) {
      await Promise.all(queue.slice(i, i + CONCURRENCY).map(async (url) => {
        try { files.set(url, await download(url)); } catch (e) { failed.set(url, e.message); return; }
        for (const { spec, optional } of findRefs(files.get(url).toString('utf8'))) {
          const abs = new URL(spec, url);
          abs.hash = ''; abs.search = '';
          if (abs.origin !== new URL(BASE).origin) continue;   // 配布元以外は取りに行かない
          const key = abs.href;
          if (!optional) required.add(key);
          if (!seen.has(key)) { seen.add(key); next.push(key); }
        }
      }));
    }
    queue = next;
  }
  for (const [url, msg] of failed) {
    if (required.has(url)) throw new Error(`必須ファイルを取得できませんでした: ${url} (${msg})`);
    console.warn(`警告: 任意のファイルを取得できなかったので飛ばします: ${url} (${msg})`);
  }
  return files;
}

export function versionId(entries) {   // entries: [相対パス, Buffer][]
  const h = createHash('sha256');
  for (const [p, buf] of [...entries].sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))) {
    h.update(p); h.update('\0'); h.update(buf); h.update('\0');
  }
  return h.digest('hex').slice(0, 12);
}

function setOutput(changed, id) {
  if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `changed=${changed}\nversion=${id}\n`);
}

async function main() {
  const files = await crawl();
  const entries = [...files].map(([url, buf]) => [toRelPath(url), buf]);
  console.log(`取得したファイル数: ${entries.length}`);
  const id = versionId(entries);
  console.log(`版ID: ${id}`);

  const currentFile = path.join(ROOT, 'vendor/cubing/current.txt');
  let current = '';
  try { current = (await readFile(currentFile, 'utf8')).trim(); } catch { /* 初回 */ }
  if (current === id) { console.log('変更なし'); setOutput(false, id); return; }

  // L_Index.html の行を先に確かめる（見つからなければ何も書かない）
  const htmlPath = path.join(ROOT, 'log/L_Index.html');
  const html = await readFile(htmlPath, 'utf8');
  const lineRe = /^(\s*const\s+CUBING_URL\s*=\s*)(['"])[^'"\n]*\2(\s*;)/m;
  if (!lineRe.test(html)) throw new Error('log/L_Index.html に CUBING_URL の行が見つかりません');

  // 一時ディレクトリに組み立ててから、まとめてリポジトリに置く
  const tmp = await mkdtemp(path.join(tmpdir(), 'cubing-'));
  try {
    for (const [p, buf] of entries) {
      const dest = path.join(tmp, p);
      await mkdir(path.dirname(dest), { recursive: true });
      await writeFile(dest, buf);
    }
    const target = path.join(ROOT, 'vendor/cubing', id);
    await mkdir(target, { recursive: true });
    await cp(tmp, target, { recursive: true });
  } finally { await rm(tmp, { recursive: true, force: true }); }

  await writeFile(currentFile, id + '\n');
  const url = `${OUT_URL_PREFIX}${id}/${ENTRY_PATH}`;
  await writeFile(htmlPath, html.replace(lineRe, (_, a, q, z) => `${a}'${url}'${z}`));
  console.log(`更新しました: ${url}`);
  setOutput(true, id);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  main().catch((e) => { console.error(`エラー: ${e.message}`); process.exit(1); });
}
