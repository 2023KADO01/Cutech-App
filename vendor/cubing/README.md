# vendor/cubing

スクランブル生成ライブラリ cubing.js の、固定コピーです。Cutech Log（`log/L_Index.html`）が読み込みます。

- 出典：https://cdn.cubing.net/v0/js/cubing/scramble とそこから読み込まれるファイル一式（cubing.js の v0 ビルド）
- 出典リポジトリ：https://github.com/cubing/cubing.js
- ライセンス：MPL-2.0 と GPL-3.0-or-later のデュアルライセンス
- `current.txt` に、いま使っている版ID（中身のSHA-256の先頭12文字）が入っています。`<版ID>/` が各版のコピーです。
- 配布元の v0 は中身が更新されるため、コピーを固定しています。

## 更新の仕組み

`.github/workflows/update-cubing.yml` が毎週、配布元を確認します。ビルドが変わっていれば、新しい版を `vendor/cubing/<版ID>/` に保存し、`CUBING_URL` を書き換えるPRを自動で作ります。
古い版のフォルダは、公開中の画面が参照している可能性があるので消しません。
