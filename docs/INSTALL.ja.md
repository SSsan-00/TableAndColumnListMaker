# 導入手順

## GitHub から利用する場合

1. GitHub のリポジトリページを開く。
2. `Code` からソース一式をダウンロードする。
3. Excel を開く。
4. `Alt + F11` で VBA エディタを開く。
5. `ファイル`、`ファイルのインポート` から `src/TableAndColumnListMaker.bas` をインポートする。
6. Excel に戻り、マクロ `RunTableAndColumnListMaker` を実行する。

## 単一ファイル Bootstrap で利用する場合

ソース一式をダウンロードできないユーザー向けに、`bootstrap/InstallTableAndColumnListMaker.vbs` だけで導入を開始できます。

1. `InstallTableAndColumnListMaker.vbs` を保存する。
2. Excel の設定で `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` を有効にする。
3. `InstallTableAndColumnListMaker.vbs` を実行する。
4. デスクトップに作成された `TableAndColumnListMaker.xlsm` を開く。
5. マクロを有効化する。
6. マクロ `RunTableAndColumnListMaker` を実行する。

Bootstrap は GitHub 上の最新 `src/TableAndColumnListMaker.bas` を取得して、マクロ有効ブックへ組み込みます。

## 実行時の流れ

1. `RunTableAndColumnListMaker` を実行する。
2. フォルダ選択ダイアログで解析対象の起点フォルダを選ぶ。
3. 対象ファイルが再帰的に探索される。
4. ファイル名が条件に一致したブックの先頭シートが解析される。
5. 実行ブック内に次の4シートが作成または更新される。
   - `テーブル一覧`
   - `カラム一覧`
   - `テーブル検索`
   - `カラム検索`

## 注意事項

- Excel のマクロ実行を許可する必要があります。
- 解析対象ブックは、可能な限り閉じた状態で実行してください。
- パスワード付きブック、破損ブック、開けないブックはスキップされます。
- Bootstrap は GitHub からファイルを取得するため、インターネット接続が必要です。
- Bootstrap で作成済みの `TableAndColumnListMaker.xlsm` がデスクトップにある場合は上書きされます。
