# pandas 練習: Excel → Supabase

openpyxl_practice で作成した Excel のデータを pandas で読み込み、Supabase に投入するシンプルなスクリプトです。

## 前提条件

- openpyxl_practice の `topic9_iterate_data()` を実行して `topic9_iterate.xlsx` を生成しておく
- Supabase のプロジェクトを作成済みであること

## セットアップ

1. 依存パッケージをインストール
   ```
   uv sync
   ```

2. `.env.example` を `.env` にコピーし、`DATABASE_URL`（接続文字列）を設定

3. データベースでテーブルを作成  
   Supabase の場合は SQL Editor で `sql/create_employees_table.sql` の内容を実行。  
   `DATABASE_URL` は Database > Connection string から取得

## 実行

```
uv run python main.py
```
