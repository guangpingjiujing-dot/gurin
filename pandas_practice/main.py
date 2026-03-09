"""
Excelからデータをpandasで読み込み、データベースに投入するシンプルなスクリプト

Python標準のDB-API 2.0パターンを使用。psycopg2（PostgreSQL）の例ですが、
sqlite3、pymysql など他のDBドライバでも同じ connect → cursor → execute の流れで応用できます。

使用するExcel: openpyxl_practice の topic9_iterate_data() で作成される topic9_iterate.xlsx
  - カラム: ID, 名前, 年齢, 部署
"""

import os
from pathlib import Path

import pandas as pd
import psycopg2
from dotenv import load_dotenv

# 環境変数を読み込み
load_dotenv()

# openpyxl_practice で作成したExcelファイルのパス
EXCEL_PATH = Path(__file__).parent.parent / "openpyxl_practice" / "excel_files" / "topic9_iterate.xlsx"


def load_excel_to_dataframe() -> pd.DataFrame:
    """ExcelファイルをpandasのDataFrameに読み込む"""
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(
            f"Excelファイルが見つかりません: {EXCEL_PATH}\n"
            "先に openpyxl_practice/main.py の topic9_iterate_data() を実行してExcelを生成してください。"
        )

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    return df


def prepare_for_database(df: pd.DataFrame) -> pd.DataFrame:
    """DB用にカラム名を英語に変換"""
    df = df.rename(columns={
        "ID": "id",
        "名前": "name",
        "年齢": "age",
        "部署": "department",
    })
    return df


def insert_to_database(df: pd.DataFrame) -> None:
    """DataFrameのデータをデータベースに投入（DB-API 2.0パターン）"""
    database_url = os.environ.get("DATABASE_URL")

    if not database_url:
        raise ValueError(
            "環境変数 DATABASE_URL を設定してください。\n"
            ".env.example を参考に .env ファイルを作成してください。"
        )

    conn = psycopg2.connect(database_url)
    try:
        cur = conn.cursor()
        # プレースホルダは %s（psycopg2の形式）。SQLiteなら ?、MySQLなら %s
        cur.executemany(
            "INSERT INTO employees (id, name, age, department) VALUES (%s, %s, %s, %s)",
            [tuple(row) for row in df.to_numpy()],
        )
        conn.commit()
        print(f"{len(df)} 件のデータをデータベースに投入しました。")
    finally:
        cur.close()
        conn.close()


def main():
    print("Excel → pandas → データベース の流れでデータを投入します")
    print("-" * 50)

    # 1. Excelを読み込み
    df = load_excel_to_dataframe()
    print(f"Excelから {len(df)} 件のデータを読み込みました:")
    print(df)
    print()

    # 2. DB用にカラム名を変換
    df = prepare_for_database(df)
    print("DB用に変換したデータ:")
    print(df)
    print()

    # 3. データベースに投入
    insert_to_database(df)
    print("完了！")


if __name__ == "__main__":
    main()
