-- 従業員テーブル（topic9_iterate.xlsx のデータ用）
-- openpyxl_practice の topic9_iterate_data() で作成されるExcelの構造に合わせています

CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    age INTEGER NOT NULL,
    department TEXT NOT NULL
);
