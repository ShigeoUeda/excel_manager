from excel_manager import ExcelManager

excel = ExcelManager(r"sample.xlsx")
# ファイルの格納場所は以下のようにWindowsのパスを指定でも可能
# excel = ExcelManager(r"/mnt/c/Users/hoge/Desktop/sample.xlsx")

# シートの作成
excel.create_sheet("データ", ["ID", "名前", "値"])

# データの書き込み
excel.write_cell("データ", 2, 1, 1)  # 行列指定
excel.write_cell("データ", 2, "B", "サンプル1")  # 列記号指定
excel.write_cell("データ", 2, 3, 1000, "#,##0")  # 数値書式指定

excel.write_cell_a1("データ", "A3", 2)  # A1形式指定
excel.write_cell_a1("データ", "B3", "サンプル2")
excel.write_cell_a1("データ", "C3", 2000, "#,##0")

# データの読み込み
value1 = excel.read_cell("データ", 2, 1)
value2 = excel.read_cell_a1("データ", "B2")

# 範囲データの読み込み
range_data = excel.read_range("データ", 2, "A", 3, "C")
print(f"読み込んだデータ: {range_data}")

excel.save()