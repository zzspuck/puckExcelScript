import pandas as pd
import os

def compare_and_update_excel(huizong_file, ruzhang_file):
    """
    比较汇总Excel和入账Excel的A-F列内容
    将汇总中有但入账中没有的数据追加到入账Excel最下面

    参数:
        huizong_file: 汇总Excel文件路径
        ruzhang_file: 入账Excel文件路径
    """

    # 检查文件是否存在
    if not os.path.exists(huizong_file):
        print(f"错误: 找不到文件 {huizong_file}")
        return

    if not os.path.exists(ruzhang_file):
        print(f"错误: 找不到文件 {ruzhang_file}")
        return

    # 读取汇总Excel
    print("正在读取汇总Excel...")
    df_huizong = pd.read_excel(huizong_file)
    expected_columns = df_huizong.columns[:6].tolist()

    # 读取入账Excel - 自动找出包含正确列名的sheet
    print("正在读取入账Excel...")
    excel_file = pd.ExcelFile(ruzhang_file)
    df_ruzhang = None
    correct_sheet = None

    # 遍历所有sheet，找出包含正确列名的
    for sheet_name in excel_file.sheet_names:
        df_temp = pd.read_excel(ruzhang_file, sheet_name=sheet_name)
        # 检查前6列是否包含预期的列名
        if all(col in df_temp.columns for col in expected_columns):
            df_ruzhang = df_temp
            correct_sheet = sheet_name
            print(f"  找到正确的sheet: '{correct_sheet}'")
            break

    if df_ruzhang is None:
        print(f"错误: 在入账Excel中找不到包含正确列名的sheet")
        print(f"  预期列名: {expected_columns}")
        print(f"  可用sheets: {excel_file.sheet_names}")
        return

    # 显示两个文件的列名用于诊断
    print(f"汇总Excel的列名: {df_huizong.columns.tolist()}")
    print(f"入账Excel的列名: {df_ruzhang.columns.tolist()}")

    # 获取A-F列（前6列）
    columns_to_compare = df_huizong.columns[:6].tolist()
    print(f"\n汇总中比较的列: {columns_to_compare}")

    # 提取A-F列的数据
    huizong_data = df_huizong[columns_to_compare].copy()
    ruzhang_data = df_ruzhang[columns_to_compare].copy()

    # 统一数据格式，处理不同的数据类型
    def format_number(val):
        """将数字格式化为统一的字符串表示，避免5和5.0的不匹配"""
        try:
            num = float(val)
            if pd.isna(num):
                return 'nan'
            # 用'g'格式去掉尾部0，保留最多10位有效数字
            return f"{num:.10g}"
        except (ValueError, TypeError):
            return str(val)

    for col in columns_to_compare:
        # 处理日期列
        if col == '日期' or 'date' in col.lower():
            huizong_data[col] = pd.to_datetime(huizong_data[col], errors='coerce').dt.strftime('%Y-%m-%d')
            ruzhang_data[col] = pd.to_datetime(ruzhang_data[col], errors='coerce').dt.strftime('%Y-%m-%d')
        # 处理数字列（借方、贷方、余额）
        elif col in ['借方', '贷方', '余额']:
            huizong_data[col] = huizong_data[col].apply(format_number)
            ruzhang_data[col] = ruzhang_data[col].apply(format_number)
        # 其他列直接转为字符串
        else:
            huizong_data[col] = huizong_data[col].astype(str)

    # 创建行标识（将所有比较列合并成元组），用于精确比较
    huizong_rows = set(huizong_data.apply(tuple, axis=1))
    ruzhang_rows = set(ruzhang_data.apply(tuple, axis=1))

    # 找出在汇总中存在但在入账中不存在的行标识
    diff_row_tuples = huizong_rows - ruzhang_rows

    if len(diff_row_tuples) == 0:
        print("\n没有发现差异，入账Excel已包含汇总Excel的所有数据。")
        return

    # 找出原始汇总数据中对应的行（使用格式化后的数据来匹配）
    mask = huizong_data.apply(tuple, axis=1).isin(diff_row_tuples)
    rows_to_add = df_huizong[mask].copy()

    print(f"\n发现 {len(rows_to_add)} 行新数据需要添加到入账Excel")
    print("\n前5行新增数据预览:")
    try:
        preview = rows_to_add[columns_to_compare].head(5).to_string(index=False)
        print(preview)
    except UnicodeEncodeError:
        print(f"(包含特殊字符，显示省略，共 {len(rows_to_add)} 行)")

    # 将新数据追加到入账Excel
    df_updated = pd.concat([df_ruzhang, rows_to_add], ignore_index=True)

    # 保存更新后的入账Excel
    print(f"\n正在保存更新后的入账Excel...")
    df_updated.to_excel(ruzhang_file, index=False)

    print(f"完成! 已将 {len(rows_to_add)} 行数据追加到 {ruzhang_file}")


if __name__ == "__main__":
    # 设置文件路径
    huizong_file = "huizong.xlsx"  # 汇总Excel文件名
    ruzhang_file = "入账.xlsx"     # 入账Excel文件名

    print("=" * 60)
    print("Excel数据比较和更新工具")
    print("=" * 60)

    compare_and_update_excel(huizong_file, ruzhang_file)
