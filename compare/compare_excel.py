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

    # 读取Excel文件
    print("正在读取汇总Excel...")
    df_huizong = pd.read_excel(huizong_file)

    print("正在读取入账Excel...")
    df_ruzhang = pd.read_excel(ruzhang_file)

    # 获取A-F列（前6列）
    columns_to_compare = df_huizong.columns[:6].tolist()
    print(f"比较的列: {columns_to_compare}")

    # 提取A-F列的数据
    huizong_data = df_huizong[columns_to_compare].copy()
    ruzhang_data = df_ruzhang[columns_to_compare].copy()

    # 统一数据类型为字符串，避免类型不匹配问题
    for col in columns_to_compare:
        huizong_data[col] = huizong_data[col].astype(str)
        ruzhang_data[col] = ruzhang_data[col].astype(str)

    # 使用merge找出差异
    merged = huizong_data.merge(
        ruzhang_data,
        on=columns_to_compare,
        how='left',
        indicator=True
    )

    # 找出在汇总中存在但在入账中不存在的行
    diff_rows = merged[merged['_merge'] == 'left_only']

    if len(diff_rows) == 0:
        print("\n没有发现差异，入账Excel已包含汇总Excel的所有数据。")
        return

    # 获取原始汇总数据中对应的完整行（包括所有列）
    diff_indices = diff_rows.index
    rows_to_add = df_huizong.iloc[diff_indices].copy()

    print(f"\n发现 {len(rows_to_add)} 行新数据需要添加到入账Excel")
    print("\n新增的数据预览:")
    print(rows_to_add[columns_to_compare].to_string(index=False))

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
