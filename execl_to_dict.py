import ast  # 用于将字符串解析为 Python 对象
import pandas as pd  # 用于处理 Excel 文件


def read_excel_and_add_query(file_path):
    # 读取 Excel 文件到 DataFrame
    df = pd.read_excel(file_path)
    target_ids = [144, 145, 146]  # 目标 ID 列表，用于筛选数据
    result_dict = {}  # 存储最终结果

    for target_id in target_ids:
        # 筛选出 id 列等于当前目标 ID 的行
        rows = df[df['id'] == target_id]
        if not rows.empty:
            row_data = {}
            # 从第二列开始遍历各列
            for col in rows.columns[1:]:
                value = rows[col].values[0]
                if isinstance(value, str):
                    try:
                        # 将字符串解析为 Python 对象
                        items = ast.literal_eval(value)
                        new_items = []
                        for item in items:
                            # 剔除 examples 字段，添加 query 字段
                            new_item = {
                                'hot': item.get('hot', ''),
                                'ids': item.get('ids', []),
                                'query': []
                            }
                            new_items.append(new_item)
                        row_data[col] = new_items
                    except SyntaxError:
                        row_data[col] = value
                else:
                    row_data[col] = value
            result_dict[target_id] = row_data
    return result_dict


# 请替换为实际 Excel 文件路径
file_path = r"C:\Users\LENOVO\Desktop\副本38a399995fd46b7b8e252a3ea1ed58b6_c2a5b6652f224ed90530ffec074c1896_8.xls"
data_dict = read_excel_and_add_query(file_path)
print(data_dict)