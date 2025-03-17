import pymysql
import json
import math
from execl_to_dict import read_excel_and_add_query

# 配置数据库连接信息
db_config = {
    'host': '',
    'user': '',
    'password': '',
    'database': '',
    'charset': 'utf8'
}
# 建立数据库连接
conn = pymysql.connect(**db_config)
cursor = conn.cursor()

# file_path = r"C:\Users\LENOVO\Desktop\副本38a399995fd46b7b8e252a3ea1ed58b6_c2a5b6652f224ed90530ffec074c1896_8.xls"
file_path = r"C:\Users\LENOVO\Desktop\2.xls"
data = read_excel_and_add_query(file_path)
# 将键转换为普通整数类型
data = {int(key): value for key, value in data.items()}

# 遍历数据
for key, value in data.items():
    g____id = 0
    g____re = 0
    w____id = 0
    w____re = 0
    for category, items in value.items():
        if isinstance(items, float) and math.isnan(items):  # 解析excel文件时遇到空的单元格，会被设为 nan
            print(f"{category} 的值为 nan，进行相应处理（比如跳过）")
            continue  # 跳过本次循环，不进行后续操作
        elif isinstance(items, list):
            for item in items:
                ids = item['ids']
                print(f"当前查询的 IDs: {ids}")
                if category in ['native_gc', 'native_hj', 'native_kf', 'native_zx']:
                    # id数字位数较少的的ID查询
                    digit_ids = [str(id) for id in ids]
                    if digit_ids:
                        sql = f"SELECT chat_id, title FROM wechat_summary WHERE chat_id IN ({','.join(digit_ids)})"
                        try:
                            cursor.execute(sql)
                            result = cursor.fetchall()    # 结果为元组列表，每个元素为 (chat_id, title)
                            w____re += len(result)  # 统计微信数据的数量
                            print(f"查询结果：{result}")
                            if result:
                                item['query']= [{row[0]: row[1]} for row in result]    # 将查询结果的两列数据，直接设为键值对
                            else:
                                item['query']= []
                            item['ids'] = ', '.join(map(str, ids))
                            w____id += len(ids)      # 验证微信数据的数量
                        except pymysql.Error as e:
                            print(f"执行 SQL 语句时出错: {e}")
                else:
                    # id数字位数较多的ID查询
                    digit_ids = [str(id) for id in ids]
                    if digit_ids:
                        sql = f"SELECT id, content FROM ticket WHERE id IN ({','.join(digit_ids)})"
                        try:
                            cursor.execute(sql)
                            result = cursor.fetchall()
                            g____re += len(result)  # 统计工单数据的数量
                            print(f"查询结果：{result}")
                            if result:
                                item['query'] = [{row[0]: row[1]} for row in result]
                            else:
                                item['query'] = []
                            item['ids'] = ', '.join(map(str, ids))
                            g____id += len(ids)  # 验证工单数据的数量
                        except pymysql.Error as e:
                            print(f"执行 SQL 语句时出错: {e}")

    # 输出每个id对应下的工单数据和微信数据的数量
    print(key, "的查询结果：g____, w____",g____re, g____id, w____re, w____id)

# 关闭数据库连接
cursor.close()
conn.close()

# 输出结果
print("输出结果：-----------------------")
print(json.dumps(data, ensure_ascii=False, indent=4))   # 参数 ensure_ascii=False 避免中文字符被转义    indent=4 输出格式化 JSON

# 写入output2.json文件
try:
    with open('output2.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print("数据已成功写入 output2.json 文件。")
except Exception as e:
    print(f"写入文件时出现错误: {e}")



