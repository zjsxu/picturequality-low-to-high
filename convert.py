import json

def convert_list_to_dict(json_in, json_out):
    with open(json_in, 'r', encoding='utf-8') as f:
        data = json.load(f)

    result = {}
    for item in data:
        phash = item.get("phash")
        path = item.get("path")
        if phash and path:
            result[phash] = path

    with open(json_out, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"转换完成！输出保存为：{json_out}，共 {len(result)} 条记录。")

convert_list_to_dict(
    "/Users/zhangjingsen/Desktop/SIIS file/原图索引.json",         # 你现在用的那份索引
    "hash_index_fixed.json"  # 你之后要输入到 GUI 的那份
)