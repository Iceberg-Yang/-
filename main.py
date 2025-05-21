#main
from extractor import scan_directory_parallel_with_cache
from exporter import export_grouped_metadata_to_excel
from grouping import group_by_metadata_hash_with_filename_split
from collections import defaultdict
import json
from grouping import split_log

def main():
    folder = r"E:\你的路径"  # 替换为你的目录路径
    print(" 正在扫描文件夹并提取元数据...")
    metadata_list = scan_directory_parallel_with_cache(folder, max_workers=16) 
    # 线程数可调 max_workers=12：线程数建议设置为 CPU核心数 * 2 左右。可根据你的机器调整，比如 16 核可设 32。


    print(" 正在进行原始文件分组（根据作者 + 创建时间）...")
    metadata_list = group_by_metadata_hash_with_filename_split(metadata_list, threshold=90)


    #  按组输出：哪些文件属于同一源文件
    grouped = defaultdict(list)
    for item in metadata_list:
        group_id = item.get("原始文件组ID", "未分组")
        grouped[group_id].append(item)

    print("\n 同一源文件判定结果：\n")
    for group_id, files in grouped.items():
        print(f" 原始文件组 {group_id}（共 {len(files)} 个文件） → 判定为同一源文件：")
        for file in files:
            print(f" - {file['文件路径']}")
        print("-" * 80)

    print(f"\n 共扫描文件数: {len(metadata_list)}")
    print(f" 识别出原始文件组数: {len(grouped)}\n")
    
    with open("all_metadata_output.json", "w", encoding="utf-8") as f:
        json.dump(metadata_list, f, ensure_ascii=False, indent=2)
        
    # 打印冲突日志
    if split_log:
        print("\n 以下文件在文件名细分时被拆组，原因如下：\n")
        for entry in split_log:
            print(f"[拆组] {entry['文件1']} vs {entry['文件2']}")
            print(f"      → 冲突原因：{entry['拆组原因']}\n")

    
    # 可选：保存为 JSON 文件
        with open("文件名拆组日志.json", "w", encoding="utf-8") as f:
            json.dump(split_log, f, ensure_ascii=False, indent=2)
        print(" 已将冲突日志写入：文件名拆组日志.json")
    else:
        print(" 未发现任何冲突性文件名拆分。")
    
        #导出为excel
    print("正在导出 Excel 文件...")
    export_grouped_metadata_to_excel(metadata_list, output_path="原始文件分组结果.xlsx")
    
    
if __name__ == "__main__":
    main()





    #print(" 输出提取结果（便于调试）：\n")

    #for i, item in enumerate(metadata_list, 1):
        #print(f" 文件 {i}:")
        #pprint.pprint(item, sort_dicts=False)
        #print("-" * 80)

    #print(f"\n 总共处理文件数: {len(metadata_list)}")

    #print("完成！已识别出原始文件分组并导出结果。")
    #