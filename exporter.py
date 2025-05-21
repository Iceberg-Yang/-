# exporter.py
import pandas as pd
from collections import defaultdict
import re

def clean_excel_value(val):
    if isinstance(val, str):
        # 去除 openpyxl 不支持的非法字符（控制字符、非打印字符等）
        return re.sub(r"[\x00-\x1F]", "", val)
    return val
def export_metadata_to_excel(metadata_list, output_path="metadata_output.xlsx"):
    """
    将所有元数据导出为一个 Excel 文件（单 Sheet）
    """
    df = pd.DataFrame(metadata_list)
    columns_order = [
        "文件路径", "文件类型", "文件所有者",
        "作者", "最后修改人", "文档创建时间", "文档修改时间",
        "文件系统创建时间", "文件系统修改时间",
        "文件哈希值", "原始文件组ID", "预览内容"
    ]
    df = df[[col for col in columns_order if col in df.columns]]
    df.to_excel(output_path, index=False)


def export_grouped_metadata_to_excel(metadata_list, output_path="分组结果.xlsx"):
    """
    按 group_id 将每组文件分别导出为 Excel 多 Sheet 文件
    """
    grouped = defaultdict(list)
    for item in metadata_list:
        group_id = item.get("原始文件组ID", "未分组")
        grouped[group_id].append(item)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for group_id, items in grouped.items():
            df = pd.DataFrame(items)
            df = df.applymap(clean_excel_value)  #  清洗整个 DataFrame 的值
            df = df[[col for col in df.columns if col != "原始文件组ID"]]  # 避免冗余
            df.to_excel(writer, sheet_name=group_id[:31], index=False)  # Sheet 名不能超过 31 字符


    print(f" 已导出分组结果到：{output_path}")
