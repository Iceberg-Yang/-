# group
import hashlib
import os
import re
from collections import defaultdict
from rapidfuzz import fuzz

split_log = []


def has_conflicting_keywords(name1, name2):
    """
    判断两个文件名中是否包含语义上冲突的关键词
    例如：年初 vs 年末、draft vs final、v1 vs v2 等
    """
    conflict_pairs = [
        ("初", "末"),
        ("开始", "结束"),
        ("draft", "final"),
        ("初稿", "终稿"),
        ("v1", "v2"), ("v2", "v3"), ("v3", "v4"),
        #("1", "2"), ("2", "3"), ("3", "4"),
        ("旧", "新"),
        ("版一", "版二"),
    ]

    for kw1, kw2 in conflict_pairs:
        if (kw1 in name1 and kw2 in name2) or (kw2 in name1 and kw1 in name2):
            return True
    return False

def extract_years_and_versions(name):
    """
    提取文件名中的年份（如2023）和版本信息（如v1、v2、final）
    返回值：set，包含所有识别出的关键词
    """
    name = name.lower()
    parts = set()

    # 提取年份（范围大致设为2000~2099）
    parts.update(re.findall(r"20\d{2}", name))

    # 提取版本号：v1、v2、ver1、ver2、v01 等
    #parts.update(re.findall(r"\bv\d{1,2}\b", name))
    #parts.update(re.findall(r"\bver\d{1,2}\b", name))

    # 提取常见词：final、draft、rev、修订等
    for kw in ["final", "draft", "rev", "修订", "终稿", "初稿"]:
        if kw in name:
            parts.add(kw)

    # 提取纯数字尾部标号，如 1、2、3（避免误判年份）
    #trailing_num = re.findall(r"(\d{1,2})(?!\d)", name)
    #if trailing_num:
        #parts.add(f"num_{trailing_num[-1]}")

    return parts

def normalize_filename(name):
    """
    标准化文件名：
    - 去除 (2)、(3) 等副本编号
    - 保留日期/年份信息用于分组（如 2022.10）
    """
    name = name.lower()
    # 去除括号副本编号，如 (2)、(3)
    cleaned = re.sub(r"\s*\(\d+\)", "", name)

    # 提取常见日期/版本关键词
    dates = re.findall(r"20\d{2}[.\-]?\d{0,2}", name)
    if dates:
        cleaned += " " + " ".join(dates)

    return cleaned.strip()



def is_semantically_conflicting(name1, name2):
    """
    判断两个文件名是否语义冲突（关键词 + 年份/版本冲突），并记录日志
    """
    norm1 = normalize_filename(name1)
    norm2 = normalize_filename(name2)
    reason = ""
    if has_conflicting_keywords(norm1, norm2):
        reason = "命中冲突关键词"
    else:
        parts1 = extract_years_and_versions(norm1)
        parts2 = extract_years_and_versions(norm2)
        if parts1 and parts2 and parts1 != parts2:
            reason = f"版本/年份不一致 → {parts1} ≠ {parts2}"

    if reason:
        split_log.append({
            "文件1": name1,
            "文件2": name2,
            "拆组原因": reason
        })
        return True
    return False



def generate_metadata_hash(item):
    """
    生成基于元数据的哈希值（用于快速分组）
    """
    key = f"{item.get('作者', '')}|{item.get('文档创建时间', '')}|{item.get('文件所有者', '')}|{item.get('文件类型', '')}"
    return hashlib.md5(key.encode('utf-8')).hexdigest()

def split_group_by_filename_difference(group_items, threshold=90):
    #（95+ | 只保留非常非常像的为一组|90| 推荐值（可区分 2023 vs 2024）|85|比较宽容，可能合并不同年份   
    """
    将一个哈希组内的文件根据文件名相似度再次细分
    """
    subgroups = []
    used = set()

    for i, item in enumerate(group_items):
        if i in used:
            continue
        name_i = normalize_filename(os.path.basename(item["文件路径"]))
        current_subgroup = [item]
        used.add(i)

        for j in range(i + 1, len(group_items)):
            if j in used:
                continue
            name_j = normalize_filename(os.path.basename(group_items[j]["文件路径"]))
            similarity = fuzz.ratio(name_i, name_j)
            #  冲突判断使用原始文件名
            original_name_i = os.path.basename(item["文件路径"])
            original_name_j = os.path.basename(group_items[j]["文件路径"])
            if similarity >= threshold and not is_semantically_conflicting(original_name_i, original_name_j):
                current_subgroup.append(group_items[j])
                used.add(j)
        subgroups.append(current_subgroup)

    return subgroups


def group_by_metadata_hash_with_filename_split(metadata_list, threshold=90):
    """
    使用元数据哈希进行快速分组后，再按文件名相异度进行二次拆分
    """
    # 第一步：按哈希初步分组
    groups = defaultdict(list)
    for item in metadata_list:
        hash_key = generate_metadata_hash(item)
        groups[hash_key].append(item)

    # 第二步：组内按文件名再拆分
    result = []
    group_counter = 1
    for group_items in groups.values():
        subgroups = split_group_by_filename_difference(group_items, threshold)
        for sub in subgroups:
            for item in sub:
                item["原始文件组ID"] = f"Group_{group_counter}"
                result.append(item)
            group_counter += 1

    return result
