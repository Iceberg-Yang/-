#extractor.py
import os
import re
from openpyxl import load_workbook
from docx import Document
from pptx import Presentation
from PyPDF2 import PdfReader
import olefile
from metadata_db import init_db, upsert_metadata, get_cached_metadata
from utils import get_file_owner, get_file_times, convert_ole_time
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging

logging.basicConfig(
    level=logging.DEBUG,  # 日志等级，打印所有 DEBUG 及以上等级日志
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='metadata_extraction.log',  # 日志文件路径，程序目录下
    filemode='w'  # 'w' 表示每次运行清空日志文件，'a' 表示追加
)

def is_windows():
    return platform.system().lower() == 'windows'

def is_linux():
    return platform.system().lower() == 'linux'



def extract_excel_metadata(file_path):
    try:
        wb = load_workbook(file_path, read_only=True)
        props = wb.properties
        return props.creator, props.lastModifiedBy, props.created, props.modified
    except:
        return "", "", "", ""

def extract_word_metadata(file_path):
    try:
        doc = Document(file_path)
        props = doc.core_properties
        return props.author, props.last_modified_by, props.created, props.modified
    except:
        return "", "", "", ""

def extract_ppt_metadata(file_path):
    try:
        ppt = Presentation(file_path)
        props = ppt.core_properties
        return props.author, props.last_modified_by, props.created, props.modified
    except:
        return "", "", "", ""

def extract_pdf_metadata(file_path):
    try:
        reader = PdfReader(file_path)
        info = reader.metadata
        return info.get("/Author", ""), "", info.get("/CreationDate", ""), info.get("/ModDate", "")
    except:
        return "", "", "", ""
    

def extract_rtf_metadata(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        # 提取作者
        author_match = re.search(r"{\\author ([^}]+)}", content)
        # 提取创建时间
        creatim_match = re.search(r"{\\creatim[^}]*\\yr(\d+)\\mo(\d+)\\dy(\d+)\\hr(\d+)\\min(\d+)", content)
        # 提取修改时间
        revtim_match = re.search(r"{\\revtim[^}]*\\yr(\d+)\\mo(\d+)\\dy(\d+)\\hr(\d+)\\min(\d+)", content)
        author = author_match.group(1) if author_match else ""
        def parse_time(match):
            if match:
                return f"{match.group(1)}-{match.group(2).zfill(2)}-{match.group(3).zfill(2)} {match.group(4).zfill(2)}:{match.group(5).zfill(2)}"
            return ""

        created = parse_time(creatim_match)
        modified = parse_time(revtim_match)

        return author, "", created, modified,

    except Exception as e:
        return "", "", "", ""


def extract_ole_metadata(file_path):
    try:
        if not olefile.isOleFile(file_path):
            return "", "", "", ""
        ole = olefile.OleFileIO(file_path)
        meta = ole.getproperties('\x05SummaryInformation')
        author = meta.get(4, "")
        last_modified_by = meta.get(8, "")
        created = convert_ole_time(meta.get(12))
        modified = convert_ole_time(meta.get(13))
        return author, last_modified_by, created, modified
    except:
        return "", "", "", ""

def extract_metadata_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    file_type = ""
    if ext == ".docx":
        file_type = "Word (docx)"
        author, last_modified_by, created, modified = extract_word_metadata(file_path)
    elif ext == ".xlsx":
        file_type = "Excel (xlsx)"
        author, last_modified_by, created, modified = extract_excel_metadata(file_path)
    elif ext == ".pptx":
        file_type = "PowerPoint (pptx)"
        author, last_modified_by, created, modified = extract_ppt_metadata(file_path)
    elif ext == ".pdf":
        file_type = "PDF"
        author, last_modified_by, created, modified = extract_pdf_metadata(file_path)
    elif ext in [".doc", ".xls", ".ppt"]:
        file_type = "Office旧格式"
        author, last_modified_by, created, modified = extract_ole_metadata(file_path)
    elif ext == ".rtf":
        file_type = "RTF 文档"
        author, last_modified_by, created, modified = extract_rtf_metadata(file_path)
    else:
        logging.info(f"跳过不支持的文件类型: {file_path} (扩展名: {ext})")
        return None


    file_owner = get_file_owner(file_path)
    file_times = get_file_times(file_path)
    fs_created_time = file_times["created"]
    fs_modified_time = file_times["modified"]

    logging.info(f"成功提取元数据: {file_path}")
    logging.info(f"  类型: {file_type}")
    logging.info(f"  作者: {author}")
    logging.info(f"  最后修改人: {last_modified_by}")
    logging.info(f"  文档创建时间: {created}")
    logging.info(f"  文档修改时间: {modified}")
    logging.info(f"  文件所有者: {file_owner}")
    logging.info(f"  文件系统创建时间: {fs_created_time}")
    logging.info(f"  文件系统修改时间: {fs_modified_time}")


    return {
        "文件路径": file_path,
        "文件类型": file_type,
        "文件所有者": file_owner,
        "作者": author,
        "最后修改人": last_modified_by,
        "文档创建时间": str(created),
        "文档修改时间": str(modified),
        "文件系统创建时间": str(fs_created_time),
        "文件系统修改时间": str(fs_modified_time),
        #"预览内容": preview if ext in [".txt", ".rtf"] else ""
    }


from concurrent.futures import ThreadPoolExecutor, as_completed

def scan_directory_parallel_with_cache(directory, max_workers=8):
    init_db()
    file_paths = []
    for root, _, files in os.walk(directory):
        for f in files:
            file_paths.append(os.path.join(root, f))
        logging.info(f"扫描到 {len(file_paths)} 个文件，开始提取元数据")


    results = []

    def process_file(path):
        try:
            modified = str(get_file_times(path)["modified"])
            cached = get_cached_metadata(path, modified)
            if cached:
                return cached
            meta = extract_metadata_from_file(path)
            if meta:
                upsert_metadata(path, modified, meta)
            return meta
        except Exception as e:
            logging.error(f"处理文件失败: {path}, 错误: {e}")
            return None

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_file, path) for path in file_paths]
        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    return results

