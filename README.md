#  文件元数据自动聚类系统

本项目用于对大量文件（如 Word、Excel、PDF、RTF 等）进行元数据提取、智能分组与结构化导出，可用于查重、文档归档、版本分析等办公自动化任务。

---

##  核心功能

-  **智能元数据提取**：提取作者、创建时间、文件类型等关键信息
-  **增量缓存机制**：自动识别已处理文件，避免重复提取
-  **多线程提取处理**：支持上万级文件快速处理
-  **两级智能聚类**：
  - 一级：按创建人 + 创建时间聚类（哈希分组）
  - 二级：基于文件名相似度进一步细分（考虑年份/版本/关键词）
-  **冲突文件日志输出**：记录因文件名差异被强制拆分的文件
-  **Excel 分组结果导出**：每个分组一个 Sheet，便于人工审查与归档

---

##  安装依赖

请使用 Python 3.8+，并安装以下依赖：

```bash
pip install -r requirements.txt
```

示例 `requirements.txt`：

```
et_xmlfile==2.0.0
lxml==5.4.0
numpy==2.2.5
olefile==0.47
openpyxl==3.1.5
pandas==2.2.3
pillow==11.2.1
PyPDF2==3.0.1
python-dateutil==2.9.0.post0
python-docx==1.1.2
python-pptx==1.0.2
pytz==2025.2
pywin32==310
six==1.17.0
typing_extensions==4.13.2
tzdata==2025.2
XlsxWriter==3.2.3
rapidfuzz==3.13.0
```

---

##  使用方法

### 1. 配置参数

编辑 `config.py`：

```python
FOLDER_PATH = r"D:\data field\测试文件"  # 目标文件夹路径
MAX_WORKERS = 12                        # 并行线程数
SIMILARITY_THRESHOLD = 90              # 文件名相似度阈值（90 ~ 95）
```

### 2. 运行主程序

```bash
python main.py
```

---

## 📁 输出文件说明

| 文件名                  | 说明                                  |
|-------------------------|----------------------------------------|
| `all_metadata_output.json` | 全部文件的元数据集合                    |
| `分组结果.xlsx`         | 每组文件一个 Sheet 的分组结果           |
| `文件名拆组日志.json`     | 文件名冲突拆组的日志，标明拆分原因       |
| `file_metadata.db`      | SQLite 数据库存储，支持增量更新         |

---

## 📂 项目结构示意

```
metadata_project/
├── main.py                      # 主程序入口
├── config.py                    # 配置文件
├── extractor.py                 # 多线程元数据提取模块
├── grouping.py                  # 分组逻辑与冲突判断
├── exporter.py                  # Excel 导出功能
├── metadata_db.py               # 元数据数据库缓存
├── utils.py                     # 工具函数（权限、时间等）
├── all_metadata_output.json
├── 分组结果.xlsx
└── 文件名拆组日志.json
```

---

##  可扩展建议

-  自动按组生成本地文件夹
-  支持图形界面（PyQt 或 Web 前端）
-  自动邮件发送分组结果
-  支持内容相似度判断（如嵌入式语义聚类）

---

##  作者建议

- 将 `MAX_WORKERS` 设置为 CPU 核心数 × 2，可大幅提速
- `SIMILARITY_THRESHOLD` 越高 → 越容易将相似文件分为不同组（建议 85~95）
- 分组依据中的“冲突关键词”与“年份/版本提取”可在 `grouping.py` 中扩展自定义

---

##  联系与支持

如需定制功能、部署脚本或技术支持，可联系开发者。
