# sheetflow

`sheetflow` 是一组面向工作表驱动工作流的办公自动化工具。分发包名是
`sheetflow`，Python 导入包名是 `sheetflow`。项目底层目前使用 xlwings，
但 xlwings 只是 Excel 适配层，不是项目的核心抽象。

项目通过 Excel 工作表的命名区域接收数据，查找内置 Word/Excel 模板，生成业务文档，适用于工商登记、税务申报、个税处理和数据分析等场景。

## 项目结构

| 路径 | 作用 |
| --- | --- |
| `src/sheetflow/` | 核心 Python 包 |
| `src/sheetflow/Template/Word/` | 内置 Word 模板 |
| `src/sheetflow/Template/Excel/` | 内置 Excel 模板 |
| `py_script/` | Excel xlwings `Run main` 入口脚本 |
| `tests/` | pytest 单元测试 |

核心模块包括：

- `named_ranges.py`：读取和创建工作表级命名区域。
- `pipeline.py`：将命名区域、模板搜索和 Word 渲染串成通用流程。
- `records.py`：把列式数据转换成逐行记录。
- `render.py`：使用 `docxtpl` 批量渲染 Word 模板。
- `time_period.py`：按月、季度或年度切分日期范围。
- `personal_income_tax.py`：批量填充 2018/2019 个税 Excel 模板。
- `autozip.py`：压缩生成的 Excel 文件并按年份整理。
- `user_errors.py`：统一维护用户错误码并生成 ASCII-safe 错误文本。
- `common.py` 和 `smart_path_manager.py`：入口共用流程和输出目录管理。

## 安装

在项目目录执行：

```bash
uv sync --extra dev
```

构建 wheel：

```bash
uv build
```

安装构建产物：

```bash
uv pip install dist/sheetflow-0.1.0-py3-none-any.whl
```

开发时也可以使用可编辑安装：

```bash
uv pip install -e .
```

## Excel 入口

`py_script/` 中的脚本由 Excel 的 xlwings `Run main` 调用。当前 Python 环境需要已经安装本包，并且本机需要有可操作的 Excel/Office 环境。首次使用前还需安装 Excel 加载项：

```bash
uv run xlwings addin install
```

加载项需与 Python 包版本保持一致，具体配置参见 [xlwings Add-in & Settings](https://docs.xlwings.org/en/stable/addin.html)。

| 脚本 | 用途 | 必需命名区域 |
| --- | --- | --- |
| `py_script/namedrange.py` | 批量创建命名区域 | `Named`、`Range`、`Sheet_Name` |
| `py_script/工商.py` | 生成工商登记 Word 文档 | `Template` 和模板字段 |
| `py_script/税务.py` | 生成税务 Word 文档或个税压缩包 | `Template`、`start`、`end`、`freq` |
| `py_script/数据分析.py` | 填充 Excel 模板 | `Template`，可选 `Sheet` |

`NamedRangeDict` 只读取当前工作表的 `sheet.names`，不会混入工作簿级全局名称。
工商登记的列式数据只跳过所有字段均为空的占位行；可选字段为空的记录仍会生成。
入口会尽量在创建输出目录前完成配置、数据形状和模板校验。目录达到阈值而需要转移
到桌面时，Windows 会使用用户重定向后的实际桌面；桌面无法解析时继续使用工作簿目录。

## 模板规则

模板默认从 `src/sheetflow/Template/` 递归搜索。Excel 中的 `Template` 命名区域填写不带扩展名和路径分隔符的纯模板名，例如 `企业所得税`；调用方会根据业务流程搜索 `.docx`、`.xlsx` 或 `.xls` 文件。

模板搜索缓存只保存仍存在的成功结果；未找到的模板不会缓存，运行中新增或恢复模板后
可在下一次调用被发现。

构建 wheel 时，`src/sheetflow/Template/**/*` 会随包一起发布。

## 错误信息

Excel/VBA 链路可能无法可靠显示中文异常。因此 Run main 入口使用带 `ERR_*` 错误码的 ASCII-safe 错误信息，并在 `CN_ESCAPED:` 后附带转义后的中文说明。

生产错误码包括：

| 错误码 | 含义 |
| --- | --- |
| `ERR_MISSING_NAMED_RANGE` | 缺少必需命名区域 |
| `ERR_NAMED_RANGE_MAP` | Named/Range 长度不一致或名称重复 |
| `ERR_NAMED_RANGE_PARTIAL` | 命名区域批量创建部分失败，工作簿未保存 |
| `ERR_NAMED_RANGE_RESTORE_FAILED` | 命名区域覆盖失败且旧定义恢复失败 |
| `ERR_TEMPLATE_NOT_FOUND` | 找不到模板 |
| `ERR_TEMPLATE_NAME` | 模板名包含路径或不是有效的纯文件名 |
| `ERR_TEMPLATE_TYPE` | 模板文件类型不受支持 |
| `ERR_TEMPLATE_WRITE_FAILED` | Excel 模板字段写入失败 |
| `ERR_FILE_NOT_FOUND` | 指定文件不存在 |
| `ERR_NOT_A_FILE` | 指定路径不是文件 |
| `ERR_SEARCH_PATH` | 模板搜索路径无效 |
| `ERR_INVALID_PATH` | 指定目录路径无效 |
| `ERR_OUTPUT_DIR` | 无法创建输出目录 |
| `ERR_OUTPUT_PATH` | 生成文件将逃出输出目录或本批次输出路径重复 |
| `ERR_NO_NAMED_RANGES` | 当前工作表没有命名区域 |
| `ERR_MIXED_RECORD_SHAPE` | 同时混用了列表字段和标量字段 |
| `ERR_RECORD_DATA_TYPE` | 记录集输入不是映射 |
| `ERR_RECORD_KEY_TYPE` | 记录集列名不是字符串 |
| `ERR_RECORD_COLUMN_TYPE` | 记录集列值不是列表 |
| `ERR_RECORD_LENGTH_MISMATCH` | 记录集各列长度不一致 |
| `ERR_PERMISSION_DENIED` | 无权限读取模板 |
| `ERR_RUN_MAIN_FAILED` | Run main 捕获到未格式化异常 |
| `ERR_ARCHIVE_PARTIAL` | 批量压缩失败或多个源文件对应同一压缩包 |

## 测试与构建检查

```bash
uv run pytest
uv run python -m compileall py_script src/sheetflow
uv build
```

修改入口脚本、错误格式、模板搜索或包结构后，建议运行 pytest、compileall 和 uv build。
