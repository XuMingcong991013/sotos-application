# Resume Converter 项目协作说明

## 1. 项目目标与当前状态

本项目为中国大陆公司的招聘部门开发简历标准化工具。当前已经实现完整主流程：

```text
原始简历
  -> 百度文档解析生成中间 Markdown
  -> 内网 LLM 恢复 Markdown 阅读结构
  -> 内网 LLM 提取结构化简历信息并由 Python 校验证据
  -> 按模板标签生成标准 Word 简历
  -> 生成待补充信息清单和全过程记录
```

当前模板标签只有 `SOTOS`，支持“有照片占位”和“无照片”两种模式。后续会增加其他模板；文档解析阶段对所有模板共用，信息提取和 Word 生成阶段允许按模板标签扩展。

不要把项目退回到“只生成 Markdown”的早期状态。除非用户明确提出，否则不要开发 JD 匹配、候选人排序、内容润色、简历改写或 Web 层。

## 2. 部署环境、网络与密钥

- 项目部署在中国大陆公司网络中，不能依赖境外服务或境外网络。
- 文档解析使用百度智能云“文档解析”API。
- LLM 使用公司内网的 OpenAI 兼容服务。
- 默认内网 LLM Base URL：`http://192.168.11.40:8000/v1`
- 默认模型：`qwen3.6-27b-fp8`
- 所有密钥只从 `.env` 读取，绝对不能写入代码、日志、测试、文档或回复。
- `.env` 应保留在本机并由 `.gitignore` 忽略。

当前环境变量：

```env
BAIDU_API_KEY=...
BAIDU_SECRET_KEY=...
LLM_BASE_URL=http://192.168.11.40:8000/v1
LLM_MODEL=qwen3.6-27b-fp8
LLM_API_KEY=...
```

## 3. 输入格式与基础解析策略

正式支持：

```python
SUPPORTED_SUFFIXES = {
    ".pdf",
    ".docx",
    ".jpg",
    ".jpeg",
    ".png",
}
```

旧版 `.doc` 明确不支持，必须在调用百度 API 前拒绝，并提示用户先转换为 `.docx`。不要为 `.doc` 引入 LibreOffice、COM 自动转换或新的解析依赖，除非用户重新明确授权。

| 文件类型 | 百度解析 | LLM 辅助参考文本 |
|---|---|---|
| 原生 PDF | 百度文档解析 | 使用 PyMuPDF 文本层 |
| 扫描 PDF | 百度 OCR | 文本层为空时允许无参考文本 |
| JPG/JPEG/PNG | 百度 OCR | 当前允许无辅助参考文本 |
| DOCX | 百度文档解析 | 使用 `python-docx` 提取段落和表格 |

PDF 必须使用 `import pymupdf`，不要改回已弃用的 `import fitz`。图片不做本地 OCR 或内容校验；扫描 PDF、空 DOCX 和图片应使用明确的“无参考文本”提示继续处理，不能仅因为缺少参考文本报错。

原始文件可能包含任意颜色、透明度、角度和形式的水印。Markdown 应保留正文内容和阅读结构，可以忽略人像照片、装饰图片、Logo、水印及重复页眉页脚。

## 4. 当前代码结构与公开入口

### 4.1 批量命令行入口：`main.py`

运行方式：

```powershell
python main.py
```

用户依次输入原始简历文件夹、输出文件夹和模板数字选项。当前菜单：

1. `SOTOS - 有照片占位`
2. `SOTOS - 无照片`

批量入口只扫描输入文件夹当前层级，不递归进入子文件夹。它会忽略 `~$` 开头的 Word/WPS 临时锁文件和不支持格式；文件无法读取时让用户重试、跳过或结束；单份失败不影响后续文件。

不同扩展名但 `stem` 相同的文件会生成同名过程数据和 Word，因此必须在调用外部 API 前阻止整批处理并提示用户重命名。

模板菜单由 `TEMPLATE_OPTIONS` 注册，后续新增模板时不要把交互和批处理逻辑写死成只支持 SOTOS。

### 4.2 单文件完整入口：`resume_pipeline.py`

```python
def ResumeConverter(
    input_file_path: str,
    output_dir: str,
    template_tag: str,
    with_photo: bool,
) -> str | None:
    ...
```

职责：依次执行文档解析、模板对应的信息提取和 Word 生成。当前模板注册表为 `TEMPLATE_HANDLERS`，标签不区分大小写。成功返回最终 Word 绝对路径；任一阶段真正失败返回 `None`。

### 4.3 文档解析入口：`data_parser.py`

```python
def DataParser(
    input_file_path: str,
    output_dir: str,
) -> str | None:
    ...
```

内部继续保持以下两个公开函数的名称和参数不变：

```python
def BaiduParser(input_file_path: str, output_dir: str) -> str:
    ...

def RestoreMarkdown(
    input_file_path: str,
    input_file: str,
    output_file: str,
) -> str:
    ...
```

输出路径：

```text
output_dir/process_data/document_parsing/XXX.md
output_dir/process_data/document_parsing/XXX_restored.md
```

所有格式统一使用原始文件 `stem` 命名。成功返回 `XXX_restored.md` 绝对路径；真正异常记录后返回 `None`。

### 4.4 信息提取入口：`information_extraction`

```python
def InformationExtractor(
    input_file_path: str,
    input_file: str,
    output_dir: str,
) -> str | None:
    ...
```

输入是原始文件和恢复后的 Markdown。成功返回：

```text
output_dir/process_data/information_extraction/XXX_extracted.json
```

每次成功还会根据该目录中的全部 JSON 重建：

```text
output_dir/process_data/information_extraction/resume_information.xlsx
```

JSON 是后续 Word 生成的主要数据源；Excel 用于批量查看、筛选和人工检查，包含“候选人汇总、工作经历、项目经历、教育经历、提取问题”五张工作表。

### 4.5 Word 生成入口：`resume_generation`

```python
def ResumeGenerator(
    input_file: str,
    output_dir: str,
    with_photo: bool = False,
) -> str | None:
    ...
```

输入是结构化 JSON。当前使用 `templates/standard_resume.docx` 生成 SOTOS 标准简历，成功返回：

```text
output_dir/final_resumes/原始文件名_标准简历.docx
```

同时创建或更新：

```text
output_dir/final_resumes/待补充信息.xlsx
```

该表固定三列：“候选人姓名、最终简历路径、需要补充的内容”，并按最终简历绝对路径保持一份简历一行。

`resume_generation/template_builder.py` 是模板设计变化时重新提炼清洁模板的维护工具，不是日常入口。不要在普通业务模块底部添加测试调用或写死业务文件路径。

## 5. Markdown 恢复原则

百度 Markdown 是内容主体，但可能出现标题层级、多栏顺序、左右文本归属、物理换行、遗漏、水印、重复页眉页脚和异常断词等问题。

LLM 只能恢复结构、顺序、物理换行、明显断词及原文件参考文本中确实存在而百度遗漏的正文。不得润色、概括、纠错、扩写或编造简历内容。

自动校验只能产生警告，不能阻断 `XXX_restored.md` 写入和返回。列表序号等 Markdown 结构数字不能当成简历事实数字。模型新增内容只要确实存在于原始参考文本中，就属于合法恢复。

## 6. 信息提取与证据规则

必须提取六部分：

- 基本资料：姓名、性别、出生年份、籍贯；
- 专业技能：保留简历原文形态的一段 Markdown 文本；
- 工作经历：时间、公司、岗位及可选的原文描述，按开始时间由近到远排序；
- 项目经历：项目名称、可选的岗位和时间、必需的原文项目描述，全部保留；
- 教育经历：时间、学校、学历、专业；
- 工作年限：按时间上最后一段教育的毕业年月计算。

最高原则是“宁可留空，也不能编造”：

1. 每个非空事实必须有可在恢复后 Markdown 中核验的原文证据。
2. LLM 返回后必须由 Python 再次校验。
3. 无法核验、原文缺失或含义模糊的字段应清空，并写入 JSON 的 `extraction_issues` 和 Excel 的“提取问题”。
4. 专业技能有明确栏目时尽量原样提取；没有独立栏目时只能拼接工作或项目中逐字可核验的技能原文，不能根据岗位或常识总结。
5. 工作描述是选填项，原文有就提取，没有就留空，不要求人工补充。
6. 项目岗位和项目时间是选填项，缺失不要求人工补充；项目名称和项目描述是必填项。

日期规则：

- 教育经历只写 `2022-2026` 时，按已确认业务规则标准化为 `2022/09-2026/06`，同时保存原始时间和规则说明。
- 工作和项目的纯年份区间不得擅自补月份。
- 学历只使用可核验并可规范到“博士、硕士、本科、大专”的值；模糊值不能强行映射。
- 工作年限由 Python 根据时间上最后一段学历的毕业年月计算，按 0.5 年向下取整；LLM 不参与计算。
- 计算依据只有最后学历毕业时间，不根据简历中可能不完整的工作经历累计。

## 7. SOTOS Word 排版与内容规则

整个文档统一使用微软雅黑，不允许其他字体声明残留。当前模板包含六个分区：基本资料、专业技能、工作经历、项目经验、教育经历、工作年限；项目经验英文标题必须为 `Project Experience`。

### 基本资料

- 固定显示姓名、性别、出生年份、籍贯和“岗位职称”。
- 岗位职称当前不自动提取，始终保留人工填写占位。
- 无照片模式使用两栏；有照片模式使用 `2:2:1` 三栏，第三栏只显示证件照占位框，不自动插入照片。
- 小四、1.5 倍行距。

### 专业技能

- 保留 Markdown 中的原文段落、换行、序号和项目符号。
- 小四、1.5 倍行距。

### 工作经历

- 标题行显示时间、公司、岗位，小四、加粗、模板蓝色。
- 三列宽度按本份简历内容统一计算；时间后保留视觉间距，岗位列起点对齐。
- 工作描述是选填项，原文有就写，没有就省略，不生成“待补充”。
- 工作描述五号、黑色、1.5 倍行距。
- 多段有描述的工作经历之间保留合理间距。

### 项目经历

- 按顺序编号“项目一、项目二……”；项目标题不包含公司名称。
- 标题格式：`项目一：项目名称 | 岗位 | 时间`，其中岗位和时间为选填，缺失时直接省略且不能留下多余分隔符或“待补充”。
- 标题小四、加粗、模板蓝色。
- 项目描述必填，位于标题下方，五号、黑色、1.5 倍行距。
- 分区标题和第一个项目之间、相邻项目之间应保留清晰但不过度的间距。

### 教育经历与工作年限

- 教育格式：`时间 | 学校 | 学历 | 专业`，小四、加粗、模板蓝色。
- 工作年限直接显示 `XX年`，支持 `.5`，小四、加粗、模板蓝色。

### 缺失内容与 Markdown 渲染

- 必填内容缺失时必须用“待补充”维持稳定版式；整个分区缺失也至少生成一行占位，不能直接跳过。
- 最终 Word 中只有“待补充”三个字使用纯红色 `#FF0000`；例如“公司名称待补充”中的前缀保持原有蓝色或黑色。
- `待补充信息.xlsx` 只汇总必填缺失项和始终需要人工确认的岗位职称，不收集工作描述、项目岗位或项目时间。
- 专业技能、工作描述和项目描述共用通用 Markdown 渲染器，不得针对具体候选人或具体词语修复。
- 标题、粗体、斜体、删除线、链接、引用、代码、表格及嵌套列表应转成相应 Word 语义；Markdown 图片忽略。
- `**文字**`、`##### 标题` 等控制符不能作为可见文字泄漏到 Word；`C++`、`C#`、`char*`、`#include` 等技术内容不能被误伤。
- Markdown 空行不能机械转换为额外 Word 空段落。
- Markdown 残留扫描只允许警告，不能中断最终 Word 生成。

## 8. 输出目录与过程记录

```text
output_dir/
  final_resumes/
    XXX_标准简历.docx
    待补充信息.xlsx
  process_data/
    processing_records.xlsx
    error_YYYYMMDD.txt
    document_parsing/
      XXX.md
      XXX_restored.md
    information_extraction/
      XXX_extracted.json
      resume_information.xlsx
```

解析 Markdown、结构化 JSON、汇总 Excel、错误日志和处理记录都是过程数据，只能放在 `process_data`。最终 Word 和面向用户的待补充清单放在 `final_resumes`，不要混放。

`processing_records.xlsx` 记录原始文件路径、百度 Markdown、恢复 Markdown、结构化 JSON、最终 Word、各阶段状态和错误信息。重复处理同一路径时应更新对应记录，不能无意义地产生重复行。

## 9. 错误、警告与可恢复行为

真正错误包括：

- 文件不存在、格式不支持或文件无法读取；
- 百度 API 鉴权、上传、轮询或下载失败；
- LLM API 调用失败或没有返回内容；
- 中间文件无法读取、JSON 无效、模板缺失或输出无法写入。

真正错误由对应公开阶段捕获，追加写入：

```text
output_dir/process_data/error_YYYYMMDD.txt
```

日志包含处理时间、处理阶段、原始文件路径和完整异常堆栈。同一天多次错误必须追加，不能覆盖。记录后返回 `None`，不要继续向终端抛出完整 Traceback。

证据校验失败、Markdown 恢复校验警告、最终 Word Markdown 残留警告都属于警告：记录或打印后继续生成可用结果。批处理中一份失败不能阻止下一份。

## 10. 开发与验证要求

- 每次开始修改前完整阅读本文件，并检查 Git 工作区和实际相关代码，以本地现状为准。
- 当前 Git 仓库根目录是 `D:/PythonCode/sotos-application`，项目位于其 `resume_converter` 子目录；注意不要误改相邻项目。
- 保留用户已有修改，不覆盖无关内容；不要使用 `git reset --hard` 等破坏性操作。
- 代码文件开头使用中文三引号 docstring，关键逻辑使用中文注释。
- 优先最小、清晰、可维护的修改；不要针对候选人、公司、项目或某份测试数据写特殊规则。
- 样例只用于发现通用问题，修复必须抽象为格式、语义或流程层面的通用规则。
- 不要在业务模块底部添加测试调用，不要写死用户桌面路径，不要把临时输出放在项目根目录。
- `main.py` 是已确认的正式批量 CLI，不能因旧任务描述将其删除；暂不新增 Web 层。
- 不要运行会消耗真实百度或 LLM 额度的测试，除非用户明确授权。
- 至少执行 `python -m py_compile ...` 和不访问外部服务的单元测试。
- 完整离线回归命令：

```powershell
python -m unittest discover -s tests -p "test_*.py"
```

- 测试至少覆盖：路径拼接、后缀分流、`.doc` 提前拒绝、扫描 PDF 空文本层、DOCX 参考文本、图片无参考文本、错误日志追加、警告不阻断、证据不匹配清空、专业技能原文保留、教育日期标准化、工作年限、JSON/Excel、批量文件发现与占用处理、模板分发、Word 占位、补充清单、Markdown 转 Word 和红色“待补充”。
- 涉及 Word 排版的修改，除结构测试外还应实际生成 DOCX，并渲染成页面图片检查全部页面；确认无重叠、截断、异常空行、Markdown 残留或分页问题。
- 完成后说明修改文件、行为变化、验证命令、验证结果和仍存在的限制。

## 11. 当前已知限制

- `.doc` 不支持，用户需自行转换为 `.docx`。
- 图片没有本地辅助 OCR，结构恢复主要依赖百度 Markdown。
- 扫描 PDF 没有文本层时无法做原文文本交叉核验。
- Word 生成当前只有 SOTOS 模板；其他标签尚未注册。
- 证件照模式只保留占位，不自动提取或插入照片。
- Markdown 图片不写入标准简历；Markdown 表格当前以可读文本行呈现，不生成复杂 Word 表格。
- 批量入口只处理输入文件夹当前层级，不递归子目录。

## 12. 用户沟通偏好

- 使用中文交流，先给结论和实际结果，再解释原因。
- 避免碎片化排版，不要两三个字就换行。
- 用户要求直接修改仓库时优先实施并验证，不要只给伪代码。
- 用户要求先讨论、先查看或先找方案时，不要提前修改代码。
- 不要过度设计；优先用最省力、可维护的方法得到可靠、可追溯、便于人工复核的结果。
