# 标准简历Word生成

## 单份原始简历完整流程

日常批次中的单文件处理优先调用统一入口，它会依次完成文档解析、信息提取和对应模板的Word生成：

```python
from resume_pipeline import ResumeConverter

word_file = ResumeConverter(
    input_file_path=r"D:\input\张三.pdf",
    output_dir=r"D:\output",
    template_tag="SOTOS",
    with_photo=False,
)

print(word_file)
```

模板标签不区分大小写，当前支持 `SOTOS` 和 `优族`。SOTOS 的 `with_photo=False` 生成三行两列基本资料，`with_photo=True` 生成带证件照占位的 `2:2:1` 基本资料；优族固定复用 SOTOS 无照片正文布局，并移除全部页眉。模板标签或证件照模式无效时，会在调用百度和LLM之前被拒绝，并写入统一错误日志。

## 只执行Word生成阶段

公开入口：

```python
from resume_generation import ResumeGenerator

result = ResumeGenerator(
    input_file=r"D:\output\process_data\information_extraction\张三_extracted.json",
    output_dir=r"D:\output",
    with_photo=False,
)

print(result)
```

参数说明：

- `input_file`：`InformationExtractor` 生成的结构化JSON路径。
- `output_dir`：本次任务的输出根目录。
- `with_photo=False`：基本资料使用两栏。
- `with_photo=True`：基本资料使用 `2:2:1` 三栏，第三栏只生成证件照占位框。

缺失信息会使用“待补充”保留可人工编辑的稳定版式：

- 基本资料缺失值及岗位职称显示“待补充”。
- 专业技能、工作经历、教育经历或工作年限整段缺失时，生成对应占位行。
- 工作描述是可选项，原简历有就写入，没有就直接省略，不生成占位。
- 项目名称或必需的项目描述缺失时生成占位；项目岗位和时间是可选项，缺失时直接省略，不生成占位或多余分隔符。

成功时返回：

```text
output_dir/final_resumes/原始文件名_标准简历.docx
```

同一目录还会创建或更新：

```text
output_dir/final_resumes/待补充信息.xlsx
```

该表按最终简历绝对路径保持一份简历一行，包含“候选人姓名”、
“最终简历路径”和“需要补充的内容”三列。工作描述、项目岗位和项目时间
属于选填信息，缺失时不会列入补充清单。

失败时返回 `None`，完整异常追加到：

```text
output_dir/process_data/error_YYYYMMDD.txt
```

如果处理记录表存在，生成结果还会写入：

```text
output_dir/process_data/processing_records.xlsx
```

`ResumeGenerator` 用于已有结构化JSON时单独调试Word生成阶段。`template_builder.py` 仅用于标准模板设计发生变化时，从新的人工参考文档重新提炼清洁模板。

优族单独调试 Word 生成阶段时，可调用 `YouzuResumeGenerator(input_file, output_dir)`；它与 `ResumeGenerator(..., with_photo=False)` 共用正文生成逻辑，只额外移除页眉。
