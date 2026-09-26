# PDF to MathType DOCX Skill

把学术 PDF 重建成正文可编辑、公式可在 MathType 中继续编辑的 Word 文档。重点是公式完整显示、按段落组织和全文复核。

这是一份供 AI 编程代理执行的 **skill + 检查工具**。公式识别、数学转录和版式决策仍需代理结合原 PDF 完成；仓库没有一键把任意 PDF 转成原生 MathType 的功能。

## 最快可靠路径

```mermaid
flowchart LR
    A[PDF 盘点] --> B[复杂公式小样]
    B --> C[一次转录与清单]
    C --> D[连续段落及 MathType 对象]
    D --> E[Word 导出 PDF]
    E --> F[全文内容与视觉复核]
    F --> G[针对问题修复并交付]
```

先验证最难公式的生成与重开编辑，再批量生成。允许重新排版时，正文按自然段流动；公式、图表、算法按实际尺寸分配空间。

## 安装与调用

下载仓库 ZIP 并解压，把含 `SKILL.md` 的目录命名为 `pdf-to-mathtype-docx`，放入你的技能目录。例如常见本地 Codex 配置使用 `$CODEX_HOME/skills`（未自定义时通常为 `~/.codex/skills`）；也可按所用代理支持的方式加载该文件夹。

调用示例：

> 使用 $pdf-to-mathtype-docx，只以这篇 PDF 为参照，转换为可编辑 Word。全部公式使用桌面 MathType OLE，完整显示优先，可以重新规划排版。正文两端对齐且词距尽量均匀，最后检查全文并修复问题。

也支持继续修复已有 DOCX 的重叠、断括号、模糊图片和错位；基本合格的文档采用局部修复。

## 环境

- Python 3.10+；通过 `python -m pip install -r requirements.txt` 安装辅助工具依赖。
- PDF 盘点、结构检查和正文区域核对可跨平台运行。
- 原生 Word 渲染需要 Windows、Microsoft Word 和 pywin32。
- 桌面 MathType OLE 的生成/编辑需要兼容的 MathType 桌面版及有效使用环境。Microsoft 365 MathType 插件与桌面 OLE 不等同。

本仓库不包含 Word、MathType、厂商 DLL、字体、用户论文或任何原文图像。

## 文件

| 入口 | 用途 |
| --- | --- |
| [SKILL.md](SKILL.md) | 代理执行流程、优先级与结束条件 |
| [MathType 路线](references/mathtype.md) | 导入路径、小样验证与 OLE 包装注意点 |
| [排版规则](references/layout.md) | 段落、对齐、公式、图像与原版面回退 |
| [内容清单](references/manifest.md) | 来源追踪、公式出现位置与正文核对区域 |
| [复核修复](references/qa-and-repair.md) | 全文检查顺序与常见问题的修复方法 |
| `scripts/pdf_inventory.py` | 提取 PDF 文本/几何信息并渲染逐页预览 |
| `scripts/audit_docx.py` | 检查 OLE 内容、预览关联和排版风险 |
| `scripts/word_render.py` | 在独立 Word 实例导出 PDF |
| `scripts/text_integrity.py` | 比较已映射正文区域，保留标点与数字 |

## 辅助工具示例

在仓库目录中执行，输入和输出路径可换为自己的路径：

```powershell
python scripts/pdf_inventory.py source.pdf --out work/source
python scripts/audit_docx.py output.docx --expected-mathtype 12 --out work/docx-audit.json
python scripts/word_render.py output.docx --output work/word.pdf
python scripts/text_integrity.py work/word.pdf work/text-regions.json --out work/text-audit.json
```

`12` 是示例数量；以本次清单实际出现的公式数为准。正文区域 JSON 格式见 [清单说明](references/manifest.md)。PDF 盘点不执行 OCR，文本存在也不代表 OCR 正确。

Word 渲染脚本串行执行，不绑定或关闭用户现有的 Word；遇到插件弹窗可能需要在系统界面处理。不要并发运行多个 Word/MathType 写入任务。

检查结果只覆盖其声明范围：ZIP/OLE 结构通过不能证明实际可编辑，正文区域匹配不能证明已覆盖全文，几何碰撞候选也不能直接证明实际重叠。最终需进入 MathType 验证，并查看 Word 渲染的全部页面。

## 开发验证

```powershell
python -m unittest discover -s tests -v
```

测试仅使用临时合成数据，覆盖损坏 OLE 关系、空原生流、预览不匹配、缺标点、缺数字、真连字符、无效核对区域与 Word 自动化清理。Word COM 单元测试使用 mock；真实 Word 冒烟验证需在安装了 Word 的 Windows 环境另行运行。

## English summary

A reusable agent skill for rebuilding formula-heavy academic PDFs into editable Word documents with native desktop MathType OLE equations. It favors paragraph reflow, proves the equation workflow on difficult examples before batching, and verifies the final document through Word rendering, content checks, and visual inspection. The bundled utilities support inspection and QA; they do not automatically transcribe arbitrary equations or replace MathType.
