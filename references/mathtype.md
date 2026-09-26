# MathType：先证明对象可编辑，再批量生成

## 选定交付格式

本技能的 OLE 路线面向 Windows Word 与支持 `Equation.DSMT4` 的 MathType 桌面版。MathType for Microsoft 365 的对象格式不同；从 MathType 7 转换到 Microsoft 365 后不能再用 MathType 7 编辑，见 [Wiris 使用说明](https://docs.wiris.com/mathtype-for-microsoft-365/get-started-with-mathtype-in-microsoft-word)。如果用户明确要桌面 OLE，不能悄悄换成该格式。

先在本机生成并重开一个对象，记录 Word/MathType 版本、32/64 位、对象 ProgID 和所用生成路线。不要依赖固定安装目录或复制其他稿件的公式。

## 按已验证能力选择最快路线

### A. 桌面 MathType 已能直接接受当前表达式

将 PDF 中的公式转录为 TeX 或 Presentation MathML，放入 MathType 编辑器确认，再插入 Word。对简单 TeX 可用 Word 中 MathType 的 **Toggle TeX**；复杂的矩阵、分段函数、重音和嵌套上下标需要逐类小样。

普通 `\[...\]` 不等于带 MTEF 信息的 MathType translator text；不能将普通 TeX 文本交给 **Convert Equations** 后假设会整篇成功。该对话框对 translator text 有格式要求，**Toggle TeX** 是不同入口。[Wiris 桌面 Word 集成说明](https://docs.wiris.com/mathtype-7-with-microsoft-office-2016-or-later?kb_language=en_US)

### B. 已有可靠的 MathML → OMML 工具链

先从本次 PDF 转录生成 OMML，然后在桌面 MathType 的 **Convert Equations** 中只选当前确实存在的 OMML 类型，目标为 MathType 对象。先对复杂小样验证，确认后整批转换；中间 OMML 不作为最终 OLE 交付。

检查转换前后公式清单，每个 ID 都应对应正确原生对象。宏/插件不可用时不要全局降低 Word 安全设置，改用可用的原生编辑入口或报告需要用户启用已安装的受信任组件。

### C. 已安装并验证的 MathType SDK/自动化适配器

适合大量公式，使用本机版本的正式 SDK 声明。典型生命周期是初始化 API → 创建有效空 MathType OLE → 激活对象 → 写入 MathML → 保存并关闭对象 → 存 DOCX → 终止 API。具体函数、返回值、字符串长度单位及位数必须以该版本 SDK 为准。

历史版本可能提供 `MTInitAPI`、`MTSetEqnFromLangStr` 和 `MTCloseOleObject` 等入口；本仓库不附厂商 DLL、空公式模板或未经版本验证的底层绑定。不能从 CPython 对象地址加固定偏移获取 COM 指针：这依赖解释器/架构内存布局，存在崩溃或写错对象风险。使用有类型的 COM 接口或厂商示例适配器。

没有经过小样验证的 SDK 路线时，用 A/B；不要把临时内存技巧包装成通用转换器。Word/MathType 同一实例只允许一个写入者，按稳定 ID 每批保存。断点续跑先核对已保存对象，再从缺项继续。

## 公式清单和语义

保留 TeX/MathML 源及 PDF 页/区域，防止仅看预览而丢失真实内容。需要重点核对：大括号、绝对值、向量粗体、正斜体、转置/共轭、上横线/帽子、上下标、分段条件、编号。

若 MathML 导入缺字，先确认字体和导入器支持。数学字母可表示为普通字符加 `mathvariant`；转换时保持各字符样式，不对整个公式做 NFKC 折叠。上横线/帽子应采用重音结构；不要全局替换 Unicode 字符“碰运气”。删掉不可见分隔符前核对是否改变函数应用语义。

## 将公式画廊插入正文

最容易验证的是 Word 自身复制原生对象并保持 OLE。大量对象若用 Open XML 复制，必须同时复制：

- `w:object`、VML/相应预览节点及尺寸；
- 内嵌 `.bin` 与预览媒体文件；
- 目标 part 的 relationship、唯一 rId/ShapeID/ObjectID；
- `[Content_Types].xml` 中所需内容类型。

只复制 XML 标签或 PNG 都不够。不能把标记为 `Equation.DSMT4` 的随机 CFB 文件视为成功；要检查有效的 `Equation Native` 数据流并实际进入 MathType 编辑。

## 保存后的验证

核对全部对象清单，至少对每类复杂结构和被修复公式做编辑→保存→关闭→重开验证。在最终 DOCX 中重新抽查，而非只在独立画廊抽查。若预览来自 PDF 回退，必须另外确认原生内容一致，并说明再次编辑后预览可能变化。
