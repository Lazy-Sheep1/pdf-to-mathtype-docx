# 内容清单与正文核对区域

工作目录保存用户文稿内容；不要将它放进公开技能仓库。

## 建议的主清单

```json
{
  "source_sha256": "SHA256_OF_INPUT_PDF",
  "items": [
    {"id":"p001","kind":"paragraph","source_page":1,"source_bbox":[40,100,550,180],"text":"Example paragraph.","checked":false},
    {"id":"eq001","kind":"inline_equation","source_page":1,"source_bbox":[100,120,140,140],"tex":"x_i^2","mathml":null,"status":"transcribed","checked":false},
    {"id":"fig001","kind":"figure","source_page":2,"source_bbox":[40,80,500,260],"asset":"work/figure-001.png","checked":false}
  ]
}
```

页码从 1 开始，坐标单位 point，原点在左上；旋转 PDF 先统一页面方向并保存变换。来源与 Word 导出后的坐标分别保存，允许页码、栏数变化。公式状态可依次为 `transcribed`、`native_created`、`inserted`、`visually_checked`、`edit_checked`；失败项保留错误，不能静默删除。

重复公式用 `formula_id` 复用转录缓存，用唯一 `occurrence_id` 标记每个出现位置；每个位置都需实际插入一个正确对象。清单还可保存目标 Word 书签或内容控件 ID，用于重新分页后定位正文核对区域，减少手工圈选。

主清单由代理根据 PDF 建立；`pdf_inventory.py` 仅提取页面文字/几何信息，不自动认定自然段或公式。

## `text_integrity.py` 的输入

脚本核对 **Word 导出 PDF** 中已确认的正文区域。跨页/跨栏段落按阅读顺序列出多个区域；坐标必须来自当前渲染，不能沿用来源 PDF 坐标。为 OLE 公式两侧正文分别建项，公式由公式清单另行核对。

```json
{
  "entries": [
    {
      "id": "paragraph-001-part-a",
      "text": "This is a complete sentence [12].",
      "regions": [{"page":1,"bbox":[40,100,550,180]}]
    }
  ]
}
```

脚本默认只归一化空白、软连字符和常见排版连字；数字、普通连字符、大小写和标点保留。合法断词若在 PDF 中成为普通 `-`，会报差异；检查源词后在该项提供 `discretionary_breaks`，例如 `["para-\ngraph"]`，只消除明确列出的换行断词。不得全局删除连字符或标点来“通过”检查。

没有建立覆盖所有正文的清单时，报告只能说已核对列出的区域，不能说全文内容完整。页眉、页脚、脚注、表格和引用也需要各自登记。此脚本不识别视觉碰撞、字距松散或公式语义。
