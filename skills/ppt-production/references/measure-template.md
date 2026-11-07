# 测量模板

先回答画布多大、页标题与正文分别多大、页眉/页脚占用哪里。显式字号分布只是样本，不能把模板内小字号注释直接当正文下限。

```python
from collections import Counter
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = Presentation(template_path)
print('canvas_in', prs.slide_width / 914400, prs.slide_height / 914400)

def leaf_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from leaf_shapes(sh.shapes)
        else:
            yield sh

sizes = Counter()
unresolved_characters = 0
for slide_no, slide in enumerate(prs.slides, 1):
    for sh in leaf_shapes(slide.shapes):
        frames = [sh.text_frame] if sh.has_text_frame else []
        if sh.has_table:
            frames += [c.text_frame for row in sh.table.rows
                       for c in row.cells if not c.is_spanned]
        for frame in frames:
            for p in frame.paragraphs:
                for r in p.runs:
                    n = len(r.text.strip())
                    size = r.font.size if r.font.size is not None else p.font.size
                    if size is None:
                        unresolved_characters += n
                    else:
                        sizes[size.pt] += n
print('explicit_or_paragraph_sizes', sorted(sizes.items()))
print('unresolved_characters', unresolved_characters)
```

- 未解析字号要继续检查版式、母版和主题，不按 0 pt 统计。标题、正文、图注分别抽样，不只看最大字号或字符量最多的字号。
- 页眉线可能在母版、组合图形或背景图片中；标题框底边不是页眉线的位置。先检查 XML/图片，再从渲染图确认。
- `<a:srgbClr>` 是局部颜色；主题 `<a:clrScheme>` 与 `<a:schemeClr>` 是另一层。某页出现紫色不代表整套主题都用该紫色。
- 解析某页失败时记录页号和错误；可使用 ZIP 中的 OOXML 及渲染图补查。未能检查的页必须标为未核实，不能静默跳过后声称已完成测量。
- 若要编辑字体节点，使用 `qn('a:ea')` 等完整命名空间，先 `el = rpr.find(qn('a:ea'))`，再以 `if el is None:` 创建。不要用 `find(...) or ...`，空 lxml 元素可能被误判。