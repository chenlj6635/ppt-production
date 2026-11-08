# 可编辑原生输出：一个技能内的可选工具路径

适用于从零制作新稿。已有原稿修改优先保持原有原生对象，见 `editing-and-verification.md`。不是所有 SVG 都受支持；转换器明确拒绝时应改成受支持的结构，不用截图冒充可编辑输出。

## 项目输入

```text
work/
  design.json
  outline.md
  svg_output/01.svg
  svg_output/02.svg
  objects.json           有原生数据图表或表格时提供
  images/                使用到的本地图片
```

文件按零补齐页号排序。设计参数必须来自当前项目要求或实际模板测量。下面仅示意字段，不是默认主题；颜色、字体和画布均应换成当前目标稿的值。

```json
{
  "title": "本次汇报标题",
  "canvas": {"width_px": 1280, "height_px": 720},
  "body_min_pt": 20,
  "font_latin": "Arial",
  "font_east_asian": "Microsoft YaHei",
  "text_color": "#222222",
  "background_color": "#FFFFFF",
  "table_header_fill": "#EDEDED",
  "palette": ["#222222", "#EDEDED", "#FFFFFF"]
}
```

SVG 采用 96 px/in；例如 20 pt 对应约 26.67 SVG px。不得把 SVG 的 `font-size="20"` 当成 20 pt。声明字号低于 20 pt 时，必须以 `body_min_exception` 记录当次用户授权的例外。

## SVG 原生形状与 Office 公式

SVG 用普通 rect/path/text/image 等受支持元素表达页面。字号、字体、颜色必须显式或可继承解析，视口严格等于 design.json；不隐式开启动画和过渡。真实图像保留为图像，文字与结构图使用可编辑文字/形状。

结构公式采用明确的公式标记，LaTeX 在本机编译为 OMML。只支持转换器的 LaTeX 输入范围，不执行完整 TeX 程序、任意宏包或 shell。未知命令报错，不自动发送到外部公式服务。

```xml
<g data-pptx-replace-with="formula">
  <metadata type="application/json"><![CDATA[
    {"latex":"\\alpha=\\frac{k}{\\rho c_p}",
     "x":100,"y":220,"width":520,"height":120,
     "font_size":36,"color":"#222222","align":"left"}
  ]]></metadata>
  <text x="110" y="290" font-family="Cambria Math"
        font-size="36" fill="#222222">α = k/(ρ cₚ)</text>
</g>
```

JSON 中反斜杠按 JSON 规则转义。预览文字是等价的普通 SVG 表示，禁止直接显示 LaTeX 代码。PowerPoint 中应检查分式、上下标和希腊字母的可编辑性与高度；预览本身不能证明 OMML 实际显示正确。

## 原生数据图表和表格

在 `objects.json` 提供真实数据。矩形范围用与 SVG 相同的 px 坐标，`slide` 从 1 开始。对应位置在 SVG 中保留空位；这些对象在 SVG 形状之后加入，会覆盖同位置内容，因此最终必须读 PPTX 导出图，不能仅预览 SVG。

```json
[
  {"slide":1,"type":"chart","chart_type":"column",
   "bounds":{"x":80,"y":200,"width":540,"height":320},
   "categories":["方案 A","方案 B"],
   "series":[{"name":"示意数据","values":[2,3]}],"legend":false},
  {"slide":1,"type":"table",
   "bounds":{"x":730,"y":240,"width":440,"height":220},
   "rows":[["参数","单位"],["长度","mm"],["温度","K"]]}
]
```

数值只作格式示例，不可当科研数据。当前封装验证了柱形、条形、折线图与规则矩阵表格；图表包含嵌入的 Excel 数据。图例与坐标轴采用项目字体/字号，颜色来自项目 palette，也可明确提供 `series_colors`。复杂组合图、散点图、特殊表格合并可使用宿主的原生能力，不宣称此接口已覆盖。不要把上游全功能直接等同于本封装已测试的能力。

## 执行与验收

只安装了技能目录时，用 `python -m pip install -r <skill_dir>/requirements.txt` 安装工具依赖。运行目录无需等于技能目录，脚本路径展开为当前宿主实际安装位置。

```sh
python scripts/native_export.py check work --report work/check.json
python scripts/native_export.py export work --output work/exports/deck-v1.pptx --report work/export-report.json
python scripts/pptx_audit.py work/exports/deck-v1.pptx --min-pt 20
```

`export` 每次重新执行静态检查，不依赖旧的成功报告。输出文件必须不存在，避免直接覆盖已有交付物。显式六位十六进制颜色、原生数据对象颜色和文字字体必须符合设计规范；SVG 其他颜色写法、复杂变换及真实字形边界仍须读图检查。图表/表格通过 objects.json 构建，SVG 中仅开放公式原生替换标记，避免绕过统一风格设置。明确错误会阻断导出；警告保留待复核。`visual_reviewed` 始终是 false，因为脚本看不到实际渲染，不能自动替 agent 签署验收。

随后按 SKILL.md 导出 PDF/PNG 并逐页读图；检查图表坐标轴与图例、表格换行、公式高度和组图图注。修改后重新导出。校验文件存在、静态检查通过、包内存在 `c:chart`/`a:tbl`/`m:oMath` 仅分别证明文件/结构，不代表可见效果。

## 实现来源与更新

SVG/OMML 转换部分来自 Hugo He 的 PPT Master，固定到 `f6cfd05fe1e3f2144906c5f9ae8ec5d91612eb0b`。归档位于 `assets/engines/`，保留原始文件、MIT 许可和上游署名检查；哈希清单随附。运行时仅解包到临时缓存，不额外安装可触发的 ppt-master 技能，也不读取其设计指令作为本技能的风格。

原生数据图表/表格封装使用 python-pptx。资料提取使用 pypdf、python-docx 和 openpyxl，依赖见仓库 requirements.txt。当前不引入上游图片生成、网页登录、声音、视频或自动更新功能。

不要直接浮动升级转换库；更新时重新固定提交与哈希，保留第三方署名，跑原生对象测试及至少一份渲染验收样例。
