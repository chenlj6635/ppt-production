# PPT Production · 区域结构化演示文稿


让 agent 先组织信息，再组织页面：文字、图片和表格组成有明确职责的区域，区域之间表达比较、流程、因果或证据关系，最后统一对齐、留白与视觉主次。


这是一份可安装的 Agent Skill，包含中文制作规范、13 页可编辑结构素材、按用途检索的图库，以及 PPTX 字号/布局辅助审计和素材一致性检查工具。


## 核心方法


1. **结构单元**：标题、容器、证据与解释组成一个有含义的整体，避免游离小文本。
2. **页面逻辑**：按内容安排各区域的角色与关系，允许大小和数量不同。
3. **平面整合**：统一网格、间距、配色、字体层级和视觉重心。


优先采用合适的“标题—框—内容”组合。区域也可以通过邻近关系、对齐或底色建立，不要求每层都套框，不固定每页三张卡片。项目和用户当次要求优先于技能默认值。


## 内容入口


- [技能入口](skills/ppt-production/SKILL.md)
- [区域结构化原则](skills/ppt-production/references/regional-structure.md)
- [素材索引](skills/ppt-production/assets/library/index.md)
- [素材图库 HTML](skills/ppt-production/assets/library/gallery.html)：克隆后用浏览器打开，可按“标题框、流程、图注、双证据”检索。
- [可编辑素材源](skills/ppt-production/assets/library/source.pptx)
- [素材更新流程](skills/ppt-production/references/library-workflow.md)


素材库是组件和结构范例的集合，不是要求所有项目照搬的完整模板。科研图片、姓名、原作者数据和结论不会自动成为新稿的事实。


## 渲染与审计脚本


- `scripts/render_deck.ps1`：一条命令导出 PDF 并栅格化成逐页 PNG。Windows 走 PowerPoint/WPS COM，装了 LibreOffice 可加 `-Engine soffice`。
- `scripts/pptx_audit.py`：字号、字形和包围框审计，`--min-pt 20` 是默认下限。
- `scripts/check_library.py`、`scripts/index_library.py`、`scripts/render_catalog.py`：素材库一致性检查、索引重建与图库渲染。

```powershell
& .\skills\ppt-production\scripts\render_deck.ps1 -Pptx .\deck.pptx -Pages 3,8
python skills/ppt-production/scripts/pptx_audit.py .\deck.pptx --min-pt 20
```

脚本退出码 0 只代表没发现明确违规，不等于视觉验收通过；交付前必须逐页读图。


## 安装和使用


仓库地址：https://github.com/chenlj6635/ppt-production 。


克隆或下载仓库后，将 `skills/ppt-production/` 整个目录安装到宿主支持的个人技能目录。必须保留 `references/`、`scripts/`、`assets/`，不能只复制 SKILL.md。已有同名技能时先备份，再决定是否替换。


例如，在本仓库根目录用 PowerShell 安装到 Codex 的个人技能目录：


```powershell
$skillHome = if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $env:USERPROFILE '.codex\skills' }
$target = Join-Path $skillHome 'ppt-production'
if (Test-Path -LiteralPath $target) { throw '目标技能已存在，请先备份并明确选择是否替换。' }
New-Item -ItemType Directory -Path $skillHome -Force | Out-Null
Copy-Item -LiteralPath '.\skills\ppt-production' -Destination $target -Recurse
```


按宿主方式刷新技能或开启新会话。可显式调用：


> 使用 $ppt-production 制作一份可编辑的组会 PPT。先按区域结构化组织每页逻辑，再选用素材库中的合适结构，完成后渲染逐页检查。


也可让其他 agent 直接读取 `skills/ppt-production/SKILL.md`。本仓库不修改 Magpie 全局指令，也不要求 Magpie 才能使用。


## 工具与依赖


只读取技能和素材不需要 Python。执行工具使用 Python 3.10+；`requirements.txt` 包含原生 PPTX、资料提取和嵌入图表数据需要的依赖。素材索引、图库生成与一致性检查仅依赖 Python 标准库。

同一个 `ppt-production` 入口现在覆盖资料提取、内容策划、项目设计规范、页面制作、静态检查和渲染验收。使用说明见 [全流程](skills/ppt-production/references/production-workflow.md)、[原生对象制作](skills/ppt-production/references/native-authoring.md) 和 [模板复用维度](skills/ppt-production/references/template-layers.md)。项目风格约束持续生效，不继承转换引擎的设计主题。

原生路径支持受支持的 SVG 转可编辑 DrawingML 形状、本地 LaTeX 转 Office Math，以及柱形/条形/折线数据图表和矩阵表格。SVG/OMML 后端取自固定提交的 PPT Master，随技能携带，不额外安装另一个技能；作者与许可见 [第三方实现说明](skills/ppt-production/THIRD_PARTY_NOTICES.md)。这不代表兼容任意 SVG、TeX 宏包或上游全部功能。


```sh
python -m pip install -r requirements.txt
python skills/ppt-production/scripts/source_intake.py report.pdf --out work/sources
python skills/ppt-production/scripts/native_export.py check work --report work/check.json
python skills/ppt-production/scripts/native_export.py export work --output work/deck-v1.pptx
python skills/ppt-production/scripts/pptx_audit.py your-deck.pptx --min-pt 20
python skills/ppt-production/scripts/check_library.py skills/ppt-production/assets/library
python skills/ppt-production/scripts/render_catalog.py skills/ppt-production/assets/library
python -m unittest discover -s tests -v
```


审计默认正文、表格和图注不低于 20 pt，可通过 `--min-pt` 调整。真正页码字段/占位符豁免。退出码：0 表示没有明确字号违规，1 表示有违规，2 表示输入无法检查。重叠、继承字号和符号字形是人工复核提示，不代表真实遮挡或缺字。审计不做 OCR，不解析所有母版字号；自动检查通过不等于视觉验收通过。


PPTX 转 PDF 和逐页渲染依赖本机可用的演示软件与 PDF 栅格化工具，由宿主选择。不要把 Word COM 用来转换 PPTX。图库不会后台自动跟随源文件变化。


## 素材与发布范围


本仓库公开提供技能规范、辅助工具和个人收集的参考结构素材。原始个人文件未改写；仓库副本已去掉 Office 作者属性和本机绝对路径。


自创代码与技能文字采用 [MIT License](LICENSE)；`skills/ppt-production/assets/library/` 的素材集合、派生索引、预览图和图库展示不适用该许可证。第三方 PPT 图片、图表和摘录仅供本仓库公开查看与研究，不代表已授予再分发、发布或商业使用授权；详见 [素材与许可说明](NOTICE.md)。
