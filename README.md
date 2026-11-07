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

只读取技能和素材不需要 Python。运行审计需要 Python 3.10+ 和 python-pptx；原生素材提取、图库生成与一致性检查只依赖 Python 标准库。

```sh
python -m pip install -r requirements.txt
python skills/ppt-production/scripts/pptx_audit.py your-deck.pptx --min-pt 20
python skills/ppt-production/scripts/check_library.py skills/ppt-production/assets/library
python skills/ppt-production/scripts/render_catalog.py skills/ppt-production/assets/library
python -m unittest discover -s tests -v
```

审计默认正文、表格和图注不低于 20 pt，可通过 `--min-pt` 调整。真正页码字段/占位符豁免。退出码：0 表示没有明确字号违规，1 表示有违规，2 表示输入无法检查。重叠、继承字号和符号字形是人工复核提示，不代表真实遮挡或缺字。审计不做 OCR，不解析所有母版字号；自动检查通过不等于视觉验收通过。

PPTX 转 PDF 和逐页渲染依赖本机可用的演示软件与 PDF 栅格化工具，由宿主选择。不要把 Word COM 用来转换 PPTX。图库不会后台自动跟随源文件变化。

## 素材与发布范围

本仓库公开提供技能规范、辅助工具和个人收集的参考结构素材。原始个人文件未改写；仓库副本已去掉 Office 作者属性和本机绝对路径。

本次没有为仓库授予开源许可证。技能文字/工具与第三方 PPT 图片、图表、摘录的来源及授权范围不同；详见 [素材说明](NOTICE.md)。仓库公开可见不代表对第三方素材授予再分发许可；使用时请核对相应来源与授权。