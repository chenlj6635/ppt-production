# 素材库维护

默认素材目录是技能根目录下的 `assets/library/`。所有 source_path 与 preview 路径都相对于素材目录，不依赖某位用户的盘符或 HOME。项目 AGENTS.md 可指定另一份库。

## 使用前

运行 `python scripts/check_library.py assets/library`（工作目录为技能目录）。它验证源 PPTX 指纹、原生形状提取、页面顺序、slide_id、XML 指纹和每张预览的指纹。报错时先修复索引，不盲信旧页码。该工具验证文件一致性，不证明文字/图表可读或语义标签正确。

## 新增或修改素材

1. 在自己的素材库副本添加、替换或删页。保留原生文字、表格、图片和组合关系。不要把多个风格来源自动当成一套统一主题。
2. 提取更新后的对象信息：`python scripts/index_library.py assets/library/source.pptx --out assets/library/inventory.json`。命令默认把源路径写成相对路径；源文件不被修改。
3. 用本机 PowerPoint/WPS 或可用演示软件导出 PDF，再渲染为 `previews/page-01.png` 等编号图片。每页都打开检查，不能只更新 JSON 而继续使用旧预览。
4. 更新 `catalog.json`。优先根据 slide_id 对照旧条目；重建或跨文件复制后不能假定 id 不变。删除的页移出当前索引，XML 指纹或内容变化的页重新判断结构用途。
5. 每个条目保留 `id`、`page`、`slide_id`、`slide_part`、`slide_xml_sha256`、`name`、`level`、`tags`、`regions`、`use_for`、`caution`、`preview`、`preview_sha256`、`reviewed_at` 和 `review_scope`。source_path 与 source_sha256 对齐 inventory；preview_sha256 在实际重新渲染并检查后再更新。
6. 运行一致性检查，再运行 `python scripts/render_catalog.py assets/library` 生成 `index.md` 和 `gallery.html`。生成器不会自动猜测语义标签或悄悄更新哈希，未同步的素材会被拒绝生成。
7. 本地打开图库筛选，并查看变化页。检查已无效的旧预览不再进入索引；提交前查看 Git diff，确认未包含私有信息。

## 复制结构到新稿

通过 slide_id/slide_part 定位页，查看 inventory 中的 shape id 和 children 定位结构。group 子形状坐标需要套用 chOff/chExt 的缩放，不能当作顶层坐标直接使用。

将标题、框、图片、说明作为完整结构单元选取，或保留语义组织后重建可编辑对象。跨文件复制后确认图片/图表关系、主题字体和布局完整；预览截图只用于挑选，不替代可编辑成品。

## 指纹的边界

PPTX 的 Office 元数据变化也会改变文件 SHA-256，即使幻灯片内容没有变。遇到这种情况比较 slide_xml_sha256 和相关媒体后再更新，不能简单关闭校验。指纹证明文件对应关系，不证明来源权属或渲染正确性。