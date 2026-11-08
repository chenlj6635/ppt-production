# 第三方实现说明

本技能部分原生转换实现来自 **PPT Master**，作者 **Hugo He**。

- 上游仓库：https://github.com/hugohe3/ppt-master
- 固定提交：`f6cfd05fe1e3f2144906c5f9ae8ec5d91612eb0b`
- 适用文件：`assets/engines/ppt-master-native.zip` 中的文件，完整清单和 SHA-256 见 `assets/engines/manifest.json`。
- 许可证：MIT，`Copyright (c) 2025-2026 Hugo He`；完整许可保存在 `assets/engines/LICENSE.ppt-master` 及归档内部。
- 上游文件保持原文与原始署名检查。该归档是转换工具运行依赖，不是将 PPT Master 改名为本项目，也不表示上游作者为本项目背书。
- 本项目新增的流程、个人风格规则和调用封装由本项目维护；上游工具的固定设计风格、模板资产、图像生成服务和自动更新流程未作为本技能默认工作方式。

python-pptx、pypdf、python-docx、openpyxl、XlsxWriter 等通过 requirements 安装，未将第三方安装包再打包进本仓库。个人素材库的授权说明仍见仓库 `NOTICE.md`，它不与转换引擎的 MIT 许可混同。
