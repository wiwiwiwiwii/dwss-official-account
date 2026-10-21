# 报告转公众号 · 维护说明

本文件随 EXE 分发（`<exe> --export-maintenance <空目录>` 导出），面向接手修改的开发者或
AI agent。导出目录与开发工作区同布局：`src/` 为完整源码，`docs/wechat-mp-official/` 为
公众号排版规范清单，`tests/wechat_compliance_check.py` 为离线合规检查。

## 产品

Windows 本机工具：把东吴证券 2026 标准研报模板的 Word 报告（.docx/.doc）转成可直接
粘贴进微信公众号编辑器的 HTML。浏览器里的本地工作台负责选文件和预览、复制。
唯一外部前置是已激活的 Microsoft Word 桌面版（WPS 不能替代）；其余依赖都在 EXE 内。

## 源码结构

| 文件 | 责任 |
| --- | --- |
| `src/app.py` | 本机 HTTP 服务（127.0.0.1 随机端口 + 随机令牌）、任务队列、标题装配、ETA、命令行入口 |
| `src/maintenance.py` | `--help`/`--version` 与参数错误输出到终端、维护包导出 |
| `src/convert_worker.py` | 隔离的转换子进程：组合装配、摘要 CSS 内联、`[STEP i/n]` 进度行 |
| `src/converter.py` | 唯一转换核心：Word 会话、摘要提取、HTML 清洗、三种模式差异 |
| `src/chart_relink.py` | 正文导出前把本地图表的数据源改写为外链到不存在的工作簿（纯 zip/XML），Word 导出时不再逐张等待 Excel |
| `src/table_image.py`、`valuation_table_image.py`、`toc_image.py` | 表格、估值表、目录渲染成图片 |
| `src/configs/*.json` | 摘要、深度正文、周报三套版式配置 |
| `src/web/index.html`、`result.html` | 工作台首页；结果页（预览、复制、进度） |
| `src/workbench.spec`、`build.ps1` | PyInstaller 单文件构建 |

## 不变量

- `converter.py` 是唯一转换核心，保留三套模式配置，不另建模式引擎。
- 摘要在前，正文或周报在后，合成一份 HTML；正文与周报互斥。
- 每份输入用本程序自己的 Word 实例，禁用宏，用完关闭；绝不结束用户已打开的 Word。
- 源 Word 文件只读，转换在私有临时副本上进行。
- 正文导出筛选 HTML 用的是图表外链改写后的副本（开关 `document.relink_local_charts`，body 与 weekly 配置）：
  只改图表的数据源关系，缓存值与格式不动；命中守卫的图表或文档、改写出错时用原副本，改写副本导出失败时
  丢弃这次导出并用原副本重导（Word 失效时先重建实例）。摘要、目录与样式探测始终用原副本。
- 不影响成品的辅助步骤（关闭文档、退出 Word 后的收尾、临时目录删除、自检、元数据、目录图）出错只打印 `[warn]` 并继续；
  正文单块渲染出错只留空行并报警（Word 页码、章节，图表另报序号），块内部分失败（表格单元格图片、图框里贴入的表）留空并同样报警，出口对记过问题却未告警的块补告警、删去未内嵌的本机图片，摘要 + 正文/周报时摘要取不到只交付正文并报警；
  只转摘要时摘要取不到、正文或周报整体导出失败、成品为空时必须失败。
- 结果页控件与提示不进入复制内容；复制标题只给纯文本。
- 生成标记必须符合公众号《内容排版规范》：改动 `converter.py`、图片辅助模块、
  `convert_worker.py` 或 `result.html` 的复制逻辑前后，对照
  `docs/wechat-mp-official/checklist.md`，并对成品运行
  `python tests/wechat_compliance_check.py <成品.html>`（FAIL 须为 0）。
  字体只声明平台默认字体栈（`converter.WECHAT_FONT_FAMILY`）。

## 维护包里能做的检查

导出包不含测试套件、样本报告和已验收快照（它们含私有报告内容，只在开发工作区）。
包内可做的检查：

- 用 `python src/convert_worker.py --source <报告.docx> --modes summary,body --toc --output out.html`
  转一份自己的报告，再运行 `python tests/wechat_compliance_check.py out.html`，FAIL 须为 0。
- 在结果页点「复制正文」，把剪贴板 HTML 原样存成文件（不外套 `<body>`），同样跑检查脚本。
- `checklist.md` 里提到的 unittest、real-parity 快照比对、`HANDOFF.md` 属于开发工作区的完整门禁，
  导出包里不可用。

## 运行与构建（PowerShell 7，在导出目录根下执行）

```powershell
python -m pip install -r src/requirements.txt
python src/app.py
pwsh -NoLogo -NoProfile -File src/build.ps1 -InstallDependencies
```

构建产物按 `APP_VERSION` 命名：`v1.1.0` → `报告转公众号v1_1.exe`，`v1.1.0 beta 2` →
`报告转公众号v1_1_beta2.exe`，写到 `release/` 并复制到根目录，旧版本文件不覆盖。

## 版本

- 版本号唯一来源是 `src/app.py` 的 `APP_VERSION`。
- 功能更新加次版本，修 bug 加补丁版本，架构变化加主版本；高位提升时低位归零。
- 修改轮次以 `vX.Y.Z beta N` 交付验收，验收通过后去掉 beta 后缀再发布。
- 公开分发仓库：`wiwiwiwiwii/dwss-official-account`（MIT），EXE 作为 Release 附件。

## 公众号排版规范

- 权威来源：微信开发者文档「内容排版规范」
  https://developers.weixin.qq.com/doc/subscription/guide/product/plugin_spec.html ，
  其检测规则以官方仓库 https://github.com/wechatjs/verify-article-structure-spec 的
  `verify_article_structure.md` 为准（随包附原文与 MIT 许可）。
- 规范更新时重新抓取两处，逐节比对，同步 `checklist.md` 与 `wechat_compliance_check.py`。
