# 公众号排版规范对照清单

改动生成标记（`src/converter.py`、图片辅助模块、`src/convert_worker.py`）或复制逻辑
（`src/web/result.html`）前后，逐条对照本清单。规则编号 C1~C17 与
`tests/wechat_compliance_check.py` 的输出一致。

## 官方文档

| 文件 | 来源 | 版本 |
| --- | --- | --- |
| `plugin_spec.md` | 微信开发者文档「内容排版规范」 https://developers.weixin.qq.com/doc/subscription/guide/product/plugin_spec.html | 2026-10-20 抓取 |
| `verify_article_structure.md` | 官方检测仓库 https://github.com/wechatjs/verify-article-structure-spec 的规则定义 | main 提交 `0fa8ad2`（2026-10-19），2026-10-20 抓取 |
| `LICENSE-verify-article-structure-spec.txt` | 上述仓库的 MIT 许可（腾讯） | 同上 |

开发者文档页声明：检测规则以仓库内 `verify_article_structure.md` 为权威定义。两份文本
有三处不同，判定以仓库版为准：

- 2.1 嵌套：文档页写「同标签、同样式、单子节点的连续嵌套不超过 10 层」；仓库版改为
  从第 15 层起检查冗余嵌套。本清单按更严的 10 层机检。
- 1.4.3 图片宽度：仓库版写明带 `data-w` 的图片不等待加载，其余最长等待 5 秒。
- 1.8 pre：仓库版允许带横向滚动或 `white-space:pre-wrap` 不溢出的代码块。

更新方法：重新抓取上面两个地址，覆盖两份文件并改文件头的日期与提交号；逐节比对差异，
同步本清单与 `tests/wechat_compliance_check.py`。

## 两种检查对象

- 文件视图：转换写出的成品 HTML（`tests/real-parity/<样本>/new.html`）。
- 复制视图：结果页「复制正文」时，`copyFragment` 把 `<body>` 换成 `<section>`，并把每个
  元素的 18 项计算文字样式写进内联样式，其中包括 `font-family` 和 `text-align`。所以
  成品中任何一个元素继承到的字体或对齐不合规，粘贴进编辑器后每个元素都会带上违规值。

## 清单

「机检」列：FAIL 为违规，WARN 为需浏览器实测，NOTE 为提示；「人工」表示脚本查不到。

| 编号 | 规范条目 | 可检查项 | 本项目做法 | 机检 |
| --- | --- | --- | --- | --- |
| C1 | 1.1 opacity | `img` 不设 `opacity:0`，不用透明图叠 SVG 背景图 | 不使用 opacity | FAIL |
| C2 | 1.2 caret-color | 不设透明光标色 | 不使用 | FAIL |
| C3 | 1.3 line-height | 含文字且可能多行的元素，行高不小于字号；纯图片容器、单行文字除外 | 正文 2.0 倍，标题 1.5~1.6 倍，周报标题 30px/18px；图片段落 `line-height:1` 不含文字 | FAIL（静态估计，折行与否需浏览器确认） |
| C4 | 1.4 width | 非图片元素不写固定 px 宽；左偏移不致窄屏溢出；有意固定宽度时加 `data-ignore-width` | 只有周报一级标题背景框 270px、文字框 200px，均带 `max-width`，背景框在外层居中 | WARN |
| C5 | 1.4.3 data-w | 建议 `img` 带原始像素宽 `data-w` | 未写；图片全部是 data: URI，即时加载，不触发超时误报 | NOTE |
| C6 | 1.5.1 height:0、1.5.2 内容溢出 | 含文字的容器不设 `height:0`；固定高度加 `overflow:hidden` 时内容不得超出 | 周报标题背景框用 `min-height`，长标题折行时框随内容撑高 | FAIL（height:0）/ WARN（固定高度裁剪） |
| C7 | 1.6 text-align | 不用 `start`、`end`；复制视图中每个元素的计算对齐都是 left/center/right/justify | `<body>` 显式 `text-align:left`，各块再写自己的对齐 | FAIL（两种视图） |
| C8 | 1.8 pre | 不用 `<pre>` 包段落 | 不生成 `<pre>` | FAIL |
| C9 | 2.1 嵌套、2.2 span[leaf]、2.3 section[nodeleaf] | 同标签同样式单子节点连续嵌套不超过 10 层；不用 leaf/nodeleaf 包块级元素 | `<body>` 以下最深 5 层，相邻层样式各不相同；不使用 leaf 属性 | FAIL |
| C10 | 3 字体 | 内联 `font-family` 只能以平台默认字体栈 `mp-quote, PingFang SC, system-ui, -apple-system` 开头（官方引擎去引号后比前缀）；复制视图中每个含文字元素的计算字体同样如此 | 全部字体键统一为平台默认字体栈 `converter.WECHAT_FONT_FAMILY`，`load_config` 把不合规的配置值换成它 | FAIL（两种视图） |
| C11 | 4.5.2 !important | 样式不用 `!important` | 不使用 | FAIL |
| C12 | 4.1.2 渐变、4.2.2 嵌套顺序、4.5.1/4.6 豁免属性 | 文字背后不用渐变；不用绝对定位或变换打乱结构顺序；`data-no-dark`、`data-ignore-dm`、`data-ignore-width` 只在确认后使用 | 不使用渐变、定位、变换和豁免属性；摘要框等背景写在外层容器 | FAIL（渐变、定位）/ NOTE（豁免属性） |
| C13 | 编辑器不执行脚本（项目规则） | 无 `script`、`style` 样式表、`iframe`、`object`、表单，无 `on*` 事件属性 | 摘要 CSS 由 `convert_worker` 转成内联样式 | FAIL |
| C14 | 资源地址（项目规则） | `src`、`style` 中 `url()` 只用 `data:` 或 `https:`；图片必须内嵌为 base64 | 图片全部内嵌；周报标题背景图由 `convert_worker` 内嵌为 data: URI | FAIL |
| C15 | 标签与属性（项目规则） | 正文只用 `section p br span strong b em i u ul ol li img`；不依赖 class、id | 唯一的 class 是公众号自己的列表类 `list-paddingleft-2` | FAIL（白名单外标签）/ NOTE（class、id） |
| C16 | 4.3.1 图片承载文字、4.3.2 透明底图片、4.3.3 背景图 | PNG 不带透明像素（深色模式底色 #191919 上看不见深色内容）；格式限 png/jpeg/gif；背景图用 data: URI 时粘贴后是否转存未经验证 | 内嵌 PNG 有透明像素时 `flatten_transparent_png` 铺白底；表格与目录做成白底图片是有意取舍（编辑器还原表格很差） | FAIL（透明 PNG）/ WARN（格式、data: 背景图） |
| C17 | 1.7 SVG 动画 begin、4.4.1 SVG | 不内嵌 SVG；动画 begin 同时写 `touchstart; click` | 正文丢弃 SVG | FAIL |
| — | 4.1.1 深色模式对比度 | 文字与背景对比度经深色转换后不低于 1.5 | 正文黑字白底；二三级标题白字 #76A2E9 底（浅色模式对比度约 2.6:1） | 人工（需官方 CLI 的 mp-darkmode 转换） |
| — | 1.3、1.4、1.5.2 的排版实测 | 叠字、多屏宽度差异、高度溢出 | 见 C3、C4、C6 | 人工（需官方 CLI） |

## 检查流程

1. 改动后运行 `python -m unittest discover -s tests -p 'test_*.py'`；其中
   `tests/test_wechat_compliance.py` 覆盖各条机检规则、转换核心的各类块，以及与
   `approved-hashes.json` 一致的 9 份样本快照。
2. 引擎改动后按 `HANDOFF.md` 重转 9 份样本，再运行
   `python tests/wechat_compliance_check.py tests/real-parity/*/new.html`，FAIL 必须为 0；
   WARN 逐条确认。
3. 复制逻辑改动后，用真实结果页复制正文，把剪贴板 HTML 原样存成文件（不要外套 `<body>`），
   同样跑一遍检查脚本：没有 `<body>` 的文件按剪贴板片段检查。
4. 需要权威结论时在本机运行官方 CLI（`npm run check ./article.html`，需 Node.js 与
   puppeteer 自带的浏览器）；只在本机运行，不把报告内容上传到任何外部服务。
