# 零安装交付形态：网页版 / exe 小程序

本文件记录两套「双击即用」形态的交互约定、已知坑、构建与同步步骤。改动界面、排查下载问题或重新打包前读这里。

## 目录

- 两种形态与入口
- 筛选与交互约定（两端必须一致）
- 下载与生成（QQ 浏览器兼容）
- 构建、打包与同步流程
- 回归测试

## 两种形态与入口

| 形态 | 文件 | 说明 |
|---|---|---|
| 单文件网页版 | `04_免安装生成器/学生档案表生成器-网页版.html` | 完全离线、零 CDN，内联 SheetJS / ExcelJS / JSZip；浏览器本地运行，名册不出本机。线上版：<https://sixwateer.github.io/student-archive-forms/>（`gh-pages` 分支） |
| exe 小程序 | `06_构建与测试归档/_exebuild/app_v2.py` | 本地起 `127.0.0.1` 服务，浏览器界面选名册 → 筛选 → 生成；`/combine` + `/download_combine` 提供合并；`/web` 与 `/guide` 分别是网页版备用入口与完整操作指引 |

两者都内置生成 + 合并能力，结果与命令行脚本一致。exe 被安全策略拦住时，用网页版即可。

## 筛选与交互约定（两端必须一致）

- **搜索即定位**：学院、班级各有一个搜索框；输入关键词实时过滤、命中项黄底高亮、首个匹配自动滚动到可视区，并显示「找到 N 个」。
- **一键全选匹配**：「全选匹配 / 清空匹配」只对当前关键词命中的项生效。
- **学院 → 班级作用域**：班级列表只展示**已勾选学院**下的班级；切换学院时重新计算勾选范围与人数，不再混入其他学院班级。
- **年级 / 层次联动**：切换年级或层次时，班级按 `classMatch` 自动勾选、其余置灰（仍可手勾回来），学院人数按 `年级×层次` 实时重算。年级口径统一为「入学年份优先」，见 SKILL.md。
- **说明文字同步**：班级列表下方的说明必须实时写出当前所选学院、层次、年级，以及范围内班级数与已勾选数；任意筛选变化都要刷新，不能残留旧文案。
- **生成前拦截**：未选学院 / 未选班级 / 名册无有效行时提示并中止，不生成空表。
- **名册自检面板**：读入后立即显示 错误 / 警告 / 提示 三级问题（空学号、空姓名、空班级、空学院、重复学号、年级列缺失等）。
- **高级设置预设**：本科·按学院分（`college-level` + 本科）、专升本·按学院分、全部·按层次分（`level`）、全部·只按学院分（`college`）、全部班级·一个文件（`class`）。
- **默认范围**：默认「全部年级 + 全部层次 + 全部学院」，避免往届生被静默漏掉。

## 下载与生成（QQ 浏览器兼容）

纯静态页面没有后端，下载只能靠浏览器能力。QQ / 微信 / 手机内核（含 `content:` 协议）对 `blob:` + 隐藏 `<a download>` 的自动下载支持很差，典型症状是「0% 下载失败」和落成一个 `UUID.bin`。

实现要点（在 `page_template.html` 的 `download()`）：

1. 按扩展名设正确 MIME：`.zip` → `application/zip`，`.xlsx` → `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`；不要用泛化的 `application/octet-stream`。
2. 用 `weakBrowser()` 识别 `MQQBrowser|QQBrowser|MicroMessenger|WeChat|Android|iPhone|iPad|Mobile|HarmonyOS` 与 `content:` 协议。
3. 弱浏览器走兼容模式：把 Blob 用 `FileReader` 转 `data URI`（≤ 2 MB），显示醒目的「点这里保存文件」按钮 + 备用 blob 链接 + 「重新下载」按钮，**不自动触发隐藏下载**（那正是失败项与 `.bin` 的来源），并提示「··· → 用系统浏览器打开」。
4. 桌面 Edge / Chrome 保持原有自动下载，同时提供手动兜底链接。
5. `URL.revokeObjectURL` 只在 `beforeunload` 时执行；过早回收会让异步下载管理器取不到数据。
6. 生成后用 `assertBuffer()` 校验文件非空、ZIP 头为 `PK`，不合法就提示重试而不是给出坏文件。
7. 生成结果与合并结果共用同一个 `download()`，改一处要同时验证两条链路。

> 受浏览器内核限制，静态页无法强制 QQ/微信接受下载；兼容模式需要用户点一下按钮。若仍失败，引导用系统浏览器打开在线版。

## 构建、打包与同步流程

源与产物都在 `06_构建与测试归档/`：

| 步骤 | 命令 / 位置 |
|---|---|
| 改网页版 | 编辑 `_网页版构建/page_template.html`（交互）或 `core.js`（生成算法） |
| 重建单文件 HTML | `python _网页版构建/_build_html.py` → 写出 `04_免安装生成器/学生档案表生成器-网页版.html` |
| 同步交付物 | `python _网页版构建/_sync_dist.py` → 复制 HTML 与操作指引到 `发给别人-.../` 和 `_exebuild/`，重打两个 zip |
| 重打包 exe | `cd _exebuild && python -m PyInstaller 学生档案表生成器.spec --noconfirm`；`.spec` 的 `datas` 需含网页版 HTML、`操作指引.html`、`内置模板.xlsx` |
| 发布网页版 | 把最新 HTML 复制到 `gh-pages` 工作副本的 `index.html` 后提交推送 |

`app_v2.py` 的 `WEB_FALLBACK` / `GUIDE_FILE` 会**优先读 exe 同目录**的文件，所以 `_sync_dist.py` 放了最新副本后，即使 exe 内嵌的是旧资源，运行时也用最新的。

## 回归测试

在 `_网页版构建/` 与 `_exebuild/` 下（需要 `node_modules/puppeteer-core` 与本机 Chrome）：

```bash
cd 06_构建与测试归档/_网页版构建
node _test_ui.js       # 原有 20 项：年级人数联动、搜索定位、自检、预设
node _test_fixes.js    # 31 项：学院→班级过滤、说明同步、QQ UA 兼容下载、合并下载
cd ../_exebuild && NODE_PATH=../_网页版构建/node_modules node _test_mp.js   # 小程序/打包后 exe
```

新增交互或修改下载链路后必须全跑一遍，确保未涉及功能不受影响。
