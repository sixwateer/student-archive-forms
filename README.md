# student-archive-forms · 学生档案纸质表生成器

**只想用、不想装环境**：直接打开在线网页版 <https://sixwateer.github.io/student-archive-forms/>（单文件离线运行，名册不上传）；或下载仓库 Release 里的单文件 HTML / exe 交付包。

把一份**学生名册**和一份 **Excel 纸质表模板**，自动转换成可直接打印的分班表格：

- **每个班级一个 sheet**，表头、页脚、字体边框全部沿用你的模板
- 按「培养层次 → 学院 → 班级」分组，可产出 1 / 2 / 9 / 17 个文件，或直接打成两个 zip
- 每个 sheet 都设好 **A4 纵向、四周 2cm、缩放至一页宽、水平居中、每页重复表头**
- 自动按宋体字宽反推列宽，长班名（如 `会计(产业学院)2026-1`）不会被截断
- **表格线一定看得见**：表头 + 学生行统一写「纯黑细实线」，签名区（应收/实收/接收人…）留白不带框；不依赖模板自带的 `auto` 边框（那种在部分 Excel/WPS 和打印驱动下会消失）
- **先筛选再生成**：按年级、本科/专升本/一起、学院、班级任意组合挑人，不再一股脑全出
- **成果可一键合并**：散落的十几个学院表、或「本科(9学院).zip + 专升本(8学院).zip」，一条命令合成一个总 ZIP（附合并清单 + 完整性校验），方便整体转发
- 自带校验脚本：人数、学号集合、班级数、表头页脚、打印参数逐项核对；筛选时自动按同一批学生核对
- **学院 → 班级严格联动**：勾选某学院后班级列表只出现该院班级，说明文字实时写出当前学院、层次、年级与班级数，不再混入其他学院或残留旧文案
- **QQ 浏览器也能下载**：自动识别 QQ / 微信 / 手机内核，改用 `data URI` + 醒目的手动保存按钮（附备用链接与重试），避免「0% 下载失败」和 `UUID.bin`；桌面 Edge / Chrome 的自动下载行为不变

既是一套独立命令行工具，也是一个 [Agent Skills](https://agentskills.io) 格式的 skill（`SKILL.md`）。

## 隐私说明

**本仓库不包含任何真实学生数据。** `examples/` 里的姓名、学号、学院全是虚构的，由脚本现场生成。
`.gitignore` 默认忽略所有 `.xls/.xlsx/.zip` 与 `output/`，避免误把真实名册提交上去。
使用真实数据时，请把名册放在仓库外，或至少确认它已被 `.gitignore` 覆盖。

## 快速开始（3 条命令跑通）

需要 Python 3.9+。**请始终在仓库根目录执行**，因为示例配置里的路径是相对当前目录的。

```bash
python -m pip install -r requirements.txt
python examples/make_demo.py                                  # 生成 21 名虚构学生 / 5 个班 + 模板
python scripts/generate.py --config scripts/config.example.json
python scripts/verify.py \
  --source ./examples/示例学生名册.xlsx \
  --template ./examples/示例档案表模板.xlsx \
  --out ./output --include-zip
```

生成结果在 `output/`：

```
示例档案纸质表格-示例理工学院-本科.xlsx        (2 个班级 sheet)
示例档案纸质表格-示例理工学院-专升本.xlsx      (1 个班级 sheet)
示例档案纸质表格-示例商学院-本科.xlsx          (1 个班级 sheet)
示例档案纸质表格-示例商学院-专升本.xlsx        (1 个班级 sheet)
示例档案纸质表格-本科(2个文件).zip
示例档案纸质表格-专升本(2个文件).zip
_生成报告.txt / _校验报告.txt                  (UTF-8)
```

校验报告出现「**全部校验通过…异常项 0**」才算完成。

## 用自己的数据

把 `scripts/config.example.json` 复制一份，改三个路径，其余按需调整：

```json
{
  "source": "D:/data/学生基本信息.xls",
  "template": "D:/data/档案纸质表格模板.xlsx",
  "out": "D:/data/输出",
  "header_row": 3,
  "group_mode": "college-level",
  "column_map": "",
  "level_map": "普通本科=本科,专升本=专升本"
}
```

```bash
python scripts/generate.py --config 我的配置.json
python scripts/verify.py --source <名册> --template <模板> --out <输出目录>
```

- 教务系统导出的名册常在第 1 行放大标题、第 2 行留空、**第 3 行才是表头**，所以 `header_row` 默认 3。
- 列名不同时**不用改代码**，用 `--column-map` 覆盖，例如 `--column-map '{"所属学院":"院系名称"}'`。
- 用户口语里的「专科」常常其实指「专升本」，用 `--level-map` 决定文件里显示成什么名字。
- 完整参数表见 [`references/parameters.md`](references/parameters.md)。

### 先筛选，再生成

不想一次把全校都倒出来，就先用 `--list-filters` 看看名册里有什么，再挑条件：

```bash
# 1) 列出可筛选的年级、层次、学院、班级及各有多少人（不生成文件）
python scripts/generate.py --source 学生基本信息.xls --header-row 3 --list-filters

# 2) 只要 2025 级 + 本科
python scripts/generate.py --config 我的配置.json --grade 2025 --level 本科

# 3) 只要某学院的专升本，或只要某几个班
python scripts/generate.py --config 我的配置.json --college 商学院 --level 专升本
python scripts/generate.py --config 我的配置.json --class '会计本2026-1,会计本2026-2'

# 4) 两个层次都要（默认），但排除若干学号
python scripts/generate.py --config 我的配置.json --exclude-ids @不打印的学号.txt
```

| 参数 | 作用 |
|---|---|
| `--grade 2025` / `2025级` | 按年级筛（**入学年份优先**，缺失时用「当前所在级 / 年级」兜底），并自动把 `2025级` 写进文件名 |
| `--level 本科` / `专升本` / `本科,专升本` | 只要某个层次；留空 = 两个层次一起出 |
| `--college 商学院,法学院` | 只要这些学院 |
| `--class 会计本2026-1` | 只要这些班级；也支持 JSON 数组 `["…","…"]` |
| `--grade-column` | 年级列名特殊时手动指定，可逗号分隔写多列 |
| `--exclude-ids` | 排除个别学号，逗号分隔或 `@文本文件`（每行一个） |
| `--list-filters` | 只列出可筛选项，不生成文件 |

条件可以任意叠加。年级取自名册的年级列；**名册没有年级列时**，`--grade` 只用于命名，不筛数据（报告里会提示）。`--level` 用的是 `--level-map` 映射后的**显示名**（如「本科」），写原始值「普通本科」也能识别。

筛选后脚本会在输出目录写入 `_筛选条件.json` 和 `_筛选信息/筛选后名册.xlsx`，**`verify.py` 会自动改按这份子集核对**，所以校验命令不用改；想强制按全量名册校验就加 `--ignore-filter`。

> 上面的参数也可以直接写进 `--config` 的 JSON：`grade`、`level`、`college`、`class`、`exclude_ids`（键名见 `scripts/config.example.json`）。

### 四种分组粒度

| `--group-mode` | 产出 | 典型用途 |
|---|---|---|
| `college-level`（默认） | 每个学院 × 层次一个文件 | 分发给各学院分层次打印 |
| `college` | 每个学院一个文件，本科班在前 | 只按学院分发 |
| `level` | 本科 / 专升本两个文件 | 只按层次汇总 |
| `class` | 单个工作簿，全部班级 sheet | 在线查看 |

加 `--stats` 输出学院 × 层次人数占比；加 `--zip-by-level` 直接产出两个压缩包；加 `--exclude 专升本` 只出本科。

## 合并分发（合成一个总 ZIP）

按学院分层次生成后，成果通常是十几个 `.xlsx`，或两个 zip（本科 9 个学院、专升本 8 个学院）。整体发给别人前先合成一个：

```bash
# 直接合并整个输出目录
python scripts/combine.py --in ./output --out 档案表成果-合并全部.zip

# 或者喂已有的 zip / 单个 xlsx（可多次 --in，混合都行）
python scripts/combine.py \
  --in "本科(9个学院).zip" --in "专升本(8个学院).zip" \
  --out 档案表成果-合并全部.zip
```

- 输入可以是 **文件夹 / `.zip` / `.xlsx` / `.txt` 任意混合**；zip 里的表格会被展平收进总包，不会嵌套一层 zip
- **同名自动去重**：不同来源的同名学生表会加「来源__」前缀，不会互相覆盖
- 自动写 `_合并清单.txt`（记录每个文件来自哪里、共多少表格），并做 `testzip` 完整性校验
- `--include-history` 一并收进历史批次；`--no-reports` 不收 `_生成报告.txt` 这类报告

参数全见 `python scripts/combine.py --help`。

## 作为 Agent Skill 使用

本仓库同时是一个符合 Agent Skills 规范的 skill 包：`SKILL.md` 只保留核心流程，细节按需从 `references/` 加载（渐进披露）。把整个目录放到 skills 目录即可，例如 WorkBuddy 的 `~/.workbuddy/skills/student-archive-forms/`；也可直接分发 `student-archive-forms.skill`（由 skill-creator 校验并打包的 zip 格式）。

- 调参数前读 `references/parameters.md`
- Windows 下跑脚本前读 `references/environment.md`
- 改界面、排查下载或重新打包前读 `references/webapp.md`
- 要交付给非技术同事时，复制 `assets/操作指引.md`

## 目录结构

```
SKILL.md                       skill 入口（name/description + 核心工作流）
scripts/generate.py            主生成器（全部参数见 --help）
scripts/verify.py              输出校验器
scripts/combine.py             把散落成果合并成一个总 ZIP
scripts/config.example.json    可直接试跑的示例配置（虚构数据路径）
references/parameters.md       参数详解、模板结构要求、A4 列宽算法
references/environment.md      Windows 环境与编码注意事项
references/webapp.md           网页版 / exe 小程序的交互约定、下载兼容与构建流程
assets/操作指引.md             面向非技术用户的一页纸操作指引
examples/make_demo.py          生成虚构演示数据（生成的 .xlsx 不入库）
```

## 打印

用 Excel 打开 → `Ctrl+P` → 打印范围选「**整个工作簿**」，一次打完该文件里所有班级。
选「活动工作表」则只打当前班级。「是否存档」「备注」两列留空供现场手写，页脚「应收/实收/接收人…」保持模板原文。

## 已知限制

- `verify.py` 按默认 7 列布局（序号/姓名/学号/学院/班级/是否存档/备注）核对第 3 列学号、第 6/7 列留空；若换成列数或列序不同的模板，需同步调整校验逻辑。
- 模板数据区若被填了示例学生，脚本按「最大空行间隔」自动切掉示例、只保留页脚（`--footer-gap`，默认 3）。若示例区与页脚之间没有空行，请清空示例或调整该参数。
- `xlrd 2.x` 只支持 `.xls`；`.xlsx` 由 `openpyxl` 读取，两者都需要安装。
- 用过筛选后，输出目录会多出 `_筛选条件.json` 与 `_筛选信息/筛选后名册.xlsx`；它们是给校验器用的中间文件，**不是成果表**，也不会被打进 zip。换一批条件重跑时会被自动覆盖。

## 许可证

[MIT](LICENSE)
