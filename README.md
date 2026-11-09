# student-archive-forms · 学生档案纸质表生成器

把一份**学生名册**和一份 **Excel 纸质表模板**，自动转换成可直接打印的分班表格：

- **每个班级一个 sheet**，表头、页脚、字体边框全部沿用你的模板
- 按「培养层次 → 学院 → 班级」分组，可产出 1 / 2 / 9 / 17 个文件，或直接打成两个 zip
- 每个 sheet 都设好 **A4 纵向、四周 2cm、缩放至一页宽、水平居中、每页重复表头**
- 自动按宋体字宽反推列宽，长班名（如 `会计(产业学院)2026-1`）不会被截断
- 自带校验脚本：人数、学号集合、班级数、表头页脚、打印参数逐项核对

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

### 四种分组粒度

| `--group-mode` | 产出 | 典型用途 |
|---|---|---|
| `college-level`（默认） | 每个学院 × 层次一个文件 | 分发给各学院分层次打印 |
| `college` | 每个学院一个文件，本科班在前 | 只按学院分发 |
| `level` | 本科 / 专升本两个文件 | 只按层次汇总 |
| `class` | 单个工作簿，全部班级 sheet | 在线查看 |

加 `--stats` 输出学院 × 层次人数占比；加 `--zip-by-level` 直接产出两个压缩包；加 `--exclude 专升本` 只出本科。

## 作为 Agent Skill 使用

把整个目录放到 skills 目录即可，例如 WorkBuddy 的 `~/.workbuddy/skills/student-archive-forms/`。
`SKILL.md` 里写了触发条件、工作流和关键约定；Agent 说一句「按学院分层次做档案表」就会走这套流程。
`references/environment.md` 记录了在 Agent / PowerShell 环境里调用时的注意事项。

## 目录结构

```
SKILL.md                       skill 入口（触发条件 + 工作流）
scripts/generate.py            主生成器（全部参数见 --help）
scripts/verify.py              输出校验器
scripts/config.example.json    可直接试跑的示例配置
references/parameters.md       参数详解、模板结构要求、A4 列宽算法
references/environment.md      环境与编码注意事项
examples/make_demo.py          生成虚构演示数据（不入库）
```

## 打印

用 Excel 打开 → `Ctrl+P` → 打印范围选「**整个工作簿**」，一次打完该文件里所有班级。
选「活动工作表」则只打当前班级。「是否存档」「备注」两列留空供现场手写，页脚「应收/实收/接收人…」保持模板原文。

## 已知限制

- `verify.py` 按默认 7 列布局（序号/姓名/学号/学院/班级/是否存档/备注）核对第 3 列学号、第 6/7 列留空；若换成列数或列序不同的模板，需同步调整校验逻辑。
- 模板数据区若被填了示例学生，脚本按「最大空行间隔」自动切掉示例、只保留页脚（`--footer-gap`，默认 3）。若示例区与页脚之间没有空行，请清空示例或调整该参数。
- `xlrd 2.x` 只支持 `.xls`；`.xlsx` 由 `openpyxl` 读取，两者都需要安装。

## 许可证

[MIT](LICENSE)
