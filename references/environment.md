# Windows 环境与数据注意事项

在 Windows / Agent 环境里调用 `generate.py`、`verify.py` 时的已知坑。改参数见 `references/parameters.md`，改界面见 `references/webapp.md`。

## 目录

- 依赖
- 读报告，不要读控制台
- 编码坑
- 删除文件的安全守卫
- 名册与模板的常见形态
- 校验要点

## 依赖

Python 3.9+。直接安装三个依赖（仓库根目录另附 `requirements.txt`）：

```bash
python -m pip install "pandas>=1.5" "openpyxl>=3.1" "xlrd>=2.0"
```

| 包 | 用途 |
|---|---|
| `pandas` | 读名册、分组聚合 |
| `openpyxl` | 写 `.xlsx`、设置样式与打印参数 |
| `xlrd` | 读 `.xls`（2.x **只支持 .xls**，不支持 .xlsx） |

`.xls` 必须由 `xlrd` 读，`.xlsx` 由 `openpyxl` 读；缺任一依赖会在读文件时报错，先装齐再排查数据。

## 读报告，不要读控制台

某些 Agent / PowerShell 工具执行后只返回退出码，不回显 stdout。**让脚本自己写 UTF-8 报告文件**，再读文件：`generate.py` / `verify.py` 已内置 `Report`，默认写到输出目录的 `_生成报告.txt` / `_校验报告.txt`，也可用 `--log` 指定。

不要依赖 `print` 的回显判断成败，一律看报告里的「全部校验通过…异常项 0」。

## 编码坑

不要用 PowerShell 管道转存 Python 的 stdout：PowerShell 会把 Python 输出的 UTF-8 字节按 GBK 解码再重新编码，中文会变成「鍓湰瀛︾敓」类乱码。脚本内部用 `io.open(path, "w", encoding="utf-8")` 写文件才安全。

配置文件、报告、JSON 一律 UTF-8；Windows 记事本另存时留意编码。

## 删除文件的安全守卫

部分环境会拦截直接删除（`Remove-Item` 报 `SAFE_DELETE_FAIL_CLOSED`，`Add-Type` 加载 .NET 被拦）。不要用 `os.remove` 绕过守卫。可行做法：

- 用 `Move-Item` / `Rename-Item` 把旧文件挪到备份目录，提示用户手动清理；
- 或调用系统回收站 API（Windows 需显式允许）；
- 删除前先打一份 zip 备份并校验完整性。

## 名册与模板的常见形态

- 教务系统导出的名册常在第 1 行放大标题、第 2 行留空、**第 3 行才是表头**，所以 `--header-row` 默认 3。列名与默认不同时用 `--column-map` 覆盖，不要改代码。
- 「培养层次」常只有 `普通本科` / `专升本`，**没有严格意义的专科**；用户口语里的「专科」先确认是否指专升本，再用 `--level-map` 决定显示名。
- 一个班级通常只属于一个学院、一个培养层次；按 层次 + 学院 + 班级 分组不会串组。若名册违反这一点，先核对数据。
- 模板的数据区有时被填入若干示例学生（只有序号和姓名），页脚仍在更下方。脚本按「最大空行间隔」自动跳过示例、只搬页脚（`--footer-gap`，默认 3）；换模板后看报告里的「模板页脚单元格数」是否符合预期。
- 长班名（如含「产业学院」「专升本」）比普通班名宽近一倍，列宽必须按全量数据最长字符串反推。

## 校验要点

`verify.py` 会检查：

- 输出学生行数 == 源表行数；学号集合零缺失、零多余、零重复
- sheet 总数 == 源表班级数
- 每个 sheet 表头一致、序号连续、「是否存档 / 备注」为空、页脚紧随末位学生
- A4 纵向、`fitToWidth=1`、`fitToPage=True`、四周边距等于 `--margin-cm`、`print_title_rows` 已设、`print_area` 已设
- `print_title_rows` 用 openpyxl 读回是 `$1:$1`（写入时是 `1:1`），比较时要去掉 `$`

用过筛选时输出目录会多出 `_筛选条件.json` 与 `_筛选信息/筛选后名册.xlsx`，校验器会自动改按这份子集核对；想强制按全量核对加 `--ignore-filter`。
