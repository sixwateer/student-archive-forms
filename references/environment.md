# 环境与编码注意事项

## 依赖

Python 3.9+，依赖见仓库根目录的 `requirements.txt`：

```bash
python -m pip install -r requirements.txt
```

| 包 | 用途 |
|---|---|
| `pandas` | 读名册、分组聚合 |
| `openpyxl` | 读写 `.xlsx`、设置打印版式 |
| `xlrd>=2.0` | 读老式 `.xls`（**xlrd 2.x 不支持 .xlsx**，.xlsx 走 openpyxl） |

## 中文路径与控制台编码

- 路径含中文没问题，配置文件用 UTF-8 保存即可。
- Windows 控制台默认可能是 GBK，直接 `print` 中文可能乱码或抛 `UnicodeEncodeError`。两个脚本内置 `Report` 类，会把完整报告**同时写成 UTF-8 文件**（默认 `<out>/_生成报告.txt`、`_校验报告.txt`）。以报告文件为准，不要只看控制台回显。
- 在 PowerShell 里不要用 `| Out-File` 转存 Python 的 stdout：PowerShell 会按 GBK 解码再编码，中文会变成乱码。让脚本自己用 `io.open(..., encoding="utf-8")` 写文件。
- 配置文件里的相对路径**相对于当前工作目录**，请从仓库根目录运行命令。

## 页边距单位

openpyxl 的 `page_margins` 用**英寸**，2cm = `2 × 0.3937 = 0.7874`。
`fitToWidth` 必须先设置 `ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)` 才会生效，否则 Excel 会忽略。
`print_title_rows` 写入 `1:1`，读回来是 `$1:$1`，比较时要去掉 `$`。

## 在 Agent / WorkBuddy 环境里调用（可选阅读）

以下只适用于由 AI Agent 在 WorkBuddy 会话内调用脚本的场景：

- 某些会话里 PowerShell 工具不回显 stdout，只返回退出码 → 一律读脚本生成的报告文件。
- 删除文件可能被 `[safe-delete]` 安全守卫拦截（回收站调用失败）。不要用 Python `os.remove` 绕过；改用 `Move-Item` / `Rename-Item` 把旧文件挪到 `_待删除-…` 目录，提示用户手动清理，删除前先打一份 zip 备份并校验完整性。
- 这类环境里 `Bash` 工具可能异常，优先用 PowerShell。

## 模板与校验的耦合

`verify.py` 假定默认 7 列布局，检查第 3 列学号、第 6/7 列留空、页脚紧随末位学生。
换模板时若列数或列序变了，需要同步调整校验逻辑；生成脚本本身是按表头文字映射的，不受影响。
