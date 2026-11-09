# 示例数据

`make_demo.py` 会现场生成两个文件，供你试跑整套流程：

| 文件 | 内容 |
|---|---|
| `示例学生名册.xlsx` | 21 名**虚构**学生、5 个班、2 个学院 × 2 个培养层次；第 1 行标题、第 3 行表头 |
| `示例档案表模板.xlsx` | 7 列表头、空白数据区（带细边框）、D10:D15 六行页脚 |

```bash
python examples/make_demo.py
python scripts/generate.py --config scripts/config.example.json
```

所有姓名、学号、学院、班级都是编造的，**不含任何真实学生信息**。
这两个 `.xlsx` 被 `.gitignore` 忽略，不会进入仓库；也请不要把真实名册放进这个目录。
