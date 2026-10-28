# shipment-customer-report — 客户出货报告 skill

根据**出货表**（只读取第 2 张工作表「单证部--工厂实际提货数据汇总」）更新**客户出货跟踪 Excel**：你指定开始和结束日期，脚本逐个检查客户表里每个 PO 明细工作表（如 `PO#_260528TP - REORDER` 对应订单 `260528`），把该时间段内 **LAURA / FLEXPORT 各自出了多少货**（数量、比例、体积）写入对应 PO 工作表。`TOTALS` 汇总页完全不动，由它自己的原有公式汇总。

原客户表不会被修改，结果写到一个新的 `.xlsx` 副本；同时生成同名 `.issues.json` 核对清单。

仓库地址：

- CNB（可直接下载）：<https://cnb.cool/Sherlock2026/shipment-customer-report>
- GitHub：<https://github.com/joyziteng/shipment-customer-report>

---

## 一、下载豆包电脑版

Skill / 办公任务功能仅限 **Windows / Mac 电脑客户端**，网页版和手机 App 不支持。

1. 打开豆包官方下载页：<https://www.doubao.com/download>
2. 选择系统（Windows 或 macOS），下载安装，用手机号或抖音账号登录。
3. 参考教程：[豆包 Skill 添加方法（CSDN）](https://blog.csdn.net/weixin_47106289/article/details/163822422)、[豆包办公任务详解（腾讯新闻）](https://new.qq.com/rain/a/20261019A0ALLC00)

## 二、下载 skill

**方式 A：直接下载两个文件（推荐）**，右键另存为：

- 技能说明：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/raw/main/SKILL.md>
- 执行脚本：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/raw/main/scripts/shipment_report.py>

**方式 B：整包下载（tar.gz）**：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/archive/main>

**方式 C：git 克隆**

```bash
git clone https://cnb.cool/Sherlock2026/shipment-customer-report.git
# 或 GitHub：git clone https://github.com/joyziteng/shipment-customer-report.git
```

**安装位置**（供本地 AI 编码工具自动识别）：

```text
~/.agents/skills/shipment-customer-report/
├── SKILL.md
└── scripts/
    └── shipment_report.py
```

**安装 Python 依赖**（需要 Python 3.9+）：

```bash
python3 -m pip install --user openpyxl
```

如果系统不允许全局安装，用虚拟环境：

```bash
python3 -m venv ~/.agents/skills/shipment-customer-report/.venv
~/.agents/skills/shipment-customer-report/.venv/bin/pip install openpyxl
```

## 三、在豆包「办公任务」中新增技能

1. 打开豆包电脑客户端，把输入框上方的模式切换为 **【工作任务】**。
2. 新增技能，任选一种：
   - **创建自定义技能**：输入框发送 `/创建技能`，在表单里填技能名称（如「客户出货报告」）和描述——可把 `SKILL.md` 的内容粘贴进去，点【保存】。
   - **先跑一遍再沉淀**：先用第四节提示词让豆包完整执行一次任务，满意后发送「把本次执行规则保存为技能，技能名称：客户出货报告」。
   - **技能中心**：左侧栏「技能·连接器·伙伴」→ 搜索或按分类浏览 → 点技能右侧「+」添加到「我的技能」。
3. 之后用 `调用【客户出货报告】，任务内容：…` 或 `/客户出货报告 …` 触发。

## 四、使用提示词

把下面这段发给豆包（或任何支持该 skill 的 AI 工具），替换尖括号内容：

> 请使用 skill `shipment-customer-report`，根据我提供的客户表和出货表生成新的客户出货报告。只读取出货表第 2 张“单证部--工厂实际提货数据汇总”工作表；统计 **2026-10-12 到 2026-10-16** 这段时间，把每个 PO 明细工作表在客户表里对应订单的出货数据，按货代填写到该时间段的列。只更新 PO 明细工作表，`TOTALS` 不要动。客户表只生成副本，不覆盖原文件。完成后列出待确认事项。
>
> 客户表：`<例如 /Users/xxx/Desktop/客户表.xlsx>`
> 出货表：`<例如 /Users/xxx/Desktop/出货表.xlsx>`
> 开始日期：`<例如 2026-10-12>`
> 结束日期：`<例如 2026-10-16>`
> 输出位置：`<例如 /Users/xxx/Desktop/客户出货报告.xlsx>`

也可以说“统计过去 7 天”，此时结束日是昨天。直接命令行运行：

```bash
python3 ~/.agents/skills/shipment-customer-report/scripts/shipment_report.py \
  --customer '/path/客户表.xlsx' --shipment '/path/出货表.xlsx' \
  --output '/path/客户报告副本.xlsx' \
  --start 2026-10-12 --end 2026-10-16 --historical-year 2026
```

`--historical-year` 只用于客户表里 `BEFORE 8/21` 这种没写年份的历史表头；脚本不会猜年份，必须由你确认。

## 五、结果怎么读

- **报告 xlsx**：只修改 PO 明细工作表。已有相同日期区间时更新原列（数量 / 比例 / 体积），原比例公式保留；没有该区间时在页面最右侧追加两组列（LAURA、FLEXPORT 各一组），表头格式如 `10/24-10/26`。
- **`.issues.json` 核对清单**：按 PO + 货代列出源表行号、数量、体积、已归集和未归集数量，另外列出备注行、货代无法识别、颜色无法唯一匹配、体积为零、无客户明细页的订单等。**清单未清零前，报告按草稿对待。**

## 六、重要注意事项

- **`TOTALS` 不改**：已有日期区间通常能由 `TOTALS` 原有公式自动汇总（用 Excel 打开或重算后生效）；新增日期列不会自动出现在 `TOTALS`，`TOTALS` 里原本的常量、断链或错引用也不会自动修正。
- 匹配规则：出货表 G 列外合同 = 客户行 A 的完整 PO（保留 `CA`、`UK`、连字符，去掉开头 `PO` 和结尾 `TP`）；E 列款号忽略结尾 `TP`；F 列描述中按完整颜色词匹配，`CHESTNUT`/`CHESNUT` 同义。
- 日期用 N 列工厂提货日，数量用 K 列，体积用 S 列，货代用 AF 列。AF 只认 `LAURA`、`FLEXPORT`；写成“预计 FLEXPORT”“0”或备注文字的行不猜货代，保留原值并提示。
- Y 列备注不再整行跳过：按表中数量和体积统计，同时在清单里列出备注、行号和原值，供人工核对（例如“实际发 1554，待核实”）。
- 如果指定区间与客户表某页已有区间部分重叠，脚本会中止并提示，避免重复统计。如果你确认要按本次区间重新统计，可加 `--allow-overlap`：原列保持不动，本次区间作为新列追加（交叠日期会同时出现在两组列里，报告会提示）。
- 客户表里同一 PO 有可见页和隐藏页时，两页各自独立更新；核对清单按来源行只统计一次，不会把两页相加。
- 数量或体积一项缺失时，另一项仍可写入；缺失项保留原值并在清单中提示。
- 原始文件不会修改；输出文件名不能与输入或已有文件重复。
