# shipment-customer-report — 客户出货报告 skill

根据**出货表**（只读取第 2 张工作表「单证部--工厂实际提货数据汇总」）自动更新**客户出货跟踪 Excel**：按 PO、款号、颜色、货代和工厂提货日匹配出货记录，在客户表副本中新增「过去 N 天」的出货数量、比例、体积列，并生成一份待确认清单（`.issues.json`）。原表不会被修改。

仓库同时发布在：

- GitHub：<https://github.com/joyziteng/shipment-customer-report>
- CNB：<https://cnb.cool/Sherlock2026/shipment-customer-report>

---

## 一、下载豆包电脑版

本教程使用豆包电脑客户端（Skill / 办公任务功能仅限 Windows / Mac 客户端，网页版和手机 App 不支持）。

1. 打开豆包官方下载页：<https://www.doubao.com/download>
2. 选择你的系统（Windows 或 macOS），下载并安装。
3. 用手机号或抖音账号登录。

> 参考教程：[豆包 Skill 添加方法（CSDN）](https://blog.csdn.net/weixin_47106289/article/details/163822422)、[豆包办公任务详解（腾讯新闻）](https://new.qq.com/rain/a/20261019A0ALLC00)

## 二、下载 skill

**方式 A：直接下载两个文件（推荐，最简单）**

在浏览器打开下面的链接，右键另存为即可：

- 技能说明：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/raw/main/SKILL.md>
- 执行脚本：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/raw/main/scripts/shipment_report.py>

**方式 B：整包下载**

- 打包下载（tar.gz）：<https://cnb.cool/Sherlock2026/shipment-customer-report/-/git/archive/main>

**方式 C：git 克隆**

```bash
# CNB
git clone https://cnb.cool/Sherlock2026/shipment-customer-report.git
# 或 GitHub
git clone https://github.com/joyziteng/shipment-customer-report.git
```

**安装位置**（供 ZCode、Claude Code 等本地 AI 编码工具自动识别）：

```text
~/.agents/skills/shipment-customer-report/
├── SKILL.md
└── scripts/
    └── shipment_report.py
```

**安装 Python 依赖**（电脑需要 Python 3.8+）：

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
   - **创建自定义技能**：在输入框发送 `/创建技能`，在表单中填写技能名称和描述。可打开上面下载的 `SKILL.md`，把 frontmatter 里的描述和正文规则粘贴进去，点【保存】。
   - **先跑一遍再沉淀**：先用第四节提示词让豆包完整执行一次任务，满意后发送「把本次执行规则保存为技能，技能名称：客户出货报告」。
   - **技能中心**：左侧栏「技能·连接器·伙伴」→ 搜索或按分类浏览 → 点技能右侧「+」添加到「我的技能」。
3. 之后在【工作任务】里用 `调用【客户出货报告】，任务内容：…` 或 `/客户出货报告 …` 触发。

## 四、使用提示词

把下面这段发给豆包（或任何支持该 skill 的 AI 工具），替换尖括号内容即可：

> 请使用 skill `shipment-customer-report`，根据我提供的客户表和出货表生成更新后的客户出货报告。只读取出货表第 2 张“单证部--工厂实际提货数据汇总”工作表；统计【过去 N 天】的记录，日期以昨天为结束日。客户表只生成副本，不覆盖原文件。完成后提供报告和待确认清单，并列出需要我补充或核对的项目。
>
> 客户表：`<例如 /Users/xxx/Desktop/客户表.xlsx>`
> 出货表：`<例如 /Users/xxx/Desktop/出货表.xlsx>`
> 过去天数：`<例如 7>`
> 输出位置：`<例如 /Users/xxx/Desktop/客户出货报告.xlsx>`

也可以直接命令行运行（不经过豆包）：

```bash
python3 ~/.agents/skills/shipment-customer-report/scripts/shipment_report.py \
  --customer '/path/客户表.xlsx' --shipment '/path/出货表.xlsx' \
  --output '/path/客户报告副本.xlsx' --days 7
```

## 五、结果怎么读

每次运行产出两个文件：

- **报告 xlsx**：含 `TOTALS` 汇总页和每个 PO 一个明细页。历史日期区间列全部保留，本次区间的新列追加在最右侧（每个货代一组：数量、比例、体积）。
- **`.issues.json` 待确认清单**：记录需要人工核对的事项（新 PO 缺基础资料、提货日期为空、无法唯一匹配、疑似超发、原表公式错误等）。**清单未清零前，报告按草稿对待。**

## 六、重要注意事项

- 客户表历史日期表头（如 `BEFORE 8/21`）**没有年份**时，脚本会主动停止，要求你确认年份后加参数 `--historical-year 2026`（替换为确认的年份）重跑，绝不自行猜测。
- 检测到本次区间与任何已有历史区间重叠时会**中止**，避免重复统计。
- 脚本不计算公式；用 Excel 打开报告时会自动重算。原客户表如已有公式错误，会被列进待确认清单，不会被悄悄修改。
- 匹配规则：出货表 G 列外合同 = PO（保留 CA、UK 前缀，忽略末尾 TP）；E 列款号忽略末尾 TP；在产品描述中匹配颜色，`CHESTNUT`/`CHESNUT` 视为同一颜色。货代目前只自动处理 LAURA 和 FLEXPORT。
- 两份源文件均为只读，输出文件不能与输入或已有文件同名。
