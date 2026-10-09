# 农夫开票明细 · 正单/返单 自动合并工具

> 一个面向「供应链系统开票明细」的小工具：自动抓取 **正单 / 返单** 数据，按商品编码 **合计对冲**，产出「正单+返单_按品类合计」汇总表。
> 既可作为独立 Python 脚本双击运行，也可作为 WorkBuddy 技能包分发。

---

## ✨ 功能

| 模块 | 脚本 | 作用 |
|---|---|---|
| 网站抓取 | `scrape.py` | 登录供应链系统，逐单抓取「购进查询 / 返厂查询」明细，输出含 `正单`、`返单` 两个 sheet 的 xlsx |
| 按品类合计 | `consolidate.py` | 读取含 `正单`、`返单` 的 xlsx，按商品编码聚合（正单为正、返单转负），产出合计主表 + 明细备查表 |

**核心口径**
- 正单视为正值，返单视为负值（返单数量/金额若为正会自动转负）
- 净数量 = 正单数量 + 返单数量；净额为负表示退货多于进货
- 同一商品编码多行单价不一致时，仅给出「平均含税单价(参考)」

---

## 📁 目录结构

```
fapiao-zhengfan-consolidate/
├─ scrape.py             # 网站自动抓取脚本
├─ consolidate.py        # 按品类合计脚本
├─ run_scrape.bat        # Windows 一键抓取+合计（自动探测 Python / 装依赖）
├─ run.bat               # Windows 拖拽 xlsx 一键合计
├─ config.example.json   # 配置模板（复制为 config.json 后填账号）
├─ SOP.md                # 中文使用说明
├─ SKILL.md              # WorkBuddy 技能定义
├─ 交给对方的使用说明.md  # 分发给他人时的安装指引
├─ README.md
└─ .gitignore            # 已忽略 config.json 与 *.xlsx（防泄密）
```

---

## 🚀 快速开始

### 1. 准备环境（仅一次）
- 安装 Python 3.8+，**务必勾选 `Add Python to PATH`**：https://www.python.org/downloads/
- 依赖会在首次运行时自动安装；也可手动：`pip install requests openpyxl beautifulsoup4`

### 2. 填写账号
复制 `config.example.json` → 重命名为 `config.json`，用记事本打开填入：

```json
{
  "base_url": "http://gyl.zyred.com",
  "user": "你的登录账号",
  "password": "你的登录密码",
  "gysbm": "你的供应商编码",
  "bmbm": "你的商户编码",
  "month": "202609",
  "output": "开票明细_网站抓取.xlsx"
}
```

### 3. 运行
- **推荐**：双击 `run_scrape.bat`，自动抓取并生成 `开票明细_网站抓取.xlsx` 与 `开票明细_网站抓取_按品类合计.xlsx`
- **命令行**：
  ```cmd
  python scrape.py --config config.json
  ```

---

## 💻 命令行参数

```cmd
python scrape.py --user 账号 --password 密码 \
  --gysbm 供应商编码 --bmbm 商户编码 \
  --month 202609 --output 开票明细.xlsx
```
- `--start-date` / `--end-date`（YYYYMMDD）：指定任意日期区间，替代 `--month`
- `--base-url`：覆盖系统地址（默认 `http://gyl.zyred.com`）
- `--no-consolidate`：只抓取、不合计

**仅对已有 xlsx 做合计**（无需联网）：
```cmd
python consolidate.py 你的文件.xlsx
# 或直接把文件拖到 run.bat 上
```

---

## 🔧 配置字段说明

| 字段 | 含义 |
|---|---|
| `base_url` | 供应链系统地址 |
| `user` / `password` | 系统登录账号密码 |
| `gysbm` | 供应商编码（要查询的供应商） |
| `bmbm` | 商户编码（由系统管理员提供） |
| `month` | 抓取月份 `YYYYMM` |
| `start_date` / `end_date` | 起止日期 `YYYYMMDD`（与 `month` 二选一） |
| `output` | 输出文件名 |

---

## ⚠️ 安全须知

1. **`config.json` 含明文密码，已被 `.gitignore` 忽略，请勿提交到 git 或随工具转发他人。**
2. 目标站点为 **HTTP 明文传输**，建议在公司内网使用，避免在公共网络运行。
3. 时间范围不要一次拉太大，建议按月 / 按季度抓取，以免页面超时。
4. 本工具仅用于**你已获授权**访问的账户与系统，请遵守对方网站的服务条款与使用规范。

---

## 📄 License

[MIT](./LICENSE) © mabelwang457-ship-it
