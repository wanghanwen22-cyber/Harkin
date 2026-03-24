# 自动化对账工具（CSV）

这是一个零依赖的 Python 命令行工具，用于自动对账两份 CSV（例如：银行流水 vs 内部台账）。

## 本地 Web 应用
- 页面入口：`/Users/wanghanwen/Desktop/Vscode/工具/ui/dashboard.html`
- 审核日志：`/Users/wanghanwen/Desktop/Vscode/工具/ui/audit-log.html`
- 月度汇总：`/Users/wanghanwen/Desktop/Vscode/工具/ui/monthly-report.html`
- 设置页：`/Users/wanghanwen/Desktop/Vscode/工具/ui/settings.html`
- 共享样式：`/Users/wanghanwen/Desktop/Vscode/工具/ui/app.css`
- 共享逻辑：`/Users/wanghanwen/Desktop/Vscode/工具/ui/app.js`

建议用本地静态服务器打开：

```bash
cd /Users/wanghanwen/Desktop/Vscode/工具
python3 -m http.server 8000
```

然后访问：

```text
http://127.0.0.1:8000/ui/dashboard.html
```

说明：
- 不依赖后端，所有数据都在浏览器本地完成解析和存储
- 当前本地版支持 `CSV`、`JSON`、`XLSX` 文件上传
- `XLSX` 解析在浏览器端完成，不依赖后端；要求浏览器支持 `DecompressionStream`
- 任务详情页支持查看 matched / unmatched 明细，并分别导出 `matched.csv`、`unmatched_bank.csv`、`unmatched_ledger.csv`、`summary.json`

## 功能
- 自动匹配：
  - 第一轮：按 `tx_id + 金额 + 日期容差` 匹配
  - 第二轮：按 `金额 + 日期容差` 贪心匹配
- 输出结果：
  - `matched.csv`（匹配成功记录）
  - `unmatched_bank.csv`（银行侧未匹配）
  - `unmatched_ledger.csv`（台账侧未匹配）
  - `summary.json`（统计摘要）

## 运行要求
- Python 3.9+

## 快速开始
```bash
python3 reconciler.py \
  --bank examples/bank.csv \
  --ledger examples/ledger.csv \
  --out-dir reconcile_output
```

## 常用参数
- `--bank` / `--ledger`：两侧 CSV 文件路径（必填）
- `--out-dir`：输出目录（默认 `./reconcile_output`）
- `--amount-tolerance`：金额容差（默认 `0.00`，如 `0.01`）
- `--date-tolerance-days`：日期容差天数（默认 `0`）

字段映射参数（用于你的真实字段名与默认不一致的情况）：
- 银行侧：
  - `--bank-amount-field`（默认 `amount`）
  - `--bank-date-field`（默认 `date`）
  - `--bank-id-field`（默认 `tx_id`）
- 台账侧：
  - `--ledger-amount-field`（默认 `amount`）
  - `--ledger-date-field`（默认 `date`）
  - `--ledger-id-field`（默认 `tx_id`）

## 自动化运行（macOS/Linux cron 示例）
每天凌晨 1 点执行一次：
```cron
0 1 * * * cd /Users/wanghanwen/Desktop/Vscode/工具 && /usr/bin/python3 reconciler.py --bank /path/to/bank.csv --ledger /path/to/ledger.csv --out-dir /path/to/out
```

## 日期格式支持
默认支持：
- `YYYY-MM-DD`
- `YYYY/MM/DD`
- `YYYYMMDD`
- `DD/MM/YYYY`
- `MM/DD/YYYY`
