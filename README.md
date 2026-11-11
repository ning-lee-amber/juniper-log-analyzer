# Juniper Log Analyzer

A browser-based tool for quickly analyzing **Juniper Junos error logs** — built by a network engineer who works hands-on with Juniper MX Series routers (MX960 / MX2020 / MX480).

🔗 **Live demo:** https://ning-lee-amber.github.io/juniper-log-analyzer/

<!-- 填:放一張工具截圖。截圖存成 screenshot.png 放在 repo 根目錄,這行就會顯示圖片 -->
![Screenshot](screenshot.png)

---

## Why I built this

During daily device health checks and maintenance windows, reading through long Junos logs line by line is slow and error-prone. This tool helps me spot critical events faster and turn raw logs into a readable summary.

## Features

- **Multi-device support** — parses syslog output from Juniper **SRX / MX / EX / QFX**, including raw SSH session captures (e.g. SecureCRT logs)
- **Batch upload** — select multiple `.log` / `.txt` files at once
- **Auto-classification** — groups messages by **severity** (Critical / Error / Warning / Notice) and by **module**
- **Keyword detection** — flags lines containing `chip`, `error`, `alarm`, `fail`, `warning`, `mqss`, `link_down`, and more
- **Flapping detection** — identifies messages that repeat rapidly (e.g. interface flaps, recurring alarms)
- **Dashboard** — summary counts of Critical / Error / Warning events across all uploaded files, plus an "urgent action" list
- **CPU / Memory Top** — ranks devices by CPU and memory usage against an 80% threshold
- **Down / Up event tracking** — pairs interface / session / slot down and up events, showing first & last down time, count, and which ones have not recovered
- **Device inventory** — lists all devices detected in the logs
- **Keyword quick filters** — one-click filters for `chip`, `error`, `alarm`, `fail`, `warning`, `mqss`, `fpc`, `link_down`
- **Viewed history** — hides logs you have already reviewed, so you focus on what's new
- **Search** — filter by PID, interface, IP address, error code, or any keyword
- **Export to Excel** — export results for reporting
- **今日判讀** — replaces the old urgent list: events are sorted into 要寫 Daily / 要看一下 / 已知雜訊 using the daily-check rules (PFEMAN repeats, logrotate, sshd, single I2C/Readback, BGP connect retries and service-card boot messages are known noise); the same message on several devices is merged into one row
- **介面狂跳** — interfaces and BFD/BGP peers that went down 3+ times; the same peer flapping on two devices at once is merged and flagged as a peer/line problem
- **只看這個時間之後** — hide events before a chosen time (e.g. the last Daily Check export)
- **Daily Check export (F1)** — one click builds the daily check sheet with fixed rules, no AI needed (see below)
- **No installation** — runs in the browser; UI in Traditional Chinese

## How to use

1. Open the [live demo](https://ning-lee-amber.github.io/juniper-log-analyzer/)
2. Click **上傳日誌檔案 (Upload log files)** and select one or more `.log` / `.txt` files
3. Review results grouped by severity and module
4. Use the search bar to drill down by PID / interface / IP / error code

## Daily Check 匯出

1. 第一次使用：按 **設定 Daily 清單**，匯入上一份 Daily check Excel（工具會讀最新一天工作表的「F1」區塊），也可以用 txt/csv（每行：名稱,版本,型號）。同時可以指定要完全忽略的設備。
2. 每天：上傳當天 raw log，按 **Daily Check 匯出**，下載 `YYYYMMDD_Daily_check_F1.xlsx`。

規則：

- 依匯入清單的順序輸出；缺資料的設備照樣列出並填「無資料」
- 欄位：設備 / Version / Model / Memory Usage / CPU Usage / log Analysis / 檢查人員（Memory 在 CPU 前）
- Memory / CPU 為整數 %（同一台多顆 RE 取最大值），超過 80% 會寫進 log Analysis
- 忽略清單中的設備完全不出現
- 介面 down 後有 up → `瞬斷，已恢復` 或 `flapping`，不標黃；down 後沒 up → 整列標黃並列入「需通知客戶」工作表
- 「異常備註」工作表列出缺資料、清單外設備、版本變動

### Python 版：直接寫進總表

**不用命令列：** 直接點兩下 `daily_check.py`（改名成 `daily_check.pyw` 就不會跳出黑色視窗），在視窗裡把總表和當天數據檔一起選起來（按住 Ctrl）、填檢查人員即可；也可以把總表和數據檔一起拖到程式圖示上。新檔會存在總表旁邊。

`daily_check.py` 會在 Daily check 總表裡複製範本工作表（格式完全沿用），新增當天的工作表，另存新檔（不覆蓋原檔）。資料可以給 raw log，也可以給網頁工具匯出的 xlsx。

```bash
pip install openpyxl
python daily_check.py 總表.xlsx 今天的log或匯出檔 --template 前一天工作表 --inspector 你的名字 --replace-name 舊名字
```

**設備清單不會寫進程式或上傳到任何地方**，只存在這台電腦的瀏覽器；換電腦時用「匯出備份 / 匯入備份」搬移。

## Data privacy

- All processing happens **locally in your browser**; logs are not uploaded anywhere
- Sample logs in this repository are **sanitized** — no real hostnames, IP addresses, or customer data

<!-- 填:如果工具其實會把資料送到外部(例如 API),請刪掉上面第一點 -->

## Tech stack

- HTML / CSS / JavaScript
- Hosted on GitHub Pages

## Roadmap

<!-- 填:之後想做的功能,沒有就整段刪掉 -->
- [ ] English UI option
- [ ] Python CLI version for batch processing

## About me

Network engineer working with Juniper MX Series — maintenance, acceptance testing, and lab builds.
**JNCIA-Junos** certified. Currently exploring network automation and cloud.

- GitHub: [@ning-lee-amber](https://github.com/ning-lee-amber)

## License

MIT
