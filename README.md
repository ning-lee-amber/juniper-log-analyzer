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
- **No installation** — runs in the browser; UI in Traditional Chinese

## How to use

1. Open the [live demo](https://ning-lee-amber.github.io/juniper-log-analyzer/)
2. Click **上傳日誌檔案 (Upload log files)** and select one or more `.log` / `.txt` files
3. Review results grouped by severity and module
4. Use the search bar to drill down by PID / interface / IP / error code

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
